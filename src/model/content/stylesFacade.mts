// CR-002 section 22.2: `pkg.styles.ensure(ids)`, splicing a style definition a document lacks out
// of the same defaults `createPackage()` writes, with what it is based on and linked to.
//
// The need is an editor's: inserting a footnote needs `FootnoteText` and `FootnoteReference` and
// their linked character styles, and a document that has neither shows the note in Normal. Doing it
// in a consumer means carrying a table of Word's definitions, which is what this replaces.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { unmarshalNode, type Jsonix } from '@docx4j/generated-objects-ts';
import { linkParents } from './tree.mjs';
import { parseXml } from '../../xml/dom.mjs';
import { DEFAULT_STYLES_XML } from '../../parts/wml/defaultStyles.mjs';
import { SPLICEABLE_STYLES_XML } from '../../parts/wml/spliceableStyles.mjs';
import { Docx4JException } from '../../opc/exceptions.mjs';
import type { StyleDefinitionsPart, MainDocumentPart } from '../../parts/wml/index.mjs';

/** What the facade needs of its package; `WordprocessingMLPackage` satisfies it. */
export interface StylePackageLike {
  getMainDocumentPart(): MainDocumentPart;
  refreshPropertyResolver(): void;
  /** Creates the styles part, with the defaults, when the document has none. */
  ensureStyleDefinitionsPart(): StyleDefinitionsPart;
}

/**
 * Where a definition is looked up, unmarshalled once and never handed out (a copy is spliced in):
 * the defaults `createPackage()` writes, and then the nine styles docx4j's own `styles.xml` carries
 * **commented out** - which is why a created document has no `FootnoteText` although the file
 * appears to contain one, and why this facade exists (CR-002 section 22.2).
 */
let sources: Promise<wml.Style[]> | undefined;

function styleSources(): Promise<wml.Style[]> {
  return (sources ??= (async () => {
    const out: wml.Style[] = [];
    for (const xml of [DEFAULT_STYLES_XML, SPLICEABLE_STYLES_XML]) {
      const element = await unmarshalNode<Jsonix.TypedNamedValue<wml.Styles>>(parseXml(xml));
      out.push(...(element.value.style ?? []));
    }
    return out;
  })());
}

/**
 * The document's style definitions as verbs that need no `Body`: `pkg.styles`.
 *
 * `ensure` is the one verb, and the point of it is the closure: a style is useless without what it
 * is `w:basedOn` and `w:link`ed to, and a caller asking for `FootnoteText` means the four styles
 * that make a footnote look like a footnote, not one.
 */
export class StylesFacade {
  constructor(private readonly pkg: StylePackageLike) {}

  /**
   * Adds each style the document lacks, from the same defaults `createPackage()` writes, together
   * with the styles it is `w:basedOn` and `w:link`ed to, recursively. Returns the ids actually
   * added, in the order they were written, so a caller can tell whether it changed anything.
   *
   * A style already in the document is left exactly as it is - the document's own definition wins
   * over the default, always, since a consumer must not silently restyle a document it was asked to
   * add a footnote to. An id the defaults do not carry throws `ItemNotFound`: the caller asked for a
   * definition nothing here has, and inventing one would be worse than saying so.
   *
   * A document that already has every id keeps its styles part **byte for byte**: the part is not
   * unmarshalled at all when its XML names them, the check `ensureCommentStyles` uses.
   */
  async ensure(ids: string | readonly string[]): Promise<string[]> {
    const wanted = typeof ids === 'string' ? [ids] : [...ids];
    if (wanted.length === 0) return [];
    const part = this.pkg.getMainDocumentPart().styleDefinitionsPart ?? this.pkg.ensureStyleDefinitionsPart();

    // The cheap path: an untouched part whose XML names them all is left alone, so a document that
    // has them keeps its styles part byte for byte. The comments must go first - docx4j's own
    // default styles carry nine definitions COMMENTED OUT, so a plain `includes` finds
    // `w:styleId="FootnoteText"` in a document that does not have the style and this would answer
    // "nothing to do" for exactly the documents the facade exists for.
    if (!part.isUnmarshalled) {
      const xml = (await part.getXml()).replace(/<!--[\s\S]*?-->/g, '');
      if (wanted.every((id) => xml.includes(`w:styleId="${id}"`))) return [];
    }

    const styles = await part.getContents();
    const list = (styles.style ??= []);
    const has = (id: string): boolean => list.some((s) => s.styleId === id);
    const source = await styleSources();
    const added: string[] = [];

    const addWithClosure = (id: string, seen: Set<string>): void => {
      if (has(id) || seen.has(id)) return;
      seen.add(id);
      const template = source.find((s) => s.styleId === id);
      if (template === undefined) {
        throw new Docx4JException(`No definition to splice for the style ${id}: neither docx4j's default styles nor the nine it comments out carry one. Define it yourself on the styles part.`);
      }
      // what it is based on and linked to must be there too, and before it reads better in the part
      for (const dependency of [template.basedOn?.val, template.link?.val]) {
        if (typeof dependency === 'string') addWithClosure(dependency, seen);
      }
      const style = copyStyle(template);
      list.push(style);
      linkParents(style, styles);
      added.push(id);
    };

    for (const id of wanted) {
      if (typeof id !== 'string' || id === '') throw new Docx4JException(`Not a styleId: ${String(id)}`);
      addWithClosure(id, new Set());
    }
    if (added.length > 0) this.pkg.refreshPropertyResolver();
    return added;
  }
}

/**
 * A style copied out of the shared defaults. `structuredClone` rather than the facade's `deepCopy`
 * because the tree carries `PARENT` back-references, which `linkParents` re-establishes for the
 * copy's new home; cloning them would drag the whole default tree along.
 */
function copyStyle(style: wml.Style): wml.Style {
  const copy: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(style)) {
    if (key === 'PARENT') continue;
    copy[key] = value === null || typeof value !== 'object' ? value : structuredClone(stripParents(value));
  }
  return copy as unknown as wml.Style;
}

/** `structuredClone` cannot take a cycle through `PARENT`, so they go before it is called. */
function stripParents(value: object): object {
  const out: Record<string, unknown> = Array.isArray(value) ? ([] as unknown as Record<string, unknown>) : {};
  for (const [key, v] of Object.entries(value)) {
    if (key === 'PARENT') continue;
    out[key] = v === null || typeof v !== 'object' ? v : stripParents(v);
  }
  return Array.isArray(value) ? Object.values(out) : out;
}
