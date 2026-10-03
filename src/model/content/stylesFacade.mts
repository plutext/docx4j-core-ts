// CR-002 section 22.2: `pkg.styles.ensure(ids)`, splicing a style definition a document lacks out
// of the same defaults `createPackage()` writes, with what it is based on and linked to.
//
// The need is an editor's: inserting a footnote needs `FootnoteText` and `FootnoteReference` and
// their linked character styles, and a document that has neither shows the note in Normal. Doing it
// in a consumer means carrying a table of Word's definitions, which is what this replaces.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy, unmarshalNode, type Jsonix } from '@docx4j/generated-objects-ts';
import { linkParents } from './tree.mjs';
import { parseXml } from '../../xml/dom.mjs';
import { DEFAULT_STYLES_XML } from '../../parts/wml/defaultStyles.mjs';
import { SPLICEABLE_STYLES_XML } from '../../parts/wml/spliceableStyles.mjs';
import { commentStyles } from '../../parts/wml/commentStyles.mjs';
import { placeholderStyle } from '../../parts/wml/placeholderStyle.mjs';
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
 * appears to contain one, and why this facade exists (CR-002 section 22.2) - with Word's two comment
 * styles and its `PlaceholderText` after them.
 */
let sources: Promise<StyleSources> | undefined;

/** The defaults a created styles part holds, and those with the nine commented-out ones after them. */
interface StyleSources {
  defaults: wml.Style[];
  all: wml.Style[];
}

function styleSources(): Promise<StyleSources> {
  return (sources ??= (async () => {
    const read = async (xml: string): Promise<wml.Style[]> =>
      (await unmarshalNode<Jsonix.TypedNamedValue<wml.Styles>>(parseXml(xml))).value.style ?? [];
    const defaults = await read(DEFAULT_STYLES_XML);
    // and Word's two comment styles, which the comment parts add when they are made, for a caller
    // that makes the comment itself (CR-002 section 37), and Word's PlaceholderText, which a
    // placeholder run names (CR-002 section 40)
    return { defaults, all: [...defaults, ...(await read(SPLICEABLE_STYLES_XML)), ...commentStyles(), placeholderStyle()] };
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
   * add a footnote to. An id the defaults do not carry throws a `Docx4JException` naming it: the
   * caller asked for a definition nothing here has, and inventing one would be worse than saying so.
   * The whole closure is worked out before anything is written, so a call that throws has changed
   * nothing (CR-002 section 31).
   *
   * A document that already has every id keeps its styles part **byte for byte**: the part is not
   * unmarshalled at all when its XML names them, the check `ensureCommentStyles` uses, and when that
   * text check misses (another prefix, single quotes) it is read privately rather than unmarshalled,
   * so the part is still written back from its source unless something is added.
   */
  async ensure(ids: string | readonly string[]): Promise<string[]> {
    const wanted = typeof ids === 'string' ? [ids] : [...ids];
    for (const id of wanted) {
      if (typeof id !== 'string' || id === '') throw new Docx4JException(`Not a styleId: ${String(id)}`);
    }
    if (wanted.length === 0) return [];
    const existing = this.pkg.getMainDocumentPart().styleDefinitionsPart;

    // The cheap path: an untouched part whose XML names them all is left alone, so a document that
    // has them keeps its styles part byte for byte. The comments must go first - docx4j's own
    // default styles carry nine definitions COMMENTED OUT, so a plain `includes` finds
    // `w:styleId="FootnoteText"` in a document that does not have the style and this would answer
    // "nothing to do" for exactly the documents the facade exists for.
    if (existing !== undefined && !existing.isUnmarshalled) {
      const xml = (await existing.getXml()).replace(/<!--[\s\S]*?-->/g, '');
      if (wanted.every((id) => xml.includes(`w:styleId="${id}"`))) return [];
    }

    // What the document holds now, read without unmarshalling; a part created below holds the defaults.
    const source = await styleSources();
    const present = existing === undefined ? source.defaults : ((await existing.readContents()).style ?? []);
    const plan = closureOf(wanted, present, source.all);

    // a document with no styles part gets the defaults, whether or not anything else is needed
    const part = existing ?? this.pkg.ensureStyleDefinitionsPart();
    if (plan.length === 0) return [];
    const styles = await part.getContents();
    const list = (styles.style ??= []);
    for (const template of plan) {
      const style = deepCopy(template);
      list.push(style);
      linkParents(style, styles);
    }
    this.pkg.refreshPropertyResolver();
    return plan.map((style) => style.styleId as string);
  }
}

/**
 * The definitions to splice, in the order they are written: each wanted id the document lacks,
 * after what it is `w:basedOn` and `w:link`ed to, recursively. Throws for an id nothing carries.
 */
function closureOf(wanted: readonly string[], present: readonly wml.Style[], source: readonly wml.Style[]): wml.Style[] {
  const plan: wml.Style[] = [];
  const has = (id: string): boolean => present.some((s) => s.styleId === id) || plan.some((s) => s.styleId === id);
  const add = (id: string, seen: Set<string>): void => {
    if (has(id) || seen.has(id)) return;
    seen.add(id);
    const template = source.find((s) => s.styleId === id);
    if (template === undefined) {
      throw new Docx4JException(`No definition to splice for the style ${id}: neither docx4j's default styles, the nine it comments out, the two comment styles nor PlaceholderText carry one. Define it yourself on the styles part.`);
    }
    // what it is based on and linked to must be there too, and before it reads better in the part
    for (const dependency of [template.basedOn?.val, template.link?.val]) {
      if (typeof dependency === 'string') add(dependency, seen);
    }
    plan.push(template);
  };
  for (const id of wanted) add(id, new Set());
  return plan;
}
