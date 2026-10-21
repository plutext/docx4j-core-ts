// CR-012 phase 2: what one reading of a package needs, read before the walk begins (`PackageTrees`), and
// what a story needs from the rest of the document (`TextContext`, CR-012 section 4.4): the note references'
// order, the numbering states per story, the bookmarks, and the views' plans.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as w15 from '@docx4j/generated-objects-ts/modules/org_docx4j_w15';
import type * as core from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_core';
import type * as extended from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_extended';
import type * as custom from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_custom';
import type { Part } from '../../parts/Part.mjs';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import type { PropertyResolver } from '../properties/PropertyResolver.mjs';
import type { NumberingDefinitions } from '../listnumbering/definitions.mjs';
import type { NumberingPartLike } from '../listnumbering/Emulator.mjs';
import { NumberingStates } from '../listnumbering/state.mjs';
import { Body } from '../content/Body.mjs';
import { ResolvedView } from '../content/ResolvedView.mjs';
import type { XPathEngine } from '../customxml/xpath.mjs';
import { parsePrefixMappings } from '../customxml/xpath.mjs';
import { parseXml, decodeXmlText } from '../../xml/dom.mjs';
import { walk } from '@docx4j/generated-objects-ts/builders/wml';
import type { NoteType } from './TextResult.mjs';

/** The `w:storeItemID`s of the core and application properties, which a control can be bound to (docx4j `BindingHandler`). */
const CORE_PROPERTIES_STOREITEMID = '{6C3C8BC8-F283-45AE-878A-BAB7291924A1}';
const EXTENDED_PROPERTIES_STOREITEMID = '{6668398D-A668-4E3E-A5EB-62B293D839F1}';

/** As much of a `WordprocessingMLPackage` as a reading uses, so that this module does not import the package. */
export interface TextPackage {
  getMainDocumentPart(): MainLike;
  getPropertyResolver(): Promise<PropertyResolver>;
  getNumberingEmulator(): Promise<unknown>;
  readonly propertyResolverOrUndefined?: PropertyResolver;
  readonly docPropsCorePart?: XmlPart<core.CoreProperties>;
  readonly docPropsExtendedPart?: XmlPart<extended.Properties>;
  readonly docPropsCustomPart?: XmlPart<custom.Properties>;
  readonly customXmlDataStorageParts: ReadonlyMap<string, { parseDocument(): Promise<Document> }>;
  xpathEngine: XPathEngine;
}

/** As much of a `MainDocumentPart` as a reading uses. */
export interface MainLike extends XmlPart<wml.Document> {
  readonly styleDefinitionsPart?: Part;
  readonly numberingDefinitionsPart?: XmlPart<wml.Numbering> & NumberingPartLike & { getDefinitions(): Promise<NumberingDefinitions> };
  readonly documentSettingsPart?: XmlPart<wml.CTSettings>;
  readonly footnotesPart?: XmlPart<wml.CTFootnotes>;
  readonly endnotesPart?: XmlPart<wml.CTEndnotes>;
  readonly commentsPart?: XmlPart<wml.Comments>;
  readonly commentsExtendedPart?: XmlPart<w15.CTCommentsEx>;
  readonly glossaryDocumentPart?: XmlPart<wml.GlossaryDocument>;
  readonly headerParts: XmlPart<wml.Hdr>[];
  readonly footerParts: XmlPart<wml.Ftr>[];
}

/** `readContents()`, undefined for a part that cannot be read. */
async function readQuietly<T>(part: XmlPart<T> | undefined): Promise<T | undefined> {
  if (!part) return undefined;
  try {
    return await part.readContents();
  } catch {
    return undefined;
  }
}

/**
 * The trees one reading walks, each read **privately** where its part was not unmarshalled
 * (`XmlPart.readContents`), so that reading a document costs no part its byte-for-byte round trip
 * (REQ-204), and the live tree where it was. Every story is read, so that the views' plan covers them
 * all ("the view is the package's"); the glossary only when it is included.
 */
export class PackageTrees {
  main: { part: MainLike; document: wml.Document } | undefined;
  footnotes: { part: Part; tree: wml.CTFootnotes } | undefined;
  endnotes: { part: Part; tree: wml.CTEndnotes } | undefined;
  comments: { part: Part; tree: wml.Comments } | undefined;
  commentsExtended: w15.CTCommentsEx | undefined;
  /** Each header and footer part related from the main document part, with its content. */
  readonly headersAndFooters = new Map<Part, wml.Hdr | wml.Ftr>();
  glossary: { part: Part; tree: wml.GlossaryDocument } | undefined;
  settings: wml.CTSettings | undefined;
  /** The document properties under `core`, `app` and `custom` (REQ-515). */
  properties: Record<string, unknown> | undefined;
  /** The resolver, where the package has a styles part (docx4j makes none without one). */
  resolver: PropertyResolver | undefined;
  numberingPart: NumberingPartLike | undefined;
  numbering: NumberingDefinitions | undefined;
  /** Each binding store's DOM, by its upper-cased item id. */
  private readonly stores = new Map<string, Document | undefined>();
  private xpath: XPathEngine | undefined;
  private anchored: Set<string> | undefined;

  private constructor(readonly pkg: TextPackage | undefined) {}

  /** The trees of a package. */
  static async read(pkg: TextPackage): Promise<PackageTrees> {
    const trees = new PackageTrees(pkg);
    let main: MainLike | undefined;
    try {
      main = pkg.getMainDocumentPart();
    } catch {
      main = undefined;
    }
    if (main) {
      trees.main = { part: main, document: await main.readContents() };
      if (main.styleDefinitionsPart) {
        try {
          trees.resolver = await pkg.getPropertyResolver();
        } catch {
          trees.resolver = undefined;                                  // hidden text and indents not resolved through styles
        }
      }
      const numbering = main.numberingDefinitionsPart;
      if (numbering) {
        try {
          await pkg.getNumberingEmulator();
          trees.numbering = await numbering.getDefinitions();
          trees.numberingPart = numbering;
        } catch {
          trees.numberingPart = undefined;
        }
      }
      trees.settings = await readQuietly(main.documentSettingsPart);
      const footnotes = await readQuietly(main.footnotesPart);
      if (footnotes && main.footnotesPart) trees.footnotes = { part: main.footnotesPart, tree: footnotes };
      const endnotes = await readQuietly(main.endnotesPart);
      if (endnotes && main.endnotesPart) trees.endnotes = { part: main.endnotesPart, tree: endnotes };
      const comments = await readQuietly(main.commentsPart);
      if (comments && main.commentsPart) trees.comments = { part: main.commentsPart, tree: comments };
      trees.commentsExtended = await readQuietly(main.commentsExtendedPart);
      for (const part of [...main.headerParts, ...main.footerParts] as XmlPart<wml.Hdr | wml.Ftr>[]) {
        const content = await readQuietly(part);
        if (content) trees.headersAndFooters.set(part, content);
      }
    }
    trees.properties = await properties(pkg);
    await trees.readStores(pkg);
    return trees;
  }

  /** Trees with no package: content read as it stands (docx4j's planless `Texts.of(content)`). */
  static none(): PackageTrees {
    return new PackageTrees(undefined);
  }

  /** The glossary document, read when `include: glossary` first asks for it. */
  async readGlossary(): Promise<void> {
    const part = this.main?.part.glossaryDocumentPart;
    if (this.glossary || !part) return;
    const tree = await readQuietly(part);
    if (tree) this.glossary = { part, tree };
  }

  // --- bound values (REQ-509) ---------------------------------------------------------------------

  private async readStores(pkg: TextPackage): Promise<void> {
    let context: Node | undefined;
    for (const [id, part] of pkg.customXmlDataStorageParts) {
      try {
        const doc = await part.parseDocument();
        this.stores.set(id.toUpperCase(), doc);
        context ??= doc;
      } catch {
        this.stores.set(id.toUpperCase(), undefined);
      }
    }
    for (const [id, part] of [[CORE_PROPERTIES_STOREITEMID, pkg.docPropsCorePart], [EXTENDED_PROPERTIES_STOREITEMID, pkg.docPropsExtendedPart]] as const) {
      if (!part) continue;
      try {
        // the part's XML as it would be saved, parsed apart from it: the part keeps its round trip
        const doc = parseXml(decodeXmlText(await part.getBytes()));
        this.stores.set(id, doc);
        context ??= doc;
      } catch {
        this.stores.set(id, undefined);
      }
    }
    if (context === undefined) return;
    try {
      await pkg.xpathEngine.ready(context);
      this.xpath = pkg.xpathEngine;
    } catch {
      this.xpath = undefined;                                         // no XPath here: every control's surface is read
    }
  }

  /**
   * A data binding's value, docx4j `BoundValues`: the binding's XPath, with its own prefix mappings,
   * against the custom XML part its `w:storeItemID` names (or the core or application properties); the
   * string value of the first node selected. Undefined where the binding does not resolve.
   */
  boundValue(binding: wml.CTDataBinding): string | undefined {
    if (!binding.xpath || !binding.storeItemID || !this.xpath) return undefined;
    const doc = this.stores.get(binding.storeItemID.toUpperCase());
    if (!doc) return undefined;
    try {
      const mappings = parsePrefixMappings(binding.prefixMappings);
      delete mappings[''];
      const nodes = this.xpath.select(binding.xpath, doc, mappings);
      const node = nodes[0];
      if (!node) return undefined;
      return node.nodeType === 2 ? (node as Attr).value : (node.textContent ?? '');
    } catch {
      return undefined;
    }
  }

  // --- the comments anchored in the document (REQ-504) -----------------------------------------------

  /**
   * The ids of the comments anchored in the document: a `w:commentReference` or a `w:commentRangeStart`
   * in the main document, a header or footer, or a note, in any branch.
   */
  anchoredComments(): Set<string> {
    if (this.anchored) return this.anchored;
    const ids = new Set<string>();
    const roots: unknown[] = [this.main?.document, ...this.headersAndFooters.values(), this.footnotes?.tree, this.endnotes?.tree];
    for (const root of roots) {
      if (root === undefined) continue;
      walk(root, (v) => {
        const tn = (v as { TYPE_NAME?: string }).TYPE_NAME;
        if (tn === 'org_docx4j_wml.R.CommentReference' || tn === 'org_docx4j_wml.CommentRangeStart') {
          const id = (v as { id?: unknown }).id;
          if (id !== undefined && id !== null) ids.add(String(id));
        }
      });
    }
    return (this.anchored = ids);
  }

  /** Every story's body, over these trees, for the views' plan: the main document, each note and comment, the headers and footers. */
  stories(): Body[] {
    const pkg = this.pkg as unknown as ConstructorParameters<typeof Body>[3];
    const out: Body[] = [];
    if (this.main?.document.body) out.push(new Body(this.main.part, this.main.document.body as { content?: never[] }, 'body', pkg));
    for (const [holder, kind] of [[this.footnotes, 'footnote'], [this.endnotes, 'endnote']] as const) {
      if (!holder) continue;
      const notes = (kind === 'footnote' ? (holder.tree as wml.CTFootnotes).footnote : (holder.tree as wml.CTEndnotes).endnote) ?? [];
      for (const note of notes) out.push(new Body(holder.part as XmlPart<unknown>, note as { content?: never[] }, `${kind}:${note.id}`, pkg));
    }
    for (const comment of this.comments?.tree.comment ?? []) {
      out.push(new Body(this.comments!.part as XmlPart<unknown>, comment as { content?: never[] }, `comment:${comment.id}`, pkg));
    }
    for (const [part, content] of this.headersAndFooters) out.push(new Body(part as XmlPart<unknown>, content as { content?: never[] }, 'story', pkg));
    return out;
  }

  /** The part whose tree holds an object, by its `PARENT`s, or undefined. */
  partOf(object: unknown): Part | undefined {
    let root: unknown = object;
    for (let o = object as { PARENT?: unknown } | undefined; o; o = o.PARENT as { PARENT?: unknown } | undefined) root = o;
    if (this.main && (root === this.main.document || root === this.main.document.body)) return this.main.part;
    for (const [part, content] of this.headersAndFooters) if (root === content) return part;
    if (this.footnotes && root === this.footnotes.tree) return this.footnotes.part;
    if (this.endnotes && root === this.endnotes.tree) return this.endnotes.part;
    if (this.comments && root === this.comments.tree) return this.comments.part;
    if (this.glossary && root === this.glossary.tree) return this.glossary.part;
    return undefined;
  }
}

/**
 * What a story needs from the rest of the document (CR-012 section 4.4): the note references' order, for
 * `[n]`; the numbering states per story, for list labels and `value`; the bookmarks, for `fields: update`;
 * and, made once, each view's plan of what Accept All or Reject All would leave. A call given a context
 * uses it and adds to it, so that the editor's calls over one part at a time, in the order the
 * specification reads the stories (REQ-501), number the notes and count the lists as one reading of the
 * whole package does; `TextResult.concat` assembles their results.
 *
 * A context is a snapshot of the package as it was read (`Texts.contextOf`): after the package changes,
 * make another. It is for one reading, in one view: the calls under it take the same options, as the parts
 * of one `Texts.of(pkg)` would.
 */
export class TextContext {
  /** Notes referenced, in reference order: id to number, 0 for a note with a custom mark. */
  readonly footnotes = new Map<string, number>();
  readonly endnotes = new Map<string, number>();
  private footnoteCount = 0;
  private endnoteCount = 0;
  /** The numbering states, one per story (`Emulator`). */
  readonly numberingStates = new NumberingStates();
  /** The `w:lastRenderedPageBreak`s of the main story so far (REQ-708). */
  pageMarkers = 0;
  /** The bookmarks' text, by lower-cased name (REQ-506's REF, CR-012 phase 5a). */
  readonly bookmarks = new Map<string, string>();
  /** Whether the main story has been read under this context, which numbers the notes. */
  mainRead = false;
  /**
   * The view the context's first call read in: its note numbers and counts are that view's, so a call in
   * another view is refused (`Texts.of`).
   */
  view: 'accepted' | 'original' | 'markup' | undefined;
  private readonly plans = new Map<boolean, ResolvedView>();
  /** Per container, then per options, what the blocks begin inside (`Extractor.fieldsAtBlocks`), made once; shared with siblings. */
  private starts = new WeakMap<object, Map<string, unknown>>();

  constructor(readonly trees: PackageTrees) {}

  /** What `make` gives for a container's blocks under options keyed `key`, made once in this context (CR-012 section 17.13). */
  blockStarts<T>(container: object, key: string, make: () => T): T {
    let byKey = this.starts.get(container);
    if (!byKey) this.starts.set(container, byKey = new Map());
    if (!byKey.has(key)) byKey.set(key, make());
    return byKey.get(key) as T;
  }

  /** The package this context reads. */
  get pkg(): TextPackage | undefined {
    return this.trees.pkg;
  }

  /** A note's number, given at its first reference: in reference order, footnotes and endnotes apart, 0 for a custom mark (REQ-503). */
  noteNumber(type: NoteType, id: string, customMark: boolean): number {
    const numbers = type === 'footnote' ? this.footnotes : this.endnotes;
    let n = numbers.get(id);
    if (n === undefined) {
      n = customMark ? 0 : type === 'footnote' ? ++this.footnoteCount : ++this.endnoteCount;
      numbers.set(id, n);
    }
    return n;
  }

  /** The plan of Accept All (`accepted`) or Reject All, over every story, made once. */
  plan(accepted: boolean): ResolvedView {
    let plan = this.plans.get(accepted);
    if (!plan) {
      const stories = this.trees.stories();
      plan = accepted ? ResolvedView.accepted(stories) : ResolvedView.original(stories);
      this.plans.set(accepted, plan);
    }
    return plan;
  }

  /** A context over the same trees and plans, its own notes and states: for a reading whose blocks are not the result's. */
  sibling(): TextContext {
    const c = new TextContext(this.trees);
    for (const [k, v] of this.plans) c.plans.set(k, v);
    c.starts = this.starts;
    return c;
  }

  /** Adopts the note numbers another context gave (`sibling`'s, for a notes part read before its story). */
  adoptNotes(other: TextContext): void {
    for (const [id, n] of other.footnotes) if (!this.footnotes.has(id)) this.footnotes.set(id, n);
    for (const [id, n] of other.endnotes) if (!this.endnotes.has(id)) this.endnotes.set(id, n);
    this.footnoteCount = Math.max(this.footnoteCount, other.footnoteCount);
    this.endnoteCount = Math.max(this.endnoteCount, other.endnoteCount);
  }
}

// --- the document properties (REQ-515; docx4j `Metadata`) --------------------------------------------

/** The properties under `core`, `app` and `custom`, each where the package has that part. */
async function properties(pkg: TextPackage): Promise<Record<string, unknown>> {
  const out: Record<string, unknown> = {};
  const c = await readQuietly(pkg.docPropsCorePart);
  if (c) out.core = coreProperties(c);
  const a = await readQuietly(pkg.docPropsExtendedPart);
  if (a) out.app = appProperties(a);
  const x = await readQuietly(pkg.docPropsCustomPart);
  if (x) out.custom = customProperties(x);
  return out;
}

function put(m: Record<string, unknown>, name: string, value: unknown): void {
  if (value === undefined || value === null || value === '') return;
  m[name] = value;
}

function literal(o: unknown): string | undefined {
  const v = typeof o === 'object' && o !== null && 'value' in o && 'name' in o ? (o as { value: unknown }).value : o;
  const content = (v as { content?: string[] } | undefined)?.content;
  return content === undefined || content.length === 0 ? undefined : content.join('');
}

/** The core properties in the schema's order. */
function coreProperties(c: core.CoreProperties): Record<string, unknown> {
  const m: Record<string, unknown> = {};
  put(m, 'title', literal(c.title));
  put(m, 'subject', literal(c.subject));
  put(m, 'creator', literal(c.creator));
  put(m, 'keywords', c.keywords);
  put(m, 'description', literal(c.description));
  put(m, 'category', c.category);
  put(m, 'contentStatus', c.contentStatus);
  put(m, 'language', literal(c.language));
  put(m, 'identifier', literal(c.identifier));
  put(m, 'version', c.version);
  put(m, 'revision', c.revision);
  put(m, 'lastModifiedBy', c.lastModifiedBy);
  put(m, 'created', literal(c.created));
  put(m, 'modified', literal(c.modified));
  put(m, 'lastPrinted', calendarText(c.lastPrinted));
  return m;
}

/** The application properties that describe the document; the statistics, which depend on layout, left out. */
function appProperties(p: extended.Properties): Record<string, unknown> {
  const m: Record<string, unknown> = {};
  put(m, 'Application', p.application);
  put(m, 'AppVersion', p.appVersion);
  put(m, 'Company', p.company);
  put(m, 'Manager', p.manager);
  put(m, 'Template', p.template);
  put(m, 'HyperlinkBase', p.hyperlinkBase);
  put(m, 'TotalTime', p.totalTime);
  put(m, 'DocSecurity', p.docSecurity);
  return m;
}

/** Each custom property's typed value, in the part's order; values of other types left out. */
function customProperties(p: custom.Properties): Record<string, unknown> {
  const m: Record<string, unknown> = {};
  for (const property of p.property ?? []) {
    if (property.name === undefined) continue;
    put(m, property.name, customValue(property));
  }
  return m;
}

function customValue(p: custom.Properties.Property): unknown {
  for (const key of ['lpwstr', 'lpstr', 'bstr', 'bool', 'i1', 'i2', 'i4', 'i8', '_int', 'ui1', 'ui2', 'ui4', 'ui8', 'uint', 'r4', 'r8', 'decimal'] as const) {
    const v = p[key];
    if (v !== undefined && v !== null) return v;
  }
  if (p.filetime !== undefined) return calendarText(p.filetime);
  if (p.date !== undefined) return calendarText(p.date);
  if (p.cy !== undefined) return p.cy;
  if (p.error !== undefined) return p.error;
  return undefined;
}

/** An XML calendar as its lexical form. */
export function calendarText(date: unknown): string | undefined {
  if (date === undefined || date === null) return undefined;
  if (typeof date === 'string') return date;
  const c = date as { year?: number; month?: number; day?: number; hour?: number; minute?: number; second?: number; fractionalSecond?: number; timezone?: number };
  if (typeof c.year !== 'number') return String(date);
  const p2 = (n: number | undefined): string => String(n ?? 0).padStart(2, '0');
  let out = `${String(c.year).padStart(4, '0')}-${p2(c.month)}-${p2(c.day)}`;
  if (c.hour !== undefined) {
    out += `T${p2(c.hour)}:${p2(c.minute)}:${p2(c.second)}`;
    if (c.fractionalSecond) out += String(c.fractionalSecond).replace(/^0/, '');
  }
  if (c.timezone !== undefined) {
    if (c.timezone === 0) out += 'Z';
    else {
      const sign = c.timezone < 0 ? '-' : '+';
      const minutes = Math.abs(c.timezone);
      out += `${sign}${p2(Math.floor(minutes / 60))}:${p2(minutes % 60)}`;
    }
  }
  return out;
}
