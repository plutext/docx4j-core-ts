// SPDX-License-Identifier: Apache-2.0
//
// A Word document given as Flat OPC (`pkg:package`), read into a document as Word 365 resolves an altChunk holding it (CR-005
// section 11; the editor's ED-005 proposal 46, part 2; Word's saves of 46a and 46b). `contentOf` takes a package's body and the
// parts its relationships reach; this does that and what Word does besides:
//
// - **Styles**: a style the content uses that the document lacks is copied with its `basedOn`, `link` and `next` chain, its
//   `w:default` dropped; a style both define is the document's. Matched by `w:styleId`.
// - **Numbering**: each list the content uses becomes a new `w:num` over a copy of its `w:abstractNum`, under fresh ids and a
//   fresh `w:nsid`, and the paragraphs' `w:numPr` (and a copied style's) name it: the document's own lists go on unbroken.
// - **Notes**: a footnote or endnote the content references is copied into the document's notes part under a free id, the
//   reference rewritten, the parts its own relationships reach copied too.
// - **Dropped**: the final section (silently, as Word does), a section break inside the content, and comments' ranges and
//   references (unmeasured: the check's fragment held none), each counted.
// - **Ids**: a bookmark's, a control's `w:id` and a drawing's `wp:docPr` id made fresh only where the document already uses it.
// - **What the package names and doesn't define** (Word's `insertOoxml`, insertooxml-46 case 4): a style reference reads as the
//   default, so it goes; a list becomes a new one over Word's default definition, never the document's list of that number.
// - **`keepSourceFormatting`** (Word's `w:matchSrc`, 46b): the body content's paragraph and character styles written as direct
//   formatting, `w:pStyle` and `w:rStyle` removed, the styles still copied; a note's content keeps its styles, as Word's did.
//
// Not merged, as Word merges none of them: the theme, the fonts and the document defaults. `insertOoxml` reads a package through
// this too (`insertedContentOf`, CR-005 section 11.6), as Word 365's `insertOoxml` does by the same rule.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy, unmarshalNode } from '@docx4j/generated-objects-ts';
import { parseXml } from '../../xml/dom.mjs';
import { sdtProperty } from '@docx4j/generated-objects-ts/builders/wml';
import { PartName } from '../../opc/PartName.mjs';
import type { Part } from '../../parts/Part.mjs';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import { applyPPr, applyRPr, applyStyleLevel, isEmptyPPr, isEmptyRPr } from '../properties/styleUtil.mjs';
import type { PropertyResolver } from '../properties/PropertyResolver.mjs';
import { linkParents, type Element } from './tree.mjs';
import { loadFlatOpc, copyReferencedParts, rewriteRelationshipIds, contentOf, isFlatOpc } from './ooxml.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** What the import needs of the document's package: the main part, a styles part made where there is none, and a refresh. */
export interface FragmentTarget {
  getMainDocumentPart(): FragmentMainPart;
  ensureStyleDefinitionsPart(): XmlPart<wml.Styles>;
  /** The resolver and the numbering definitions built again over the parts the package now holds. */
  refresh(): Promise<unknown>;
}

/** The main document part as the import reads and writes it. */
export interface FragmentMainPart extends Part {
  getContents(): Promise<wml.Document>;
  readonly isUnmarshalled: boolean;
  readonly headerParts: readonly (Part & { readonly isUnmarshalled: boolean; getContents(): Promise<object> })[];
  readonly footerParts: readonly (Part & { readonly isUnmarshalled: boolean; getContents(): Promise<object> })[];
  footnotesPart: (XmlPart<wml.CTFootnotes>) | undefined;
  endnotesPart: (XmlPart<wml.CTEndnotes>) | undefined;
  numberingDefinitionsPart: XmlPart<wml.Numbering> | undefined;
  styleDefinitionsPart: XmlPart<wml.Styles> | undefined;
}

export interface FragmentOptions {
  /** The part the content goes into: the main document part, a header or a footer. Its relationships take the copied parts. */
  target: Part;
  /** Word's `w:matchSrc`: the content's styles as direct formatting (the styles are copied all the same). */
  keepSourceFormatting?: boolean;
  /** The DOM preprocessor the source package is loaded with (the document's MCE one, as `contentOf`'s). */
  preprocess?: (doc: Document) => void;
}

/**
 * What an import added to the document, for a caller that binds again and removes what the run before made once nothing uses
 * it (the editor's refresh): the `w:num` and `w:abstractNum` ids, the style ids copied, and the part names added (pictures and
 * other targets, a notes, numbering or styles part made, an altChunk's part). The pass itself keeps no record between runs.
 */
export interface FragmentCreated {
  numIds: string[];
  abstractNumIds: string[];
  styleIds: string[];
  parts: string[];
}

/** The content, ready to place, what was left out, counted by kind, and what was added to the document. */
export interface ImportedFragment {
  content: Element[];
  dropped: Record<string, number>;
  created: FragmentCreated;
}

/** The names of a package's parts, to tell afterwards which an import added. */
export function partNamesOf(pkg: { parts: { values(): IterableIterator<Part> } } | undefined): Set<string> {
  return new Set(pkg ? [...pkg.parts.values()].map((p) => p.partName.name) : []);
}

type Obj = Record<string, unknown>;
interface Wrapper { name: { localPart: string; namespaceURI?: string }; value: unknown }

function isWrapper(v: unknown): v is Wrapper {
  const w = v as Partial<Wrapper> | null;
  return !!w && typeof w === 'object' && typeof w.name === 'object' && w.name !== null && typeof (w.name as { localPart?: unknown }).localPart === 'string' && 'value' in w;
}

function isPlain(v: unknown): v is Obj {
  return !!v && typeof v === 'object' && (v as { nodeType?: unknown }).nodeType === undefined;
}

/** Every typed object under `root`, with the element name its wrapper gives it (a DOM node, and what it holds, passed over). */
function visit(root: unknown, fn: (value: Obj, name: string | undefined) => void, name?: string): void {
  if (!isPlain(root)) return;
  if (Array.isArray(root)) { for (const item of root) visit(item, fn); return; }
  if (isWrapper(root)) { visit(root.value, fn, root.name.localPart); return; }
  fn(root, name);
  for (const value of Object.values(root)) if (typeof value === 'object') visit(value, fn);
}

/** Removes, from every list under `root`, each item `drop` names (an element by its wrapper's name and value); the count removed. */
function removeWhere(root: unknown, drop: (name: string | undefined, value: unknown) => boolean): number {
  let removed = 0;
  const walk = (v: unknown): void => {
    if (!isPlain(v)) return;
    if (Array.isArray(v)) {
      for (let i = v.length - 1; i >= 0; i--) {
        const item = v[i];
        if (drop(isWrapper(item) ? item.name.localPart : undefined, isWrapper(item) ? item.value : item)) { v.splice(i, 1); removed++; }
      }
      for (const item of v) walk(item);
      return;
    }
    if (isWrapper(v)) { walk(v.value); return; }
    for (const value of Object.values(v)) if (typeof value === 'object') walk(value);
  };
  walk(root);
  return removed;
}

function count(dropped: Record<string, number>, kind: string, n: number): void {
  if (n > 0) dropped[kind] = (dropped[kind] ?? 0) + n;
}

const COMMENT_MARKUP = new Set(['commentRangeStart', 'commentRangeEnd', 'commentReference']);

/** Comments' ranges and references, and a section break inside the content: what the import leaves out (D5, D6). */
function dropUnplaced(content: unknown, dropped: Record<string, number>): void {
  count(dropped, 'comments (their ranges and references)', removeWhere(content, (name) => name !== undefined && COMMENT_MARKUP.has(name)));
  let sections = 0;
  visit(content, (value) => {
    const pPr = (value as { pPr?: wml.PPr }).pPr;
    if (value.TYPE_NAME === 'org_docx4j_wml.P' && pPr?.sectPr !== undefined) { delete pPr.sectPr; sections++; }
  });
  count(dropped, 'section breaks inside the fragment', sections);
}

/** The notes the content references, by kind and id. */
function noteReferences(content: unknown): { footnote: Obj[]; endnote: Obj[] } {
  const out = { footnote: [] as Obj[], endnote: [] as Obj[] };
  visit(content, (value, name) => {
    if (name === 'footnoteReference') out.footnote.push(value);
    else if (name === 'endnoteReference') out.endnote.push(value);
  });
  return out;
}

const SEPARATORS = (kind: 'footnote' | 'endnote'): string => `<w:${kind}s xmlns:w="${W}">`
  + `<w:${kind} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:${kind}>`
  + `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`
  + `</w:${kind}s>`;

/** The document's notes part of a kind, made with Word's two separators where there is none. */
async function notesPartOf(main: FragmentMainPart, kind: 'footnote' | 'endnote'): Promise<XmlPart<wml.CTFootnotes | wml.CTEndnotes>> {
  const existing = kind === 'footnote' ? main.footnotesPart : main.endnotesPart;
  if (existing) { await existing.getContents(); return existing as XmlPart<wml.CTFootnotes | wml.CTEndnotes>; }
  // Imported here: parts/wml builds the Body views, so a static import would be a cycle (as ooxml.mts has it).
  const { FootnotesPart, EndnotesPart } = await import('../../parts/wml/index.mjs');
  const part = kind === 'footnote' ? new FootnotesPart() : new EndnotesPart();
  part.setXml(SEPARATORS(kind));
  main.addTargetPart(part);
  await part.getContents();
  return part as XmlPart<wml.CTFootnotes | wml.CTEndnotes>;
}

/**
 * Copies the notes the content references into the document's notes parts, each under a free id, and rewrites the references.
 * The copies are returned, for the styles, numbering and ids passes; what a note's relationships reach is copied with it.
 */
async function copyNotes(content: Element[], source: FragmentMainPart, main: FragmentMainPart, target: Part, dropped: Record<string, number>): Promise<object[]> {
  const refs = noteReferences(content);
  if (refs.footnote.length + refs.endnote.length === 0) return [];
  if (target !== main) {
    // a note's reference has no place in a header or a footer
    const n = removeWhere(content, (name) => name === 'footnoteReference' || name === 'endnoteReference');
    count(dropped, 'notes, in a header or a footer', n);
    return [];
  }
  const copies: object[] = [];
  for (const kind of ['footnote', 'endnote'] as const) {
    const references = refs[kind];
    if (references.length === 0) continue;
    const sourcePart = kind === 'footnote' ? source.footnotesPart : source.endnotesPart;
    const sourceNotes = sourcePart ? ((await sourcePart.getContents()) as { [k: string]: wml.CTFtnEdn[] | undefined })[kind] ?? [] : [];
    const part = await notesPartOf(main, kind);
    const notes = ((await part.getContents()) as unknown as { [k: string]: wml.CTFtnEdn[] });
    const list = (notes[kind] ??= []);
    let next = Math.max(0, ...list.map((n) => Number(n.id))) + 1;
    const mapped = new Map<string, number>();
    let missing = 0;
    for (const ref of references) {
      const key = String(ref.id);
      let id = mapped.get(key);
      if (id === undefined) {
        const note = sourceNotes.find((n) => String(n.id) === key);
        if (!note) { missing++; continue; }
        const copy = deepCopy(note) as wml.CTFtnEdn;
        id = next++;
        copy.id = id as never;
        const noteContent = (copy as { content?: Element[] }).content ?? [];
        dropUnplaced(noteContent, dropped);
        if (sourcePart) {
          const rewrites = await copyReferencedParts(noteContent, sourcePart, part);
          if (rewrites.size > 0) rewriteRelationshipIds(noteContent, rewrites);
        }
        list.push(copy);
        linkParents(copy, notes);
        copies.push(copy);
        mapped.set(key, id);
      }
      ref.id = id as never;
    }
    count(dropped, `${kind} references naming no ${kind}`, missing);
  }
  return copies;
}

// ------------------------------------------------------------------------------------------------ styles and numbering

const STYLE_REFERENCES = ['pStyle', 'rStyle', 'tblStyle'] as const;

/** The style ids the trees name: a paragraph's, a run's and a table's. */
function stylesUsedBy(trees: unknown[]): Set<string> {
  const used = new Set<string>();
  for (const tree of trees) {
    visit(tree, (value) => {
      for (const key of STYLE_REFERENCES) {
        const ref = value[key] as { val?: string } | undefined;
        if (ref && typeof ref === 'object' && typeof ref.val === 'string') used.add(ref.val);
      }
    });
  }
  return used;
}

/**
 * A style reference the fragment doesn't define reads as its default, so it goes, as Word 365's `insertOoxml` takes it (its
 * insertooxml-46 round, case 4: a package without its styles part, its Heading 1 paragraph read as Normal though the document
 * defines Heading 1). The count removed.
 */
function dropUndefinedStyleReferences(trees: unknown[], sourceStyles: wml.Style[]): number {
  const defined = new Set(sourceStyles.map((s) => s.styleId).filter((id): id is string => id !== undefined));
  let removed = 0;
  for (const tree of trees) {
    visit(tree, (value) => {
      for (const key of STYLE_REFERENCES) {
        const ref = value[key] as { val?: string } | undefined;
        if (ref && typeof ref === 'object' && typeof ref.val === 'string' && !defined.has(ref.val)) { delete value[key]; removed++; }
      }
    });
  }
  return removed;
}

/**
 * The source's styles to copy: each used style the document lacks, with its `basedOn`, `link` and `next` chain (each of those
 * the document lacks). A style both define is the document's, and its chain is the document's too.
 */
function stylesToCopy(used: Set<string>, sourceStyles: wml.Style[], documentIds: Set<string>): wml.Style[] {
  const byId = new Map(sourceStyles.filter((s) => s.styleId).map((s) => [s.styleId!, s]));
  const out: wml.Style[] = [];
  const taken = new Set<string>();
  const take = (id: string | undefined): void => {
    if (id === undefined || taken.has(id) || documentIds.has(id)) return;
    const style = byId.get(id);
    if (!style) return;
    taken.add(id);
    out.push(style);
    take(style.basedOn?.val);
    take(style.link?.val);
    take(style.next?.val);
  };
  for (const id of used) take(id);
  return out;
}

/** Every `w:numPr` in the trees whose `w:numId` names a list (not 0). */
function numPrsIn(trees: unknown[]): wml.PPrBase.NumPr[] {
  const out: wml.PPrBase.NumPr[] = [];
  for (const tree of trees) {
    visit(tree, (value) => {
      const numPr = value.numPr as wml.PPrBase.NumPr | undefined;
      if (numPr && typeof numPr === 'object' && numPr.numId?.val !== undefined && Number(numPr.numId.val) !== 0) out.push(numPr);
    });
  }
  return out;
}

const NUMBERING_XML = `<w:numbering xmlns:w="${W}"/>`;

/**
 * Word 365's list for a `w:numId` a fragment names and doesn't define (its insertooxml-46 round, case 4: a package without its
 * numbering part): nine decimal levels, `%1.` to `%9.`, each indented 720 twips more with a hanging 720 and a tab stop there, as
 * Word wrote it; the text model's default definition (REQ-510) in Word's own form.
 */
const WORD_DEFAULT_LIST_XML = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>`
  + [0, 1, 2, 3, 4, 5, 6, 7, 8].map((l) => `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%${l + 1}."/>`
    + `<w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="${720 * (l + 1)}"/></w:tabs><w:ind w:left="${720 * (l + 1)}" w:hanging="720"/></w:pPr></w:lvl>`).join('')
  + '</w:abstractNum></w:numbering>';

let wordDefaultList: Promise<wml.Numbering.AbstractNum> | undefined;
function wordDefaultAbstract(): Promise<wml.Numbering.AbstractNum> {
  wordDefaultList ??= (async () => ((await unmarshalNode(parseXml(WORD_DEFAULT_LIST_XML))).value as wml.Numbering).abstractNum![0]!)();
  return wordDefaultList;
}

function hexId(): string {
  return Math.floor(Math.random() * 0x7fffffff).toString(16).toUpperCase().padStart(8, '0');
}

/**
 * Each list the trees use becomes a new `w:num` over a copy of its `w:abstractNum`, as Word 365 resolves a chunk (46a): fresh
 * ids and a fresh `w:nsid`, the source's `w:lvlOverride`s kept; the `w:numPr`s rewritten. Returns the lists copied.
 */
async function copyNumbering(trees: unknown[], source: FragmentMainPart, main: FragmentMainPart, dropped: Record<string, number>, created: FragmentCreated): Promise<void> {
  const numPrs = numPrsIn(trees);
  if (numPrs.length === 0) return;
  const sourceNumbering = source.numberingDefinitionsPart ? await source.numberingDefinitionsPart.getContents() : undefined;
  let part = main.numberingDefinitionsPart;
  if (!part) {
    const { NumberingDefinitionsPart } = await import('../../parts/wml/index.mjs');
    const created = new NumberingDefinitionsPart();
    created.setXml(NUMBERING_XML);
    main.addTargetPart(created);
    part = created;
  }
  const numbering = await part.getContents();
  const abstracts = (numbering.abstractNum ??= []);
  const nums = (numbering.num ??= []);
  let nextAbstract = Math.max(-1, ...abstracts.map((a) => Number(a.abstractNumId))) + 1;
  let nextNum = Math.max(0, ...nums.map((n) => Number(n.numId))) + 1;
  const mapped = new Map<string, number | undefined>();
  for (const numPr of numPrs) {
    const key = String(numPr.numId!.val);
    if (!mapped.has(key)) {
      const num = sourceNumbering?.num?.find((n) => String(n.numId) === key);
      const abstract = num ? sourceNumbering?.abstractNum?.find((a) => String(a.abstractNumId) === String(num.abstractNumId?.val)) : undefined;
      // a list the fragment names and doesn't define: a new one over Word's default, as Word makes it (case 4), never the
      // document's list of the same number
      const abstractCopy = deepCopy(abstract ?? await wordDefaultAbstract()) as wml.Numbering.AbstractNum;
      abstractCopy.abstractNumId = nextAbstract++ as never;
      abstractCopy.nsid = { TYPE_NAME: 'org_docx4j_wml.CTLongHexNumber', val: hexId() } as never;
      abstracts.push(abstractCopy);
      linkParents(abstractCopy, numbering);
      const numCopy = (num ? deepCopy(num) : { TYPE_NAME: 'org_docx4j_wml.Numbering.Num' }) as wml.Numbering.Num;
      numCopy.numId = nextNum++ as never;
      numCopy.abstractNumId = { ...(numCopy.abstractNumId ?? { TYPE_NAME: 'org_docx4j_wml.Numbering.Num.AbstractNumId' }), val: abstractCopy.abstractNumId } as never;
      nums.push(numCopy);
      linkParents(numCopy, numbering);
      mapped.set(key, Number(numCopy.numId));
      created.numIds.push(String(numCopy.numId));
      created.abstractNumIds.push(String(abstractCopy.abstractNumId));
    }
    const id = mapped.get(key);
    if (id !== undefined) numPr.numId!.val = id as never;
  }
}

// ------------------------------------------------------------------------------------------------ ids

interface UsedIds { bookmark: Set<number>; sdt: Set<number>; docPr: Set<number> }

const DRAWINGS = new Set(['org_docx4j_dml_wordprocessingDrawing.Inline', 'org_docx4j_dml_wordprocessingDrawing.Anchor']);

function idsIn(trees: unknown[], into: UsedIds): void {
  for (const tree of trees) {
    visit(tree, (value, name) => {
      if (name === 'bookmarkStart' && value.id !== undefined) into.bookmark.add(Number(value.id));
      else if (value.TYPE_NAME === 'org_docx4j_wml.SdtPr') {
        const id = (sdtProperty(value as wml.SdtPr, 'id')?.value as { val?: unknown } | undefined)?.val;
        if (id !== undefined) into.sdt.add(Number(id));
      } else if (DRAWINGS.has(String(value.TYPE_NAME))) {
        const id = (value.docPr as { id?: unknown } | undefined)?.id;
        if (id !== undefined) into.docPr.add(Number(id));
      }
    });
  }
}

/** The document's ids, made fresh in the incoming trees only where they collide (a bookmark's start and end alike). */
function freshenIds(trees: unknown[], used: UsedIds): void {
  const incoming: UsedIds = { bookmark: new Set(), sdt: new Set(), docPr: new Set() };
  idsIn(trees, incoming);
  const nextOf = (kind: keyof UsedIds): (() => number) => {
    let next = Math.max(0, ...used[kind], ...incoming[kind]) + 1;
    return () => next++;
  };
  const next = { bookmark: nextOf('bookmark'), sdt: nextOf('sdt'), docPr: nextOf('docPr') };
  const bookmarks = new Map<number, number>();
  for (const tree of trees) {
    visit(tree, (value, name) => {
      if (name === 'bookmarkStart' && value.id !== undefined) {
        const id = Number(value.id);
        if (used.bookmark.has(id)) { const fresh = next.bookmark(); bookmarks.set(id, fresh); value.id = fresh; }
      } else if (value.TYPE_NAME === 'org_docx4j_wml.SdtPr') {
        const holder = sdtProperty(value as wml.SdtPr, 'id')?.value as { val?: unknown } | undefined;
        if (holder?.val !== undefined && used.sdt.has(Number(holder.val))) holder.val = next.sdt();
      } else if (DRAWINGS.has(String(value.TYPE_NAME))) {
        const docPr = value.docPr as { id?: unknown } | undefined;
        if (docPr?.id !== undefined && used.docPr.has(Number(docPr.id))) docPr.id = next.docPr();
      }
    });
  }
  if (bookmarks.size === 0) return;
  for (const tree of trees) {
    visit(tree, (value, name) => {
      if (name === 'bookmarkEnd' && value.id !== undefined) {
        const fresh = bookmarks.get(Number(value.id));
        if (fresh !== undefined) value.id = fresh;
      }
    });
  }
}

/** The ids the document uses: its main part's, and every header's, footer's and notes part's already read. */
async function documentIds(main: FragmentMainPart): Promise<UsedIds> {
  const used: UsedIds = { bookmark: new Set(), sdt: new Set(), docPr: new Set() };
  const trees: unknown[] = [await main.getContents()];
  for (const part of [...main.headerParts, ...main.footerParts]) if (part.isUnmarshalled) trees.push(await part.getContents());
  for (const part of [main.footnotesPart, main.endnotesPart]) if (part?.isUnmarshalled) trees.push(await part.getContents());
  idsIn(trees, used);
  return used;
}

// ------------------------------------------------------------------------------------------------ keepSourceFormatting

const PARA_RPR_EXTRAS = ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange'] as const;

/**
 * Word's `w:matchSrc` (46b): each paragraph's style chain (its `w:pStyle`'s, else the source's default paragraph style's, without
 * the document defaults) under its own `w:pPr`, the style's run properties under the mark's and every run's, a run's character
 * style a level between (its toggles combined as the resolver combines them), and `w:pStyle` and `w:rStyle` removed.
 */
function keepSourceFormatting(content: Element[], resolver: PropertyResolver): void {
  const defaults = resolver.getDocumentDefaultRPr();
  const defaultStyle = resolver.getDefaultParagraphStyleId();
  const runsOf = (value: unknown, chainR: wml.RPr): void => {
    if (!isPlain(value)) return;
    if (Array.isArray(value)) { for (const item of value) runsOf(item, chainR); return; }
    if (isWrapper(value)) { runsOf(value.value, chainR); return; }
    if (value.TYPE_NAME === 'org_docx4j_wml.P') { paragraph(value as wml.P); return; }
    if (value.TYPE_NAME === 'org_docx4j_wml.R') {
      const run = value as wml.R;
      const own = run.rPr;
      const rStyle = own?.rStyle?.val;
      const merged = deepCopy(chainR) as wml.RPr;
      if (rStyle) applyStyleLevel(resolver.getChainRPr(rStyle), merged, defaults);
      if (own) { const direct = deepCopy(own) as wml.RPr; delete direct.rStyle; applyRPr(direct, merged); if (own.rPrChange) merged.rPrChange = own.rPrChange; }
      delete merged.rStyle;
      if (isEmptyRPr(merged)) delete run.rPr; else { run.rPr = merged; linkParents(merged, run); }
    }
    for (const [key, item] of Object.entries(value)) if (key !== 'rPr' && key !== 'pPr' && typeof item === 'object') runsOf(item, chainR);
  };
  const paragraph = (p: wml.P): void => {
    const own = p.pPr;
    const styleId = own?.pStyle?.val ?? defaultStyle;
    const chainP = deepCopy(resolver.getChainPPr(styleId)) as wml.PPr;
    const chainR = deepCopy(resolver.getChainRPr(styleId)) as wml.RPr;
    const merged = chainP;
    delete merged.rPr;
    if (own) {
      const direct = deepCopy(own) as wml.PPr;
      delete direct.rPr;
      applyPPr(direct, merged);
      if (own.pPrChange) merged.pPrChange = own.pPrChange;
    }
    delete merged.pStyle;
    // the mark: the style's run properties under its own
    const mark = deepCopy(chainR) as unknown as wml.ParaRPr;
    (mark as { TYPE_NAME?: string }).TYPE_NAME = 'org_docx4j_wml.ParaRPr';
    if (own?.rPr) {
      applyRPr(own.rPr, mark);
      for (const key of PARA_RPR_EXTRAS) if ((own.rPr as Obj)[key] !== undefined) (mark as unknown as Obj)[key] = (own.rPr as Obj)[key];
    }
    delete (mark as { rStyle?: unknown }).rStyle;
    if (!isEmptyRPr(mark as never)) merged.rPr = mark;
    if (isEmptyPPr(merged) && merged.rPr === undefined) delete p.pPr; else { p.pPr = merged; linkParents(merged, p); }
    for (const item of (p as { content?: unknown[] }).content ?? []) runsOf(item, chainR);
  };
  runsOf(content, { TYPE_NAME: 'org_docx4j_wml.RPr' } as wml.RPr);
}

// ------------------------------------------------------------------------------------------------ the import

/**
 * A Flat OPC package's body content, read into the document `host` holds as Word 365 resolves an altChunk holding it (the
 * module's head). The content is returned for the caller to place; the parts, styles, lists and notes it needs are already in
 * the document. Throws where the package can't be read, or holds an element Word doesn't know (CR-026 D12, as `contentOf`).
 */
export async function importFlatOpc(ooxml: string, host: FragmentTarget, options: FragmentOptions): Promise<ImportedFragment> {
  const dropped: Record<string, number> = {};
  const created: FragmentCreated = { numIds: [], abstractNumIds: [], styleIds: [], parts: [] };
  const main = host.getMainDocumentPart();
  const before = partNamesOf(main.package);
  const { source, content } = await loadFlatOpc(ooxml, options.preprocess);
  if (content.length === 0) return { content, dropped, created };
  const sourceMain = source.getMainDocumentPart() as unknown as FragmentMainPart;
  dropUnplaced(content, dropped);
  const rewrites = await copyReferencedParts(content, sourceMain, options.target);
  if (rewrites.size > 0) rewriteRelationshipIds(content, rewrites);
  const notes = await copyNotes(content, sourceMain, main, options.target, dropped);

  // the styles the content and its notes use, before keepSourceFormatting takes the content's references away
  const sourceStyles = sourceMain.styleDefinitionsPart ? ((await sourceMain.styleDefinitionsPart.getContents()).style ?? []) : [];
  dropUndefinedStyleReferences([content, notes], sourceStyles);
  const existingStyles = main.styleDefinitionsPart ? await main.styleDefinitionsPart.getContents() : undefined;
  const idsOf = (styles: wml.Styles | undefined): Set<string> => new Set((styles?.style ?? []).map((s) => s.styleId).filter((id): id is string => id !== undefined));
  let toCopy = stylesToCopy(stylesUsedBy([content, notes]), sourceStyles, idsOf(existingStyles));
  // a styles part only where a style is copied: made with the defaults, which then count as the document's
  const styles = toCopy.length === 0 ? existingStyles : (existingStyles ?? await host.ensureStyleDefinitionsPart().getContents());
  if (styles !== existingStyles) toCopy = stylesToCopy(stylesUsedBy([content, notes]), sourceStyles, idsOf(styles));
  toCopy = toCopy.map((s) => {
    const copy = deepCopy(s) as wml.Style;
    delete copy._default;
    return copy;
  });

  if (options.keepSourceFormatting) keepSourceFormatting(content, await source.getPropertyResolver());

  await copyNumbering([content, notes, toCopy], sourceMain, main, dropped, created);
  if (toCopy.length > 0 && styles) {
    const list = (styles.style ??= []);
    for (const style of toCopy) { list.push(style); linkParents(style, styles); created.styleIds.push(style.styleId!); }
  }
  freshenIds([content, notes], await documentIds(main));
  await host.refresh();
  created.parts = [...partNamesOf(main.package)].filter((name) => !before.has(name));
  return { content, dropped, created };
}

/** The package a body is in as the import's target, where it can be one (a WordprocessingML package); else undefined. */
export function fragmentTargetOf(pkg: unknown): FragmentTarget | undefined {
  const p = pkg as Partial<FragmentTarget> | undefined;
  return p && typeof p.ensureStyleDefinitionsPart === 'function' && typeof p.refresh === 'function' && typeof p.getMainDocumentPart === 'function'
    ? p as FragmentTarget : undefined;
}

/**
 * What `insertOoxml` inserts (CR-002 open question 5, decided by Word's own run: CR-005 section 11.6). A flat OPC package's body is
 * read into the document as Word 365's `insertOoxml` reads it, its styles, lists and notes reconciled by `importFlatOpc`, where
 * the body is in a WordprocessingML package; a bare fragment, or a package inserted into a body with none, as `contentOf` reads it.
 */
export async function insertedContentOf(ooxml: string, options: { preprocess?: ((doc: Document) => void) | undefined; target?: Part | undefined; package?: unknown }): Promise<Element[]> {
  const host = fragmentTargetOf(options.package);
  if (isFlatOpc(ooxml) && host && options.target) {
    return (await importFlatOpc(ooxml, host, { target: options.target, ...(options.preprocess ? { preprocess: options.preprocess } : {}) })).content;
  }
  return contentOf(ooxml, { ...(options.preprocess ? { preprocess: options.preprocess } : {}), ...(options.target ? { target: options.target } : {}) });
}

/** A free part name beside `beside`'s: `<dir>/<stem>N<ext>`, N from 1. */
export function freePartName(pkg: { getPart(name: PartName | string): unknown }, beside: Part, stem: string, ext: string): PartName {
  const name = beside.partName.name;
  const dir = name.substring(0, name.lastIndexOf('/'));
  for (let i = 1; ; i++) {
    const candidate = PartName.of(`${dir}/${stem}${i}${ext}`);
    if (!pkg.getPart(candidate)) return candidate;
  }
}

