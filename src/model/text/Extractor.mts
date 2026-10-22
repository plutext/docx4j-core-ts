// CR-012 phase 2: the text model's walk, docx4j's org.docx4j.model.text.Extractor (its CR-046 section 2.3):
// scope(view(resolve(package))) in one pass over the unresolved tree, producing the block model. Branches of
// mc:AlternateContent are the content API's selected ones (`selectedBranchOf`, CR-010); the accepted and
// original views ask a `ResolvedView` what Accept All or Reject All would remove or join; markup reads every
// run with its revision. Nothing in the package is changed (REQ-204): every tree here was read before the
// walk began (`Texts`), privately where its part was not unmarshalled.
//
// Not yet here (CR-012 section 5, phase 5): fields computed under `fields: update` (5a, refused by
// `TextOptions`), math as LaTeX (5b: an equation is its m:t text inside the delimiters, REQ-604's fallback),
// charts and SmartArt (5c: nothing), embedded packages (5d, refused) and an altChunk's expansion (5e: not
// expanded, `altchunk.not_expanded`).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import type { Part } from '../../parts/Part.mjs';
import { AlternativeFormatInputPart } from '../../parts/BinaryPart.mjs';
import type { PropertyResolver } from '../properties/PropertyResolver.mjs';
import { Emulator, type NumberingResult, type NumberedPackageLike, type NumberingPartLike } from '../listnumbering/Emulator.mjs';
import type { NumberingState } from '../listnumbering/state.mjs';
import { type Element, runItemsOf, selectedBranchOf, segmentsOf, hiddenRunsOf, SDT_TYPES } from '../content/tree.mjs';
import type { ResolvedView } from '../content/ResolvedView.mjs';
import type { TextOptions, TextModelView } from './TextOptions.mjs';
import {
  TextBlock, TextRun, TextRow, TextCell, TextField, TextNumbering, TextRevision, TextWarning, TextComment, TextImage,
  TextResult, type TextStory, type TextRunKind, type NoteType,
} from './TextResult.mjs';
import { type TextContext, type PackageTrees, calendarText } from './TextContext.mjs';
import { normalize } from './normalize.mjs';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * Fonts whose glyphs are all one width, for the block model's `monospace` and Markdown's code. The WD's
 * definition of a monospaced font (1.3.2) is this list, docx4j's: a change goes to the specification first.
 */
const MONOSPACE = new Set([
  'courier new', 'courier', 'consolas', 'lucida console', 'lucida sans typewriter', 'menlo', 'monaco',
  'cascadia mono', 'cascadia code', 'source code pro', 'dejavu sans mono', 'liberation mono',
  'andale mono', 'fira code', 'fira mono', 'inconsolata', 'roboto mono', 'ubuntu mono', 'jetbrains mono',
  'noto sans mono', 'ocr a extended', 'letter gothic', 'pt mono', 'sf mono', 'ibm plex mono',
]);

const DECORATIVE_EXT = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}';
/** The extension Office 2016 writes an SVG picture's SVG in, beside the PNG copy `a:blip/@r:embed` names. */
const SVG_BLIP_EXTENSION = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';

/** The run items that open, divide, close or instruct a complex field: all `fieldsAtBlocks` reads of a run. */
const FIELD_ITEMS = new Set(['fldChar', 'instrText', 'delInstrText']);

/** How the view is read (docx4j `Extractor.Mode`). */
export type ExtractorMode =
  /** The view's rules: a plan for accepted and original, revisions marked for markup. */
  | 'view'
  /** A package already resolved by Accept All or Reject All (the copy route, REQ-406): no view logic. */
  | 'resolved'
  /** Content with no package to plan over: run-level revisions only, no joins or moves. */
  | 'planless';

// --- the object tree ------------------------------------------------------------------------------

type Obj = Record<string, unknown> & { TYPE_NAME?: string; PARENT?: object };

/** A list member's value: an element pair's, else the member. docx4j `unwrap`. */
function unwrap(item: unknown): unknown {
  return typeof item === 'object' && item !== null && 'value' in item && 'name' in item ? (item as Element).value : item;
}

/** A list member's element name, or ''. docx4j `name`. */
function nameOf(item: unknown): string {
  return typeof item === 'object' && item !== null && 'name' in item ? ((item as Element).name?.localPart ?? '') : '';
}

function typeOf(v: unknown): string | undefined {
  return typeof v === 'object' && v !== null ? (v as Obj).TYPE_NAME : undefined;
}

const P = 'org_docx4j_wml.P';
const R = 'org_docx4j_wml.R';
const TBL = 'org_docx4j_wml.Tbl';
const TR = 'org_docx4j_wml.Tr';
const TC = 'org_docx4j_wml.Tc';
const AC = 'org_docx4j_mce.AlternateContent';
const REVISION_TYPES = new Set(['org_docx4j_wml.RunIns', 'org_docx4j_wml.RunDel', 'org_docx4j_wml.RunTrackChange']);

const isSdt = (v: unknown): boolean => SDT_TYPES.has(typeOf(v) ?? '');
const isMath = (v: unknown): boolean => typeOf(v) === 'org_docx4j_math.CTOMath' || typeOf(v) === 'org_docx4j_math.CTOMathPara';
const on = (b: { val?: boolean } | undefined): boolean => b !== undefined && b !== null && b.val !== false;

/** The `w:ins`, `w:del`, `w:moveFrom` or `w:moveTo` a list member is, or undefined. */
function revisionHolder(item: unknown): wml.CTTrackChange | undefined {
  const v = unwrap(item);
  return REVISION_TYPES.has(typeOf(v) ?? '') ? v as wml.CTTrackChange : undefined;
}

const isInsertionName = (name: string): boolean => name === 'ins' || name === 'moveTo';

/** Run-level content a holder keeps (docx4j `runItemsOf`); not a paragraph's. */
function runItemsOfValue(v: unknown): Element[] | undefined {
  if (typeof v !== 'object' || v === null || typeOf(v) === P) return undefined;
  return runItemsOf(v);
}

/** Block-level content a container keeps (docx4j `blockItems`). */
function blockItems(v: unknown): Element[] | undefined {
  if (typeof v !== 'object' || v === null) return undefined;
  switch (typeOf(v)) {
    case 'org_docx4j_wml.Body': case 'org_docx4j_wml.Hdr': case 'org_docx4j_wml.Ftr': case 'org_docx4j_wml.CTFtnEdn':
    case 'org_docx4j_wml.Comments.Comment': case TC: case 'org_docx4j_wml.CTTxbxContent': case 'org_docx4j_wml.CTCustomXmlBlock':
    case 'org_docx4j_wml.CTCustomXmlRow': case 'org_docx4j_wml.CTCustomXmlCell':
      return ((v as { content?: Element[] }).content ??= []);
    case 'org_docx4j_wml.Document': {
      const body = (v as wml.Document).body;
      return body ? ((body as { content?: Element[] }).content ?? []) : undefined;
    }
    default: return undefined;
  }
}

/** A range marker, bookmark, permission or proofing mark: a join passes over it. */
function isMarker(item: unknown): boolean {
  const name = nameOf(item);
  return name.endsWith('RangeStart') || name.endsWith('RangeEnd') || name === 'bookmarkStart' || name === 'bookmarkEnd'
    || name === 'permStart' || name === 'permEnd' || name === 'proofErr';
}

const isBookmarkMarker = (item: unknown): boolean => nameOf(item) === 'bookmarkStart' || nameOf(item) === 'bookmarkEnd';

/** The `w:sdtPr` child of a name, or undefined. */
function sdtPrChild(pr: wml.SdtPr | undefined, local: string): Element | undefined {
  return (pr?.rPrOrAliasOrLock as Element[] | undefined)?.find((o) => nameOf(o) === local && (o as Element).name.namespaceURI === W_NS);
}

function showingPlaceholder(pr: wml.SdtPr | undefined): boolean {
  const e = sdtPrChild(pr, 'showingPlcHdr');
  return e !== undefined && on(e.value as { val?: boolean });
}

/** What a reflective walk does at each object (docx4j `Visitor`). */
interface Visitor {
  /** A typed object reached; false to go no deeper. */
  enter?(value: Obj): boolean;
  /** A list member or property value reached, before its value is entered; false to skip it. */
  item?(item: unknown): boolean;
}

/** A walk of the typed objects under `value`, in property and list order; DOM is not entered. docx4j `walk`. */
function walk(value: unknown, depth: number, visitor: Visitor): void {
  if (typeof value !== 'object' || value === null || depth > 200 || typeOf(value) === undefined) return;
  if (visitor.enter && !visitor.enter(value as Obj)) return;
  for (const key of Object.keys(value)) {
    if (key === 'PARENT' || key === 'TYPE_NAME') continue;
    const fv = (value as Obj)[key];
    if (Array.isArray(fv)) {
      for (const o of [...fv]) {
        if (typeof o !== 'object' || o === null) continue;
        if (visitor.item && !visitor.item(o)) continue;
        walk(unwrap(o), depth + 1, visitor);
      }
    } else if (typeof fv === 'object' && fv !== null) {
      if (visitor.item && !visitor.item(fv)) continue;
      walk(unwrap(fv), depth + 1, visitor);
    }
  }
}

/** The typed objects directly under `value`: its list members and object properties. docx4j `children`. */
function children(value: unknown): unknown[] {
  const out: unknown[] = [];
  if (typeof value !== 'object' || value === null || typeOf(value) === undefined) return out;
  for (const key of Object.keys(value)) {
    if (key === 'PARENT' || key === 'TYPE_NAME') continue;
    const fv = (value as Obj)[key];
    if (Array.isArray(fv)) out.push(...fv.filter((o) => typeof o === 'object' && o !== null));
    else if (typeof fv === 'object' && fv !== null) out.push(fv);
  }
  return out;
}

function isMonospace(font: string | undefined): boolean {
  return font !== undefined && MONOSPACE.has(font.trim().toLowerCase());
}

function asciiFont(rPr: wml.RPr | undefined): string | undefined {
  return rPr?.rFonts?.ascii;
}

function revisionOf(tc: { author?: string; date?: unknown }, type: 'INSERTION' | 'DELETION'): TextRevision {
  return new TextRevision(type, tc.author, calendarText(tc.date));
}

// --- where Office JS's paragraph.text puts a run's text (CR-012 section 4.4) -------------------------

interface OfficeSpan { offset: number; length: number }

/**
 * Each item's place in `Paragraph.text` (CR-012 D2, section 4.4): the content API's own reader since phase 4
 * (`segmentsOf`, Office JS's characters, hidden runs left out by the same rule), so an item that adds a
 * character there has its offset, and one it leaves out (hidden text, a comment reference) has none.
 */
function officeSpans(p: wml.P, resolver: PropertyResolver | undefined): Map<unknown, OfficeSpan> {
  const out = new Map<unknown, OfficeSpan>();
  for (const seg of segmentsOf(p, { hidden: hiddenRunsOf(resolver, p) })) {
    if (!seg.math) out.set(seg.item, { offset: seg.start, length: seg.end - seg.start });
  }
  return out;
}

// --- the walk's state -----------------------------------------------------------------------------

/** A bookmark open in the walk. */
interface Bookmark { name: string; text: string; paragraphEnded: boolean }

/** An open complex field. */
class FieldState {
  instr = '';
  result = '';
  separated = false;
  /** `w:fldLock` on its beginning: never recomputed (REQ-506). */
  locked = false;
  /** Its result computed under `fields: update` (REQ-506): the cached result's text is not read. */
  computed: string | undefined;
  /** A HYPERLINK field's target, once its code is read (REQ-703). */
  link: string | undefined;
  /** A table of contents under `toc: false` (D12): nothing of it is read. */
  toc = false;
  /** Where its code's runs begin, for a table of contents' to be taken out again. */
  codeBlock: TextBlock | undefined;
  codeStart = 0;
  /** The paragraph whose merge floor its code raised, and the floor before. */
  floorParagraph: ParagraphState | undefined;
  previousFloor = 0;

  /** The field as a block after the one it began in starts inside it: its code's block and paragraph left behind. */
  copy(): FieldState {
    const f = new FieldState();
    f.instr = this.instr;
    f.result = this.result;
    f.separated = this.separated;
    f.locked = this.locked;
    f.computed = this.computed;
    f.link = this.link;
    f.toc = this.toc;
    return f;
  }
}

/** The complex fields open where a container's blocks begin, by the block's index (`Extractor.fieldsAtBlocks`). */
export type BlockStarts = ReadonlyMap<number, readonly FieldState[]>;

/** The state of a run-level walk: the revision and the control the runs are in. */
class RunState {
  constructor(
    /** markup: the innermost revision; else undefined. */
    readonly revision: TextRevision | undefined,
    /** accepted, original: inside a revision holder the view keeps (its runs are text whatever their elements). */
    readonly kept: boolean,
    readonly placeholder: boolean,
  ) {}

  withRevision(r: TextRevision | undefined, kept: boolean): RunState {
    return new RunState(r, kept, this.placeholder);
  }

  withPlaceholder(): RunState {
    return new RunState(this.revision, this.kept, true);
  }
}

const PLAIN = new RunState(undefined, false, false);

/** Where a run's text came from, for its `source` and `officeOffset`. */
interface ItemSource { run: wml.R; item: Element }

/** What follows its anchoring paragraph, or the table its anchor is in (REQ-511, REQ-518): a text box, legacy WordArt. */
interface Anchored {
  prefix: string;
  emit(x: Extractor, anchoring: Story, sink: TextBlock[]): void;
}

/** The walk of one story. */
class Story {
  /** Open complex fields; [0] is the innermost (docx4j's deque head). */
  readonly fields: FieldState[] = [];
  /** What was anchored since the last flush, to follow its anchor (or its table). */
  readonly pending: Anchored[] = [];
  tableDepth = 0;
  /** The sections ended so far (the main story only). */
  sections: (wml.SectPr | undefined)[] | undefined;
  noteType: NoteType | undefined;
  noteNumber = 0;
  noteId: string | undefined;
  private numberingState: NumberingState | undefined;

  constructor(private readonly x: Extractor, readonly kind: TextStory, readonly prefix: string, readonly part: Part | undefined) {}

  /** The resolver the walk reads styles with, for `officeSpans`' hidden rule. */
  get resolver(): PropertyResolver | undefined {
    return this.x.resolver;
  }

  numbering(): NumberingState {
    this.numberingState ??= this.part === undefined ? this.x.context.numberingStates.main() : this.x.context.numberingStates.forPart(this.part);
    return this.numberingState;
  }
}

/** A paragraph being read: the block its runs go to, its address, what it anchored so far. */
class ParagraphState {
  boxes = 0;
  charts = 0;
  smartArts = 0;
  objects = 0;
  /** Whether every run with text so far is in a monospaced font; undefined before the first. */
  monospace: boolean | undefined;
  /** The formatting of the run being read (a run carrying only it), or undefined outside a run. */
  format: TextRun | undefined;
  /** The target of the `w:hyperlink` being read. */
  link: string | undefined;
  /** An equation being written, and whether it is a display one. */
  math = false;
  displayMath = false;
  /** Whether a table of contents left out (D12) is in the paragraph. */
  toc = false;
  /** The first of the block's runs a run may be merged into: a field's code starts a run of its own where it may be taken out. */
  mergeFloor = 0;
  private office: Map<unknown, OfficeSpan> | undefined;

  constructor(readonly block: TextBlock, readonly p: wml.P, readonly address: string, readonly story: Story) {}

  /** Where an item's text is in this paragraph's `Paragraph.text`. */
  officeSpan(item: Element): OfficeSpan | undefined {
    this.office ??= officeSpans(this.p, this.story.resolver);
    return this.office.get(item);
  }
}

/**
 * The text model's walk (CR-046 section 2.3): `scope(view(resolve(package)))` in one pass over the
 * unresolved tree, producing the block model. docx4j `Extractor`, over trees read before it starts.
 */
export class Extractor {
  readonly view: TextModelView;
  readonly mode: ExtractorMode;
  /** What a view removes and joins; undefined in markup, and in the resolved and planless modes. */
  readonly plan: ResolvedView | undefined;
  readonly resolver: PropertyResolver | undefined;
  readonly warnings: TextWarning[] = [];
  private readonly warned = new Set<string>();
  /** What every address this extraction makes starts with. */
  root = '';
  /** The run properties a paragraph style and character style give, by their ids. */
  private readonly styleRPrCache = new Map<string, wml.RPr | undefined>();
  /** Bookmarks open in the walk, by id (REQ-506's REF, phase 5a). */
  private readonly openBookmarks = new Map<string, Bookmark>();
  /** `fieldsAtBlocks`' walk: paragraphs read for their field characters and instructions alone, nothing emitted. */
  private fieldsOnly = false;

  constructor(
    readonly trees: PackageTrees,
    readonly options: TextOptions,
    mode: ExtractorMode,
    readonly context: TextContext,
  ) {
    this.view = options.view;
    this.mode = trees.pkg === undefined && mode === 'view' ? 'planless' : mode;
    this.plan = this.mode === 'view' && this.view !== 'markup' ? context.plan(this.view === 'accepted') : undefined;
    this.resolver = trees.resolver;
  }

  // --- the package ------------------------------------------------------------------------------

  extract(): TextResult {
    const out: TextBlock[] = [];
    const comments = new Map<string, TextComment>();
    const metadata = this.metadata();
    const main = this.trees.main;
    const body = main?.document.body as { content?: Element[]; sectPr?: wml.SectPr } | undefined;
    if (!main || !body) return new TextResult(out, comments, this.warnings, metadata, this.options);

    // the main story, each text box after its anchor; sections recorded as they end
    const story = new Story(this, 'MAIN', `${this.root}body`, main.part);
    story.sections = [];
    const mainBlocks: TextBlock[] = [];
    this.blocks(body.content ?? [], `${this.root}body`, story, mainBlocks);
    story.sections.push(body.sectPr);
    this.context.mainRead = true;

    const headers = this.options.includes.has('headers');
    const footers = this.options.includes.has('footers');
    if (headers || footers) this.sectionsWithHeaders(story.sections, mainBlocks, headers, footers, out);
    else out.push(...mainBlocks);

    // the notes, in reference order (REQ-501, REQ-503)
    if (this.options.notes) {
      this.notes('footnote', out);
      this.notes('endnote', out);
    }

    // the glossary, last (REQ-501, REQ-516)
    if (this.options.includes.has('glossary')) this.glossary(out);

    // comments, rendered at their references (REQ-504)
    if (this.commentsIncluded()) this.comments(comments);

    return new TextResult(out, comments, this.warnings, metadata, this.options);
  }

  /** The result's metadata: `fields_update_on_open` and the document properties (REQ-801, REQ-515). */
  private metadata(): Record<string, unknown> {
    const metadata: Record<string, unknown> = { [TextResult.FIELDS_UPDATE_ON_OPEN]: this.fieldsUpdateOnOpen() };
    Object.assign(metadata, this.trees.properties ?? {});
    return metadata;
  }

  private fieldsUpdateOnOpen(): boolean {
    return on(this.trees.settings?.updateFields);
  }

  commentsIncluded(): boolean {
    return this.view === 'markup' || this.options.comments;
  }

  /** A part's or object's content as one story (`Texts.of(part)`, `Texts.of(pkg, content)`). */
  extractContent(content: unknown, kind: TextStory, prefix: string, part: Part | undefined): TextResult {
    const out: TextBlock[] = [];
    const story = new Story(this, kind, prefix, part);
    const v = unwrap(content);
    const items = blockItems(v);
    if (items !== undefined && typeOf(v) !== P) this.blocks(items, prefix, story, out);
    else this.blocks([content as Element], prefix, story, out);
    if (kind === 'MAIN' && prefix === 'body') this.context.mainRead = true;
    const metadata: Record<string, unknown> = {};
    if (this.trees.pkg !== undefined) metadata[TextResult.FIELDS_UPDATE_ON_OPEN] = this.fieldsUpdateOnOpen();
    return new TextResult(out, new Map(), this.warnings, metadata, this.options);
  }

  /** The notes of one kind, in the context's reference order (`Texts.of(notesPart)`). */
  extractNotes(type: NoteType): TextResult {
    const out: TextBlock[] = [];
    this.notes(type, out);
    return new TextResult(out, new Map(), this.warnings, { [TextResult.FIELDS_UPDATE_ON_OPEN]: this.fieldsUpdateOnOpen() }, this.options);
  }

  /** The comments (`Texts.of(commentsPart)`): what the view includes, in the comments part's order. */
  extractComments(): TextResult {
    const comments = new Map<string, TextComment>();
    this.comments(comments);
    return new TextResult([], comments, this.warnings, { [TextResult.FIELDS_UPDATE_ON_OPEN]: this.fieldsUpdateOnOpen() }, this.options);
  }

  // --- the glossary (REQ-516) -------------------------------------------------------------------

  private glossary(out: TextBlock[]): void {
    const g = this.trees.glossary;
    const parts = g?.tree.docParts?.docPart ?? [];
    for (let i = 0; i < parts.length; i++) {
      const body = parts[i]!.docPartBody as { content?: Element[] } | undefined;
      if (!body) continue;
      const prefix = `${this.root}glossary:${i}`;
      this.blocks(body.content ?? [], prefix, new Story(this, 'GLOSSARY', prefix, g!.part), out);
    }
  }

  // --- block level ------------------------------------------------------------------------------

  /** Block-level content into `sink`: paragraphs (joined in a view), tables, and what is transparent. */
  blocks(items: Element[] | undefined, prefix: string, story: Story, sink: TextBlock[], state: RunState = PLAIN): void {
    if (!items) return;
    for (let i = 0; i < items.length; i++) i = this.block(items, i, prefix, story, sink, state);
  }

  /** The block at `items[i]` into `sink`, with the paragraphs a view joins to it; the index of the last item it took. */
  private block(items: Element[], i: number, prefix: string, story: Story, sink: TextBlock[], state: RunState): number {
    const item = items[i]!;
    if (this.removes(item)) return i;
    const v = unwrap(item);
    const address = `${prefix}/${i}`;
    const tn = typeOf(v);
    if (tn === P) {
      const chain: wml.P[] = [v as wml.P];
      const addresses = [address];
      let j = i;
      while (this.joinsNext(chain[chain.length - 1]!)) {
        const k = this.nextParagraph(items, j);
        if (k < 0) break;                                            // no paragraph next: only the mark goes
        chain.push(unwrap(items[k]) as wml.P);
        addresses.push(`${prefix}/${k}`);
        j = k;
      }
      this.paragraph(chain, addresses, story, sink, state);
      return j;
    }
    if (tn === TBL) {
      this.table(v as wml.Tbl, address, story, sink, state);
    } else if (isSdt(v) && this.unwraps(item)) {
      // an inserted control rejected: its content in its place, its properties gone
      this.blocks(sdtContentOf(v), address, story, sink, state);
    } else if (isSdt(v)) {
      const sdt = v as wml.SdtBlock;
      const bound = this.bound(sdt, address);
      if (bound !== undefined) {
        // a plain-text control bound at block level: its bound value as one paragraph (REQ-509)
        const block = paragraphBlock(story.kind, address, bound);
        block.section = story.sections?.length ?? 0;
        block.source = sdt;
        sink.push(block);
        this.flushIfOutsideTable(story, sink);
        return i;
      }
      const placeholder = showingPlaceholder(sdt.sdtPr);
      if (placeholder && !this.options.placeholders) return i;
      this.blocks(sdtContentOf(v), address, story, sink, placeholder ? state.withPlaceholder() : state);
    } else if (tn === 'org_docx4j_wml.CTAltChunk') {
      this.altChunk(v as wml.CTAltChunk, address, story);
    } else if (isBookmarkMarker(item)) {
      this.bookmark(item);
    } else if (tn === AC) {
      const ac = v as mce.AlternateContent;
      this.compareBranches(ac, address);
      this.blocks(selectedBranchOf(ac), address, story, sink, state);
    } else if (revisionHolder(item) !== undefined) {
      this.blocks(runItemsOfValue(v), address, story, sink, this.holderState(item, state));
    } else {
      const inner = blockItems(v);
      if (inner) this.blocks(inner, address, story, sink, state);
    }
    return i;
  }

  // --- some of a container's blocks (CR-012 section 17.13) -----------------------------------------

  /**
   * Some of a container's blocks, by their indices, each as `extractContent` reads it in the whole: a
   * paragraph the view joins to one before it gives nothing (its text is in that one's block), and each
   * block starts inside the complex fields open where it begins (`starts`, from `fieldsAtBlocks`), so that
   * a table of contents left out (D12) is left out here too. Each block is read as a story of its own.
   */
  extractBlocks(content: unknown, kind: TextStory, prefix: string, part: Part | undefined, indices: readonly number[], starts: BlockStarts): TextResult {
    const out: TextBlock[] = [];
    const items = blockItems(unwrap(content)) ?? [];
    for (const i of [...new Set(indices)].sort((a, b) => a - b)) {
      if (!(i >= 0 && i < items.length) || this.joinedFromBefore(items, i)) continue;
      const story = new Story(this, kind, prefix, part);
      for (const f of starts.get(i) ?? []) story.fields.push(f.copy());
      this.block(items, i, prefix, story, out, PLAIN);
    }
    return new TextResult(out, new Map(), this.warnings, { [TextResult.FIELDS_UPDATE_ON_OPEN]: this.fieldsUpdateOnOpen() }, this.options);
  }

  /**
   * The complex fields open where each of a container's blocks begins, for `extractBlocks`: a walk of the
   * blocks reading field characters and instructions alone (`fieldsOnly`), as the whole walk reads them.
   * Blocks with none open are not in the map. Its own context and warnings, so the reading leaves no trace.
   */
  static fieldsAtBlocks(trees: PackageTrees, options: TextOptions, mode: ExtractorMode, context: TextContext, content: unknown, kind: TextStory, prefix: string, part: Part | undefined): BlockStarts {
    const x = new Extractor(trees, options, mode, context.sibling());
    x.fieldsOnly = true;
    const starts = new Map<number, FieldState[]>();
    const items = blockItems(unwrap(content)) ?? [];
    const story = new Story(x, kind, prefix, part);
    const sink: TextBlock[] = [];
    for (let i = 0; i < items.length; i++) {
      if (story.fields.length > 0) starts.set(i, story.fields.map((f) => f.copy()));
      i = x.block(items, i, prefix, story, sink, PLAIN);
      sink.length = 0;
    }
    return starts;
  }

  /** Whether the view joins the paragraph at `items[i]` to one before it: the item before it, past markers and what the view removes, is a paragraph whose mark goes. */
  private joinedFromBefore(items: Element[], i: number): boolean {
    if (typeOf(unwrap(items[i])) !== P) return false;
    for (let k = i - 1; k >= 0; k--) {
      const item = items[k];
      if (isMarker(item) || this.removes(item)) continue;
      const v = unwrap(item);
      return typeOf(v) === P && this.joinsNext(v as wml.P);
    }
    return false;
  }

  /** The index of the paragraph a paragraph joins: the next in the list past markers and what the view removes, or -1. */
  private nextParagraph(items: Element[], from: number): number {
    for (let k = from + 1; k < items.length; k++) {
      const item = items[k];
      if (isMarker(item) || this.removes(item)) continue;
      return typeOf(unwrap(item)) === P ? k : -1;
    }
    return -1;
  }

  private paragraph(chain: wml.P[], addresses: string[], story: Story, sink: TextBlock[], state: RunState): void {
    const first = chain[0]!;
    const last = chain[chain.length - 1]!;
    const block = new TextBlock('PARAGRAPH', story.kind, addresses[0]!);
    if (this.fieldsOnly) {
      for (let c = 0; c < chain.length; c++) this.runs(chain[c]!.content as Element[] | undefined, new ParagraphState(block, chain[c]!, addresses[c]!, story), state);
      return;
    }
    if (chain.length > 1) block.joined = addresses.slice(1);
    block.paraId = first.paraId;
    block.source = first;
    // a joined paragraph keeps the last one's properties (CR-043's join)
    block.style = last.pPr?.pStyle?.val;
    block.noteType = story.noteType;
    block.noteNumber = story.noteNumber;
    block.noteId = story.noteId;
    let monospace: boolean | undefined;
    let toc = false;
    for (let c = 0; c < chain.length; c++) {
      const p = chain[c]!;
      block.revisions += markRevisions(p);
      const para = new ParagraphState(block, p, addresses[c]!, story);
      const startsInToc = inToc(story);
      this.runs(p.content as Element[] | undefined, para, state);
      if (startsInToc || para.toc || inToc(story)) toc = true;
      if (para.monospace !== undefined) monospace = monospace === undefined ? para.monospace : monospace && para.monospace;
    }
    for (const b of this.openBookmarks.values()) b.paragraphEnded = true;
    block.monospace = monospace === true;
    this.properties(block, last);
    const markHidden = this.isHiddenMark(last.pPr);
    const anyVisible = block.runs.some((r) => !r.hidden);
    const hiddenParagraph = markHidden && !anyVisible;
    if (hiddenParagraph && !this.options.hidden) {
      // runs and mark hidden: no line at all (REQ-507)
      block.runs.length = 0;
      this.flushIfOutsideTable(story, sink);
      this.endSection(last, story);
      return;
    }
    block.hidden = hiddenParagraph;
    if (!this.options.hidden) {
      for (let k = block.runs.length - 1; k >= 0; k--) if (block.runs[k]!.hidden) block.runs.splice(k, 1);
    }
    if (toc && block.runs.length === 0) {
      // a table of contents' paragraph, under toc: false (D12): no line at all
      this.flushIfOutsideTable(story, sink);
      this.endSection(last, story);
      return;
    }
    // one count per paragraph: the label (list_labels) and the numbering's value come from it
    if (this.trees.pkg !== undefined && (this.options.listLabels || block.numbering !== undefined)) {
      const counted = this.count(last, story);
      if (this.options.listLabels) block.listLabel = listLabel(counted);
      if (block.numbering !== undefined && counted !== undefined && !counted.notNumbered) block.numbering.value = counted.count;
    }
    block.section = story.sections?.length ?? 0;
    sink.push(block);
    this.flushIfOutsideTable(story, sink);
    this.endSection(last, story);
  }

  private endSection(last: wml.P, story: Story): void {
    if (story.sections && story.tableDepth === 0 && last.pPr?.sectPr) story.sections.push(last.pPr.sectPr);
  }

  private flushIfOutsideTable(story: Story, sink: TextBlock[]): void {
    if (story.tableDepth > 0 || story.pending.length === 0) return;
    const items = story.pending.splice(0);
    for (const item of items) {
      const blocks: TextBlock[] = [];
      item.emit(this, story, blocks);
      const section = story.sections?.length ?? 0;
      for (const b of blocks) b.section = section;
      sink.push(...blocks);
    }
  }

  // --- altChunk (REQ-513) -----------------------------------------------------------------------

  /** An `altChunk`: not expanded until CR-012 phase 5e, which raises `altchunk.not_expanded` with its content type. */
  private altChunk(chunk: wml.CTAltChunk, address: string, story: Story): void {
    const part = related(story.part, chunk.id);
    if (!(part instanceof AlternativeFormatInputPart)) return;
    const ct = part.contentType ?? '';
    const semi = ct.indexOf(';');
    this.warn(new TextWarning(TextWarning.ALTCHUNK_NOT_EXPANDED, address, undefined, (semi >= 0 ? ct.slice(0, semi) : ct).trim()));
  }

  // --- bound values (REQ-509) -------------------------------------------------------------------

  /**
   * A plain-text control's bound value where its binding resolves and is what Word shows: not where the
   * value is empty and the control shows its placeholder, nor under `bindings: surface`. Undefined to read
   * the surface. Raises `sdt.binding_differs` where the two differ.
   */
  private bound(sdt: { sdtPr?: wml.SdtPr; sdtContent?: unknown }, address: string): string | undefined {
    const pr = sdt.sdtPr;
    if (this.trees.pkg === undefined || !pr || sdtPrChild(pr, 'text') === undefined) return undefined;
    const binding = sdtPrChild(pr, 'dataBinding')?.value as wml.CTDataBinding | undefined;
    if (!binding) return undefined;
    const value = this.trees.boundValue(binding);
    if (value === undefined) return undefined;
    const placeholder = showingPlaceholder(pr);
    if (value === '' && placeholder) return undefined;
    const surface = placeholder ? undefined : this.surfaceText(sdt);
    if (value !== surface) this.warn(new TextWarning(TextWarning.SDT_BINDING_DIFFERS, address));
    return this.options.bindings === 'surface' ? undefined : value;
  }

  /**
   * A control's surface in the view: its `w:t`, tabs and breaks, what the view removes left out, and of an
   * mc:AlternateContent the selected branch only, as the reading takes it (CR-013: every branch's text
   * made a surface no binding could match, and a false `sdt.binding_differs`).
   */
  private surfaceText(sdt: { sdtContent?: unknown }): string {
    let out = '';
    // every item, a branch's own included (a branch in a w:r holds w:t directly), goes through item()
    const item = (o: unknown): boolean => {
      if (this.removes(o)) return false;
      const name = nameOf(o);
      const v = unwrap(o);
      if (typeOf(v) === AC) {
        visit(selectedBranchOf(v as mce.AlternateContent));
        return false;
      }
      if (name === 't' && typeOf(v) === 'org_docx4j_wml.Text') out += String((v as wml.Text).value ?? '');
      else if (name === 'tab') out += '\t';
      else if (name === 'br' || name === 'cr') out += '\n';
      return name !== 'instrText';
    };
    const visit = (items: unknown[] | undefined): void => {
      for (const o of items ?? []) if (item(o)) walk(unwrap(o), 0, { item });
    };
    visit(sdtContentOf(sdt));
    return out;
  }

  // --- bookmarks, for REF under fields: update (REQ-506) -----------------------------------------

  private bookmark(item: unknown): void {
    if (this.options.fields !== 'update') return;                     // phase 5a
    const v = unwrap(item) as { id?: unknown; name?: string };
    if (nameOf(item) === 'bookmarkStart') {
      if (v.id !== undefined && v.name !== undefined) this.openBookmarks.set(String(v.id), { name: v.name, text: '', paragraphEnded: false });
    } else if (v.id !== undefined) {
      const b = this.openBookmarks.get(String(v.id));
      this.openBookmarks.delete(String(v.id));
      if (b) this.context.bookmarks.set(b.name.toLowerCase(), b.text);
    }
  }

  private properties(block: TextBlock, p: wml.P): void {
    const own = p.pPr;
    block.bottomBorder = own?.pBdr?.bottom !== undefined;
    let effective: wml.PPr | undefined = own;
    if (this.resolver) {
      try {
        effective = this.resolver.getEffectivePPr(own);
      } catch {
        // not resolved: the paragraph's own properties
      }
    }
    if (!effective) return;
    const ind = effective.ind;
    if (ind) {
      block.indentLeft = twips(ind.left ?? ind.start);
      block.indentHanging = twips(ind.hanging);
      block.indentFirstLine = twips(ind.firstLine);
    }
    if (effective.jc?.val !== undefined) block.alignment = String(effective.jc.val);
    const level = effective.outlineLvl?.val;
    if (level !== undefined && level !== null) {
      const n = Number(level);
      if (n >= 0 && n <= 8) block.outlineLevel = n;                    // 9 is body text
    }
    if (effective.contextualSpacing !== undefined) block.contextualSpacing = on(effective.contextualSpacing);
    const directNumId = own?.numPr?.numId !== undefined;
    block.numbering = this.numbering(effective.numPr, directNumId, own?.pStyle?.val);
  }

  /**
   * A paragraph's numbering from its effective `w:numPr` and the numbering definitions; undefined for none,
   * for `w:numId` 0, and where Word gives the paragraph no number: a `w:numId` its style brought, whose
   * level belongs to another paragraph style (`Emulator.levelStyleId`, REQ-510).
   */
  private numbering(numPr: wml.PPrBase.NumPr | undefined, directNumId: boolean, pStyle: string | undefined): TextNumbering | undefined {
    const numIdValue = numPr?.numId?.val;
    if (numIdValue === undefined || numIdValue === null) return undefined;
    const numId = String(numIdValue);
    if (Number(numIdValue) === 0) return undefined;
    const ilvlValue = numPr?.ilvl?.val;
    const ilvl = ilvlValue === undefined || ilvlValue === null ? 0 : Number(ilvlValue);
    let format: string | undefined;
    let text: string | undefined;
    let abstractNumId: string | undefined;
    let start = 1;
    const part = this.trees.numberingPart;
    const definitions = this.trees.numbering;
    try {
      if (part && !directNumId && this.linkedElsewhere(part, numId, String(ilvl), pStyle)) return undefined;
      const list = definitions?.list(numId);
      abstractNumId = list?.abstractDefinition?.id;
      // the instance's level, its w:lvlOverride applied
      const level = list?.level(ilvl) ?? list?.abstractDefinition?.levels.get(String(ilvl));
      if (level) {
        format = level.numFmt === undefined ? (level.isBullet ? 'bullet' : undefined) : String(level.numFmt);
        text = level.lvlText;
        start = level.startValue + 1;
      }
    } catch {
      // the definitions could not be read
    }
    const numbering = new TextNumbering(numId, ilvl, format, text, start);
    numbering.abstractNumId = abstractNumId;
    return numbering;
  }

  /** Whether the level belongs to a paragraph style that is neither the paragraph's nor one its style is based on. */
  private linkedElsewhere(part: NumberingPartLike, numId: string, ilvl: string, pStyle: string | undefined): boolean {
    const linked = Emulator.levelStyleId(part, numId, ilvl);
    if (linked === undefined || linked === '') return false;
    // a paragraph naming no style is in the default paragraph style, which can own a level
    let id = pStyle ?? this.resolver?.getDefaultParagraphStyleId();
    const seen = new Set<string>();
    while (id !== undefined && !seen.has(id)) {
      seen.add(id);
      if (linked === id) return false;
      id = this.resolver?.getStyle(id)?.basedOn?.val;
    }
    return true;
  }

  /** The paragraph counted in its story's numbering state; undefined where it is not numbered. */
  private count(p: wml.P, story: Story): NumberingResult | undefined {
    const pkg = this.trees.pkg as NumberedPackageLike | undefined;
    if (!pkg || !this.trees.numberingPart) return undefined;
    try {
      return Emulator.getNumber(pkg, p.pPr, story.numbering());
    } catch {
      return undefined;
    }
  }

  // --- tables -----------------------------------------------------------------------------------

  private table(tbl: wml.Tbl, address: string, story: Story, sink: TextBlock[], state: RunState): void {
    const block = new TextBlock('TABLE', story.kind, address);
    block.section = story.sections?.length ?? 0;
    block.source = tbl;
    story.tableDepth++;
    try {
      this.rows(tbl.content as Element[] | undefined, address, story, block, state);
    } finally {
      story.tableDepth--;
    }
    if (block.rows.length > 0) sink.push(block);                     // a table left with no rows goes (CR-043)
    this.flushIfOutsideTable(story, sink);
  }

  private rows(items: Element[] | undefined, prefix: string, story: Story, table: TextBlock, state: RunState): void {
    if (!items) return;
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (this.removes(item)) continue;
      const v = unwrap(item);
      const address = `${prefix}/${i}`;
      if (typeOf(v) === TR) {
        const tr = v as wml.Tr;
        const row = new TextRow(address);
        const trPr = tr.trPr;
        let rowState = state;
        if (trPr) {
          for (const e of (trPr.cnfStyleOrDivIdOrGridBefore ?? []) as Element[]) {
            if (nameOf(e) === 'tblHeader') row.header = on(e.value as { val?: boolean });
          }
          if (this.view === 'markup' && this.mode === 'view') {
            if (trPr.ins) rowState = state.withRevision(revisionOf(trPr.ins, 'INSERTION'), false);
            if (trPr.del) rowState = state.withRevision(revisionOf(trPr.del, 'DELETION'), false);
          }
        }
        this.cells(tr.content as Element[] | undefined, address, story, row, rowState);
        table.rows.push(row);
      } else if (isSdt(v)) {
        this.rows(sdtContentOf(v), address, story, table, state);
      } else if (typeOf(v) === AC) {
        // a row-level mc:AlternateContent: its selected branch's rows, addressed from the element (CR-013)
        const ac = v as mce.AlternateContent;
        this.compareBranches(ac, address);
        this.rows(selectedBranchOf(ac), address, story, table, state);
      } else {
        const inner = blockItems(v);
        if (inner && typeOf(v) !== TC) this.rows(inner, address, story, table, state);
      }
    }
  }

  private cells(items: Element[] | undefined, prefix: string, story: Story, row: TextRow, state: RunState): void {
    if (!items) return;
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (this.removes(item)) continue;
      const v = unwrap(item);
      const address = `${prefix}/${i}`;
      if (typeOf(v) === TC) {
        const tc = v as wml.Tc;
        const cell = new TextCell(address);
        const tcPr = tc.tcPr;
        let cellState = state;
        if (tcPr) {
          const span = tcPr.gridSpan?.val;
          if (span !== undefined && span !== null) cell.gridSpan = Number(span);
          if (tcPr.vMerge !== undefined) {
            if (tcPr.vMerge.val === 'restart') cell.mergeRestart = true;
            else cell.mergeContinue = true;
          }
          if (this.view === 'markup' && this.mode === 'view') {
            if (tcPr.cellIns) cellState = state.withRevision(revisionOf(tcPr.cellIns, 'INSERTION'), false);
            if (tcPr.cellDel) cellState = state.withRevision(revisionOf(tcPr.cellDel, 'DELETION'), false);
          }
        }
        this.blocks(tc.content as Element[] | undefined, address, story, cell.blocks, cellState);
        row.cells.push(cell);
      } else if (isSdt(v)) {
        this.cells(sdtContentOf(v), address, story, row, state);
      } else if (typeOf(v) === AC) {
        // a cell-level mc:AlternateContent: its selected branch's cells, addressed from the element (CR-013)
        const ac = v as mce.AlternateContent;
        this.compareBranches(ac, address);
        this.cells(selectedBranchOf(ac), address, story, row, state);
      } else {
        const inner = blockItems(v);
        if (inner) this.cells(inner, address, story, row, state);
      }
    }
  }

  // --- run level --------------------------------------------------------------------------------

  private runs(items: Element[] | undefined, para: ParagraphState, state: RunState): void {
    if (!items) return;
    for (const item of items) {
      if (this.removes(item)) continue;
      const v = unwrap(item);
      const tn = typeOf(v);
      if (revisionHolder(item) !== undefined) {
        para.block.revisions++;
        this.runs(runItemsOfValue(v), para, this.holderState(item, state));
      } else if (tn === R) {
        this.run(v as wml.R, para, state);
      } else if (tn === AC) {
        const ac = v as mce.AlternateContent;
        this.compareBranches(ac, para.block.address);
        this.runs(selectedBranchOf(ac), para, state);
      } else if (isSdt(v) && this.unwraps(item)) {
        this.runs(sdtContentOf(v), para, state);
      } else if (isSdt(v)) {
        const sdt = v as wml.SdtRun;
        const bound = this.bound(sdt, para.block.address);
        if (bound !== undefined) {
          this.text(para, bound, false, state);
          continue;
        }
        const placeholder = showingPlaceholder(sdt.sdtPr);
        if (placeholder && !this.options.placeholders) continue;
        this.runs(sdtContentOf(v), para, placeholder ? state.withPlaceholder() : state);
      } else if (isBookmarkMarker(item)) {
        this.bookmark(item);
      } else if (tn === 'org_docx4j_wml.CTSimpleField') {
        this.simpleField(v as wml.CTSimpleField, para, state);
      } else if (isMath(v)) {
        this.math(v as Obj, para, state);
      } else if (tn === 'org_docx4j_wml.P.Hyperlink') {
        const saved = para.link;
        const target = hyperlinkTarget(v as wml.P.Hyperlink, para.story);
        if (target !== undefined) para.link = target;
        try {
          this.runs((v as wml.P.Hyperlink).content as Element[] | undefined, para, state);
        } finally {
          para.link = saved;
        }
      } else {
        const inner = runItemsOfValue(v);
        if (inner) this.runs(inner, para, state);
      }
    }
  }

  private holderState(item: unknown, state: RunState): RunState {
    if (this.view === 'markup' && this.mode === 'view') {
      const tc = revisionHolder(item)!;
      return state.withRevision(revisionOf(tc, isInsertionName(nameOf(item)) ? 'INSERTION' : 'DELETION'), false);
    }
    return state.withRevision(undefined, true);
  }

  private run(r: wml.R, para: ParagraphState, state: RunState): void {
    if (this.fieldsOnly) {
      for (const item of (r.content ?? []) as Element[]) if (FIELD_ITEMS.has(nameOf(item))) this.runItem(item, r, false, para, state);
      return;
    }
    const hidden = this.isHiddenRun(r, para.p);
    const before = textLength(para.block);
    const saved = para.format;
    para.format = this.formatOf(r, para.p);
    try {
      for (const item of (r.content ?? []) as Element[]) this.runItem(item, r, hidden, para, state);
    } finally {
      para.format = saved;
    }
    if (textLength(para.block) > before) {
      const mono = isMonospace(this.runFont(r, para.p));
      para.monospace = para.monospace === undefined ? mono : para.monospace && mono;
    }
  }

  private runItem(item: Element, r: wml.R, hidden: boolean, para: ParagraphState, state: RunState): void {
    const v = unwrap(item);
    const name = nameOf(item);
    const source: ItemSource = { run: r, item };
    switch (name) {
      case 'fldChar': {
        const fc = v as wml.FldChar;
        const fields = para.story.fields;
        if (fc.fldCharType === 'begin') {
          const f = new FieldState();
          f.locked = fc.fldLock === true;
          f.codeBlock = para.block;
          f.codeStart = para.block.runs.length;
          if (!this.options.toc && this.options.fieldCodes) {
            f.floorParagraph = para;
            f.previousFloor = para.mergeFloor;
            para.mergeFloor = para.block.runs.length;
          }
          fields.unshift(f);
        } else if (fc.fldCharType === 'separate') {
          const f = fields[0];
          if (f) {
            f.separated = true;
            f.link = fieldHyperlinkTarget(f.instr);
            if (!this.options.toc && !suppressed(para.story) && isTocField(f.instr)) {
              f.toc = true;
              para.toc = true;
              removeFieldCode(f, para);
            }
          }
        } else if (fc.fldCharType === 'end') {
          const f = fields.shift();
          if (f) {
            // a field inside a computed field's cached result is not read, nor recorded
            if (suppressed(para.story)) return;
            if (f.toc || (!this.options.toc && isTocField(f.instr))) {
              if (!f.toc) removeFieldCode(f, para);
              para.toc = true;
              return;
            }
            if (f.computed !== undefined) para.block.fields.push(new TextField(f.instr, f.computed, true));
            else this.endField(f.instr, f.separated ? f.result : undefined, para);
          }
        }
        return;
      }
      case 'instrText':
      case 'delInstrText': {
        if (name === 'delInstrText' && !state.kept && state.revision === undefined && this.view === 'accepted') return;
        const s = (v as wml.Text | undefined)?.value;
        if (s === undefined || s === null) return;
        const f = para.story.fields[0];
        if (f) f.instr += s;
        if (this.options.fieldCodes && !suppressed(para.story)) this.addText(para, s, 'FIELD_CODE', hidden, state);
        return;
      }
      case 't':
        this.text(para, (v as wml.Text | undefined)?.value, hidden, state, source);
        return;
      case 'delText': {
        // the holder decides; outside one, the element: no text after Accept All (REQ-404)
        if (!state.kept && state.revision === undefined) {
          if (this.view === 'accepted') return;
          if (this.view === 'markup' && this.mode === 'view') state = state.withRevision(new TextRevision('DELETION', undefined, undefined), false);
        }
        this.text(para, (v as wml.DelText | undefined)?.value, hidden, state, source);
        return;
      }
      case 'tab':
      case 'ptab':
        this.text(para, '\t', hidden, state, source);
        return;
      case 'br':
      case 'cr':
        this.text(para, '\n', hidden, state, source);
        return;
      case 'noBreakHyphen':
        this.text(para, '\u2011', hidden, state, source);
        return;
      case 'softHyphen':
        this.text(para, '\u00AD', hidden, state, source);
        return;
      case 'sym': {
        const code = (v as wml.R.Sym | undefined)?._char;
        if (typeof code === 'string' && /^[0-9A-Fa-f]{1,6}$/.test(code)) {
          const cp = parseInt(code, 16);
          if (cp <= 0x10FFFF && !(cp >= 0xD800 && cp <= 0xDFFF)) this.text(para, String.fromCodePoint(cp), hidden, state, source);
        }
        return;
      }
      case 'footnoteReference':
      case 'endnoteReference': {
        // in a computed field's cached result, a copy of a reference: neither written nor numbered
        if (inCode(para.story) || suppressed(para.story)) return;
        const foot = name === 'footnoteReference';
        const ref = v as wml.CTFtnEdnRef;
        const id = ref.id === undefined || ref.id === null ? '' : String(ref.id);
        const run = new TextRun('NOTE_REFERENCE', '', state.revision);
        run.hidden = hidden;
        run.noteType = foot ? 'footnote' : 'endnote';
        run.noteId = id;
        run.source = { run: r, paragraph: para.p };
        if (!(hidden && !this.options.hidden)) {
          // a reference carrying a custom mark is given no number (REQ-503); its note is still read
          run.noteNumber = this.context.noteNumber(run.noteType, id, ref.customMarkFollows === true);
        }
        para.block.runs.push(run);
        return;
      }
      case 'footnoteRef':
      case 'endnoteRef': {
        if (para.story.noteType === undefined) return;
        const run = new TextRun('NOTE_MARK', '', state.revision);
        run.noteType = para.story.noteType;
        run.noteNumber = para.story.noteNumber;
        run.noteId = para.story.noteId;
        para.block.runs.push(run);
        return;
      }
      case 'commentReference': {
        if (!this.commentsIncluded() || inCode(para.story) || suppressed(para.story)) return;
        const id = (v as wml.R.CommentReference).id;
        if (this.plan?.removesComment(id)) return;                    // gone with the revision it was on
        const run = new TextRun('COMMENT_REFERENCE', '', state.revision);
        run.hidden = hidden;
        run.commentId = id === undefined || id === null ? '' : String(id);
        run.source = { run: r, paragraph: para.p };
        para.block.runs.push(run);
        return;
      }
      case 'drawing':
      case 'pict':
      case 'object':
        if (!inCode(para.story) && !suppressed(para.story) && !(hidden && !this.options.hidden)) this.anchored(v, para, state);
        return;
      case 'lastRenderedPageBreak': {
        if (para.story.kind !== 'MAIN' || inCode(para.story) || (hidden && !this.options.hidden)) return;
        if (suppressed(para.story)) return;
        const run = new TextRun('PAGE_MARKER', '', state.revision);
        run.pageNumber = ++this.context.pageMarkers + 1;
        para.block.runs.push(run);
        return;
      }
      case 'AlternateContent': {
        const ac = v as mce.AlternateContent;
        this.compareBranches(ac, para.block.address);
        for (const o of selectedBranchOf(ac) ?? []) this.runItem(o, r, hidden, para, state);
        return;
      }
      case 'ruby': {
        const base = (v as wml.CTRuby).rubyBase as { egRubyContent?: Element[] } | undefined;
        this.runs(base?.egRubyContent, para, state);
        return;
      }
      default:
        // a run within a run (w:r in m:r, a ruby base's) or nothing that is text
        if (typeOf(v) === R) this.run(v as wml.R, para, state);
    }
  }

  /** Document text: a field's code where a field is in its code (shown only under `field_codes`), else text. */
  private text(para: ParagraphState, s: string | undefined | null, hidden: boolean, state: RunState, source?: ItemSource): void {
    if (s === undefined || s === null || s === '') return;
    if (inCode(para.story)) {
      if (this.options.fieldCodes) this.addText(para, s, 'FIELD_CODE', hidden, state);
      return;
    }
    if (suppressed(para.story)) return;
    if (!(hidden && !this.options.hidden)) {
      for (const f of para.story.fields) f.result += s;
      for (const b of this.openBookmarks.values()) {
        if (b.paragraphEnded && b.text.length > 0) b.text += '\n';
        b.paragraphEnded = false;
        b.text += s;
      }
    }
    this.addText(para, s, 'TEXT', hidden, state, source);
  }

  private addText(para: ParagraphState, s: string, kind: TextRunKind, hidden: boolean, state: RunState, source?: ItemSource): void {
    const run = new TextRun(kind, s, state.revision);
    run.hidden = hidden;
    run.placeholder = state.placeholder;
    if (kind === 'TEXT') {
      if (para.format) run.copyFormat(para.format);
      run.link = linkOf(para);
      run.math = para.math;
      run.displayMath = para.displayMath;
      if (source) {
        run.source = { run: source.run, paragraph: para.p };
        const span = para.officeSpan(source.item);
        if (span !== undefined && span.length === s.length) run.officeOffset = span.offset;
      }
    }
    const runs = para.block.runs;
    if (runs.length > para.mergeFloor && !run.math) {
      const last = runs[runs.length - 1]!;
      // one run as long as its text and Paragraph.text's agree character for character (CR-012 section 4.4)
      const contiguous = last.officeOffset === undefined ? run.officeOffset === undefined
        : run.officeOffset === last.officeOffset + last.text.length && last.source?.paragraph === run.source?.paragraph;
      if (last.kind === kind && last.hidden === hidden && last.placeholder === state.placeholder
        && sameRevision(last.revision, state.revision) && last.sameFormat(run) && contiguous) {
        const merged = new TextRun(kind, last.text + s, last.revision);
        merged.hidden = hidden;
        merged.placeholder = last.placeholder;
        merged.copyFormat(last);
        merged.officeOffset = last.officeOffset;
        merged.source = last.source;
        runs[runs.length - 1] = merged;
        return;
      }
    }
    runs.push(run);
  }

  private endField(instr: string, result: string | undefined, para: ParagraphState): void {
    para.block.fields.push(new TextField(instr, result));
    if (result === undefined) this.warn(new TextWarning(TextWarning.FIELD_NO_RESULT, para.block.address));
  }

  private simpleField(field: wml.CTSimpleField, para: ParagraphState, state: RunState): void {
    if (suppressed(para.story)) return;                              // inside a computed field's cached result
    const instr = field.instr ?? '';
    if (!this.options.toc && isTocField(instr)) {
      para.toc = true;                                               // a table of contents left out (D12)
      return;
    }
    if (this.options.fieldCodes && !inCode(para.story)) this.addText(para, instr, 'FIELD_CODE', false, state);
    const before = para.block.runs.length;
    const text = this.resultOf(field.content as Element[] | undefined, para, state, fieldHyperlinkTarget(instr));
    const hasResult = text !== '' || para.block.runs.length > before;
    para.block.fields.push(new TextField(instr, hasResult ? text : undefined));
    if (!hasResult) this.warn(new TextWarning(TextWarning.FIELD_NO_RESULT, para.block.address));
  }

  /** The runs of a simple field's result, added to the paragraph; their text. */
  private resultOf(content: Element[] | undefined, para: ParagraphState, state: RunState, link: string | undefined): string {
    const capture = new FieldState();
    capture.separated = true;
    capture.link = link;
    para.story.fields.push(capture);
    try {
      this.runs(content, para, state);
    } finally {
      const fields = para.story.fields;
      const at = fields.lastIndexOf(capture);
      if (at >= 0) fields.splice(at, 1);
    }
    return capture.result;
  }

  // --- math (REQ-604) ---------------------------------------------------------------------------

  /**
   * An equation, `$...$` (`$$...$$` for an `m:oMathPara`). Until CR-012 phase 5b ports a LaTeX converter
   * it is its `m:t` text in document order inside the delimiters, REQ-604's form for an equation that
   * can't be converted. In markup, an equation with tracked changes is its Reject All form deleted and
   * its Accept All form inserted, each as the plan for that view leaves it.
   */
  private math(v: Obj, para: ParagraphState, state: RunState): void {
    const display = typeOf(v) === 'org_docx4j_math.CTOMathPara';
    const savedMath = para.math;
    const savedDisplay = para.displayMath;
    para.math = true;
    para.displayMath = display;
    try {
      this.mathText(v, display, para, state);
    } finally {
      para.math = savedMath;
      para.displayMath = savedDisplay;
    }
  }

  private mathText(v: Obj, display: boolean, para: ParagraphState, state: RunState): void {
    if (this.view === 'markup' && this.mode === 'view') {
      const revisions = mathRevisions(v);
      if (revisions[0] !== undefined || revisions[1] !== undefined) {
        const original = this.context.plan(false);
        const accepted = this.context.plan(true);
        const before = mathRunsText(v, (o) => !original.removes(o));
        const after = mathRunsText(v, (o) => !accepted.removes(o));
        if (before !== after) {
          if (before !== '') this.text(para, delimit(before, display), false, state.withRevision(revisions[0] ?? new TextRevision('DELETION', undefined, undefined), false));
          if (after !== '') this.text(para, delimit(after, display), false, state.withRevision(revisions[1] ?? new TextRevision('INSERTION', undefined, undefined), false));
          return;
        }
        if (after !== '') this.text(para, delimit(after, display), false, state);
        return;
      }
    }
    const text = mathRunsText(v, (o) => this.keepMathInView(o));
    if (text !== '') this.text(para, delimit(text, display), false, state);
  }

  /** The view's rule for an equation's list members: what the plan removes goes; without a plan, a revision by its type. */
  private keepMathInView(o: unknown): boolean {
    if (this.removes(o)) return false;
    if (this.mode !== 'planless' || this.view === 'markup') return true;
    const ins = mathRevisionType(o);
    return ins === undefined || ins === (this.view === 'accepted');
  }

  // --- what a run anchors: text boxes (REQ-518), WordArt, pictures (REQ-514), objects (REQ-512) ---

  /**
   * What a run's drawing, picture or object holds, in document order, branches selected: text boxes and
   * WordArt to follow the paragraph; pictures and objects as runs where they are. A chart or SmartArt
   * contributes nothing until CR-012 phase 5c.
   */
  private anchored(v: unknown, para: ParagraphState, state: RunState): void {
    const visitor: Visitor = {
      enter: (value) => {
        switch (value.TYPE_NAME) {
          case 'org_docx4j_wml.CTTxbxContent':
            if (this.options.textBoxes) para.story.pending.push(new Box(value as wml.CTTxbxContent, `${para.address}/textbox:${para.boxes++}`));
            return false;
          case AC: {
            const ac = value as unknown as mce.AlternateContent;
            this.compareBranches(ac, para.block.address);
            for (const o of selectedBranchOf(ac) ?? []) walk(unwrap(o), 1, visitor);
            return false;
          }
          case 'org_docx4j_wml.CTObject':
            this.object(value as wml.CTObject, para, state);
            return false;
          case 'org_docx4j_wml.CTControl':
            objectRun(para, undefined, state);
            return false;
          case 'org_docx4j_dml_wordprocessingDrawing.Inline':
          case 'org_docx4j_dml_wordprocessingDrawing.Anchor':
            return this.graphic(value, para, state);
          case 'org_docx4j_dml_picture.Pic': {
            const pic = value as { nvPicPr?: { cNvPr?: NonVisual } };
            this.image(pic.nvPicPr?.cNvPr, blipPart(value, para), false, value, para, state);
            return false;
          }
          case 'org_docx4j_vml.CTShape':
            return this.vmlShape(value, para, state);
          default:
            return true;
        }
      },
    };
    walk(v, 0, visitor);
  }

  /** A `wp:inline` or `wp:anchor`: a chart, SmartArt, a picture, or a shape to walk into. */
  private graphic(anchor: Obj, para: ParagraphState, state: RunState): boolean {
    const graphic = anchor.graphic as { graphicData?: { any?: unknown[]; uri?: string } } | undefined;
    const docPr = anchor.docPr as NonVisual | undefined;
    if (!graphic?.graphicData) return false;
    const uri = graphic.graphicData.uri ?? '';
    // a chart (c:chart, cx:chart): CR-012 phase 5c
    if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/chart' || uri.startsWith('http://schemas.microsoft.com/office/drawing/2014/chartex')) return false;
    for (const o of graphic.graphicData.any ?? []) {
      const v = unwrap(o);
      const tn = typeOf(v);
      if (tn === 'org_docx4j_dml_chart.CTRelId' || tn === 'org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex.CTRelId') return false;
      if (tn === 'org_docx4j_dml_diagram.CTRelIds') return false;    // SmartArt: CR-012 phase 5c
      if (tn === 'org_docx4j_dml_picture.Pic') {
        this.image(docPr, blipPart(v as Obj, para), false, anchor, para, state);
        return false;
      }
    }
    return true;                                                     // a shape, group or canvas: its text boxes and pictures
  }

  /** A VML shape: legacy WordArt, a picture, or a shape to walk into for its text box. */
  private vmlShape(shape: Obj, para: ParagraphState, state: RunState): boolean {
    let imageData: { id?: string } | undefined;
    for (const e of (shape.pathOrFormulasOrHandles ?? []) as Element[]) {
      const v = unwrap(e);
      const tn = typeOf(v);
      if (tn === 'org_docx4j_vml.CTTextPath') {
        const s = (v as { string?: string }).string;
        if (s !== undefined && s !== '') {
          if (this.options.textBoxes) para.story.pending.push(new WordArt(s, `${para.address}/textbox:${para.boxes++}`));
          return false;
        }
      } else if (tn === 'org_docx4j_vml.CTImageData') {
        imageData = v as { id?: string };
      }
    }
    if (imageData) {
      this.imageOf(shape.alt as string | undefined, undefined, undefined, false, true, related(para.story.part, imageData.id), shape, para, state);
    }
    return true;
  }

  /** A `w:object`: an object run with its ProgID (REQ-512); its package is read under `embedded: true`, phase 5d. */
  private object(object: wml.CTObject, para: ParagraphState, state: RunState): void {
    if (object.control) {
      objectRun(para, undefined, state);
      return;
    }
    for (const o of (object.anyAndAny ?? []) as unknown[]) {
      const v = unwrap(o);
      const tn = typeOf(v);
      if (tn === 'org_docx4j_vml_officedrawing.CTOLEObject') {
        objectRun(para, (v as { progID?: string }).progID, state);
        return;
      }
      if (tn === 'org_docx4j_wml.CTControl') {
        objectRun(para, undefined, state);
        return;
      }
    }
  }

  /**
   * A DrawingML picture: its `wp:docPr` (or a nested picture's `pic:cNvPr`). `descr` and `title` are the
   * schema's default, `''`, where the element doesn't give them, as docx4j's JAXB getters read them (CR-012
   * section 20).
   */
  private image(nv: NonVisual | undefined, part: Part | undefined, vml: boolean, source: object, para: ParagraphState, state: RunState): void {
    this.imageOf(nv ? nv.descr ?? '' : undefined, nv?.name, nv ? nv.title ?? '' : undefined, decorative(nv), vml, part, source, para, state);
  }

  private imageOf(descr: string | undefined, name: string | undefined, title: string | undefined, isDecorative: boolean, vml: boolean,
    part: Part | undefined, source: object, para: ParagraphState, state: RunState): void {
    const image = new TextImage(descr, name, title, isDecorative, vml, part, source);
    const run = new TextRun('IMAGE', '', state.revision);
    run.image = image;
    try {
      run.description = this.options.imageDescriber(image);
    } catch {
      this.warn(new TextWarning(TextWarning.IMAGE_NOT_DESCRIBED, para.block.address, undefined, undefined, part?.partName.name));
    }
    run.placeholder = state.placeholder;
    para.block.runs.push(run);
  }

  // --- markup compatibility: REQ-307 ------------------------------------------------------------

  /**
   * `mce.branches_differ` (REQ-307): where the selected branch and another both hold text, and the texts
   * differ after REQ-1201's normalisation. A branch's text is every `w:t`, `m:t` and `a:t` in it, whatever
   * the view and hidden or not, as docx4j reads it: a comparison made at load knows no view, and a branch
   * that differs only in hidden or deleted text is one an agent should hear of (Jason, 2026-10-21, for WD
   * 0.29; CR-012 section 15.6).
   */
  private compareBranches(ac: mce.AlternateContent, address: string): void {
    if (!this.options.warnings || this.fieldsOnly) return;
    const selected = selectedBranchOf(ac);
    if (selected === undefined) return;
    const mine = branchText(selected);
    if (mine === '') return;
    const branches: { any?: unknown[]; requires?: string }[] = [...(ac.choice ?? [])];
    if (ac.fallback) branches.push(ac.fallback);
    for (const other of branches) {
      if ((other.any ?? []) === selected) continue;
      const theirs = branchText(other.any ?? []);
      if (theirs !== '' && theirs !== mine) {
        const chosen = (ac.choice ?? []).find((c) => (c.any ?? []) === selected);
        const branch = chosen ? `mc:Choice Requires="${chosen.requires}"` : 'mc:Fallback';
        this.warn(new TextWarning(TextWarning.MCE_BRANCHES_DIFFER, address, branch));
        return;
      }
    }
  }

  warn(w: TextWarning): void {
    if (this.warned.has(w.key)) return;
    this.warned.add(w.key);
    this.warnings.push(w);
  }

  // --- headers and footers by section (REQ-502) ---------------------------------------------------

  private sectionsWithHeaders(sections: (wml.SectPr | undefined)[], mainBlocks: TextBlock[], headers: boolean, footers: boolean, out: TextBlock[]): void {
    const evenAndOdd = on(this.trees.settings?.evenAndOddHeaders);
    const emitted = new Set<string>();
    const lastHeader = new Map<string, string>();
    const lastFooter = new Map<string, string>();
    let b = 0;
    for (let s = 0; s < sections.length; s++) {
      const sectPr = sections[s];
      let titlePg = false;
      if (sectPr) {
        for (const ref of (sectPr.egHdrFtrReferences ?? []) as unknown[]) {
          const v = unwrap(ref) as { TYPE_NAME?: string; type?: string; id?: string };
          if (v.TYPE_NAME === 'org_docx4j_wml.HeaderReference' && v.type !== undefined && v.id !== undefined) lastHeader.set(v.type, v.id);
          else if (v.TYPE_NAME === 'org_docx4j_wml.FooterReference' && v.type !== undefined && v.id !== undefined) lastFooter.set(v.type, v.id);
        }
        titlePg = on(sectPr.titlePg);
      }
      const shown: string[] = [];
      if (titlePg) shown.push('first');
      shown.push('default');
      if (evenAndOdd) shown.push('even');
      if (headers) for (const type of shown) this.headerOrFooter(lastHeader.get(type), 'HEADER', emitted, out);
      while (b < mainBlocks.length && mainBlocks[b]!.section <= s) out.push(mainBlocks[b++]!);
      if (s === sections.length - 1) while (b < mainBlocks.length) out.push(mainBlocks[b++]!);
      if (footers) for (const type of shown) this.headerOrFooter(lastFooter.get(type), 'FOOTER', emitted, out);
    }
  }

  private headerOrFooter(rId: string | undefined, kind: 'HEADER' | 'FOOTER', emitted: Set<string>, out: TextBlock[]): void {
    if (rId === undefined || emitted.has(rId)) return;
    emitted.add(rId);
    const main = this.trees.main;
    const part = main ? related(main.part, rId) : undefined;
    const content = part ? this.trees.headersAndFooters.get(part) : undefined;
    if (!content) return;
    const prefix = `${this.root}${kind === 'HEADER' ? 'header:' : 'footer:'}${rId}`;
    const story = new Story(this, kind, prefix, part);
    this.blocks(blockItems(content), prefix, story, out);
  }

  /** A header or footer part's content as its own story (`Texts.of(headerPart)`). */
  headerOrFooterPart(part: Part, kind: 'HEADER' | 'FOOTER', rId: string): TextResult {
    const out: TextBlock[] = [];
    const content = this.trees.headersAndFooters.get(part);
    const prefix = `${this.root}${kind === 'HEADER' ? 'header:' : 'footer:'}${rId}`;
    if (content) this.blocks(blockItems(content), prefix, new Story(this, kind, prefix, part), out);
    return new TextResult(out, new Map(), this.warnings, { [TextResult.FIELDS_UPDATE_ON_OPEN]: this.fieldsUpdateOnOpen() }, this.options);
  }

  // --- notes (REQ-503) ------------------------------------------------------------------------

  private notes(type: NoteType, out: TextBlock[]): void {
    const holder = type === 'footnote' ? this.trees.footnotes : this.trees.endnotes;
    const numbers = type === 'footnote' ? this.context.footnotes : this.context.endnotes;
    if (!holder || numbers.size === 0) return;
    const notes = ((type === 'footnote' ? (holder.tree as wml.CTFootnotes).footnote : (holder.tree as wml.CTEndnotes).endnote) ?? []) as wml.CTFtnEdn[];
    const byId = new Map<string, wml.CTFtnEdn>();
    for (const note of notes) if (note.id !== undefined && note.id !== null) byId.set(String(note.id), note);
    for (const [id, number] of numbers) {
      const note = byId.get(id);
      if (!note) continue;
      if (note.type !== undefined && note.type !== 'normal') continue;
      const prefix = `${this.root}${type}:${id}`;
      const story = new Story(this, type === 'footnote' ? 'FOOTNOTE' : 'ENDNOTE', prefix, holder.part);
      story.noteType = type;
      story.noteNumber = number;
      story.noteId = id;
      const blocks: TextBlock[] = [];
      this.blocks(note.content as Element[] | undefined, prefix, story, blocks);
      if (blocks.length === 0) continue;
      // the note is prefixed [n]: its own mark, or one added at its start
      const marked = blocks.some((b) => b.runs.some((r) => r.kind === 'NOTE_MARK'));
      if (!marked) {
        const first = blocks.find((b) => b.kind === 'PARAGRAPH');
        if (first) {
          const mark = new TextRun('NOTE_MARK', '', undefined);
          mark.noteType = type;
          mark.noteNumber = number;
          mark.noteId = id;
          first.runs.unshift(mark);
        }
      }
      // every block of the note says whose it is, a table's too
      for (const b of blocks) {
        if (b.noteType !== undefined) continue;
        b.noteType = type;
        b.noteNumber = number;
        b.noteId = id;
      }
      out.push(...blocks);
    }
  }

  // --- comments (REQ-504) ---------------------------------------------------------------------

  private comments(out: Map<string, TextComment>): void {
    const part = this.trees.comments;
    if (!part) return;
    const extended = new Map<string, { done?: unknown; paraIdParent?: string }>();
    for (const ex of (this.trees.commentsExtended?.commentEx ?? []) as { paraId?: string; done?: unknown; paraIdParent?: string }[]) {
      if (ex.paraId) extended.set(ex.paraId.toUpperCase(), ex);
    }
    const idByParaId = new Map<string, string>();
    const anchored = this.trees.anchoredComments();
    for (const c of (part.tree.comment ?? []) as wml.Comments.Comment[]) {
      if (this.plan?.removesComment(c.id)) continue;
      const id = c.id === undefined || c.id === null ? '' : String(c.id);
      // a comment with no anchor in the document is none of its text: Word drops it on save (REQ-504)
      if (!anchored.has(id)) continue;
      const comment = new TextComment(id);
      comment.author = c.author;
      comment.initials = c.initials;
      comment.date = calendarText(c.date);
      const prefix = `${this.root}comment:${id}`;
      this.blocks(c.content as Element[] | undefined, prefix, new Story(this, 'COMMENT', prefix, part.part), comment.blocks);
      const paraId = lastParaId(c.content as Element[] | undefined);
      if (paraId !== undefined) {
        idByParaId.set(paraId.toUpperCase(), id);
        const ex = extended.get(paraId.toUpperCase());
        if (ex) {
          comment.resolved = ex.done === true || ex.done === 1 || ex.done === '1' || ex.done === 'true';
          comment.parentId = ex.paraIdParent;
        }
      }
      out.set(id, comment);
    }
    for (const comment of out.values()) {
      if (comment.parentId === undefined) continue;
      const parent = idByParaId.get(comment.parentId.toUpperCase());
      comment.parentId = parent;
      if (parent !== undefined) out.get(parent)?.replies.push(comment);
    }
  }

  // --- the view -------------------------------------------------------------------------------

  /** Whether the view removes a list member (always false in markup and on a resolved copy). */
  removes(item: unknown): boolean {
    if (this.plan) return this.plan.removes(item);
    if (this.mode === 'planless' && this.view !== 'markup') {
      if (revisionHolder(item) === undefined) return false;
      return isInsertionName(nameOf(item)) !== (this.view === 'accepted');
    }
    return false;
  }

  /** Whether the view takes a content control away and leaves its content (an inserted one, rejected). */
  private unwraps(item: unknown): boolean {
    return this.plan !== undefined && this.plan.unwraps(item);
  }

  private joinsNext(p: wml.P): boolean {
    return this.plan !== undefined && this.plan.removesMark(p);
  }

  // --- hidden text (REQ-507) and formatting -----------------------------------------------------

  private isHiddenRun(r: wml.R, p: wml.P): boolean {
    const own = r.rPr?.vanish;
    if (own !== undefined) return on(own);
    return on(this.styleRPr(r.rPr?.rStyle?.val, p)?.vanish);
  }

  /**
   * A run's formatting beyond its paragraph style's (CR-046 phase 3): bold, italic, struck through,
   * monospaced, each as the run's own property gives it, else its styles', where the paragraph style
   * alone doesn't give it. A run carrying only that.
   */
  private formatOf(r: wml.R, p: wml.P): TextRun {
    const own = r.rPr;
    const styles = this.styleRPr(own?.rStyle?.val, p);
    const paragraph = this.styleRPr(undefined, p);
    const f = new TextRun('TEXT', '', undefined);
    f.bold = onOf(own?.b, styles?.b) && !on(paragraph?.b);
    f.italic = onOf(own?.i, styles?.i) && !on(paragraph?.i);
    f.strike = onOf(own?.strike, styles?.strike) && !on(paragraph?.strike);
    f.monospace = isMonospace(this.runFont(r, p)) && !isMonospace(asciiFont(paragraph));
    return f;
  }

  /** The run's font for `monospace`: its own `w:rFonts/@w:ascii`, else its styles'. */
  private runFont(r: wml.R, p: wml.P): string | undefined {
    return asciiFont(r.rPr) ?? asciiFont(this.styleRPr(r.rPr?.rStyle?.val, p));
  }

  /** The effective run properties of the paragraph's style and a character style, cached by their ids; undefined without styles. */
  private styleRPr(rStyle: string | undefined, p: wml.P): wml.RPr | undefined {
    if (!this.resolver) return undefined;
    const pStyle = p.pPr?.pStyle?.val;
    const key = `${String(pStyle)}|${String(rStyle)}`;
    if (this.styleRPrCache.has(key)) return this.styleRPrCache.get(key);
    let effective: wml.RPr | undefined;
    try {
      const styleOnly = rStyle === undefined ? undefined : { TYPE_NAME: 'org_docx4j_wml.RPr', rStyle: { TYPE_NAME: 'org_docx4j_wml.RStyle', val: rStyle } } as wml.RPr;
      const pPrStyle = pStyle === undefined ? undefined : { TYPE_NAME: 'org_docx4j_wml.PPr', pStyle: { TYPE_NAME: 'org_docx4j_wml.PPrBase.PStyle', val: pStyle } } as wml.PPr;
      effective = this.resolver.getEffectiveRPr(styleOnly, pPrStyle);
    } catch {
      effective = undefined;
    }
    this.styleRPrCache.set(key, effective);
    return effective;
  }

  private isHiddenMark(pPr: wml.PPr | undefined): boolean {
    const own = pPr?.rPr?.vanish;
    if (own !== undefined) return on(own);
    if (!this.resolver) return false;
    try {
      return on(this.resolver.getEffectiveParagraphMarkRPr(pPr)?.vanish);
    } catch {
      return false;
    }
  }
}

// --- the anchored items ---------------------------------------------------------------------------

/** A text box: its content, read as a story of its own. */
class Box implements Anchored {
  constructor(private readonly content: wml.CTTxbxContent, readonly prefix: string) {}

  emit(x: Extractor, anchoring: Story, sink: TextBlock[]): void {
    const inner = new Story(x, 'TEXT_BOX', this.prefix, anchoring.part);
    inner.noteType = anchoring.noteType;
    inner.noteNumber = anchoring.noteNumber;
    inner.noteId = anchoring.noteId;
    x.blocks(this.content.content as Element[] | undefined, this.prefix, inner, sink);
  }
}

/** Legacy VML WordArt: a one-paragraph text box holding its `v:textpath/@string` (REQ-511). */
class WordArt implements Anchored {
  constructor(private readonly text: string, readonly prefix: string) {}

  emit(_x: Extractor, _anchoring: Story, sink: TextBlock[]): void {
    sink.push(paragraphBlock('TEXT_BOX', `${this.prefix}/0`, this.text));
  }
}

// --- helpers ----------------------------------------------------------------------------------------

/** A paragraph block of one text run (docx4j `Embedded.paragraph`). */
function paragraphBlock(story: TextStory, address: string, text: string): TextBlock {
  const b = new TextBlock('PARAGRAPH', story, address);
  if (text !== '') b.runs.push(new TextRun('TEXT', text, undefined));
  return b;
}

/** A content control's content list, whatever its level. */
function sdtContentOf(v: unknown): Element[] | undefined {
  const content = (v as { sdtContent?: { content?: Element[] } } | undefined)?.sdtContent;
  return content?.content;
}

function textLength(block: TextBlock): number {
  let n = 0;
  for (const run of block.runs) if (run.kind === 'TEXT') n += run.text.length;
  return n;
}

function sameRevision(a: TextRevision | undefined, b: TextRevision | undefined): boolean {
  return a === undefined ? b === undefined : a.equals(b);
}

/** An on/off property: the run's own where it has one, else its styles'. */
function onOf(own: { val?: boolean } | undefined, styles: { val?: boolean } | undefined): boolean {
  if (own !== undefined) return on(own);
  return on(styles);
}

function twips(value: unknown): number | undefined {
  if (value === undefined || value === null) return undefined;
  const n = Number(value);
  return Number.isFinite(n) ? Math.trunc(n) : undefined;
}

function markRevisions(p: wml.P): number {
  const rPr = p.pPr?.rPr;
  if (!rPr) return 0;
  return (rPr.ins ? 1 : 0) + (rPr.del ? 1 : 0) + (rPr.moveFrom ? 1 : 0) + (rPr.moveTo ? 1 : 0);
}

function listLabel(result: NumberingResult | undefined): string | undefined {
  if (!result || result.notNumbered) return undefined;
  let label = result.numString;
  if (label === undefined || label === '') label = result.bullet;
  if (label === undefined) return undefined;
  const suff = result.lvl?.suff?.val;
  const suffix = suff === 'space' ? ' ' : suff === 'nothing' ? '' : '\t';
  return label + suffix;
}

function inCode(story: Story): boolean {
  return story.fields.some((f) => !f.separated);
}

/** Inside the cached result of a field whose result was computed (REQ-506), or a table of contents left out. */
function suppressed(story: Story): boolean {
  return story.fields.some((f) => f.computed !== undefined || f.toc);
}

/** Inside a table of contents left out (D12). */
function inToc(story: Story): boolean {
  return story.fields.some((f) => f.toc);
}

/** A TOC field's instruction: a table of contents, of figures (`\c`), of anything it collects. */
export function isTocField(instr: string): boolean {
  const tokens = fieldTokens(instr);
  return tokens.length > 0 && tokens[0]!.toUpperCase() === 'TOC';
}

/** A table of contents' code, shown under `field_codes` until it was known to be one, taken out again. */
function removeFieldCode(f: FieldState, para: ParagraphState): void {
  // what follows the field may merge with what preceded it, as it would with no field codes
  if (f.floorParagraph) f.floorParagraph.mergeFloor = f.previousFloor;
  if (f.codeBlock) {
    const runs = f.codeBlock.runs;
    for (let k = runs.length - 1; k >= Math.min(f.codeStart, runs.length); k--) if (runs[k]!.kind === 'FIELD_CODE') runs.splice(k, 1);
  }
  if (para.block !== f.codeBlock) {
    // a code over more than one paragraph: this one's code too
    const runs = para.block.runs;
    for (let k = runs.length - 1; k >= 0; k--) if (runs[k]!.kind === 'FIELD_CODE') runs.splice(k, 1);
  }
}

/** The link text being read is in: the `w:hyperlink`'s, else the innermost HYPERLINK field's whose result it is. */
function linkOf(para: ParagraphState): string | undefined {
  if (para.link !== undefined) return para.link;
  for (const f of para.story.fields) if (f.separated && f.link !== undefined) return f.link;
  return undefined;
}

/** The part a relationship of `host` names. */
function related(host: Part | undefined, rId: string | undefined): Part | undefined {
  if (!host || !rId) return undefined;
  const rels = host.getRelationshipsPart(false);
  if (!rels) return undefined;
  try {
    return rels.getPart(rId);
  } catch {
    return undefined;
  }
}

/** A `w:hyperlink`'s target: its relationship's target, and `#` and its anchor; else `#` and its anchor. */
function hyperlinkTarget(h: wml.P.Hyperlink, story: Story): string | undefined {
  let target: string | undefined;
  if (h.id && story.part) {
    const rel = story.part.getRelationshipsPart(false)?.getRelationshipById(h.id);
    if (rel?.target) target = rel.target;
  }
  const anchor = h.anchor === undefined || h.anchor === '' ? undefined : h.anchor;
  if (anchor !== undefined) return target === undefined ? `#${anchor}` : `${target}#${anchor}`;
  return target;
}

/**
 * A HYPERLINK field's target from its instruction (WD REQ-703): its address, and `#` and the bookmark its
 * `\l` names; undefined for any other field, or one with neither.
 */
export function fieldHyperlinkTarget(instr: string): string | undefined {
  const tokens = fieldTokens(instr);
  if (tokens.length === 0 || tokens[0]!.toUpperCase() !== 'HYPERLINK') return undefined;
  let address: string | undefined;
  let bookmark: string | undefined;
  for (let i = 1; i < tokens.length; i++) {
    const t = tokens[i]!;
    if (t.length === 2 && t[0] === '\\') {
      const sw = t[1]!.toLowerCase();
      // \l, \o, \t take an argument; \m and \n don't
      if ((sw === 'l' || sw === 'o' || sw === 't') && i + 1 < tokens.length) {
        const arg = tokens[++i]!;
        if (sw === 'l') bookmark = arg;
      }
    } else if (address === undefined) {
      address = t;
    }
  }
  if (address === '') address = undefined;
  if (bookmark === '') bookmark = undefined;
  if (bookmark !== undefined) return address === undefined ? `#${bookmark}` : `${address}#${bookmark}`;
  return address;
}

const isJavaWhitespace = (c: string): boolean => /[\t\n\x0B\f\r\x1C-\x1F \u1680\u2000-\u2006\u2008-\u200A\u2028\u2029\u205F\u3000]/.test(c);

/** A field instruction's words: a quoted argument is one, its quotes removed and `\\` and `\"` unescaped. */
function fieldTokens(instr: string | undefined): string[] {
  const tokens: string[] = [];
  if (!instr) return tokens;
  let i = 0;
  const n = instr.length;
  while (i < n) {
    const c = instr[i]!;
    if (isJavaWhitespace(c)) {
      i++;
    } else if (c === '"') {
      let out = '';
      i++;
      while (i < n && instr[i] !== '"') {
        let d = instr[i]!;
        if (d === '\\' && i + 1 < n && (instr[i + 1] === '\\' || instr[i + 1] === '"')) d = instr[++i]!;
        out += d;
        i++;
      }
      i++;                                                           // the closing quote
      tokens.push(out);
    } else {
      let j = i;
      while (j < n && !isJavaWhitespace(instr[j]!) && instr[j] !== '"') j++;
      tokens.push(instr.slice(i, j));
      i = j;
    }
  }
  return tokens;
}

function objectRun(para: ParagraphState, progId: string | undefined, state: RunState): void {
  const run = new TextRun('OBJECT', '', state.revision);
  run.progId = progId;
  para.block.runs.push(run);
}

/** A picture's non-visual properties. */
interface NonVisual { descr?: string; name?: string; title?: string; extLst?: { ext?: { uri?: string; any?: unknown }[] } }

/** `adec:decorative val="1"` among the properties' extensions. */
function decorative(nv: NonVisual | undefined): boolean {
  for (const ext of nv?.extLst?.ext ?? []) {
    if ((ext.uri ?? '').toUpperCase() !== DECORATIVE_EXT) continue;
    const any = ext.any;
    let val: unknown;
    if (typeof any === 'object' && any !== null && typeof (any as { getAttribute?: unknown }).getAttribute === 'function') {
      val = (any as globalThis.Element).getAttribute('val');
    } else {
      val = (unwrap(any) as { val?: unknown } | undefined)?.val;
    }
    if (val === true || val === 1 || val === '1' || String(val).toLowerCase() === 'true') return true;
  }
  return false;
}

function blipPart(pic: Obj, para: ParagraphState): Part | undefined {
  const blip = (pic.blipFill as { blip?: { embed?: string; extLst?: { ext?: { uri?: string; any?: unknown }[] } } } | undefined)?.blip;
  if (!blip) return undefined;
  // an SVG picture with no PNG copy: its svgBlip's (REQ-514, docx4j #699)
  return related(para.story.part, blip.embed) ?? related(para.story.part, svgBlipEmbed(blip));
}

/** An `a:blip`'s `asvg:svgBlip/@r:embed`. */
function svgBlipEmbed(blip: { extLst?: { ext?: { uri?: string; any?: unknown }[] } }): string | undefined {
  for (const ext of blip.extLst?.ext ?? []) {
    if ((ext.uri ?? '').toUpperCase() !== SVG_BLIP_EXTENSION) continue;
    const any = ext.any;
    if (typeof any === 'object' && any !== null && typeof (any as { getAttributeNS?: unknown }).getAttributeNS === 'function') {
      const id = (any as globalThis.Element).getAttributeNS(R_NS, 'embed');
      if (id) return id;
      continue;
    }
    const v = unwrap(any) as { embed?: string } | undefined;
    if (v?.embed) return v.embed;
  }
  return undefined;
}

/** A branch's text (`w:t`, `m:t`, `a:t`), normalised as REQ-1201 normalises. docx4j `normalized`. */
function branchText(branch: unknown[]): string {
  let out = '';
  const visitor: Visitor = {
    item: (item) => {
      const v = unwrap(item);
      const tn = typeOf(v);
      if (tn === 'org_docx4j_wml.Text' && nameOf(item) === 't') out += String((v as wml.Text).value ?? '');
      else if (tn === 'org_docx4j_math.CTText') out += String((v as { value?: string }).value ?? '');
      else if (tn === 'org_docx4j_dml.CTRegularTextRun') out += String((v as { t?: string }).t ?? '');
      return true;
    },
  };
  // each member as an item first: a run-level element's branch holds bare w:t
  for (const o of branch) if (visitor.item!(o) !== false) walk(unwrap(o), 0, visitor);
  return normalize(out);
}

/** The last paragraph's `w14:paraId` in a comment's content (a branch's too). */
function lastParaId(content: Element[] | undefined): string | undefined {
  let id: string | undefined;
  for (const o of content ?? []) {
    const v = unwrap(o);
    if (typeOf(v) === P && (v as wml.P).paraId) id = (v as wml.P).paraId;
    else if (typeOf(v) === AC) {
      const inner = lastParaId(selectedBranchOf(v as mce.AlternateContent));
      if (inner !== undefined) id = inner;
    }
  }
  return id;
}

function delimit(text: string, display: boolean): string {
  return display ? `$$${text}$$` : `$${text}$`;
}

/** For a `w:ins` or `w:del` in an equation's run, whether it is an insertion; else undefined. */
function mathRevisionType(o: unknown): boolean | undefined {
  if (typeOf(unwrap(o)) === 'org_docx4j_wml.CTMathRunTrackChange') return nameOf(o) === 'ins';
  return undefined;
}

/** The first deletion's and the first insertion's revision in an equation. */
function mathRevisions(v: unknown): [TextRevision | undefined, TextRevision | undefined] {
  const out: [TextRevision | undefined, TextRevision | undefined] = [undefined, undefined];
  walk(v, 0, {
    item: (item) => {
      const ins = mathRevisionType(item);
      if (ins !== undefined) {
        const tc = unwrap(item) as wml.CTMathRunTrackChange;
        const i = ins ? 1 : 0;
        out[i] ??= revisionOf(tc, ins ? 'INSERTION' : 'DELETION');
      }
      return true;
    },
  });
  return out;
}

/** An equation's runs' text (`m:t`, `w:t`), what `keep` refuses left out: REQ-604's text for an equation not converted. */
function mathRunsText(v: unknown, keep: (o: unknown) => boolean): string {
  let out = '';
  const item = (o: unknown): void => {
    const ov = unwrap(o);
    const tn = typeOf(ov);
    if (tn === 'org_docx4j_math.CTText') out += String((ov as { value?: string }).value ?? '');
    else if (tn === 'org_docx4j_wml.Text' && nameOf(o) === 't') out += String((ov as wml.Text).value ?? '');
  };
  const visit = (value: unknown, depth: number): void => {
    if (value === undefined || value === null || depth > 64) return;
    const tn = typeOf(value);
    if (tn === 'org_docx4j_math.CTR') {
      for (const o of ((value as { content?: unknown[] }).content ?? [])) {
        if (!keep(o)) continue;
        const ov = unwrap(o);
        if (typeOf(ov) === 'org_docx4j_wml.CTMathRunTrackChange') {
          for (const inner of ((ov as { rest?: unknown[] }).rest ?? [])) item(inner);
          continue;
        }
        item(o);
      }
      return;
    }
    if (tn === 'org_docx4j_math.CTCtrlPr') return;
    for (const child of children(value)) {
      if (!keep(child)) continue;
      visit(unwrap(child), depth + 1);
    }
  };
  visit(v, 0);
  return out;
}

