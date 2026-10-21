// CR-012 phase 2: what a text includes and how it is written, docx4j's org.docx4j.model.text.TextOptions.
// The option names are the specification's section 10 names verbatim (`note_references`, `field_codes`),
// not camelCase (CR-012 D1), so that the suite's scenarios, the editor's tools and the MCP tools pass
// options through unchanged; docx4j's `of(Map)` and `set(name, value)` take the same.
import { Docx4JException } from '../../opc/exceptions.mjs';
import type { ImageDescriber } from './TextResult.mjs';
import { AltTextDescriber } from './TextResult.mjs';

/**
 * The three revision views (WD section 4, REQ-401): the document as Word's Accept All leaves it, as
 * Reject All leaves it, and every run as the package holds it, insertions and deletions marked.
 * `current` is accepted as a synonym of `accepted` (WD section 4.1's note). docx4j's `TextView`, by
 * another name: the content API's `TextView` (CR-002 phase F) is `getText`'s two views.
 */
export type TextModelView = 'accepted' | 'original' | 'markup';

/** The stories `include` adds (REQ-501, REQ-502). */
export type TextInclude = 'headers' | 'footers' | 'comments' | 'glossary';
/** `fields` (REQ-506). */
export type TextFields = 'cached' | 'update';
/** `bindings` (REQ-509). */
export type TextBindings = 'bound' | 'surface';
/** `symbols` (REQ-601). */
export type TextSymbols = 'asis' | 'unicode';
/** `math` (REQ-604). */
export type TextMath = 'latex' | 'unicodemath';
/** `tables` (REQ-605). */
export type TextTables = 'auto' | 'tsv' | 'grid';
/** `markup_detail` (REQ-702). */
export type TextMarkupDetail = 'type' | 'author' | 'author_date';
/** `layout` (REQ-709). */
export type TextLayout = 'none' | 'indent';

/** A boolean option's value: a boolean, or `"true"` / `"false"` as a tool passes it. */
type Bool = boolean | 'true' | 'false';

/**
 * The options by the specification's names (WD section 10), every one optional, each absent one at its
 * default. `include` takes a list of story names or a string of them separated by commas or spaces;
 * a boolean takes `"true"` or `"false"` as well; any other value is its name, in any case.
 * `imageDescriber` is the host's (REQ-514), not one of the specification's options.
 */
export interface TextOptionsInit {
  view?: TextModelView | 'current';
  include?: readonly TextInclude[] | string;
  text_boxes?: Bool;
  notes?: Bool;
  note_references?: Bool;
  comments?: Bool;
  hidden?: Bool;
  field_codes?: Bool;
  fields?: TextFields;
  toc?: Bool;
  placeholders?: Bool;
  bindings?: TextBindings;
  list_labels?: Bool;
  charts?: Bool;
  smartart?: Bool;
  embedded?: Bool;
  altchunks?: Bool;
  alt_text?: Bool;
  metadata?: Bool;
  symbols?: TextSymbols;
  math?: TextMath;
  tables?: TextTables;
  markup_detail?: TextMarkupDetail;
  page_markers?: Bool;
  layout?: TextLayout;
  warnings?: Bool;
  /** The describer images take their text from (REQ-514); `AltTextDescriber` by default. */
  imageDescriber?: ImageDescriber;
}

/**
 * An option refused (REQ-1001): an unknown name, a value the specification does not define, or a
 * value this package does not implement yet. `statement` names the requirement, `option` the option.
 */
export class TextOptionError extends Docx4JException {
  constructor(message: string, readonly option: string, readonly statement: string,
    readonly reason: 'unknown' | 'invalid' | 'not_implemented') {
    super(message);
    this.name = 'TextOptionError';
  }
}

/** The option names of WD section 10, in its order. docx4j `TextOptions.NAMES`. */
export const TEXT_OPTION_NAMES: readonly string[] = Object.freeze([
  'view', 'include', 'text_boxes', 'notes', 'note_references', 'comments', 'hidden', 'field_codes',
  'fields', 'toc', 'placeholders', 'bindings', 'list_labels', 'charts', 'smartart', 'embedded', 'altchunks',
  'alt_text', 'metadata', 'symbols', 'math', 'tables', 'markup_detail', 'page_markers', 'layout',
  'warnings',
]);

const INCLUDES: readonly TextInclude[] = ['headers', 'footers', 'comments', 'glossary'];

/**
 * What a text includes and how it is written: the specification's section 10 options (REQ-1001), as
 * docx4j's `TextOptions` has them. Mutable; `Texts.of` reads it once.
 *
 * Every option name and value of WD section 10 is accepted by `set`. Some values are not implemented,
 * and `Texts.of` refuses them with a `TextOptionError` naming the statement rather than ignoring them
 * (REQ-1001): `math: unicodemath` and `symbols: unicode`, as docx4j refuses them (no converter, no glyph
 * tables), and, until CR-012 phase 5, `fields: update` (REQ-506) and `embedded: true` (REQ-512). The
 * defaults `charts`, `smartart` and `altchunks` are accepted, and until phase 5 a chart or SmartArt
 * contributes nothing and an `altChunk` is not expanded, raising `altchunk.not_expanded` (CR-012
 * section 5).
 */
export class TextOptions {
  view: TextModelView = 'accepted';
  private readonly includeSet = new Set<TextInclude>();
  textBoxes = true;
  notes = true;
  noteReferences = true;
  private commentsOption = false;
  hidden = false;
  fieldCodes = false;
  fields: TextFields = 'cached';
  /**
   * Whether a table of contents is read: false by default (WD REQ-520; docx4j CR-046 D12). A table of
   * contents is the result of a TOC field; under `toc: false` nothing of the field is read, its code
   * included under `field_codes`, and a paragraph with nothing else in it is no paragraph.
   */
  toc = false;
  placeholders = true;
  bindings: TextBindings = 'bound';
  listLabels = false;
  charts = true;
  smartart = true;
  embedded = false;
  altchunks = true;
  altText = false;
  metadata = false;
  symbols: TextSymbols = 'asis';
  math: TextMath = 'latex';
  tables: TextTables = 'auto';
  markupDetail: TextMarkupDetail = 'type';
  /**
   * `page_markers` (REQ-708): `[page n]` at each `w:lastRenderedPageBreak`. The markers are where Word's
   * pages began when it last laid the document out: they are current only when the package was last
   * paginated (saved by Word) after its last edit, and an extractor cannot tell whether it was.
   */
  pageMarkers = false;
  layout: TextLayout = 'none';
  warnings = true;
  /** The describer images take their text from (REQ-514). */
  imageDescriber: ImageDescriber = AltTextDescriber.INSTANCE;

  /** Options from the specification's names and values (REQ-1001); those not named keep their defaults. */
  static of(init?: TextOptionsInit | TextOptions): TextOptions {
    if (init instanceof TextOptions) return init;
    const o = new TextOptions();
    if (init === undefined || init === null) return o;
    for (const [name, value] of Object.entries(init)) {
      if (value === undefined) continue;
      if (name === 'imageDescriber') o.imageDescriber = (value as ImageDescriber | null) ?? AltTextDescriber.INSTANCE;
      else o.set(name, value);
    }
    return o;
  }

  /** Sets one option by its specification name and value; throws `TextOptionError` for one it does not define. */
  set(name: string, value: unknown): this {
    const n = (name ?? '').trim();
    switch (n) {
      case 'view': this.view = viewOf(value); break;
      case 'include': this.setInclude(value); break;
      case 'text_boxes': this.textBoxes = bool(n, value, 'REQ-518'); break;
      case 'notes': this.notes = bool(n, value, 'REQ-503'); break;
      case 'note_references': this.noteReferences = bool(n, value, 'REQ-503'); break;
      case 'comments': this.commentsOption = bool(n, value, 'REQ-504'); break;
      case 'hidden': this.hidden = bool(n, value, 'REQ-507'); break;
      case 'field_codes': this.fieldCodes = bool(n, value, 'REQ-505'); break;
      case 'fields': this.fields = choice(n, value, ['cached', 'update'], 'REQ-506'); break;
      case 'toc': this.toc = bool(n, value, 'REQ-520'); break;
      case 'placeholders': this.placeholders = bool(n, value, 'REQ-508'); break;
      case 'bindings': this.bindings = choice(n, value, ['bound', 'surface'], 'REQ-509'); break;
      case 'list_labels': this.listLabels = bool(n, value, 'REQ-510'); break;
      case 'charts': this.charts = bool(n, value, 'REQ-511'); break;
      case 'smartart': this.smartart = bool(n, value, 'REQ-511'); break;
      case 'embedded': this.embedded = bool(n, value, 'REQ-512'); break;
      case 'altchunks': this.altchunks = bool(n, value, 'REQ-513'); break;
      case 'alt_text': this.altText = bool(n, value, 'REQ-514'); break;
      case 'metadata': this.metadata = bool(n, value, 'REQ-515'); break;
      case 'symbols': this.symbols = choice(n, value, ['asis', 'unicode'], 'REQ-601'); break;
      case 'math': this.math = choice(n, value, ['latex', 'unicodemath'], 'REQ-604'); break;
      case 'tables': this.tables = choice(n, value, ['auto', 'tsv', 'grid'], 'REQ-605'); break;
      case 'markup_detail': this.markupDetail = choice(n, value, ['type', 'author', 'author_date'], 'REQ-702'); break;
      case 'page_markers': this.pageMarkers = bool(n, value, 'REQ-708'); break;
      case 'layout': this.layout = choice(n, value, ['none', 'indent'], 'REQ-709'); break;
      case 'warnings': this.warnings = bool(n, value, 'REQ-801'); break;
      default:
        throw new TextOptionError(`Unknown text option '${String(name)}'; the options are ${TEXT_OPTION_NAMES.join(', ')} (WD REQ-1001)`,
          String(name), 'REQ-1001', 'unknown');
    }
    return this;
  }

  private setInclude(value: unknown): void {
    this.includeSet.clear();
    const items: unknown[] = Array.isArray(value) ? value
      : value === undefined || value === null ? [] : String(value).split(/[,\s]+/).filter((s) => s !== '');
    for (const item of items) this.includeSet.add(choice('include', item, INCLUDES, 'REQ-501'));
  }

  /** Adds stories to `include` (docx4j's builder `include(Include...)`). */
  include(...stories: TextInclude[]): this {
    for (const s of stories) this.includeSet.add(choice('include', s, INCLUDES, 'REQ-501'));
    return this;
  }

  /** The stories `include` names; `comments: true` adds `comments`. */
  get includes(): ReadonlySet<TextInclude> {
    const out = new Set(this.includeSet);
    if (this.commentsOption) out.add('comments');
    return out;
  }

  /** `comments: true` or `include: comments`. */
  get comments(): boolean {
    return this.commentsOption || this.includeSet.has('comments');
  }

  set comments(value: boolean) {
    this.commentsOption = value;
  }

  /** A copy, for a reading one level down or a first reading. */
  copy(): TextOptions {
    const o = new TextOptions();
    Object.assign(o, this);
    (o as unknown as { includeSet: Set<TextInclude> }).includeSet = new Set(this.includeSet);
    return o;
  }

  /** The options as the specification names them, with their values, in section 10's order. */
  toMap(): Record<string, unknown> {
    return {
      view: this.view, include: [...this.includeSet], text_boxes: this.textBoxes, notes: this.notes,
      note_references: this.noteReferences, comments: this.commentsOption, hidden: this.hidden,
      field_codes: this.fieldCodes, fields: this.fields, toc: this.toc, placeholders: this.placeholders,
      bindings: this.bindings, list_labels: this.listLabels, charts: this.charts, smartart: this.smartart,
      embedded: this.embedded, altchunks: this.altchunks, alt_text: this.altText, metadata: this.metadata,
      symbols: this.symbols, math: this.math, tables: this.tables, markup_detail: this.markupDetail,
      page_markers: this.pageMarkers, layout: this.layout, warnings: this.warnings,
    };
  }

  toJSON(): Record<string, unknown> {
    return this.toMap();
  }

  /**
   * Throws for a value this package does not honour (REQ-1001: documented, and refused rather than
   * ignored): the two docx4j refuses, and the two whose statements wait on CR-012 phase 5.
   */
  checkImplemented(): void {
    if (this.math === 'unicodemath') {
      throw new TextOptionError('Not implemented (WD REQ-1001): math: unicodemath (no OMML to UnicodeMath converter)', 'math', 'REQ-604', 'not_implemented');
    }
    if (this.symbols === 'unicode') {
      throw new TextOptionError('Not implemented (WD REQ-1001): symbols: unicode (no glyph tables for Symbol and Wingdings)', 'symbols', 'REQ-601', 'not_implemented');
    }
    if (this.fields === 'update') {
      throw new TextOptionError('Not implemented yet (WD REQ-1001): fields: update (REQ-506, CR-012 phase 5a)', 'fields', 'REQ-506', 'not_implemented');
    }
    if (this.embedded) {
      throw new TextOptionError('Not implemented yet (WD REQ-1001): embedded: true (REQ-512, CR-012 phase 5d)', 'embedded', 'REQ-512', 'not_implemented');
    }
  }
}

function viewOf(value: unknown): TextModelView {
  const n = String(value ?? '').trim().toLowerCase();
  if (n === 'current') return 'accepted';
  if (n === 'accepted' || n === 'original' || n === 'markup') return n;
  throw new TextOptionError(`view: '${String(value)}' is not accepted, original, markup or current (WD REQ-401)`, 'view', 'REQ-401', 'invalid');
}

function bool(name: string, value: unknown, statement: string): boolean {
  if (typeof value === 'boolean') return value;
  const s = String(value ?? '').trim().toLowerCase();
  if (s === 'true') return true;
  if (s === 'false') return false;
  throw new TextOptionError(`${name}: '${String(value)}' is not true or false (WD REQ-1001)`, name, statement, 'invalid');
}

function choice<T extends string>(name: string, value: unknown, values: readonly T[], statement: string): T {
  const s = String(value ?? '').trim().toLowerCase();
  const found = values.find((v) => v === s);
  if (found !== undefined) return found;
  throw new TextOptionError(`${name}: '${String(value)}' is not one of ${values.join(', ')} (WD REQ-1001)`, name, statement, 'invalid');
}
