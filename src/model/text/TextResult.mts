// CR-012 phase 2: the block model (WD sections 7.4 and 8), docx4j's org.docx4j.model.text result classes:
// TextResult, TextBlock, TextRun, TextRow, TextCell, TextField, TextNumbering, TextRevision, TextWarning,
// TextComment, TextImage, ImageDescription, ImageDescriber and AltTextDescriber. docx4j's enum values are
// kept verbatim as strings ('PARAGRAPH', 'TEXT_BOX', 'NOTE_REFERENCE'), as CR-011 kept its change kinds, so
// that a block model passed as JSON reads the same from either implementation.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Part } from '../../parts/Part.mjs';
import type { TextOptions } from './TextOptions.mjs';

/** What a block is (docx4j `TextBlock.Kind`). */
export type TextBlockKind = 'PARAGRAPH' | 'TABLE' | 'CHART';

/**
 * The story a block is in (WD section 5.1; docx4j `TextBlock.Story`). SmartArt's text and an embedded
 * package's content follow their anchoring paragraph but are kept apart from its story; an `altChunk`'s
 * content is in the story the chunk is in.
 */
export type TextStory = 'MAIN' | 'TEXT_BOX' | 'FOOTNOTE' | 'ENDNOTE' | 'HEADER' | 'FOOTER' | 'COMMENT' | 'GLOSSARY' | 'SMARTART' | 'EMBEDDED';

/** What a run is (docx4j `TextRun.Kind`). */
export type TextRunKind =
  /** Document text: `w:t`, tabs and breaks as REQ-601 maps them, a field's result. */
  | 'TEXT'
  /** A field's instruction, present only under `field_codes: true` (REQ-505). */
  | 'FIELD_CODE'
  /** A footnote or endnote reference in the story (REQ-503). */
  | 'NOTE_REFERENCE'
  /** The note's own mark at its start (`w:footnoteRef`, or added where it has none). */
  | 'NOTE_MARK'
  /** A comment's reference, where the comment is rendered in plain text (REQ-504). */
  | 'COMMENT_REFERENCE'
  /** A list label from the numbering model, under `list_labels: true` (REQ-510). */
  | 'LIST_LABEL'
  /** A picture, with its description where the describer gives one (REQ-514). */
  | 'IMAGE'
  /** An embedded object, OLE link or ActiveX control (`w:object`; REQ-512): no text. */
  | 'OBJECT'
  /** A `w:lastRenderedPageBreak`: where the page began when Word last laid the document out (REQ-708). */
  | 'PAGE_MARKER';

/** `footnote` or `endnote`, as a note reference, mark or block says whose it is. */
export type NoteType = 'footnote' | 'endnote';

// The tree objects a block or run was read from: kept out of the block model's JSON (REQ-706 holds no
// tree objects; CR-012 section 4.4), so held here rather than as properties.
const blockSources = new WeakMap<TextBlock, object>();
const runSources = new WeakMap<TextRun, TextRunSource>();

/** Where a run's text was read from: its run and its paragraph (CR-012 section 4.4). */
export interface TextRunSource {
  run: wml.R;
  paragraph: wml.P;
}

/**
 * A run's revision in the `markup` view (WD REQ-706): `w:ins` and `w:moveTo` are insertions, `w:del`
 * and `w:moveFrom` deletions, with their author and date as written. docx4j `TextRevision`.
 */
export class TextRevision {
  constructor(
    readonly type: 'INSERTION' | 'DELETION',
    readonly author: string | undefined,
    /** The date as the package wrote it (`w:date`), or undefined. */
    readonly date: string | undefined,
  ) {}

  equals(o: TextRevision | undefined): boolean {
    return o !== undefined && this.type === o.type && this.author === o.author && this.date === o.date;
  }

  toString(): string {
    return this.type + (this.author === undefined ? '' : ` by ${this.author}`) + (this.date === undefined ? '' : ` ${this.date}`);
  }
}

/**
 * An image's description (WD REQ-514) and where it came from: the document's author, or a describer
 * that generated it. Plain text writes it `[image: ...]` or `[image, generated: ...]` under
 * `alt_text: true` (REQ-707). docx4j `ImageDescription`.
 */
export class ImageDescription {
  constructor(readonly text: string, readonly provenance: 'AUTHOR' | 'GENERATED') {}

  /** An author's description. */
  static author(text: string): ImageDescription {
    return new ImageDescription(text, 'AUTHOR');
  }

  /** A generated description. */
  static generated(text: string): ImageDescription {
    return new ImageDescription(text, 'GENERATED');
  }

  toString(): string {
    return this.provenance === 'GENERATED' ? `generated: ${this.text}` : this.text;
  }
}

/**
 * A picture as the document gives it, for an `ImageDescriber` (WD REQ-514): its alt text
 * (`wp:docPr/@descr`, or VML's `v:shape/@alt`), its name and title, whether it is marked decorative
 * (`adec:decorative`), and its image part where the package holds one. The title is never read as text.
 * docx4j `TextImage`.
 */
export class TextImage {
  constructor(
    /** The alt text: `wp:docPr/@descr` (or a nested picture's `pic:cNvPr/@descr`), VML `v:shape/@alt`. */
    readonly descr: string | undefined,
    /** `wp:docPr/@name` (`pic:cNvPr/@name`); undefined for a VML picture. */
    readonly name: string | undefined,
    /** `wp:docPr/@title`. */
    readonly title: string | undefined,
    /** True for a picture marked decorative (`adec:decorative val="1"` in `wp:docPr`'s extensions). */
    readonly decorative: boolean,
    /** True for a VML picture (`v:shape` with `v:imagedata`). */
    readonly vml: boolean,
    /**
     * The image part (`a:blip/@r:embed`, else its `asvg:svgBlip`'s; `v:imagedata/@r:id`), or undefined
     * for a linked or missing one.
     */
    readonly part: Part | undefined,
    /** The tree object the picture is read from: a `wp:inline`, `wp:anchor`, `pic:pic` or `v:shape`. */
    readonly source: object | undefined,
  ) {}

  toJSON(): Record<string, unknown> {
    return { descr: this.descr, name: this.name, title: this.title, decorative: this.decorative, vml: this.vml, part: this.part?.partName.name };
  }
}

/**
 * Gives an image its text (WD REQ-514): one describer per extraction, `AltTextDescriber` by default. A
 * host may replace or wrap the default (a captioning model); what it generates it marks `GENERATED`.
 * An SVG's `<text>` and the glyph runs of EMF and WMF are not the document's text. docx4j `ImageDescriber`.
 */
export type ImageDescriber = (image: TextImage) => ImageDescription | undefined;

/**
 * The default `ImageDescriber` (WD REQ-514): the author's alt text, `wp:docPr/@descr`, else the
 * picture's `@name`; for VML, `v:shape/@alt`; nothing for a picture marked decorative.
 * docx4j `AltTextDescriber`.
 */
export const AltTextDescriber = {
  INSTANCE: ((image: TextImage): ImageDescription | undefined => {
    if (!image || image.decorative) return undefined;
    if (image.descr) return ImageDescription.author(image.descr);
    if (image.name) return ImageDescription.author(image.name);
    return undefined;
  }) as ImageDescriber,
};

/**
 * A piece of a block's content (WD REQ-706): document text, or a typed item whose text, if any, the
 * serialiser adds (a note reference is written `[n]` in plain text, REQ-707; it is a field of the run
 * here, never text). docx4j `TextRun`.
 *
 * `source` and `officeOffset` are this package's (CR-012 section 4.4; WD REQ-202's emulated-API
 * exemption): where the run's text was read from, and where it lies in its paragraph's `Paragraph.text`.
 */
export class TextRun {
  /** True for hidden text included under `hidden: true` (REQ-507). */
  hidden = false;
  /** True for a content control's placeholder text (REQ-508). */
  placeholder = false;
  /** For a note reference or mark: whose. */
  noteType: NoteType | undefined;
  /** For a note reference or mark: its number in reference order, from 1; 0 for a custom mark. */
  noteNumber = 0;
  /** For a note reference or mark: the note's `w:id`. */
  noteId: string | undefined;
  /** For a comment reference: the comment's `w:id`. */
  commentId: string | undefined;
  /** For an image: what the document says of it. */
  image: TextImage | undefined;
  /** For an image: the describer's description; undefined for none (a decorative picture's). */
  description: ImageDescription | undefined;
  /** For an object: its ProgID (`o:OLEObject/@ProgID`). */
  progId: string | undefined;
  /** For a page marker: the number of the page that begins there, from 2. */
  pageNumber = 0;
  /** Bold where its paragraph's style is not: the run's `w:b`, else its styles' (CR-046 phase 3). */
  bold = false;
  /** Italic where its paragraph's style is not. */
  italic = false;
  /** Struck through where its paragraph's style is not. */
  strike = false;
  /** In a monospaced font where its paragraph's style is not in one. */
  monospace = false;
  /** For text in a link: its target (a `w:hyperlink`'s, or a HYPERLINK field's, REQ-703). */
  link: string | undefined;
  /** True for an equation (REQ-604). */
  math = false;
  /** True for a display equation (`m:oMathPara`). */
  displayMath = false;
  /**
   * Where the run's text starts in its paragraph's `Paragraph.text` (the Office JS reader, CR-012 D2),
   * where it contributes there; within the run the two texts agree character for character. Undefined
   * for a run only the model writes: a note reference or mark, a list label, a page marker, a field code,
   * an image, a field result or bound value computed here, text `Paragraph.text` does not hold.
   */
  officeOffset: number | undefined;

  constructor(
    readonly kind: TextRunKind,
    /** The text: the document's for `TEXT`, `FIELD_CODE` and `LIST_LABEL`; empty for the others. */
    readonly text: string,
    /** The run's revision in the `markup` view; undefined in the other views and for an unrevised run. */
    readonly revision: TextRevision | undefined,
  ) {}

  /** Where the run's text was read from (not in the JSON). */
  get source(): TextRunSource | undefined {
    return runSources.get(this);
  }

  set source(value: TextRunSource | undefined) {
    if (value) runSources.set(this, value); else runSources.delete(this);
  }

  /** Whether two runs' text can be one run: the same formatting, link and equation. */
  sameFormat(o: TextRun): boolean {
    return this.bold === o.bold && this.italic === o.italic && this.strike === o.strike && this.monospace === o.monospace
      && this.link === o.link && this.math === o.math && this.displayMath === o.displayMath;
  }

  copyFormat(from: TextRun): void {
    this.bold = from.bold;
    this.italic = from.italic;
    this.strike = from.strike;
    this.monospace = from.monospace;
    this.link = from.link;
    this.math = from.math;
    this.displayMath = from.displayMath;
  }

  toString(): string {
    return this.kind + (this.text === '' ? '' : ` '${this.text}'`) + (this.revision ? ` [${this.revision.toString()}]` : '');
  }
}

/**
 * A field in a block (WD REQ-505): its instruction, and its cached result, undefined where it has none.
 * docx4j `TextField`.
 */
export class TextField {
  constructor(
    /** The instruction (`w:instrText`, `w:fldSimple/@instr`), as written. */
    readonly instr: string,
    /** The cached result's text in the view; undefined for a field with none. */
    readonly result: string | undefined,
    /** True when the result was computed under `fields: update` (REQ-506), rather than cached. */
    readonly updated = false,
  ) {}

  toString(): string {
    return `field{${this.instr}${this.updated ? ' => ' : ' -> '}${String(this.result)}}`;
  }
}

/**
 * A paragraph's numbering, as its effective `w:numPr` and the numbering definitions give it: what
 * Markdown's list markers are made from (WD REQ-510, REQ-703). The label itself is the block's
 * `listLabel`, under `list_labels: true`. docx4j `TextNumbering`.
 */
export class TextNumbering {
  /** The count the paragraph's level shows, as the list labels count it (the 3 of "1.3."), or undefined. */
  value: number | undefined;
  /** The list the paragraph belongs to: its `w:abstractNum`'s id, or undefined. */
  abstractNumId: string | undefined;

  constructor(
    /** The `w:numId`. */
    readonly numId: string,
    /** The level, `w:ilvl`: 0 to 8. */
    readonly level: number,
    /** The level's `w:numFmt` value, or undefined where the definitions don't give one. */
    readonly format: string | undefined,
    /** The level's `w:lvlText` value, or undefined. */
    readonly levelText: string | undefined,
    /** The level's `w:start`: 1 where it has none. */
    readonly start: number,
  ) {}

  /** True for a bullet level. */
  get bullet(): boolean {
    return this.format === 'bullet';
  }

  /** True for a level that is a label rather than a number: not a bullet, and its format `none` or its text with no `%n`. */
  get label(): boolean {
    if (this.bullet || (this.format === undefined && this.levelText === undefined)) return false;
    return this.format === 'none' || this.levelText === undefined || !this.levelText.includes('%');
  }

  toString(): string {
    return `${this.numId}/${this.level} ${String(this.format)} '${String(this.levelText)}' from ${this.start}${this.value === undefined ? '' : ` = ${this.value}`}`;
  }
}

/**
 * A table row of the block model: its cells, and whether it is a header row (`w:tblHeader`).
 * docx4j `TextRow`.
 */
export class TextRow {
  readonly cells: TextCell[] = [];
  /** True for a row marked `w:tblHeader`. */
  header = false;

  constructor(readonly address: string) {}
}

/**
 * A table cell of the block model: its blocks (paragraphs, and nested tables), how many grid columns it
 * spans, and whether it is in a vertical merge. docx4j `TextCell`.
 */
export class TextCell {
  readonly blocks: TextBlock[] = [];
  /** `w:gridSpan`, 1 by default. */
  gridSpan = 1;
  /** True for a cell continuing a vertical merge (`w:vMerge` without `restart`). */
  mergeContinue = false;
  /** True for a cell starting a vertical merge (`w:vMerge w:val="restart"`). */
  mergeRestart = false;

  constructor(readonly address: string) {}
}

/**
 * A block of the block model (WD REQ-706): a paragraph, a table whose cells hold blocks, or a chart. A
 * paragraph's text is its runs' in the view; what an extractor adds in plain text (note numbers, a
 * table-format note, an image's bracket) is not in it, but carried as fields (REQ-707). docx4j `TextBlock`.
 *
 * Addresses are CR-002 section 3.3's ordinal form, on the unresolved document (REQ-405, REQ-518):
 * `body/3`, `body/4/1/0/2` (a table's row 1, cell 0, block 2), `header:rId3/0`, `footnote:2/0`,
 * `comment:5/0`, and what follows an anchoring paragraph is addressed from it, each kind counted apart
 * (REQ-519): `body/3/textbox:0/1` (legacy WordArt too), `body/3/chart:0`, `body/3/smartart:0/1`,
 * `body/3/object:0/...`, `body/7/altchunk:0/...`; the glossary's `glossary:0/1`. A paragraph joined to
 * the next in the `accepted` view (a deleted mark) takes the first paragraph's address and `w14:paraId`.
 */
export class TextBlock {
  /** The paragraph's `w14:paraId` where Word wrote one. */
  paraId: string | undefined;
  /** The paragraph's style id (`w:pStyle`). */
  style: string | undefined;
  readonly runs: TextRun[] = [];
  /** The fields whose result or instruction is in the block (REQ-505). */
  readonly fields: TextField[] = [];
  /** A table's rows; empty for a paragraph. */
  readonly rows: TextRow[] = [];
  /** True for a hidden paragraph included under `hidden: true`: its runs and mark are hidden (REQ-507). */
  hidden = false;
  /** How many revisions (run-level revisions and paragraph marks) the block's paragraphs hold, in any view. */
  revisions = 0;
  /** The effective left indent in twips. */
  indentLeft: number | undefined;
  /** The effective hanging indent in twips. */
  indentHanging: number | undefined;
  /** The effective first-line indent in twips. */
  indentFirstLine: number | undefined;
  /** The effective alignment (`w:jc`'s value). */
  alignment: string | undefined;
  /** True when every text run is in a monospaced font, and at least one is. */
  monospace = false;
  /** The paragraph's list label from the numbering model, under `list_labels: true`. */
  listLabel: string | undefined;
  /** The effective outline level (`w:outlineLvl`), 0 to 8; undefined for body text. */
  outlineLevel: number | undefined;
  /** The paragraph's numbering (its effective `w:numPr` with a `w:numId` other than 0). */
  numbering: TextNumbering | undefined;
  /** The effective `w:contextualSpacing`, undefined where neither the paragraph nor its styles set it. */
  contextualSpacing: boolean | undefined;
  /** True when the paragraph has a bottom border of its own (`w:pBdr/w:bottom` in its `w:pPr`). */
  bottomBorder = false;
  /** For a note's block: whose. */
  noteType: NoteType | undefined;
  /** For a note's block: the note's number in reference order; else 0. */
  noteNumber = 0;
  /** For a note's block: the note's `w:id`. */
  noteId: string | undefined;
  /** The section the block ends or is in (main story only). */
  section = 0;
  /** For a table: always written in REQ-605's simple form (an embedded workbook's sheet). */
  simpleForm = false;
  /** For a chart: its title as Office shows it. */
  chartTitle: string | undefined;
  /** For a chart: its description (docx4j CR-047's `ChartDescription`); CR-012 phase 5c. */
  chart: { note: string; kind?: string; axes?: unknown[]; secondarySeries?: string[] } | undefined;
  /** For a chart: its table's column headings. */
  chartHeader: string[] = [];
  /** For a chart: its table's rows. */
  chartRows: string[][] = [];

  constructor(
    readonly kind: TextBlockKind,
    /** The story; not readonly, since an embedded package's or a chunk's content takes the story it is read into. */
    public story: TextStory,
    readonly address: string,
  ) {}

  /**
   * What the block was read from (CR-012 section 4.4; not in the JSON): a paragraph's `w:p` (a joined
   * block's first, as its paraId is, REQ-405), a table's `w:tbl`, a bound control's `w:sdt`; undefined
   * for a block the model makes (a chart, a WordArt line).
   */
  get source(): object | undefined {
    return blockSources.get(this);
  }

  set source(value: object | undefined) {
    if (value) blockSources.set(this, value); else blockSources.delete(this);
  }

  /** True when the block holds placeholder text (REQ-508). */
  get placeholder(): boolean {
    return this.runs.some((r) => r.placeholder);
  }

  /** True for a block in a text box (REQ-518). */
  get textbox(): boolean {
    return this.story === 'TEXT_BOX';
  }

  /** True for a paragraph whose only content is one or more pictures marked decorative (REQ-514). */
  get decorative(): boolean {
    let decorative = false;
    for (const r of this.runs) {
      if (r.kind === 'IMAGE') {
        if (!r.image || !r.image.decorative) return false;
        decorative = true;
      } else if (r.kind === 'TEXT' && r.text !== '') {
        return false;
      }
    }
    return decorative;
  }

  /**
   * The block's text in the view: its text runs (and field codes, where included); for a table, its
   * rows a line apart, their cells a tab apart. No list label, note number or comment.
   */
  get text(): string {
    if (this.kind === 'CHART') {
      let out = this.chartTitle ?? '';
      if (this.chartHeader.length > 0) out += (out === '' ? '' : '\n') + this.chartHeader.join('\t');
      for (const row of this.chartRows) out += `\n${row.join('\t')}`;
      return out;
    }
    if (this.kind === 'TABLE') {
      return this.rows.map((row) => row.cells.map((cell) => cell.blocks.map((b) => b.text).join(' ')).join('\t')).join('\n');
    }
    let out = '';
    for (const r of this.runs) if (r.kind === 'TEXT' || r.kind === 'FIELD_CODE') out += r.text;
    return out;
  }

  /** The block model's JSON (REQ-706): the fields, the text and the flags; never a tree object. */
  toJSON(): Record<string, unknown> {
    const out: Record<string, unknown> = { ...(this as unknown as Record<string, unknown>), text: this.text, placeholder: this.placeholder, textbox: this.textbox, decorative: this.decorative };
    if (this.kind !== 'CHART') { delete out.chartTitle; delete out.chart; delete out.chartHeader; delete out.chartRows; }
    return out;
  }

  toString(): string {
    return `${this.address} ${this.story} '${this.text}'`;
  }
}

/**
 * A comment (WD REQ-504): its author, date and blocks, the comment it replies to, and whether it is
 * resolved (`w15:done`). Plain text renders a comment at its reference, replies after it. docx4j `TextComment`.
 */
export class TextComment {
  author: string | undefined;
  initials: string | undefined;
  /** The date as the package wrote it. */
  date: string | undefined;
  /** The `w:id` of the comment this one replies to (`w15:paraIdParent`). */
  parentId: string | undefined;
  /** True when the comment is resolved (`w15:done`). */
  resolved = false;
  readonly blocks: TextBlock[] = [];
  /** The replies, in the comments part's order. */
  readonly replies: TextComment[] = [];

  constructor(/** The comment's `w:id`. */ readonly id: string) {}

  /** The comment's text: its blocks a line apart. */
  get text(): string {
    return this.blocks.map((b) => b.text).join('\n');
  }

  toJSON(): Record<string, unknown> {
    return { id: this.id, author: this.author, initials: this.initials, date: this.date, parentId: this.parentId, resolved: this.resolved, blocks: this.blocks, replies: this.replies.map((r) => r.id) };
  }
}

/**
 * Where a result is not the whole document (WD section 8, REQ-801): a kind, the address it is at where
 * there is one, and its selector: for `mce.branches_differ` the name of the branch selected, for
 * `altchunk.not_expanded` the chunk's content type, for `content.not_read` and `image.not_described` the
 * part's name. Never the document's text (REQ-802). docx4j `TextWarning`.
 */
export class TextWarning {
  /** Both branches of an `mc:AlternateContent` hold text, and the texts differ (REQ-307). */
  static readonly MCE_BRANCHES_DIFFER = 'mce.branches_differ';
  /** A field with no cached result (REQ-505). */
  static readonly FIELD_NO_RESULT = 'field.no_result';
  /** A data-bound control whose surface differs from its bound value (REQ-509). */
  static readonly SDT_BINDING_DIFFERS = 'sdt.binding_differs';
  /** An `altChunk` not expanded (REQ-513). */
  static readonly ALTCHUNK_NOT_EXPANDED = 'altchunk.not_expanded';
  /** A chart, a SmartArt data part, or an embedded package or a part inside it, that can't be read. */
  static readonly CONTENT_NOT_READ = 'content.not_read';
  /** An image the `ImageDescriber` failed to describe (REQ-514). */
  static readonly IMAGE_NOT_DESCRIBED = 'image.not_described';

  constructor(
    readonly kind: string,
    /** The address of the block it is in (REQ-518's grammar), or undefined. */
    readonly address?: string,
    /** For `mce.branches_differ`, the selected branch (`mc:Choice Requires="wps"`, `mc:Fallback`). */
    readonly branch?: string,
    /** For `altchunk.not_expanded`, the chunk's content type as the package declares it. */
    readonly contentType?: string,
    /** For `content.not_read` and `image.not_described`, the part's name. */
    readonly partName?: string,
  ) {}

  /** The selector of REQ-801: the branch, the content type or the part's name, whichever the kind has. */
  get selector(): string | undefined {
    return this.branch ?? this.contentType ?? this.partName;
  }

  /** The identity a result keeps each warning once by. */
  get key(): string {
    return [this.kind, this.address, this.branch, this.contentType, this.partName].map((v) => v ?? '\u0000').join('\u0001');
  }

  toJSON(): Record<string, unknown> {
    return { kind: this.kind, address: this.address, selector: this.selector };
  }

  toString(): string {
    return this.kind + (this.address === undefined ? '' : ` at ${this.address}`) + (this.branch === undefined ? '' : ` (${this.branch})`)
      + (this.contentType === undefined ? '' : ` (${this.contentType})`) + (this.partName === undefined ? '' : ` (${this.partName})`);
  }
}

/**
 * A text in the block model (WD sections 7.4 and 8): the blocks in reading order, the comments, the
 * warnings and the metadata. Plain text (`PlainText`) and the other serialisations are written from it
 * (REQ-202). docx4j `TextResult`.
 */
export class TextResult {
  /** The metadata's key for the settings' `w:updateFields` (REQ-506, REQ-801): a boolean. */
  static readonly FIELDS_UPDATE_ON_OPEN = 'fields_update_on_open';
  /** The metadata's key for the core properties (REQ-515): each property's element name to its value. */
  static readonly CORE = 'core';
  /** The metadata's key for the application properties (REQ-515), the statistics left out. */
  static readonly APP = 'app';
  /** The metadata's key for the custom properties (REQ-515): each property's name to its typed value. */
  static readonly CUSTOM = 'custom';

  constructor(
    /** The blocks in reading order (REQ-501, REQ-502, REQ-518, REQ-606). Comments are not among them. */
    readonly blocks: readonly TextBlock[],
    /** The comments the view includes, by `w:id`, in the comments part's order (REQ-504). */
    readonly comments: ReadonlyMap<string, TextComment>,
    /** The warnings raised (REQ-801), never with the document's text (REQ-802). */
    readonly warnings: readonly TextWarning[],
    /**
     * The result's metadata (REQ-801, REQ-515), whatever the options: `fields_update_on_open`, and the
     * document properties under `core`, `app` and `custom`, each a record, present when the package has
     * that part. Property values are document text, in the text stream only under `metadata: true`.
     */
    readonly metadata: Readonly<Record<string, unknown>>,
    /** The options the result was made with. */
    readonly options: TextOptions,
  ) {}

  /**
   * Several results as one, in the order given (CR-012 section 4.4): the blocks one after another, the
   * comments together, each warning once, the metadata and options the first's. What the editor
   * assembles from its per-part calls under one `TextContext`.
   */
  static concat(...results: TextResult[]): TextResult {
    if (results.length === 0) throw new RangeError('TextResult.concat needs a result');
    const blocks: TextBlock[] = [];
    const comments = new Map<string, TextComment>();
    const warnings: TextWarning[] = [];
    const seen = new Set<string>();
    const metadata: Record<string, unknown> = {};
    for (const r of results) {
      blocks.push(...r.blocks);
      for (const [id, c] of r.comments) if (!comments.has(id)) comments.set(id, c);
      for (const w of r.warnings) if (!seen.has(w.key)) { seen.add(w.key); warnings.push(w); }
      for (const [k, v] of Object.entries(r.metadata)) if (!(k in metadata)) metadata[k] = v;
    }
    return new TextResult(blocks, comments, warnings, metadata, results[0]!.options);
  }

  toJSON(): Record<string, unknown> {
    return { blocks: this.blocks, comments: Object.fromEntries(this.comments), warnings: this.warnings, metadata: this.metadata, options: this.options };
  }
}
