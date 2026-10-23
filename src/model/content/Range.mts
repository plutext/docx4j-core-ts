import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { Font, MixedFont, readingOf, markHolderOf, type RPrHolder } from './Font.mjs';
import { runFontSelectorOf } from '../fonts/lookup.mjs';
import type { Paragraph, FormattingOptions } from './Paragraph.mjs';
import type { Body, BlockElement } from './Body.mjs';
import { type Element, typeNameOf, linkParents, runItemsOf, runsOf, revisionKindOf, WHOLE, type SpanEdges, type RevisionHolder, type TextSegment, type TextViewOptions } from './tree.mjs';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { NotSupportedError } from '../../office-js/errors.mjs';
import { ContentControl, assertControlHolds, controlsOfTypes, BLOCK_CONTROL, type ContentControlType, type ContentControlOptions } from './ContentControl.mjs';
import { sdt as sdtOf, nextSdtId, applyRunOptions, t as textElement } from '@docx4j/generated-objects-ts/builders/wml';
import * as f from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { hyperlink as hyperlinkOf } from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { Namespaces } from '../../parts/Namespaces.mjs';
import { controlIdScope, sdtKindFor, declareSdtExtensions } from '../customxml/insert.mjs';
import { contentOf } from './ooxml.mjs';
import { markHolderInserted } from './tracking.mjs';
import { officeMatches, expandReplacement, paragraphMarkPieces, refuseMarkReplacement, type SearchOptions } from './search.mjs';
import { commentApi, type CommentContent, type CommentOptions } from './comments.mjs';
import type { Comment } from './Comment.mjs';
import type { TrackedChange } from './TrackedChange.mjs';
import type { Table } from './Table.mjs';
import type { TableCell } from './Table.mjs';
import { storyLayout, storyOf, type StoryLayout } from './story.mjs';
import { styleNameOf, builtInOf } from './styles.mjs';
import { GeneralException } from './generalException.mjs';
import { extentOf, assertTrackable, deleteSpan, firstCharacterOf, placeReplacement, insertParagraphAt, formatSpan, type Extent } from './spanEdits.mjs';

/**
 * A subset of Office JS `Word.Range`: a span of text from an offset in one paragraph to an offset in the same
 * paragraph or a later one of the same story, [start, end), `start` in `paragraph`'s text and `end` in
 * `endParagraph`'s. Positions are recomputed from the tree on every use, so a range stays valid across edits
 * made through it; edits made elsewhere may shift it.
 *
 * A range within one paragraph is what it has been since CR-002: `endParagraph` is `paragraph`. A range across
 * paragraphs (CR-014 phase 1) reads as Office JS's does (check 52): its `text` in `body.text`'s form between
 * its ends, `paragraphs` every paragraph it touches (a table's cells' among them), `font` `null` where the runs
 * differ, `search`, the tracked changes and the comments over all of it. Its writes (phase 2) are Word's too:
 * `delete` and `insertText('Replace')` join the paragraphs in the last one's properties, `font` formats every
 * run and mark in it, `insertText` goes at its ends, and `insertParagraph` splits the paragraph at its start or
 * end; the rest of its writes throw `NotSupportedError`.
 *
 * A picture has a place in the text and no text of its own, so a picture at the range's start or
 * end offset may be inside it or outside it. `whole` says: a range from `Paragraph.getRange('Whole')`
 * or `'Content'` holds the paragraph's whole content, the pictures at its edges too, and deleting or
 * replacing it deletes them; any other range holds the pictures strictly inside it, so that a search
 * hit replaced does not take the picture beside it (CR-002 section 48).
 */
export class Range {
  private startParagraph: Paragraph;
  private lastParagraph: Paragraph;
  private holdsWhole: boolean;
  private holdsMark: boolean;
  private lastMarkText: string | undefined;

  constructor(paragraph: Paragraph, public start: number, public end: number, whole = false,
    /**
     * A paragraph's 'Whole' range: its text ends with the paragraph's mark, as Office JS's does (check 48:
     * `\r`, and `\t` for a cell's last paragraph, the cell's end), where 'Content' ends with its text.
     * The mark is no offset: every edit and search works on the text before it.
     */
    mark = false,
    /** The paragraph holding `end`, where it isn't `paragraph` (CR-014 phase 1). */
    endParagraph?: Paragraph,
    /**
     * What the range's text ends with in place of `endParagraph`'s mark, where `mark` is set: a table's
     * 'Whole' range ends with its last row's end, `\r\n` (check 52).
     */
    endMark?: string) {
    this.startParagraph = paragraph;
    this.lastParagraph = endParagraph ?? paragraph;
    this.holdsWhole = whole;
    this.holdsMark = mark;
    this.lastMarkText = endMark;
  }

  /** The paragraph holding `start`. */
  get paragraph(): Paragraph {
    return this.startParagraph;
  }

  /** The paragraph holding `end`: `paragraph` for a range within one paragraph (CR-014 phase 1). */
  get endParagraph(): Paragraph {
    return this.lastParagraph;
  }

  /** The range holds its paragraphs' whole content, the pictures at its edges too (CR-002 section 48). */
  get whole(): boolean {
    return this.holdsWhole;
  }

  /** The range's text ends with its last paragraph's mark: a 'Whole' range. */
  get mark(): boolean {
    return this.holdsMark;
  }

  /** What the text ends with in place of the last paragraph's mark: a table's 'Whole' range's `\r\n`. */
  get endMark(): string | undefined {
    return this.lastMarkText;
  }

  /**
   * True where the range holds its last paragraph's mark (a 'Whole' range ending at the paragraph's end), or a
   * table's end in its place: what `delete` takes and `Replace` keeps (check 52, W1 and W2). An extension.
   */
  get holdsEndMark(): boolean {
    return this.holdsMark && this.end >= this.lastParagraph.text.length;
  }

  /** The range collapsed where an edit through it closed it up, or over what the edit put there. */
  /**
   * This range moved onto another's span: its paragraphs and offsets (a replacement's, CR-002 section 60); `fresh`, it
   * no longer holds a mark or a paragraph's whole content either, as `landAt` leaves one.
   */
  private cover(r: Range, fresh = false): void {
    this.startParagraph = r.paragraph;
    this.lastParagraph = r.endParagraph;
    this.start = r.start;
    this.end = r.end;
    if (fresh) {
      this.holdsWhole = false;
      this.holdsMark = false;
      this.lastMarkText = undefined;
    }
  }

  private landAt(paragraph: Paragraph, start: number, end = start): void {
    this.startParagraph = paragraph;
    this.lastParagraph = paragraph;
    this.start = start;
    this.end = end;
    this.holdsWhole = false;
    this.holdsMark = false;
    this.lastMarkText = undefined;
  }

  /** True for a range across paragraphs: its ends are in two (extension; CR-014). */
  get spansParagraphs(): boolean {
    return this.endParagraph.element !== this.paragraph.element;
  }

  /** The edges `Paragraph.splice` is given: both for a whole range. */
  private get edges(): SpanEdges | undefined {
    return this.whole ? WHOLE : undefined;
  }

  /** The story's layout, and where this range's two paragraphs are in it. */
  private located(): { layout: StoryLayout; first: number; last: number } {
    const layout = storyLayout(storyOf(this.paragraph.p).items, (element, container) => this.viewOf(element, container).text);
    const first = layout.indexOf(this.paragraph.p);
    const last = layout.indexOf(this.endParagraph.p);
    if (first < 0 || last < 0) throw new Docx4JException('This range\'s paragraphs are not in one story as a reader reads it');
    return { layout, first, last };
  }

  /** The view of a paragraph of this range's story: this range's own where it is one of its two. */
  private viewOf(element: Element<wml.P>, container: Element[]): Paragraph {
    if (element === this.paragraph.element) return this.paragraph;
    if (element === this.endParagraph.element) return this.endParagraph;
    return this.paragraph.viewOf(element, container);
  }

  /**
   * The range as one range per paragraph it touches: its part of the first, the whole of each between, its
   * part of the last. A range within one paragraph is itself.
   */
  private pieces(): Range[] {
    if (!this.spansParagraphs) return [this];
    const { layout, first, last } = this.located();
    return layout.paragraphs.slice(first, last + 1).map((sp, i, all) => {
      const p = this.viewOf(sp.element, sp.container);
      if (i === 0) return new Range(p, this.start, p.text.length, this.whole);
      if (i === all.length - 1) return new Range(p, 0, this.end, this.whole);
      return new Range(p, 0, p.text.length, true);
    });
  }

  /**
   * Refuses a hyperlink set on a range across paragraphs. Word wrote a HYPERLINK field from the first paragraph into
   * the last (check 54 case 34), where this package's links are `w:hyperlink` elements, one paragraph's each, which
   * its getter reads (CR-014 section 12).
   */
  private assertOneParagraph(member: string): void {
    if (this.spansParagraphs) {
      throw new NotSupportedError('Range', member, "this range spans paragraphs, where Word writes a HYPERLINK field across them, which isn't "
        + "written here (CR-014); link each paragraph's range");
    }
  }

  /** Refuses, before anything changes, a write over a paragraph of the range that is read-only (a text box's, a branch's). */
  private assertWritableSpan(): void {
    for (const p of this.paragraphs) p.assertWritable();
  }

  /** Every paragraph the range touches, in document order, a table's cells' among them (check 52). */
  get paragraphs(): Paragraph[] {
    if (!this.spansParagraphs) return [this.paragraph];
    const { layout, first, last } = this.located();
    return layout.paragraphs.slice(first, last + 1).map((sp) => this.viewOf(sp.element, sp.container));
  }

  /**
   * The range's text, in `body.text`'s form across paragraphs (check 52): a mark `\r` between paragraphs, a
   * table's cells a `\t` apart and each row ended `\r\n`; with the last paragraph's mark where the range
   * holds it (a 'Whole' range).
   */
  get text(): string {
    if (!this.spansParagraphs) {
      const text = this.paragraph.text;
      const span = text.substring(this.start, this.end);
      return this.mark && this.end >= text.length ? span + (this.endMark ?? this.paragraph.markText) : span;
    }
    const tail = this.mark && this.end >= this.endParagraph.text.length ? this.endMark ?? this.endParagraph.markText : '';
    const { layout, first, last } = this.located();
    return layout.text.slice(layout.paragraphs[first]!.start + this.start, layout.paragraphs[last]!.start + this.end) + tail;
  }

  set text(value: string) {
    this.insertText(value, 'Replace');
  }

  /** True when the range holds no text (Office JS `isEmpty`). */
  get isEmpty(): boolean {
    return this.text === '';
  }

  /**
   * The paragraph's style id (extension, as on Paragraph); a range across paragraphs gives the first one's. Set
   * across paragraphs, each paragraph's (check 54 cases 32 and 33: Word gave both Heading 2).
   */
  get styleId(): string {
    return this.paragraph.styleId;
  }
  set styleId(id: string) {
    for (const p of this.styledParagraphs()) p.styleId = id;
  }

  /**
   * The range's style name, as Office JS reads it (checks 52 and 53). A range holding its paragraphs whole reads
   * their paragraph style, where they share one, else `Normal`. Any other range reads its runs' character style
   * where they share one (`Emphasis`), `''` where their character styles differ (some runs styled, some not: a
   * footnote reference among plain text), and the paragraph style where none has one. Setting it sets the
   * paragraph's style.
   */
  get style(): string {
    const read = this.styleRead();
    return read.charStyle === undefined ? read.paragraphs((p) => p.style) : read.charStyle === '' ? '' : styleNameOf(this.paragraph.parentBody.package_, read.charStyle);
  }
  set style(id: string) {
    for (const p of this.styledParagraphs()) p.style = id;
  }

  /** The built-in name of `style` (Office JS `styleBuiltIn`): `Other` where it is `''`, `Normal` where paragraphs across it differ. */
  get styleBuiltIn(): string {
    const read = this.styleRead();
    return read.charStyle === undefined ? read.paragraphs((p) => p.styleBuiltIn) : read.charStyle === '' ? 'Other' : builtInOf(read.charStyle);
  }
  /** Sets each paragraph's style by its built-in name, as `style` does. */
  set styleBuiltIn(value: string) {
    for (const p of this.styledParagraphs()) p.styleBuiltIn = value;
  }

  /**
   * The paragraphs a style set on the range goes to: its own, or each across paragraphs (check 54 cases 32 and 33:
   * Word gave a body paragraph and a list item Heading 2). Word also took away the body paragraph's centring and
   * italic and the list item's numbering, which `Paragraph.style` doesn't (CR-014 section 12).
   */
  private styledParagraphs(): Paragraph[] {
    if (!this.spansParagraphs) return [this.paragraph];
    this.assertWritableSpan();
    return this.paragraphs;
  }

  /**
   * What `style` reads: the runs' shared character style id, `''` where their character styles differ, or
   * undefined where the paragraph style applies, with how to read that.
   */
  private styleRead(): { charStyle: string | undefined; paragraphs: (read: (p: Paragraph) => string) => string } {
    const paragraphs = (read: (p: Paragraph) => string): string => {
      if (!this.spansParagraphs) return read(this.paragraph);
      const values = new Set(this.paragraphs.map(read));
      return values.size === 1 ? [...values][0]! : 'Normal';
    };
    const holdsWhole = this.start === 0 && this.end >= this.endParagraph.text.length;
    if (holdsWhole) return { charStyle: undefined, paragraphs };
    const ids = new Set(this.runs.map((run) => run.value.rPr?.rStyle?.val ?? ''));
    if (ids.size === 0) return { charStyle: undefined, paragraphs };
    if (ids.size > 1) return { charStyle: '', paragraphs };
    const only = [...ids][0]!;
    return { charStyle: only === '' ? undefined : only, paragraphs };
  }

  /** The table cell holding the range's start, if it is in one (Office JS `parentTableCell`, which throws instead). */
  get parentTableCell(): TableCell | undefined {
    return this.paragraph.parentTableCell;
  }

  /** The table holding the range's start, if it is in one (Office JS `parentTable`, which throws instead). */
  get parentTable(): Table | undefined {
    return this.paragraph.parentTableCell?.parentTable;
  }

  /**
   * The tables in the range (Office JS `Range.tables`): a table the range covers some of, unless the range is
   * inside one of its cells, where that cell's tables are looked at instead (check 52: a range across a table,
   * or into or out of one, lists it; a cell's whole range lists none).
   */
  get tables(): Table[] {
    const inRange = new Set(this.paragraphs.map((p) => p.element.value));
    const pick = (tables: Table[]): Table[] => tables.flatMap((table) => {
      const cells = table.rows.flatMap((row) => row.cells);
      if (!cells.some((cell) => cell.paragraphs.some((p) => inRange.has(p.p)))) return [];
      const home = cells.find((cell) => [...inRange].every((p) => cell.paragraphs.some((cp) => cp.p === p)));
      return home ? pick(home.body.tables) : [table];
    });
    return pick(this.storyBody().tables);
  }

  /**
   * The content controls in the range (Office JS `Range.contentControls`): each control the range overlaps,
   * except one holding the whole range, in document order (check 53: a range from a paragraph into a control
   * lists it, one inside a control, its content whole, doesn't, one over an inline control lists it).
   */
  get contentControls(): ContentControl[] {
    const layout = storyLayout(storyOf(this.paragraph.p).items, (element, container) => this.viewOf(element, container).text);
    const at = (p: Paragraph, offset: number): number => {
      const i = layout.indexOf(p.p);
      return i < 0 ? Number.NaN : layout.paragraphs[i]!.start + offset;
    };
    const from = at(this.paragraph, this.start);
    const to = at(this.endParagraph, this.end);
    return this.storyBody().contentControls.filter((control) => {
      let start: number;
      let end: number;
      if (control.form === 'Run') {
        let range: Range;
        try { range = control.getRange(); } catch { return false; }
        [start, end] = [at(range.paragraph, range.start), at(range.paragraph, range.end)];
      } else {
        const paragraphs = control.paragraphs;
        if (paragraphs.length === 0) return false;
        const last = paragraphs[paragraphs.length - 1]!;
        [start, end] = [at(paragraphs[0]!, 0), at(last, last.text.length)];
      }
      if (Number.isNaN(start) || Number.isNaN(end)) return false;
      const overlaps = start < to && end > from || (start === end && from <= start && start <= to);
      const holdsRange = start <= from && to <= end;
      return overlaps && !holdsRange;
    });
  }

  /**
   * Office JS `getContentControls` (WordApi 1.5): the content controls `contentControls` lists, those of `options.types`
   * where given. Word's `contentControls` lists rich-text controls only; this package's lists every kind, a recorded
   * departure (CR-002 section 60 G1, check 55b), so this is how add-in code that asks for plain-text controls finds them.
   */
  getContentControls(options?: ContentControlOptions): ContentControl[] {
    return controlsOfTypes(this.contentControls, options);
  }

  /** A body over this range's story, for its tables and controls. */
  private storyBody(): Body {
    const { owner } = storyOf(this.paragraph.p);
    return this.paragraph.parentBody.sub(owner as { content?: Element[] });
  }

  /**
   * The text in one of the two views (extension; `text` is the accepted one). `{ view: 'original' }`
   * is offered on `Paragraph` and `Body`; on a `Range` it is the accepted-view span only, since a
   * range's offsets are accepted-view offsets.
   */
  getText(options?: TextViewOptions): string {
    if (options?.view === 'original') throw new Docx4JException("A Range's offsets are accepted-view offsets; read the original view on its Paragraph");
    return this.text;
  }

  /**
   * The formatting of exactly this span: reads the first run's *effective* properties, writes
   * direct formatting, splitting the runs at the boundaries so that a write touches only the
   * span.
   */
  get font(): Font {
    return this.getFont();
  }

  /** `font`, with `{ direct: true }` for the direct formatting of the first run (extension). */
  getFont(options?: FormattingOptions): Font {
    // across paragraphs: each run read in its own paragraph, null where they differ, as Office JS's (check 52)
    if (this.spansParagraphs) {
      return new MixedFont(() => {
        const resolver = this.paragraph.parentBody.propertyResolver;
        const selector = runFontSelectorOf(this.paragraph.parentBody.package_);
        return this.pieces().flatMap((piece) => piece.runs.map((run) => {
        const rPr = options?.direct === true ? run.value.rPr
          : resolver.getEffectiveRPr(run.value.rPr, piece.paragraph.p.pPr, piece.paragraph.cellContext);
          return readingOf(rPr, options?.direct !== true, options?.direct === true ? undefined : selector?.asciiFontName(rPr));
        }));
      }, (opts) => {
        // every run in it and every mark it holds (CR-014 phase 2, check 52)
        this.assertWritableSpan();
        const pieces = this.pieces();
        formatSpan(pieces.map((piece, i) => ({
          paragraph: piece.paragraph, start: piece.start, end: piece.end, runs: () => piece.runs,
          holdsMark: i < pieces.length - 1 || this.holdsEndMark,
        })), opts);
      });
    }
    const holders = (write: boolean): wml.R[] => {
      if (this.start === this.end) return [];
      // a write formats runs holding the span and nothing else: a tab, a break or text after the span in
      // the same run is split off too, where splitting at the two offsets left it in the run and formatted
      // it; a read splits nothing, the first run's properties being the same either way (CR-012 phase 4)
      if (write) this.paragraph.splitRunsAround(this.start, this.end);
      return this.runs.map((r) => r.value);
    };
    // a range holding its paragraph's mark and no run, an empty paragraph's whole range, reads the mark (CR-002 section 59)
    const mark = (): RPrHolder | undefined => (this.holdsEndMark ? markHolderOf(this.paragraph.p) : undefined);
    if (options?.direct === true) return new Font(holders, () => this.paragraph.fontTracking(), undefined, undefined, mark);
    return new Font(holders, () => this.paragraph.fontTracking(),
      (rPr) => this.paragraph.parentBody.propertyResolver.getEffectiveRPr(rPr, this.paragraph.p.pPr, this.paragraph.cellContext),
      (rPr) => runFontSelectorOf(this.paragraph.parentBody.package_)?.asciiFontName(rPr), mark);
  }

  /** The runs the span covers (a run partly inside counts). */
  get runs(): Element<wml.R>[] {
    if (this.spansParagraphs) return this.pieces().flatMap((piece) => piece.runs);
    const out: Element<wml.R>[] = [];
    const seen = new Set<wml.R>();
    for (const seg of this.paragraph.segments()) {
      // an equation is in no run, and its math runs' formatting is their own (CR-012 phase 4)
      if (seg.end <= this.start || seg.start >= this.end || seg.math) continue;
      if (seen.has(seg.run)) continue;
      seen.add(seg.run);
      out.push(seg.runOwner[seg.runIndex] as Element<wml.R>);
    }
    return out;
  }

  /**
   * Text put at the range's start ('Start', 'Before'), at its end ('End', 'After'), or in its place ('Replace').
   * Across paragraphs (CR-014 phase 2, check 52) the start is in the first paragraph and the end in the last,
   * before a last mark the range holds, and a replacement is `delete()` keeping a last mark, then the text where the
   * paragraphs joined: untracked, in the first replaced character's look (`lookOf`); tracked, a `w:ins` before the
   * deletion, as Word writes it. A range holding its paragraph's mark is replaced so too under tracking.
   */
  insertText(text: string, location: 'Before' | 'After' | 'Start' | 'End' | 'Replace'): Range {
    // tracked, the w:ins goes before the w:del within one paragraph too, as Word's (check 54 case 4; Jason,
    // 2026-10-23: "follow Word (insertion first)"), where CR-002 wrote the w:del first
    if (location === 'Replace' && (this.spansParagraphs || this.paragraph.changeTracker)) return this.replaceAcross(text);
    if (this.spansParagraphs) {
      this.assertWritableSpan();
      if (location === 'End' || location === 'After') return this.endParagraph.splice(this.end, this.end, text);
      const r = this.paragraph.splice(this.start, this.start, text, 'after');
      this.startParagraph = r.endParagraph;   // the text may have broken the paragraph (CR-002 section 60 E1)
      this.start = r.end;
      return r;
    }
    let r: Range;
    if (location === 'Replace') r = this.paragraph.splice(this.start, this.end, text, 'before', this.edges);
    // text put before the range joins the run the range starts in - inside the link the range is on,
    // as Office JS (check 32: insertText('Before') on a link's text went into the link, as 'After' did)
    else if (location === 'Before' || location === 'Start') r = this.paragraph.splice(this.start, this.start, text, 'after');
    else r = this.paragraph.splice(this.end, this.end, text);
    // this range keeps covering the same text: a replacement resizes it, text before it shifts it, text after it leaves
    // it; what was put in is the returned range's, which a line break may have taken into the paragraphs it made
    if (location === 'Replace') this.cover(r);
    else if (location === 'Before' || location === 'Start') {
      const length = this.end - this.start;
      this.startParagraph = this.lastParagraph = r.endParagraph;
      this.start = r.end;
      this.end = r.end + length;
    }
    return r;
  }

  /**
   * A new paragraph of the text where the range starts ('Before') or ends ('After'), splitting the paragraph there,
   * as Word does (`insertParagraphAt`): across paragraphs (check 52) and within one (check 54 cases 1 to 3; until
   * then a range within one paragraph put it before or after the whole paragraph). At a paragraph's start or end,
   * or after a range holding its mark, it goes before or after the paragraph. This range keeps covering its text.
   */
  insertParagraph(text: string, location: 'Before' | 'After'): Paragraph {
    this.assertWritableSpan();
    const spans = this.spansParagraphs;
    if (location === 'Before') {
      const { inserted, firstHalf } = insertParagraphAt(this.paragraph, this.start, text);
      if (firstHalf) {                                     // the paragraph keeps the text from the split on
        if (!spans) this.end -= this.start;
        this.start = 0;
      }
      return inserted;
    }
    const at = this.holdsEndMark ? this.endParagraph.text.length : this.end;
    const { inserted, firstHalf } = insertParagraphAt(this.endParagraph, at, text);
    if (firstHalf) {                                       // the text before the split moved into the first half
      if (!spans) this.startParagraph = firstHalf;
      this.lastParagraph = firstHalf;
    }
    return inserted;
  }

  /**
   * Word's `insertOoxml` at range level: a flat OPC `pkg:package` (or a bare fragment). One
   * incoming paragraph's runs go at the start or the end of the span, or in its place; anything
   * else is inserted as blocks before or after the paragraph. Returns what was inserted, as
   * `insertXml` does. Across paragraphs (check 54 case 35) one paragraph's runs replace the range as
   * `insertText('Replace')` replaces it, where the paragraphs joined; the start is the first paragraph's,
   * the end the last's; several blocks go before the first paragraph or after the last, and replacing a
   * range across paragraphs with several blocks is refused, as unmeasured.
   */
  async insertOoxml(ooxml: string, location: 'Before' | 'After' | 'Replace'): Promise<(Paragraph | BlockElement)[]> {
    this.assertWritableSpan();
    const body = this.paragraph.parentBody;
    const preprocess = body.package_?.loadOptions.preprocessor;
    const elements = await contentOf(ooxml, { preprocess: preprocess ? (doc) => preprocess(doc) : undefined, target: body.part });
    if (elements.length === 0) return [];
    const only = elements.length === 1 && typeNameOf(elements[0]!) === 'org_docx4j_wml.P' ? elements[0]!.value as { content?: Element[] } : undefined;
    if (only) {
      if (location === 'Replace') {
        const landing = this.removeKeepingMark('insertOoxml');
        landing.paragraph.insertItemsAt(landing.offset, (only.content ?? []));
        return [landing.paragraph];
      }
      const [paragraph, at] = location === 'Before' ? [this.paragraph, this.start] : [this.endParagraph, this.end];
      paragraph.insertItemsAt(at, (only.content ?? []));
      return [paragraph];
    }
    if (location === 'Replace' && this.spansParagraphs) {
      throw new NotSupportedError('Range', 'insertOoxml', "several blocks in place of a range across paragraphs, which Word wasn't measured doing (CR-014)");
    }
    const target = location === 'Before' ? this.paragraph : this.endParagraph;
    body.insertElement(elements, location === 'Before' ? 'Before' : 'After', target);
    if (location === 'Replace') this.removeKeepingMark('insertOoxml');
    return elements.map((el) => typeNameOf(el) === 'org_docx4j_wml.P'
      ? body.paragraphFor(el.value as never) ?? { element: el, container: target.container }
      : { element: el, container: target.container });
  }

  /**
   * The range's content removed as a replacement removes it, keeping a last mark the range holds (check 52):
   * `delete()`'s join across paragraphs, else the paragraph's splice. Where the paragraphs closed up.
   */
  private removeKeepingMark(member: string): { paragraph: Paragraph; offset: number } {
    if (!this.spansParagraphs && !this.holdsEndMark) {
      this.paragraph.splice(this.start, this.end, '', 'before', this.edges);
      this.end = this.start;
      return { paragraph: this.paragraph, offset: this.start };
    }
    const extent = this.replaceableExtent(member);
    if (this.paragraph.changeTracker) assertTrackable(extent, member);
    const landing = deleteSpan(this, extent, true)!;
    this.landAt(landing.paragraph, landing.offset);
    return landing;
  }

  /**
   * The extent of a range a replacement may take: refused where it ends with a whole table, as Word refused it
   * (`GeneralException`, check 54 case 7), and where it starts with a whole table or control (unmeasured).
   */
  private replaceableExtent(member: string): Extent {
    const extent = extentOf(this, member);
    if (!extent.tail) {
      throw new GeneralException(`Range.${member}: this range ends with a whole table or control, which Office JS refuses to replace (check 54)`);
    }
    if (!extent.head) {
      throw new NotSupportedError('Range', member, "this range starts with a whole table or content control, so there is no paragraph for the text to go in, and what Word does isn't measured");
    }
    this.assertWritableSpan();
    return extent;
  }

  /**
   * Removes the span's text, and the pictures in it (CR-002 section 48). Across paragraphs, or holding its
   * paragraph's mark, it is Word's (CR-014 phase 2, check 52; `deleteSpan`): the paragraphs join in the last one's
   * properties, the blocks between go, a last mark the range holds goes with it, joining the paragraph after; a
   * footnote whose reference goes goes on save. Tracked, the text is deleted and every mark in the range marked
   * deleted. A range starting or ending inside a table is refused with `GeneralException`, as Office JS refuses it.
   */
  delete(): void {
    if (this.spansParagraphs || this.holdsEndMark) {
      const extent = extentOf(this, 'delete', true);
      this.assertWritableSpan();
      if (this.paragraph.changeTracker) assertTrackable(extent, 'delete');
      const landing = deleteSpan(this, extent, false);
      if (landing) this.landAt(landing.paragraph, landing.offset);
      return;
    }
    this.paragraph.splice(this.start, this.end, '', 'before', this.edges);
    this.end = this.start;
  }

  /** `insertText(text, 'Replace')` across paragraphs, and any under tracking (CR-014 phases 2 and 3). */
  private replaceAcross(text: string): Range {
    const extent = this.replaceableExtent('insertText');
    if (this.paragraph.changeTracker) {
      assertTrackable(extent, 'insertText');
      // tracked: the insertion before the deletion (check 52 cases 39 to 48), in the first replaced run's formatting
      const length = this.spansParagraphs ? undefined : this.end - this.start;
      const inserted = this.paragraph.splice(this.start, this.start, text, 'after');
      // the range's own text follows what was put in, in the paragraph the insertion ended in (E1)
      this.startParagraph = inserted.endParagraph;
      this.start = inserted.end;
      if (length !== undefined) { this.lastParagraph = inserted.endParagraph; this.end = inserted.end + length; }
      deleteSpan(this, extentOf(this, 'insertText'), true);
      this.cover(inserted, true);
      return new Range(inserted.paragraph, inserted.start, inserted.end, false, false, inserted.endParagraph);
    }
    const first = firstCharacterOf(this);
    // a line break in the text is a paragraph break (E1), refused before anything changes where none can be made
    if (/[\r\n]/.test(text)) this.paragraph.assertBreakableAt(this.start, 'before');
    const landing = deleteSpan(this, extent, true)!;
    const placed = placeReplacement(landing, first, text);
    const r = new Range(placed.first, placed.start, placed.end, false, false, placed.last);
    this.cover(r, true);
    return r;
  }

  /**
   * The hits of a query in the range, as Office JS's `search` gives them. A plain query holding `^p`, Word's
   * paragraph mark, matches across paragraphs (`searchMarks`); any other matches a paragraph at a time.
   */
  search(text: string, options?: SearchOptions): Range[] {
    const marks = options?.matchWildcards || options?.matchRegExp ? undefined : paragraphMarkPieces(text);
    if (marks) return this.searchMarks(marks, options);
    if (this.spansParagraphs) return this.pieces().flatMap((piece) => piece.search(text, options));
    return officeMatches(this.paragraph.segments({ objects: true }), text, options, this.start, this.end)
      .map(({ start, end }) => new Range(this.paragraph, start, end));
  }

  /**
   * A query's hits where it holds `^p`, Word's paragraph mark, as Word finds them (check 52 case 2; CR-014 phase
   * 3): `pieces` is the query split at its marks. A hit starts in one paragraph with the first piece matching to
   * its end, takes each middle piece as a whole paragraph's text, and ends with the last piece matching from the
   * next paragraph's start; a query ending with `^p` ends with that paragraph's mark, which the hit holds. A mark
   * matches only a paragraph's own mark the range holds: never a cell's or a row's end, and never across the edge
   * of a table, so every paragraph of a hit is in the same cell or in none (`^p` between a list item and the first
   * cell after it found nothing). `^p^p` matches an empty paragraph's mark after another mark. Plain queries only:
   * Word's wildcards write the mark another way.
   */
  private searchMarks(pieces: string[], options?: SearchOptions): Range[] {
    const paragraphs = this.paragraphs;
    const last = paragraphs.length - 1;
    const marks = pieces.length - 1;
    const endsWithMark = pieces[marks] === '';
    const cell = (p: Paragraph): object | undefined => p.parentTableCell?.element.value as object | undefined;
    // a paragraph's mark the search may match: the range holds it, and it is a paragraph's, not a cell's end
    const markIn = (j: number): boolean => (j < last || this.holdsEndMark) && paragraphs[j]!.markText === '\r';
    const matches = (p: Paragraph, query: string, from: number, to: number) => officeMatches(p.segments({ objects: true }), query, options, from, to);
    const out: Range[] = [];
    for (let i = 0; i <= last; i++) {
      const end = i + marks - (endsWithMark ? 1 : 0);                  // the paragraph the hit ends in
      if (end > last) break;
      const scope = cell(paragraphs[i]!);
      let ok = true;
      for (let j = i; j < i + marks && ok; j++) ok = markIn(j);
      for (let j = i + 1; j <= end && ok; j++) ok = cell(paragraphs[j]!) === scope;
      // the middle pieces: each a whole paragraph's text (an empty one for an empty piece)
      for (let j = 1; j < marks && ok; j++) {
        const p = paragraphs[i + j]!;
        ok = pieces[j] === '' ? p.text === '' : matches(p, pieces[j]!, 0, p.text.length).some((m) => m.start === 0 && m.end === p.text.length);
      }
      if (!ok) continue;
      const first = paragraphs[i]!;
      const from = i === 0 ? this.start : 0;
      const starts = pieces[0] === '' ? [first.text.length]
        : matches(first, pieces[0]!, from, first.text.length).filter((m) => m.end === first.text.length).map((m) => m.start);
      if (starts.length === 0 || starts[0]! < from) continue;
      const endParagraph = paragraphs[end]!;
      let to: number;
      if (endsWithMark) to = endParagraph.text.length;
      else {
        const limit = end === last ? this.end : endParagraph.text.length;
        const hit = matches(endParagraph, pieces[marks]!, 0, limit).find((m) => m.start === 0);
        if (!hit) continue;
        to = hit.end;
      }
      out.push(new Range(first, starts[0]!, to, false, endsWithMark, endParagraph.element === first.element ? undefined : endParagraph));
    }
    return out;
  }

  /**
   * Replaces every match in this span, last first so the offsets stay valid; returns the count
   * (CR-002 section 3.7). The replacement's `$1` or Word's `\1` expands as it does on a paragraph
   * (section 26; this one was missed there, section 31).
   */
  replaceText(find: string, replace: string, options?: SearchOptions): number {
    refuseMarkReplacement('Range', find, options);
    // across paragraphs a paragraph at a time, as `search` (CR-014 section 3), the last first
    if (this.spansParagraphs) {
      this.assertWritableSpan();
      return this.pieces().reverse().reduce((count, piece) => count + piece.replaceText(find, replace, options), 0);
    }
    const matches = officeMatches(this.paragraph.segments({ objects: true }), find, options, this.start, this.end);
    for (let i = matches.length - 1; i >= 0; i--) {
      const { start, end, match } = matches[i]!;
      new Range(this.paragraph, start, end).insertText(expandReplacement(replace, match, options), 'Replace');
    }
    return matches.length;
  }

  /** The tracked changes this span covers, in document order (CR-002 phase F). */
  getTrackedChanges(): TrackedChange[] {
    if (this.spansParagraphs) {
      // the story's listing, whose changes run across paragraphs as Word's do (check 52: a deletion through a
      // mark is one change), each that overlaps the range; a paragraph's own listing cuts a change at its edges
      const { layout, first, last } = this.located();
      const at = (p: Paragraph, offset: number): number => {
        const i = layout.indexOf(p.p);
        return i < 0 ? Number.NaN : layout.paragraphs[i]!.start + offset;
      };
      const from = layout.paragraphs[first]!.start + this.start;
      const to = layout.paragraphs[last]!.start + this.end;
      return this.storyBody().getTrackedChanges().filter((change) => {
        const range = change.getRange();
        if (!range) return false;
        const start = at(range.paragraph, range.start);
        const end = at(range.endParagraph, range.end);
        return !Number.isNaN(start) && !Number.isNaN(end) && start <= to && end >= from;
      });
    }
    return this.paragraph.getTrackedChanges().filter((change) => {
      const range = change.getRange();
      if (!range) return false;
      // a change across paragraphs (CR-014) covers this paragraph from where it starts in it to where it ends in it
      const start = range.paragraph.element === this.paragraph.element ? range.start : 0;
      const end = range.endParagraph.element === this.paragraph.element ? range.end : Number.POSITIVE_INFINITY;
      return start <= this.end && end >= this.start;
    });
  }

  /**
   * A range of this one, as Office JS's `getRange` (check 52): 'Whole' is the range itself, 'Content' the same
   * without a last mark, 'Start' and 'End' collapsed at its ends, and 'After' collapsed just after it - at the
   * start of the next paragraph where the range holds its last paragraph's mark.
   */
  getRange(location: 'Whole' | 'Start' | 'End' | 'Content' | 'After' = 'Whole'): Range {
    if (location === 'Start') return new Range(this.paragraph, this.start, this.start);
    if (location === 'End') return new Range(this.endParagraph, this.end, this.end);
    if (location === 'After') {
      if (this.mark && this.end >= this.endParagraph.text.length) {
        const layout = storyLayout(storyOf(this.endParagraph.p).items, (element, container) => this.viewOf(element, container).text);
        const next = layout.paragraphs[layout.indexOf(this.endParagraph.p) + 1];
        if (next) return new Range(this.viewOf(next.element, next.container), 0, 0);
      }
      return new Range(this.endParagraph, this.end, this.end);
    }
    return new Range(this.paragraph, this.start, this.end, this.whole, this.mark && location === 'Whole',
      this.spansParagraphs ? this.endParagraph : undefined, this.endMark);
  }

  /**
   * A range from this range's start or `range`'s, whichever comes first, to the end that comes last (Office JS
   * `expandTo`; which is first doesn't matter, check 52). Both must be in one story.
   */
  expandTo(range: Range): Range {
    const story = storyOf(this.paragraph.p).items;
    // Office JS's ItemNotFound (check 53: a body range expandTo a footnote's)
    if (storyOf(range.paragraph.p).items !== story) throw new ItemNotFound('expandTo: the two ranges are in different stories');
    const layout = storyLayout(story, (element, container) => this.viewOf(element, container).text);
    const at = (p: Paragraph, offset: number): number => {
      const i = layout.indexOf(p.p);
      if (i < 0) throw new Docx4JException('expandTo: a range is outside the story as a reader reads it');
      return layout.paragraphs[i]!.start + offset;
    };
    const first = at(range.paragraph, range.start) < at(this.paragraph, this.start) ? range : this;
    const thisEnd = at(this.endParagraph, this.end) + (this.mark ? 1 : 0);
    const rangeEnd = at(range.endParagraph, range.end) + (range.mark ? 1 : 0);
    const last = rangeEnd > thisEnd ? range : this;
    return new Range(first.paragraph, first.start, last.end, first.whole && last.whole, last.mark,
      last.endParagraph.element === first.paragraph.element ? undefined : last.endParagraph, last.endMark);
  }

  /** The span's text as it stands (extension; Office JS getOoxml wraps it in a package). */
  toString(): string {
    return this.text;
  }

  // --- comments (CR-002 phase G) ---

  /** The comments this span touches, replies nested under their parent. */
  async getComments(): Promise<Comment[]> {
    if (!this.spansParagraphs) return commentApi().commentsOf(this);
    const seen = new Set<number>();
    const out: Comment[] = [];
    for (const piece of this.pieces()) {
      for (const comment of await commentApi().commentsOf(piece)) if (!seen.has(comment.id)) { seen.add(comment.id); out.push(comment); }
    }
    return out;
  }

  /**
   * Comments the span: the markers and the reference run around it (runs are split at the
   * boundaries, as `font` does), the comment itself and the side-part entries, creating any of
   * the comment parts the document lacks. Asynchronous because it unmarshals those parts.
   */
  async insertComment(content: CommentContent, options?: CommentOptions): Promise<Comment> {
    // across paragraphs (CR-014 phase 3) the markers go in the first and the last paragraph, as Word's (check 52)
    this.assertWritableSpan();
    return commentApi().insertComment(this, content, options);
  }

  // --- content controls (CR-002 phase E) ---

  /**
   * Wraps the runs this span covers in a run-level content control (Office JS
   * `Range.insertContentControl`), splitting the runs at the boundaries as `font` does. An empty
   * span gets an empty control at its position. Returns the control.
   */
  insertContentControl(kind?: ContentControlType): ContentControl {
    if (this.spansParagraphs) return this.wrapParagraphs(kind);
    this.paragraph.assertWritable();
    const paragraph = this.paragraph;
    const body = paragraph.parentBody;
    // The control is built empty first: `sdt` refuses a kind that cannot be run-level (a repeating
    // section), and that refusal must come before any run is split.
    const sdt = sdtOf([], { kind: sdtKindFor(kind), id: nextSdtId(controlIdScope(body)), form: 'run' }) as Element<wml.SdtRun>;
    declareSdtExtensions(body, sdt);
    if (this.start === this.end) {
      paragraph.insertItemsAt(this.start, [sdt as Element]);
      return new ContentControl(sdt, containerOf(sdt, paragraph), body);
    }
    // Every refusal before anything changes (CR-002 section 43): until then the runs were split first.
    const covered = coveredSegments(paragraph, this.start, this.end);
    if (covered.length === 0) throw new Docx4JException('This range covers no run');
    const coveredOwner = covered[0]!.runOwner;
    if (covered.some((s) => s.runOwner !== coveredOwner)) {
      throw new Docx4JException('This range spans more than one run holder (a hyperlink or a tracked change); wrap a narrower span');
    }
    refuseTrackedDeletion(coveredOwner, covered);
    const tracker = paragraph.changeTracker;
    const coveredRevision = revisionHolding(covered[0]!, coveredOwner);
    if (tracker && coveredRevision && itemsKeyOf(coveredRevision) === undefined) {
      throw new Docx4JException('This range is in a revision that cannot be split around a content control');
    }

    paragraph.splitAt(this.start);
    paragraph.splitAt(this.end);
    const segments = paragraph.segments().filter((s) => s.start >= this.start && s.end <= this.end);
    const owner = segments[0]!.runOwner;
    const { at, items } = wrappedItems(owner, segments);
    const revision = revisionHolding(segments[0]!, owner);
    if (tracker && revision) {
      // The runs are in a revision: the control goes between its halves, not in it, and the runs go
      // into the control in a revision of the same kind and attributes with a new id - Word's form for
      // inserted text wrapped in a new control (check 30, inserted-text-wrapped: the w:ins inside the
      // control, the control's own markers outside it).
      revision.items.splice(at, items.length);
      const wrapped = revisionAround(revision, items, tracker.nextId());
      (sdt.value.sdtContent as { content?: Element[] }).content = [wrapped];
      linkParents(wrapped, sdt.value.sdtContent);
      if (!placeInRevision(paragraph, revision, at, sdt as Element, () => tracker.nextId())) {
        throw new Docx4JException('This range is in a revision that cannot be split around a content control');
      }
      markHolderInserted(tracker, sdt as Element, revision.owner, () => {});
      return new ContentControl(sdt, revision.owner, body);
    }
    (sdt.value.sdtContent as { content?: Element[] }).content = items;
    owner.splice(at, items.length, sdt as Element);
    linkParents(sdt, (segments[0]!.run as { PARENT?: object }).PARENT ?? paragraph.p);
    // under tracking the control itself is the insertion, the runs in it what they were (check 30)
    if (tracker) markHolderInserted(tracker, sdt as Element, owner, () => {});
    return new ContentControl(sdt, owner, body);
  }

  /**
   * A block content control around the paragraphs of a range across paragraphs, as Word puts one (check 54 cases
   * 38 and 43, untracked and tracked): only a range holding its paragraphs whole, from the first's start to the
   * last's end, its mark or not. Any other, starting or ending inside a paragraph or a table, is refused with
   * `GeneralException`, as Word refused it (cases 36, 37 and 39 to 42). Tracked, the control is an insertion, as a
   * paragraph's control is (check 30's markers).
   */
  private wrapParagraphs(kind?: ContentControlType): ContentControl {
    if (this.start !== 0 || this.end < this.endParagraph.text.length) {
      throw new GeneralException('Range.insertContentControl: a range across paragraphs takes a control only where it holds its paragraphs whole, as Office JS (check 54)');
    }
    const extent = extentOf(this, 'insertContentControl');
    this.assertWritableSpan();
    const body = this.paragraph.parentBody;
    const { list, first, last } = extent;
    const parent = (list[first]!.value as { PARENT?: object }).PARENT ?? body.container;
    // a bound block control holds one paragraph (CR-002 section 61): checked before sdtOf, which takes the paragraphs
    assertControlHolds(body, parent, list.filter((_, i) => i < first || i > last), [BLOCK_CONTROL], undefined, ['Range', 'insertContentControl']);
    const sdt = sdtOf(list.slice(first, last + 1), { kind: sdtKindFor(kind), id: nextSdtId(controlIdScope(body)), form: 'block' }) as Element<wml.SdtBlock>;
    list.splice(first, last - first + 1, sdt as Element);
    linkParents(sdt, parent);
    declareSdtExtensions(body, sdt);
    const tracker = this.paragraph.changeTracker;
    if (tracker) markHolderInserted(tracker, sdt as Element, list, () => {});
    return new ContentControl(sdt, list, body);
  }

  /**
   * Office JS `Range.hyperlink`: the address of the first hyperlink in this range, or a new
   * hyperlink over it. `address#location` separates the address from an optional location within
   * it, so `"#heading"` is a link inside this document (a `w:anchor`) and
   * `"https://example.com"` an external one (a relationship). Reading gives `""` where the range
   * is in no hyperlink.
   *
   * Setting removes the hyperlinks the range already carries and wraps its runs in one
   * `w:hyperlink`, splitting runs at the boundaries as `font` and `insertContentControl` do;
   * setting `""` removes them and wraps nothing. An external address becomes a relationship of
   * this range's part, which is why the part must exist (CR-002 section 20). Only the hyperlinks
   * the range touches are removed; before 0.2.1 every one in the paragraph was (section 25). An
   * empty range is refused with a `GeneralException`, as Office JS refuses it (section 25, check 19):
   * to change a link under a caret, set it on the link's whole range.
   *
   * The runs are styled as Word styles them, measured through Office JS (`test/README.md` check
   * 15, CR-002 section 25): each run a link wraps takes the `Hyperlink` character style, over any
   * character style it had, and loses its direct colour, keeping the rest of its direct formatting;
   * a link removed takes the style off its runs again. The `Hyperlink` definition is added to the
   * styles part when the document lacks it, at the latest when the package is saved - a setter
   * cannot await it - so `await pkg.styles.ensure('Hyperlink')` first where an effective read
   * before the save must see it. Under change tracking the restyle is a formatting revision, as a
   * `font` write is.
   */
  get hyperlink(): string {
    if (this.spansParagraphs) return this.pieces().map((piece) => piece.hyperlink).find((link) => link !== '') ?? '';
    const holder = this.hyperlinkHolder();
    if (holder === undefined) return '';
    const anchor = holder.value.anchor ?? '';
    const address = holder.value.id === undefined ? '' : this.relationshipTarget(holder.value.id);
    return anchor === '' ? address : `${address}#${anchor}`;
  }

  set hyperlink(value: string) {
    this.assertOneParagraph('hyperlink');
    this.paragraph.assertWritable();
    // Office JS refuses a hyperlink set on an empty range - inside a link, in plain text, just after a
    // link, and "" alike - with GeneralException, and changes nothing (test/README.md check 19, Word
    // 16.0.20326.20158; CR-002 section 25). Before, the setter removed the link under a caret and
    // left an empty w:hyperlink there.
    if (this.start === this.end) {
      throw new GeneralException('A hyperlink cannot be set on an empty range, as in Office JS: select the text to link, or the whole of a link to change it');
    }
    const paragraph = this.paragraph;
    // A range that crosses the edge of a link takes the whole of that link into the new one; a range
    // inside a link replaces it with a link over the range alone. Measured in Word through Office JS
    // (test/README.md check 16, cases F and G; CR-002 section 25): "alpha beta" linked and then
    // "beta gamma" is one link over "alpha beta gamma", where "alpha beta gamma" linked and then
    // "beta" leaves "alpha " and " gamma" unlinked.
    let start = this.start;
    let end = this.end;
    for (const holder of hyperlinkHolders(paragraph)) {
      const span = this.spanOf(holder);
      if (span === undefined || !this.touchesSpan(span)) continue;
      if (span[0] <= this.start && this.end <= span[1]) continue;
      start = Math.min(start, span[0]);
      end = Math.max(end, span[1]);
    }
    const hash = value.indexOf('#');
    const address = hash === -1 ? value : value.slice(0, hash);
    const location = hash === -1 ? '' : value.slice(hash + 1);
    // Every refusal comes before anything changes, judged on what the range would hold once the links
    // it touches were unwrapped (CR-002 section 43). Until then the setter unwrapped them, added the
    // relationship and split the runs before it refused, so a refused call left the links gone, an
    // orphan relationship and split runs (found by the docx4j-python session).
    if (value !== '') this.refuseUnlinkable(start, end, address);
    this.removeHyperlinks();
    if (value === '') return;

    // w:history="1" is what Word writes on every hyperlink it makes, and what docx4j writes when
    // it builds one (MainDocumentPart.addHyperlink): "add the target to the viewed hyperlinks"
    // (ECMA-376 17.16.22). Readers do not depend on it, but a document holding one holder from
    // each source should not differ over it. Found by the docx4j-ts-editor session, ED-003
    // section 11.8.
    const holder = hyperlinkOf({ content: [], history: true }) as Element<wml.P.Hyperlink>;
    if (location !== '') holder.value.anchor = location;
    if (address !== '') holder.value.id = this.addHyperlinkRelationship(address);

    paragraph.splitAt(start);
    paragraph.splitAt(end);
    const segments = paragraph.segments().filter((seg) => seg.start >= start && seg.end <= end);
    const owner = segments[0]!.runOwner;
    const { at, items } = wrappedItems(owner, segments);
    const revision = revisionHolding(segments[0]!, owner);
    let runsParent: object = holder.value;
    if (revision) {
      // A w:hyperlink cannot be in a w:ins or w:moveTo (CT_RunTrackChange holds EG_ContentRunContent
      // only), so over runs in one, tracked or not, the link goes between the revision's halves and the
      // runs into it in a revision of the same kind and attributes with a new id, as a content control
      // over inserted text does (CR-002 section 43; the docx4j-python session's form). Until then the
      // link went into the revision, which marshalled with an xsi:type. Word's own form for a link set
      // on inserted text is a HYPERLINK field in the one w:ins (check 30, link-on-inserted-text), which
      // this setter's getter would not read back.
      revision.items.splice(at, items.length);
      const inner = revisionAround(revision, items, paragraph.nextAnnotationId());
      holder.value.content = [inner] as never;
      linkParents(inner, holder.value);
      if (!placeInRevision(paragraph, revision, at, holder as Element, () => paragraph.nextAnnotationId())) {
        throw new Docx4JException('This range is in a revision that cannot be split around a hyperlink');
      }
      runsParent = inner.value as object;
    } else {
      holder.value.content = items as never;
      owner.splice(at, items.length, holder as Element);
      linkParents(holder, (segments[0]!.run as { PARENT?: object }).PARENT ?? paragraph.p);
    }
    this.restyle(holder.value, true);
    mergeTextRuns(items, runsParent);
    (paragraph.parentBody.package_ as { requireStyles?: (ids: string[]) => void } | undefined)
      ?.requireStyles?.([HYPERLINK_STYLE]);
  }

  /**
   * Refuses, before anything changes, a link the setter could not make over [start, end): one with an
   * address on a range with no part, over no run, over runs in more than one holder, or around a
   * tracked deletion - each judged as the range would stand once the links it touches were unwrapped,
   * their runs then in the list holding them.
   */
  private refuseUnlinkable(start: number, end: number, address: string): void {
    const paragraph = this.paragraph;
    if (address !== '' && !paragraph.parentBody.part) {
      throw new Docx4JException('This range has no part, so a hyperlink cannot be related to it');
    }
    const touched = hyperlinkHolders(paragraph).filter((h) => this.touches(h)) as Element[];
    const unwrapsInto = new Map<Element[], Element[]>();
    for (const holder of touched) {
      const inner = runItemsOf(holder.value as object) as Element[] | undefined;
      if (inner) unwrapsInto.set(inner, containerOf(holder, paragraph));
    }
    const ownerOnceUnwrapped = (owner: Element[]): Element[] => {
      let o = owner;
      for (let guard = 0; guard < 100 && unwrapsInto.has(o); guard++) o = unwrapsInto.get(o)!;
      return o;
    };
    const covered = coveredSegments(paragraph, start, end);
    if (covered.length === 0) throw new Docx4JException('This range covers no run');
    const owner = ownerOnceUnwrapped(covered[0]!.runOwner);
    if (covered.some((seg) => ownerOnceUnwrapped(seg.runOwner) !== owner)) {
      throw new Docx4JException('This range spans more than one run holder (a hyperlink or a tracked change); wrap a narrower span');
    }
    const unwrapped = new Set(touched);
    const flatten = (items: Element[]): Element[] => items.flatMap((item) => (unwrapped.has(item)
      ? flatten((runItemsOf(item.value as object) ?? []) as Element[]) : [item]));
    refuseTrackedDeletion(flatten(owner), covered);
    const revision = revisionHolding(covered[0]!, owner);
    if (revision && itemsKeyOf(revision) === undefined) {
      throw new Docx4JException('This range is in a revision that cannot be split around a hyperlink');
    }
  }

  /** The first `w:hyperlink` this range touches, if any. */
  private hyperlinkHolder(): Element<wml.P.Hyperlink> | undefined {
    return hyperlinkHolders(this.paragraph).find((holder) => this.touches(holder));
  }

  /**
   * Unwraps every `w:hyperlink` the range touches, keeping its runs where they are and taking the
   * `Hyperlink` style off them. Until 0.2.1 this unwrapped every hyperlink of the paragraph, so
   * linking one word unlinked the others (CR-002 section 25).
   */
  private removeHyperlinks(): void {
    for (const holder of hyperlinkHolders(this.paragraph).filter((h) => this.touches(h))) {
      const owner = containerOf(holder as Element, this.paragraph);
      const at = owner.indexOf(holder as Element);
      if (at === -1) continue;
      this.restyle(holder.value, false);
      owner.splice(at, 1, ...((holder.value.content ?? []) as Element[]));
      for (const item of (holder.value.content ?? []) as Element[]) {
        linkParents(item, (holder as { PARENT?: object }).PARENT ?? this.paragraph.p);
      }
    }
  }

  /**
   * Whether this range touches a hyperlink's text: overlaps it, or for an empty range lies in it or
   * at either end of it - Office JS reads a caret just after a link as in that link (check 19, case
   * K; the start edge was not measured). A hyperlink holding no text touches nothing.
   */
  private touches(holder: Element<wml.P.Hyperlink>): boolean {
    const span = this.spanOf(holder);
    return span !== undefined && this.touchesSpan(span);
  }

  private touchesSpan([start, end]: [number, number]): boolean {
    return this.start === this.end ? start <= this.start && this.start <= end : start < this.end && end > this.start;
  }

  /** The text a hyperlink holds, as [start, end) in the paragraph, or undefined when it holds none. */
  private spanOf(holder: Element<wml.P.Hyperlink>): [number, number] | undefined {
    let start: number | undefined;
    let end = 0;
    for (const segment of this.paragraph.segments()) {
      if (!within(segment.run, holder.value, this.paragraph.p)) continue;
      start ??= segment.start;
      end = segment.end;
    }
    return start === undefined ? undefined : [start, end];
  }

  /**
   * The `Hyperlink` character style on every run of a hyperlink, in place of any character style
   * and with the direct colour removed, or off again when the link goes: what Word does (CR-002
   * section 25). A run whose properties are then empty loses its `w:rPr`, as in Word's output.
   */
  private restyle(holder: wml.P.Hyperlink, linked: boolean): void {
    const tracking = this.paragraph.fontTracking();
    const runs = new Set([...runsOf(holder), ...runsOf(holder, { view: 'original' })].map((r) => r.value));
    for (const run of runs) {
      if (!linked && run.rPr?.rStyle?.val !== HYPERLINK_STYLE) continue;
      run.rPr ??= { TYPE_NAME: 'org_docx4j_wml.RPr' };
      // the properties as they are now become w:rPrChange, unless this run is our own insertion
      if (tracking && !tracking.ownInsertions.has(run)) tracking.tracker.recordRPrChange(run.rPr);
      applyRunOptions(run.rPr, linked ? { style: HYPERLINK_STYLE, color: '' } : { style: '' });
      if (tracking) tracking.tracker.settleRPrChange(run);
      if (!run.rPr) continue;
      if (Object.keys(run.rPr).every((key) => key === 'TYPE_NAME')) delete run.rPr;
      else linkParents(run.rPr, run);
    }
  }

  /** The target of a relationship of this range's part, for the getter. */
  private relationshipTarget(id: string): string {
    const part = this.paragraph.parentBody.part;
    return part?.relationshipsPart?.getRelationshipById(id)?.target ?? '';
  }

  /** An external relationship of this range's part for the address, and its id. */
  private addHyperlinkRelationship(address: string): string {
    const part = this.paragraph.parentBody.part;
    if (!part) throw new Docx4JException('This range has no part, so a hyperlink cannot be related to it');
    return part.getRelationshipsPart(true)!.addExternalRelationship(Namespaces.HYPERLINK, address).id;
  }
}

/** What Office JS throws where nothing matches (`ItemNotFound`): its `code`, as `GeneralException` carries its own. */
class ItemNotFound extends Docx4JException {
  readonly code = 'ItemNotFound';
  constructor(message: string) {
    super(message);
    this.name = 'ItemNotFound';
  }
}

/** The character style Word gives the runs of a hyperlink it makes. */
const HYPERLINK_STYLE = 'Hyperlink';

/**
 * Adjacent runs holding nothing but text and the same properties become one, as Word writes a link
 * it makes: linking "beta gamma" over the end of a link on "alpha beta" gives Word one run
 * "alpha beta gamma" (`test/README.md` check 16, case F), where splitting and restyling left two.
 * Runs whose properties differ at all - a `w:rPrChange` under tracking among them - stay apart.
 */
function mergeTextRuns(items: Element[], parent: object): void {
  for (let i = items.length - 1; i > 0; i--) {
    const [a, b] = [items[i - 1]!, items[i]!];
    if (!isTextRun(a) || !isTextRun(b)) continue;
    const [ra, rb] = [a.value as wml.R, b.value as wml.R];
    if (canonical(ra.rPr) !== canonical(rb.rPr)) continue;
    const text = [...(ra.content ?? []), ...(rb.content ?? [])].map((item) => (item.value as wml.Text).value ?? '').join('');
    ra.content = [textElement(text)] as never;
    linkParents(ra.content, ra);
    items.splice(i, 1);
  }
  linkParents(items, parent);
}

function isTextRun(item: Element): boolean {
  return typeNameOf(item) === 'org_docx4j_wml.R'
    && ((item.value as wml.R).content ?? []).every((c) => (c as Element).name?.localPart === 't');
}

/** A value as text with its keys in order and `PARENT` left out, to compare run properties by content. */
function canonical(value: unknown): string {
  return JSON.stringify(value, (key, v) => {
    if (key === 'PARENT') return undefined;
    if (v === null || typeof v !== 'object' || Array.isArray(v)) return v;
    return Object.fromEntries(Object.keys(v).sort().map((k) => [k, (v as Record<string, unknown>)[k]]));
  });
}

/** Whether `object` is `ancestor` or below it, following `PARENT` no further than `stop`. */
function within(object: object, ancestor: object, stop: object): boolean {
  for (let o: object | undefined = object; o !== undefined && o !== stop; o = (o as { PARENT?: object }).PARENT) {
    if (o === ancestor) return true;
  }
  return false;
}

/** Every `w:hyperlink` of a paragraph, outermost first (they do not nest in practice). */
function hyperlinkHolders(paragraph: Paragraph): Element<wml.P.Hyperlink>[] {
  const out: Element<wml.P.Hyperlink>[] = [];
  const visit = (items: Element[] | undefined): void => {
    for (const item of items ?? []) {
      if (typeNameOf(item) === 'org_docx4j_wml.P.Hyperlink') out.push(item as Element<wml.P.Hyperlink>);
      const value = item.value;
      if (typeof value === 'object' && value !== null) visit(runItemsOf(value) as Element[] | undefined);
    }
  };
  visit(runItemsOf(paragraph.p) as Element[] | undefined);
  return out;
}

/** The run-level array holding an element, after `insertItemsAt` put it somewhere in the paragraph. */
/**
 * The segments a range covers, read before its runs are split at its edges: those inside it, and those
 * a split would cut, whose inner part then is. What `segments()` gives inside the range once it is
 * split, so that a wrapper's refusals can come before anything changes.
 */
function coveredSegments(paragraph: Paragraph, start: number, end: number): TextSegment[] {
  return paragraph.segments().filter((s) => (s.start >= start && s.end <= end) || (s.start < end && s.end > start));
}

/**
 * The items a range's runs are wrapped in place of: everything from the first run it covers to the
 * last, in the one list holding them, so that what lies between - a bookmark, a comment's range, a
 * proofing mark - goes into the wrapper with them. Until CR-002 section 43 as many items were taken as
 * there were runs, from the first, which dropped what lay between and left a run in two places.
 */
function wrappedItems(owner: Element[], segments: TextSegment[]): { at: number; items: Element[] } {
  const indexes = segments.map((s) => owner.indexOf(s.runOwner[s.runIndex] as Element)).filter((i) => i >= 0);
  const at = Math.min(...indexes);
  return { at, items: owner.slice(at, Math.max(...indexes) + 1) };
}

/**
 * Refuses a range with a tracked deletion (or text moved away) between its runs: having no text in
 * the range's, it would be wrapped unseen, and a range spanning a tracked change is refused (CR-002
 * section 43).
 */
function refuseTrackedDeletion(items: Element[], segments: TextSegment[]): void {
  const indexes = segments.map((s) => items.indexOf(s.runOwner[s.runIndex] as Element)).filter((i) => i >= 0);
  if (indexes.length === 0) return;
  for (const item of items.slice(Math.min(...indexes), Math.max(...indexes) + 1)) {
    const kind = revisionKindOf(item);
    if (kind === 'del' || kind === 'moveFrom') {
      throw new Docx4JException('This range spans a tracked deletion; wrap the text on either side of it');
    }
  }
}

/** The `w:ins` or `w:moveTo` a segment's run is directly in, when `owner` is its own item list. */
function revisionHolding(segment: TextSegment, owner: Element[]): RevisionHolder | undefined {
  return segment.revision !== undefined && segment.revision.items === owner ? segment.revision : undefined;
}

/** The key under which a revision's value holds its items, or undefined where none does. */
function itemsKeyOf(revision: RevisionHolder): string | undefined {
  const value = revision.value as unknown as Record<string, unknown>;
  return Object.keys(value).find((key) => value[key] === revision.items);
}

/** A revision of the same kind and attributes as `revision`, with a new id, holding `items`. */
function revisionAround(revision: RevisionHolder, items: Element[], id: number): Element {
  const value = revision.value as unknown as Record<string, unknown>;
  const itemsKey = itemsKeyOf(revision)!;
  const inner: Record<string, unknown> = {};
  for (const [key, v] of Object.entries(value)) if (key !== itemsKey && key !== 'PARENT') inner[key] = v;
  inner.id = id;
  inner[itemsKey] = items;
  linkParents(items, inner);
  return { name: revision.element.name, value: inner } as Element;
}

/**
 * Puts `item` where items from `index` on were taken out of `revision`: in its place where it holds
 * nothing now, before it where they were its first, else between its halves. False where the revision
 * cannot be split.
 */
function placeInRevision(paragraph: Paragraph, revision: RevisionHolder, index: number, item: Element, nextId: () => number): boolean {
  const owner = revision.owner;
  const parent = (revision.value as { PARENT?: object }).PARENT ?? paragraph.p;
  if (revision.items.length === 0) {
    owner.splice(owner.indexOf(revision.element), 1, item);                       // the whole revision wrapped
    linkParents(item, parent);
    return true;
  }
  if (index === 0) {
    owner.splice(owner.indexOf(revision.element), 0, item);                       // its start wrapped
    linkParents(item, parent);
    return true;
  }
  return paragraph.splitHolderAround(revision, index, item, nextId);
}

function containerOf(element: Element, paragraph: Paragraph): Element[] {
  const seek = (items: Element[] | undefined): Element[] | undefined => {
    if (!items) return undefined;
    if (items.includes(element)) return items;
    for (const item of items) {
      const value = item.value;
      if (typeof value !== 'object' || value === null) continue;
      const found = seek(runItemsOf(value));
      if (found) return found;
    }
    return undefined;
  };
  return seek(runItemsOf(paragraph.p)) ?? ((paragraph.p.content ??= []) as Element[]);
}
