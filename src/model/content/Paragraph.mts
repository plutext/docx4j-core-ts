import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy, marshalString } from '@docx4j/generated-objects-ts';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { r as textRun, t as textItem, br as breakItem } from '@docx4j/generated-objects-ts/builders/wml';
import { type Element, segmentsOf, runsOf, runItemsOf, linkParents, textOf, textOfView, typeNameOf, W_NS, runOf, paragraphOf, revisionKindOf, type TextSegment, type TextViewOptions } from './tree.mjs';

function withRPr(run: Element<wml.R>, rPr: wml.RPr | undefined): Element<wml.R> {
  if (rPr) run.value.rPr = rPr;
  return run;
}
import { Font, type FontTracking } from './Font.mjs';
import { runFontSelectorOf } from '../fonts/lookup.mjs';
import { Range } from './Range.mjs';
import { searchPattern, findAll, type SearchOptions } from './search.mjs';
import type { Body, BlockElement } from './Body.mjs';
import { builtInOf, idOfBuiltIn, styleNameOf, styleIdOf } from './styles.mjs';
import { cellOf, type TableCell } from './Table.mjs';
import { ContentControl, collectRunControls, type ContentControlType } from './ContentControl.mjs';
import { sdt as sdtOf, nextSdtId } from '@docx4j/generated-objects-ts/builders/wml';
import { controlIdScope, sdtKindFor } from '../customxml/insert.mjs';
import { InlinePicture, addImage, writableWidthEmu, type InlinePictureOptions } from './InlinePicture.mjs';
import { contentOf } from './ooxml.mjs';
import { commentApi } from './comments.mjs';
import type { Comment } from './Comment.mjs';
import { type ChangeTracker, copyRPr, markDeleted, toDeletedText } from './tracking.mjs';
import { TrackedChange, trackedChangesOfParagraph } from './TrackedChange.mjs';

/**
 * Office JS `Word.Alignment`, as an *effective* read reports it: an absent `w:jc` resolves to
 * `'Left'`, which is what Word does and what Office JS answers. `'Unknown'` is what a direct
 * read (`formatting({ direct: true })`) says for "no `w:jc` here", and is accepted by the
 * setter, where it removes the element (CR-001 Phase B step 2 decision 4).
 */
export type Alignment = 'Left' | 'Centered' | 'Right' | 'Justified';

/** {@link Alignment} plus the direct reads' "nothing stated" and the setter's "remove it". */
export type AlignmentOrUnknown = Alignment | 'Unknown';

/** Whether a read reports direct formatting instead of the effective value. */
export interface FormattingOptions {
  /**
   * `true` reads the direct formatting only - the `w:pPr` or `w:rPr` on the element itself -
   * as this package did before CR-001 Phase B. The default is the effective value, resolved
   * through the document defaults and the style chain, which is what Office JS reports.
   */
  direct?: boolean;
}

/** Office JS's paragraph formatting properties, read in one call (extension). */
export interface ParagraphFormatting {
  alignment: AlignmentOrUnknown;
  leftIndent: number;
  rightIndent: number;
  firstLineIndent: number;
  spaceBefore: number;
  spaceAfter: number;
  lineSpacing: number;
  outlineLevel: number;
}

const TWIPS_PER_POINT = 20;

/**
 * A subset of Office JS `Word.Paragraph` over a `w:p` in a body, cell, header, footer or
 * content control. A light view: it holds the element and the array that contains it; nothing
 * is cached, so views may be created and discarded freely. Every write mutates the tree in
 * place and links PARENT on what it adds.
 */
export class Paragraph {
  constructor(
    /** The `w:p` element pair. */
    readonly element: Element<wml.P>,
    /** The content array holding the element (a body's, a cell's, a content control's). */
    readonly container: Element[],
    /** The nearest body view (the part's), for addresses and paragraph ids. */
    readonly parentBody: Body,
  ) {}

  /** The paragraph value (docx4j: the P). */
  get p(): wml.P {
    return this.element.value;
  }

  get index(): number {
    return this.container.indexOf(this.element);
  }

  /** The paragraph's text: runs joined, tabs as \t, breaks as \n; deleted text excluded. */
  get text(): string {
    return textOf(this.p);
  }

  /** Replaces the content with one run of the text, keeping the first run's formatting. */
  set text(value: string) {
    if (this.changeTracker) { this.splice(0, this.text.length, value); return; }
    const rPr = runsOf(this.p)[0]?.value.rPr;
    this.p.content = [withRPr(textRun(value), rPr ? deepCopy(rPr) : undefined)] as wml.P['content'];
    linkParents(this.p.content, this.p);
  }

  /**
   * The text in one of the two views (extension; `text` is the accepted one). `{ view: 'original' }`
   * reads the document as it was before the tracked changes: `w:del` included, `w:ins` excluded.
   */
  getText(options?: TextViewOptions): string {
    return options?.view === 'original' ? textOfView(this.p, options) : this.text;
  }

  /** The style id (w:pStyle; docx4j's name for it); 'Normal' when none is set. Extension: Office JS has `style` and `styleBuiltIn` only. */
  get styleId(): string {
    return this.p.pPr?.pStyle?.val ?? 'Normal';
  }
  set styleId(id: string) {
    // through pPr(), so that a tracked write records w:pPrChange first (CR-002 phase F)
    const pPr = this.pPr();
    if (id === '' || id === 'Normal') { delete pPr.pStyle; return; }
    pPr.pStyle = { val: id };
  }

  /**
   * The style's display name ('Heading 1', 'My Style'), as Office JS: from the styles part's
   * `w:name` when that part is unmarshalled, else derived from the id. Setting accepts a display
   * name, a stored name, or an id.
   */
  get style(): string {
    return styleNameOf(this.parentBody.package_, this.styleId);
  }
  set style(name: string) {
    this.styleId = styleIdOf(this.parentBody.package_, name);
  }

  /** The `Word.Style` value ('Heading1'), or 'Other' when the style is not a built-in one, as Office JS. */
  get styleBuiltIn(): string {
    return builtInOf(this.styleId);
  }
  set styleBuiltIn(value: string) {
    this.styleId = idOfBuiltIn(value);
  }

  /**
   * The paragraph properties which actually apply: the document defaults, the style chain and
   * the numbering level, then this paragraph's own `w:pPr` (docx4j
   * `PropertyResolver.getEffectivePPr`). A live, possibly cached object: copy it before
   * changing it. Needs the package's resolver; see `Body.propertyResolver`.
   */
  get effectivePPr(): wml.PPr {
    return this.parentBody.propertyResolver.getEffectivePPr(this.p.pPr);
  }

  /** The paragraph mark's effective run properties (docx4j `getEffectiveParagraphMarkRPr`). */
  get effectiveParagraphMarkRPr(): wml.RPr {
    return this.parentBody.propertyResolver.getEffectiveParagraphMarkRPr(this.p.pPr);
  }

  /** The `w:pPr` a formatting read works from: the resolved one, or this paragraph's own. */
  private readPPr(options?: FormattingOptions): wml.PPr | undefined {
    return options?.direct === true ? this.p.pPr : this.effectivePPr;
  }

  /**
   * Office JS's paragraph formatting in one call, and the way to ask for direct formatting:
   * `formatting({ direct: true })` is what the individual getters reported before CR-001
   * Phase B step 2 (an absent property reads 0, 10 or `'Unknown'`).
   */
  formatting(options?: FormattingOptions): ParagraphFormatting {
    const pPr = this.readPPr(options);
    const effective = options?.direct !== true;
    return {
      alignment: alignmentOf(pPr, effective),
      leftIndent: (pPr?.ind?.left ?? 0) / TWIPS_PER_POINT,
      rightIndent: (pPr?.ind?.right ?? 0) / TWIPS_PER_POINT,
      firstLineIndent: firstLineIndentOf(pPr),
      spaceBefore: (pPr?.spacing?.before ?? 0) / TWIPS_PER_POINT,
      spaceAfter: (pPr?.spacing?.after ?? 0) / TWIPS_PER_POINT,
      lineSpacing: lineSpacingOf(pPr),
      outlineLevel: outlineLevelOf(pPr),
    };
  }

  /** The effective alignment; an absent `w:jc` after resolution is `'Left'`, as Office JS. */
  get alignment(): Alignment {
    return alignmentOf(this.effectivePPr, true) as Alignment;
  }
  set alignment(v: AlignmentOrUnknown) {
    const pPr = this.pPr();
    const val: wml.JcEnumeration | undefined = v === 'Left' ? 'left' : v === 'Centered' ? 'center' : v === 'Right' ? 'right' : v === 'Justified' ? 'both' : undefined;
    if (val === undefined) { delete pPr.jc; return; }
    pPr.jc = { val };
  }

  /** Indents and spacing in points, as Office JS; the effective values (CR-001 Phase B step 2). */
  get leftIndent(): number { return (this.effectivePPr.ind?.left ?? 0) / TWIPS_PER_POINT; }
  set leftIndent(pt: number) { this.ind().left = Math.round(pt * TWIPS_PER_POINT); }
  get rightIndent(): number { return (this.effectivePPr.ind?.right ?? 0) / TWIPS_PER_POINT; }
  set rightIndent(pt: number) { this.ind().right = Math.round(pt * TWIPS_PER_POINT); }
  /** First-line indent; negative for a hanging indent, as Office JS. */
  get firstLineIndent(): number {
    return firstLineIndentOf(this.effectivePPr);
  }
  set firstLineIndent(pt: number) {
    const ind = this.ind();
    if (pt < 0) { ind.hanging = Math.round(-pt * TWIPS_PER_POINT); delete ind.firstLine; } else { ind.firstLine = Math.round(pt * TWIPS_PER_POINT); delete ind.hanging; }
  }
  get spaceBefore(): number { return (this.effectivePPr.spacing?.before ?? 0) / TWIPS_PER_POINT; }
  set spaceBefore(pt: number) { this.spacing().before = Math.round(pt * TWIPS_PER_POINT); }
  get spaceAfter(): number { return (this.effectivePPr.spacing?.after ?? 0) / TWIPS_PER_POINT; }
  set spaceAfter(pt: number) { this.spacing().after = Math.round(pt * TWIPS_PER_POINT); }
  /** Line spacing in points: w:line/240 lines of 12pt when the rule is auto, else w:line twips; 0 when not set. */
  get lineSpacing(): number {
    return lineSpacingOf(this.effectivePPr);
  }
  set lineSpacing(pt: number) { const s = this.spacing(); s.line = Math.round(pt * TWIPS_PER_POINT); s.lineRule = 'exact'; }
  /** w:outlineLvl + 1 (1 to 9); 10 for body text, as Office JS. A heading style supplies it. */
  get outlineLevel(): number { return outlineLevelOf(this.effectivePPr); }
  set outlineLevel(level: number) { const pPr = this.pPr(); if (level >= 10 || level < 1) delete pPr.outlineLvl; else pPr.outlineLvl = { val: level - 1 }; }

  /**
   * The runs' formatting: reads the first run's *effective* properties, writes direct
   * formatting to all (extension: `Font` is over runs, not the paragraph mark).
   */
  get font(): Font {
    return this.getFont();
  }

  /** `font`, with `{ direct: true }` for the direct formatting of the first run (extension). */
  getFont(options?: FormattingOptions): Font {
    const holders = (): wml.R[] => runsOf(this.p).map((r) => r.value);
    if (options?.direct === true) return new Font(holders, () => this.fontTracking());
    return new Font(holders, () => this.fontTracking(),
      (rPr) => this.parentBody.propertyResolver.getEffectiveRPr(rPr, this.p.pPr),
      (rPr) => runFontSelectorOf(this.parentBody.package_)?.asciiFontName(rPr));
  }

  /** w14:paraId, the stable address Word gives paragraphs. */
  get paraId(): string | undefined {
    return this.p.paraId;
  }
  set paraId(id: string | undefined) {
    if (id === undefined) delete this.p.paraId; else this.p.paraId = id;
  }

  /** The runs, direct and nested (hyperlinks, content controls, insertions), in order. */
  get runs(): Element<wml.R>[] {
    return runsOf(this.p);
  }

  /** The cell this paragraph is in, or undefined when it is not in a table (Office JS). */
  get parentTableCell(): TableCell | undefined {
    return cellOf(this.p, this.parentBody);
  }

  /** The run-level content controls in this paragraph, in order, nested ones included. */
  get contentControls(): ContentControl[] {
    const out: ContentControl[] = [];
    collectRunControls(this.p, this.parentBody, out);
    return out;
  }

  /** The inline pictures in this paragraph, in order (Office JS Paragraph.inlinePictures). */
  get inlinePictures(): InlinePicture[] {
    const out: InlinePicture[] = [];
    for (const run of runsOf(this.p)) {
      for (const item of (run.value.content ?? []) as Element[]) {
        if (item.name.localPart !== 'drawing') continue;
        const inline = (item.value as wml.Drawing).anchorOrInline?.[0];
        if (inline && (inline as { TYPE_NAME?: string }).TYPE_NAME === 'org_docx4j_dml_wordprocessingDrawing.Inline') {
          out.push(new InlinePicture(item as Element<wml.Drawing>, run, this));
        }
      }
    }
    return out;
  }

  /** Text inserted at the start, the end, or replacing the whole paragraph's text. */
  insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range {
    const length = this.text.length;
    if (location === 'Start') return this.splice(0, 0, text);
    if (location === 'End') return this.splice(length, length, text);
    return this.splice(0, length, text);
  }

  /** A new paragraph of the text before or after this one, with this paragraph's properties (as Word). */
  insertParagraph(text: string, location: 'Before' | 'After'): Paragraph {
    const pPr = this.p.pPr ? deepCopy(this.p.pPr) : undefined;
    if (pPr) { delete pPr.rPr; delete (pPr as { sectPr?: unknown }).sectPr; delete pPr.pPrChange; }
    const el = paragraphOf(text === '' ? [] : [textRun(text)], pPr);
    return this.parentBody.insertElement(el, location, this) as Paragraph;
  }

  insertBreak(type: 'Page' | 'Line', location: 'Before' | 'After' | 'Start' | 'End'): void {
    const brk = breakItem(type === 'Page' ? 'page' : undefined);
    if (location === 'Before' || location === 'After') {
      const p = this.insertParagraph('', location);
      const run = runOf([brk]);
      const tracker = p.changeTracker;
      const item = tracker ? tracker.ins([run]) : run;
      p.p.content!.push(item as never);
      linkParents(item, p.p);
      return;
    }
    const run = runOf([brk]);
    const tracker = this.changeTracker;
    const item = tracker ? tracker.ins([run]) : run;
    const content = (this.p.content ??= []);
    if (location === 'Start') content.unshift(item as never); else content.push(item as never);
    linkParents(item, this.p);
  }

  /**
   * Adds the image as a part of this paragraph's part, with a relationship, and shows it at the
   * start or the end of this paragraph, or in place of its content (Office JS).
   */
  insertInlinePictureFromBase64(base64: string, location: 'Start' | 'End' | 'Replace', options?: InlinePictureOptions): InlinePicture {
    const part = this.parentBody.part;
    if (!part) throw new Docx4JException('This body has no part, so an image cannot be related to it');
    const { run, drawing } = addImage(part, base64, this.parentBody.container, writableWidthEmu(this.parentBody.container), options);
    const content = (this.p.content ??= []) as Element[];
    if (location === 'Replace') content.length = 0;
    if (location === 'Start') content.unshift(run as Element); else content.push(run as Element);
    linkParents(run, this.p);
    return new InlinePicture(drawing, run, this);
  }

  /**
   * Word's `insertOoxml` at paragraph level: a flat OPC `pkg:package` (or a bare fragment). One
   * incoming paragraph is merged into this one at 'Start' and 'End', as Word's paste does;
   * anything else is inserted as blocks before or after. 'Replace' puts the content where this
   * paragraph was and removes it.
   */
  async insertOoxml(ooxml: string, location: 'Before' | 'After' | 'Start' | 'End' | 'Replace'): Promise<(Paragraph | BlockElement)[]> {
    const preprocess = this.parentBody.package_?.loadOptions.preprocessor;
    const elements = await contentOf(ooxml, { preprocess: preprocess ? (doc) => preprocess(doc) : undefined, target: this.parentBody.part });
    if (elements.length === 0) return [];
    const only = elements.length === 1 && typeNameOf(elements[0]!) === 'org_docx4j_wml.P' ? (elements[0]!.value as wml.P) : undefined;
    if (only && (location === 'Start' || location === 'End')) {
      this.insertItemsAt(location === 'Start' ? 0 : this.text.length, (only.content ?? []) as Element[]);
      return [this];
    }
    if (location === 'Replace') {
      const inserted = this.parentBody.insertElement(elements, 'Before', this);
      void inserted;
      this.delete();
    } else {
      this.parentBody.insertElement(elements, location === 'Start' ? 'Before' : location === 'End' ? 'After' : location, this);
    }
    return elements.map((el) => typeNameOf(el) === 'org_docx4j_wml.P'
      ? new Paragraph(el as Element<wml.P>, this.container, this.parentBody)
      : { element: el, container: this.container });
  }

  /** Matches within this paragraph. */
  search(text: string, options?: SearchOptions): Range[] {
    return findAll(this.text, searchPattern(text, options)).map(([s, e]) => new Range(this, s, e));
  }

  /**
   * Replaces every match of `find` with `replace`, last match first so that the offsets of the
   * ones still to do stay valid (CR-002 section 3.7). Returns the number replaced; tracked when
   * the package's `changeTrackingMode` is on.
   */
  replaceText(find: string, replace: string, options?: SearchOptions): number {
    const matches = this.search(find, options);
    for (let i = matches.length - 1; i >= 0; i--) matches[i]!.insertText(replace, 'Replace');
    return matches.length;
  }

  /** The tracked changes in this paragraph, in document order (CR-002 phase F). */
  getTrackedChanges(): TrackedChange[] {
    return trackedChangesOfParagraph(this);
  }

  getRange(location: 'Whole' | 'Start' | 'End' | 'Content' = 'Whole'): Range {
    const length = this.text.length;
    if (location === 'Start') return new Range(this, 0, 0);
    if (location === 'End') return new Range(this, length, length);
    return new Range(this, 0, length);
  }

  /**
   * Removes the paragraph from its container. While the package tracks changes, the content
   * becomes a `w:del` and the mark is marked deleted (`w:pPr/w:rPr/w:del`) instead, unless the
   * whole paragraph was this author's own insertion, which Word simply takes back.
   */
  delete(): void {
    const tracker = this.changeTracker;
    if (tracker) { this.trackedDelete(tracker); return; }
    const i = this.index;
    if (i >= 0) this.container.splice(i, 1);
  }

  private trackedDelete(tracker: ChangeTracker): void {
    if (markDeleted(this.p)) throw new Docx4JException('This paragraph is already marked deleted');
    const length = this.text.length;
    if (length > 0) this.deleteText(tracker, 0, length);
    const rPr = this.p.pPr?.rPr;
    const ownMark = rPr?.ins !== undefined && rPr.ins.author === tracker.author;
    if (ownMark && (this.p.content ?? []).length === 0) {
      const i = this.index;
      if (i >= 0) this.container.splice(i, 1);
      return;
    }
    tracker.markParagraphDeleted(this.p);
  }

  /** The paragraph as XML (extension; Office JS getOoxml wraps it in a package). */
  getXml(): Promise<string> {
    return marshalString(this.element as Element);
  }

  // --- docx4j names ---

  /** docx4j: the P's content list. */
  getContent(): Element[] {
    return (this.p.content ??= []) as Element[];
  }

  // --- editing primitives, shared with Range ---

  /**
   * The paragraph's properties, created when absent. Every property setter goes through this
   * one accessor, so this is where a tracked write records `w:pPrChange` with the properties as
   * they stand before it.
   */
  private pPr(): wml.PPr {
    const pPr = (this.p.pPr ??= { TYPE_NAME: 'org_docx4j_wml.PPr' });
    linkParents(pPr, this.p);
    this.changeTracker?.recordPPrChange(pPr);
    return pPr;
  }
  private ind(): wml.PPrBase.Ind {
    return (this.pPr().ind ??= {});
  }
  private spacing(): wml.PPrBase.Spacing {
    return (this.pPr().spacing ??= {});
  }

  /** The package's change tracker while `changeTrackingMode` is on; undefined when it is off. */
  get changeTracker(): ChangeTracker | undefined {
    return this.parentBody.changeTracker;
  }

  /** What `Font` needs to write `w:rPrChange`: the tracker, and the runs that are our own insertion (`Range` shares it). */
  fontTracking(): FontTracking | undefined {
    const tracker = this.changeTracker;
    if (!tracker) return undefined;
    const ownInsertions = new Set<object>();
    for (const seg of this.segments()) if (tracker.ownInsertion(seg.revision)) ownInsertions.add(seg.run);
    return { tracker, ownInsertions };
  }

  /** The text segments (w:t, w:tab, ...) with their offsets. */
  segments(options?: TextViewOptions): TextSegment[] {
    return segmentsOf(this.p, options);
  }

  /**
   * Replaces the text in [start, end) with `text` and returns the range of the inserted text.
   * Runs are edited in place: a w:t loses or gains characters, a tab or break inside the range
   * is removed, a run left empty is removed. Text is inserted into the w:t at `start`,
   * else into a new run with the formatting of the nearest run.
   */
  splice(start: number, end: number, text: string): Range {
    const length = this.text.length;
    start = Math.max(0, Math.min(start, length));
    end = Math.max(start, Math.min(end, length));
    const tracker = this.changeTracker;
    if (tracker) {
      // a replacement is the w:del first and the w:ins after it, as Word writes one
      const deletions = end > start ? this.deleteText(tracker, start, end) : [];
      if (text.length > 0) this.insertTracked(tracker, start, text, deletions[deletions.length - 1]);
      return new Range(this, start, start + text.length);
    }
    let segs = this.segments();
    let inserted = false;
    if (end > start) {
      // Replace: the new text takes the place (and formatting) of the first replaced character, as Word does.
      const first = segs.find((s) => s.start <= start && start < s.end);
      const toRemove: TextSegment[] = [];
      for (const seg of segs) {
        if (seg.end <= start || seg.start >= end) continue;
        const from = Math.max(0, start - seg.start);
        const to = Math.min(seg.text.length, end - seg.start);
        if (seg.editable) {
          const t = seg.item.value as wml.Text;
          const middle = seg === first && !inserted ? text : '';
          if (seg === first) inserted = true;
          const value = seg.text.substring(0, from) + middle + seg.text.substring(to);
          if (value === '') toRemove.push(seg); else { t.value = value; if (/^\s|\s$|\s\s/.test(value)) t.space = 'preserve'; }
        } else toRemove.push(seg);
      }
      for (const seg of toRemove.reverse()) {
        const i = seg.owner.indexOf(seg.item);
        if (i >= 0) seg.owner.splice(i, 1);
      }
      if (inserted && text.length > 0) {
        this.removeEmptyRuns();
        return new Range(this, start, start + text.length);
      }
      this.removeEmptyRuns();
      segs = this.segments();
    }
    if (text.length > 0) {
      // Insert: extend the w:t at `start` (the one containing it, else the one ending there, else the one starting there)
      const target = segs.find((s) => s.editable && s.start < start && start < s.end)
        ?? segs.find((s) => s.editable && s.end === start)
        ?? segs.find((s) => s.editable && s.start === start);
      if (target) {
        const t = target.item.value as wml.Text;
        const at = start - target.start;
        t.value = target.text.substring(0, at) + text + target.text.substring(at);
        if (/^\s|\s$|\s\s/.test(t.value)) t.space = 'preserve';
      } else {
        // no w:t to extend: a new run next to the segment at `start`, with that run's formatting
        const before = [...segs].reverse().find((s) => s.end <= start);
        const after = segs.find((s) => s.start >= start);
        const neighbour = before ?? after;
        const firstRun = runsOf(this.p)[0]?.value.rPr;
        const rPr = neighbour?.run.rPr ? deepCopy(neighbour.run.rPr) : firstRun ? deepCopy(firstRun) : undefined;
        const run = runOf([textItem(text)], rPr);
        if (before) before.runOwner.splice(before.runIndex + 1, 0, run);
        else if (after) after.runOwner.splice(after.runIndex, 0, run);
        else {
          const content = (this.p.content ??= []);
          if (start === 0) content.unshift(run as never); else content.push(run as never);
        }
        linkParents(run, this.p);
      }
    }
    return new Range(this, start, start + text.length);
  }

  /**
   * Inserts run-level items (runs, hyperlinks, content controls) at a text offset, splitting the
   * run there when the offset falls inside one. PARENT is linked on what is inserted.
   */
  insertItemsAt(offset: number, items: Element[]): void {
    if (items.length === 0) return;
    this.splitAt(offset);
    const segs = this.segments();
    const after = segs.find((s) => s.start >= offset);
    const before = [...segs].reverse().find((s) => s.end <= offset);
    const neighbour = after ?? before;
    const list = neighbour ? neighbour.runOwner : ((this.p.content ??= []) as Element[]);
    const at = neighbour ? (after ? neighbour.runIndex : neighbour.runIndex + 1) : (offset === 0 ? 0 : list.length);
    list.splice(at, 0, ...items);
    const owner = (neighbour?.run as { PARENT?: object } | undefined)?.PARENT ?? this.p;
    for (const item of items) linkParents(item, owner);
  }

  /** Splits the run at a text offset so that [offset, ...) begins a run; returns nothing when the offset is already a boundary. */
  splitAt(offset: number): void {
    const seg = this.segments().find((s) => s.editable && s.start < offset && offset < s.end);
    if (!seg) return;
    const at = offset - seg.start;
    const t = seg.item.value as wml.Text;
    const head = seg.text.substring(0, at);
    const tail = seg.text.substring(at);
    t.value = head;
    if (/^\s|\s$|\s\s/.test(head)) t.space = 'preserve';
    const rest = seg.owner.splice(seg.index + 1);
    const second = runOf([textItem(tail), ...rest], seg.run.rPr ? deepCopy(seg.run.rPr) : undefined);
    seg.runOwner.splice(seg.runIndex + 1, 0, second);
    linkParents(second, (seg.run as { PARENT?: object }).PARENT ?? this.p);
  }

  // --- comments (CR-002 phase G) ---

  /** The comments anchored in this paragraph, replies nested under their parent. */
  async getComments(): Promise<Comment[]> {
    return commentApi().commentsOf(this);
  }

  /** Comments the whole paragraph (sugar over `getRange().insertComment`). */
  async insertComment(text: string): Promise<Comment> {
    return commentApi().insertComment(this.getRange(), text);
  }

  // --- content controls (CR-002 phase E) ---

  /**
   * Wraps this paragraph in a new block-level content control (Office JS
   * `Paragraph.insertContentControl`), typed for the kind when one is given. The `w:id` is free in
   * the body. Returns the control.
   */
  insertContentControl(kind?: ContentControlType): ContentControl {
    const at = this.index;
    if (at < 0) throw new Docx4JException('This paragraph is not in its container any more');
    const sdt = sdtOf([this.element as Element], { kind: sdtKindFor(kind), id: nextSdtId(controlIdScope(this.parentBody)), form: 'block' });
    this.container.splice(at, 1, sdt as Element);
    linkParents(sdt, (this.p as { PARENT?: object }).PARENT ?? this.parentBody.container);
    return new ContentControl(sdt as Element<wml.SdtBlock>, this.container, this.parentBody);
  }

  private removeEmptyRuns(): void {
    const prune = (items: Element[]): void => {
      for (let i = items.length - 1; i >= 0; i--) {
        const el = items[i]!;
        const v = el.value as { TYPE_NAME?: string; content?: Element[]; sdtContent?: { content?: Element[] } };
        if (v.TYPE_NAME === 'org_docx4j_wml.R') {
          if (!v.content || v.content.length === 0) items.splice(i, 1);
        } else if (v.TYPE_NAME !== 'org_docx4j_wml.P') {
          const nested = runItemsOf(v);
          if (nested) {
            prune(nested);
            // a revision left with nothing in it goes too
            if (nested.length === 0 && revisionKindOf(el) !== undefined) items.splice(i, 1);
          }
        }
      }
    };
    prune(this.p.content ?? []);
  }

  // --- tracked editing (CR-002 phase F) ---------------------------------------------------

  /**
   * The text in [start, end) becomes a deletion: the runs holding it are isolated, their `w:t`
   * turned into `w:delText`, and each run of consecutive ones moved into one `w:del`. Text this
   * author had inserted is simply removed, as Word does, rather than nested in a `w:del`.
   * Returns the `w:del` elements it made, in document order.
   */
  private deleteText(tracker: ChangeTracker, start: number, end: number): Anchor[] {
    const targets = this.isolate(start, end);
    for (const seg of targets) tracker.assertEditable(seg.revision);
    type Target = { run: Element; owner: Element[]; own: boolean; revision?: RevisionHolderOf };
    const list: Target[] = [];
    for (const seg of targets) {
      const run = seg.runOwner[seg.runIndex]!;
      if (list.length > 0 && list[list.length - 1]!.run === run) continue;
      const target: Target = { run, owner: seg.runOwner, own: tracker.ownInsertion(seg.revision) };
      if (seg.revision) target.revision = seg.revision;
      list.push(target);
    }
    // groups of consecutive runs in the same array, so that one w:del holds them all
    const groups: Target[][] = [];
    for (const t of list) {
      const last = groups[groups.length - 1];
      const previous = last?.[last.length - 1];
      if (last && previous && previous.own === t.own && previous.owner === t.owner
        && t.owner.indexOf(t.run) === t.owner.indexOf(previous.run) + 1) last.push(t);
      else groups.push([t]);
    }
    const made: Anchor[] = [];
    for (let g = groups.length - 1; g >= 0; g--) {
      const group = groups[g]!;
      const owner = group[0]!.owner;
      const at = owner.indexOf(group[0]!.run);
      const runs = owner.splice(at, group.length);
      if (group[0]!.own) {
        // taking back our own insertion: the runs go, and an emptied w:ins with them
        const revision = group[0]!.revision;
        if (revision && revision.items.length === 0) {
          const i = revision.owner.indexOf(revision.element);
          if (i >= 0) revision.owner.splice(i, 1);
        }
        continue;
      }
      for (const run of runs) toDeletedText(run.value as wml.R);
      const del = tracker.del(runs);
      owner.splice(at, 0, del);
      linkParents(del, (group[0]!.revision?.value ?? (group[0]!.run.value as { PARENT?: object }).PARENT ?? this.p) as object);
      made.unshift({ owner, element: del, parent: (group[0]!.revision?.value ?? (del.value as { PARENT?: object }).PARENT ?? this.p) as object });
    }
    return made;
  }

  /**
   * Text inserted at `at` as a `w:ins`. A run of this author's own that is already inside a
   * `w:ins` is extended rather than nested in another one, as Word does. `anchor` is the `w:del`
   * of a replacement, which the insertion must follow.
   */
  private insertTracked(tracker: ChangeTracker, at: number, text: string, anchor?: Anchor): void {
    if (anchor) {
      const rPr = copyRPr(runsOf(anchor.element.value as object, { view: 'original' })[0]?.value.rPr);
      const ins = tracker.ins([runOf([textItem(text)], rPr)]);
      anchor.owner.splice(anchor.owner.indexOf(anchor.element) + 1, 0, ins);
      linkParents(ins, anchor.parent);
      return;
    }
    this.splitAt(at);
    const segs = this.segments();
    const before = [...segs].reverse().find((s) => s.end <= at);
    const after = segs.find((s) => s.start >= at);
    for (const [seg, side] of [[before, 'after'], [after, 'before']] as const) {
      if (!seg || !tracker.ownInsertion(seg.revision)) continue;
      if (seg.editable) {
        const t = seg.item.value as wml.Text;
        t.value = side === 'after' ? seg.text + text : text + seg.text;
        if (/^\s|\s$|\s\s/.test(t.value)) t.space = 'preserve';
      } else {
        const run = runOf([textItem(text)], copyRPr(seg.run.rPr));
        seg.runOwner.splice(side === 'after' ? seg.runIndex + 1 : seg.runIndex, 0, run);
        linkParents(run, seg.revision!.value);
      }
      return;
    }
    const neighbour = before ?? after;
    tracker.assertEditable(neighbour?.revision);
    const rPr = copyRPr(neighbour?.run.rPr ?? runsOf(this.p)[0]?.value.rPr);
    const ins = tracker.ins([runOf([textItem(text)], rPr)]);
    if (neighbour) {
      const owner = neighbour.revision ? neighbour.revision.owner : neighbour.runOwner;
      const item = neighbour.revision ? neighbour.revision.element : neighbour.runOwner[neighbour.runIndex]!;
      const parent = (neighbour.revision ? neighbour.revision.value : neighbour.run) as { PARENT?: object };
      owner.splice(owner.indexOf(item) + (neighbour === before ? 1 : 0), 0, ins);
      linkParents(ins, parent.PARENT ?? this.p);
    } else {
      const content = (this.p.content ??= []);
      if (at === 0) content.unshift(ins as never); else content.push(ins as never);
      linkParents(ins, this.p);
    }
  }

  /**
   * Splits runs so that every run holding text in [start, end) holds nothing else, and returns
   * the segments inside the span. `splitAt` does the `w:t` boundaries; this also moves the items
   * of a run that straddles a boundary (a tab, a break, a drawing) out into runs of their own.
   */
  private isolate(start: number, end: number): TextSegment[] {
    this.splitAt(start);
    this.splitAt(end);
    for (let guard = 0; guard < 10_000; guard++) {
      const inside = this.segments().filter((s) => s.start >= start && s.end <= end && s.text.length > 0);
      let split = false;
      const seen = new Set<wml.R>();
      for (const seg of inside) {
        if (seen.has(seg.run)) continue;
        seen.add(seg.run);
        const ofRun = inside.filter((s) => s.run === seg.run);
        const first = ofRun[0]!;
        const last = ofRun[ofRun.length - 1]!;
        if (first.index > 0) { this.splitRunBefore(first); split = true; break; }
        if (last.index < last.owner.length - 1) { this.splitRunAfter(last); split = true; break; }
      }
      if (!split) return inside;
    }
    return this.segments().filter((s) => s.start >= start && s.end <= end && s.text.length > 0);
  }

  /** Moves the items before `seg` into a run of their own, in front of `seg`'s run. */
  private splitRunBefore(seg: TextSegment): void {
    const head = seg.owner.splice(0, seg.index);
    const first = runOf(head, copyRPr(seg.run.rPr));
    seg.runOwner.splice(seg.runIndex, 0, first);
    linkParents(first, (seg.run as { PARENT?: object }).PARENT ?? this.p);
  }

  /** Moves the items after `seg` into a run of their own, behind `seg`'s run. */
  private splitRunAfter(seg: TextSegment): void {
    const tail = seg.owner.splice(seg.index + 1);
    if (tail.length === 0) return;
    const rest = runOf(tail, copyRPr(seg.run.rPr));
    seg.runOwner.splice(seg.runIndex + 1, 0, rest);
    linkParents(rest, (seg.run as { PARENT?: object }).PARENT ?? this.p);
  }
}

type RevisionHolderOf = NonNullable<TextSegment['revision']>;

/** Where a `w:ins` goes when it must follow the `w:del` of a replacement. */
interface Anchor {
  owner: Element[];
  element: Element;
  parent: object;
}

export { W_NS };

// --- reading one w:pPr, direct or resolved -----------------------------------------------

function alignmentOf(pPr: wml.PPr | undefined, effective: boolean): AlignmentOrUnknown {
  switch (pPr?.jc?.val) {
    case 'left': case 'start': return 'Left';
    case 'center': return 'Centered';
    case 'right': case 'end': return 'Right';
    case 'both': case 'distribute': return 'Justified';
    // Word lays a paragraph with no w:jc out left-aligned, so that is the effective value
    default: return effective ? 'Left' : 'Unknown';
  }
}

function firstLineIndentOf(pPr: wml.PPr | undefined): number {
  const ind = pPr?.ind;
  if (ind?.hanging !== undefined) return -ind.hanging / TWIPS_PER_POINT;
  return (ind?.firstLine ?? 0) / TWIPS_PER_POINT;
}

function lineSpacingOf(pPr: wml.PPr | undefined): number {
  const spacing = pPr?.spacing;
  if (spacing?.line === undefined) return 0;
  return spacing.lineRule === undefined || spacing.lineRule === 'auto'
    ? (spacing.line / 240) * 12 : spacing.line / TWIPS_PER_POINT;
}

function outlineLevelOf(pPr: wml.PPr | undefined): number {
  const level = pPr?.outlineLvl?.val;
  return level === undefined ? 10 : level + 1;
}
