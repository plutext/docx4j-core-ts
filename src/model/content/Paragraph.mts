import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy, marshalString } from '@docx4j/generated-objects-ts';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { r as textRun, t as textItem, br as breakItem } from '@docx4j/generated-objects-ts/builders/wml';
import { type Element, segmentsOf, runsOf, runItemsOf, linkParents, textOf, typeNameOf, W_NS, runOf, paragraphOf, type TextSegment } from './tree.mjs';

function withRPr(run: Element<wml.R>, rPr: wml.RPr | undefined): Element<wml.R> {
  if (rPr) run.value.rPr = rPr;
  return run;
}
import { Font } from './Font.mjs';
import { Range } from './Range.mjs';
import { searchPattern, findAll, type SearchOptions } from './search.mjs';
import type { Body, BlockElement } from './Body.mjs';
import { builtInOf, idOfBuiltIn, styleNameOf, styleIdOf } from './styles.mjs';
import { cellOf, type TableCell } from './Table.mjs';
import { type ContentControl, collectRunControls } from './ContentControl.mjs';
import { InlinePicture, addImage, writableWidthEmu, type InlinePictureOptions } from './InlinePicture.mjs';
import { contentOf } from './ooxml.mjs';
import { commentApi } from './comments.mjs';
import type { Comment } from './Comment.mjs';

/** Office JS Word.Alignment. */
export type Alignment = 'Unknown' | 'Left' | 'Centered' | 'Right' | 'Justified';

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
    const rPr = runsOf(this.p)[0]?.value.rPr;
    this.p.content = [withRPr(textRun(value), rPr ? deepCopy(rPr) : undefined)] as wml.P['content'];
    linkParents(this.p.content, this.p);
  }

  /** The style id (w:pStyle; docx4j's name for it); 'Normal' when none is set. Extension: Office JS has `style` and `styleBuiltIn` only. */
  get styleId(): string {
    return this.p.pPr?.pStyle?.val ?? 'Normal';
  }
  set styleId(id: string) {
    if (id === '' || id === 'Normal') { if (this.p.pPr) delete this.p.pPr.pStyle; return; }
    this.pPr().pStyle = { val: id };
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

  get alignment(): Alignment {
    switch (this.p.pPr?.jc?.val) {
      case 'left': case 'start': return 'Left';
      case 'center': return 'Centered';
      case 'right': case 'end': return 'Right';
      case 'both': case 'distribute': return 'Justified';
      default: return 'Unknown';
    }
  }
  set alignment(v: Alignment) {
    const val: wml.JcEnumeration | undefined = v === 'Left' ? 'left' : v === 'Centered' ? 'center' : v === 'Right' ? 'right' : v === 'Justified' ? 'both' : undefined;
    if (val === undefined) { if (this.p.pPr) delete this.p.pPr.jc; return; }
    this.pPr().jc = { val };
  }

  /** Indents and spacing in points, as Office JS; 0 when not set directly. */
  get leftIndent(): number { return (this.p.pPr?.ind?.left ?? 0) / TWIPS_PER_POINT; }
  set leftIndent(pt: number) { this.ind().left = Math.round(pt * TWIPS_PER_POINT); }
  get rightIndent(): number { return (this.p.pPr?.ind?.right ?? 0) / TWIPS_PER_POINT; }
  set rightIndent(pt: number) { this.ind().right = Math.round(pt * TWIPS_PER_POINT); }
  /** First-line indent; negative for a hanging indent, as Office JS. */
  get firstLineIndent(): number {
    const ind = this.p.pPr?.ind;
    if (ind?.hanging !== undefined) return -ind.hanging / TWIPS_PER_POINT;
    return (ind?.firstLine ?? 0) / TWIPS_PER_POINT;
  }
  set firstLineIndent(pt: number) {
    const ind = this.ind();
    if (pt < 0) { ind.hanging = Math.round(-pt * TWIPS_PER_POINT); delete ind.firstLine; } else { ind.firstLine = Math.round(pt * TWIPS_PER_POINT); delete ind.hanging; }
  }
  get spaceBefore(): number { return (this.p.pPr?.spacing?.before ?? 0) / TWIPS_PER_POINT; }
  set spaceBefore(pt: number) { this.spacing().before = Math.round(pt * TWIPS_PER_POINT); }
  get spaceAfter(): number { return (this.p.pPr?.spacing?.after ?? 0) / TWIPS_PER_POINT; }
  set spaceAfter(pt: number) { this.spacing().after = Math.round(pt * TWIPS_PER_POINT); }
  /** Line spacing in points: w:line/240 lines of 12pt when the rule is auto, else w:line twips; 0 when not set. */
  get lineSpacing(): number {
    const s = this.p.pPr?.spacing;
    if (s?.line === undefined) return 0;
    return s.lineRule === undefined || s.lineRule === 'auto' ? (s.line / 240) * 12 : s.line / TWIPS_PER_POINT;
  }
  set lineSpacing(pt: number) { const s = this.spacing(); s.line = Math.round(pt * TWIPS_PER_POINT); s.lineRule = 'exact'; }
  /** w:outlineLvl + 1 (1 to 9); 10 for body text, as Office JS. */
  get outlineLevel(): number { const l = this.p.pPr?.outlineLvl?.val; return l === undefined ? 10 : l + 1; }
  set outlineLevel(level: number) { if (level >= 10 || level < 1) { if (this.p.pPr) delete this.p.pPr.outlineLvl; } else this.pPr().outlineLvl = { val: level - 1 }; }

  /** Direct formatting of the runs: reads the first run, writes all (extension: `Font` is over runs, not the paragraph mark). */
  get font(): Font {
    return new Font(() => runsOf(this.p).map((r) => r.value));
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
    if (pPr) { delete pPr.rPr; delete (pPr as { sectPr?: unknown }).sectPr; }
    const el = paragraphOf(text === '' ? [] : [textRun(text)], pPr);
    return this.parentBody.insertElement(el, location, this) as Paragraph;
  }

  insertBreak(type: 'Page' | 'Line', location: 'Before' | 'After' | 'Start' | 'End'): void {
    const brk = breakItem(type === 'Page' ? 'page' : undefined);
    if (location === 'Before' || location === 'After') {
      const p = this.insertParagraph('', location);
      p.p.content!.push(runOf([brk]) as never);
      linkParents(p.p.content, p.p);
      return;
    }
    const run = runOf([brk]);
    const content = (this.p.content ??= []);
    if (location === 'Start') content.unshift(run as never); else content.push(run as never);
    linkParents(run, this.p);
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

  getRange(location: 'Whole' | 'Start' | 'End' | 'Content' = 'Whole'): Range {
    const length = this.text.length;
    if (location === 'Start') return new Range(this, 0, 0);
    if (location === 'End') return new Range(this, length, length);
    return new Range(this, 0, length);
  }

  /** Removes the paragraph from its container. */
  delete(): void {
    const i = this.index;
    if (i >= 0) this.container.splice(i, 1);
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

  private pPr(): wml.PPr {
    return (this.p.pPr ??= { TYPE_NAME: 'org_docx4j_wml.PPr' });
  }
  private ind(): wml.PPrBase.Ind {
    return (this.pPr().ind ??= {});
  }
  private spacing(): wml.PPrBase.Spacing {
    return (this.pPr().spacing ??= {});
  }

  /** The text segments (w:t, w:tab, ...) with their offsets. */
  segments(): TextSegment[] {
    return segmentsOf(this.p);
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

  private removeEmptyRuns(): void {
    const prune = (items: Element[]): void => {
      for (let i = items.length - 1; i >= 0; i--) {
        const el = items[i]!;
        const v = el.value as { TYPE_NAME?: string; content?: Element[]; sdtContent?: { content?: Element[] } };
        if (v.TYPE_NAME === 'org_docx4j_wml.R') {
          if (!v.content || v.content.length === 0) items.splice(i, 1);
        } else if (v.TYPE_NAME !== 'org_docx4j_wml.P') {
          const nested = runItemsOf(v);
          if (nested) prune(nested);
        }
      }
    };
    prune(this.p.content ?? []);
  }
}

export { W_NS };
