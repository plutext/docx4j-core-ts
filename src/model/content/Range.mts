import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { Font } from './Font.mjs';
import { runFontSelectorOf } from '../fonts/lookup.mjs';
import type { Paragraph, FormattingOptions } from './Paragraph.mjs';
import type { BlockElement } from './Body.mjs';
import { type Element, typeNameOf, linkParents, runItemsOf, type TextViewOptions } from './tree.mjs';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { ContentControl, type ContentControlType } from './ContentControl.mjs';
import { sdt as sdtOf, nextSdtId } from '@docx4j/generated-objects-ts/builders/wml';
import * as f from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { hyperlink as hyperlinkOf } from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { Namespaces } from '../../parts/Namespaces.mjs';
import { controlIdScope, sdtKindFor } from '../customxml/insert.mjs';
import { contentOf } from './ooxml.mjs';
import { searchPattern, findAll, type SearchOptions } from './search.mjs';
import { commentApi } from './comments.mjs';
import type { Comment } from './Comment.mjs';
import type { TrackedChange } from './TrackedChange.mjs';

/**
 * A subset of Office JS `Word.Range`: a span of text within one paragraph, [start, end) in the
 * paragraph's text. Positions are recomputed from the tree on every use, so a range stays
 * valid across edits made through it; edits made elsewhere may shift it.
 */
export class Range {
  constructor(readonly paragraph: Paragraph, public start: number, public end: number) {}

  get paragraphs(): Paragraph[] {
    return [this.paragraph];
  }

  get text(): string {
    return this.paragraph.text.substring(this.start, this.end);
  }
  set text(value: string) {
    this.insertText(value, 'Replace');
  }

  /** The paragraph's style id (extension, as on Paragraph). */
  get styleId(): string {
    return this.paragraph.styleId;
  }
  set styleId(id: string) {
    this.paragraph.styleId = id;
  }

  /** The paragraph's style name (Office JS reports the range's paragraph style). */
  get style(): string {
    return this.paragraph.style;
  }
  set style(id: string) {
    this.paragraph.style = id;
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
    const holders = (): wml.R[] => {
      if (this.start === this.end) return [];
      this.paragraph.splitAt(this.start);
      this.paragraph.splitAt(this.end);
      return this.runs.map((r) => r.value);
    };
    if (options?.direct === true) return new Font(holders, () => this.paragraph.fontTracking());
    return new Font(holders, () => this.paragraph.fontTracking(),
      (rPr) => this.paragraph.parentBody.propertyResolver.getEffectiveRPr(rPr, this.paragraph.p.pPr),
      (rPr) => runFontSelectorOf(this.paragraph.parentBody.package_)?.asciiFontName(rPr));
  }

  /** The runs the span covers (a run partly inside counts). */
  get runs(): Element<wml.R>[] {
    const out: Element<wml.R>[] = [];
    const seen = new Set<wml.R>();
    for (const seg of this.paragraph.segments()) {
      if (seg.end <= this.start || seg.start >= this.end) continue;
      if (seen.has(seg.run)) continue;
      seen.add(seg.run);
      out.push(seg.runOwner[seg.runIndex] as Element<wml.R>);
    }
    return out;
  }

  insertText(text: string, location: 'Before' | 'After' | 'Start' | 'End' | 'Replace'): Range {
    let r: Range;
    if (location === 'Replace') r = this.paragraph.splice(this.start, this.end, text);
    else if (location === 'Before' || location === 'Start') r = this.paragraph.splice(this.start, this.start, text);
    else r = this.paragraph.splice(this.end, this.end, text);
    // this range keeps covering the same text: a replacement resizes it, text before it shifts it, text after it leaves it
    if (location === 'Replace') this.end = this.start + text.length;
    else if (location === 'Before' || location === 'Start') { this.start += text.length; this.end += text.length; }
    return r;
  }

  insertParagraph(text: string, location: 'Before' | 'After'): Paragraph {
    return this.paragraph.insertParagraph(text, location);
  }

  /**
   * Word's `insertOoxml` at range level: a flat OPC `pkg:package` (or a bare fragment). One
   * incoming paragraph's runs go at the start or the end of the span, or in its place; anything
   * else is inserted as blocks before or after the paragraph. Returns what was inserted, as
   * `insertXml` does.
   */
  async insertOoxml(ooxml: string, location: 'Before' | 'After' | 'Replace'): Promise<(Paragraph | BlockElement)[]> {
    const body = this.paragraph.parentBody;
    const preprocess = body.package_?.loadOptions.preprocessor;
    const elements = await contentOf(ooxml, { preprocess: preprocess ? (doc) => preprocess(doc) : undefined, target: body.part });
    if (elements.length === 0) return [];
    const only = elements.length === 1 && typeNameOf(elements[0]!) === 'org_docx4j_wml.P' ? elements[0]!.value as { content?: Element[] } : undefined;
    if (only) {
      if (location === 'Replace') { this.delete(); this.paragraph.insertItemsAt(this.start, (only.content ?? [])); }
      else this.paragraph.insertItemsAt(location === 'Before' ? this.start : this.end, (only.content ?? []));
      return [this.paragraph];
    }
    body.insertElement(elements, location === 'Before' ? 'Before' : 'After', this.paragraph);
    if (location === 'Replace') this.delete();
    return elements.map((el) => typeNameOf(el) === 'org_docx4j_wml.P'
      ? body.paragraphFor(el.value as never) ?? { element: el, container: this.paragraph.container }
      : { element: el, container: this.paragraph.container });
  }

  /** Removes the span's text. */
  delete(): void {
    this.paragraph.splice(this.start, this.end, '');
    this.end = this.start;
  }

  search(text: string, options?: SearchOptions): Range[] {
    const base = this.start;
    return findAll(this.text, searchPattern(text, options)).map(([s, e]) => new Range(this.paragraph, base + s, base + e));
  }

  /** Replaces every match in this span, last first so the offsets stay valid; returns the count (CR-002 section 3.7). */
  replaceText(find: string, replace: string, options?: SearchOptions): number {
    const matches = this.search(find, options);
    for (let i = matches.length - 1; i >= 0; i--) matches[i]!.insertText(replace, 'Replace');
    return matches.length;
  }

  /** The tracked changes this span covers, in document order (CR-002 phase F). */
  getTrackedChanges(): TrackedChange[] {
    return this.paragraph.getTrackedChanges().filter((change) => {
      const range = change.getRange();
      if (!range) return false;
      return range.start <= this.end && range.end >= this.start;
    });
  }

  getRange(location: 'Whole' | 'Start' | 'End' | 'Content' = 'Whole'): Range {
    if (location === 'Start') return new Range(this.paragraph, this.start, this.start);
    if (location === 'End') return new Range(this.paragraph, this.end, this.end);
    return new Range(this.paragraph, this.start, this.end);
  }

  /** The span's text as it stands (extension; Office JS getOoxml wraps it in a package). */
  toString(): string {
    return this.text;
  }

  // --- comments (CR-002 phase G) ---

  /** The comments this span touches, replies nested under their parent. */
  async getComments(): Promise<Comment[]> {
    return commentApi().commentsOf(this);
  }

  /**
   * Comments the span: the markers and the reference run around it (runs are split at the
   * boundaries, as `font` does), the comment itself and the side-part entries, creating any of
   * the comment parts the document lacks. Asynchronous because it unmarshals those parts.
   */
  async insertComment(text: string): Promise<Comment> {
    return commentApi().insertComment(this, text);
  }

  // --- content controls (CR-002 phase E) ---

  /**
   * Wraps the runs this span covers in a run-level content control (Office JS
   * `Range.insertContentControl`), splitting the runs at the boundaries as `font` does. An empty
   * span gets an empty control at its position. Returns the control.
   */
  insertContentControl(kind?: ContentControlType): ContentControl {
    const paragraph = this.paragraph;
    const body = paragraph.parentBody;
    // The control is built empty first: `sdt` refuses a kind that cannot be run-level (a repeating
    // section), and that refusal must come before any run is split.
    const sdt = sdtOf([], { kind: sdtKindFor(kind), id: nextSdtId(controlIdScope(body)), form: 'run' }) as Element<wml.SdtRun>;
    if (this.start === this.end) {
      paragraph.insertItemsAt(this.start, [sdt as Element]);
      return new ContentControl(sdt, containerOf(sdt, paragraph), body);
    }
    paragraph.splitAt(this.start);
    paragraph.splitAt(this.end);
    const segments = paragraph.segments().filter((s) => s.start >= this.start && s.end <= this.end);
    if (segments.length === 0) throw new Docx4JException('This range covers no run');
    const owner = segments[0]!.runOwner;
    if (segments.some((s) => s.runOwner !== owner)) {
      throw new Docx4JException('This range spans more than one run holder (a hyperlink or a tracked change); wrap a narrower span');
    }
    const items: Element[] = [];
    for (const segment of segments) {
      const element = owner[segment.runIndex] as Element;
      if (!items.includes(element)) items.push(element);
    }
    const at = owner.indexOf(items[0]!);
    (sdt.value.sdtContent as { content?: Element[] }).content = items;
    owner.splice(at, items.length, sdt as Element);
    linkParents(sdt, (segments[0]!.run as { PARENT?: object }).PARENT ?? paragraph.p);
    return new ContentControl(sdt, owner, body);
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
   * this range's part, which is why the part must exist (CR-002 section 20).
   */
  get hyperlink(): string {
    const holder = this.hyperlinkHolder();
    if (holder === undefined) return '';
    const anchor = holder.value.anchor ?? '';
    const address = holder.value.id === undefined ? '' : this.relationshipTarget(holder.value.id);
    return anchor === '' ? address : `${address}#${anchor}`;
  }

  set hyperlink(value: string) {
    const paragraph = this.paragraph;
    this.removeHyperlinks();
    if (value === '') return;

    const hash = value.indexOf('#');
    const address = hash === -1 ? value : value.slice(0, hash);
    const location = hash === -1 ? '' : value.slice(hash + 1);
    const holder = hyperlinkOf({ content: [] }) as Element<wml.P.Hyperlink>;
    if (location !== '') holder.value.anchor = location;
    if (address !== '') holder.value.id = this.addHyperlinkRelationship(address);

    if (this.start === this.end) {
      paragraph.insertItemsAt(this.start, [holder as Element]);
      return;
    }
    paragraph.splitAt(this.start);
    paragraph.splitAt(this.end);
    const segments = paragraph.segments().filter((seg) => seg.start >= this.start && seg.end <= this.end);
    if (segments.length === 0) throw new Docx4JException('This range covers no run');
    const owner = segments[0]!.runOwner;
    if (segments.some((seg) => seg.runOwner !== owner)) {
      throw new Docx4JException('This range spans more than one run holder (a hyperlink or a tracked change); wrap a narrower span');
    }
    const items: Element[] = [];
    for (const segment of segments) {
      const element = owner[segment.runIndex] as Element;
      if (!items.includes(element)) items.push(element);
    }
    const at = owner.indexOf(items[0]!);
    holder.value.content = items as never;
    owner.splice(at, items.length, holder as Element);
    linkParents(holder, (segments[0]!.run as { PARENT?: object }).PARENT ?? paragraph.p);
  }

  /** The `w:hyperlink` this range's first run sits in, if any. */
  private hyperlinkHolder(): Element<wml.P.Hyperlink> | undefined {
    for (const segment of this.paragraph.segments()) {
      if (segment.end <= this.start && this.start !== this.end) continue;
      if (segment.start >= this.end && this.start !== this.end) break;
      const parent = (segment.run as { PARENT?: object }).PARENT;
      const found = hyperlinkHolders(this.paragraph).find((h) => h.value === parent);
      if (found) return found;
    }
    return undefined;
  }

  /** Unwraps every `w:hyperlink` the range touches, keeping its runs where they are. */
  private removeHyperlinks(): void {
    for (const holder of hyperlinkHolders(this.paragraph)) {
      const owner = containerOf(holder as Element, this.paragraph);
      const at = owner.indexOf(holder as Element);
      if (at === -1) continue;
      owner.splice(at, 1, ...((holder.value.content ?? []) as Element[]));
      for (const item of (holder.value.content ?? []) as Element[]) {
        linkParents(item, (holder as { PARENT?: object }).PARENT ?? this.paragraph.p);
      }
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
