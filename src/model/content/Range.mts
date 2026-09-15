import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { Font } from './Font.mjs';
import type { Paragraph } from './Paragraph.mjs';
import type { BlockElement } from './Body.mjs';
import { type Element, typeNameOf } from './tree.mjs';
import { contentOf } from './ooxml.mjs';
import { searchPattern, findAll, type SearchOptions } from './search.mjs';
import { commentApi } from './comments.mjs';
import type { Comment } from './Comment.mjs';

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

  /** Direct formatting of exactly this span: runs are split at the boundaries so that a write touches only the span. */
  get font(): Font {
    return new Font(() => {
      if (this.start === this.end) return [];
      this.paragraph.splitAt(this.start);
      this.paragraph.splitAt(this.end);
      return this.runs.map((r) => r.value);
    });
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
}
