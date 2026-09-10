import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { Font } from './Font.mjs';
import type { Paragraph } from './Paragraph.mjs';
import type { Element } from './tree.mjs';
import { searchPattern, findAll, type SearchOptions } from './search.mjs';

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

  /** The paragraph's style (Office JS reports the range's paragraph style). */
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
}
