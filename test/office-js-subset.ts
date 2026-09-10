// Compile-time check that the content API is assignable to a subset of Office JS's Word types,
// so that a function written against the subset runs against a live document and a package
// (CR-002 section 3.4). The interfaces are the Office JS members this package implements,
// copied from @types/office-js with load/sync/context removed and arrays for collections.
import type { Body, Paragraph, Range, Font, WordprocessingMLPackage } from '../src/index.mjs';

namespace OfficeSubset {
  export type InsertLocation = 'Start' | 'End';
  export type ParagraphLocation = 'Before' | 'After';
  export interface SearchOptions { matchCase?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean }
  export interface Font {
    bold: boolean; italic: boolean; underline: string; strikeThrough: boolean; subscript: boolean; superscript: boolean;
    name: string; size: number; color: string; highlightColor: string | null;
  }
  export interface Range {
    text: string; readonly font: Font; style: string;
    insertText(text: string, location: 'Before' | 'After' | 'Start' | 'End' | 'Replace'): Range;
    insertParagraph(text: string, location: ParagraphLocation): Paragraph;
    search(text: string, options?: SearchOptions): ArrayLike<Range>;
    delete(): void;
    getRange(location?: 'Whole' | 'Start' | 'End' | 'Content'): Range;
  }
  export interface Paragraph {
    text: string; style: string; styleBuiltIn: string; alignment: string;
    leftIndent: number; rightIndent: number; firstLineIndent: number; spaceBefore: number; spaceAfter: number; lineSpacing: number; outlineLevel: number;
    readonly font: Font;
    insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range;
    insertParagraph(text: string, location: ParagraphLocation): Paragraph;
    insertBreak(type: 'Page' | 'Line', location: 'Before' | 'After' | 'Start' | 'End'): void;
    search(text: string, options?: SearchOptions): ArrayLike<Range>;
    getRange(location?: 'Whole' | 'Start' | 'End' | 'Content'): Range;
    delete(): void;
  }
  export interface Body {
    readonly paragraphs: ArrayLike<Paragraph>;
    readonly text: string;
    insertParagraph(text: string, location: InsertLocation): Paragraph;
    insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range;
    insertBreak(type: 'Page' | 'Line', location: InsertLocation): void;
    search(text: string, options?: SearchOptions): ArrayLike<Range>;
    clear(): void;
  }
}

declare const pkg: WordprocessingMLPackage;
const body: OfficeSubset.Body = pkg.body;
const paragraph: OfficeSubset.Paragraph = body.insertParagraph('x', 'End');
const range: OfficeSubset.Range = paragraph.insertText('y', 'End');
const font: OfficeSubset.Font = range.font;

// a function written against the subset
function shout(b: OfficeSubset.Body): void {
  for (const r of Array.from(b.search('hello'))) { r.font.bold = true; r.insertText('!', 'After'); }
}
shout(pkg.body);
export { body, paragraph, range, font, shout };
export type { Body, Paragraph, Range, Font };
