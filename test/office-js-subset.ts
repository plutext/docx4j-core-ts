// Compile-time check that the content API is assignable to a subset of Office JS's Word types,
// so that a function written against the subset runs against a live document and a package
// (CR-002 section 3.4). The interfaces are the Office JS members this package implements,
// copied from @types/office-js with load/sync/context removed and arrays for collections.
import type { Body, Paragraph, Range, Font, Table, TableRow, TableCell, InlinePicture, ContentControl, Comment, WordprocessingMLPackage } from '../src/index.mjs';

namespace OfficeSubset {
  export type InsertLocation = 'Start' | 'End';
  export type ParagraphLocation = 'Before' | 'After';
  export interface SearchOptions { matchCase?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean }
  export interface Font {
    bold: boolean; italic: boolean; underline: string; strikeThrough: boolean; subscript: boolean; superscript: boolean;
    name: string; size: number; color: string; highlightColor: string | null;
  }
  /**
   * Word.Comment, minus contentRange; Office JS has a separate CommentReply type, and this subset
   * uses Comment for replies too (CR-002 section 9). Promises where this package unmarshals the
   * comment parts.
   */
  export interface Comment {
    readonly authorEmail: string;
    readonly authorName: string;
    content: string;
    readonly creationDate: Date | undefined;
    resolved: boolean;
    readonly replies: ArrayLike<Comment>;
    reply(replyText: string): Promise<Comment>;
    delete(): Promise<void>;
  }
  export interface Range {
    text: string; readonly font: Font; style: string;
    insertText(text: string, location: 'Before' | 'After' | 'Start' | 'End' | 'Replace'): Range;
    insertParagraph(text: string, location: ParagraphLocation): Paragraph;
    search(text: string, options?: SearchOptions): ArrayLike<Range>;
    delete(): void;
    getRange(location?: 'Whole' | 'Start' | 'End' | 'Content'): Range;
    getComments(): Promise<ArrayLike<Comment>>;
    insertComment(content: string): Promise<Comment>;
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
    getComments(): Promise<ArrayLike<Comment>>;
    insertComment(content: string): Promise<Comment>;
  }
  export interface Body {
    readonly paragraphs: ArrayLike<Paragraph>;
    readonly text: string;
    insertParagraph(text: string, location: InsertLocation): Paragraph;
    insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range;
    insertBreak(type: 'Page' | 'Line', location: InsertLocation): void;
    search(text: string, options?: SearchOptions): ArrayLike<Range>;
    clear(): void;
    getComments(): Promise<ArrayLike<Comment>>;
    // CR-002 phase C
    readonly tables: ArrayLike<Table>;
    readonly contentControls: ArrayLike<ContentControl>;
    readonly inlinePictures: ArrayLike<InlinePicture>;
    insertTable(rowCount: number, columnCount: number, location: InsertLocation, values?: string[][]): Table;
    insertInlinePictureFromBase64(base64EncodedImage: string, location: InsertLocation): InlinePicture;
  }

  // CR-002 phase C: Word.Table, Word.TableRow, Word.TableCell, Word.InlinePicture, Word.ContentControl.
  export interface Table {
    readonly rowCount: number;
    readonly rows: ArrayLike<TableRow>;
    values: string[][];
    style: string;
    styleBuiltIn: string;
    headerRowCount: number;
    getCell(rowIndex: number, cellIndex: number): TableCell;
    addRows(insertLocation: InsertLocation, rowCount: number, values?: string[][]): ArrayLike<TableRow>;
    deleteRows(rowIndex: number, rowCount?: number): void;
    delete(): void;
  }
  export interface TableRow {
    readonly cellCount: number;
    readonly cells: ArrayLike<TableCell>;
    readonly rowIndex: number;
    readonly parentTable: Table;
    values: string[];
    isHeader: boolean;
    insertRows(insertLocation: ParagraphLocation, rowCount: number, values?: string[][]): ArrayLike<TableRow>;
    delete(): void;
  }
  export interface TableCell {
    readonly body: Body;
    readonly cellIndex: number;
    readonly rowIndex: number;
    readonly parentRow: TableRow;
    readonly parentTable: Table;
    value: string;
    /** Points, as Office JS reports them (CR-002 open question 2). */
    width: number;
    columnWidth: number;
  }
  export interface InlinePicture {
    altTextDescription: string;
    altTextTitle: string;
    width: number;
    height: number;
    readonly imageFormat: string;
    readonly paragraph: Paragraph;
    // Office JS returns a ClientResult; this package marshals, so it is a promise (CR-002 section 3.4)
    getBase64ImageSrc(): Promise<string>;
    delete(): void;
  }
  export interface ContentControl {
    tag: string;
    title: string;
    readonly id: number;
    readonly type: string;
    readonly text: string;
    readonly paragraphs: ArrayLike<Paragraph>;
    readonly tables: ArrayLike<Table>;
    readonly contentControls: ArrayLike<ContentControl>;
    insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range;
    insertParagraph(text: string, location: 'Start' | 'End' | 'Before' | 'After'): Paragraph;
    search(text: string, options?: SearchOptions): ArrayLike<Range>;
    getRange(location?: 'Whole' | 'Start' | 'End' | 'Content'): Range;
    delete(keepContent: boolean): void;
  }
}

declare const pkg: WordprocessingMLPackage;
const body: OfficeSubset.Body = pkg.body;
const paragraph: OfficeSubset.Paragraph = body.insertParagraph('x', 'End');
const range: OfficeSubset.Range = paragraph.insertText('y', 'End');
const font: OfficeSubset.Font = range.font;

const table: OfficeSubset.Table = body.insertTable(2, 2, 'End', [['a', 'b'], ['c', 'd']]);
const row: OfficeSubset.TableRow = table.rows[0]!;
const tableCell: OfficeSubset.TableCell = table.getCell(0, 0);
const cellBody: OfficeSubset.Body = tableCell.body;
const picture: OfficeSubset.InlinePicture = body.insertInlinePictureFromBase64('', 'End');
declare const someControl: ContentControl;
const contentControl: OfficeSubset.ContentControl = someControl;
declare const commentView: Comment;
const comment: OfficeSubset.Comment = commentView;

// a function written against the subset
function shout(b: OfficeSubset.Body): void {
  for (const r of Array.from(b.search('hello'))) { r.font.bold = true; r.insertText('!', 'After'); }
}
shout(pkg.body);

// one written against the phase C members
function fill(b: OfficeSubset.Body, values: string[][]): OfficeSubset.Table {
  const t = b.insertTable(values.length, values[0]!.length, 'End', values);
  t.headerRowCount = 1;
  t.styleBuiltIn = 'TableGrid';
  return t;
}
fill(pkg.body, [['a']]);

// comments, the same way: this compiles against a Word add-in's objects too
async function resolveComments(b: OfficeSubset.Body): Promise<void> {
  for (const c of Array.from(await b.getComments())) {
    if (c.content.startsWith('TODO')) await c.reply('done');
    c.resolved = true;
  }
}
void resolveComments(pkg.body);

export { body, paragraph, range, font, table, row, tableCell, cellBody, picture, contentControl, comment, shout, fill, resolveComments };
export type { Body, Paragraph, Range, Font, Table, TableRow, TableCell, InlinePicture, ContentControl, Comment };
