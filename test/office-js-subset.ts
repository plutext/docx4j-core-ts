// Compile-time check that the content API is assignable to a subset of Office JS's Word types,
// so that a function written against the subset runs against a live document and a package
// (CR-002 section 3.4). The interfaces are the Office JS members this package implements,
// copied from @types/office-js with load/sync/context removed and arrays for collections.
import type { Body, Paragraph, Range, Font, Table, TableRow, TableCell, InlinePicture, ContentControl, Comment, CustomXmlPart, CustomXmlNode, XmlMapping, TrackedChange, List, ListItem, WordprocessingMLPackage } from '../src/index.mjs';

namespace OfficeSubset {
  export type InsertLocation = 'Start' | 'End';
  export type ParagraphLocation = 'Before' | 'After';
  export interface SearchOptions { matchCase?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean }
  export type ChangeTrackingMode = 'Off' | 'TrackAll' | 'TrackMineOnly';
  export type ChangeTrackingState = 'Unknown' | 'Added' | 'Deleted' | 'Formatted' | 'None';
  export interface TrackedChange {
    readonly author: string;
    readonly date: Date | undefined;
    readonly text: string;
    readonly type: ChangeTrackingState;
    accept(): void;
    reject(): void;
    getRange(): Range | undefined;
  }
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
    getTrackedChanges(): ArrayLike<TrackedChange>;
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
    getTrackedChanges(): ArrayLike<TrackedChange>;
    // CR-002 phase H
    readonly isListItem: boolean;
    readonly list: List;
    readonly listItem: ListItem;
    readonly listOrNullObject: List;
    readonly listItemOrNullObject: ListItem;
    /** Office JS's is synchronous; here the default definitions are XML to unmarshal (section 17). */
    startNewList(): Promise<List>;
    attachToList(listId: number, level: number): List;
    detachFromList(): void;
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
    getTrackedChanges(): ArrayLike<TrackedChange>;
    // CR-002 phase C
    readonly tables: ArrayLike<Table>;
    readonly contentControls: ArrayLike<ContentControl>;
    readonly inlinePictures: ArrayLike<InlinePicture>;
    insertTable(rowCount: number, columnCount: number, location: InsertLocation, values?: string[][]): Table;
    insertInlinePictureFromBase64(base64EncodedImage: string, location: InsertLocation): InlinePicture;
    // CR-002 phase H
    readonly lists: ArrayLike<List>;
  }

  // CR-002 phase H: Word.List and Word.ListItem.
  export type ListLevelType = 'Bullet' | 'Number' | 'Picture';
  export type ListNumbering = 'None' | 'Arabic' | 'UpperRoman' | 'LowerRoman' | 'UpperLetter' | 'LowerLetter';
  export type ListBullet = 'Custom' | 'Solid' | 'Hollow' | 'Square' | 'Diamonds' | 'Arrow' | 'Checkmark';
  export interface List {
    readonly id: number;
    readonly levelExistences: boolean[];
    readonly levelTypes: ListLevelType[];
    readonly paragraphs: ArrayLike<Paragraph>;
    getLevelFont(level: number): Font;
    getLevelParagraphs(level: number): ArrayLike<Paragraph>;
    /** Office JS returns a ClientResult; nothing is marshalled here, so it is the string. */
    getLevelString(level: number): string;
    insertParagraph(paragraphText: string, insertLocation: InsertLocation): Paragraph;
    setLevelAlignment(level: number, alignment: string): void;
    setLevelBullet(level: number, listBullet: ListBullet, charCode?: number, fontName?: string): void;
    setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number): void;
    setLevelNumbering(level: number, listNumbering: ListNumbering, formatString?: (string | number)[]): void;
    setLevelStartingNumber(level: number, startingNumber: number): void;
  }
  export interface ListItem {
    level: number;
    readonly listString: string;
    readonly siblingIndex: number;
    getAncestor(parentOnly?: boolean): Paragraph;
    getDescendants(directChildrenOnly?: boolean): ArrayLike<Paragraph>;
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
    // CR-002 phase E (WordApiDesktop 1.3): the XML mapping, the w:sdtPr properties and the kinds
    readonly xmlMapping: XmlMapping;
    placeholderText: string;
    appearance: string;
    color: string;
    cannotDelete: boolean;
    cannotEdit: boolean;
    removeWhenEdited: boolean;
    readonly checkboxContentControl: CheckboxContentControl | undefined;
    readonly datePickerContentControl: DatePickerContentControl | undefined;
    readonly dropDownListContentControl: ListContentControl | undefined;
    readonly comboBoxContentControl: ListContentControl | undefined;
    readonly pictureContentControl: PictureContentControl | undefined;
    readonly repeatingSectionContentControl: RepeatingSectionContentControl | undefined;
    readonly groupContentControl: object | undefined;
  }

  // CR-002 phase E: Word.CustomXmlPart, Word.CustomXmlNode and Word.XmlMapping.
  export interface XmlMapping {
    readonly isMapped: boolean;
    readonly xpath: string;
    readonly prefixMappings: string;
    readonly customXmlPart: CustomXmlPart | undefined;
    readonly customXmlNode: CustomXmlNode | undefined;
    setMapping(xpath: string, prefixMappings?: string, part?: CustomXmlPart): boolean;
    setMappingByNode(node: CustomXmlNode): boolean;
    delete(): void;
  }
  export interface CustomXmlPart {
    readonly id: string;
    readonly namespaceUri: string;
    readonly builtIn: boolean;
    readonly documentElement: CustomXmlNode;
    readonly namespaceManager: CustomXmlPrefixMappingCollection;
    readonly schemaCollection: string[];
    /** Office JS returns a ClientResult; a DOM part does not go through a marshaller here, so it is the string (CR-002 section 12). */
    getXml(): string;
    setXml(xml: string): void;
    selectNodes(xpath: string, namespaceMappings?: string): ArrayLike<CustomXmlNode>;
    selectSingleNode(xpath: string, namespaceMappings?: string): CustomXmlNode | undefined;
    delete(): void;
    // insertElement / updateElement / deleteElement and the three attribute methods are deliberately
    // left out of the promise: Office JS's desktop-only forms put namespaceMappings second, and this
    // package puts it last and optional, as CR-002 section 3.5 specifies (section 12).
  }
  export interface CustomXmlPrefixMappingCollection {
    readonly items: ArrayLike<{ prefix: string; namespaceUri: string }>;
    addNamespace(prefix: string, namespaceUri: string): void;
    lookupNamespace(prefix: string): string;
    lookupPrefix(namespaceUri: string): string;
  }
  export interface CustomXmlNode {
    readonly baseName: string;
    readonly namespaceUri: string;
    readonly nodeType: string;
    nodeValue: string;
    text: string;
    readonly xml: string;
    readonly xpath: string;
    readonly attributes: ArrayLike<CustomXmlNode>;
    readonly childNodes: ArrayLike<CustomXmlNode>;
    readonly parentNode: CustomXmlNode | undefined;
    readonly firstChild: CustomXmlNode | undefined;
    readonly lastChild: CustomXmlNode | undefined;
    readonly nextSibling: CustomXmlNode | undefined;
    readonly previousSibling: CustomXmlNode | undefined;
    readonly ownerPart: CustomXmlPart;
    hasChildNodes(): boolean;
    selectNodes(xpath: string, namespaceMappings?: string): ArrayLike<CustomXmlNode>;
    selectSingleNode(xpath: string, namespaceMappings?: string): CustomXmlNode | undefined;
    insertNodeBefore(xml: string, nextSibling?: CustomXmlNode): CustomXmlNode;
    removeChild(child: CustomXmlNode): void;
    replaceChildNode(oldNode: CustomXmlNode, xml: string): CustomXmlNode;
    delete(): void;
  }
  export interface CheckboxContentControl {
    isChecked: boolean;
  }
  export interface DatePickerContentControl {
    dateDisplayFormat: string;
    dateDisplayLocale: string;
    dateCalendarType: string;
    dateStorageFormat: string;
  }
  export interface ContentControlListItem {
    displayText: string;
    value: string;
    readonly index: number;
    delete(): void;
  }
  export interface ListContentControl {
    readonly listItems: ArrayLike<ContentControlListItem>;
    addListItem(displayText: string, value?: string, index?: number): ContentControlListItem;
    deleteAllListItems(): void;
  }
  export interface PictureContentControl {
    readonly inlinePicture: InlinePicture | undefined;
  }
  export interface RepeatingSectionContentControl {
    readonly items: ArrayLike<ContentControl>;
    allowInsertDeleteSection: boolean;
    sectionTitle: string;
  }
  export interface CustomXmlPartCollection {
    readonly items: ArrayLike<CustomXmlPart>;
    getByNamespace(namespaceUri: string): ArrayLike<CustomXmlPart>;
    /** Office JS hands out a proxy that errors when there is none; here it is undefined. */
    getItem(id: string): CustomXmlPart | undefined;
    add(xml: string): CustomXmlPart;
  }

  /** Word.Document: the members the shim offers on `context.document` (CR-002 phases F and I). */
  export interface Document {
    changeTrackingMode: ChangeTrackingMode;
    readonly body: Body;
    getTrackedChanges(): ArrayLike<TrackedChange>;
  }
}

declare const pkg: WordprocessingMLPackage;
const body: OfficeSubset.Body = pkg.body;
const paragraph: OfficeSubset.Paragraph = body.insertParagraph('x', 'End');
const range: OfficeSubset.Range = paragraph.insertText('y', 'End');
const font: OfficeSubset.Font = range.font;
const document_: OfficeSubset.Document = pkg;
const change: OfficeSubset.TrackedChange | undefined = Array.from(body.getTrackedChanges())[0];

const table: OfficeSubset.Table = body.insertTable(2, 2, 'End', [['a', 'b'], ['c', 'd']]);
const row: OfficeSubset.TableRow = table.rows[0]!;
const tableCell: OfficeSubset.TableCell = table.getCell(0, 0);
const cellBody: OfficeSubset.Body = tableCell.body;
const picture: OfficeSubset.InlinePicture = body.insertInlinePictureFromBase64('', 'End');
declare const someControl: ContentControl;
const contentControl: OfficeSubset.ContentControl = someControl;
declare const commentView: Comment;
const comment: OfficeSubset.Comment = commentView;
declare const listView: List;
declare const listItemView: ListItem;
const list: OfficeSubset.List = listView;
const listItem: OfficeSubset.ListItem = listItemView;

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

// CR-002 phase E: the custom XML model and the typed content controls
declare const someCustomXmlPart: CustomXmlPart;
const customXmlPart: OfficeSubset.CustomXmlPart = someCustomXmlPart;
const customXmlNode: OfficeSubset.CustomXmlNode = someCustomXmlPart.documentElement;
const namespaceManager: OfficeSubset.CustomXmlPrefixMappingCollection = someCustomXmlPart.namespaceManager;
const xmlMapping: OfficeSubset.XmlMapping = someControl.xmlMapping;
const customXmlParts: OfficeSubset.CustomXmlPartCollection = pkg.customXmlParts;
const listContentControl: OfficeSubset.ListContentControl | undefined = someControl.dropDownListContentControl;

// one written against the phase E members: it runs against a Word add-in's objects too
function bindToFirstNode(control: OfficeSubset.ContentControl, part: OfficeSubset.CustomXmlPart): boolean {
  const node = part.selectSingleNode('/*[1]/*[1]');
  if (!node) return false;
  return control.xmlMapping.setMappingByNode(node);
}
void bindToFirstNode(someControl, someCustomXmlPart);

export { customXmlPart, customXmlNode, namespaceManager, xmlMapping, customXmlParts, listContentControl, bindToFirstNode };

// CR-002 phase H: one written against the list members, which runs in an add-in too
async function outlineList(b: OfficeSubset.Body): Promise<string[]> {
  const out: string[] = [];
  for (const p of Array.from(b.paragraphs)) {
    if (!p.isListItem) continue;
    const item = p.listItem;
    out.push(`${'  '.repeat(item.level)}${item.listString} ${p.text}`);
  }
  const first = Array.from(b.paragraphs)[0]!;
  const made = await first.startNewList();
  made.setLevelNumbering(0, 'UpperRoman');
  made.setLevelIndents(0, 36, 18);
  Array.from(b.lists).forEach((l) => out.push(String(l.id) + ' ' + l.levelTypes[0]));
  return out;
}
void outlineList(pkg.body);

// and one that reviews, as an add-in would (CR-002 phase F)
function review(d: OfficeSubset.Document): string[] {
  d.changeTrackingMode = 'TrackAll';
  return Array.from(d.body.getTrackedChanges()).map((c) => `${c.type} by ${c.author}: ${c.text}`);
}
review(pkg);

export { body, paragraph, range, font, table, row, tableCell, cellBody, picture, contentControl, comment, list, listItem, document_, change, shout, fill, resolveComments, review, outlineList };
export type { Body, Paragraph, Range, Font, Table, TableRow, TableCell, InlinePicture, ContentControl, Comment, CustomXmlPart, CustomXmlNode, XmlMapping, TrackedChange, List, ListItem };
