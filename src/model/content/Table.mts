// CR-002 phase C: Table, TableRow and TableCell, subsets of Office JS's Word.Table and friends
// over w:tbl, w:tr and w:tc. Views, as Paragraph is: the element plus the array holding it,
// nothing cached, PARENT linked on everything inserted.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { p as paragraphOfText } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { type Element, typeNameOf, childrenOf, linkParents, textOf, rowsOf, cellsOf } from './tree.mjs';
import type { Body } from './Body.mjs';
import { builtInOf, idOfBuiltIn, styleNameOf, styleIdOf } from './styles.mjs';
import type { Paragraph } from './Paragraph.mjs';
import type { Range } from './Range.mjs';
import { type ChangeTracker, trackInsertedParagraph } from './tracking.mjs';

/** The default table width in twips (A4 with 2.54 cm margins), as the objects package's `tbl` builder uses. */
const DEFAULT_WIDTH = 9026;
const TWIPS_PER_POINT = 20;

/**
 * A subset of Office JS `Word.Table` over a `w:tbl`. Rows are read through row-level content
 * controls (an OpenDoPE repeat wraps its `w:tr` in a `w:sdt`), so `rows` is what Word shows.
 */
export class Table {
  constructor(
    /** The `w:tbl` element pair. */
    readonly element: Element<wml.Tbl>,
    /** The content array holding the element. */
    readonly container: Element[],
    /** The body the table is in (the part's), for addresses and paragraph ids. */
    readonly parentBody: Body,
  ) {}

  /** The table value (docx4j: the Tbl). */
  get tbl(): wml.Tbl {
    return this.element.value;
  }

  get index(): number {
    return this.container.indexOf(this.element);
  }

  get rowCount(): number {
    return rowsOf(this.tbl).length;
  }

  get rows(): TableRow[] {
    return rowsOf(this.tbl).map((r) => new TableRow(r.element, r.container, this));
  }

  /** The text of every cell, row by row (Office JS values). */
  get values(): string[][] {
    return this.rows.map((row) => row.values);
  }
  set values(values: string[][]) {
    const rows = this.rows;
    values.forEach((row, i) => { if (i < rows.length) rows[i]!.values = row; });
  }

  /** The table style id (w:tblStyle; docx4j's name for it); '' when none is set. Extension, as on Paragraph. */
  get styleId(): string {
    return this.tbl.tblPr?.tblStyle?.val ?? '';
  }
  set styleId(id: string) {
    if (id === '') { if (this.tbl.tblPr) delete this.tbl.tblPr.tblStyle; return; }
    this.tblPr().tblStyle = { val: id };
  }

  /** The style's display name ('Table Grid'), as Office JS; '' when none is set. Setting accepts a name or an id, as Paragraph. */
  get style(): string {
    const id = this.styleId;
    return id === '' ? '' : styleNameOf(this.parentBody.package_, id);
  }
  set style(name: string) {
    this.styleId = name === '' ? '' : styleIdOf(this.parentBody.package_, name);
  }

  /** The `Word.Style` value ('TableGrid'), or 'Other', as Office JS. */
  get styleBuiltIn(): string {
    return builtInOf(this.styleId);
  }
  set styleBuiltIn(value: string) {
    this.styleId = idOfBuiltIn(value);
  }

  /** The number of leading rows marked as header rows (w:tblHeader), as Office JS. */
  get headerRowCount(): number {
    const rows = this.rows;
    let n = 0;
    while (n < rows.length && rows[n]!.isHeader) n++;
    return n;
  }
  set headerRowCount(count: number) {
    this.rows.forEach((row, i) => { row.isHeader = i < count; });
  }

  /** The cell at a row and cell index; throws when there is none, as Office JS's getCell does. */
  getCell(rowIndex: number, cellIndex: number): TableCell {
    const row = this.rows[rowIndex];
    if (!row) throw new Docx4JException(`No row ${rowIndex} in this table (${this.rowCount} rows)`);
    const cell = row.cells[cellIndex];
    if (!cell) throw new Docx4JException(`No cell ${cellIndex} in row ${rowIndex} (${row.cellCount} cells)`);
    return cell;
  }

  /** Adds rows at the start or the end, with the column count and widths of this table. */
  addRows(location: 'Start' | 'End', rowCount: number, values?: string[][]): TableRow[] {
    const widths = this.columnWidths();
    const added: Element<wml.Tr>[] = [];
    for (let i = 0; i < rowCount; i++) added.push(rowElement(widths, values?.[i]));
    // New rows go into the table's own content, not into a row-level content control.
    const container = (this.tbl.content ??= []) as Element[];
    const first = rowsOf(this.tbl)[0];
    const at = location === 'Start' && first && first.container === container ? container.indexOf(first.element) : location === 'Start' ? 0 : container.length;
    container.splice(at, 0, ...added);
    for (const row of added) linkParents(row, this.tbl);
    markRowsInserted(this.changeTracker, added);
    return added.map((element) => new TableRow(element, container, this));
  }

  /**
   * Removes `rowCount` rows from `rowIndex` (one by default), as Office JS. While the package
   * tracks changes the rows stay and take a `w:trPr/w:del` instead (CR-002 phase F).
   */
  deleteRows(rowIndex: number, rowCount = 1): void {
    const tracker = this.changeTracker;
    const rows = rowsOf(this.tbl).slice(rowIndex, rowIndex + rowCount);
    for (const row of rows.reverse()) {
      if (tracker) { tracker.markRowDeleted(row.element.value); continue; }
      const i = row.container.indexOf(row.element);
      if (i >= 0) row.container.splice(i, 1);
    }
  }

  /** Removes the table from its container; tracked, every row is marked deleted instead. */
  delete(): void {
    const tracker = this.changeTracker;
    if (tracker) { for (const row of rowsOf(this.tbl)) tracker.markRowDeleted(row.element.value); return; }
    const i = this.index;
    if (i >= 0) this.container.splice(i, 1);
  }

  /** The package's change tracker while `changeTrackingMode` is on (CR-002 phase F). */
  get changeTracker(): ChangeTracker | undefined {
    return this.parentBody.changeTracker;
  }

  /** The cell this table is nested in, if any (extension). */
  get parentTableCell(): TableCell | undefined {
    return cellOf(this.tbl, this.parentBody);
  }

  /** The text of the table: a line per paragraph, as docx4j TextUtils (extension). */
  get text(): string {
    return textOf(this.tbl);
  }

  /** The column widths in twips, from the grid or from the first row (extension). */
  columnWidths(): number[] {
    const grid = this.tbl.tblGrid?.gridCol;
    if (grid && grid.length) return grid.map((c) => c.w ?? Math.floor(DEFAULT_WIDTH / grid.length));
    const first = rowsOf(this.tbl)[0];
    const cells = first ? cellsOf(first.element.value) : [];
    const n = Math.max(1, cells.length);
    const each = Math.floor(DEFAULT_WIDTH / n);
    return Array.from({ length: n }, (_, i) => cells[i]?.element.value.tcPr?.tcW?.w ?? each);
  }

  private tblPr(): wml.TblPr {
    return (this.tbl.tblPr ??= {});
  }
}

/** A subset of Office JS `Word.TableRow` over a `w:tr`. */
export class TableRow {
  constructor(
    readonly element: Element<wml.Tr>,
    /** The array holding the row: the table's content, or a row-level content control's. */
    readonly container: Element[],
    readonly parentTable: Table,
  ) {}

  get tr(): wml.Tr {
    return this.element.value;
  }

  /** The row's index among the table's rows. */
  get rowIndex(): number {
    return rowsOf(this.parentTable.tbl).findIndex((r) => r.element === this.element);
  }

  get cellCount(): number {
    return cellsOf(this.tr).length;
  }

  get cells(): TableCell[] {
    return cellsOf(this.tr).map((c) => new TableCell(c.element, c.container, this));
  }

  /** The text of the row's cells. */
  get values(): string[] {
    return this.cells.map((c) => c.text);
  }
  set values(values: string[]) {
    const cells = this.cells;
    values.forEach((v, i) => { if (i < cells.length) cells[i]!.value = v; });
  }

  /** w:tblHeader: the row repeats at the top of every page (Office JS Table.headerRowCount). */
  get isHeader(): boolean {
    const item = this.trPrItem('tblHeader');
    return item !== undefined && (item.value as wml.BooleanDefaultTrue).val !== false;
  }
  set isHeader(header: boolean) {
    const items = this.tr.trPr?.cnfStyleOrDivIdOrGridBefore;
    const at = items?.findIndex((i) => i.name.localPart === 'tblHeader') ?? -1;
    if (!header) { if (items && at >= 0) items.splice(at, 1); return; }
    if (at >= 0) return;
    const trPr = (this.tr.trPr ??= {});
    (trPr.cnfStyleOrDivIdOrGridBefore ??= []).push(el.tblHeader({}) as never);
  }

  /** New rows before or after this one, with this table's columns. */
  insertRows(location: 'Before' | 'After', rowCount: number, values?: string[][]): TableRow[] {
    const widths = this.parentTable.columnWidths();
    const added: Element<wml.Tr>[] = [];
    for (let i = 0; i < rowCount; i++) added.push(rowElement(widths, values?.[i]));
    const at = this.container.indexOf(this.element) + (location === 'After' ? 1 : 0);
    this.container.splice(at, 0, ...added);
    const owner = ownerOfArray(this.container, this.parentTable.tbl) ?? this.parentTable.tbl;
    for (const row of added) linkParents(row, owner);
    markRowsInserted(this.parentTable.changeTracker, added);
    return added.map((element) => new TableRow(element, this.container, this.parentTable));
  }

  /** Removes the row; tracked, it stays and takes a `w:trPr/w:del` (CR-002 phase F). */
  delete(): void {
    const tracker = this.parentTable.changeTracker;
    if (tracker) { tracker.markRowDeleted(this.tr); return; }
    const i = this.container.indexOf(this.element);
    if (i >= 0) this.container.splice(i, 1);
  }

  private trPrItem(localPart: string): Element | undefined {
    return this.tr.trPr?.cnfStyleOrDivIdOrGridBefore?.find((i) => i.name.localPart === localPart);
  }
}

/** A subset of Office JS `Word.TableCell` over a `w:tc`: a body of its own, plus its position. */
export class TableCell {
  constructor(
    readonly element: Element<wml.Tc>,
    /** The array holding the cell: the row's content, or a cell-level content control's. */
    readonly container: Element[],
    readonly parentRow: TableRow,
  ) {}

  get tc(): wml.Tc {
    return this.element.value;
  }

  get parentTable(): Table {
    return this.parentRow.parentTable;
  }

  get rowIndex(): number {
    return this.parentRow.rowIndex;
  }

  get cellIndex(): number {
    return cellsOf(this.parentRow.tr).findIndex((c) => c.element === this.element);
  }

  /**
   * The cell's content as a `Body` (Office JS TableCell.body): the same part and package as the
   * table's body, addressed from the cell (`body/4/0/1`).
   */
  get body(): Body {
    const parent = this.parentTable.parentBody;
    return parent.sub(this.tc as { content?: Element[] }, parent.addressOf(this) ?? parent.prefix);
  }

  get paragraphs(): Paragraph[] {
    return this.body.paragraphs;
  }

  get tables(): Table[] {
    return this.body.tables;
  }

  /** The cell's text, a line per paragraph. */
  get text(): string {
    return textOf(this.tc);
  }

  /** Office JS TableCell.value: the text, replaced as one paragraph when set. */
  get value(): string {
    return this.text;
  }
  set value(text: string) {
    this.body.clear();
    this.body.insertParagraph(text, 'End');
  }

  insertParagraph(text: string, location: 'Start' | 'End'): Paragraph {
    return this.body.insertParagraph(text, location);
  }

  insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range {
    return this.body.insertText(text, location);
  }

  /** The cell width in points (w:tcW in twips), as Office JS reports it; 0 when it is not a fixed width. */
  get width(): number {
    const w = this.tc.tcPr?.tcW;
    return w?.type === 'dxa' ? (w.w ?? 0) / TWIPS_PER_POINT : 0;
  }
  set width(points: number) {
    const tcW = ((this.tc.tcPr ??= {}).tcW ??= {});
    tcW.w = Math.round(points * TWIPS_PER_POINT);
    tcW.type = 'dxa';
  }

  /** Office JS TableCell.columnWidth, the same value. */
  get columnWidth(): number {
    return this.width;
  }
  set columnWidth(points: number) {
    this.width = points;
  }
}

/** A `w:tr` of empty cells of the given widths, one paragraph per cell (docx4j's table helpers). */
function rowElement(widths: number[], values?: string[]): Element<wml.Tr> {
  return el.tr({
    content: widths.map((w, i) => el.tc({ tcPr: { tcW: { w, type: 'dxa' } }, content: [paragraphOfText(values?.[i] ?? '')] })),
  }) as Element<wml.Tr>;
}

/** The typed object whose content array this is, for PARENT links (the table, or a row-level control's content). */
function ownerOfArray(container: Element[], tbl: wml.Tbl): object | undefined {
  if (tbl.content === (container as never)) return tbl;
  for (const item of (tbl.content ?? []) as Element[]) {
    const v = item.value;
    if (typeof v !== 'object' || v === null) continue;
    const sdtContent = (v as { sdtContent?: { content?: Element[] } }).sdtContent;
    if (sdtContent && sdtContent.content === container) return sdtContent;
  }
  return undefined;
}

/**
 * The cell a value is in, from its PARENT links; undefined when it is not in a table. Shared by
 * `Paragraph.parentTableCell` and `Table.parentTableCell`.
 */
export function cellOf(value: object, parentBody: Body): TableCell | undefined {
  let tc: wml.Tc | undefined;
  let current: object | undefined = value;
  for (let guard = 0; current && guard < 64; guard++) {
    if (typeNameOf(current) === 'org_docx4j_wml.Tc') { tc = current as wml.Tc; break; }
    current = (current as { PARENT?: object }).PARENT;
  }
  if (!tc) return undefined;
  // a cell-level content control sits between the cell and the row
  const tr = ancestorOf(tc, 'org_docx4j_wml.Tr', 4);
  if (!tr) return undefined;
  const tbl = ancestorTable(tr);
  if (!tbl) return undefined;
  const tblContainer = childrenOf(((tbl as { PARENT?: object }).PARENT ?? {}) as object);
  const tblElement = tblContainer?.find((e) => e.value === tbl) as Element<wml.Tbl> | undefined;
  if (!tblElement || !tblContainer) return undefined;
  const table = new Table(tblElement, tblContainer, parentBody);
  const row = rowsOf(tbl).find((r) => r.element.value === tr);
  if (!row) return undefined;
  const rowView = new TableRow(row.element, row.container, table);
  const cell = cellsOf(tr).find((c) => c.element.value === tc);
  return cell ? new TableCell(cell.element, cell.container, rowView) : undefined;
}

/** The w:tbl a row is in, through a row-level content control if there is one. */
function ancestorTable(tr: object): wml.Tbl | undefined {
  return ancestorOf(tr, 'org_docx4j_wml.Tbl', 4) as wml.Tbl | undefined;
}

/** The nearest ancestor of a type, at most `depth` PARENT steps up (the sdt wrappers are steps). */
function ancestorOf(value: object, typeName: string, depth: number): object | undefined {
  let current: object | undefined = (value as { PARENT?: object }).PARENT;
  for (let guard = 0; current && guard < depth; guard++) {
    if (typeNameOf(current) === typeName) return current;
    current = (current as { PARENT?: object }).PARENT;
  }
  return undefined;
}

/** A new row is an insertion: `w:trPr/w:ins` on it and a `w:ins` around every run it holds. */
function markRowsInserted(tracker: ChangeTracker | undefined, rows: Element<wml.Tr>[]): void {
  if (!tracker) return;
  for (const row of rows) {
    tracker.markRowInserted(row.value);
    for (const cell of cellsOf(row.value)) {
      for (const block of childrenOf(cell.element.value) ?? []) {
        if (typeNameOf(block) === 'org_docx4j_wml.P') trackInsertedParagraph(tracker, block.value as wml.P);
      }
    }
  }
}
