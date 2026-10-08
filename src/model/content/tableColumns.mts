// CR-002 section 42: columns added and removed, in Word's three table layout modes, as Word 365
// writes them (test/README.md check 39, 2026-10-09). Everything here works on the tree; Table.mts
// makes the views.
//
// Word's three choices, as its Table Properties states them: `w:tblLayout fixed` is fixed; a
// `w:tblW` in percent is AutoFit to window; anything else AutoFit to contents. Insert Right wrote:
//   fixed     the new column copies the width beside the caret; the grid gains it; a `w:tblW` in
//             twips grows by it, past the margin; nothing is rescaled;
//   window    the table keeps 100%: the grid is rescaled to the width it had, and every cell's
//             `w:tcW` is its share of the table in fiftieths of a percent, the cumulative shares
//             floored (1666, 1667, 1667);
//   contents  `auto` throughout; the grid is Word's cache of its last measurement, which a consumer
//             measures afresh at open, so the grid here is only of the right shape.
// A cell's width is written in the form the cell states, else the table's, else twips.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy } from '@docx4j/generated-objects-ts';
import { tc as cellElementOf } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { type Element, rowsOf, cellsOf, childrenOf, linkParents, paragraphOf } from './tree.mjs';
import { type ChangeTracker, trackInsertedBlocks } from './tracking.mjs';

/** The default text width in twips (A4 with 2.54 cm margins), as the objects package's `tbl` builder uses. */
export const DEFAULT_TEXT_WIDTH = 9026;
/** A table's width in fiftieths of a percent: `w:tblW w:type="pct" w:w="5000"` is the text width. */
const FULL_PCT = 5000;

export type TableLayoutMode = 'fixed' | 'window' | 'contents';

/** Which of Word's three layout modes a table is in, read as Word's Table Properties reads the file. */
export function layoutModeOf(tbl: wml.Tbl): TableLayoutMode {
  if (tbl.tblPr?.tblLayout?.type === 'fixed') return 'fixed';
  if (tbl.tblPr?.tblW?.type === 'pct') return 'window';
  return 'contents';
}

/** The section's text width in twips: the page less its margins and gutter, else the default. */
export function textWidthTwips(container: { sectPr?: wml.SectPr } | undefined): number {
  const sectPr = container?.sectPr;
  const pgSz = sectPr?.pgSz;
  if (!pgSz?.w) return DEFAULT_TEXT_WIDTH;
  const width = pgSz.w - (sectPr?.pgMar?.left ?? 0) - (sectPr?.pgMar?.right ?? 0) - (sectPr?.pgMar?.gutter ?? 0);
  return width > 0 ? width : DEFAULT_TEXT_WIDTH;
}

type RowMember = Element<{ val?: number | string }>;

function rowMember(tr: wml.Tr, name: 'gridBefore' | 'gridAfter'): RowMember | undefined {
  return (tr.trPr?.cnfStyleOrDivIdOrGridBefore ?? []).find((item) => (item as Element).name?.localPart === name) as RowMember | undefined;
}

function spanOf(tc: wml.Tc): number {
  return Number(tc.tcPr?.gridSpan?.val ?? 1) || 1;
}

function setSpan(tc: wml.Tc, count: number): void {
  const tcPr = (tc.tcPr ??= {});
  if (count > 1) tcPr.gridSpan = { ...(tcPr.gridSpan ?? {}), val: count } as wml.TcPrInner.GridSpan;
  else delete tcPr.gridSpan;
}

/** The grid's column widths in twips; a table without a `w:tblGrid` gets one from its first row (equal columns over the text width when the cells state none). */
export function gridWidthsOf(tbl: wml.Tbl, textWidth: number): number[] {
  const grid = tbl.tblGrid?.gridCol;
  if (grid && grid.length > 0) return grid.map((col) => Number(col.w ?? 0));
  const first = rowsOf(tbl)[0];
  const cells = first ? cellsOf(first.element.value) : [];
  const columns = Math.max(1, cells.reduce((n, cell) => n + spanOf(cell.element.value), 0));
  const each = Math.floor(textWidth / columns);
  const widths: number[] = [];
  for (const cell of cells) {
    const tc = cell.element.value;
    const span = spanOf(tc);
    const stated = tc.tcPr?.tcW?.type === 'dxa' ? Number(tc.tcPr.tcW.w ?? 0) : 0;
    for (let i = 0; i < span; i++) widths.push(stated > 0 ? Math.floor(stated / span) : each);
  }
  return widths.length > 0 ? widths : [each];
}

function writeGrid(tbl: wml.Tbl, widths: number[]): void {
  const grid = (tbl.tblGrid ??= {});
  grid.gridCol = widths.map((w) => ({ TYPE_NAME: 'org_docx4j_wml.TblGridCol', w } as wml.TblGridCol));
  linkParents(grid, tbl);
}

/**
 * Widths rescaled to a total: each floored, the remainder given to the columns with the largest
 * fractions, the earlier first on a tie, so that equal columns inserted and then removed come
 * back equal (Word's own grid, 3005/3006/3006, is its layout's measurement and not reproduced).
 */
function rescaled(widths: number[], total: number): number[] {
  const sum = widths.reduce((a, b) => a + b, 0);
  if (sum <= 0 || total <= 0) return widths.slice();
  const exact = widths.map((w) => (w * total) / sum);
  const out = exact.map((w) => Math.floor(w));
  let remainder = total - out.reduce((a, b) => a + b, 0);
  const order = exact.map((w, i) => ({ i, fraction: w - Math.floor(w) })).sort((a, b) => b.fraction - a.fraction || a.i - b.i);
  for (const { i } of order) { if (remainder <= 0) break; out[i]! += 1; remainder--; }
  return out;
}

/** A cell's share of the table in fiftieths of a percent over grid columns [start, end): the cumulative shares floored, as Word writes them. */
function pctShare(widths: number[], start: number, end: number): number {
  const total = widths.reduce((a, b) => a + b, 0);
  if (total <= 0) return 0;
  const cumulative = (to: number): number => Math.floor((FULL_PCT * widths.slice(0, to).reduce((a, b) => a + b, 0)) / total);
  return cumulative(end) - cumulative(start);
}

/** The form a cell's width takes: its own `w:tcW` type, else the table's `w:tblW` type, else twips (section 42.2 item 4). */
export function widthTypeOf(tc: wml.Tc | undefined, tbl: wml.Tbl): string {
  return tc?.tcPr?.tcW?.type ?? tbl.tblPr?.tblW?.type ?? 'dxa';
}

/** Writes a cell's `w:tcW` for the grid columns it covers, in the form it takes; `auto` and `nil` widths are left as they are. */
function writeCellWidth(tc: wml.Tc, tbl: wml.Tbl, widths: number[], start: number, end: number): void {
  const type = widthTypeOf(tc, tbl);
  const tcPr = (tc.tcPr ??= {});
  if (type === 'pct') tcPr.tcW = { ...(tcPr.tcW ?? {}), w: pctShare(widths, start, end), type: 'pct' };
  else if (type === 'dxa') tcPr.tcW = { ...(tcPr.tcW ?? {}), w: widths.slice(start, end).reduce((a, b) => a + b, 0), type: 'dxa' };
  else if (!tcPr.tcW) tcPr.tcW = { w: 0, type };
}

/** The typed object a row's cell array belongs to: the row, or a cell-level control's content. */
function ownerOfCells(tr: wml.Tr, container: Element[]): object {
  if (tr.content === (container as never)) return tr;
  for (const item of (tr.content ?? []) as Element[]) {
    const v = item.value as { sdtContent?: { content?: Element[] } } | null;
    if (v && typeof v === 'object' && v.sdtContent?.content === container) return v.sdtContent;
  }
  return tr;
}

/** The cell properties a new cell takes from the cell beside it: everything but its span, merges, markers and record. */
const NOT_COPIED = ['gridSpan', 'vMerge', 'hMerge', 'tcPrChange', 'cellIns', 'cellDel', 'cellMerge', 'PARENT'];

function cellPropertiesFrom(template: wml.Tc | undefined): wml.TcPr | undefined {
  if (!template?.tcPr) return undefined;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(template.tcPr)) if (!NOT_COPIED.includes(key)) out[key] = deepCopy(value);
  return out as wml.TcPr;
}

/** Where a new column's cells go in one row, and which cell they take their properties from. */
interface RowPlan {
  tr: wml.Tr;
  /** Insert new cells into this array at this index (the row's content or a cell-level control's). */
  container?: Element[];
  at?: number;
  template?: wml.Tc;
  /** A cell whose span covers the boundary grows by the count instead. */
  widen?: wml.Tc;
  /** The row's skipped columns before or after its cells grow by the count instead. */
  skipped?: RowMember;
}

/** The new cells of one row, each a `w:tc` of one paragraph with the template's properties. */
function newCells(count: number, values: string[] | undefined, template: wml.Tc | undefined): Element<wml.Tc>[] {
  const out: Element<wml.Tc>[] = [];
  for (let i = 0; i < count; i++) {
    // an empty cell holds an empty `w:p`, as Word writes one; a value its text
    const text = values?.[i] ?? '';
    const cell = cellElementOf(text === '' ? [paragraphOf([]) as Element] : text) as Element<wml.Tc>;
    const tcPr = cellPropertiesFrom(template);
    if (tcPr) cell.value.tcPr = tcPr; else delete cell.value.tcPr;
    out.push(cell);
  }
  return out;
}

export interface ColumnInsertion {
  /** The new cells per row, in row order; a row whose spanning cell was widened, or whose skipped columns grew, has none. */
  cells: Element<wml.Tc>[][];
}

/**
 * Inserts `count` grid columns at `gridIndex` (0 to the column count), the new cells taking the
 * properties and width of the column beside the boundary: the one to its left (an insert after a
 * cell, or at the end) or to its right (before a cell, or at the start). A cell spanning the
 * boundary grows by the count; a row's `w:gridBefore` or `w:gridAfter` covering it grows the same
 * way. `values` is by row, then by new column. While changes are tracked the table change is
 * recorded as Word records one (section 42.3): the old properties, grid and every cell's
 * properties, and the new cells' text inserted.
 */
export function insertGridColumns(tbl: wml.Tbl, textWidth: number, gridIndex: number, count: number, values: string[][] | undefined, neighbour: 'left' | 'right', tracker: ChangeTracker | undefined): ColumnInsertion {
  if (!Number.isInteger(count) || count < 1) throw new Docx4JException(`columnCount must be a positive integer, not ${count}`);
  const old = gridWidthsOf(tbl, textWidth);
  if (gridIndex < 0 || gridIndex > old.length) throw new Docx4JException(`No column boundary ${gridIndex} in a grid of ${old.length} columns`);
  const mode = layoutModeOf(tbl);
  const neighbourColumn = Math.min(old.length - 1, Math.max(0, neighbour === 'left' ? gridIndex - 1 : gridIndex));
  const copied = old[neighbourColumn]!;
  let widths = [...old.slice(0, gridIndex), ...Array.from({ length: count }, () => copied), ...old.slice(gridIndex)];
  if (mode === 'window') widths = rescaled(widths, old.reduce((a, b) => a + b, 0));

  const rows = rowsOf(tbl).map((row) => row.element.value);
  const plans: RowPlan[] = rows.map((tr) => planRow(tr, gridIndex, neighbour));
  const record = tracker ? recordTable(tbl, rows, tracker) : undefined;

  const result: ColumnInsertion = { cells: [] };
  plans.forEach((plan, rowIndex) => {
    const tr = plan.tr;
    if (plan.skipped) { plan.skipped.value.val = Number(plan.skipped.value.val ?? 0) + count; result.cells.push([]); return; }
    if (plan.widen) { setSpan(plan.widen, spanOf(plan.widen) + count); result.cells.push([]); return; }
    const cells = newCells(count, values?.[rowIndex], plan.template);
    plan.container!.splice(plan.at!, 0, ...cells);
    const owner = ownerOfCells(tr, plan.container!);
    for (const cell of cells) linkParents(cell, owner);
    if (tracker) for (const cell of cells) trackInsertedBlocks(tracker, childrenOf(cell.value));
    result.cells.push(cells);
  });

  // the grid, the table's width, and the cells' widths in their forms
  writeGrid(tbl, widths);
  const tblW = tbl.tblPr?.tblW;
  if (mode === 'fixed' && tblW?.type === 'dxa') tblW.w = Number(tblW.w ?? 0) + copied * count;
  writeCellWidths(tbl, widths, mode === 'window' ? 'all' : result.cells.flat().map((c) => c.value));
  if (record) record();
  return result;
}

/** Where the boundary falls in a row: before a cell, inside a spanning cell, in its skipped columns, or after its last cell. */
function planRow(tr: wml.Tr, gridIndex: number, neighbour: 'left' | 'right'): RowPlan {
  const before = rowMember(tr, 'gridBefore');
  const after = rowMember(tr, 'gridAfter');
  const skippedBefore = Number(before?.value.val ?? 0);
  if (before && (gridIndex < skippedBefore || (gridIndex === skippedBefore && neighbour === 'left' && skippedBefore > 0))) return { tr, skipped: before };
  const cells = cellsOf(tr);
  let at = skippedBefore;
  for (let i = 0; i < cells.length; i++) {
    const cell = cells[i]!;
    const tc = cell.element.value;
    const end = at + spanOf(tc);
    if (gridIndex === at) {
      const previous = i > 0 ? cells[i - 1]!.element.value : undefined;
      return { tr, container: cell.container, at: cell.container.indexOf(cell.element), template: neighbour === 'left' && previous ? previous : tc };
    }
    if (gridIndex > at && gridIndex < end) return { tr, widen: tc };
    at = end;
  }
  const skippedAfter = Number(after?.value.val ?? 0);
  if (after && (gridIndex < at + skippedAfter || (gridIndex === at && neighbour === 'right' && skippedAfter > 0))) return { tr, skipped: after };
  const last = cells[cells.length - 1];
  if (!last) return { tr, container: (tr.content ??= []) as Element[], at: (tr.content as Element[]).length, template: undefined };
  return { tr, container: last.container, at: last.container.indexOf(last.element) + 1, template: last.element.value };
}

/** Every cell's `w:tcW` from the grid (`all`), or only the given cells', each in the form it takes. */
function writeCellWidths(tbl: wml.Tbl, widths: number[], which: 'all' | wml.Tc[]): void {
  for (const row of rowsOf(tbl)) {
    const tr = row.element.value;
    let at = Number(rowMember(tr, 'gridBefore')?.value.val ?? 0);
    for (const cell of cellsOf(tr)) {
      const tc = cell.element.value;
      const end = at + spanOf(tc);
      if (which === 'all' || which.includes(tc)) writeCellWidth(tc, tbl, widths, at, end);
      at = end;
    }
  }
}

/**
 * What Word records for a table change (section 42.3, from section 29's measurement): a
 * `w:tblPrChange` with the properties as they were, a `w:tblGridChange` with the old grid, and a
 * `w:tcPrChange` on every existing cell with its old properties - taken before the change and
 * written after it, so that `TrackedChange` lists the whole as one table change and rejecting it
 * puts the recorded table back. A record already there (an earlier pending change) is kept, so
 * the original state stays the recorded one.
 */
function recordTable(tbl: wml.Tbl, rows: wml.Tr[], tracker: ChangeTracker): () => void {
  const tblPr = tbl.tblPr;
  const oldTblPr = tblPr && !tblPr.tblPrChange ? without(tblPr, ['tblPrChange']) : undefined;
  const oldGrid = tbl.tblGrid && !tbl.tblGrid.tblGridChange ? (tbl.tblGrid.gridCol ?? []).map((col) => deepCopy(col)) : undefined;
  const oldCells: [wml.Tc, wml.TcPrInner][] = [];
  for (const tr of rows) for (const cell of cellsOf(tr)) {
    const tc = cell.element.value;
    if (!tc.tcPr?.tcPrChange) oldCells.push([tc, without(tc.tcPr, [...NOT_COPIED.filter((k) => k !== 'gridSpan' && k !== 'vMerge' && k !== 'hMerge'), 'PARENT']) as wml.TcPrInner]);
  }
  return () => {
    if (oldTblPr && tbl.tblPr) {
      tbl.tblPr.tblPrChange = { TYPE_NAME: 'org_docx4j_wml.CTTblPrChange', ...tracker.markup(), tblPr: { TYPE_NAME: 'org_docx4j_wml.CTTblPrBase', ...oldTblPr } as wml.CTTblPrBase };
      linkParents(tbl.tblPr.tblPrChange, tbl.tblPr);
    }
    if (oldGrid && tbl.tblGrid) {
      tbl.tblGrid.tblGridChange = { TYPE_NAME: 'org_docx4j_wml.CTTblGridChange', id: tracker.nextId(), tblGrid: { TYPE_NAME: 'org_docx4j_wml.TblGridBase', gridCol: oldGrid } };
      linkParents(tbl.tblGrid.tblGridChange, tbl.tblGrid);
    }
    for (const [tc, old] of oldCells) {
      const tcPr = (tc.tcPr ??= {});
      tcPr.tcPrChange = { TYPE_NAME: 'org_docx4j_wml.CTTcPrChange', ...tracker.markup(), tcPr: { TYPE_NAME: 'org_docx4j_wml.TcPrInner', ...old } as wml.TcPrInner };
      linkParents(tcPr.tcPrChange, tcPr);
    }
  };
}

function without(value: object | undefined, keys: string[]): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, v] of Object.entries(value ?? {})) if (!keys.includes(key) && key !== 'PARENT') out[key] = deepCopy(v);
  return out;
}

/**
 * Removes `count` grid columns from `columnIndex`: a cell covering only removed columns goes, one
 * spanning further shrinks, a row's skipped columns shrink; the grid loses the columns, a `w:tblW`
 * in twips loses their width, and in window mode the grid is rescaled to the width it had and the
 * cells' percentages written again. Not tracked: Word does not track a column deleted (known
 * issues, entry 2), so a tracked document's deletion is as untracked as Word's own.
 */
export function deleteGridColumns(tbl: wml.Tbl, textWidth: number, columnIndex: number, count: number): void {
  if (!Number.isInteger(count) || count < 1) throw new Docx4JException(`columnCount must be a positive integer, not ${count}`);
  const old = gridWidthsOf(tbl, textWidth);
  if (columnIndex < 0 || columnIndex + count > old.length) throw new Docx4JException(`No columns ${columnIndex} to ${columnIndex + count - 1} in a grid of ${old.length} columns`);
  const mode = layoutModeOf(tbl);
  const from = columnIndex;
  const to = columnIndex + count;
  const overlap = (start: number, end: number): number => Math.max(0, Math.min(end, to) - Math.max(start, from));
  for (const row of rowsOf(tbl)) {
    const tr = row.element.value;
    const before = rowMember(tr, 'gridBefore');
    const skippedBefore = Number(before?.value.val ?? 0);
    if (before) before.value.val = skippedBefore - overlap(0, skippedBefore);
    let at = skippedBefore;
    for (const cell of cellsOf(tr)) {
      const tc = cell.element.value;
      const span = spanOf(tc);
      const end = at + span;
      const covered = overlap(at, end);
      if (covered === span) {
        const i = cell.container.indexOf(cell.element);
        if (i >= 0) cell.container.splice(i, 1);
      } else if (covered > 0) {
        setSpan(tc, span - covered);
      }
      at = end;
    }
    const after = rowMember(tr, 'gridAfter');
    if (after) { const skippedAfter = Number(after.value.val ?? 0); after.value.val = skippedAfter - overlap(at, at + skippedAfter); }
  }
  const removed = old.slice(from, to).reduce((a, b) => a + b, 0);
  let widths = [...old.slice(0, from), ...old.slice(to)];
  if (mode === 'window') widths = rescaled(widths, old.reduce((a, b) => a + b, 0));
  writeGrid(tbl, widths);
  const tblW = tbl.tblPr?.tblW;
  if (tblW?.type === 'dxa') tblW.w = Math.max(0, Number(tblW.w ?? 0) - removed);
  writeCellWidths(tbl, widths, mode === 'window' ? 'all' : shrunkCells(tbl));
}

/** The cells whose span no longer matches their stated twips width: those that shrank. */
function shrunkCells(tbl: wml.Tbl): wml.Tc[] {
  const out: wml.Tc[] = [];
  const widths = tbl.tblGrid?.gridCol?.map((col) => Number(col.w ?? 0)) ?? [];
  for (const row of rowsOf(tbl)) {
    const tr = row.element.value;
    let at = Number(rowMember(tr, 'gridBefore')?.value.val ?? 0);
    for (const cell of cellsOf(tr)) {
      const tc = cell.element.value;
      const end = at + spanOf(tc);
      const tcW = tc.tcPr?.tcW;
      if (tcW?.type === 'dxa' && Number(tcW.w ?? 0) !== widths.slice(at, end).reduce((a, b) => a + b, 0)) out.push(tc);
      at = end;
    }
  }
  return out;
}

/** The grid columns a cell covers: its start and the one past its end. */
export function gridRangeOf(tr: wml.Tr, tc: wml.Tc): { start: number; end: number } | undefined {
  let at = Number(rowMember(tr, 'gridBefore')?.value.val ?? 0);
  for (const cell of cellsOf(tr)) {
    const end = at + spanOf(cell.element.value);
    if (cell.element.value === tc) return { start: at, end };
    at = end;
  }
  return undefined;
}
