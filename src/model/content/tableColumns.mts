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
// Check 40 (2026-10-09) settled the rest: the new column copies the column to the right of the
// boundary when there is one, else the left; the window rescale is proportional; a boundary inside
// a cell spanning columns puts the new cell before that cell, its span kept; and through Office JS
// with tracking on, a column deleted loses its cells untracked but records the table change, a
// column added records a `w:tcPrChange` on the new cells too and their text in a `w:ins` with no
// mark insertion, and the `w:tblGridChange` holds the grid as it is after the change, both ways.
// Check 41 (2026-10-09) added: a new cell takes the right-hand cell's shading and its paragraph's
// properties, not its borders or vertical alignment; an AutoFit-to-contents add records nothing;
// an AutoFit-to-window add records the grid of the old and new column edges together, the old
// cells spanning their former extents over it, a new cell's width a placeholder (check 27's form).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy } from '@docx4j/generated-objects-ts';
import { tc as cellElementOf } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { type Element, rowsOf, cellsOf, childrenOf, linkParents, paragraphOf } from './tree.mjs';
import { type ChangeTracker, wrapInsertedRuns } from './tracking.mjs';

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

/**
 * The cell properties a new cell takes from the cell to the right of the boundary: its width (the
 * form; the value is written from the grid) and its shading, and nothing else - check 41 saw Word
 * copy the shading and leave the borders and the vertical alignment behind.
 */
const COPIED = ['tcW', 'shd'];

function cellPropertiesFrom(template: wml.Tc | undefined): wml.TcPr | undefined {
  if (!template?.tcPr) return undefined;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(template.tcPr)) if (COPIED.includes(key)) out[key] = deepCopy(value);
  return Object.keys(out).length > 0 ? (out as wml.TcPr) : undefined;
}

/** The paragraph properties a new cell's paragraph takes: the template's first paragraph's, its mark's run properties included, less any revision (check 41: centred bold, right-aligned italic, as the neighbour's). */
function paragraphPropertiesFrom(template: wml.Tc | undefined): wml.PPr | undefined {
  const first = (childrenOf(template ?? {}) ?? []).find((item) => (item.value as { TYPE_NAME?: string } | null)?.TYPE_NAME === 'org_docx4j_wml.P');
  const pPr = (first?.value as wml.P | undefined)?.pPr;
  if (!pPr) return undefined;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(pPr)) if (!['pPrChange', 'sectPr', 'PARENT'].includes(key)) out[key] = deepCopy(value);
  const rPr = out.rPr as Record<string, unknown> | undefined;
  if (rPr) for (const key of ['ins', 'del', 'rPrChange', 'moveFrom', 'moveTo']) delete rPr[key];
  return Object.keys(out).length > 0 ? (out as wml.PPr) : undefined;
}

/** Where a new column's cells go in one row, and which cell they take their properties from. */
interface RowPlan {
  tr: wml.Tr;
  /** Insert new cells into this array at this index (the row's content or a cell-level control's). */
  container?: Element[];
  at?: number;
  template?: wml.Tc;
  /** The row's skipped columns before or after its cells grow by the count instead. */
  skipped?: RowMember;
}

/** The new cells of one row, each a `w:tc` of one paragraph with the template's shading and paragraph properties. */
function newCells(count: number, values: string[] | undefined, template: wml.Tc | undefined): Element<wml.Tc>[] {
  const out: Element<wml.Tc>[] = [];
  for (let i = 0; i < count; i++) {
    // an empty cell holds an empty `w:p`, as Word writes one; a value its text
    const text = values?.[i] ?? '';
    const cell = cellElementOf(text === '' ? [paragraphOf([]) as Element] : text) as Element<wml.Tc>;
    const tcPr = cellPropertiesFrom(template);
    if (tcPr) cell.value.tcPr = tcPr; else delete cell.value.tcPr;
    const pPr = paragraphPropertiesFrom(template);
    const paragraph = (childrenOf(cell.value) ?? []).find((item) => (item.value as { TYPE_NAME?: string }).TYPE_NAME === 'org_docx4j_wml.P');
    if (pPr && paragraph) { (paragraph.value as wml.P).pPr = pPr; linkParents(pPr, paragraph.value as object); }
    out.push(cell);
  }
  return out;
}

export interface ColumnInsertion {
  /** The new cells per row, in row order; a row whose spanning cell was widened, or whose skipped columns grew, has none. */
  cells: Element<wml.Tc>[][];
}

/**
 * Inserts `count` grid columns at `gridIndex` (0 to the column count). The new column copies the
 * width of the column to the right of the boundary when there is one, else the left (check 40:
 * Insert Right in a narrow cell beside a wide one took the wide width; Insert Right in a spanning
 * cell took the column after it), and each row's new cell takes the properties of the cell on the
 * same side, less its span and merges. A cell spanning the boundary keeps its span and the new
 * cell goes before it (check 40b1); a row's `w:gridBefore` or `w:gridAfter` covering the boundary
 * grows instead. `values` is by row, then by new column. While changes are tracked the table
 * change is recorded as Word's Office JS records one (check 40c): the old properties, every
 * cell's properties - the new cells' as made - the grid as it is after, and the new cells' text in
 * a `w:ins`, with no mark insertion.
 */
export function insertGridColumns(tbl: wml.Tbl, textWidth: number, gridIndex: number, count: number, values: string[][] | undefined, tracker: ChangeTracker | undefined): ColumnInsertion {
  if (!Number.isInteger(count) || count < 1) throw new Docx4JException(`columnCount must be a positive integer, not ${count}`);
  const old = gridWidthsOf(tbl, textWidth);
  if (gridIndex < 0 || gridIndex > old.length) throw new Docx4JException(`No column boundary ${gridIndex} in a grid of ${old.length} columns`);
  const mode = layoutModeOf(tbl);
  const copied = old[gridIndex < old.length ? gridIndex : old.length - 1]!;
  const unscaled = [...old.slice(0, gridIndex), ...Array.from({ length: count }, () => copied), ...old.slice(gridIndex)];
  const widths = mode === 'window' ? rescaled(unscaled, old.reduce((a, b) => a + b, 0)) : unscaled;

  const rows = rowsOf(tbl).map((row) => row.element.value);
  const plans: RowPlan[] = rows.map((tr) => planRow(tr, gridIndex));
  // an AutoFit-to-contents add changes no width, and Office JS records nothing for it (check 41c)
  const record = tracker && mode !== 'contents' ? recordTable(tbl, rows, tracker) : undefined;

  const result: ColumnInsertion = { cells: [] };
  plans.forEach((plan, rowIndex) => {
    const tr = plan.tr;
    if (plan.skipped) { plan.skipped.value.val = Number(plan.skipped.value.val ?? 0) + count; result.cells.push([]); return; }
    const cells = newCells(count, values?.[rowIndex], plan.template);
    plan.container!.splice(plan.at!, 0, ...cells);
    const owner = ownerOfCells(tr, plan.container!);
    for (const cell of cells) linkParents(cell, owner);
    if (tracker) for (const cell of cells) for (const block of childrenOf(cell.value) ?? []) wrapInsertedRuns(tracker, childrenOf(block.value as object), block.value as object);
    result.cells.push(cells);
  });

  // the grid, the table's width, and the cells' widths in their forms
  writeGrid(tbl, widths);
  const tblW = tbl.tblPr?.tblW;
  if (mode === 'fixed' && tblW?.type === 'dxa') tblW.w = Number(tblW.w ?? 0) + copied * count;
  writeCellWidths(tbl, widths, mode === 'window' ? 'all' : result.cells.flat().map((c) => c.value));
  if (record) record(result.cells.flat().map((c) => c.value), { unscaled, scaled: widths });
  return result;
}

/** Where the boundary falls in a row: at or inside a cell (the new cells go before it), in its skipped columns, or after its last cell. */
function planRow(tr: wml.Tr, gridIndex: number): RowPlan {
  const before = rowMember(tr, 'gridBefore');
  const after = rowMember(tr, 'gridAfter');
  const skippedBefore = Number(before?.value.val ?? 0);
  if (before && gridIndex < skippedBefore) return { tr, skipped: before };
  const cells = cellsOf(tr);
  let at = skippedBefore;
  for (const cell of cells) {
    const tc = cell.element.value;
    const end = at + spanOf(tc);
    // at the cell's start, or inside its span (check 40b1: the span kept, the new cell before it)
    if (gridIndex >= at && gridIndex < end) return { tr, container: cell.container, at: cell.container.indexOf(cell.element), template: tc };
    at = end;
  }
  const skippedAfter = Number(after?.value.val ?? 0);
  if (after && gridIndex < at + skippedAfter) return { tr, skipped: after };
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

/** The grids a window-mode insert went through: the old widths with the new column at its copied width, and the same rescaled to the width the table had. */
interface Rescale {
  unscaled: number[];
  scaled: number[];
}

/**
 * What Word records for a column added or removed through Office JS with tracking on (checks 40c
 * and 41c, the form of section 35's table change): a `w:tblPrChange` with the properties as they
 * were, a `w:tcPrChange` on every cell with its properties as they were - a new cell's as made -
 * and a `w:tblGridChange` holding the grid **as it is after the change**. Where the change
 * rescaled the columns (AutoFit to window), the recorded grid is instead the old and new column
 * edges together - the unscaled grid's edges and the scaled grid's, merged - with every cell's
 * record spanning the merged columns its unscaled extent covers, and a new cell's recorded width a
 * placeholder (`1 pct`; check 27's form, as Word wrote it in check 41c), so that rejecting puts
 * the old widths back over a grid that `collapseGrid` then folds. Taken before the change and
 * written after it, so that `TrackedChange` lists the whole as one table change. A record already
 * there (an earlier pending change) is kept, so the original state stays the recorded one.
 */
function recordTable(tbl: wml.Tbl, rows: wml.Tr[], tracker: ChangeTracker): (added?: wml.Tc[], rescale?: Rescale) => void {
  const tblPr = tbl.tblPr;
  const oldTblPr = tblPr && !tblPr.tblPrChange ? without(tblPr, ['tblPrChange']) : undefined;
  const recordGrid = !tbl.tblGrid?.tblGridChange;
  const oldCells: [wml.Tc, wml.TcPrInner][] = [];
  for (const tr of rows) for (const cell of cellsOf(tr)) {
    const tc = cell.element.value;
    if (!tc.tcPr?.tcPrChange) oldCells.push([tc, recordedCellProperties(tc)]);
  }
  return (added: wml.Tc[] = [], rescale?: Rescale) => {
    if (oldTblPr && tbl.tblPr) {
      tbl.tblPr.tblPrChange = { TYPE_NAME: 'org_docx4j_wml.CTTblPrChange', ...tracker.markup(), tblPr: { TYPE_NAME: 'org_docx4j_wml.CTTblPrBase', ...oldTblPr } as wml.CTTblPrBase };
      linkParents(tbl.tblPr.tblPrChange, tbl.tblPr);
    }
    const merged = rescale && rescale.unscaled.some((w, i) => w !== rescale.scaled[i]) ? mergedEdges(rescale.unscaled, rescale.scaled) : undefined;
    if (recordGrid && tbl.tblGrid) {
      const gridCol = merged ? merged.widths.map((w) => ({ TYPE_NAME: 'org_docx4j_wml.TblGridCol', w } as wml.TblGridCol)) : (tbl.tblGrid.gridCol ?? []).map((col) => deepCopy(col));
      tbl.tblGrid.tblGridChange = { TYPE_NAME: 'org_docx4j_wml.CTTblGridChange', id: tracker.nextId(), tblGrid: { TYPE_NAME: 'org_docx4j_wml.TblGridBase', gridCol } };
      linkParents(tbl.tblGrid.tblGridChange, tbl.tblGrid);
    }
    const extents = merged ? cellExtents(tbl) : undefined;
    const cells: [wml.Tc, wml.TcPrInner][] = [...oldCells.filter(([tc]) => isInTable(tc, tbl)), ...added.map((tc) => [tc, recordedCellProperties(tc)] as [wml.Tc, wml.TcPrInner])];
    for (const [tc, old] of cells) {
      const tcPr = (tc.tcPr ??= {});
      const recorded: Record<string, unknown> = { ...old };
      if (merged && extents) {
        // over the merged grid: the columns the cell's unscaled extent covers
        const extent = extents.get(tc);
        if (extent) {
          const span = merged.index(unscaledEdge(rescale!.unscaled, extent.end)) - merged.index(unscaledEdge(rescale!.unscaled, extent.start));
          if (span > 1) recorded.gridSpan = { val: span }; else delete recorded.gridSpan;
        }
        if (added.includes(tc) && (recorded.tcW as wml.TblWidth | undefined)?.type === 'pct') recorded.tcW = { w: 1, type: 'pct' };
      }
      tcPr.tcPrChange = { TYPE_NAME: 'org_docx4j_wml.CTTcPrChange', ...tracker.markup(), tcPr: { TYPE_NAME: 'org_docx4j_wml.TcPrInner', ...recorded } as wml.TcPrInner };
      linkParents(tcPr.tcPrChange, tcPr);
    }
  };
}

/** The cumulative edge, in twips, of a column index over a grid. */
function unscaledEdge(widths: number[], column: number): number {
  return widths.slice(0, column).reduce((a, b) => a + b, 0);
}

/** The old and new grids' edges together, as merged columns, and where an edge falls among them. */
function mergedEdges(unscaled: number[], scaled: number[]): { widths: number[]; index: (edge: number) => number } {
  const edges = new Set<number>([0]);
  let at = 0; for (const w of unscaled) { at += w; edges.add(at); }
  at = 0; for (const w of scaled) { at += w; edges.add(at); }
  const sorted = [...edges].sort((a, b) => a - b);
  return { widths: sorted.slice(1).map((edge, i) => edge - sorted[i]!), index: (edge) => sorted.indexOf(edge) };
}

/** Each cell's column extent in the grid as it is now, through the rows' spans and skipped columns. */
function cellExtents(tbl: wml.Tbl): Map<wml.Tc, { start: number; end: number }> {
  const out = new Map<wml.Tc, { start: number; end: number }>();
  for (const row of rowsOf(tbl)) {
    const tr = row.element.value;
    let at = Number(rowMember(tr, 'gridBefore')?.value.val ?? 0);
    for (const cell of cellsOf(tr)) {
      const tc = cell.element.value;
      const end = at + spanOf(tc);
      out.set(tc, { start: at, end });
      at = end;
    }
  }
  return out;
}

/** A cell's properties as a record holds them: everything but the markers and an earlier record. */
function recordedCellProperties(tc: wml.Tc): wml.TcPrInner {
  return without(tc.tcPr, ['tcPrChange', 'cellIns', 'cellDel', 'cellMerge']) as wml.TcPrInner;
}

/** Whether a cell is still in the table (a deleted column's cells are not). */
function isInTable(tc: wml.Tc, tbl: wml.Tbl): boolean {
  return rowsOf(tbl).some((row) => cellsOf(row.element.value).some((cell) => cell.element.value === tc));
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
 * cells' percentages written again. While changes are tracked the cells still go, untracked, as
 * Word's own Delete Columns and Office JS's `deleteColumns` leave them (known issues, entry 2;
 * check 40c), and the table change is recorded as Office JS records it: the old properties, the
 * remaining cells' properties, the grid as it is after.
 */
export function deleteGridColumns(tbl: wml.Tbl, textWidth: number, columnIndex: number, count: number, tracker: ChangeTracker | undefined): void {
  if (!Number.isInteger(count) || count < 1) throw new Docx4JException(`columnCount must be a positive integer, not ${count}`);
  const old = gridWidthsOf(tbl, textWidth);
  if (columnIndex < 0 || columnIndex + count > old.length) throw new Docx4JException(`No columns ${columnIndex} to ${columnIndex + count - 1} in a grid of ${old.length} columns`);
  const mode = layoutModeOf(tbl);
  const record = tracker && mode !== 'contents' ? recordTable(tbl, rowsOf(tbl).map((row) => row.element.value), tracker) : undefined;
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
  if (record) record();
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
