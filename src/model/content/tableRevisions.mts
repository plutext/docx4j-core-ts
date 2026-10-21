// CR-002 section 35: a table's property revisions - w:tblPrChange, w:tblGridChange,
// w:tblPrExChange, w:trPrChange and w:tcPrChange - which Word writes as one set for one change and
// Office JS lists as one change per table, or per row where only a row changed (test/README.md
// checks 18, 27 and 28). Everything here works on the tree; TrackedChange.mts makes the views.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy } from '@docx4j/generated-objects-ts';
import { type Element, typeNameOf, rowsOf, cellsOf, cellLayoutsOf, childrenOf, linkParents, textOfView } from './tree.mjs';
import { canonical } from './tracking.mjs';

/** The marker members of a `w:tcPr` that are revisions of their own, kept whatever a reject restores. */
const CELL_MARKERS = ['cellIns', 'cellDel', 'cellMerge'];
/** The members of a `w:trPr` that are the row's own revisions. */
const ROW_MARKERS = ['ins', 'del'];

/** A value without some of its members, for comparing what a record holds with what stands. */
function without(value: object | undefined, keys: string[]): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, v] of Object.entries(value ?? {})) if (!keys.includes(key)) out[key] = v;
  return out;
}

function differs(current: object | undefined, recorded: object | undefined): boolean {
  return canonical(current ?? {}) !== canonical(recorded ?? {});
}

/**
 * A table's rows and a row's cells in every branch of a kept row- or cell-level `mc:AlternateContent`:
 * the table's property records are one change whichever branch holds them, and resolving it reaches
 * every branch (CR-013 phase 3; docx4j's `Lister.allRowsOf` and `allCellsOf`, its CR-051 phase 2, from
 * docx4j-python's finding). The change's text reads the selected branch (`tableChangeText`).
 */
function allRows(tbl: wml.Tbl): wml.Tr[] {
  return rowsOf(tbl, 'every').map((row) => row.element.value);
}

function allCells(tr: wml.Tr): wml.Tc[] {
  return cellsOf(tr, 'every').map((cell) => cell.element.value);
}

/** The rows a change takes in: a row with a `w:trPrChange` or a `w:tblPrExChange`, or a cell with a `w:tcPrChange`. */
function inChange(tr: wml.Tr): boolean {
  return tr.trPr?.trPrChange !== undefined || tr.tblPrEx?.tblPrExChange !== undefined
    || allCells(tr).some((tc) => tc.tcPr?.tcPrChange !== undefined);
}

/**
 * True when a row in the change is not as its records have it: its properties (a row without a
 * `w:trPrChange` had none before, check 28), a cell's, or its `w:tblPrEx`.
 */
function rowDiffers(tr: wml.Tr): boolean {
  if (differs(without(tr.trPr, [...ROW_MARKERS, 'trPrChange']), tr.trPr?.trPrChange?.trPr)) return true;
  if (tr.tblPrEx?.tblPrExChange && differs(without(tr.tblPrEx, ['tblPrExChange']), tr.tblPrEx.tblPrExChange.tblPrEx)) return true;
  return allCells(tr).some((tc) => {
    const tcPr = tc.tcPr;
    return tcPr?.tcPrChange !== undefined && differs(without(tcPr, [...CELL_MARKERS, 'tcPrChange']), tcPr.tcPrChange.tcPr);
  });
}

/** True when the table's own properties or its grid are not as recorded. */
function tableDiffers(tbl: wml.Tbl): boolean {
  const change = tbl.tblPr?.tblPrChange;
  if (change && differs(without(tbl.tblPr, ['tblPrChange']), change.tblPr)) return true;
  const grid = tbl.tblGrid?.tblGridChange;
  return grid !== undefined && differs(without(tbl.tblGrid, ['tblGridChange']), grid.tblGrid);
}

/** A table's property records, as the listing needs them; undefined when it has none that change anything. */
export interface TableRecords {
  /** The record whose author and date the change reports: the table's, else the first row's or cell's. */
  value: wml.CTTrackChange;
}

export function tableRecordsOf(tbl: wml.Tbl): TableRecords | undefined {
  const changed = allRows(tbl).filter(inChange);
  if (!tbl.tblPr?.tblPrChange && !tbl.tblGrid?.tblGridChange && changed.length === 0) return undefined;
  // a set of records that holds no difference is no change (checks 24 and 26's rule)
  if (!tableDiffers(tbl) && !changed.some(rowDiffers)) return undefined;
  const value = tbl.tblPr?.tblPrChange
    ?? changed.map((tr) => tr.trPr?.trPrChange ?? tr.tblPrEx?.tblPrExChange ?? allCells(tr).map((tc) => tc.tcPr?.tcPrChange).find(Boolean)).find(Boolean);
  return value ? { value } : undefined;
}

/** True when the table has any property record at all, changing anything or not. */
export function hasTableRecords(tbl: wml.Tbl): boolean {
  return tbl.tblPr?.tblPrChange !== undefined || tbl.tblGrid?.tblGridChange !== undefined || allRows(tbl).some(inChange);
}

/** A cell's text as Office JS reports it in a table's change: its paragraphs, a `\r` between them. */
function cellText(tc: wml.Tc): string {
  return (childrenOf(tc) ?? []).filter((el) => typeNameOf(el) === 'org_docx4j_wml.P').map((el) => textOfView(el.value as object, { view: 'accepted', characters: 'textOf' })).join('\r');
}

/** A row's text as Office JS reports it: its cells with a tab between them, then `\r\n` (checks 27 and 28). */
export function rowText(tr: wml.Tr): string {
  return cellsOf(tr, 'read').map((cell) => cellText(cell.element.value)).join('\t') + '\r\n';
}

/**
 * The text of a table's change: the table's where its own properties or grid changed, else the changed
 * rows', read in the selected branch, what Word shows: so a change whose records are all in other
 * branches has no text (docx4j's `Tables.changeText`, its `aTableChangesTextReadsTheBranchWordReads`).
 */
export function tableChangeText(tbl: wml.Tbl): string {
  const rows = rowsOf(tbl, 'read').map((row) => row.element.value);
  const shown = tableDiffers(tbl) ? rows : rows.filter((tr) => inChange(tr) && rowDiffers(tr));
  return shown.map(rowText).join('');
}

/** Accepting keeps the properties as they are and drops every record, in every branch; a row's `w:trPr` left empty goes. */
export function acceptTable(tbl: wml.Tbl): void {
  delete tbl.tblPr?.tblPrChange;
  delete tbl.tblGrid?.tblGridChange;
  for (const tr of allRows(tbl)) {
    delete tr.trPr?.trPrChange;
    delete tr.tblPrEx?.tblPrExChange;
    pruneRow(tr);
    for (const tc of allCells(tr)) delete tc.tcPr?.tcPrChange;
  }
}

/**
 * Rejecting puts the recorded table back: its properties, its grid, and each row and cell in the
 * change as recorded - a row in the change without a `w:trPrChange` cleared, having had no properties
 * (check 28, from a saved file) - then collapses a merged grid (check 27: a width change records the
 * old cells over a grid of the old and new column edges together). Rows outside the change are left.
 * Every branch of a kept row- or cell-level `mc:AlternateContent` is restored (CR-013 phase 3).
 */
export function rejectTable(tbl: wml.Tbl): void {
  const tableChange = tbl.tblPr?.tblPrChange;
  if (tableChange && tbl.tblPr) replaceProperties(tbl.tblPr, tableChange.tblPr, [], tbl);
  const gridChange = tbl.tblGrid?.tblGridChange;
  if (gridChange && tbl.tblGrid) replaceProperties(tbl.tblGrid, gridChange.tblGrid, [], tbl);
  for (const tr of allRows(tbl)) {
    if (!inChange(tr)) continue;
    const exChange = tr.tblPrEx?.tblPrExChange;
    if (exChange && tr.tblPrEx) replaceProperties(tr.tblPrEx, exChange.tblPrEx, [], tr);
    if (tr.trPr) replaceProperties(tr.trPr, tr.trPr.trPrChange?.trPr, ROW_MARKERS, tr);
    pruneRow(tr);
    for (const tc of allCells(tr)) {
      const tcPr = tc.tcPr;
      if (tcPr?.tcPrChange) replaceProperties(tcPr, tcPr.tcPrChange.tcPr, CELL_MARKERS, tc);
    }
  }
  if (gridChange) collapseGrid(tbl);
}

/**
 * Replaces a properties object's members with a record's, keeping the object (and so its type, which
 * the marshaller writes by) and the members named in `keep`.
 */
function replaceProperties(target: object, recorded: object | undefined, keep: string[], parent: object): void {
  const t = target as Record<string, unknown>;
  for (const key of Object.keys(t)) if (key !== 'TYPE_NAME' && key !== 'PARENT' && !keep.includes(key)) delete t[key];
  for (const [key, value] of Object.entries(recorded ?? {})) {
    if (key === 'TYPE_NAME' || key === 'PARENT' || keep.includes(key)) continue;
    t[key] = deepCopy(value);
  }
  linkParents(target, parent);
}

/** A row's `w:trPr` with nothing left in it goes, as Word writes none. */
export function pruneRow(tr: wml.Tr): void {
  const trPr = tr.trPr;
  if (trPr && !trPr.ins && !trPr.del && !trPr.trPrChange && !(trPr.cnfStyleOrDivIdOrGridBefore?.length)) delete tr.trPr;
}

// --- a merged grid collapsed -------------------------------------------------------------------------

type RowMember = Element<{ val?: number | string }>;

function rowMember(tr: wml.Tr, name: 'gridBefore' | 'gridAfter'): RowMember | undefined {
  return (tr.trPr?.cnfStyleOrDivIdOrGridBefore ?? []).find((el) => (el as Element).name?.localPart === name) as RowMember | undefined;
}

function span(tc: wml.Tc): number {
  return Number(tc.tcPr?.gridSpan?.val ?? 1) || 1;
}

/**
 * Grid columns that no row keeps apart become one: the edges any row's cells (or its `w:gridBefore`
 * and `w:gridAfter`) fall on are kept, the others merged, and each cell's `w:gridSpan` and each row's
 * skipped columns counted again over the kept columns. The grid is every branch's: every row-level
 * branch's rows are counted, and a row's cells along each of its layouts (`cellLayoutsOf`, one per
 * reading of a cell-level `mc:AlternateContent`), every branch spanning the same columns (CR-013 phase 3,
 * docx4j's `Lister.cellLayoutsOf`; phase 2 counted every branch's cells along one row).
 */
export function collapseGrid(tbl: wml.Tbl): void {
  const cols = tbl.tblGrid?.gridCol ?? [];
  if (cols.length < 2) return;
  const rows = allRows(tbl);
  const layouts = new Map(rows.map((tr) => [tr, cellLayoutsOf(tr).map((layout) => layout.map((cell) => cell.element.value))]));
  const edges = new Set<number>([0, cols.length]);
  for (const tr of rows) {
    for (const layout of layouts.get(tr)!) {
      let at = Number(rowMember(tr, 'gridBefore')?.value.val ?? 0);
      edges.add(at);
      for (const tc of layout) { at += span(tc); edges.add(at); }
    }
  }
  const kept = [...edges].filter((edge) => edge >= 0 && edge <= cols.length).sort((a, b) => a - b);
  if (kept.length - 1 === cols.length) return;
  const index = (edge: number): number => kept.indexOf(edge);
  const widths = kept.slice(1).map((edge, i) => cols.slice(kept[i]!, edge).reduce((sum, col) => sum + Number(col.w ?? 0), 0));
  tbl.tblGrid!.gridCol = widths.map((w, i) => ({ ...cols[Math.min(kept[i]!, cols.length - 1)]!, w }));
  linkParents(tbl.tblGrid!, tbl);
  for (const tr of rows) {
    const before = rowMember(tr, 'gridBefore');
    const start = Number(before?.value.val ?? 0);
    // the new spans counted along each layout, last first, a cell outside the elements once; the row's
    // skipped columns after from the selected branches' layout (the first)
    const spans = new Map<wml.Tc, number>();
    let end = start;
    const all = layouts.get(tr)!;
    for (let l = all.length - 1; l >= 0; l--) {
      let at = start;
      for (const tc of all[l]!) {
        const next = at + span(tc);
        spans.set(tc, index(next) - index(at));
        at = next;
      }
      end = at;
    }
    for (const [tc, count] of spans) {
      if (!tc.tcPr) continue;
      if (count > 1) tc.tcPr.gridSpan = { ...(tc.tcPr.gridSpan ?? {}), val: count } as wml.TcPrInner.GridSpan;
      else delete tc.tcPr.gridSpan;
    }
    if (before) before.value.val = index(start);
    const after = rowMember(tr, 'gridAfter');
    if (after) after.value.val = kept.length - 1 - index(end);
  }
}
