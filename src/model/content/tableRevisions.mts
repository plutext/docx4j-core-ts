// CR-002 section 35: a table's property revisions - w:tblPrChange, w:tblGridChange,
// w:tblPrExChange, w:trPrChange and w:tcPrChange - which Word writes as one set for one change and
// Office JS lists as one change per table, or per row where only a row changed (test/README.md
// checks 18, 27 and 28). Everything here works on the tree; TrackedChange.mts makes the views.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy } from '@docx4j/generated-objects-ts';
import { type Element, typeNameOf, rowsOf, cellsOf, childrenOf, linkParents, textOfView } from './tree.mjs';
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

/** The rows a change takes in: a row with a `w:trPrChange` or a `w:tblPrExChange`, or a cell with a `w:tcPrChange`. */
function inChange(tr: wml.Tr): boolean {
  return tr.trPr?.trPrChange !== undefined || tr.tblPrEx?.tblPrExChange !== undefined
    || cellsOf(tr).some((cell) => cell.element.value.tcPr?.tcPrChange !== undefined);
}

/**
 * True when a row in the change is not as its records have it: its properties (a row without a
 * `w:trPrChange` had none before, check 28), a cell's, or its `w:tblPrEx`.
 */
function rowDiffers(tr: wml.Tr): boolean {
  if (differs(without(tr.trPr, [...ROW_MARKERS, 'trPrChange']), tr.trPr?.trPrChange?.trPr)) return true;
  if (tr.tblPrEx?.tblPrExChange && differs(without(tr.tblPrEx, ['tblPrExChange']), tr.tblPrEx.tblPrExChange.tblPrEx)) return true;
  return cellsOf(tr).some(({ element }) => {
    const tcPr = element.value.tcPr;
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
  const rows = rowsOf(tbl).map((row) => row.element.value);
  const changed = rows.filter(inChange);
  if (!tbl.tblPr?.tblPrChange && !tbl.tblGrid?.tblGridChange && changed.length === 0) return undefined;
  // a set of records that holds no difference is no change (checks 24 and 26's rule)
  if (!tableDiffers(tbl) && !changed.some(rowDiffers)) return undefined;
  const value = tbl.tblPr?.tblPrChange
    ?? changed.map((tr) => tr.trPr?.trPrChange ?? tr.tblPrEx?.tblPrExChange ?? cellsOf(tr).map((c) => c.element.value.tcPr?.tcPrChange).find(Boolean)).find(Boolean);
  return value ? { value } : undefined;
}

/** True when the table has any property record at all, changing anything or not. */
export function hasTableRecords(tbl: wml.Tbl): boolean {
  return tbl.tblPr?.tblPrChange !== undefined || tbl.tblGrid?.tblGridChange !== undefined || rowsOf(tbl).some((row) => inChange(row.element.value));
}

/** A cell's text as Office JS reports it in a table's change: its paragraphs, a `\r` between them. */
function cellText(tc: wml.Tc): string {
  return (childrenOf(tc) ?? []).filter((el) => typeNameOf(el) === 'org_docx4j_wml.P').map((el) => textOfView(el.value as object, { view: 'accepted' })).join('\r');
}

/** A row's text as Office JS reports it: its cells with a tab between them, then `\r\n` (checks 27 and 28). */
export function rowText(tr: wml.Tr): string {
  return cellsOf(tr).map((cell) => cellText(cell.element.value)).join('\t') + '\r\n';
}

/** The text of a table's change: the table's where its own properties or grid changed, else the changed rows'. */
export function tableChangeText(tbl: wml.Tbl): string {
  const rows = rowsOf(tbl).map((row) => row.element.value);
  const shown = tableDiffers(tbl) ? rows : rows.filter((tr) => inChange(tr) && rowDiffers(tr));
  return shown.map(rowText).join('');
}

/** Accepting keeps the properties as they are and drops every record; a row's `w:trPr` left empty goes. */
export function acceptTable(tbl: wml.Tbl): void {
  delete tbl.tblPr?.tblPrChange;
  delete tbl.tblGrid?.tblGridChange;
  for (const { element } of rowsOf(tbl)) {
    const tr = element.value;
    delete tr.trPr?.trPrChange;
    delete tr.tblPrEx?.tblPrExChange;
    pruneRow(tr);
    for (const cell of cellsOf(tr)) delete cell.element.value.tcPr?.tcPrChange;
  }
}

/**
 * Rejecting puts the recorded table back: its properties, its grid, and each row and cell in the
 * change as recorded - a row in the change without a `w:trPrChange` cleared, having had no properties
 * (check 28, from a saved file) - then collapses a merged grid (check 27: a width change records the
 * old cells over a grid of the old and new column edges together). Rows outside the change are left.
 */
export function rejectTable(tbl: wml.Tbl): void {
  const tableChange = tbl.tblPr?.tblPrChange;
  if (tableChange && tbl.tblPr) replaceProperties(tbl.tblPr, tableChange.tblPr, [], tbl);
  const gridChange = tbl.tblGrid?.tblGridChange;
  if (gridChange && tbl.tblGrid) replaceProperties(tbl.tblGrid, gridChange.tblGrid, [], tbl);
  for (const { element } of rowsOf(tbl)) {
    const tr = element.value;
    if (!inChange(tr)) continue;
    const exChange = tr.tblPrEx?.tblPrExChange;
    if (exChange && tr.tblPrEx) replaceProperties(tr.tblPrEx, exChange.tblPrEx, [], tr);
    if (tr.trPr) replaceProperties(tr.trPr, tr.trPr.trPrChange?.trPr, ROW_MARKERS, tr);
    pruneRow(tr);
    for (const cell of cellsOf(tr)) {
      const tcPr = cell.element.value.tcPr;
      if (tcPr?.tcPrChange) replaceProperties(tcPr, tcPr.tcPrChange.tcPr, CELL_MARKERS, cell.element.value);
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
 * skipped columns counted again over the kept columns.
 */
export function collapseGrid(tbl: wml.Tbl): void {
  const cols = tbl.tblGrid?.gridCol ?? [];
  if (cols.length < 2) return;
  const rows = rowsOf(tbl).map((row) => row.element.value);
  const edges = new Set<number>([0, cols.length]);
  for (const tr of rows) {
    let at = Number(rowMember(tr, 'gridBefore')?.value.val ?? 0);
    edges.add(at);
    for (const cell of cellsOf(tr)) { at += span(cell.element.value); edges.add(at); }
  }
  const kept = [...edges].filter((edge) => edge >= 0 && edge <= cols.length).sort((a, b) => a - b);
  if (kept.length - 1 === cols.length) return;
  const index = (edge: number): number => kept.indexOf(edge);
  const widths = kept.slice(1).map((edge, i) => cols.slice(kept[i]!, edge).reduce((sum, col) => sum + Number(col.w ?? 0), 0));
  tbl.tblGrid!.gridCol = widths.map((w, i) => ({ ...cols[Math.min(kept[i]!, cols.length - 1)]!, w }));
  linkParents(tbl.tblGrid!, tbl);
  for (const tr of rows) {
    const before = rowMember(tr, 'gridBefore');
    let at = Number(before?.value.val ?? 0);
    if (before) before.value.val = index(at);
    for (const cell of cellsOf(tr)) {
      const tc = cell.element.value;
      const next = at + span(tc);
      const count = index(next) - index(at);
      if (tc.tcPr) {
        if (count > 1) tc.tcPr.gridSpan = { ...(tc.tcPr.gridSpan ?? {}), val: count } as wml.TcPrInner.GridSpan;
        else delete tc.tcPr.gridSpan;
      }
      at = next;
    }
    const after = rowMember(tr, 'gridAfter');
    if (after) after.value.val = kept.length - 1 - index(at);
  }
}
