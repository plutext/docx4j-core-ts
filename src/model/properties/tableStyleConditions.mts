// docx4j `org.docx4j.model.table.TableStyleConditions` (CR-007; docx4j CR-030 and 17.2.0): a
// table style's conditional formatting (`w:tblStylePr`, ECMA-376-1 17.7.6) - which of its
// conditions a given row, cell or paragraph is under, and which of the style's `w:tblStylePr`
// entries therefore apply to it, in the order they are to be applied.
//
// Exported as the namespace `TableStyleConditions`, so the Java's static methods read the same
// (`TableStyleConditions.resolve`, `.applicable`, `.key`); `Look.of` is `lookOf`, `Look.DEFAULT`
// is `DEFAULT_LOOK`, `Look.allows` is `lookAllows`, and an `EnumSet<STTblStyleOverrideType>` is
// a `Set` of the schema's strings.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { log } from './log.mjs';
import { isEmptyPPr, isEmptyRPr } from './styleUtil.mjs';

export type Condition = wml.STTblStyleOverrideType;

/**
 * The order in which a table style's conditional formats are applied to a cell, later
 * overriding earlier (ECMA-376-1 17.7.6).
 */
export const PRECEDENCE: readonly Condition[] = [
  'wholeTable',
  'band1Vert', 'band2Vert',
  'band1Horz', 'band2Horz',
  'firstCol', 'lastCol',
  'firstRow', 'lastRow',
  'nwCell', 'neCell',
  'swCell', 'seCell',
];

/** The conditions `w:cnfStyle` carries on a row: the row-axis bits. */
const ROW_AXIS: ReadonlySet<Condition> = new Set<Condition>(['firstRow', 'lastRow', 'band1Horz', 'band2Horz']);
/** The conditions `w:cnfStyle` carries on a cell: the column-axis bits. */
const COLUMN_AXIS: ReadonlySet<Condition> = new Set<Condition>(['firstCol', 'lastCol', 'band1Vert', 'band2Vert']);

/** Which conditional formats a table asks for: `w:tblPr/w:tblLook`, resolved. */
export interface Look {
  readonly firstRow: boolean;
  readonly lastRow: boolean;
  readonly firstColumn: boolean;
  readonly lastColumn: boolean;
  /** Horizontal (row) banding is on: `w:noHBand` is off. */
  readonly hBand: boolean;
  /** Vertical (column) banding is on: `w:noVBand` is off. */
  readonly vBand: boolean;
}

/**
 * What Word assumes where a table states no `w:tblLook` at all: its own default, `04A0` - first
 * row, first column and row banding on, column banding off.
 */
export const DEFAULT_LOOK: Look = Object.freeze({
  firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, hBand: true, vBand: false,
});

function on(v: string | undefined): boolean {
  return v === '1' || v === 'true' || v === 'on';
}

/**
 * The look a `w:tblLook` states. The six attributes win where any of them is present; otherwise
 * the legacy `w:val` bitmask (0x0020 firstRow, 0x0040 lastRow, 0x0080 firstColumn, 0x0100
 * lastColumn, 0x0200 noHBand, 0x0400 noVBand); none, or an empty element, is {@link DEFAULT_LOOK}.
 */
export function lookOf(tblLook: wml.CTTblLook | undefined): Look {
  if (tblLook === undefined) return DEFAULT_LOOK;
  if (tblLook.firstRow !== undefined || tblLook.lastRow !== undefined
    || tblLook.firstColumn !== undefined || tblLook.lastColumn !== undefined
    || tblLook.noHBand !== undefined || tblLook.noVBand !== undefined) {
    return {
      firstRow: on(tblLook.firstRow), lastRow: on(tblLook.lastRow),
      firstColumn: on(tblLook.firstColumn), lastColumn: on(tblLook.lastColumn),
      hBand: !on(tblLook.noHBand), vBand: !on(tblLook.noVBand),
    };
  }
  const val = tblLook.val;
  if (val === undefined || val.trim().length === 0) return DEFAULT_LOOK;
  const text = val.trim();
  // Integer.parseInt(s, 16): an optional sign, then hex digits only
  if (!/^[+-]?[0-9a-fA-F]+$/.test(text)) {
    log.warn(`Unreadable w:tblLook/@w:val '${val}'; using Word's default`);
    return DEFAULT_LOOK;
  }
  const bits = parseInt(text, 16);
  return {
    firstRow: (bits & 0x0020) !== 0, lastRow: (bits & 0x0040) !== 0,
    firstColumn: (bits & 0x0080) !== 0, lastColumn: (bits & 0x0100) !== 0,
    hBand: (bits & 0x0200) === 0, vBand: (bits & 0x0400) === 0,
  };
}

/** Whether a look lets the condition through at all. */
export function lookAllows(look: Look, t: Condition): boolean {
  switch (t) {
    case 'wholeTable': return true;
    case 'firstRow': return look.firstRow;
    case 'lastRow': return look.lastRow;
    case 'firstCol': return look.firstColumn;
    case 'lastCol': return look.lastColumn;
    case 'band1Horz': case 'band2Horz': return look.hBand;
    case 'band1Vert': case 'band2Vert': return look.vBand;
    case 'nwCell': return look.firstRow && look.firstColumn;
    case 'neCell': return look.firstRow && look.lastColumn;
    case 'swCell': return look.lastRow && look.firstColumn;
    case 'seCell': return look.lastRow && look.lastColumn;
    default: return false;
  }
}

function atLeastOne(v: number | undefined): number {
  return v === undefined || !(v >= 1) ? 1 : Math.trunc(v);
}

/**
 * `w:tblStyleRowBandSize`, at least 1 where it is stated, and 0 where it is not: Word bands no
 * row of a table whose style chain and own `w:tblPr` state no band size, whatever its `w:tblLook`
 * and the style's band conditions (measured, docx4j CR-030 probes T9 and T10, D11: Word's re-save
 * writes no band bit, and its PDF has no band formatting; a size stated by the table alone is
 * enough). Until docx4j 0349796f9 an absent size was taken as 1.
 */
export function rowBandSize(tblPr: wml.CTTblPrBase | undefined): number {
  return tblPr?.tblStyleRowBandSize === undefined ? 0 : atLeastOne(tblPr.tblStyleRowBandSize.val);
}

/**
 * `w:tblStyleColBandSize`, as {@link rowBandSize} for the columns: 0, no vertical banding, where
 * it is not stated. (By the rows' rule; the columns themselves were not probed in Word.)
 */
export function colBandSize(tblPr: wml.CTTblPrBase | undefined): number {
  return tblPr?.tblStyleColBandSize === undefined ? 0 : atLeastOne(tblPr.tblStyleColBandSize.val);
}

/** The bit order of a `w:cnfStyle` value (ECMA-376-1 17.18.6). */
const CNF_BITS: readonly Condition[] = [
  'firstRow', 'lastRow', 'firstCol', 'lastCol',
  'band1Vert', 'band2Vert', 'band1Horz', 'band2Horz',
  'nwCell', 'neCell', 'swCell', 'seCell',
];

/**
 * The conditions a `w:cnfStyle` bitmask names: twelve characters, firstRow, lastRow,
 * firstColumn, lastColumn, oddVBand, evenVBand, oddHBand, evenHBand, then the corners nw, ne,
 * sw, se. A short string is read as far as it goes; none is the empty set.
 */
export function fromCnf(cnf: wml.CTCnf | string | undefined): Set<Condition> {
  const out = new Set<Condition>();
  const val = typeof cnf === 'string' ? cnf : cnf?.val;
  if (val === undefined || val === null) return out;
  const s = String(val).trim();
  for (let i = 0; i < CNF_BITS.length && i < s.length; i++) {
    if (s.charAt(i) === '1') out.add(CNF_BITS[i]!);
  }
  return out;
}

type TrPrItem = { name?: { localPart?: string }; value?: unknown };

/** The `w:cnfStyle` of a row, or undefined where the row states none. */
export function rowCnf(trPr: wml.CTTrPrBase | undefined): wml.CTCnf | undefined {
  for (const item of (trPr?.cnfStyleOrDivIdOrGridBefore ?? []) as TrPrItem[]) {
    if (item?.name?.localPart === 'cnfStyle' && typeof item.value === 'object' && item.value !== null) {
      return item.value as wml.CTCnf;
    }
  }
  return undefined;
}

function addCorners(s: Set<Condition>): void {
  const fr = s.has('firstRow');
  const lr = s.has('lastRow');
  const fc = s.has('firstCol');
  const lc = s.has('lastCol');
  if (fr && fc) s.add('nwCell');
  if (fr && lc) s.add('neCell');
  if (lr && fc) s.add('swCell');
  if (lr && lc) s.add('seCell');
}

/** Remove the conditions the look does not allow (in place), and return the set. */
export function gate(conditions: Set<Condition>, look: Look = DEFAULT_LOOK): Set<Condition> {
  for (const t of [...conditions]) if (!lookAllows(look, t)) conditions.delete(t);
  return conditions;
}

/**
 * The conditions a cell is under by its position alone, with no `w:cnfStyle` to go by.
 *
 * Banding counts from the first row (column) which is not under a condition of its own: where
 * the look has `firstRow` on, row 1 is the first banded row; a last row (column) the look gives
 * its own condition is left out of the bands too. A band is `w:tblStyleRowBandSize` rows deep,
 * and a size of 0 (none stated anywhere, see {@link rowBandSize}) bands nothing.
 * A cell spanning several columns is placed by its first column, and is in the last column
 * where its span reaches it.
 */
export function atPosition(look: Look, rowBand: number, colBand: number,
  row: number, rowCount: number, col: number, colSpan: number, colCount: number): Set<Condition> {
  const out = new Set<Condition>();
  const first = look.firstRow && row === 0;
  const last = look.lastRow && rowCount > 0 && row === rowCount - 1;
  if (first) out.add('firstRow');
  if (last) out.add('lastRow');
  if (look.hBand && rowBand > 0 && !first && !last) {
    const offset = look.firstRow ? 1 : 0;
    const band = Math.trunc((row - offset) / Math.max(1, rowBand));
    out.add(band % 2 === 0 ? 'band1Horz' : 'band2Horz');
  }
  const span = Math.max(1, colSpan);
  const firstCol = look.firstColumn && col === 0;
  const lastCol = look.lastColumn && colCount > 0 && col + span - 1 >= colCount - 1;
  if (firstCol) out.add('firstCol');
  if (lastCol) out.add('lastCol');
  if (look.vBand && colBand > 0 && !firstCol && !lastCol) {
    const offset = look.firstColumn ? 1 : 0;
    const band = Math.trunc((col - offset) / Math.max(1, colBand));
    out.add(band % 2 === 0 ? 'band1Vert' : 'band2Vert');
  }
  addCorners(out);
  return out;
}

function axis(bits: Set<Condition>, para: Set<Condition>, keep: ReadonlySet<Condition>): Set<Condition> {
  const out = new Set<Condition>();
  for (const t of bits) if (keep.has(t)) out.add(t);
  for (const t of para) if (keep.has(t)) out.add(t);
  return out;
}

/**
 * The conditions a cell (or a paragraph in it) is under: the `w:cnfStyle` caches where the
 * document has them, the position where it has not, gated by the look (docx4j
 * `TableStyleConditions.resolve`).
 *
 * Word writes the row-axis bits on the `w:tr`, the column-axis bits on the `w:tc` and the
 * row-axis bits again on each `w:p`, and never the corner bits. So the row's cache, where there
 * is one, decides the row axis, the cell's the column axis, a paragraph's bits are added on
 * whichever axis they name, and the corners are derived from the result. An axis with no cache
 * falls back to {@link atPosition}. Every bit is then gated by the look, so a stale cache cannot
 * switch on a format the table has turned off.
 */
export function resolve(look: Look, rowBand: number, colBand: number,
  row: number, rowCount: number, col: number, colSpan: number, colCount: number,
  rowCnfStyle: wml.CTCnf | undefined, cellCnf: wml.CTCnf | undefined, paraCnf: wml.CTCnf | undefined): Set<Condition> {
  const computed = atPosition(look, rowBand, colBand, row, rowCount, col, colSpan, colCount);
  const para = fromCnf(paraCnf);
  const out = new Set<Condition>();

  const rowBits = axis(fromCnf(rowCnfStyle), para, ROW_AXIS);
  if (rowCnfStyle !== undefined || (paraCnf !== undefined && rowBits.size > 0)) {
    for (const t of rowBits) out.add(t);
  } else {
    for (const t of computed) if (ROW_AXIS.has(t)) out.add(t);
  }

  const colBits = axis(fromCnf(cellCnf), para, COLUMN_AXIS);
  if (cellCnf !== undefined || (paraCnf !== undefined && colBits.size > 0)) {
    for (const t of colBits) out.add(t);
  } else {
    for (const t of computed) if (COLUMN_AXIS.has(t)) out.add(t);
  }

  gate(out, look);
  addCorners(out);
  return out;
}

/**
 * The row-axis conditions a row is under - first row, last row, or a horizontal band - for the
 * row's own properties: its `w:cnfStyle` where it has one, else its position, gated by the look.
 */
export function rowConditions(look: Look, rowBand: number, row: number, rowCount: number,
  rowCnfStyle: wml.CTCnf | undefined): Set<Condition> {
  const from = rowCnfStyle === undefined
    ? atPosition(look, rowBand, 1, row, rowCount, 1, 1, 3)
    : fromCnf(rowCnfStyle);
  const out = new Set<Condition>();
  for (const t of from) if (ROW_AXIS.has(t)) out.add(t);
  return gate(out, look);
}

/**
 * The style's `w:tblStylePr` entry for a condition, or undefined. Pass the merged table style
 * (`PropertyResolver.getTableStyleChain`) to see inherited conditions.
 */
export function lookup(tableStyle: wml.Style | undefined, type: Condition): wml.CTTblStylePr | undefined {
  return tableStyle?.tblStylePr?.find((pr) => pr !== undefined && pr !== null && pr.type === type);
}

/**
 * The style's `w:tblStylePr` entries which apply under these conditions, in the order they are
 * to be applied ({@link PRECEDENCE}); a `wholeTable` entry, where the style has one, always
 * applies and comes first.
 */
export function applicable(tableStyle: wml.Style | undefined, conditions: ReadonlySet<Condition> | undefined): wml.CTTblStylePr[] {
  const out: wml.CTTblStylePr[] = [];
  if (tableStyle?.tblStylePr === undefined || tableStyle.tblStylePr.length === 0) return out;
  for (const t of PRECEDENCE) {
    if (t !== 'wholeTable' && (conditions === undefined || !conditions.has(t))) continue;
    const pr = lookup(tableStyle, t);
    if (pr !== undefined) out.push(pr);
  }
  return out;
}

/** Whether the entry has any paragraph or run formatting to contribute. */
export function formatsText(pr: wml.CTTblStylePr | undefined): boolean {
  return pr !== undefined && (!isEmptyPPr(pr.pPr) || !isEmptyRPr(pr.rPr));
}

/**
 * A short name for a set of conditions, in precedence order - `firstCol-firstRow-nwCell` - the
 * form docx4j's goldens and synthetic style ids use; the empty string for none.
 */
export function key(conditions: ReadonlySet<Condition> | undefined): string {
  if (conditions === undefined || conditions.size === 0) return '';
  return PRECEDENCE.filter((t) => t !== 'wholeTable' && conditions.has(t)).join('-');
}
