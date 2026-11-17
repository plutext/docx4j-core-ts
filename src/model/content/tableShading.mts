// CR-002 section 67: Office JS's `shadingColor` on `Table`, `TableRow` and `TableCell`, as Word 365 writes and reads it.
// Measured by the portfolio's Word bridge round core-ts-table-shading (Word 16.0.20326.20318, section 67.3, its files in
// test/table-shading/) and decided as section 67.3's E1 to E5 (Jason, 2026-11-18):
// - a write is `w:shd w:val="clear" w:color="RRGGBB" w:fill="RRGGBB"` on each cell written, the theme attributes gone;
//   a row writes each of its `w:tc`s, a merge's continuation among them, never a restart in another row; a table writes its
//   `w:tblPr/w:shd` and every cell's (E2, E3);
// - a clear (`""`, `null`), which raises WD-004's false dialog in Word, takes the cells' own `w:shd` away, and a table's
//   own too, so a table style's shading shows again (E1, a departure);
// - tracked, each cell written records its old `w:tcPrChange`, a table its `w:tblPrChange`; Word's records of no change
//   (the unchanged grid, an empty `w:tblPrExChange`, an unchanged neighbour) are left out (E4, a departure);
// - a getter reads what Word draws: the cell's own shading, else its table style's through the conditions it is under, a
//   theme colour as drawn, never the table's own `w:tblPr/w:shd`; '' for none, null where a row's or table's cells differ,
//   a merge's continuation not among a row's cells (E5).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { createCTShd } from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { WORD_COLOR_NAMES } from '@docx4j/generated-objects-ts/helpers/wml';
import { linkParents, rowsOf, cellsOf } from './tree.mjs';
import type { ChangeTracker } from './tracking.mjs';
import { InvalidArgument } from './generalException.mjs';
import type { PropertyResolver } from '../properties/PropertyResolver.mjs';
import type { TableContext } from '../properties/table.mjs';

/**
 * CSS's colour names, which Office JS's `shadingColor` takes in any case (`yellow`, `Red` and `lightgreen` measured, ts03):
 * the objects package's 140 that Word reads in a colour value, and the seven `grey` spellings CSS adds, as its
 * `highlightNameForColor` takes them.
 */
const CSS_COLOR_NAMES: Readonly<Record<string, string>> = Object.freeze(Object.assign(Object.create(null) as Record<string, string>,
  WORD_COLOR_NAMES, {
    darkgrey: 'A9A9A9', darkslategrey: '2F4F4F', dimgrey: '696969', grey: '808080', lightgray: 'D3D3D3',
    lightslategrey: '778899', slategrey: '708090',
  }));

/**
 * The colour a `shadingColor` write names, as `RRGGBB`, or undefined for a clear (`""` or `null`). Word takes `#RRGGBB`
 * and six bare hex digits in either case, and a CSS name in any case; anything else throws `InvalidArgument` (ts02,
 * ts03), before anything changes.
 */
export function shadingColorOf(value: unknown, member: string): string | undefined {
  if (value === null || value === undefined || value === '') return undefined;
  if (typeof value === 'string') {
    const hex = /^#?([0-9A-Fa-f]{6})$/.exec(value);
    if (hex) return hex[1]!.toUpperCase();
    const named = CSS_COLOR_NAMES[value.toLowerCase()];
    if (named !== undefined) return named;
  }
  throw new InvalidArgument(`${member}: ${JSON.stringify(value)} is not a colour; give '#RRGGBB' or a colour name`);
}

/** Whether a shading is already Word's form of this colour, or absent for a clear, so that writing it changes nothing. */
function holds(shd: wml.CTShd | undefined, color: string | undefined): boolean {
  if (color === undefined) return shd === undefined;
  return shd !== undefined && shd.val === 'clear' && shd.color === color && shd.fill === color
    && shd.themeColor === undefined && shd.themeTint === undefined && shd.themeShade === undefined
    && shd.themeFill === undefined && shd.themeFillTint === undefined && shd.themeFillShade === undefined;
}

/** Word's shading of a colour, linked to its holder. */
function shadingOf(color: string, holder: object): wml.CTShd {
  const shd = createCTShd({ val: 'clear', color, fill: color });
  linkParents(shd, holder);
  return shd;
}

/** A cell given Word's shading for the colour, or its own taken away (`undefined`); tracked where `tracker` is set. */
export function shadeCell(tc: wml.Tc, color: string | undefined, tracker: ChangeTracker | undefined): void {
  if (holds(tc.tcPr?.shd, color)) return;
  const tcPr = (tc.tcPr ??= {} as wml.TcPr);
  tracker?.recordTcPrChange(tcPr);
  if (color === undefined) delete tcPr.shd;
  else tcPr.shd = shadingOf(color, tcPr);
  tracker?.settleTcPrChange(tcPr);
}

/** Every `w:tc` of a row, in every branch of a kept `mc:AlternateContent`, a merge's continuation among them (ts05). */
export function shadeRow(tr: wml.Tr, color: string | undefined, tracker: ChangeTracker | undefined): void {
  for (const cell of cellsOf(tr, 'every')) shadeCell(cell.element.value, color, tracker);
}

/** The table's own `w:tblPr/w:shd` and every cell of every row, as Word's `Table.shadingColor` writes them (ts08). */
export function shadeTable(tbl: wml.Tbl, color: string | undefined, tracker: ChangeTracker | undefined): void {
  if (!holds(tbl.tblPr?.shd, color)) {
    const tblPr = (tbl.tblPr ??= {} as wml.TblPr);
    tracker?.recordTblPrChange(tblPr);
    if (color === undefined) delete tblPr.shd;
    else tblPr.shd = shadingOf(color, tblPr);
    tracker?.settleTblPrChange(tblPr);
  }
  for (const row of rowsOf(tbl, 'every')) shadeRow(row.element.value, color, tracker);
}

/** Whether a cell is a vertical merge's continuation, which Office JS doesn't count among its row's cells (ts01). */
export function isContinuation(tc: wml.Tc): boolean {
  const vMerge = tc.tcPr?.vMerge;
  return vMerge !== undefined && vMerge !== null && vMerge.val !== 'restart';
}

/** A shading that states one: `w:val="nil"` is none, so a table style's shows through it (section 67.1). */
function stated(shd: wml.CTShd | undefined | null): wml.CTShd | undefined {
  return shd === undefined || shd === null || shd.val === 'nil' ? undefined : shd;
}

/**
 * The colour Word reads for a cell, `#RRGGBB`, or '' for none: its own shading, else its table style's (the style's own
 * `w:tcPr`, then each conditional format it is under, in the order they apply), a theme colour as Word draws it. The
 * table's own `w:tblPr/w:shd` is never read (ts01: T7's cells read '').
 */
export function cellShadingOf(resolver: PropertyResolver, context: TableContext, tr: wml.Tr, tc: wml.Tc): string {
  let shd = stated(tc.tcPr?.shd);
  if (shd === undefined) {
    const conditions = context.forCell(tr, tc).conditions;
    for (const level of [context.tableStyle.tcPr?.shd, ...context.applicable(conditions).map((pr) => pr.tcPr?.shd)]) {
      shd = stated(level) ?? shd;
    }
  }
  const drawn = resolver.resolveFill(shd);
  return drawn === undefined ? '' : `#${drawn.toHex()}`;
}

/** The colours read for a row's or a table's cells, as Office JS gives them: one colour where all agree, '' for none, null where they differ. */
export function agreed(colors: readonly string[]): string | null {
  if (colors.length === 0) return '';
  return colors.every((c) => c === colors[0]) ? colors[0]! : null;
}

/** The colours of a row's cells, a merge's continuation left out (ts01: T4's second row reads its one other cell). */
export function rowShadings(resolver: PropertyResolver, context: TableContext, tr: wml.Tr): string[] {
  return cellsOf(tr, 'read').map((c) => c.element.value).filter((tc) => !isContinuation(tc))
    .map((tc) => cellShadingOf(resolver, context, tr, tc));
}
