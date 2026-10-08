// The files for Word check 40 (CR-002 section 42.5's three open points): what Word 365 writes for a
// column inserted into an AutoFit-to-window table whose columns are unequal, for Insert Right beside
// and inside a cell spanning columns, and what Office JS does to a column deleted and one added
// while changes are tracked. Run from the repository root after `npm run build`:
//
//   node test/fixtures/check40/build.mjs
//
// writes into this directory (README.md says what to do with each):
//   40a1-window-unequal.docx   one AutoFit-to-window table, columns 3,000 and 6,026 twips: Insert Right in the narrow first cell
//   40a2-window-unequal.docx   the same table: Insert Left in the wide last cell
//   40b1-span.docx             a fixed table, row 1 a cell spanning columns 1 and 2 then a cell, row 2 three cells: Insert Right in row 2's first cell
//   40b2-span.docx             the same table: Insert Right in the spanning cell
//   40c-officejs.docx          two fixed 3 x 2 tables for the Script Lab snippet (check40-script.js): a column deleted and one added under tracking
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const TEXT_WIDTH = 9026;

const cell = (text, tcPr) => `<w:tc><w:tcPr>${tcPr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (cells) => `<w:tr>${cells.join('')}</w:tr>`;
const pct = (twips) => Math.floor((5000 * twips) / TEXT_WIDTH);

/** An AutoFit-to-window table of two unequal columns, 3,000 and 6,026 twips (1661 and 3339 pct). */
function windowUnequal() {
  const widths = [3000, 6026];
  const shares = [pct(3000), 5000 - pct(3000)];
  const r = (prefix) => row(widths.map((_, i) => cell(`${prefix}${'AB'[i]}`, `<w:tcW w:w="${shares[i]}" w:type="pct"/>`)));
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${r('W1')}${r('W2')}</w:tbl>`;
}

/** A fixed table of three equal columns; row 1 a cell spanning the first two then a cell, row 2 three cells. */
function spanTable() {
  const w = 3009;
  const grid = `<w:tblGrid><w:gridCol w:w="${w}"/><w:gridCol w:w="${w}"/><w:gridCol w:w="${w - 1}"/></w:tblGrid>`;
  const row1 = row([cell('S1AB (spans columns 1 and 2)', `<w:tcW w:w="${2 * w}" w:type="dxa"/><w:gridSpan w:val="2"/>`), cell('S1C', `<w:tcW w:w="${w - 1}" w:type="dxa"/>`)]);
  const row2 = row([cell('S2A', `<w:tcW w:w="${w}" w:type="dxa"/>`), cell('S2B', `<w:tcW w:w="${w}" w:type="dxa"/>`), cell('S2C', `<w:tcW w:w="${w - 1}" w:type="dxa"/>`)]);
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${TEXT_WIDTH}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>${grid}${row1}${row2}</w:tbl>`;
}

/** A fixed 3 x 2 table of equal columns with a letter per cell. */
function fixedTable(prefix) {
  const w = 3009;
  const grid = `<w:tblGrid><w:gridCol w:w="${w}"/><w:gridCol w:w="${w}"/><w:gridCol w:w="${w - 1}"/></w:tblGrid>`;
  const r = (n) => row(['A', 'B', 'C'].map((c, i) => cell(`${prefix}${n}${c}`, `<w:tcW w:w="${i === 2 ? w - 1 : w}" w:type="dxa"/>`)));
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${TEXT_WIDTH}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>${grid}${r(1)}${r(2)}</w:tbl>`;
}

async function document(heading, blocks) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  body.insertParagraph(heading, 'End');
  for (const block of blocks) {
    if (block.startsWith('<w:tbl>')) { await body.insertXml(block, 'End'); body.insertParagraph('', 'End'); }
    else body.insertParagraph(block, 'End');
  }
  return pkg.save();
}

await writeFile(join(here, '40a1-window-unequal.docx'), await document('Check 40a1: AutoFit to window, columns 3000 and 6026 twips (1661 and 3339 pct). Put the caret in W1A (the narrow first cell), Table Layout > Insert Right, Save As 40a1-word365.docx.', [windowUnequal()]));
await writeFile(join(here, '40a2-window-unequal.docx'), await document('Check 40a2: the same table. Put the caret in W1B (the wide last cell), Table Layout > Insert Left, Save As 40a2-word365.docx.', [windowUnequal()]));
await writeFile(join(here, '40b1-span.docx'), await document('Check 40b1: fixed, three equal columns; row 1 has a cell spanning columns 1 and 2. Put the caret in S2A (row 2, first cell), Table Layout > Insert Right, Save As 40b1-word365.docx.', [spanTable()]));
await writeFile(join(here, '40b2-span.docx'), await document('Check 40b2: the same table. Put the caret in S1AB (the spanning cell), Table Layout > Insert Right, Save As 40b2-word365.docx.', [spanTable()]));
await writeFile(join(here, '40c-officejs.docx'), await document('Check 40c: for the Script Lab snippet check40-script.js. Table 1 (D) loses its middle column, table 2 (E) gains one at the end, both with tracking on, through Office JS. Then Save As 40c-word365.docx.', [fixedTable('D'), 'Table 2:', fixedTable('E')]));
console.log('written the five files in', here);
