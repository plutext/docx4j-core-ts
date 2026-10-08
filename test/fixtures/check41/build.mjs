// The files for Word check 41 (CR-002 section 42.6's two unmeasured points): whose properties a new
// column's cells take when the two neighbouring columns differ, and what Office JS writes and lists
// for a column added, tracking on, to an AutoFit-to-contents table and an AutoFit-to-window one.
// Run from the repository root after `npm run build`:
//
//   node test/fixtures/check41/build.mjs
//
// writes into this directory (README.md says what to do with each):
//   41a-properties.docx   a fixed two-column table whose columns differ in every way a cell and its paragraph can:
//                         Insert Right in the first cell
//   41b-properties.docx   the same table: Insert Left in the last cell
//   41c-officejs.docx     an AutoFit-to-contents table (F) and an AutoFit-to-window one (G) for the Script Lab snippet
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const TEXT_WIDTH = 9026;

/** Column 1: yellow, centred vertically, a thick red left border, a centred bold paragraph. Column 2: blue, bottom, a dashed green right border, a right-aligned italic paragraph. */
const LEFT_TCPR = '<w:tcW w:w="4513" w:type="dxa"/><w:tcBorders><w:left w:val="single" w:sz="24" w:space="0" w:color="FF0000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:vAlign w:val="center"/>';
const RIGHT_TCPR = '<w:tcW w:w="4513" w:type="dxa"/><w:tcBorders><w:right w:val="dashed" w:sz="12" w:space="0" w:color="00AA00"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="99CCFF"/><w:vAlign w:val="bottom"/>';
const LEFT_P = (text) => `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const RIGHT_P = (text) => `<w:p><w:pPr><w:jc w:val="right"/><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

function propertiesTable() {
  const row = (n) => `<w:tr><w:tc><w:tcPr>${LEFT_TCPR}</w:tcPr>${LEFT_P(`P${n}A`)}</w:tc><w:tc><w:tcPr>${RIGHT_TCPR}</w:tcPr>${RIGHT_P(`P${n}B`)}</w:tc></w:tr>`;
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${TEXT_WIDTH}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid>${row(1)}${row(2)}</w:tbl>`;
}

const cell = (text, tcPr) => `<w:tc><w:tcPr>${tcPr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (cells) => `<w:tr>${cells.join('')}</w:tr>`;

/** An AutoFit-to-contents table (w:tblW auto, every w:tcW auto), 3 x 2, with a letter per cell. */
function contentsTable(prefix) {
  const grid = `<w:tblGrid><w:gridCol w:w="3009"/><w:gridCol w:w="3009"/><w:gridCol w:w="3008"/></w:tblGrid>`;
  const r = (n) => row(['A', 'B', 'C'].map((c) => cell(`${prefix}${n}${c}`, '<w:tcW w:w="0" w:type="auto"/>')));
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>${grid}${r(1)}${r(2)}</w:tbl>`;
}

/** An AutoFit-to-window table (w:tblW 5000 pct, cells in pct), 3 x 2. */
function windowTable(prefix) {
  const grid = `<w:tblGrid><w:gridCol w:w="3009"/><w:gridCol w:w="3009"/><w:gridCol w:w="3008"/></w:tblGrid>`;
  const r = (n) => row(['A', 'B', 'C'].map((c, i) => cell(`${prefix}${n}${c}`, `<w:tcW w:w="${i === 0 ? 1666 : 1667}" w:type="pct"/>`)));
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/></w:tblPr>${grid}${r(1)}${r(2)}</w:tbl>`;
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

await writeFile(join(here, '41a-properties.docx'), await document('Check 41a: column 1 is yellow, centred vertically, with a thick red left border and centred bold text; column 2 is blue, bottom-aligned, with a dashed green right border and right-aligned italic text. Put the caret in P1A, Table Layout > Insert Right, Save As 41a-word365.docx.', [propertiesTable()]));
await writeFile(join(here, '41b-properties.docx'), await document('Check 41b: the same table. Put the caret in P1B, Table Layout > Insert Left, Save As 41b-word365.docx.', [propertiesTable()]));
await writeFile(join(here, '41c-officejs.docx'), await document('Check 41c: for the Script Lab snippet check41-script.js. Table 1 (F) is AutoFit to contents, table 2 (G) AutoFit to window; each gains a column at the end through Office JS with tracking on. Then Save As 41c-word365.docx.', [contentsTable('F'), 'Table 2:', windowTable('G')]));
console.log('written the three files in', here);
