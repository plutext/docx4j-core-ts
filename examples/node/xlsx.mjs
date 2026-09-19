// A workbook from nothing: create, add two sheets, save.
//
//   node examples/node/xlsx.mjs [out.xlsx]
//
// SpreadsheetML has no content API either; a sheet's rows are `sheet.contents.sheetData.row`
// from the object model, built with the generated factories.
//
// Published, the import is `from '@docx4j/core-ts'`; here it is the built dist of this checkout.
import { writeFile } from 'node:fs/promises';
import { SpreadsheetMLPackage } from '../../dist/index.mjs';
// The generated factories of the object model; never write qualified names by hand.
import { createRow, createCell } from '@docx4j/generated-objects-ts/factory/org_xlsx4j_sml';

const out = process.argv[2] ?? 'hello.xlsx';

const pkg = await SpreadsheetMLPackage.createPackage();
const sheet = pkg.createWorksheetPart('Sales');
pkg.createWorksheetPart('Notes');

// Three rows of inline numbers; `t: 'n'` is the default, so only the value is needed.
const rows = [[1, 2], [3, 4], [5, 6]];
sheet.contents.sheetData.row = rows.map((values, r) => createRow({
  r: r + 1,
  c: values.map((v, c) => createCell({ r: `${String.fromCharCode(65 + c)}${r + 1}`, v: String(v) })),
}));

console.log(`sheets: ${pkg.getWorkbookPart().contents.sheets.sheet.map((s) => s.name).join(', ')}`);
console.log(`rows in ${sheet.partName.name}: ${sheet.contents.sheetData.row.length}`);

await writeFile(out, await pkg.save());
console.log(`wrote ${out}`);
