// The files for Word check 72 (CR-016 D5: what Word 365 does with each misplaced child docx4j CR-064's rules refuse). Run
// from the repository root after `npm run build`:
//
//   node test/fixtures/check72/build.mjs
//
// writes one .docx per form into this directory, which git ignores (Word reads the share's copies); `check72Files` gives a
// test the same parts. The parts are written as text: the forms are what no builder here will make, and what a direct
// edit of the tree from JavaScript, or through `any`, saves today (CR-016 section 2).
//
// Each file holds one misplaced child, between a paragraph BEFORE and a paragraph AFTER, so that what Word drops or keeps
// shows against them. One form per file, since Word refuses a whole document for some faults (docx4j CR-065 round 1), and
// a refusal of a file holding several forms would say nothing about each. The forms are one per rule of CR-064 section 8's
// five sets, and the string in a run, whose save fails here with no cause given.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const t = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const cell = (inner, w = 3000) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const row = (...cells) => `<w:tr>${cells.join('')}</w:tr>`;
const tblPr = '<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
  + '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
  + '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr>';
const grid = (n) => `<w:tblGrid>${'<w:gridCol w:w="3000"/>'.repeat(n)}</w:tblGrid>`;
const table = (inner, columns = 2) => `<w:tbl>${tblPr}${grid(columns)}${inner}</w:tbl>`;
const plainRow = row(cell(para(t('CELL A1'))), cell(para(t('CELL A2'))));

/**
 * The forms, one file each: `name`, the CR-064 set and rule (`set`, `child`), what the file holds (`form`), and the body
 * content between BEFORE and AFTER (`body`). MISPLACED marks the misplaced child's text where it has one.
 */
export const FORMS = [
  { name: '72a-run-in-body', set: 'block-level container (body)', child: 'a run', form: 'a w:r as a child of w:body', body: t('MISPLACED run in the body') },
  { name: '72b-text-in-body', set: 'block-level container (body)', child: 'bare text', form: 'a w:t as a child of w:body', body: '<w:t>MISPLACED text in the body</w:t>' },
  { name: '72c-cell-in-body', set: 'block-level container (body)', child: 'a cell', form: 'a w:tc as a child of w:body', body: cell(para(t('MISPLACED cell in the body'))) },
  { name: '72d-row-in-body', set: 'block-level container (body)', child: 'a row', form: 'a w:tr as a child of w:body', body: row(cell(para(t('MISPLACED row in the body')))) },
  { name: '72e-break-in-body', set: 'block-level container (body)', child: 'a break', form: 'a w:br as a child of w:body', body: '<w:br/>' },
  { name: '72f-run-in-cell', set: 'block-level container (cell)', child: 'a run', form: 'a w:r as a child of w:tc, beside its paragraph', body: table(row(cell(t('MISPLACED run in a cell') + para(t('CELL A1'))), cell(para(t('CELL A2'))))) },
  { name: '72g-table-in-paragraph', set: 'run-level container (paragraph)', child: 'a table', form: 'a w:tbl inside a w:p', body: para(t('paragraph holding ') + table(plainRow) + t(' a table')) },
  { name: '72h-paragraph-in-paragraph', set: 'run-level container (paragraph)', child: 'a paragraph', form: 'a w:p inside a w:p', body: para(t('outer paragraph holding ') + para(t('MISPLACED inner paragraph')) + t(' then more')) },
  { name: '72i-text-in-paragraph', set: 'run-level container (paragraph)', child: 'bare text', form: 'a w:t inside a w:p, outside any run', body: para(t('a run, then ') + '<w:t>MISPLACED bare text</w:t>' + t(', then a run')) },
  { name: '72j-break-in-paragraph', set: 'run-level container (paragraph)', child: 'a break', form: 'a w:br inside a w:p, outside any run', body: para(t('a run, then a bare break') + '<w:br/>' + t('then a run')) },
  { name: '72k-run-in-run', set: 'run (r)', child: 'a run', form: 'a w:r inside a w:r', body: para(`<w:r>${t('MISPLACED inner run')}<w:t xml:space="preserve"> outer text</w:t></w:r>`) },
  { name: '72l-paragraph-in-run', set: 'run (r)', child: 'a paragraph', form: 'a w:p inside a w:r', body: para(`<w:r><w:t xml:space="preserve">outer text </w:t>${para(t('MISPLACED paragraph in a run'))}</w:r>`) },
  { name: '72m-string-in-run', set: 'run (r)', child: 'a string', form: 'character data directly in a w:r, outside w:t', body: para('<w:r>MISPLACED string in a run</w:r>') },
  { name: '72n-cell-in-table', set: 'table (tbl)', child: 'a cell', form: 'a w:tc as a child of w:tbl, beside its row', body: table(cell(para(t('MISPLACED cell in a table'))) + plainRow) },
  { name: '72o-paragraph-in-table', set: 'table (tbl)', child: 'a paragraph', form: 'a w:p as a child of w:tbl, beside its row', body: table(para(t('MISPLACED paragraph in a table')) + plainRow) },
  { name: '72p-table-in-table', set: 'table (tbl)', child: 'a table', form: 'a w:tbl as a child of w:tbl, beside its row', body: table(table(row(cell(para(t('MISPLACED nested table')))), 1) + plainRow) },
  { name: '72q-paragraph-in-row', set: 'row (tr)', child: 'a paragraph', form: 'a w:p as a child of w:tr, beside its cells', body: table(row(para(t('MISPLACED paragraph in a row')) + cell(para(t('CELL A1'))) + cell(para(t('CELL A2'))))) },
  { name: '72r-run-in-row', set: 'row (tr)', child: 'a run', form: 'a w:r as a child of w:tr, beside its cells', body: table(row(t('MISPLACED run in a row') + cell(para(t('CELL A1'))) + cell(para(t('CELL A2'))))) },
];

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** A form's document part. */
export const documentOf = (f) => `${declaration}<w:document xmlns:w="${W}"><w:body>`
  + para(t(`Check 72, ${f.name}: ${f.form} (CR-064: ${f.child} in a ${f.set}).`))
  + para(t('BEFORE')) + f.body + para(t('AFTER')) + `${sectPr}</w:body></w:document>`;

/** The files, { name, set, child, form, bytes }. */
export async function check72Files() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  if (!entries['word/document.xml']) throw new Error('no word/document.xml in the created package');
  return FORMS.map((f) => ({ ...f, bytes: zipSync({ ...entries, 'word/document.xml': strToU8(documentOf(f)) }) }));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  for (const f of await check72Files()) await writeFile(join(here, `${f.name}.docx`), f.bytes);
  console.log(`wrote ${FORMS.length} files`);
}
