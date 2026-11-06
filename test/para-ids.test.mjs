// CR-002 section 66: what the content API adds takes a w14:paraId, so a file Word wrote keeps its ids. In some files one paragraph
// or row without an id makes Word 365 renumber the part (checks 79, 80): each route here runs on Word's own save, one such file,
// and leaves none without.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const WORD_SAVE = new Uint8Array(readFileSync(join(here, 'check79', '79c-save-1.word.docx')));
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const TABLE = `<w:tbl ${W}><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>from xml</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

/** The paragraphs and rows in document.xml with no id, and the ids used twice. */
function audit(bytes) {
  const xml = strFromU8(unzipSync(bytes)['word/document.xml']);
  const tags = [...xml.matchAll(/<w:(p|tr)(\s[^>]*?)?\/?>/g)];
  const ids = tags.map((m) => /w14:paraId="([^"]+)"/.exec(m[2] ?? '')?.[1]);
  return { total: tags.length, missing: ids.filter((id) => id === undefined).length, repeated: ids.filter((id, i) => id && ids.indexOf(id) !== i), xml };
}

async function wordFile() {
  const pkg = await WordprocessingMLPackage.load(WORD_SAVE);
  await pkg.getBody();
  return pkg;
}

test('the file these run on has an id on every paragraph and row, as Word wrote it', () => {
  const { total, missing, repeated } = audit(WORD_SAVE);
  assert.equal(missing, 0);
  assert.deepEqual(repeated, []);
  assert.ok(total > 15);
});

const firstTable = (pkg) => pkg.body.tables[0];
const firstControl = async (pkg) => (await pkg.body.getContentControls())[0];
const ROUTES = {
  'Table.addRows at the end': (pkg) => firstTable(pkg).addRows('End', 2, [['a', 'b'], ['c', 'd']]),
  'Table.addRows at the start': (pkg) => firstTable(pkg).addRows('Start', 1),
  'TableRow.insertRows after': (pkg) => firstTable(pkg).rows[0].insertRows('After', 1, [['a', 'b']]),
  'TableRow.insertRows before': (pkg) => firstTable(pkg).rows[0].insertRows('Before', 2),
  'Table.addColumns at the end': (pkg) => firstTable(pkg).addColumns('End', 1, [['a']]),
  'Table.addColumns at the start': (pkg) => firstTable(pkg).addColumns('Start', 2),
  'TableCell.insertColumns': (pkg) => firstTable(pkg).getCell(0, 0).insertColumns('After', 1),
  'Body.insertTable at the end': (pkg) => pkg.body.insertTable(2, 2, 'End', [['a', 'b'], ['c', 'd']]),
  'Body.insertTable at the start': (pkg) => pkg.body.insertTable(1, 1, 'Start'),
  'a page break in a paragraph of its own': (pkg) => pkg.body.paragraphs[1].insertBreak('Page', 'After'),
  'a table in a cell': (pkg) => firstTable(pkg).getCell(0, 0).body.insertTable(1, 1, 'End'),
  'a paragraph in a new cell': (pkg) => firstTable(pkg).addRows('End', 1)[0].cells[0].body.insertParagraph('more', 'End'),
  'a table in a new cell': (pkg) => firstTable(pkg).addRows('End', 1)[0].cells[0].body.insertTable(1, 1, 'End'),
  'insertXml holding a table': (pkg) => pkg.body.insertXml(TABLE, 'End'),
  'a tracked row': (pkg) => { pkg.changeTrackingMode = 'TrackAll'; firstTable(pkg).addRows('End', 1, [['a', 'b']]); },
  'a tracked table': (pkg) => { pkg.changeTrackingMode = 'TrackAll'; pkg.body.insertTable(1, 1, 'End'); },
  'a tracked column': (pkg) => { pkg.changeTrackingMode = 'TrackAll'; firstTable(pkg).addColumns('End', 1, [['a']]); },
  'a cell\'s last paragraph taken, Word\'s repair': (pkg) => firstTable(pkg).getCell(0, 0).paragraphs[0].delete(),
  'a paragraph added': (pkg) => pkg.body.insertParagraph('new', 'End'),
  'a paragraph after a paragraph': (pkg) => pkg.body.paragraphs[1].insertParagraph('new', 'After'),
  'a paragraph split': (pkg) => pkg.body.paragraphs[1].insertText('a\nb', 'End'),
  'a control\'s paragraph': async (pkg) => (await firstControl(pkg)).insertParagraph('new', 'End'),
  'a control\'s text over paragraphs': async (pkg) => (await firstControl(pkg)).insertText('a\nb', 'Replace'),
};

for (const [name, route] of Object.entries(ROUTES)) {
  test(`CR-002 section 66: ${name} leaves every paragraph and row with an id, none twice`, async () => {
    const pkg = await wordFile();
    const before = audit(WORD_SAVE).xml.match(/w14:paraId="[^"]+"/g);
    await route(pkg);
    const { missing, repeated, xml } = audit(await pkg.save());
    assert.equal(missing, 0, 'one without makes Word renumber the file (check 79)');
    assert.deepEqual(repeated, []);
    const kept = before.filter((id) => xml.includes(id)).length;
    assert.ok(kept >= before.length - 1, `${kept} of ${before.length} of Word's ids kept`);
  });
}

test('CR-002 section 66: the ids are there at once, not at the save', async () => {
  const pkg = await wordFile();
  const [row] = firstTable(pkg).addRows('End', 1, [['a', 'b']]);
  assert.match(row.tr.paraId, /^[0-9A-F]{8}$/);
  for (const cell of row.cells) assert.match(cell.paragraphs[0].paraId, /^[0-9A-F]{8}$/);
  assert.ok(parseInt(row.tr.paraId, 16) < 0x80000000, 'below 0x80000000, as the schema has it');
});

test('CR-002 section 66: a row or a table is addressable by its paragraphs\' ids straight after it\'s added', async () => {
  const pkg = await wordFile();
  const table = pkg.body.insertTable(1, 1, 'End', [['found']]);
  const id = table.getCell(0, 0).paragraphs[0].paraId;
  assert.equal(pkg.body.paragraphAt(`w14:${id}`)?.text, 'found');
});

test('CR-002 section 66: a document that uses no ids is left without', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.body.insertParagraph('one', 'End');
  const table = pkg.body.insertTable(2, 2, 'End');
  table.addRows('End', 1);
  table.addColumns('End', 1);
  const { total, missing } = audit(await pkg.save());
  assert.equal(missing, total);
});
