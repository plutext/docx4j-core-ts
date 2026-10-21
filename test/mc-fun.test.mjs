// CR-013: mc:AlternateContent at block, row and cell level, on docx4j's mc-fun probes (test/mc-fun/,
// copied from docx4j at the commit its manifest.json names). Phase 0 recorded what each reader made of
// them; phase 1 brought the text model's rows and cells, and mcePreprocess: true (D3), to Word 365's
// reading. The Office JS readers' assertions below are today's gaps, phase 2's: each fails when that
// phase reads the branch, as CLAUDE.md's rule has a gap announce its fix.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync } from 'node:fs';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage, Texts, PlainText, TextWarning } from '../dist/index.mjs';

const dir = new URL('./mc-fun/', import.meta.url);
const bytesOf = (name) => readFileSync(new URL(name, dir));
const load = (name, resolved) => WordprocessingMLPackage.load(bytesOf(name), resolved ? { mcePreprocess: true } : undefined);
const textOf = async (name, resolved) => PlainText.of(await Texts.of(await load(name, resolved))).trimEnd();

test('the probes are docx4j\'s, as copied: every file listed, every SHA-1 matching', () => {
  const manifest = JSON.parse(readFileSync(new URL('manifest.json', dir), 'utf8'));
  const listed = new Map(manifest.files.map((f) => [f.name, f.sha1]));
  const present = readdirSync(dir).filter((n) => n.endsWith('.docx')).sort();
  assert.deepEqual(present, [...listed.keys()].sort(), 'copy again from docx4j, never edit');
  for (const name of present) assert.equal(createHash('sha1').update(bytesOf(name)).digest('hex'), listed.get(name), name);
  assert.equal(present.length, 30);
});

// What the text model reads of each probe, kept or resolved: the branch Word 365 takes, w16sdtfl where a
// Choice requires it. The 13 with a Word save are held to it too, read the same way.
const READINGS = {
  'mc-text-newest-first': 'w:r\nI prefer w16sdtfl',
  'mc-run-newest-first': 'Run level test\n\nI prefer w16sdtfl',
  'mc-sdt-run-newest-first': 'Run sdt test\nI prefer w16sdtfl',
  'mc-block-newest-first': 'Block level test\n\nw16sdtfl',
  'mc-block-nofallback': 'Block level test\n\nw16sdtfl',
  'mc-sdt-block-newest-first': 'Block sdt test\n\nw16sdtfl',
  'mc-tc-block-newest-first': 'Block in cell test\nI prefer\tw16sdtfl',
  'mc-row-newest-first': 'Row level test\nI prefer\tbelow\nrow\tw16sdtfl',
  'mc-sdt-row-newest-first': 'Row sdt test\nI prefer\tbelow\nrow\tw16sdtfl',
  'mc-cell-newest-first': 'Cell level test\nI prefer\tw16sdtfl',
  'mc-sdt-cell-newest-first': 'Cell sdt test\nI prefer\tw16sdtfl',
  'mc-mixed': 'Mixed test\nsdt run: SDTRUN_W14\nhyperlink: HLINK_W14\nin p: INP_W14\nBODY_W14',
  'mc-txbx-block': 'Text box, block level\nBODY_CHOICE_PARA \nBODY_CHOICE_BOX',
  'mc-txbx-sdt-block': 'Text box, block sdt\nSDT_CHOICE_PARA \nSDT_CHOICE_BOX',
  'mc-txbx-tc-block': 'Text box, block in cell\nI prefer\tCELL_CHOICE_PARA \nCELL_CHOICE_BOX',
  'mc-txbx-run': "Text box, run level (Word's own form)\nRUN_PARA \nRUN_CHOICE_BOX",
};

test('the text model reads every probe as Word 365 does, at every level, kept or resolved (phase 1, D3)', async () => {
  let saves = 0;
  for (const [probe, expected] of Object.entries(READINGS)) {
    assert.equal(await textOf(`${probe}.docx`, false), expected, `${probe}, kept`);
    assert.equal(await textOf(`${probe}.docx`, true), expected, `${probe}, resolved`);
    if (readdirSync(dir).includes(`${probe}-word365.docx`)) {
      assert.equal(await textOf(`${probe}-word365.docx`, false), expected, `${probe}: Word 365's save`);
      saves++;
    }
  }
  assert.equal(saves, 13);
  // the original-order probe's save: Word took w14, the first Choice there
  assert.equal(await textOf('mc-run-word365.docx', false), 'Run level test\n\nI prefer w14');
});

test('a branch\'s rows and cells are addressed from the element (D2), and its branches compared (REQ-307)', async () => {
  const addresses = async (name) => {
    const result = await Texts.of(await load(name));
    const out = [];
    const visit = (blocks) => {
      for (const b of blocks) {
        if (b.kind !== 'TABLE') { out.push(`${b.address}=${b.text}`); continue; }
        for (const row of b.rows) for (const cell of row.cells) { out.push(`cell ${cell.address}`); visit(cell.blocks); }
      }
    };
    visit(result.blocks);
    return { out, warnings: result.warnings.map((w) => [w.kind, w.address, w.branch]) };
  };
  const differ = (address) => [[TextWarning.MCE_BRANCHES_DIFFER, address, 'mc:Choice Requires="w16sdtfl"']];
  // the block level, as before: the element's address, then the paragraph's in the branch
  assert.deepEqual(await addresses('mc-block-newest-first.docx'), { out: ['body/0=Block level test', 'body/1=', 'body/2/0=w16sdtfl'], warnings: differ('body/2') });
  // the row level: the table's second row is the element's (body/1/1), its branch's row body/1/1/0
  assert.deepEqual(await addresses('mc-row-newest-first.docx'), {
    out: ['body/0=Row level test', 'cell body/1/0/0', 'body/1/0/0/0=I prefer', 'cell body/1/0/1', 'body/1/0/1/0=below',
      'cell body/1/1/0/0', 'body/1/1/0/0/0=row', 'cell body/1/1/0/1', 'body/1/1/0/1/0=w16sdtfl', 'body/2='],
    warnings: differ('body/1/1'),
  });
  // the cell level: the row's second cell is the element's (body/1/0/1), its branch's cell body/1/0/1/0
  assert.deepEqual(await addresses('mc-cell-newest-first.docx'), {
    out: ['body/0=Cell level test', 'cell body/1/0/0', 'body/1/0/0/0=I prefer', 'cell body/1/0/1/0', 'body/1/0/1/0/0=w16sdtfl', 'body/2='],
    warnings: differ('body/1/0/1'),
  });
  // in a content control: the control's, then the element's in its content
  assert.deepEqual((await addresses('mc-sdt-row-newest-first.docx')).warnings, differ('body/1/1/0'));
  assert.deepEqual((await addresses('mc-sdt-cell-newest-first.docx')).warnings, differ('body/1/0/1/0'));
});

// The Office JS readers today (CR-013 phase 2 will change each of these): a block-, row- or cell-level
// element is invisible to Body.paragraphs, Body.text and a table's rows and cells, while TableCell.values
// reads a block-level element in its cell, which the cell's paragraphs don't list.
test('the Office JS readers miss a block-, row- or cell-level element (phase 2\'s gap, asserted as it is)', async () => {
  const office = async (name) => {
    const pkg = await load(name);
    await pkg.mainDocumentPart.getContents();
    const body = pkg.body;
    return {
      paragraphs: body.paragraphs.map((p) => p.text),
      tables: body.tables.map((t) => ({ values: t.values, cells: t.rows.map((row) => row.cells.map((c) => c.paragraphs.map((p) => p.text))) })),
    };
  };
  // the block level: the branch's paragraph isn't one of the body's
  assert.deepEqual((await office('mc-block-newest-first.docx')).paragraphs, ['Block level test', '']);
  assert.deepEqual((await office('mc-sdt-block-newest-first.docx')).paragraphs, ['Block sdt test', '']);
  assert.deepEqual((await office('mc-mixed.docx')).paragraphs, ['Mixed test', 'sdt run: SDTRUN_W14', 'hyperlink: HLINK_W14', 'in p: INP_W14']);
  assert.deepEqual((await office('mc-txbx-block.docx')).paragraphs, ['Text box, block level']);
  // the row and cell levels: the branch's row, or cell, isn't the table's
  assert.deepEqual((await office('mc-row-newest-first.docx')).tables, [{ values: [['I prefer', 'below']], cells: [[['I prefer'], ['below']]] }]);
  assert.deepEqual((await office('mc-sdt-row-newest-first.docx')).tables, [{ values: [['I prefer', 'below']], cells: [[['I prefer'], ['below']]] }]);
  assert.deepEqual((await office('mc-cell-newest-first.docx')).tables, [{ values: [['I prefer']], cells: [[['I prefer']]] }]);
  assert.deepEqual((await office('mc-sdt-cell-newest-first.docx')).tables, [{ values: [['I prefer']], cells: [[['I prefer']]] }]);
  // a block-level element in a cell: values reads it, the cell's paragraphs don't
  assert.deepEqual((await office('mc-tc-block-newest-first.docx')).tables, [{ values: [['I prefer', 'w16sdtfl']], cells: [[['I prefer'], []]] }]);
  // what Word's saves give them, every element resolved
  assert.deepEqual((await office('mc-row-newest-first-word365.docx')).tables[0].values, [['I prefer', 'below'], ['row', 'w16sdtfl']]);
  assert.deepEqual((await office('mc-cell-newest-first-word365.docx')).tables[0].values, [['I prefer', 'w16sdtfl']]);
});

test('a resolved load gives the Office JS readers Word 365\'s branch at every level (D3)', async () => {
  for (const [probe, values] of [['mc-row-newest-first', [['I prefer', 'below'], ['row', 'w16sdtfl']]], ['mc-cell-newest-first', [['I prefer', 'w16sdtfl']]], ['mc-tc-block-newest-first', [['I prefer', 'w16sdtfl']]]]) {
    const pkg = await load(`${probe}.docx`, true);
    await pkg.mainDocumentPart.getContents();
    assert.deepEqual(pkg.body.tables[0].values, values, probe);
  }
  const pkg = await load('mc-block-newest-first.docx', true);
  await pkg.mainDocumentPart.getContents();
  assert.deepEqual(pkg.body.paragraphs.map((p) => p.text), ['Block level test', '', 'w16sdtfl']);
});

test('the default load saves every element whole, every branch and its text, after a re-marshal', async () => {
  const count = (xml, re) => (xml.match(re) ?? []).length;
  const texts = (xml) => [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)</g)].map((m) => m[1]).sort();
  for (const name of readdirSync(dir).filter((n) => n.endsWith('.docx')).sort()) {
    const source = strFromU8(unzipSync(bytesOf(name))['word/document.xml']);
    const pkg = await load(name);
    const part = pkg.mainDocumentPart;
    await part.getContents();
    part.setContents(part.contents);
    const saved = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
    for (const re of [/<mc:AlternateContent\b/g, /<mc:Choice\b/g, /<mc:Fallback\b/g]) assert.equal(count(saved, re), count(source, re), `${name}: ${re.source}`);
    assert.deepEqual(texts(saved), texts(source), `${name}: every w:t`);
  }
});

// D3's round trip (CR-013 section 3, item 6): what Word writes in the four namespaces a resolving load now
// counts as understood is attributes, which survive a re-marshal unbound, so taking their Choice loses nothing.
test('a resolved load takes a Choice requiring w16du or w16sdtdh and keeps their attributes (D3)', async () => {
  const base = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const documentWith = (body) => zipSync({ ...base, 'word/document.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" mc:Ignorable="w15 w16du w16sdtdh w16cex"><w:body>${body}<w:sectPr/></w:body></w:document>`) });
  const fallback = '<mc:Choice Requires="w16cex"><w:p><w:r><w:t>w16cex</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback>';
  const cases = {
    'w16du:dateUtc="2026-10-21T00:00:00Z"': `<mc:AlternateContent><mc:Choice Requires="w16du"><w:p><w:ins w:id="1" w:author="A" w:date="2026-10-21T10:00:00" w16du:dateUtc="2026-10-21T00:00:00Z"><w:r><w:t>w16du</w:t></w:r></w:ins></w:p></mc:Choice>${fallback}</mc:AlternateContent>`,
    // check 35's binding (test/fixtures/revisions/check35/35a-word15.docx), its checksum as Word wrote it
    'w16sdtdh:storeItemChecksum="JAms/A=="': `<mc:AlternateContent><mc:Choice Requires="w16sdtdh"><w:sdt><w:sdtPr><w:id w:val="480498328"/><w15:dataBinding w:prefixMappings="" w:xpath="/data[1]/name[1]" w:storeItemID="{E0B29C1B-74EB-4E2F-A275-01AE9154A015}" w16sdtdh:storeItemChecksum="JAms/A=="/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>w16sdtdh</w:t></w:r></w:p></w:sdtContent></w:sdt></mc:Choice>${fallback}</mc:AlternateContent>`,
  };
  for (const [attribute, body] of Object.entries(cases)) {
    const pkg = await WordprocessingMLPackage.load(documentWith(body), { mcePreprocess: true });
    const part = pkg.mainDocumentPart;
    await part.getContents();
    part.setContents(part.contents);
    const saved = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
    assert.ok(saved.includes(attribute), `${attribute} kept`);
    assert.ok(!saved.includes('mc:AlternateContent') && !saved.includes('w16cex<') && !saved.includes('fallback'), 'resolved to the first Choice');
  }
});
