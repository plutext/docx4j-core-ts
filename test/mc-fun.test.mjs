// CR-013: mc:AlternateContent at block, row and cell level, on docx4j's mc-fun probes (test/mc-fun/,
// copied from docx4j at the commit its manifest.json names). Phase 0 recorded what each reader made of
// them; phase 1 brought the text model's rows and cells, and mcePreprocess: true (D3), to Word 365's
// reading; phase 2 the Office JS readers and the tracked-change listing, with every view in a branch
// read-only (D4); phase 3 the mutators reaching every branch: the other branches' revisions resolved with
// their twins (docx4j's TrackedChangesAlternateContentTest), the bulk edits, the table edits, the ids.
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
test("the Office JS readers take the selected branch at block, row and cell level, as Word 365's saves read (CR-013 phase 2)", async () => {
  const office = async (name, resolved) => {
    const pkg = await load(name, resolved);
    await pkg.mainDocumentPart.getContents();
    const body = pkg.body;
    return {
      body,
      paragraphs: body.paragraphs.map((p) => p.text),
      text: body.text,
      tables: body.tables.map((t) => ({ values: t.values, rows: t.rowCount, cells: t.rows.map((row) => row.cells.map((c) => c.paragraphs.map((p) => p.text))) })),
      controls: body.contentControls.map((c) => [c.form, c.text]),
      search: body.search('w16sdtfl').map((r) => body.addressOf(r.paragraphs[0])),
    };
  };
  const probes = readdirSync(dir).filter((n) => n.endsWith('.docx') && !n.includes('word365')).sort();
  let saves = 0;
  for (const name of probes) {
    const kept = await office(name);
    // a resolved load (D3) and a kept one read alike
    const resolved = await office(name, true);
    for (const key of ['paragraphs', 'text', 'tables', 'controls']) assert.deepEqual(kept[key], resolved[key], `${name}, resolved: ${key}`);
    // and Word 365's save of the probe, every element resolved by Word, reads alike
    const save = name.replace('.docx', '-word365.docx');
    if (!readdirSync(dir).includes(save)) continue;
    saves++;
    const word = await office(save);
    for (const key of ['paragraphs', 'text', 'tables']) assert.deepEqual(kept[key], word[key], `${name}, Word's save: ${key}`);
    // every paragraph's address (the element's form, D2) gives the paragraph back
    for (const p of kept.body.paragraphs) {
      const address = kept.body.addressOf(p);
      assert.equal(kept.body.paragraphAt(address)?.p, p.p, `${name} ${address}`);
    }
  }
  assert.equal(saves, 13);                                           // mc-run-word365 is the original-order probe's save
  // the element's form (D2): a branch's paragraph addressed from the element, as the text model and docx4j have it
  const addresses = async (name) => { const { body } = await office(name); return body.paragraphs.map((p) => body.addressOf(p)); };
  assert.deepEqual(await addresses('mc-block-newest-first.docx'), ['body/0', 'body/1', 'body/2/0']);
  assert.deepEqual(await addresses('mc-row-newest-first.docx'), ['body/0', 'body/1/0/0/0', 'body/1/0/1/0', 'body/1/1/0/0/0', 'body/1/1/0/1/0', 'body/2']);
  assert.deepEqual(await addresses('mc-cell-newest-first.docx'), ['body/0', 'body/1/0/0/0', 'body/1/0/1/0/0', 'body/2']);
  // the readers' answers on the probes
  assert.deepEqual((await office('mc-row-newest-first.docx')).tables, [{ values: [['I prefer', 'below'], ['row', 'w16sdtfl']], rows: 2, cells: [[['I prefer'], ['below']], [['row'], ['w16sdtfl']]] }]);
  assert.deepEqual((await office('mc-cell-newest-first.docx')).tables, [{ values: [['I prefer', 'w16sdtfl']], rows: 1, cells: [[['I prefer'], ['w16sdtfl']]] }]);
  assert.deepEqual((await office('mc-tc-block-newest-first.docx')).tables[0].values, [['I prefer', 'w16sdtfl']]);
  // a row control's text is its cells a \t apart, as Word 365 reads one (check 53, case 15; CR-014 section 9.7)
  assert.deepEqual((await office('mc-sdt-row-newest-first.docx')).controls, [['Row', 'row\tw16sdtfl']]);
  assert.deepEqual((await office('mc-mixed.docx')).paragraphs, ['Mixed test', 'sdt run: SDTRUN_W14', 'hyperlink: HLINK_W14', 'in p: INP_W14', 'BODY_W14']);
  assert.deepEqual((await office('mc-block-newest-first.docx')).search, ['body/2/0']);
  assert.deepEqual((await office('mc-cell-newest-first.docx')).search, ['body/1/0/1/0/0']);
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

// --- CR-013 phase 2: the tracked-change listing, the table's rows, the views' refusals (D4) -----------------

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
/** A package of this body, saved and loaded again, so that every PARENT link is the load's. */
async function mini(bodyXml) {
  const make = await WordprocessingMLPackage.createPackage();
  make.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:w14="${W14}" mc:Ignorable="w14"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`);
  const pkg = await WordprocessingMLPackage.load(await make.save());
  await pkg.getBody();
  return pkg;
}
const documentXml = async (pkg) => strFromU8(unzipSync(await pkg.save())['word/document.xml']);
const AC = (choice, fallback) => `<mc:AlternateContent><mc:Choice Requires="w14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
const ins = (id, text) => `<w:ins w:id="${id}" w:author="A"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:ins>`;
const del = (id, text) => `<w:del w:id="${id}" w:author="A"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const TBL = '<w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>';
const run = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const changed = (branch) => `${run(`${branch} `)}${ins(branch === 'C' ? 1 : 3, 'INSERTED')}${del(branch === 'C' ? 2 : 4, 'DELETED')}`;
const cell = (content) => `<w:tc><w:p>${content}</w:p></w:tc>`;
const row = (content) => `<w:tr>${cell(content)}</w:tr>`;
const deletedRow = (id, text) => `<w:tr><w:trPr><w:del w:id="${id}" w:author="A"/></w:trPr><w:tc><w:p><w:r><w:delText>${text}</w:delText></w:r></w:p></w:tc></w:tr>`;
// docx4j's TrackedChangesAlternateContentTest's four shapes (CR-051 phase 2, 9db1a3cbd)
const LEVELS = {
  block: `<w:p>${run('before')}</w:p>${AC(`<w:p>${changed('C')}</w:p>`, `<w:p>${changed('F')}</w:p>`)}<w:p>${run('after')}</w:p>`,
  paragraph: `<w:p>${run('x ')}${AC(changed('C'), changed('F'))}</w:p>`,
  row: `<w:tbl>${TBL}${AC(row(changed('C')), row(changed('F')))}</w:tbl><w:p/>`,
  cell: `<w:tbl>${TBL}<w:tr>${AC(cell(changed('C')), cell(changed('F')))}</w:tr></w:tbl><w:p/>`,
};

test("a branch's tracked changes are listed once, from the selected branch, at every level; the text model's views planned from them (phase 2); Accept All reaches every branch (phase 3)", async () => {
  for (const [level, xml] of Object.entries(LEVELS)) {
    const pkg = await mini(xml);
    const listed = (await pkg.getAllTrackedChanges()).map((c) => [c.type, c.text]);
    assert.deepEqual(listed, [['Added', 'INSERTED'], ['Deleted', 'DELETED']], level);
    // the accepted view takes the deletion out of the branch Word reads, where it read it before this phase
    const accepted = PlainText.of(await Texts.of(pkg));
    assert.ok(accepted.includes('C INSERTED') && !accepted.includes('DELETED') && !accepted.includes('F '), `${level}: ${accepted}`);
    const original = PlainText.of(await Texts.of(pkg, { view: 'original' }));
    assert.ok(original.includes('C DELETED') && !original.includes('INSERTED'), `${level}: ${original}`);
    // Accept All counts what it listed and resolves every branch (phase 3; docx4j's acceptAllResolvesEveryBranch)
    assert.equal(await pkg.acceptAllRevisions(), 2, level);
    const saved = await documentXml(pkg);
    const choice = saved.match(/<mc:Choice[^>]*>(.*?)<\/mc:Choice>/s)[1];
    const fallback = saved.match(/<mc:Fallback>(.*?)<\/mc:Fallback>/s)[1];
    for (const [name, branch] of [['Choice', choice], ['Fallback', fallback]]) {
      assert.ok(!/<w:(ins|del) /.test(branch) && branch.includes('INSERTED') && !branch.includes('DELETED'), `${level}: the ${name} resolved`);
    }
    assert.deepEqual(await pkg.getAllTrackedChanges(), [], `${level}: nothing left to list`);
  }
});

test("a deleted row accepted leaves its table while a branch has a row (docx4j's Lister.hasRows), and takes it with its twin; Table.deleteColumns likewise (phases 2 and 3)", async () => {
  // the Choice's deleted row accepted, a plain row in the Fallback: the table stays, the Choice emptied
  // (docx4j's acceptingADeletedRowKeepsATableWhoseOtherBranchHasARow)
  let pkg = await mini(`<w:tbl>${TBL}${AC(deletedRow(1, 'C'), row(run('F')))}</w:tbl><w:p/>`);
  let changes = await pkg.getAllTrackedChanges();
  assert.equal(changes.length, 1);
  changes[0].accept();
  let xml = await documentXml(pkg);
  assert.ok(xml.includes('<w:tbl') && /<mc:Fallback><w:tr/.test(xml) && !/<mc:Choice Requires="w14"><w:tr/.test(xml), xml);
  // a deleted row in every branch: listed once; accepted, its twin goes with it, and the table with both
  // (phase 3; docx4j's acceptingADeletedRowInEveryBranchRemovesTheTable)
  pkg = await mini(`<w:p>${run('before')}</w:p><w:tbl>${TBL}${AC(deletedRow(1, 'C'), deletedRow(2, 'F'))}</w:tbl><w:p/>`);
  changes = await pkg.getAllTrackedChanges();
  assert.equal(changes.length, 1);
  changes[0].accept();
  xml = await documentXml(pkg);
  assert.ok(!xml.includes('<w:tbl') && !xml.includes('<w:tr') && xml.includes('before'), xml);
  // section 7.3's shape: a plain deleted row beside the element, whose rows were not counted, took the table
  pkg = await mini(`<w:tbl>${TBL}${deletedRow(1, 'D')}${AC(row(run('C')), row(run('F')))}</w:tbl><w:p/>`);
  assert.equal(await pkg.acceptAllRevisions(), 1);
  assert.deepEqual((await pkg.getBody()).tables.map((t) => t.values), [[['C']]]);
  // Table.deleteColumns on a table whose other cells are in an element (section 7.3's second shape)
  const probe = await load('mc-cell-newest-first.docx');
  const body = await probe.getBody();
  body.tables[0].deleteColumns(0);
  assert.deepEqual(body.tables.map((t) => t.values), [[['w16sdtfl']]]);
});

test("a view in a branch is read-only (D4): paragraphs, bodies, tables, rows, cells and controls refuse every write; the bulk edits reach every branch (phase 3)", async () => {
  const refuses = (name, f) => assert.throws(f, (e) => e.name === 'NotSupportedError' && /in a branch of an mc:AlternateContent, which is read-only \(CR-013 D4\)/.test(e.message), name);
  // a paragraph in a block-level branch, and the body beside it
  let pkg = await mini(LEVELS.block);
  let body = await pkg.getBody();
  const p = body.paragraphs[1];
  assert.equal(p.text, 'C INSERTED');
  assert.equal(p.inBranch, true);
  assert.equal(body.paragraphs[0].inBranch, false);
  refuses('Paragraph.insertText', () => p.insertText('x', 'End'));
  refuses('Range.delete', () => p.search('C')[0].delete());
  refuses('Font', () => { p.getRange().font.bold = true; });
  refuses('Paragraph.delete', () => p.delete());
  refuses('Paragraph.insertParagraph', () => p.insertParagraph('x', 'After'));
  refuses('Body.insertParagraph beside it', () => body.insertElement(body.content[0], 'After', p));
  // the bulk edits reach every branch (phase 3): replaceText counts what a reader sees, and a tracked
  // clear marks every branch's paragraphs deleted
  assert.equal(body.replaceText('INSERTED', 'Z'), 1);
  let saved = await documentXml(pkg);
  assert.ok(!saved.includes('INSERTED') && saved.includes('>C Z<') && saved.includes('>F Z<'), saved);
  pkg.changeTrackingMode = 'TrackAll';
  body.clear();
  pkg.changeTrackingMode = 'Off';
  saved = await documentXml(pkg);
  for (const branch of ['C', 'F']) assert.ok(saved.includes(`<w:delText>${branch} Z</w:delText>`), `${branch}: ${saved}`);
  assert.equal((saved.match(/<w:pPr><w:rPr><w:del /g) ?? []).length, 4, 'every mark deleted, in both branches');
  // rows, cells and their bodies in a row-level branch
  pkg = await mini(`<w:tbl>${TBL}${row(run('plain'))}${AC(row(run('C')), row(run('F')))}</w:tbl><w:p/>`);
  body = await pkg.getBody();
  const table = body.tables[0];
  assert.deepEqual(table.values, [['plain'], ['C']]);
  const [plain, branch] = table.rows;
  assert.deepEqual([plain.inBranch, branch.inBranch, branch.cells[0].inBranch], [false, true, true]);
  refuses('TableRow.delete', () => branch.delete());
  refuses('TableRow.insertRows', () => branch.insertRows('After', 1));
  refuses('Table.deleteRows', () => table.deleteRows(1));
  refuses('TableCell.value', () => { branch.cells[0].value = 'x'; });
  refuses('TableCell.width', () => { branch.cells[0].width = 30; });
  refuses('TableCell.body', () => branch.cells[0].body.insertParagraph('x', 'End'));
  refuses('Table.values', () => { table.values = [['first'], ['second']]; });
  assert.deepEqual(table.values, [['plain'], ['C']], 'nothing written before the refusal');
  table.deleteRows(0);
  assert.deepEqual(table.values, [['C']], 'a row outside the branch is deleted as before');
  // a control in a block-level branch, and its kinds
  const sdt = (text) => `<w:sdt><w:sdtPr><w:tag w:val="t"/><w:dropDownList><w:listItem w:displayText="a" w:value="a"/></w:dropDownList></w:sdtPr><w:sdtContent><w:p>${run(text)}</w:p></w:sdtContent></w:sdt>`;
  pkg = await mini(AC(sdt('C'), sdt('F')));
  body = await pkg.getBody();
  const [control] = body.contentControls;
  assert.deepEqual([control.text, control.tag, control.inBranch], ['C', 't', true]);
  refuses('ContentControl.tag', () => { control.tag = 'x'; });
  refuses('ContentControl.delete', () => control.delete());
  refuses('ContentControl.color', () => { control.color = '#FF0000'; });
  refuses('ListContentControl.addListItem', () => control.dropDownListContentControl.addListItem('b'));
  refuses('ContentControlListItem.value', () => { control.dropDownListContentControl.listItems[0].value = 'z'; });
});

test("the table context counts a branch's rows, and a comment in a branch is listed and deleted from every branch (phase 2)", async () => {
  // a row-level branch's row is the table's last row for a style's conditional formats (TableContext)
  const look = '<w:tblPr><w:tblLook w:val="0060" w:firstRow="1" w:lastRow="1" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>';
  let pkg = await mini(`<w:tbl>${look}${row(run('one'))}${AC(row(run('C')), row(run('F')))}</w:tbl><w:p/>`);
  let body = await pkg.getBody();
  assert.deepEqual(body.paragraphs.slice(0, 2).map((p) => [p.text, [...p.cellContext.conditions]]), [['one', ['firstRow']], ['C', ['lastRow']]]);
  // a comment on a paragraph, which is then made a block-level element's in both branches
  const make = await mini(`<w:p>${run('before')}</w:p><w:p>${run('anchored')}</w:p>`);
  await (await make.getBody()).paragraphs[1].getRange().insertComment('on the branch');
  const files = unzipSync(await make.save());
  let xml = strFromU8(files['word/document.xml']);
  const paragraph = xml.match(/<w:p\b(?:(?!<w:p\b).)*?anchored.*?<\/w:p>/s)[0];
  xml = xml.replace(paragraph, AC(paragraph, paragraph.replace('anchored', 'fallback')));
  if (!xml.includes(`xmlns:mc="${MC}"`)) xml = xml.replace('<w:document ', `<w:document xmlns:mc="${MC}" `);
  if (!xml.includes(`xmlns:w14="${W14}"`)) xml = xml.replace('<w:document ', `<w:document xmlns:w14="${W14}" `);
  files['word/document.xml'] = strToU8(xml);
  pkg = await WordprocessingMLPackage.load(zipSync(files));
  body = await pkg.getBody();
  const [comment] = await body.getComments();
  assert.equal(comment.content, 'on the branch');
  assert.deepEqual(comment.getRanges().map((r) => [body.addressOf(r.paragraphs[0]), r.text]), [['body/1/0', 'anchored']]);
  await comment.delete();
  const saved = await documentXml(pkg);
  assert.ok(!/comment(RangeStart|RangeEnd|Reference)/.test(saved), 'the markers gone from every branch');
  assert.ok(saved.includes('>anchored<') && saved.includes('>fallback<'));
});

// --- CR-013 phase 3: the mutators reach every branch ---------------------------------------------------------

const count = (s, token) => s.split(token).length - 1;
const bodyXmlOf = async (pkg) => (await documentXml(pkg)).match(/<w:body>(.*)<\/w:body>/s)[1];

test("docx4j's TrackedChangesAlternateContentTest: each listed change has its twin, resolved with it; Accept All resolves every branch (phase 3)", async () => {
  for (const [level, xml] of Object.entries(LEVELS)) {
    // listedOnceWithTwins: two listed, from the selected branch
    let pkg = await mini(xml);
    assert.deepEqual((await pkg.getAllTrackedChanges()).map((c) => c.kind), ['RUN_INSERTION', 'RUN_DELETION'], level);
    // acceptingResolvesTheTwin: the insertion accepted in both branches, the deletion rejected in both
    (await pkg.getAllTrackedChanges())[0].accept();
    let out = await bodyXmlOf(pkg);
    assert.ok(!out.includes('<w:ins '), `${level}: ${out}`);
    assert.equal(count(out, 'INSERTED'), 2, `${level}: the insertion stays in both branches`);
    assert.equal((await pkg.getAllTrackedChanges()).length, 1, `${level}: the deletions are still listed, once`);
    (await pkg.getAllTrackedChanges())[0].reject();
    out = await bodyXmlOf(pkg);
    assert.ok(!out.includes('<w:del '), `${level}: ${out}`);
    assert.equal(count(out, 'DELETED'), 2, `${level}: the deleted text is back in both branches`);
    assert.deepEqual(await pkg.getAllTrackedChanges(), []);
    // acceptAllResolvesEveryBranch: counts what it listed
    pkg = await mini(xml);
    assert.equal(await pkg.acceptAllRevisions(), 2, level);
    out = await bodyXmlOf(pkg);
    assert.ok(!out.includes('<w:ins ') && !out.includes('<w:del '), `${level}: ${out}`);
    assert.deepEqual([count(out, 'INSERTED'), count(out, 'DELETED')], [2, 0], level);
    assert.ok(out.includes('<mc:Choice') && out.includes('<mc:Fallback'), `${level}: the element is still whole`);
    // Reject All, and a piece resolved on its own takes its twin too (docx4j leaves it to acceptAll)
    pkg = await mini(xml);
    assert.equal(await pkg.rejectAllRevisions(), 2, level);
    out = await bodyXmlOf(pkg);
    assert.deepEqual([count(out, 'INSERTED'), count(out, 'DELETED'), count(out, '<w:ins '), count(out, '<w:del ')], [0, 2, 0, 0], level);
    pkg = await mini(xml);
    (await pkg.getAllTrackedChanges())[1].pieces[0].accept();
    assert.equal(count(await bodyXmlOf(pkg), 'DELETED'), 0, level);
  }
  // the accepted view of a cell holding a block-level element reads the branch Word reads (acceptedViewReadsTheBranchWordReads)
  const pkg = await mini(`<w:tbl>${TBL}<w:tr><w:tc>${LEVELS.block}</w:tc></w:tr></w:tbl><w:p/>`);
  assert.equal((await pkg.getAllTrackedChanges())[0].text, 'INSERTED');
});

test("a resolve-all with a filter matches each branch's changes on their own, as docx4j's (phase 3)", async () => {
  const insBy = (id, author, text) => `<w:ins w:id="${id}" w:author="${author}"><w:r><w:t>${text}</w:t></w:r></w:ins>`;
  const pkg = await mini(`<w:p>${run('x')}</w:p>${AC(`<w:p>${insBy(1, 'A', 'byA')}</w:p>`, `<w:p>${insBy(2, 'B', 'byB')}</w:p>`)}`);
  assert.equal(await pkg.acceptAllRevisions({ authors: ['A'] }), 1);
  let out = await bodyXmlOf(pkg);
  assert.ok(!/<w:ins w:id="1"/.test(out) && /<w:ins w:id="2"/.test(out), out);
  assert.equal(await pkg.acceptAllRevisions({ authors: ['B'] }), 0, 'counted as listed: the Fallback lists nothing');
  out = await bodyXmlOf(pkg);
  assert.ok(!out.includes('<w:ins '), 'but resolved, matched on its own');
});

test("a deleted cell's twin goes in its own branch, and a table's property records are one change over every branch (docx4j 60bd4b9ad, 901a922a3)", async () => {
  const tcOf = (text, tcPr) => `<w:tc><w:tcPr>${tcPr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const deletedCell = (id, text) => tcOf(text, `<w:tcW w:w="1000" w:type="dxa"/><w:cellDel w:id="${id}" w:author="A"/>`);
  const plainCell = (text) => tcOf(text, '<w:tcW w:w="1000" w:type="dxa"/>');
  // acceptingADeletedCellResolvesItsTwinInItsOwnBranch: each branch's own neighbour widened once
  let pkg = await mini(`<w:tbl>${TBL}<w:tr>${AC(deletedCell(1, 'C') + plainCell('Ckeep'), deletedCell(2, 'F') + plainCell('Fkeep'))}</w:tr></w:tbl><w:p/>`);
  let changes = await pkg.getAllTrackedChanges();
  assert.deepEqual(changes.map((c) => c.kind), ['CELL_DELETION']);
  changes[0].accept();
  let out = await bodyXmlOf(pkg);
  assert.ok(!out.includes('cellDel') && !out.includes('>C<') && !out.includes('>F<') && out.includes('Ckeep') && out.includes('Fkeep'), out);
  assert.deepEqual([count(out, '<w:gridSpan w:val="2"/>'), count(out, 'w:w="2000"')], [2, 2], 'each branch\'s neighbour widened once');
  assert.deepEqual(await pkg.getAllTrackedChanges(), []);
  // acceptingADeletedCellAloneInEachBranchWidensTheRowsNeighbourOnce
  pkg = await mini(`<w:tbl>${TBL}<w:tr>${plainCell('keep')}${AC(deletedCell(1, 'C'), deletedCell(2, 'F'))}</w:tr></w:tbl><w:p/>`);
  (await pkg.getAllTrackedChanges())[0].accept();
  out = await bodyXmlOf(pkg);
  assert.ok(!out.includes('cellDel') && !out.includes('>C<') && !out.includes('>F<'), out);
  assert.deepEqual([count(out, '<w:gridSpan w:val="2"/>'), count(out, 'w:w="2000"')], [1, 1], 'the row\'s neighbour widened once');
  assert.ok(out.includes('<mc:Choice') && out.includes('<mc:Fallback'), 'the element stays, its branches empty of cells');
  // resolvingATablePropertyRecordReachesEveryBranchsCell
  const changedCell = (id, text) => tcOf(text, `<w:tcW w:w="1000" w:type="dxa"/><w:shd w:val="clear" w:fill="FF0000"/><w:tcPrChange w:id="${id}" w:author="A"><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr></w:tcPrChange>`);
  const records = `<w:tbl>${TBL}<w:tr>${AC(changedCell(1, 'C'), changedCell(2, 'F'))}</w:tr></w:tbl><w:p/>`;
  pkg = await mini(records);
  assert.deepEqual((await pkg.getAllTrackedChanges()).map((c) => [c.kind, c.text]), [['TABLE_PROPERTIES', 'C\r\n']]);
  await pkg.acceptAllRevisions();
  out = await bodyXmlOf(pkg);
  assert.deepEqual([count(out, 'tcPrChange'), count(out, 'w:fill="FF0000"')], [0, 2], 'accepted: both branches keep the shading');
  pkg = await mini(records);
  await pkg.rejectAllRevisions();
  out = await bodyXmlOf(pkg);
  assert.deepEqual([count(out, 'tcPrChange'), count(out, 'w:fill="FF0000"'), count(out, '<w:tcW')], [0, 0, 2], 'rejected: both branches lose it');
  // acceptingATablePropertyRecordReachesARowInAnotherBranch: listed, though only the Fallback's row has it,
  // with no text, since Word shows none of it
  const changedRow = `<w:tr><w:trPr><w:cantSplit/><w:trPrChange w:id="1" w:author="A"><w:trPr/></w:trPrChange></w:trPr>${plainCell('F')}</w:tr>`;
  pkg = await mini(`<w:tbl>${TBL}${AC(row(run('C')), changedRow)}</w:tbl><w:p/>`);
  assert.deepEqual((await pkg.getAllTrackedChanges()).map((c) => [c.kind, c.text]), [['TABLE_PROPERTIES', '']]);
  await pkg.acceptAllRevisions();
  out = await bodyXmlOf(pkg);
  assert.ok(!out.includes('trPrChange') && out.includes('<w:cantSplit/>'), out);
  assert.deepEqual(await pkg.getAllTrackedChanges(), []);
  // aTableChangesTextReadsTheBranchWordReads: a whole-table change's text is the selected branch's rows
  pkg = await mini(`<w:tbl><w:tblPr><w:jc w:val="center"/><w:tblPrChange w:id="1" w:author="A"><w:tblPr/></w:tblPrChange></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>${AC(row(run('C')), row(run('F')))}</w:tbl><w:p/>`);
  assert.deepEqual((await pkg.getAllTrackedChanges()).map((c) => [c.kind, c.text]), [['TABLE_PROPERTIES', 'C\r\n']]);
  // a rejected grid record collapses the grid along each reading of a row (cellLayoutsOf): the recorded four
  // columns become three, and each branch's cell keeps its span of two over them (counted along one row,
  // as phase 2 did, the Fallback's cell lost its span)
  const gridRecord = '<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblPrChange w:id="8" w:author="A"><w:tblPr/></w:tblPrChange></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:tblGridChange w:id="9"><w:tblGrid><w:gridCol w:w="250"/><w:gridCol w:w="250"/><w:gridCol w:w="500"/><w:gridCol w:w="1000"/></w:tblGrid></w:tblGridChange></w:tblGrid>';
  const spanning = (text, n) => tcOf(text, `<w:gridSpan w:val="${n}"/>`);
  pkg = await mini(`<w:tbl>${gridRecord}<w:tr>${AC(spanning('C', 3), spanning('F', 3))}${tcOf('last', '')}</w:tr><w:tr>${spanning('a', 2)}${spanning('b', 2)}</w:tr></w:tbl><w:p/>`);
  await pkg.rejectAllRevisions();
  out = await bodyXmlOf(pkg);
  assert.ok(out.includes('<w:tblGrid><w:gridCol w:w="500"/><w:gridCol w:w="500"/><w:gridCol w:w="1000"/></w:tblGrid>'), out);
  for (const text of ['C', 'F', 'b']) assert.ok(out.includes(`<w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>${text}</w:t>`), `${text}: ${out}`);
});

test("a single change in a text box resolves its mc:Fallback twin, as docx4j's (CR-011's gap, phase 3)", async () => {
  const pkg = await WordprocessingMLPackage.load(readFileSync(new URL('./fixtures/revisions/cr043/cr043-tracked.docx', import.meta.url)));
  const boxed = (await pkg.getAllTrackedChanges()).filter((c) => c.text === ' DELETEME' && c.part?.partName.name === '/word/document.xml');
  assert.equal(boxed.length, 1);
  boxed[0].accept();
  const xml = await documentXml(pkg);
  const ac = xml.match(/<mc:AlternateContent>(?:(?!<mc:AlternateContent>).)*?C22 text box.*?<\/mc:AlternateContent>/s)[0];
  assert.equal(count(ac, 'C22 text box text'), 2, 'both boxes');
  assert.equal(count(ac, 'DELETEME'), 0, 'the deletion accepted in the Choice\'s box and in its Fallback twin');
  assert.equal(count(ac, 'C22 inserted'), 2);
});

test("a move in a branch is paired within it, the selected branch with its story (phase 3)", async () => {
  const moved = (id, name) => `<w:p><w:moveFromRangeStart w:id="${id}" w:author="A" w:name="${name}"/><w:moveFrom w:id="${id + 1}" w:author="A"><w:r><w:t>moved</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="${id}"/></w:p>`
    + `<w:p>${run('between')}</w:p>`
    + `<w:p><w:moveToRangeStart w:id="${id + 2}" w:author="A" w:name="${name}"/><w:moveTo w:id="${id + 3}" w:author="A"><w:r><w:t>moved</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="${id + 2}"/></w:p>`;
  const pkg = await mini(AC(moved(10, 'move1'), moved(20, 'move1')));
  const changes = await pkg.getAllTrackedChanges();
  assert.deepEqual(changes.map((c) => c.kind), ['MOVE', 'MOVE']);
  changes[0].reject();
  const out = await bodyXmlOf(pkg);
  // rejected as a move in both branches: the source kept, the destination gone, the range markers with it
  assert.ok(!/move(From|To)/.test(out), out);
  assert.equal(count(out, '>moved<'), 2, out);
  for (const branch of [out.match(/<mc:Choice[^>]*>(.*?)<\/mc:Choice>/s)[1], out.match(/<mc:Fallback>(.*?)<\/mc:Fallback>/s)[1]]) {
    assert.ok(branch.indexOf('>moved<') < branch.indexOf('between'), branch);
  }
});

test("the bulk edits reach every branch: a tracked row or table deletion, the binding passes, a reply, inserted content (phase 3)", async () => {
  // a tracked table deletion marks every branch's rows and their content deleted
  let pkg = await mini(`<w:tbl>${TBL}${row(run('plain'))}${AC(row(run('C')), row(run('F')))}</w:tbl><w:p/>`);
  pkg.changeTrackingMode = 'TrackAll';
  (await pkg.getBody()).tables[0].delete();
  pkg.changeTrackingMode = 'Off';
  let out = await bodyXmlOf(pkg);
  assert.equal(count(out, '<w:trPr><w:del '), 3, out);
  for (const text of ['plain', 'C', 'F']) assert.ok(out.includes(`<w:delText>${text}</w:delText>`) || out.includes(`<w:delText xml:space="preserve">${text}</w:delText>`), text);
  assert.equal(await pkg.acceptAllRevisions(), 1);
  assert.ok(!(await bodyXmlOf(pkg)).includes('<w:tbl'), 'and accepted, the table goes');

  // the binding passes write every branch's controls and count what a reader sees; updateFromContentControls
  // reads the selected branch
  const ITEM = '{11111111-2222-3333-4444-555555555555}';
  const bound = (text) => `<w:sdt><w:sdtPr><w:dataBinding w:xpath="/data[1]/name[1]" w:storeItemID="${ITEM}"/><w:text/></w:sdtPr><w:sdtContent><w:p>${run(text)}</w:p></w:sdtContent></w:sdt>`;
  pkg = await mini(AC(bound('C'), bound('F')));
  pkg.customXmlParts.add('<data><name>Bound</name></data>', { itemID: ITEM });
  const result = await pkg.customXmlParts.applyBindings();
  assert.deepEqual([result.bound, result.updated], [1, 1]);
  out = await bodyXmlOf(pkg);
  assert.equal(count(out, '>Bound<'), 2, out);
  const body = await pkg.getBody();
  assert.deepEqual(body.contentControls.map((c) => c.text), ['Bound']);
  assert.deepEqual(body.contentControlsIn('every').map((c) => c.text), ['Bound', 'Bound']);
  const back = await pkg.customXmlParts.updateFromContentControls();
  assert.equal(back.bound, 1, 'the reverse pass reads what a reader sees');
  pkg.customXmlParts.removePart(pkg.customXmlParts.getItem(ITEM));
  assert.equal(count(await bodyXmlOf(pkg), 'dataBinding'), 0, 'the part removed: every branch\'s binding goes');

  // a reply is anchored in every branch its parent is
  const make = await mini(`<w:p>${run('anchored')}</w:p>`);
  await (await make.getBody()).paragraphs[0].getRange().insertComment('parent');
  const files = unzipSync(await make.save());
  let xml = strFromU8(files['word/document.xml']);
  const paragraph = xml.match(/<w:p\b(?:(?!<w:p\b).)*?anchored.*?<\/w:p>/s)[0];
  xml = xml.replace(paragraph, AC(paragraph, paragraph.replace('anchored', 'fallback')));
  if (!xml.includes(`xmlns:mc="${MC}"`)) xml = xml.replace('<w:document ', `<w:document xmlns:mc="${MC}" `);
  if (!xml.includes(`xmlns:w14="${W14}"`)) xml = xml.replace('<w:document ', `<w:document xmlns:w14="${W14}" `);
  files['word/document.xml'] = strToU8(xml);
  pkg = await WordprocessingMLPackage.load(zipSync(files));
  const [parent] = await (await pkg.getBody()).getComments();
  const reply = await parent.reply('the reply');
  out = await bodyXmlOf(pkg);
  assert.equal(count(out, `<w:commentRangeStart w:id="${reply.id}"/>`), 2, out);
  assert.equal(count(out, `<w:commentReference w:id="${reply.id}"/>`), 2, out);

  // a table inserted under tracking with a row-level element: every branch's row is an insertion
  pkg = await mini(`<w:p>${run('x')}</w:p>`);
  pkg.changeTrackingMode = 'TrackAll';
  await (await pkg.getBody()).insertXml(`<w:tbl xmlns:w="${W}" xmlns:mc="${MC}" xmlns:w14="${W14}">${TBL}${AC(row(run('C')), row(run('F')))}</w:tbl>`, 'End');
  pkg.changeTrackingMode = 'Off';
  out = await bodyXmlOf(pkg);
  assert.equal(count(out, '<w:trPr><w:ins '), 2, out);
  assert.equal(await pkg.acceptAllRevisions(), 1);
  assert.equal(count(await bodyXmlOf(pkg), '<w:ins '), 0);
});

test("the column passes reach every branch's cells: a cell-level element's branches each gain and lose a column (phase 3)", async () => {
  const probe = await load('mc-cell-newest-first.docx');
  let body = await probe.getBody();
  let table = body.tables[0];
  const branches = (xml) => [xml.match(/<mc:Choice[^>]*>(.*?)<\/mc:Choice>/gs) ?? [], xml.match(/<mc:Fallback>(.*?)<\/mc:Fallback>/gs) ?? []].flat();
  const before = branches(await documentXml(probe)).map((b) => count(b, '<w:tc>'));
  table.addColumns('End', 1, [['new']]);
  const after = branches(await documentXml(probe)).map((b) => count(b, '<w:tc>'));
  assert.deepEqual(after, before.map((n) => n + 1), 'each branch\'s cell list gained the new cell');
  assert.deepEqual(table.values, [['I prefer', 'w16sdtfl', 'new']]);
  table.deleteColumns(2);
  assert.deepEqual(branches(await documentXml(probe)).map((b) => count(b, '<w:tc>')), before, 'and lost it again');
  // a row-level element's rows gain a column too
  const pkg = await mini(`<w:tbl>${TBL}${row(run('plain'))}${AC(row(run('C')), row(run('F')))}</w:tbl><w:p/>`);
  table = (await pkg.getBody()).tables[0];
  table.addColumns('End', 1);
  const out = await bodyXmlOf(pkg);
  assert.equal(count(out, '<w:tc>') + count(out, '<w:tc '), 6, out);
  assert.deepEqual(table.values, [['plain', ''], ['C', '']]);
  // rows added beside the element go outside it, at either end
  table.addRows('Start', 1, [['first', '']]);
  table.addRows('End', 1, [['last', '']]);
  assert.deepEqual(table.values, [['first', ''], ['plain', ''], ['C', ''], ['last', '']]);
  const branchRows = (await bodyXmlOf(pkg)).match(/<mc:(Choice|Fallback)[^>]*>.*?<\/mc:(Choice|Fallback)>/gs).map((b) => count(b, '<w:tr>') + count(b, '<w:tr '));
  assert.deepEqual(branchRows, [1, 1]);
});

test("ids see every branch, a comment's paraId reads the selected branch, and toApiScript writes a branch Word reads (phase 3)", async () => {
  // a w14:paraId only in the Fallback still says the document uses them
  const pkg = await mini(`<w:p>${run('x')}</w:p>${AC(`<w:p>${run('C')}</w:p>`, `<w:p w14:paraId="12345678">${run('F')}</w:p>`)}`);
  const body = await pkg.getBody();
  const p = body.insertParagraph('new', 'End');
  assert.match(p.paraId ?? '', /^[0-9A-F]{8}$/);
  assert.notEqual(p.paraId, '12345678');
  // toApiScript: the element's selected branch through the verbs, a table with a row-level element as XML
  const { toApiScript } = await import('../dist/office-js/index.mjs');
  const script = await toApiScript(await (await mini(`${AC(`<w:p>${run('Choice text')}</w:p>`, `<w:p>${run('Fallback text')}</w:p>`)}<w:tbl>${TBL}${AC(row(run('C')), row(run('F')))}</w:tbl>`)).getBody());
  assert.ok(script.includes('Choice text') && !script.includes('Fallback text') && !/insertXml\(`<mc:AlternateContent/.test(script), script);
  assert.ok(/insertXml\(`<w:tbl[^`]*mc:AlternateContent/s.test(script), script);
});
