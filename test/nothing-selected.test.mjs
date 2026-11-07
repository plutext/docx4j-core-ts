// CR-029: an mc:AlternateContent with Choices, none understood, and no Fallback selects nothing (WD REQ-305), by the set Word was
// measured to understand (REQ-306), and says so (mce.nothing_selected). Word's rounds are in test/nothing-selected/, copied unedited
// with a SHA-1 manifest: each file read by both load routes is held to Word's text, and the warning to where Word shows neither branch.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strToU8, strFromU8, unzipSync } from 'fflate';
import {
  OpcPackage, WordprocessingMLPackage, Texts, PlainText, parseXml, preprocessAlternateContent, mceKeep, mcePreprocess, resolveAlternateContent, MCE_NS,
} from '../dist/index.mjs';
import { storedZip } from './helpers.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, 'nothing-selected');
const manifest = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8'));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const KEPT = {};
const RESOLVED = { mcePreprocess: true };
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';

const bytesOf = (name) => new Uint8Array(readFileSync(join(dir, name)));
const wordText = (bytes) => {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes[0] === 0xef && bytes[1] === 0xbb) return new TextDecoder().decode(bytes.subarray(3));
  return new TextDecoder('windows-1252').decode(bytes);
};
const norm = (text) => text.replace(/[\s\u0000-\u001f]+/g, ' ').trim();
const nothing = (result) => result.warnings.filter((w) => w.kind === 'mce.nothing_selected');
const documentOf = (bytes) => strFromU8(unzipSync(bytes)['word/document.xml']);
const read = async (bytes, options) => {
  const pkg = await OpcPackage.load(bytes, options);
  await pkg.unmarshalAll();
  return { pkg, result: await Texts.of(pkg) };
};
const p = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const ac = (requires, inner = p('HIDDEN')) => `<mc:AlternateContent><mc:Choice Requires="${requires}">${inner}</mc:Choice></mc:AlternateContent>`;
/** A package whose body is `inner`, `z` bound to a namespace nothing understands. */
const docx = (inner, { ns = 'xmlns:z="urn:docx4j:z"', extra = [], rels = '', types = '' } = {}) => storedZip([
  ['[Content_Types].xml', `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN}"/>${types}</Types>`],
  ['_rels/.rels', `${XML}<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
  ['word/document.xml', `${XML}<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:mc="${MCE_NS}" ${ns}><w:body>${inner}</w:body></w:document>`],
  ...(rels ? [['word/_rels/document.xml.rels', `${XML}<Relationships xmlns="${PR}">${rels}</Relationships>`]] : []),
  ...extra,
].map(([name, value]) => [name, strToU8(value)]));

test("CR-029: the rounds are as the manifest has them", () => {
  for (const f of manifest.files) assert.equal(createHash('sha1').update(bytesOf(f.name)).digest('hex'), f.sha1, f.name);
  assert.deepEqual(Object.fromEntries(Object.entries(manifest.rounds).map(([k, v]) => [k, v.length])), { ul: 119, uc: 13, c305: 9, rw: 2, te: 11 });
});

test("CR-029: Word's 142 files read as Word's text by both loads, and the warning stands exactly where Word shows neither branch", async () => {
  const index = new Map(strFromU8(bytesOf('understood-list.index.tsv')).split('\n').filter(Boolean).map((line) => line.split('\t')));
  const stems = [...manifest.rounds.ul, ...manifest.rounds.uc, ...manifest.rounds.c305, ...manifest.rounds.rw.filter((n) => n.startsWith('rw-25'))];
  assert.equal(stems.length, 142);
  let neither = 0;
  for (const stem of stems) {
    const word = wordText(bytesOf(`${stem}.txt`));
    const shown = word.includes('CHOICE WORDS');
    if (!shown) neither++;
    for (const [route, options] of [['kept', KEPT], ['resolved', RESOLVED]]) {
      const { pkg, result } = await read(bytesOf(`${stem}.docx`), options);
      assert.equal(norm(PlainText.of(result)), norm(word), `${stem} ${route}`);
      const warnings = nothing(result);
      assert.equal(warnings.length, shown ? 0 : 1, `${stem} ${route}`);
      assert.equal(pkg.warnings.filter((w) => w.code === 'mce.nothing_selected').length, shown ? 0 : 1, `${stem} ${route}: the part's`);
      if (shown) continue;
      assert.equal(warnings[0].selector, '/word/document.xml');
      assert.equal(warnings[0].address === undefined, route === 'resolved', `${stem}: an address where the walk meets it`);
      if (index.has(stem)) assert.equal(warnings[0].detail, index.get(stem), stem);
    }
  }
  assert.equal(neither, 24);
});

test('CR-029: at a root (rw-24) the element stays, reads empty as Word shows it, and is said, with no address', async () => {
  const stem = manifest.rounds.rw.find((n) => n.startsWith('rw-24'));
  assert.equal(norm(wordText(bytesOf(`${stem}.txt`))), '');
  for (const options of [KEPT, RESOLVED]) {
    const { pkg, result } = await read(bytesOf(`${stem}.docx`), options);
    assert.equal(norm(PlainText.of(result)), '');
    const [warning] = nothing(result);
    assert.equal(nothing(result).length, 1);
    assert.equal(warning.address, undefined);
    assert.equal(warning.selector, '/word/document.xml');
    const [record] = pkg.warnings.filter((w) => w.code === 'mce.nothing_selected');
    assert.deepEqual([record.root, record.removed], [true, false]);
    assert.match(documentOf(await pkg.save()), /AlternateContent/, 'a part keeps its root');
  }
});

test('CR-029: what a removal leaves (the tc-empty round): a cell gets a paragraph, as in Word\'s repaired save; a row and a table are left', async () => {
  for (const stem of ['te-01-tc-only', 'te-02-tc-only-one-cell', 'te-03-tc-no-branch']) {
    const { pkg } = await read(bytesOf(`${stem}.docx`), RESOLVED);
    const xml = documentOf(await pkg.save());
    assert.ok(!xml.includes('AlternateContent'), stem);
    const word = documentOf(bytesOf(`${stem}.word.docx`));
    const cells = (s) => [...s.matchAll(/<w:tc\b[^>]*>(.*?)<\/w:tc>/gs)].map((m) => /<w:p\b/.test(m[1]));
    assert.deepEqual(cells(xml), cells(word), `${stem}: a paragraph in every cell, as Word's save has`);
    assert.ok(cells(xml).every(Boolean));
  }
  // an element with no branch at all: no Choice, so no warning (REQ-305 names Choices), and its cell is kept too
  assert.deepEqual(nothing((await read(bytesOf('te-03-tc-no-branch.docx'), RESOLVED)).result), []);
  // a row left with no cell: a file Word refuses, source and save alike, so nothing is added
  const row = (await read(bytesOf('te-04-tr-only.docx'), RESOLVED)).pkg;
  assert.match(documentOf(await row.save()), /<w:tr\b[^>]*\/>|<w:tr\b[^>]*><\/w:tr>/);
  // a table left with no row, which Word drops silently, as its save has
  const table = (await read(bytesOf('te-05-tbl-only.docx'), RESOLVED)).pkg;
  assert.ok(!documentOf(await table.save()).includes('HIDDEN'));
  assert.ok(!documentOf(bytesOf('te-05-tbl-only.word.docx')).includes('<w:tbl'));
  // with a w:tcPr the paragraph follows it; a cell left holding something isn't touched
  const cell = (inner) => docx(`<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc>${inner}</w:tc></w:tr></w:tbl>${p('END')}`);
  const tcPr = '<w:tcPr><w:tcW w:w="100" w:type="dxa"/></w:tcPr>';
  const saved = documentOf(await (await read(cell(tcPr + ac('z')), RESOLVED)).pkg.save());
  assert.match(saved, /<w:tcPr>.*<\/w:tcPr><w:p\b/s);
  const kept = documentOf(await (await read(cell(ac('z') + p('STAYS')), RESOLVED)).pkg.save());
  assert.equal(kept.match(/<w:tc\b.*<\/w:tc>/s)[0].match(/<w:p\b/g).length, 1);
});

test('CR-029: the warning at each level: its address, selector and detail, one an element', async () => {
  const run = '<w:r><w:t>r</w:t></w:r>';
  const bytes = docx([
    ac('z'),                                                                                 // body/0, a block
    `<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>${ac('z', `<w:tr><w:tc>${p('H')}</w:tc></w:tr>`)}<w:tr>${ac('z', `<w:tc>${p('H')}</w:tc>`)}<w:tc>${p('IN')}</w:tc></w:tr></w:tbl>`,
    `<w:p>${run}${ac('z', '<w:r><w:t>H</w:t></w:r>')}</w:p>`,                                // among the runs
    `<w:p><w:r><w:t>r</w:t>${ac('z', '<w:t>H</w:t>')}</w:r></w:p>`,                          // in a run
  ].join(''));
  const warnings = nothing((await read(bytes, KEPT)).result);
  assert.deepEqual(warnings.map((w) => w.address), ['body/0', 'body/1/0', 'body/1/1/0', 'body/2', 'body/3'].filter((a) => warnings.some((w) => w.address === a)));
  assert.ok(warnings.length >= 5, warnings.map((w) => w.address).join());
  for (const w of warnings) {
    assert.equal(w.selector, '/word/document.xml');
    assert.equal(w.detail, 'urn:docx4j:z');
    assert.deepEqual(w.toJSON(), { kind: 'mce.nothing_selected', address: w.address, selector: '/word/document.xml', detail: 'urn:docx4j:z' });
  }
  assert.equal(norm(PlainText.of((await read(bytes, KEPT)).result)).includes('H'), false);
  // two at one address, requiring different namespaces, are two
  const two = await read(docx(`<w:p>${ac('z', run)}${ac('y', run)}</w:p>`, { ns: 'xmlns:z="urn:z" xmlns:y="urn:y"' }), KEPT);
  assert.deepEqual(nothing(two.result).map((w) => w.detail), ['urn:z', 'urn:y']);
});

test('CR-029: detail: what the Choices require, once each; an unbound prefix has none; nothing is said where a branch is selected or there is none', async () => {
  const detail = async (inner, ns) => nothing((await read(docx(inner, { ns }), KEPT)).result).map((w) => w.detail);
  assert.deepEqual(await detail(`<mc:AlternateContent><mc:Choice Requires="a">${p('H')}</mc:Choice><mc:Choice Requires="b">${p('H')}</mc:Choice></mc:AlternateContent>`, 'xmlns:a="urn:a" xmlns:b="urn:b"'), ['urn:a urn:b']);
  assert.deepEqual(await detail(`<mc:AlternateContent><mc:Choice Requires="a">${p('H')}</mc:Choice><mc:Choice Requires="a2">${p('H')}</mc:Choice></mc:AlternateContent>`, 'xmlns:a="urn:a" xmlns:a2="urn:a"'), ['urn:a']);
  assert.deepEqual(await detail(ac('w a'), 'xmlns:a="urn:a"'), [`${W} urn:a`], 'one understood is not all');
  assert.deepEqual(await detail(ac('unbound'), ''), ['']);
  assert.deepEqual(await detail(`<mc:AlternateContent><mc:Choice Requires="q" xmlns:q="urn:q">${p('H')}</mc:Choice></mc:AlternateContent>`, ''), ['urn:q']);
  // an understood namespace under another prefix is the Choice, by URI (REQ-303)
  const q = await read(docx(`<mc:AlternateContent><mc:Choice Requires="q" xmlns:q="http://schemas.microsoft.com/office/word/2010/wordml">${p('SHOWN')}</mc:Choice></mc:AlternateContent>`, { ns: '' }), KEPT);
  assert.equal(norm(PlainText.of(q.result)), 'SHOWN');
  assert.deepEqual(nothing(q.result), []);
  assert.deepEqual(await detail(`<mc:AlternateContent><mc:Choice Requires="z">${p('H')}</mc:Choice><mc:Fallback>${p('F')}</mc:Fallback></mc:AlternateContent>`), []);
  assert.deepEqual(await detail('<mc:AlternateContent/>'), [], 'no branches: not REQ-305\'s case');
  // nested in the Fallback a reader takes: said, and its own words stay absent
  const nested = await read(docx(`<mc:AlternateContent><mc:Choice Requires="z">${p('H')}</mc:Choice><mc:Fallback>${ac('z')}</mc:Fallback></mc:AlternateContent>`), KEPT);
  assert.equal(nothing(nested.result).length, 1);
  assert.equal(norm(PlainText.of(nested.result)), '');
  // under warnings: false, by either route
  const quiet = await OpcPackage.load(docx(ac('z')));
  assert.deepEqual(nothing(await Texts.of(quiet, { warnings: false })), []);
  const quietResolved = await OpcPackage.load(docx(ac('z')), RESOLVED);
  assert.deepEqual(nothing(await Texts.of(quietResolved, { warnings: false })), []);
});

test('CR-029: a story beyond the main one: a header\'s element says that part, on a kept and a resolved load', async () => {
  const header = `${XML}<w:hdr xmlns:w="${W}" xmlns:mc="${MCE_NS}" xmlns:z="urn:docx4j:z">${p('HEADER')}${ac('z')}</w:hdr>`;
  const bytes = docx(`${p('BODY')}<w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr>`, {
    rels: `<Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/>`,
    types: '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    extra: [['word/header1.xml', header]],
  });
  for (const options of [KEPT, RESOLVED]) {
    const pkg = await OpcPackage.load(bytes, options);
    assert.deepEqual(nothing(await Texts.of(pkg, { include: ['headers'] })).map((w) => w.selector), ['/word/header1.xml']);
    // not read, not said
    assert.deepEqual(nothing(await Texts.of(await OpcPackage.load(bytes, options), { include: [] })), []);
  }
});

test('CR-029: the pass: removal at every level, and where it leaves the element whole', () => {
  const wrap = (inner) => `<w:p xmlns:w="${W}" xmlns:mc="${MCE_NS}" xmlns:z="urn:z">${inner}</w:p>`;
  // in a run, as Word reads it (c305-*)
  const doc = parseXml(wrap(`<w:r><w:t>a</w:t>${ac('z', '<w:t>H</w:t>')}</w:r>`));
  const result = preprocessAlternateContent(doc);
  assert.equal(result.resolved, 1);
  assert.equal(doc.getElementsByTagNameNS(MCE_NS, 'AlternateContent').length, 0);
  assert.deepEqual(result.warnings.map((w) => [w.code, w.namespaces, w.removed, w.root]), [['mce.nothing_selected', ['urn:z'], true, false]]);
  // resolve: false changes nothing, and reports the same element as not removed
  const kept = parseXml(wrap(ac('z', '<w:r/>')));
  const report = preprocessAlternateContent(kept, { resolve: false });
  assert.equal(kept.getElementsByTagNameNS(MCE_NS, 'AlternateContent').length, 1);
  assert.deepEqual(report.warnings.map((w) => w.removed), [false]);
  // a part of another kind: left whole and said nothing of (D3), by the part's content type
  const part = (contentType) => ({ contentType, partName: { name: '/x.xml' } });
  const slide = part('application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
  const other = parseXml(wrap(ac('z', '<w:r/>')));
  assert.deepEqual(mcePreprocess(other, slide), []);
  assert.deepEqual(mceKeep(other, slide), []);
  assert.equal(other.getElementsByTagNameNS(MCE_NS, 'AlternateContent').length, 1);
  for (const contentType of [MAIN, 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml', 'application/vnd.ms-word.document.macroEnabled.main+xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml']) {
    const word = parseXml(wrap(ac('z', '<w:r/>')));
    assert.equal(mcePreprocess(word, part(contentType)).length, 1, contentType);
    assert.equal(word.getElementsByTagNameNS(MCE_NS, 'AlternateContent').length, 0, contentType);
  }
  // a KEPT_HOSTS host comes first, whatever the part
  const sheet = parseXml(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="${MCE_NS}" xmlns:z="urn:z"><mc:AlternateContent><mc:Choice Requires="z"><x/></mc:Choice></mc:AlternateContent></workbook>`);
  assert.equal(resolveAlternateContent(sheet), 0);
  assert.equal(sheet.getElementsByTagNameNS(MCE_NS, 'AlternateContent').length, 1);
  // a nested element that selects nothing gives no words to mce.branches_differ, where it gave its first Choice's
  const nested = parseXml(wrap(`<mc:AlternateContent><mc:Choice Requires="w">${ac('z', '<w:r><w:t>HIDDEN</w:t></w:r>')}</mc:Choice><mc:Fallback><w:r><w:t>F</w:t></w:r></mc:Fallback></mc:AlternateContent>`));
  const differ = preprocessAlternateContent(nested, { resolve: false }).warnings.find((w) => w.code === 'mce.branches_differ');
  assert.deepEqual(differ.branches.map((b) => b.text), ['', 'F']);
});

test('CR-029: the default load keeps the element: an untouched part byte for byte, a save with every branch, a mutator reaching it', async () => {
  const bytes = bytesOf('rw-25-body-choice-no-fallback.docx');
  const source = unzipSync(bytes)['word/document.xml'];
  const pkg = await WordprocessingMLPackage.load(bytes);
  await Texts.of(pkg);
  assert.deepEqual(unzipSync(await pkg.save())['word/document.xml'], source);
  const edited = await WordprocessingMLPackage.load(bytes);
  (await edited.getBody()).insertParagraph('MORE', 'End');
  const xml = documentOf(await edited.save());
  assert.ok(xml.includes('CHOICE WORDS') && xml.includes('AlternateContent') && xml.includes('MORE'));
  // a tracked change in a branch no reader sees is still resolved by a resolve-all, which walks every branch
  const ins = `<w:ins w:id="1" w:author="A" w:date="2026-11-07T00:00:00Z"><w:r><w:t>H</w:t></w:r></w:ins>`;
  const tracked = await WordprocessingMLPackage.load(docx(`${p('BODY')}${ac('z', `<w:p>${ins}</w:p>`)}`));
  await tracked.getBody();
  await tracked.acceptAllRevisions();
  const after = documentOf(await tracked.save());
  assert.ok(after.includes('AlternateContent') && !after.includes('<w:ins'));
});

test("CR-029: wpi, the one namespace new to the set the corpus exercises: the Choice read by both loads, and a resolved save keeps the ink", async () => {
  const bytes = new Uint8Array(readFileSync(join(here, 'text-extraction', 'suite', 'phase0', 'extraction-probe-inspect.docx')));
  const count = (xml) => (xml.match(/<w14:contentPart\b/g) ?? []).length;
  const source = count(documentOf(bytes));
  assert.ok(source > 0);
  const { pkg } = await read(bytes, RESOLVED);
  assert.deepEqual(pkg.warnings.filter((w) => w.code === 'part.skipped'), []);
  assert.equal(count(documentOf(await pkg.save())), source);
});
