// CR-016 phase 2 (2026-10-28): a child its container never admits (docx4j CR-064's rules, in the objects facade since its
// CR-017, objects 0.7.0) is refused at save with its path and docx4j's fix (D2), written as held and reported under
// pkg.checkContentOnSave = false, found on request by checkContent() without a change (D6), and a target.misplaced part
// warning on load (D7, docx4j-python's code and fields). The gate (D6): every fixture saves under the check unrefused.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { p, r, t, tbl, tc } from '@docx4j/generated-objects-ts/builders/wml';
import { WordprocessingMLPackage, OpcPackage, MisplacedContentException, Docx4JException } from '../dist/index.mjs';
import { check72Files } from './fixtures/check72/build.mjs';
import { fixturesDir } from './helpers.mjs';

const documentXml = (bytes) => strFromU8(unzipSync(bytes)['word/document.xml']);

/** A created package whose body holds one paragraph, with the tree's arrays to edit directly, as a JavaScript caller can. */
async function created() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  body.insertParagraph('BEFORE', 'Start');
  const content = pkg.getMainDocumentPart().contents.body.content;
  return { pkg, content, paragraph: content.find((c) => c.name.localPart === 'p') };
}

/** CR-016 section 2's table: each direct edit of the tree, what the save is refused for, and what the opt-out writes. */
const EDITS = [
  { name: 'a run pushed into the body', edit: ({ content }) => content.push(r('MISPLACED')),
    finding: { holder: 'Body', item: 'R', element: 'w:r', index: 1, path: '/w:document/w:body[1]/w:r[1]' }, written: '</w:p><w:r><w:t>MISPLACED</w:t></w:r>' },
  { name: "a table pushed into a paragraph's content", edit: ({ paragraph }) => paragraph.value.content.push(tbl([['MISPLACED']])),
    finding: { holder: 'P', item: 'Tbl', element: 'w:tbl', index: 1, path: '/w:document/w:body[1]/w:p[1]/w:tbl[1]' }, written: '<w:p><w:r><w:t>BEFORE</w:t></w:r><w:tbl>' },
  { name: "a paragraph pushed into a paragraph's content", edit: ({ paragraph }) => paragraph.value.content.push(p('MISPLACED')),
    finding: { holder: 'P', item: 'P', element: 'w:p', index: 1, path: '/w:document/w:body[1]/w:p[1]/w:p[1]' }, written: '<w:p><w:r><w:t>BEFORE</w:t></w:r><w:p><w:r><w:t>MISPLACED</w:t></w:r></w:p></w:p>' },
  { name: "a cell pushed straight into a table's content", edit: ({ content }) => { const table = tbl([['A1']]); table.value.content.push(tc('MISPLACED')); content.push(table); },
    finding: { holder: 'Tbl', item: 'Tc', element: 'w:tc', index: 1, path: '/w:document/w:body[1]/w:tbl[1]/w:tc[1]' }, written: '</w:tr><w:tc><w:p><w:r><w:t>MISPLACED</w:t></w:r></w:p></w:tc></w:tbl>' },
  { name: "a string pushed into a run's content", edit: ({ paragraph }) => paragraph.value.content[0].value.content.push('MISPLACED'),
    finding: { holder: 'R', index: 1, path: '/w:document/w:body[1]/w:p[1]/w:r[1]' }, written: null },
];

test("CR-016 D2: a direct edit of the tree is refused at save, nothing written, with the part, the child's path and docx4j's fix; the content API's inserts refuse before it", async () => {
  for (const { name, edit, finding } of EDITS) {
    const made = await created();
    edit(made);
    const error = await made.pkg.save().then(() => undefined, (e) => e);
    assert.ok(error instanceof MisplacedContentException, `${name}: refused (${error})`);
    assert.ok(error instanceof Docx4JException);
    assert.equal(error.name, 'MisplacedContentException');
    assert.equal(error.partName, '/word/document.xml', name);
    assert.equal(error.findings.length, 1, name);
    const [f] = error.findings;
    assert.equal(f.code, 'target.misplaced');
    assert.equal(f.partName, '/word/document.xml');
    assert.equal(f.rule, 'misplaced');
    for (const [k, v] of Object.entries(finding)) assert.equal(f[k], v, `${name}: ${k}`);
    if (finding.element === undefined) { assert.equal(f.item, undefined); assert.equal(f.element, undefined); assert.equal(f.text, undefined, 'a save never sees the string'); }
    assert.equal(f.line, undefined, 'a save gives the path, not a line');
    assert.ok(f.message.startsWith(`/word/document.xml: ${finding.path}: You can't `), `${name}: ${f.message}`);
    assert.ok(error.message.startsWith('Save refused: /word/document.xml holds a child its container never admits, which Word 365 refuses the file for or repairs by dropping content (set pkg.checkContentOnSave = false to write it as held):\n  /word/document.xml: '), error.message);
    assert.ok(error.message.includes(f.message), 'each finding in the message');
    assert.equal(made.pkg.lastSave, undefined, `${name}: no save happened`);
    if (finding.element === undefined) assert.ok(error.cause instanceof Error, 'the runtime refused the string, after the walk reported it');
  }
  // several in one part: every one listed
  const made = await created();
  made.content.push(r('ONE'), t('TWO'));
  const error = await made.pkg.save().then(() => undefined, (e) => e);
  assert.ok(error instanceof MisplacedContentException);
  assert.deepEqual(error.findings.map((f) => [f.item, f.index, f.path]), [['R', 1, '/w:document/w:body[1]/w:r[1]'], ['Text', 2, '/w:document/w:body[1]/w:t[1]']]);
  assert.ok(error.message.includes('holds 2 children their containers never admit'), error.message);
  // the content API refuses the same at the insert, before anything changes (section 2, as before)
  const body = await (await created()).pkg.getBody();
  assert.throws(() => body.insertElement(r('x')), Docx4JException);
});

test('CR-016 D2: with pkg.checkContentOnSave = false the part is written as held and each misplaced child is on pkg.lastSave.misplaced; a string in a run still fails, naming the cause', async () => {
  for (const { name, edit, finding, written } of EDITS) {
    const made = await created();
    made.pkg.checkContentOnSave = false;
    edit(made);
    if (written === null) {
      const error = await made.pkg.save().then(() => undefined, (e) => e);
      assert.ok(error instanceof Docx4JException && !(error instanceof MisplacedContentException), `${name}: the runtime's refusal, not the check's`);
      assert.ok(error.message.startsWith("Problem saving part /word/document.xml: /word/document.xml: /w:document/w:body[1]/w:p[1]/w:r[1]: You can't add a string to a run (w:r)"), error.message);
      continue;
    }
    const xml = documentXml(await made.pkg.save());
    assert.ok(xml.includes(written), `${name}: written as held\n${xml}`);
    assert.deepEqual(made.pkg.lastSave.misplaced.map((f) => [f.code, f.holder, f.item, f.path]), [['target.misplaced', finding.holder, finding.item, finding.path]], name);
    assert.deepEqual(made.pkg.lastSave.outOfRange, []);
  }
  // nothing misplaced: an empty report either way
  const clean = await created();
  clean.pkg.checkContentOnSave = false;
  await clean.pkg.save();
  assert.deepEqual(clean.pkg.lastSave.misplaced, []);
  // the setting travels with a clone
  const made = await created();
  made.pkg.checkContentOnSave = false;
  assert.equal((await made.pkg.clone()).checkContentOnSave, false);
  assert.equal((await (await created()).pkg.clone()).checkContentOnSave, true);
});

test("CR-016 D6: checkContent() on a part or the package finds what a save would refuse, changes nothing, and leaves an untouched part untouched; a document read and saved untouched is never refused", async () => {
  for (const f of await check72Files()) {
    const pkg = await WordprocessingMLPackage.load(f.bytes);
    const main = pkg.getMainDocumentPart();
    const found = await main.checkContent();
    const isString = f.name === '72m-string-in-run';
    // the string is gone at the read, so a tree holds nothing to find; its only notice is the load warning (D7)
    assert.equal(found.length, isString ? 0 : 1, f.name);
    assert.equal(main.isUntouched, true, `${f.name}: the check reads privately`);
    assert.deepEqual(await pkg.checkContent(), found, `${f.name}: the package's check is its parts'`);
    if (!isString) {
      assert.equal(found[0].code, 'target.misplaced');
      assert.ok(found[0].path.startsWith('/w:document/w:body[1]/'), found[0].path);
      assert.equal(found[0].line, undefined);
    }
    // untouched, the part is copied from the source, so the save writes it as it came, refused or checked by nothing
    const saved = documentXml(await pkg.save());
    assert.equal(saved, strFromU8(unzipSync(f.bytes)['word/document.xml']), `${f.name}: byte for byte`);
    assert.deepEqual(pkg.lastSave.misplaced, []);
    // once unmarshalled, the save is refused for the same finding
    await main.getContents();
    if (isString) await pkg.save();
    else {
      const error = await pkg.save().then(() => undefined, (e) => e);
      assert.ok(error instanceof MisplacedContentException, f.name);
      assert.deepEqual(error.findings, found, `${f.name}: the save's findings are the check's`);
    }
  }
  // a direct edit, found before any save is tried
  const made = await created();
  made.content.push(r('MISPLACED'));
  assert.deepEqual((await made.pkg.checkContent()).map((f) => f.path), ['/w:document/w:body[1]/w:r[1]']);
  assert.deepEqual(await (await created()).pkg.checkContent(), []);
});

/** D7's fields for each check 72 file's one warning: the container's class, the child's, its element and its place in the list. */
const LOAD_WARNINGS = {
  '72a-run-in-body': ['Body', 'R', 'w:r', 2], '72b-text-in-body': ['Body', 'Text', 'w:t', 2], '72c-cell-in-body': ['Body', 'Tc', 'w:tc', 2],
  '72d-row-in-body': ['Body', 'Tr', 'w:tr', 2], '72e-break-in-body': ['Body', 'Br', 'w:br', 2], '72f-run-in-cell': ['Tc', 'R', 'w:r', 0],
  '72g-table-in-paragraph': ['P', 'Tbl', 'w:tbl', 1], '72h-paragraph-in-paragraph': ['P', 'P', 'w:p', 1], '72i-text-in-paragraph': ['P', 'Text', 'w:t', 1],
  '72j-break-in-paragraph': ['P', 'Br', 'w:br', 1], '72k-run-in-run': ['R', 'R', 'w:r', 0], '72l-paragraph-in-run': ['R', 'P', 'w:p', 1],
  '72m-string-in-run': ['R', undefined, undefined, 0], '72n-cell-in-table': ['Tbl', 'Tc', 'w:tc', 0], '72o-paragraph-in-table': ['Tbl', 'P', 'w:p', 0],
  '72p-table-in-table': ['Tbl', 'Tbl', 'w:tbl', 0], '72q-paragraph-in-row': ['Tr', 'P', 'w:p', 0], '72r-run-in-row': ['Tr', 'R', 'w:r', 0],
};

test("CR-016 D7: a target.misplaced warning per misplaced child read, docx4j-python's fields, docx4j's message after the part and the place, a string's text carried", async () => {
  for (const f of await check72Files()) {
    const pkg = await WordprocessingMLPackage.load(f.bytes);
    const main = pkg.getMainDocumentPart();
    await main.getContents();
    const [holder, item, element, index] = LOAD_WARNINGS[f.name];
    assert.equal(main.warnings.length, 1, f.name);
    const [w] = main.warnings;
    assert.equal(w.code, 'target.misplaced');
    assert.equal(w.partName, '/word/document.xml');
    assert.deepEqual([w.holder, w.item, w.element, w.index, w.rule], [holder, item, element, index, 'misplaced'], f.name);
    assert.equal(w.path, undefined, 'a load gives the place in the source, not a path');
    assert.equal(typeof w.line, 'number', `${f.name}: xmldom's line`);
    assert.equal(typeof w.column, 'number');
    assert.equal(w.text, f.name === '72m-string-in-run' ? 'MISPLACED string in a run' : undefined, f.name);
    const container = { Body: 'w:body', Tc: 'w:tc', P: 'w:p', R: 'w:r', Tbl: 'w:tbl', Tr: 'w:tr' }[holder];
    const shown = element ?? 'a string "MISPLACED string in a run"';
    assert.ok(w.message.startsWith(`/word/document.xml: ${container}, child ${index} (${shown}) (line ${w.line}, column ${w.column}): You can't `), `${f.name}: ${w.message}`);
    assert.deepEqual(pkg.warnings, main.warnings, 'the package gathers it');
  }
  // a part whose content is set again starts clean
  const pkg = await WordprocessingMLPackage.load((await check72Files())[0].bytes);
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  main.setContents(main.contents);
  assert.deepEqual(main.warnings, []);
});

test('CR-016 D6, the gate: every fixture loads, unmarshals every part and saves under the check unrefused, and none warns at the load (the kept-branch fixtures included)', async () => {
  const files = [];
  const walk = (dir) => {
    for (const name of readdirSync(dir)) {
      const path = join(dir, name);
      if (statSync(path).isDirectory()) { if (name !== 'check72') walk(path); } else if (/\.(docx|pptx|xlsx)$/i.test(name)) files.push(path);
    }
  };
  walk(fixturesDir);
  walk(join(fixturesDir, '..', 'mc-fun'));
  assert.ok(files.length > 300, `${files.length} files`);
  const refused = [];
  const found = [];
  for (const file of files) {
    const pkg = await OpcPackage.load(new Uint8Array(readFileSync(file)));
    await pkg.unmarshalAll();
    assert.equal(pkg.checkContentOnSave, true);
    try { await pkg.save(); } catch (e) { refused.push(`${file}: ${e.message.split('\n')[0]}`); continue; }
    if (pkg.lastSave.misplaced.length > 0) found.push(file);
    for (const w of pkg.warnings) if (w.code === 'target.misplaced') found.push(`${file} (load): ${w.message}`);
  }
  assert.deepEqual(refused, [], 'no legal document refused (a rule that trips here reports instead, CR-064 D8)');
  assert.deepEqual(found, []);
});

// A gap in the objects facade's walk (found here 2026-10-28, sent to the objects session for its CR-017): a container whose
// value has no TYPE_NAME, as a value built by hand without its type (CR-016 section 2's second escape), is not tested, where
// the marshal still writes it by its element's declared class. This asserts the gap as it stands (CR-001 section 19): the
// facade's fix announces itself here by failing, and the assertion then flips. createPackage's document and body are typed
// since phase 2 so that a created document's body is a container the walk sees.
test("the objects facade's walk skips a container built without its TYPE_NAME (known gap, asserted as it stands)", async () => {
  const made = await created();
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  made.content.push({ name: { namespaceURI: W, localPart: 'p' }, value: { content: [r('x'), tbl([['MISPLACED']])] } });
  assert.deepEqual(await made.pkg.checkContent(), [], 'the untyped paragraph is not tested (objects 0.7.0)');
  const xml = documentXml(await made.pkg.save());
  assert.ok(xml.includes('<w:p><w:r><w:t>x</w:t></w:r><w:tbl>'), 'written by the element\'s declared class, as held');
  // typed, the same paragraph is refused (the builder itself refuses the table, objects CR-017, so the push goes past it)
  const typed = await created();
  const paragraph = p([r('x')]);
  assert.throws(() => p([r('x'), tbl([['MISPLACED']])]), TypeError, 'the builder refuses it first');
  paragraph.value.content.push(tbl([['MISPLACED']]));
  typed.content.push(paragraph);
  assert.deepEqual((await typed.pkg.checkContent()).map((f) => f.path), ['/w:document/w:body[1]/w:p[2]/w:tbl[1]']);
});
