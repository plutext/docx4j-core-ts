// CR-002 section 61, a bound block control of plain text holds one paragraph. Check 56 (test/README.md) measures it:
// four files to open, each made by this package's own writers, and a file for a Script Lab snippet that gives each
// kind of bound block control Office JS's block writes. The snippet's file and the snippet must agree, since Word runs
// one on the other and nothing can be fixed after the run.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../dist/index.mjs';
import { Word, NotSupportedError } from '../dist/office-js/index.mjs';
import { check56Bytes, check56Cases, check56OpenFiles, KINDS, STORE_ITEM_ID } from './fixtures/check56/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));

async function loaded(bytes) {
  const pkg = await WordprocessingMLPackage.load(bytes);
  await pkg.customXmlParts.load();
  return pkg;
}

/** The control's type element, or 'untyped'. */
function kindOf(control) {
  const pr = control.sdt.sdtPr;
  const names = (pr?.rPrOrAliasOrLock ?? []).map((e) => e.name?.localPart);
  if (names.includes('richText')) return 'richtext';
  if (names.includes('text')) return control.isMultiLine ? 'text-multi' : 'text-single';
  return 'untyped';
}

test("check 56's snippet file and snippet agree: every case parses to the case it was built as, each control bound and of its kind", async () => {
  const source = await readFile(join(here, 'fixtures', 'check56', 'check56-script-lab.js'), 'utf8');
  const { parseRegion, runCheck } = new Function(`${source}\nreturn CHECK56;`)();
  const pkg = await loaded(await check56Bytes());
  const body = await pkg.getBody();
  const regions = body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null);
  const cases = check56Cases();
  assert.equal(cases.length, 16);
  assert.deepEqual(regions, cases.map(({ n, op, kind }) => ({ n, op, kind })));
  const data = pkg.customXmlParts.getItem(STORE_ITEM_ID);
  for (const c of cases) {
    const control = body.contentControls.find((x) => x.tag === `c56-${c.n}`);
    assert.equal(control.form, 'Block', `case ${c.n}`);
    assert.ok(control.xmlMapping.isMapped, `case ${c.n}`);
    assert.equal(kindOf(control), c.kind, `case ${c.n}`);
    assert.deepEqual(control.paragraphs.map((p) => p.text), [`start${c.n}`]);
    assert.equal(data.selectSingleNode(`/ns0:check56/ns0:value${c.n}`, "xmlns:ns0='urn:check56'").text, `start${c.n}`);
  }
  // through the Word shim, with no error of the snippet's own: what fails is what the shim doesn't have, and the
  // multi-line control's insertParagraph, refused (CR-002 section 61 H2)
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  assert.equal(ours.cases.length, 16);
  const errors = new Set(ours.cases.flatMap((c) => [c.error, ...Object.values(c.lookups), ...Object.values(c.region)])
    .filter((v) => typeof v === 'string' && v.startsWith('error:')).map((v) => v.split(': this control')[0]));
  assert.deepEqual([...errors].sort(), [
    'error: NotImplemented: Word.ContentControl.insertOoxml is not supported by @docx4j/core-ts',
    'error: NotImplemented: Word.ContentControl.insertParagraph is not supported by @docx4j/core-ts',
    'error: NotImplemented: Word.ContentControl.insertTable is not supported by @docx4j/core-ts',
    'error: NotImplemented: Word.Document.customXmlParts is not supported by @docx4j/core-ts',
    'error: NotImplemented: Word.Paragraph.parentContentControlOrNullObject is not supported by @docx4j/core-ts',
  ]);
});

test("check 56's files to open: each a bound block control of its kind, given a second block by this package's writers", async () => {
  const files = await check56OpenFiles();
  assert.deepEqual(files.map((f) => f.name), ['56a-untyped-two-paragraphs.docx', '56b-text-two-paragraphs.docx', '56c-richtext-two-paragraphs.docx', '56d-untyped-table.docx']);
  for (const { name, kind, add, bytes } of files) {
    const pkg = await loaded(bytes);
    const control = (await pkg.getBody()).contentControls.find((x) => x.tag === 'c56-1');
    assert.ok(control.xmlMapping.isMapped, name);
    assert.equal(kindOf(control), kind, name);
    const blocks = control.sdt.sdtContent.content.map((e) => e.name.localPart);
    assert.deepEqual(blocks, add === 'table' ? ['tbl', 'p'] : ['p', 'p'], name);
    assert.ok(Object.hasOwn(KINDS, kind));
  }
});

test("check 56's run: Word's JSON is of this file's cases, and its saves and repairs are Word's, as copied", async () => {
  const { createHash } = await import('node:crypto');
  const manifest = JSON.parse(await readFile(join(here, 'check56', 'manifest.json'), 'utf8'));
  assert.deepEqual(manifest.files.map((f) => f.name), [
    '56a-untyped-two-paragraphs-repaired.docx', '56b-text-two-paragraphs-repaired.docx', '56c-richtext-two-paragraphs-repaired.docx',
    '56d-untyped-table-repaired.docx', '56-writes-word365.docx', '56-writes-word365-repaired.docx',
  ]);
  for (const { name, sha1 } of manifest.files) {
    assert.equal(createHash('sha1').update(await readFile(join(here, 'check56', name))).digest('hex'), sha1, name);
  }
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check56', '56-writes-word365-run.json'), 'utf8'));
  assert.equal(word.host.version, '16.0.20326.20232');
  assert.deepEqual(word.cases.map(({ n, op, kind }) => ({ n, op, kind })), check56Cases().map(({ n, op, kind }) => ({ n, op, kind })));
  // each repair of a file to open kept one paragraph, a single-line or multi-line w:text control refilled from its node
  for (const name of manifest.files.map((f) => f.name).filter((n) => /^56[a-d]-/.test(n))) {
    const pkg = await loaded(await readFile(join(here, 'check56', name)));
    const control = (await pkg.getBody()).contentControls.find((x) => x.tag === 'c56-1');
    assert.equal(kindOf(control), name.startsWith('56b') ? 'text-multi' : 'text-single', name);
    assert.deepEqual(control.sdt.sdtContent.content.map((e) => e.name.localPart), ['p'], name);
    assert.equal(control.text, 'start1', name);
  }
});

/** A case's paragraphs and control as the snippet reports them, for comparing two runs. */
const seen = (c) => ({ paragraphs: c.region.paragraphs, type: c.region.control?.type, text: c.region.control?.text, error: c.error?.split(':')[1]?.trim() });

test("CR-002 section 61: check 56's snippet through the Word shim gives Word 365's run, case for case, but for H2's refusal and the writes the shim hasn't", async () => {
  const source = await readFile(join(here, 'fixtures', 'check56', 'check56-script-lab.js'), 'utf8');
  const { runCheck } = new Function(`${source}\nreturn CHECK56;`)();
  const pkg = await loaded(await check56Bytes());
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check56', '56-writes-word365-run.json'), 'utf8'));
  for (const theirs of word.cases) {
    const mine = ours.cases.find((c) => c.n === theirs.n);
    if (theirs.op === 'paragraph-end' && theirs.kind === 'text-multi') {
      // H2: Word wrote a second paragraph, and then couldn't open its own save; here it is refused
      assert.equal(seen(theirs).paragraphs.length, 4);
      assert.equal(seen(mine).error, 'NotImplemented');
      assert.deepEqual(seen(mine).paragraphs, [`BEGIN${theirs.n} paragraph-end text-multi`, `start${theirs.n}`, `END${theirs.n}`]);
    } else if (theirs.op === 'table-end' || theirs.op === 'ooxml-end') {
      // the shim's ContentControl has no insertTable or insertOoxml; the content API's refusals are the next test's
      assert.equal(seen(theirs).error, theirs.op === 'table-end' ? 'InvalidArgument' : 'GeneralException');
      assert.equal(seen(mine).error, 'NotImplemented');
      assert.deepEqual(seen(mine).paragraphs, seen(theirs).paragraphs, `case ${theirs.n}`);
    } else {
      assert.deepEqual(seen(mine), seen(theirs), `case ${theirs.n}`);
    }
  }
});

/** A bound block control of a kind in a package of its own: its paragraph, its control, the package. */
async function boundControl(kind) {
  const c = check56Cases().find((x) => x.kind === kind && x.op === 'paragraph-end');
  const pkg = await loaded(await check56Bytes());
  const control = (await pkg.getBody()).contentControls.find((x) => x.tag === `c56-${c.n}`);
  return { pkg, control, paragraph: control.paragraphs[0], n: c.n };
}

test('CR-002 section 61: a bound block control holds one paragraph; every write that would give it another block is refused before anything changes, with Word\'s error where measured', async () => {
  for (const kind of Object.keys(KINDS)) {
    const { pkg, control, paragraph, n } = await boundControl(kind);
    const before = await pkg.getMainDocumentPart().getXml();
    const body = paragraph.parentBody;
    // H4: every kind reads as single-line plain text but w:text multi-line
    assert.equal(control.type, 'PlainText', kind);
    assert.equal(control.isMultiLine, kind === 'text-multi', kind);
    assert.ok(body.getContentControls({ types: ['PlainText'] }).length >= 0);
    // H3: Word's codes where measured, NotSupportedError where Word wasn't given the write
    assert.throws(() => body.insertTable(1, 2, 'End', [['A', 'B']]), (e) => e.code === 'InvalidArgument', kind);
    await assert.rejects(body.insertOoxml('<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>', 'End'), (e) => e.code === 'GeneralException', kind);
    await assert.rejects(body.insertXml('<w:p><w:r><w:t>A</w:t></w:r></w:p>', 'End'), (e) => e.code === 'GeneralException', kind);
    assert.throws(() => paragraph.insertParagraph('x', 'After'), NotSupportedError, kind);
    assert.throws(() => paragraph.getRange().search(`start${n}`)[0].insertParagraph('x', 'Before'), NotSupportedError, kind);
    assert.throws(() => paragraph.insertContentControl(), NotSupportedError, kind);
    assert.equal(await pkg.getMainDocumentPart().getXml(), before, `${kind}: nothing changed`);
    // H2: insertParagraph joins a single-line control's paragraph, and is refused in a multi-line one
    if (kind === 'text-multi') assert.throws(() => control.insertParagraph('NEWPARA', 'End'), NotSupportedError);
    else {
      const p = control.insertParagraph('NEWPARA', 'End');
      assert.equal(p.text, `start${n}NEWPARA`);
      assert.equal(control.paragraphs.length, 1);
      assert.equal(control.xmlMapping.customXmlNode.text, `start${n}NEWPARA`, `${kind}: the node follows the control (F3)`);
    }
    // section 61 item 3: a line feed is a w:br in a multi-line control and goes in a single-line one
    control.insertText('x\ny', 'End');
    assert.equal(control.paragraphs.length, 1, kind);
    assert.match(control.paragraphs[0].text, kind === 'text-multi' ? /x\u000by$/ : /xy$/, kind);
    // the file opens again as it was written
    const again = await loaded(await pkg.save());
    assert.equal((await again.getBody()).contentControls.find((x) => x.tag === `c56-${n}`).paragraphs.length, 1);
  }
});

test('CR-002 section 61: what the rule leaves alone — an unbound control, a w15:dataBinding one, a node of flat OPC, and a file already holding two paragraphs (D4)', async () => {
  // an unbound rich-text control takes a second paragraph, as before
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const unbound = body.insertParagraph('one', 'End').insertContentControl();
  unbound.insertParagraph('two', 'End');
  assert.deepEqual(unbound.paragraphs.map((p) => p.text), ['one', 'two']);
  assert.equal(unbound.type, 'RichText');
  // the file docx4j-python's 3a is, and 56a, read and saved as it is; a write removing a block is allowed, and the
  // write-back joins its paragraphs with \n (F1)
  const [file] = await check56OpenFiles();
  const two = await loaded(file.bytes);
  const control = (await two.getBody()).contentControls.find((x) => x.tag === 'c56-1');
  assert.deepEqual(control.paragraphs.map((p) => p.text), ['start1', 'second paragraph']);
  await two.customXmlParts.updateFromContentControls();
  assert.equal(control.xmlMapping.customXmlNode.text, 'start1\nsecond paragraph');
  control.paragraphs[1].delete();
  assert.deepEqual(control.paragraphs.map((p) => p.text), ['start1']);
  // a w15:dataBinding (Word's rich-text binding) is no w:dataBinding: the control stays rich text and takes blocks
  const { pkg: w15pkg, control: w15 } = await boundControl('richtext');
  const pr = w15.sdt.sdtPr;
  const binding = pr.rPrOrAliasOrLock.find((e) => e.name.localPart === 'dataBinding');
  binding.name = { namespaceURI: 'http://schemas.microsoft.com/office/word/2012/wordml', localPart: 'dataBinding', prefix: 'w15' };
  assert.equal(w15.type, 'RichText');
  w15.insertParagraph('two', 'End');
  assert.equal(w15.paragraphs.length, 2);
  void w15pkg;
});


test('CR-002 section 61.6: applyBindings over a bound block control holding more than one block writes one paragraph, as docx4j\'s bind.xslt does', async () => {
  // (56d's control, holding a table, is a container to the binding pass, as to docx4j's, and is left as it is)
  for (const file of (await check56OpenFiles()).filter((f) => ['56a-untyped-two-paragraphs.docx', '56c-richtext-two-paragraphs.docx'].includes(f.name))) {
    const pkg = await loaded(file.bytes);
    pkg.customXmlParts.getItem(STORE_ITEM_ID).selectSingleNode('/ns0:check56/ns0:value1', "xmlns:ns0='urn:check56'").text = 'bound value';
    await pkg.customXmlParts.applyBindings();
    const control = (await pkg.getBody()).contentControls.find((x) => x.tag === 'c56-1');
    const blocks = control.sdt.sdtContent.content.map((e) => e.name.localPart);
    assert.deepEqual(blocks, ['p'], file.name);
    assert.equal(control.text, 'bound value', file.name);
    await loaded(await pkg.save());
  }
});

test('CR-002 section 61.6: setMapping refuses, before anything changes, to bind a plain-text block control holding more than one paragraph', async () => {
  const pkg = await loaded(await check56Bytes());
  const body = await pkg.getBody();
  const text = body.contentControls.find((x) => x.tag === 'c56-5');   // w:text, multi-line
  const xpath = text.xmlMapping.xpath;
  const prefixes = text.xmlMapping.prefixMappings;
  text.xmlMapping.delete();
  text.insertParagraph('second', 'End');   // unbound now, so it may hold two
  assert.equal(text.paragraphs.length, 2);
  const before = await pkg.getMainDocumentPart().getXml();
  assert.throws(() => text.xmlMapping.setMapping(xpath, prefixes), (e) => e instanceof NotSupportedError && /section 61/.test(e.message));
  assert.equal(await pkg.getMainDocumentPart().getXml(), before);
  // one paragraph again, it binds
  text.paragraphs[1].delete();
  assert.equal(text.xmlMapping.setMapping(xpath, prefixes), true);
  assert.ok(text.xmlMapping.isMapped);
  // an untyped control takes Word's rich-text binding, w15:dataBinding, which may hold blocks
  const untyped = body.contentControls.find((x) => x.tag === 'c56-1');
  const xpath1 = untyped.xmlMapping.xpath;
  untyped.xmlMapping.delete();
  untyped.insertParagraph('second', 'End');
  assert.equal(untyped.paragraphs.length, 2);
  assert.equal(untyped.xmlMapping.setMapping(xpath1, prefixes), true);
  assert.ok(untyped.sdt.sdtPr.rPrOrAliasOrLock.some((e) => e.name.localPart === 'dataBinding' && /2012/.test(e.name.namespaceURI)));
  // section 61's rule doesn't cover Word's rich-text binding: its first paragraph takes the node's value, the rest stays
  assert.deepEqual(untyped.paragraphs.map((p) => p.text), ['start1', 'second']);
});
