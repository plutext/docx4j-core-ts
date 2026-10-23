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
import { Word } from '../dist/office-js/index.mjs';
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
  // through the Word shim, with no error of the snippet's own: what fails is what the shim doesn't have
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  assert.equal(ours.cases.length, 16);
  const errors = new Set(ours.cases.flatMap((c) => [c.error, ...Object.values(c.lookups), ...Object.values(c.region)])
    .filter((v) => typeof v === 'string' && v.startsWith('error:')));
  assert.deepEqual([...errors].sort(), [
    'error: NotImplemented: Word.Body.getContentControls is not supported by @docx4j/core-ts',
    'error: NotImplemented: Word.ContentControl.insertOoxml is not supported by @docx4j/core-ts',
    'error: NotImplemented: Word.ContentControl.insertTable is not supported by @docx4j/core-ts',
    'error: NotImplemented: Word.Document.customXmlParts is not supported by @docx4j/core-ts',
    'error: NotImplemented: Word.Paragraph.parentContentControlOrNullObject is not supported by @docx4j/core-ts',
  ]);
  // before section 61: an insertParagraph at a control's end gives it a second paragraph, whatever its kind
  for (const c of ours.cases.filter((x) => x.op === 'paragraph-end')) {
    assert.deepEqual(c.region.control.paragraphs, [`start${c.n}`, `NEWPARA${c.n}`], `case ${c.n}`);
  }
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
