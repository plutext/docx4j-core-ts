// CR-002 section 60, Office JS's characters written back. Check 55 (test/README.md) measures what Word 365's writers do
// with the characters Office JS's readers give: its file and its Script Lab snippet must agree, since Word runs the
// snippet on the file and nothing can be fixed after the run. Every case's BEGIN paragraph parses with the snippet's
// own parser to the case the file was built from, and every token the snippet searches for is in the file once.
// Word 365's run (2026-10-23) is beside the build as JSON, and its save in test/check55/ (outside fixtures/, so that
// the weekly harnesses cut no golden for it), with a SHA-1 manifest: CR-002 section 60.3 reads them.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../dist/index.mjs';
import { Word } from '../dist/office-js/index.mjs';
import { check55Bytes, check55Cases, STORE_ITEM_ID } from './fixtures/check55/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));

/** The snippet's CHECK55, loaded as Script Lab runs it, with no page around it. */
async function snippet() {
  const source = await readFile(join(here, 'fixtures', 'check55', 'check55-script-lab.js'), 'utf8');
  return new Function(`${source}\nreturn CHECK55;`)();
}

test("check 55's file and snippet agree: every case parses to the case it was built as, every word it searches for is in the file once", async () => {
  const { parseRegion, runCheck } = await snippet();
  const pkg = await WordprocessingMLPackage.load(await check55Bytes());
  await pkg.customXmlParts.load();
  const texts = (await pkg.getBody()).paragraphs.map((p) => p.text);
  const regions = texts.map(parseRegion).filter((r) => r !== null);
  const cases = check55Cases();
  assert.equal(cases.length, 58);
  assert.deepEqual(regions.map((r) => [r.n, r.op, r.tracked, r.chars]), cases.map((c) => [c.n, c.op, c.tracking === 'tracked', c.chars]));
  for (const r of regions) assert.equal(r.text, `x${String.fromCodePoint(...r.chars)}y`);
  const count = (token) => texts.reduce((sum, text) => sum + (text.match(new RegExp(`(?<![A-Za-z0-9])${token}(?![A-Za-z0-9])`, 'g')) ?? []).length, 0);
  for (const r of regions) {
    const tokens = [`BEGIN${r.n}`, `END${r.n}`, r.op.startsWith('bound-') ? `BOUND${r.n}` : `TARGET${r.n}`];
    for (const token of tokens) assert.equal(count(token), 1, `case ${r.n}: ${token}`);
  }
  // each bound case's control and node
  const data = pkg.customXmlParts.getItem(STORE_ITEM_ID);
  for (const r of regions.filter((x) => x.op.startsWith('bound-'))) {
    const control = (await pkg.getBody()).contentControls.find((c) => c.tag === `c55-${r.n}`);
    assert.equal(control.xmlMapping.customXmlNode.text, `start${r.n}`, `case ${r.n}`);
    assert.equal(control.isMultiLine, r.op === 'bound-multi', `case ${r.n}`);
    assert.equal(data.selectSingleNode(`/ns0:check55/ns0:value${r.n}`, "xmlns:ns0='urn:check55'").text, `start${r.n}`);
  }
  // the snippet runs every case through the Word shim with no error of its own: what fails is the custom XML
  // parts' read, which the shim doesn't have
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  assert.equal(ours.cases.length, 58);
  const errors = new Set(ours.cases.flatMap((c) => [c.error, ...Object.values(typeof c.region === 'string' ? { region: c.region } : c.region)])
    .filter((v) => typeof v === 'string' && v.startsWith('error:')));
  assert.deepEqual([...errors], ['error: NotImplemented: Word.Document.customXmlParts is not supported by @docx4j/core-ts']);
});

test("check 55's run: Word's JSONs are of this file's cases, and its save is Word's, as copied", async () => {
  const manifest = JSON.parse(await readFile(join(here, 'check55', 'manifest.json'), 'utf8'));
  assert.deepEqual(manifest.files.map((f) => f.name), ['55-characters-word365.docx']);
  for (const { name, sha1 } of manifest.files) {
    assert.equal(createHash('sha1').update(await readFile(join(here, 'check55', name))).digest('hex'), sha1, name);
  }
  const cases = check55Cases();
  for (const mode of ['run', 'reopened']) {
    const word = JSON.parse(await readFile(join(here, 'fixtures', 'check55', `55-characters-word365-${mode}.json`), 'utf8'));
    assert.equal(word.host.version, '16.0.20326.20232');
    assert.deepEqual(word.cases.map((c) => [c.n, c.op, c.tracked, c.spec]), cases.map((c) => [c.n, c.op, c.tracking === 'tracked', c.spec]), mode);
  }
});
