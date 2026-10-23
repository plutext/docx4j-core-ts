// CR-002 section 60, Office JS's characters written back. Check 55 (test/README.md) measures what Word 365's writers do
// with the characters Office JS's readers give: its file and its Script Lab snippet must agree, since Word runs the
// snippet on the file and nothing can be fixed after the run. Every case's BEGIN paragraph parses with the snippet's
// own parser to the case the file was built from, and every token the snippet searches for is in the file once.
// Word 365's run (2026-10-23) is beside the build as JSON, and its save in test/check55/ (outside fixtures/, so that
// the weekly harnesses cut no golden for it), with a SHA-1 manifest: CR-002 section 60.3 reads them.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../dist/index.mjs';
import { Word } from '../dist/office-js/index.mjs';
import { check55Bytes, check55Cases, check55bBytes, check55bCases, STORE_ITEM_ID } from './fixtures/check55/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));

/** The snippet's CHECK55 (or CHECK55B), loaded as Script Lab runs it, with no page around it. */
async function snippet(name = 'CHECK55') {
  const file = name === 'CHECK55B' ? 'check55b-script-lab.js' : 'check55-script-lab.js';
  const source = await readFile(join(here, 'fixtures', 'check55', file), 'utf8');
  return new Function(`${source}\nreturn ${name};`)();
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
  // parts' read, which the shim doesn't have, and a line feed into a run-level control, refused (CR-002 section 60 F4)
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  assert.equal(ours.cases.length, 58);
  const errors = new Set(ours.cases.flatMap((c) => [c.error, ...Object.values(typeof c.region === 'string' ? { region: c.region } : c.region)])
    .filter((v) => typeof v === 'string' && v.startsWith('error:')).map((v) => v.split(': a line feed')[0]));
  assert.deepEqual([...errors].sort(), [
    'error: NotImplemented: Word.ContentControl.insertText is not supported by @docx4j/core-ts',
    'error: NotImplemented: Word.Document.customXmlParts is not supported by @docx4j/core-ts',
  ]);
});

test("checks 55 and 55b's runs: Word's JSONs are of their files' cases, and the saves are Word's, as copied", async () => {
  const manifest = JSON.parse(await readFile(join(here, 'check55', 'manifest.json'), 'utf8'));
  assert.deepEqual(manifest.files.map((f) => f.name), ['55-characters-word365.docx', '55b-bound-word365.docx']);
  for (const { name, sha1 } of manifest.files) {
    assert.equal(createHash('sha1').update(await readFile(join(here, 'check55', name))).digest('hex'), sha1, name);
  }
  const cases = check55Cases();
  for (const mode of ['run', 'reopened']) {
    const word = JSON.parse(await readFile(join(here, 'fixtures', 'check55', `55-characters-word365-${mode}.json`), 'utf8'));
    assert.equal(word.host.version, '16.0.20326.20232');
    assert.deepEqual(word.cases.map((c) => [c.n, c.op, c.tracked, c.spec]), cases.map((c) => [c.n, c.op, c.tracking === 'tracked', c.spec]), mode);
  }
  const word55b = JSON.parse(await readFile(join(here, 'fixtures', 'check55', '55b-bound-word365-run.json'), 'utf8'));
  assert.equal(word55b.host.version, '16.0.20326.20232');
  assert.deepEqual(word55b.cases.map((c) => [c.n, c.op, c.spec]), check55bCases().map((c) => [c.n, c.op, c.spec]));
});

test("check 55b's file and snippet agree, and the snippet runs through the Word shim, every lookup recorded", async () => {
  const { parseRegion, runCheck } = await snippet('CHECK55B');
  const pkg = await WordprocessingMLPackage.load(await check55bBytes());
  await pkg.customXmlParts.load();
  const body = await pkg.getBody();
  const texts = body.paragraphs.map((p) => p.text);
  const regions = texts.map(parseRegion).filter((r) => r !== null);
  const cases = check55bCases();
  assert.equal(cases.length, 10);
  assert.deepEqual(regions.map((r) => [r.n, r.op, r.chars]), cases.map((c) => [c.n, c.op, c.chars]));
  const data = pkg.customXmlParts.getItem(STORE_ITEM_ID);
  for (const r of regions) {
    const control = body.contentControls.find((c) => c.tag === `c55b-${r.n}`);
    assert.equal(control.isMultiLine, r.op !== 'bound-single', `case ${r.n}`);
    assert.equal(data.selectSingleNode(`/ns0:check55b/ns0:value${r.n}`, "xmlns:ns0='urn:check55b'").text, r.op === 'bound-read' ? 'line1\nline2' : `start${r.n}`);
  }
  // through the shim: found by getContentControls with the plain-text type, as Word finds them (G1); what fails is
  // the custom XML read, which the shim doesn't have, and a line feed into a run-level control, refused (F4)
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  assert.equal(ours.cases.length, 10);
  for (const c of ours.cases.filter((x) => x.op !== 'bound-read' && !x.error)) assert.equal(c.result.via, 'range.getContentControls', `case ${c.n}`);
  const errors = new Set(ours.cases.flatMap((c) => [c.error, ...Object.values(c.lookups), ...Object.values(c.region)])
    .filter((v) => typeof v === 'string' && v.startsWith('error:')).map((v) => v.split(': a line feed')[0]));
  assert.deepEqual([...errors].sort(), [
    'error: NotImplemented: Word.ContentControl.insertText is not supported by @docx4j/core-ts',
    'error: NotImplemented: Word.Document.customXmlParts is not supported by @docx4j/core-ts',
  ]);
});

/** A case's paragraphs, comments and control as the snippet reports them, for comparing two runs. */
const seen = (c) => ({
  paragraphs: c.region.paragraphs,
  comments: c.region.comments,
  control: typeof c.region.control === 'object' && c.region.control ? c.region.control.text : undefined,
  error: c.error?.split(':')[1]?.trim(),
});

test("CR-002 section 60: check 55's snippet through the Word shim gives Word 365's run, case for case, but for the decided differences", async () => {
  const { runCheck } = await snippet();
  const pkg = await WordprocessingMLPackage.load(await check55Bytes());
  await pkg.customXmlParts.load();
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  const read = async (mode) => JSON.parse(await readFile(join(here, 'fixtures', 'check55', `55-characters-word365-${mode}.json`), 'utf8'));
  const [run, reopened] = [await read('run'), await read('reopened')];
  const compared = [];
  for (const theirs of run.cases) {
    const mine = ours.cases.find((c) => c.n === theirs.n);
    const chars = theirs.spec.split(',').map((u) => parseInt(u.slice(2), 16));
    if (theirs.op.startsWith('bound-')) continue;   // Word found none of these controls; check 55b measures them
    if (theirs.op === 'paragraph-after' && chars.some((c) => c === 0x0a || c === 0x0d)) {
      // E2, a recorded departure: Word writes the escape as text (x\ny), here one paragraph holding a w:br or w:cr
      assert.equal(seen(mine).paragraphs[2], chars.length === 2 || chars[0] === 0x0a ? 'x\\u000by' : 'x\\u000dy', `case ${theirs.n}`);
      continue;
    }
    if (chars.some((c) => c === 0x01 || c === 0x02)) {
      // E3: Word reads U+0001 and U+0002 until its save and drops them there; here they go at once, so its reopened report
      const after = reopened.cases.find((c) => c.n === theirs.n);
      assert.deepEqual([seen(mine).paragraphs, seen(mine).comments], [seen(after).paragraphs, seen(after).comments], `case ${theirs.n}`);
      continue;
    }
    assert.deepEqual([seen(mine).paragraphs, seen(mine).comments], [seen(theirs).paragraphs, seen(theirs).comments], `case ${theirs.n}`);
    compared.push(theirs.n);
  }
  assert.equal(compared.length, 42);
});

test("CR-002 section 60: check 55b's snippet through the Word shim gives Word 365's run, case for case, but for F4's refusal", async () => {
  const { runCheck } = await snippet('CHECK55B');
  const pkg = await WordprocessingMLPackage.load(await check55bBytes());
  await pkg.customXmlParts.load();
  await pkg.customXmlParts.applyBindings();   // Word fills a bound control from its node when it opens the file (case 10)
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check55', '55b-bound-word365-run.json'), 'utf8'));
  for (const theirs of word.cases) {
    const mine = ours.cases.find((c) => c.n === theirs.n);
    if (theirs.spec === 'U+000A' && theirs.op !== 'bound-read' || theirs.spec === 'U+000D') {
      // F4: Word moved the control to block level; here a line feed into a run-level control is refused, nothing changed
      assert.equal(seen(mine).error, 'NotImplemented', `case ${theirs.n}`);
      assert.deepEqual(seen(mine).paragraphs, [`BEGIN${theirs.n} ${theirs.op} untracked ${theirs.spec}`, `BOUND${theirs.n} start${theirs.n}`, `END${theirs.n}`]);
    } else if (theirs.spec === 'U+0002') {
      // E3: Word's text reads U+0002 until its save; here it goes at once
      assert.deepEqual(seen(mine).paragraphs, seen(theirs).paragraphs.map((p) => p.replace('\\u0002', '')), `case ${theirs.n}`);
    } else {
      assert.deepEqual([seen(mine).paragraphs, seen(mine).control], [seen(theirs).paragraphs, seen(theirs).control], `case ${theirs.n}`);
    }
  }
  // and the nodes, as section 60's D4, F1, F2 and F3 write them: Word's \r\n is \n here, its - the same
  const data = pkg.customXmlParts.getItem(STORE_ITEM_ID);
  const node = (n) => data.selectSingleNode(`/ns0:check55b/ns0:value${n}`, "xmlns:ns0='urn:check55b'").text;
  assert.deepEqual([1, 2, 5, 6, 7, 8].map(node), ['xAy', 'x\ny', 'x-y', 'xy', 'xAy', 'xy']);
});


/** A created package with one paragraph of text, its body and paragraph. */
async function one(text = 'ab') {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  return { pkg, body, p: body.insertParagraph(text, 'End') };
}

/** The saved main part, as a string. */
async function savedXml(pkg) {
  const again = await WordprocessingMLPackage.load(await pkg.save());
  return { again, xml: await again.getMainDocumentPart().getXml() };
}

test('CR-002 section 60 D1 and E3: every writer writes each character as the element it reads as, and drops what XML 1.0 excludes', async () => {
  const s = 'x\u0009\u000B\u000C\u000E\u001E\u001F\u0002\u0001\u0007￾y';
  const writers = {
    insertText: (w) => w.p.insertText(s, 'End'),
    tracked: (w) => { w.pkg.changeTrackingMode = 'TrackAll'; w.p.insertText(s, 'End'); },
    text: (w) => { w.p.text = s; },
    insertParagraph: (w) => w.p.insertParagraph(s, 'After'),
    bodyInsertParagraph: (w) => w.body.insertParagraph(s, 'End'),
    replace: (w) => w.p.getRange().insertText(s, 'Replace'),
    cellValue: (w) => { w.body.insertTable(1, 1, 'End').getCell(0, 0).value = s; },
  };
  for (const [name, write] of Object.entries(writers)) {
    const w = await one();
    write(w);
    const { again, xml } = await savedXml(w.pkg);
    assert.ok(!/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾]/.test(xml), name);
    for (const el of ['<w:tab/>', '<w:br/>', '<w:br w:type="page"/>', '<w:br w:type="column"/>', '<w:noBreakHyphen/>', '<w:softHyphen/>']) assert.ok(xml.includes(el), `${name}: ${el}`);
    // read back, each element is its character again, the excluded ones gone
    assert.ok((await again.getBody()).text.includes('x\u0009\u000B\u000C\u000E\u001E\u001Fy'), name);
  }
});

test('CR-002 section 60 E1: insertText makes a paragraph of each \\n and \\r, \\r\\n one, as Word; its range spans them; tracked, the marks inserted', async () => {
  const w = await one('ab');
  const r = w.p.insertText('x\ny\r\nz\rq', 'End');
  assert.deepEqual(w.body.paragraphs.map((p) => p.text), ['abx', 'y', 'z', 'q']);
  assert.equal(r.text, 'x\ry\rz\rq');
  assert.equal(r.endParagraph.text, 'q');
  // at a paragraph's end and start: an empty paragraph after it, before it
  const e = await one('ab');
  e.p.pPr().sectPr = { TYPE_NAME: 'org_docx4j_wml.SectPr' };   // the last paragraph of a section
  e.p.insertText('x\n', 'End');
  assert.deepEqual(e.body.paragraphs.map((p) => p.text), ['abx', '']);
  // as a paragraph mark is one: the new paragraph in front, the original keeping its mark and section break, empty
  assert.equal(e.p.text, '');
  assert.ok(e.p.p.pPr.sectPr && !e.body.paragraphs[0].p.pPr?.sectPr);
  const b = await one('ab');
  b.p.insertText('\nx', 'Start');
  assert.deepEqual(b.body.paragraphs.map((p) => p.text), ['', 'xab']);
  // Replace, and a range before which text goes keeps covering its text
  const rp = await one('one two three');
  const two = rp.p.search('two')[0];
  two.insertText('2\n2', 'Replace');
  assert.deepEqual(rp.body.paragraphs.map((p) => p.text), ['one 2', '2 three']);
  assert.equal(two.text, '2\r2');
  const before = await one('one two');
  const word = before.p.search('two')[0];
  word.insertText('A\nB ', 'Before');
  assert.equal(word.text, 'two');
  assert.deepEqual(before.body.paragraphs.map((p) => p.text), ['one A', 'B two']);
  // tracked: the inserted text in w:ins, and each new mark inserted, as Word's check 55 case 48
  const t = await one('ab');
  t.pkg.changeTrackingMode = 'TrackAll';
  t.p.insertText('x\ny', 'End');
  assert.deepEqual(t.body.paragraphs.map((p) => p.text), ['abx', 'y']);
  assert.match(t.body.paragraphs[0].p.pPr?.rPr?.ins ? 'inserted' : '', /inserted/);
  t.pkg.changeTrackingMode = 'Off';
  await t.pkg.rejectAllRevisions();
  assert.deepEqual(t.body.paragraphs.map((p) => p.text), ['ab']);
  // a Replace across paragraphs too
  const a = await one('one');
  a.body.insertParagraph('two', 'End');
  const across = a.body.paragraphs[0].search('one')[0].expandTo(a.body.paragraphs[1].search('two')[0]);
  across.insertText('A\nB', 'Replace');
  assert.deepEqual(a.body.paragraphs.map((p) => p.text), ['A', 'B']);
});

test('CR-002 section 60 E1 and F4: a paragraph break where none can be made is refused before anything changes', async () => {
  const w = await one('ab');
  const control = w.p.getRange().search('ab')[0].insertContentControl();
  const before = await w.pkg.getMainDocumentPart().getXml();
  assert.throws(() => control.insertText('x\ny', 'End'), (e) => e.code === 'NotImplemented' && /F4/.test(e.message));
  assert.throws(() => w.p.getRange().search('ab')[0].insertText('x\ny', 'Replace'), (e) => e.code === 'NotImplemented');
  assert.equal(await w.pkg.getMainDocumentPart().getXml(), before);
  // a line break (U+000B) goes in, as a w:br
  control.insertText('x\u000By', 'End');
  assert.match(await w.pkg.getMainDocumentPart().getXml(), /<w:t>x<\/w:t><w:br\/><w:t>y<\/w:t>/);
});

test('CR-002 section 60 E2 and D5: insertParagraph and Paragraph.text = keep one paragraph (\\n and \\r\\n a w:br, \\r a w:cr); a comment makes paragraphs', async () => {
  const w = await one('ab');
  const p = w.p.insertParagraph('x\ny\rz\r\nq', 'After');
  assert.equal(w.body.paragraphs.length, 2);
  assert.equal(p.text, 'x\u000By\rz\u000Bq');
  w.p.text = 'a\nb';
  assert.equal(w.p.text, 'a\u000Bb');
  const { xml } = await savedXml(w.pkg);
  assert.match(xml, /<w:t>y<\/w:t><w:cr\/><w:t>z<\/w:t>/);
  // comments: \r, \n and \r\n (one) each a paragraph, insertComment and content = alike (check 55 cases 29 to 46)
  const c = await w.p.getRange().insertComment('one\rtwo\r\nthree\nfour');
  assert.equal(c.content, 'one\rtwo\rthree\rfour');
  c.content = 'a\r\nb';
  assert.equal(c.content, 'a\rb');
});

test('CR-002 section 60 F3 and section 61 item 3: a plain-text control takes a line break as its kind does', async () => {
  const w = await one('ab');
  const single = w.p.getRange().search('ab')[0].insertContentControl('PlainText');
  single.insertText('x\u000By', 'End');
  assert.equal(single.text, 'abxy');
  const m = await one('cd');
  const multi = m.p.getRange().search('cd')[0].insertContentControl('PlainText');
  multi.sdt.sdtPr.rPrOrAliasOrLock.find((e) => e.name.localPart === 'text').value.multiLine = true;
  multi.insertText('x\u000By', 'End');
  assert.equal(multi.text, 'cdx\u000By');
});

test('CR-002 section 60 D3: a part holding a character XML 1.0 excludes is refused at save, naming it', async () => {
  const w = await one('ab');
  // a value put in the tree by hand, past the writers
  w.p.runs[0].value.content[0].value.value = 'a\u0002b';
  // the save's own error, docx4j's, names the part; its cause says why
  await assert.rejects(w.pkg.save(), (e) => /Part \/word\/document\.xml holds U\+0002, which XML 1\.0 excludes/.test(e.cause?.message));
  w.p.runs[0].value.content[0].value.value = 'ab';
  await w.pkg.save();
});

test('CR-002 section 60 G1: getContentControls lists the kinds asked for', async () => {
  const w = await one('ab cd');
  w.p.getRange().search('ab')[0].insertContentControl('PlainText');
  w.p.getRange().search('cd')[0].insertContentControl();
  assert.deepEqual(w.body.getContentControls().map((c) => c.type), ['PlainText', 'RichText']);
  assert.deepEqual(w.body.getContentControls({ types: ['PlainText'] }).map((c) => c.type), ['PlainText']);
  assert.deepEqual(w.p.getContentControls({ types: ['RichText'] }).map((c) => c.type), ['RichText']);
});

test('CR-002 section 60: text put into a paragraph whose every run is hidden can be seen, and is written as its characters are', async () => {
  const pkg = await WordprocessingMLPackage.load(readFileSync(join(here, 'text-extraction', 'suite', 'phase0', 'extraction-probe-inspect.docx')));
  const body = await pkg.getBody();
  const p = body.paragraphs[4];
  assert.equal(p.text, '');   // its one run is hidden through its style
  const r = p.insertText('X\u000BY\nZ', 'End');
  assert.equal(r.text, 'X\u000BY\rZ');
  assert.deepEqual(body.paragraphs.slice(4, 6).map((q) => q.text), ['X\u000BY', 'Z']);
  await pkg.save();   // before, the text went into a hidden run, its U+000B unmapped, and the save refused it
});
