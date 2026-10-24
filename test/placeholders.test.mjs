// CR-002 section 63, writing into a content control showing its placeholder. Check 63 (test/README.md) measures what
// Office JS's writes do to one, block-level and run-level, bound and not, with change tracking off and on: whether the
// control stays, whether w:showingPlcHdr and the Placeholder Text style go, and where the text lands. The snippet's
// file and the snippet must agree, since Word runs one on the other and nothing can be fixed after the run. Word 365's
// run (2026-10-24) is pinned as its answers, and the shim's as what this package gives until section 63 follows them.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../dist/index.mjs';
import { Word } from '../dist/office-js/index.mjs';
import { check63Bytes, check63Cases, beginText, OPS, MODES, PLACEHOLDER } from './fixtures/check63/build.mjs';
import * as check64 from './fixtures/check64/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const snippet63 = async () => new Function(`${await readFile(join(here, 'fixtures', 'check63', 'check63-script-lab.js'), 'utf8')}\nreturn CHECK63;`)();

/**
 * A region after the case as one line: `<n> <control><flag><style> <where> <paragraphs>`. The control is `C` where the
 * region's OOXML still holds w:sdt (else `x`); `P` where its w:sdtPr has w:showingPlcHdr; `S` where a run in its content
 * has the Placeholder Text style among its own properties (`s` where only deleted runs have it). Where NEW<n> landed: `in`, `out` (outside the control),
 * or `-`. Then the region's paragraphs' text, `|` between, and the write's error code if it threw.
 */
function digest(c) {
  const xml = typeof c.after?.ooxml === 'string' ? c.after.ooxml : '';
  const sdt = /<w:sdt>([\s\S]*?)<\/w:sdt>/.exec(xml);
  const pr = sdt ? (/<w:sdtPr>[\s\S]*?<\/w:sdtPr>/.exec(sdt[1]) ?? [''])[0] : '';
  const content = sdt ? (/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/.exec(sdt[1]) ?? ['', ''])[1] : '';
  // a run's own properties, not those a w:rPrChange records
  const runs = [...content.replace(/<w:rPrChange\b[\s\S]*?<\/w:rPrChange>/g, '').matchAll(/<w:r>(?:<w:rPr>([\s\S]*?)<\/w:rPr>)?<w:(t|delText)\b/g)];
  const styled = runs.filter((m) => (m[1] ?? '').includes('<w:rStyle w:val="PlaceholderText"/>')).map((m) => m[2]);
  const style = styled.includes('t') ? 'S' : styled.length > 0 ? 's' : '-';
  const token = `NEW${c.n}`;
  const where = content.includes(token) ? 'in' : xml.includes(token) ? 'out' : '-';
  const error = c.error ? ` ${/^error: (\w+)/.exec(c.error)?.[1] ?? c.error}` : '';
  const paragraphs = Array.isArray(c.after?.paragraphs) ? c.after.paragraphs.join('|') : String(c.after?.paragraphs);
  return `${c.n} ${sdt ? 'C' : 'x'}${pr.includes('showingPlcHdr') ? 'P' : '-'}${style} ${where} ${paragraphs}${error}`;
}

test("check 63's file and snippet agree: every case parses to the case it was built as, each control once, showing its placeholder", async () => {
  const { parseRegion } = await snippet63();
  const pkg = await WordprocessingMLPackage.load(await check63Bytes());
  const body = await pkg.getBody();
  const cases = check63Cases();
  const regions = body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null);
  assert.deepEqual(regions, cases.map(({ n, op, shape, mode }) => ({ n, op, shape, ...(mode ? { mode } : {}) })));
  for (const c of cases) assert.equal(parseRegion(beginText(c)).n, c.n);
  assert.equal(cases.length, 54);
  for (const c of cases) {
    if (c.op !== 'open') assert.ok(OPS.includes(c.op) && MODES.includes(c.mode), `case ${c.n}`);
    // the paragraph a run-level case's paragraph op looks for is there once
    if (c.shape.startsWith('run')) assert.equal(body.search(`PLAIN${c.n}A`, { matchCase: true, matchWholeWord: true }).length, 1, `case ${c.n}`);
  }
  const controls = body.contentControls;
  assert.deepEqual(controls.map((c) => c.tag), cases.map((c) => `c63-${c.n}`));
  assert.ok(controls.every((c) => c.isShowingPlaceholder), 'every control shows its placeholder');
  assert.deepEqual(controls.filter((c) => c.xmlMapping.isMapped).map((c) => c.tag), cases.filter((c) => c.shape.endsWith('-bound')).map((c) => `c63-${c.n}`));
  // the placeholder's text is Word's own, and reads as no text (CR-012 section 18)
  assert.equal(controls[0].text, '');
  assert.equal(controls[0].placeholderText, PLACEHOLDER);
});

test("CR-002 section 63: check 63's snippet through the Word shim, as this package gives it before section 63's changes", async () => {
  const { runCheck } = await snippet63();
  const run = async (mode) => {
    const pkg = await WordprocessingMLPackage.load(await check63Bytes());
    await pkg.getBody();
    return runCheck((f) => Word.run(pkg, f), mode);
  };
  const report = await run('report');
  assert.equal(report.caseCount, 54);
  // nothing written: every control there, showing Word's placeholder in its style, and reading as no text
  for (const c of report.cases) {
    assert.ok(/^\d+ CP[S-] - /.test(digest(c)), digest(c));
    assert.equal(c.after.control.text, '', `case ${c.n}`);
    assert.equal(c.after.showingPlaceholderText, true, `case ${c.n}`);
  }
  const result = await run('run');
  assert.equal(result.caseCount, 54);
  // the shim has no customXmlParts, so a bound control's node isn't read
  for (const c of result.cases.filter((x) => x.shape.endsWith('-bound'))) assert.match(c.after.node, /^error: NotImplemented/, `case ${c.n}`);
  assert.deepEqual(result.cases.map(digest), EXPECTED_RUN);
  // F3: a tracked insertText on a run-level control reads its text as '' where its paragraph reads NEW39
  assert.deepEqual([result.cases[38].after.control.text, result.cases[38].after.paragraphs], ['', ['PLAIN39A NEW39 PLAIN39B']]);
});

const wordRun = async (mode) => JSON.parse(await readFile(join(here, 'fixtures', 'check63', `63-placeholders-word365-${mode}.json`), 'utf8'));

test("check 63's run: Word's two JSONs are of this file's cases, and its saves are Word's, as copied", async () => {
  const cases = check63Cases();
  for (const mode of ['report', 'run']) {
    const word = await wordRun(mode);
    assert.deepEqual([word.check, word.mode, word.host.version, word.caseCount], [63, mode, '16.0.20326.20232', 54]);
    assert.deepEqual(word.cases.map(({ n, op, shape, mode: m }) => ({ n, op, shape, ...(m ? { mode: m } : {}) })), cases);
  }
  const manifest = JSON.parse(await readFile(join(here, 'check63', 'manifest.json'), 'utf8'));
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await readFile(join(here, 'check63', name))).digest('hex'), sha1, name);
});

test("CR-002 section 63: Word 365's answers (check 63): a text write into the control replaces the placeholder and clears the flag; the paragraph's Replace takes a run-level control", async () => {
  const report = await wordRun('report');
  const run = await wordRun('run');
  // opened, nothing written: every control there, showing its placeholder, reading as no text
  for (const c of report.cases) assert.ok(/^\d+ CP[S-] - (PLAIN\d+A  PLAIN\d+B)?$/.test(digest(c)), digest(c));
  assert.equal(run.cases.filter((c) => c.error).length, 0, 'Office JS refused nothing');
  assert.deepEqual(run.cases.map(digest), WORD_RUN);
  // Word gives no placeholderText without a glossary part, and no showingPlaceholderText (undefined in this build)
  for (const c of [...report.cases, ...run.cases]) {
    if (c.after.control) assert.equal(c.after.control.placeholderText, null, `case ${c.n}`);
    if (c.after.control) assert.equal(c.after.showingPlaceholderText, 'undefined', `case ${c.n}`);
  }
  // a bound control's node is left empty by every write, in Office JS's read and in Word's save
  for (const c of run.cases.filter((x) => x.shape.endsWith('-bound'))) assert.equal(c.after.node, '', `case ${c.n}`);
  const { unzipSync, strFromU8 } = await import('fflate');
  const saved = unzipSync(await readFile(join(here, 'check63', '63-run-word365.docx')));
  assert.match(strFromU8(saved['customXml/item1.xml']), /<check63 xmlns="urn:check63"><value21\/><value22\/><value23\/><value24\/><value25\/><value26\/><value47\/>/);
});

test("CR-002 section 63: the shim reads check 63's file as Word does, and after the writes differs from Word's run in the cases section 63.2 lists", async () => {
  const { runCheck } = await snippet63();
  const run = async (mode) => {
    const pkg = await WordprocessingMLPackage.load(await check63Bytes());
    await pkg.getBody();
    return runCheck((f) => Word.run(pkg, f), mode);
  };
  // what both read: the region, the control's type, text and paragraphs, and what the write returned. placeholderText
  // and showingPlaceholderText are left out: Word gives null and none, this package the placeholder's runs and the flag
  const view = (c) => ({ digest: digest(c), control: c.after.control && { type: c.after.control.type, text: c.after.control.text, paragraphs: c.after.control.paragraphs }, returned: c.result?.returned });
  const differing = {};
  for (const mode of ['report', 'run']) {
    const ours = await run(mode);
    const word = await wordRun(mode);
    differing[mode] = word.cases.filter((c, i) => JSON.stringify(view(c)) !== JSON.stringify(view(ours.cases[i]))).map((c) => c.n);
  }
  assert.deepEqual(differing.report, []);
  // 44 of the 52 writes. The eight that agree here: a block control's Content range deleted (6, 32), a bound control's
  // own insertText (21, 24), and four tracked writes over a whole control (38, 45, 46, 52), whose markup still differs:
  // Word marks the control deleted with w:customXmlDelRange, and puts a Replace's insertion first
  assert.deepEqual(differing.run, [1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 22, 23, 25, 26, 27, 28, 29, 30, 31,
    33, 34, 35, 36, 37, 39, 40, 41, 42, 43, 44, 47, 48, 49, 50, 51]);
});

// Word 365's run (2026-10-24), case by case, as `digest` reads it (CR-002 section 63.2).
const WORD_RUN = [
  '1 C-S in NEW1',
  '2 C-S in NEW2',
  '3 CPS in ',
  '4 C-S in NEW4',
  '5 C-S in NEW5',
  '6 CP- - ',
  '7 C-S in NEW7',
  '8 C-S in NEW8',
  '9 C-S in NEW9',
  '10 CPS out |NEW10',
  '11 CPS out NEW11|',
  '12 x-- - ',
  '13 C-S in PLAIN13A NEW13 PLAIN13B',
  '14 C-S in PLAIN14A NEW14 PLAIN14B',
  '15 C-S in PLAIN15A NEW15 PLAIN15B',
  '16 C-S in PLAIN16A NEW16 PLAIN16B',
  '17 CP- - PLAIN17A  PLAIN17B',
  '18 x-- out PLAIN18A NEW18 PLAIN18B',
  '19 x-- out NEW19',
  '20 x-- - ',
  '21 C-- in NEW21',
  '22 C-- in NEW22',
  '23 C-- in NEW23',
  '24 C-- in PLAIN24A NEW24 PLAIN24B',
  '25 C-- in PLAIN25A NEW25 PLAIN25B',
  '26 x-- out NEW26',
  '27 C-S in NEW27',
  '28 C-S in NEW28',
  '29 CPS in ',
  '30 C-S in NEW30',
  '31 C-S in NEW31',
  '32 CPs - ',
  '33 C-S in NEW33',
  '34 C-S in NEW34',
  '35 C-S in NEW35',
  '36 CPS out |NEW36',
  '37 CPS out NEW37|',
  '38 CPs - ',
  '39 C-S in PLAIN39A NEW39 PLAIN39B',
  '40 C-S in PLAIN40A NEW40 PLAIN40B',
  '41 C-S in PLAIN41A NEW41 PLAIN41B',
  '42 C-S in PLAIN42A NEW42 PLAIN42B',
  '43 CPs - PLAIN43A  PLAIN43B',
  '44 CPs out PLAIN44A NEW44 PLAIN44B',
  '45 CPs out NEW45',
  '46 CPs - ',
  '47 C-- in NEW47',
  '48 C-- in NEW48',
  '49 C-- in NEW49',
  '50 C-- in PLAIN50A NEW50 PLAIN50B',
  '51 C-- in PLAIN51A NEW51 PLAIN51B',
  '52 CPs out NEW52',
  '53 CP- - PLAIN53A  PLAIN53B',
  '54 CP- - ',
];

// What the shim gives (2026-10-24), case by case; CR-002 section 63 reads it. Only ContentControl.insertText clears the
// flag; every other write into the control leaves it on, so what it writes reads as no text. The untracked
// paragraph-text-replace on a run-level control (19, 26) empties the control and keeps it (docx4j-python's probe).
// F1: a tracked insertText on a bound block control (47) is refused only after it has deleted the placeholder and cleared
// the flag. F2: a run-level control's Content and Whole ranges are empty at its start, so text written to them lands
// outside it (15, 16, 18; 41, 42, 44), and its content-delete deletes nothing (17, 43). F3: a tracked write reads as text
// inside a control still showing its placeholder (33, 49), where the same write untracked reads as none (7, 23), and a
// tracked ContentControl.insertText on a run-level control (39) reads the control's text as ''.
const EXPECTED_RUN = [
  '1 C-- in NEW1',
  '2 C-- in NEW2',
  '3 CPS in |',
  '4 CP- in ',
  '5 CPS in ',
  '6 CP- - ',
  '7 CP- in ',
  '8 CPS in ',
  '9 CPS in ',
  '10 CPS in |',
  '11 CPS in |',
  '12 CP- - ',
  '13 C-- in PLAIN13A NEW13 PLAIN13B',
  '14 C-- in PLAIN14A NEW14 PLAIN14B',
  '15 CPS out PLAIN15A NEW15 PLAIN15B',
  '16 CPS out PLAIN16A NEW16 PLAIN16B',
  '17 CPS - PLAIN17A  PLAIN17B',
  '18 CPS out PLAIN18A NEW18 PLAIN18B',
  '19 CP- out NEW19',
  '20 CP- - ',
  '21 C-- in NEW21',
  '22 CP- in ',
  '23 CP- in ',
  '24 C-- in PLAIN24A NEW24 PLAIN24B',
  '25 CPS out PLAIN25A NEW25 PLAIN25B',
  '26 CP- out NEW26',
  '27 C-s in |NEW27',
  '28 C-s in |NEW28',
  '29 CPS in |',
  '30 CPs - ',
  '31 CPS in ',
  '32 CPs - ',
  '33 CPS in NEW33',
  '34 CPS in ',
  '35 CPS in ',
  '36 CPS in |',
  '37 CPS in |',
  '38 CPs - ',
  '39 C-s in PLAIN39A NEW39 PLAIN39B',
  '40 C-s in PLAIN40A NEW40 PLAIN40B',
  '41 CPS out PLAIN41A NEW41 PLAIN41B',
  '42 CPS out PLAIN42A NEW42 PLAIN42B',
  '43 CPS - PLAIN43A  PLAIN43B',
  '44 CPS out PLAIN44A NEW44 PLAIN44B',
  '45 CPs out NEW45',
  '46 CPs - ',
  '47 C-s -  NotImplemented',
  '48 CPs - ',
  '49 CPS in NEW49',
  '50 C-s in PLAIN50A NEW50 PLAIN50B',
  '51 CPS out PLAIN51A NEW51 PLAIN51B',
  '52 CPs out NEW52',
  '53 CP- - PLAIN53A  PLAIN53B',
  '54 CP- - ',
];

// --- check 64 (CR-002 section 63.4): check 63's writes on controls in Word's own form (w:placeholder and a glossary part
// Word made), controls Word makes itself, and a bound control's node with and without a placeholder

const snippet64 = async () => new Function(`${await readFile(join(here, 'fixtures', 'check64', 'check64-script-lab.js'), 'utf8')}\nreturn CHECK64;`)();

test("check 64's file and snippet agree: every case parses to the case it was built as; each control in Word's form, its glossary Word's", async () => {
  const { parseRegion } = await snippet64();
  const { unzipSync } = await import('fflate');
  const bytes = await check64.check64Bytes();
  const pkg = await WordprocessingMLPackage.load(bytes);
  await pkg.customXmlParts.load();
  const body = await pkg.getBody();
  const cases = check64.check64Cases();
  assert.equal(cases.length, 52);
  assert.deepEqual(body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null), cases);
  for (const c of cases) {
    assert.equal(parseRegion(check64.beginText(c)).n, c.n);
    if (/(^|-)run(-|$)/.test(c.shape)) assert.equal(body.search(`PLAIN${c.n}A`, { matchCase: true, matchWholeWord: true }).length, 1, `case ${c.n}`);
  }
  // a control for every case but the made ones, each naming Word's placeholder building block; the -value ones show their value
  const controls = body.contentControls;
  const withControl = cases.filter((c) => !c.shape.startsWith('made-'));
  assert.deepEqual(controls.map((c) => c.tag), withControl.map((c) => `c64-${c.n}`));
  assert.deepEqual(controls.map((c) => c.isShowingPlaceholder), withControl.map((c) => !c.shape.endsWith('-value')));
  assert.deepEqual(controls.filter((c) => c.xmlMapping.isMapped).map((c) => c.tag), withControl.filter((c) => /-(bound|value)$/.test(c.shape)).map((c) => `c64-${c.n}`));
  assert.deepEqual(controls.filter((c) => c.text !== '').map((c) => c.text), withControl.filter((c) => c.shape.endsWith('-value')).map((c) => `VALUE${c.n}`));
  const xml = new TextDecoder().decode(unzipSync(bytes)['word/document.xml']);
  assert.equal(xml.match(/<w:placeholder><w:docPart w:val="DefaultPlaceholder_-1854013440"\/><\/w:placeholder>/g).length, withControl.length);
  // the glossary parts are 34a-mapped.docx's (check 34, saved by Word), byte for byte, and hold that building block
  const { parts } = check64.glossaryParts();
  const built = unzipSync(bytes);
  assert.deepEqual(Object.keys(parts).sort(), ['word/glossary/_rels/document.xml.rels', 'word/glossary/document.xml', 'word/glossary/fontTable.xml',
    'word/glossary/settings.xml', 'word/glossary/styles.xml', 'word/glossary/webSettings.xml']);
  for (const [name, b] of Object.entries(parts)) assert.deepEqual(built[name], b, name);
  assert.match(new TextDecoder().decode(parts['word/glossary/document.xml']), /<w:docPartPr><w:name w:val="DefaultPlaceholder_-1854013440"\/>/);
});

test("CR-002 section 63.4: check 64's snippet through the Word shim, as this package gives it before section 63's changes", async () => {
  const { runCheck } = await snippet64();
  const run = async (mode) => {
    const pkg = await WordprocessingMLPackage.load(await check64.check64Bytes());
    await pkg.getBody();
    return runCheck((f) => Word.run(pkg, f), mode);
  };
  const report = await run('report');
  assert.equal(report.caseCount, 52);
  for (const c of report.cases) {
    const expected = c.shape.startsWith('made-') ? /^\d+ x-- - (PLAIN\d+A  PLAIN\d+B)?$/ : c.shape.endsWith('-value') ? /^\d+ C-- - (PLAIN\d+A )?VALUE\d+( PLAIN\d+B)?$/ : /^\d+ CPS - (PLAIN\d+A  PLAIN\d+B)?$/;
    assert.match(digest(c), expected);
  }
  const result = await run('run');
  // the shim makes a control around the empty paragraph, or after PLAIN<n>A, showing no placeholder
  for (const c of result.cases.filter((x) => x.shape.startsWith('made-'))) assert.equal(c.made.at, c.shape === 'made-block' ? 'paragraph' : 'After', `case ${c.n}`);
  assert.deepEqual(result.cases.map(digest), EXPECTED_RUN_64);
});

const wordRun64 = async (mode) => JSON.parse(await readFile(join(here, 'fixtures', 'check64', `64-placeholders-word-word365-${mode}.json`), 'utf8'));

test("check 64's run: Word's two JSONs are of this file's cases, and its saves are Word's, as copied", async () => {
  const cases = check64.check64Cases();
  for (const mode of ['report', 'run']) {
    const word = await wordRun64(mode);
    assert.deepEqual([word.check, word.mode, word.host.version, word.caseCount], [64, mode, '16.0.20326.20232', 52]);
    assert.deepEqual(word.cases.map(({ n, op, shape, mode: m }) => ({ n, op, shape, mode: m })), cases);
  }
  const manifest = JSON.parse(await readFile(join(here, 'check64', 'manifest.json'), 'utf8'));
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await readFile(join(here, 'check64', name))).digest('hex'), sha1, name);
});

test("CR-002 section 63.5: Word 365's answers (check 64): in Word's form a write takes the placeholder away, plain and untracked, and reaches a bound control's node", async () => {
  const { unzipSync, strFromU8 } = await import('fflate');
  const report = await wordRun64('report');
  const run = await wordRun64('run');
  // opened, nothing written: as built; placeholderText is the glossary's text now
  for (const c of report.cases) {
    const expected = c.shape.startsWith('made-') ? /^\d+ x-- - (PLAIN\d+A  PLAIN\d+B)?$/ : c.shape.endsWith('-value') ? /^\d+ C-- - (PLAIN\d+A )?VALUE\d+( PLAIN\d+B)?$/ : /^\d+ CPS - (PLAIN\d+A  PLAIN\d+B)?$/;
    assert.match(digest(c), expected);
    if (c.after.control) assert.equal(c.after.control.placeholderText, PLACEHOLDER, `case ${c.n}`);
  }
  assert.equal(run.cases.filter((c) => c.error).length, 0, 'Office JS refused nothing');
  // every write: the control kept, the flag gone, no run in the Placeholder Text style, the text in the control
  assert.deepEqual(run.cases.map(digest), WORD_RUN_64);
  assert.ok(run.cases.every((c) => /^\d+ C-- in /.test(digest(c))));
  // tracked, the placeholder isn't marked deleted: only the insertion is recorded
  for (const c of run.cases.filter((x) => x.mode === 'tracked')) {
    assert.match(c.after.ooxml, /<w:ins /, `case ${c.n}`);
    assert.doesNotMatch(c.after.ooxml, /<w:del |<w:delText/, `case ${c.n}`);
  }
  // Word makes a control showing its placeholder, in the Placeholder Text style, with a w:placeholder reference
  for (const c of run.cases.filter((x) => x.shape.startsWith('made-'))) {
    assert.match(c.made.ooxml, /<w:placeholder><w:docPart w:val="[^"]+"\/><\/w:placeholder><w:showingPlcHdr\/><\/w:sdtPr><w:sdtContent>(<w:p>)?<w:r><w:rPr><w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t>Click or tap here to enter text\.<\/w:t>/, `case ${c.n}`);
  }
  // every bound control's node holds the write, a placeholder's empty node too, in Office JS's read and in Word's save
  for (const c of run.cases.filter((x) => /-(bound|value)$/.test(x.shape))) assert.equal(c.after.node, `NEW${c.n}`, `case ${c.n}`);
  const saved = strFromU8(unzipSync(await readFile(join(here, 'check64', '64-run-word365.docx')))['customXml/item1.xml']);
  for (const n of [43, 44, 45, 46, 47, 48, 49, 50, 51, 52]) assert.ok(saved.includes(`<value${n}>NEW${n}</value${n}>`), `node ${n}`);
  // in Word's save every control, the 14 it made among them, names the default placeholder building block
  const document = strFromU8(unzipSync(await readFile(join(here, 'check64', '64-run-word365.docx')))['word/document.xml']);
  assert.equal(document.match(/<w:docPart w:val="DefaultPlaceholder_-1854013440"\/>/g).length, 52);
});

test("CR-002 section 63.5: the shim reads check 64's file as Word does, and after the writes differs from Word's run in the cases listed", async () => {
  const { runCheck } = await snippet64();
  const run = async (mode) => {
    const pkg = await WordprocessingMLPackage.load(await check64.check64Bytes());
    await pkg.getBody();
    return runCheck((f) => Word.run(pkg, f), mode);
  };
  const view = (c) => ({ digest: digest(c), control: c.after.control && { type: c.after.control.type, text: c.after.control.text, paragraphs: c.after.control.paragraphs }, returned: c.result?.returned });
  const differing = {};
  for (const mode of ['report', 'run']) {
    const ours = await run(mode);
    const word = await wordRun64(mode);
    differing[mode] = word.cases.filter((c, i) => JSON.stringify(view(c)) !== JSON.stringify(view(ours.cases[i]))).map((c) => c.n);
  }
  assert.deepEqual(differing.report, []);
  assert.deepEqual(differing.run, DIFFERING_64);
});

// Word 365's run of check 64 (2026-10-24), case by case, as `digest` reads it (CR-002 section 63.5).
const WORD_RUN_64 = [
  '1 C-- in NEW1',
  '2 C-- in NEW2',
  '3 C-- in NEW3',
  '4 C-- in NEW4',
  '5 C-- in NEW5',
  '6 C-- in |NEW6',
  '7 C-- in NEW7',
  '8 C-- in NEW8',
  '9 C-- in NEW9',
  '10 C-- in PLAIN10A NEW10 PLAIN10B',
  '11 C-- in PLAIN11A NEW11 PLAIN11B',
  '12 C-- in PLAIN12A NEW12 PLAIN12B',
  '13 C-- in PLAIN13A NEW13 PLAIN13B',
  '14 C-- in PLAIN14A NEW14 PLAIN14B',
  '15 C-- in NEW15',
  '16 C-- in NEW16',
  '17 C-- in NEW17',
  '18 C-- in NEW18',
  '19 C-- in PLAIN19ANEW19  PLAIN19B',
  '20 C-- in PLAIN20ANEW20  PLAIN20B',
  '21 C-- in PLAIN21ANEW21  PLAIN21B',
  '22 C-- in NEW22',
  '23 C-- in NEW23',
  '24 C-- in NEW24',
  '25 C-- in NEW25',
  '26 C-- in NEW26',
  '27 C-- in |NEW27',
  '28 C-- in NEW28',
  '29 C-- in NEW29',
  '30 C-- in NEW30',
  '31 C-- in PLAIN31A NEW31 PLAIN31B',
  '32 C-- in PLAIN32A NEW32 PLAIN32B',
  '33 C-- in PLAIN33A NEW33 PLAIN33B',
  '34 C-- in PLAIN34A NEW34 PLAIN34B',
  '35 C-- in PLAIN35A NEW35 PLAIN35B',
  '36 C-- in NEW36',
  '37 C-- in NEW37',
  '38 C-- in NEW38',
  '39 C-- in NEW39',
  '40 C-- in PLAIN40ANEW40  PLAIN40B',
  '41 C-- in PLAIN41ANEW41  PLAIN41B',
  '42 C-- in PLAIN42ANEW42  PLAIN42B',
  '43 C-- in NEW43',
  '44 C-- in NEW44',
  '45 C-- in NEW45',
  '46 C-- in NEW46',
  '47 C-- in NEW47',
  '48 C-- in NEW48',
  '49 C-- in PLAIN49A NEW49 PLAIN49B',
  '50 C-- in PLAIN50A NEW50 PLAIN50B',
  '51 C-- in PLAIN51A NEW51 PLAIN51B',
  '52 C-- in PLAIN52A NEW52 PLAIN52B',
];

// What the shim gives (2026-10-24), case by case: check 63's answers on controls in Word's form, so the w:placeholder and
// the glossary change nothing here.
const EXPECTED_RUN_64 = [
  '1 C-- in NEW1',
  '2 C-- in NEW2',
  '3 C-- in NEW3',
  '4 CP- in ',
  '5 CPS in ',
  '6 CPS in |',
  '7 CP- in ',
  '8 CPS in ',
  '9 CPS in ',
  '10 C-- in PLAIN10A NEW10 PLAIN10B',
  '11 C-- in PLAIN11A NEW11 PLAIN11B',
  '12 C-- in PLAIN12A NEW12 PLAIN12B',
  '13 CPS out PLAIN13A NEW13 PLAIN13B',
  '14 CPS out PLAIN14A NEW14 PLAIN14B',
  '15 C-- in NEW15',
  '16 C-- in NEW16',
  '17 C-- in NEW17',
  '18 C-- in NEW18',
  '19 C-- in PLAIN19ANEW19  PLAIN19B',
  '20 C-- in PLAIN20ANEW20  PLAIN20B',
  '21 C-- out PLAIN21ANEW21  PLAIN21B',
  '22 C-s in |NEW22',
  '23 C-s in |NEW23',
  '24 C-s in |NEW24',
  '25 CPs - ',
  '26 CPS in ',
  '27 CPS in |',
  '28 CPS in NEW28',
  '29 CPS in ',
  '30 CPS in ',
  '31 C-s in PLAIN31A NEW31 PLAIN31B',
  '32 C-s in PLAIN32A NEW32 PLAIN32B',
  '33 C-s in PLAIN33A NEW33 PLAIN33B',
  '34 CPS out PLAIN34A NEW34 PLAIN34B',
  '35 CPS out PLAIN35A NEW35 PLAIN35B',
  '36 C-- in |NEW36',
  '37 C-- in NEW37',
  '38 C-- in NEW38',
  '39 C-- in NEW39',
  '40 C-- in PLAIN40ANEW40  PLAIN40B',
  '41 C-- in PLAIN41ANEW41  PLAIN41B',
  '42 C-- out PLAIN42ANEW42  PLAIN42B',
  '43 C-- in NEW43',
  '44 CP- in ',
  '45 CP- in ',
  '46 C-- in NEW46',
  '47 C-- in NEW47',
  '48 C-- in NEW48',
  '49 C-- in PLAIN49A NEW49 PLAIN49B',
  '50 CPS out PLAIN50A NEW50 PLAIN50B',
  '51 C-- in PLAIN51A NEW51 PLAIN51B',
  '52 C-- in PLAIN52A NEW52 PLAIN52B',
];

// The writes of check 64 where this package differs from Word (28 of 52): only the control's own insertText clears the
// flag here, the Content range of a run-level control is outside it (F2), and a tracked write marks the placeholder
// deleted (or makes a second paragraph, F1), where Word records only the insertion. The made controls agree but for
// a run-level control's Content range (21, 42) and a tracked insertText's second paragraph (36), since this package
// makes them showing no placeholder.
const DIFFERING_64 = [4, 5, 6, 7, 8, 9, 13, 14, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 42, 44, 45, 50];
