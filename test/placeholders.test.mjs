// CR-002 section 63, writing into a content control showing its placeholder. Check 63 (test/README.md) measures what
// Office JS's writes do to one, block-level and run-level, bound and not, with change tracking off and on: whether the
// control stays, whether w:showingPlcHdr and the Placeholder Text style go, and where the text lands. The snippet's
// file and the snippet must agree, since Word runs one on the other and nothing can be fixed after the run. Word 365's
// run (2026-10-24) is pinned as its answers, and the shim's held to it since section 63 is implemented (section 63.12).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../dist/index.mjs';
import { Word } from '../dist/office-js/index.mjs';
import { textOf } from '@docx4j/generated-objects-ts/builders/wml';
import { check63Bytes, check63Cases, beginText, OPS, MODES, PLACEHOLDER } from './fixtures/check63/build.mjs';
import * as check64 from './fixtures/check64/build.mjs';
import * as check65 from './fixtures/check65/build.mjs';
import * as check66 from './fixtures/check66/build.mjs';
import * as check67 from './fixtures/check67/build.mjs';
import * as check68 from './fixtures/check68/build.mjs';
import * as check69 from './fixtures/check69/build.mjs';
import * as check70 from './fixtures/check70/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const snippet63 = async () => new Function(`${await readFile(join(here, 'fixtures', 'check63', 'check63-script-lab.js'), 'utf8')}\nreturn CHECK63;`)();

/**
 * A region after the case as one line: `<n> <control><flag><style> <where> <paragraphs>`. The control is `C` where the
 * region's OOXML still holds w:sdt (else `x`); `P` where its w:sdtPr has w:showingPlcHdr; `S` where a run in its content
 * has the Placeholder Text style among its own properties (`s` where only deleted runs have it). Where NEW<n> (FILL<n>) landed: `in`, `out` (outside the control),
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
  // what the case wrote: NEW<n>, or FILL<n> in check 67 (check 58's word)
  const tokens = [`NEW${c.n}`, `FILL${c.n}`];
  const where = tokens.some((t) => content.includes(t)) ? 'in' : tokens.some((t) => xml.includes(t)) ? 'out' : '-';
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
  // the placeholder's text is Word's own, and reads as no text (CR-012 section 18); placeholderText reads null, as Word's
  // does, since the control names no glossary entry (CR-002 section 63, D8)
  assert.equal(controls[0].text, '');
  assert.equal(controls[0].placeholderText, null);
  assert.equal(controls[0].placeholderEntry, undefined);
});

test("CR-002 section 63: check 63's snippet through the Word shim gives Word's answers with Design Mode off (run 2), case by case", async () => {
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
  // until section 63, 44 of the 52 writes differed from Word's run 1 (section 63.2), which was in Design Mode; since
  // section 63.17 they are run 2's. F1 to F3 are gone
  assert.deepEqual(result.cases.map(digest), WORD_RUN2);
  // F3: a tracked insertText on a run-level control read its text as '' where its paragraph read NEW39
  assert.deepEqual([result.cases[38].after.control.text, result.cases[38].after.paragraphs], ['NEW39', ['PLAIN39A NEW39 PLAIN39B']]);
  // F1: a tracked insertText on a bound block control was refused after it had changed it
  assert.equal(result.cases.filter((c) => c.error).length, 0);
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

test("CR-002 section 63: the shim against check 63's run 1, in Design Mode: it reads the file as Word did, and differs in the writes Design Mode made", async () => {
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
  // Design Mode's answers (section 63.14): the style kept, insertParagraph's flag left on, the placeholder marked deleted
  // under tracking, a Content delete leaving no filler. Section 63.12 followed them; 63.17 follows run 2
  assert.deepEqual(differing.run, [1, 2, 3, 4, 5, 7, 8, 9, 13, 14, 15, 16, 27, 28, 29, 30, 31, 32, 33, 34, 35, 38, 39, 40, 41, 42, 43, 44, 45, 46, 52]);
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

test("CR-002 section 63.4: check 64's snippet through the Word shim gives Word's answers, the controls it makes showing Word's placeholder", async () => {
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
  // the shim makes a control around the empty paragraph, or after PLAIN<n>A, showing Word's default placeholder in Word's
  // form, as Word's (CR-002 section 63, D8; until then it showed none)
  for (const c of result.cases.filter((x) => x.shape.startsWith('made-'))) {
    assert.equal(c.made.at, c.shape === 'made-block' ? 'paragraph' : 'After', `case ${c.n}`);
    assert.match(c.made.ooxml, /<w:placeholder><w:docPart w:val="DefaultPlaceholder_-1854013440"\/><\/w:placeholder><w:showingPlcHdr\/>(<w:tag [^>]*\/>)?<\/w:sdtPr><w:sdtContent>(<w:p>)?<w:r><w:rPr><w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t>Click or tap here to enter text\.<\/w:t>/, `case ${c.n}`);
  }
  // until section 63, 28 of the 52 writes differed (section 63.5)
  assert.deepEqual(result.cases.map(digest), WORD_RUN_64);
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

test("CR-002 section 63.5: the shim reads check 64's file as Word does, and after the writes too", async () => {
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
  assert.deepEqual(differing.run, []);
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

// --- check 65 (CR-002 section 63.7, for D8): what Word needs of a placeholder in its own form. 65a names Word's default
// entry with no glossary part; 65b has a glossary of one part, written here, holding that entry and one of its own text

const snippet65 = async () => new Function(`${await readFile(join(here, 'fixtures', 'check65', 'check65-script-lab.js'), 'utf8')}\nreturn CHECK65;`)();

test("check 65's files and snippet agree: every case parses to the case it was built as; 65a has no glossary part, 65b one of one part", async () => {
  const { parseRegion } = await snippet65();
  const { unzipSync, strFromU8 } = await import('fflate');
  assert.deepEqual(check65.CHECK65_FILES.map((f) => check65.check65Cases(f).length), [10, 14]);
  for (const file of check65.CHECK65_FILES) {
    const bytes = await check65.check65Bytes(file);
    const pkg = await WordprocessingMLPackage.load(bytes);
    const body = await pkg.getBody();
    const cases = check65.check65Cases(file);
    assert.deepEqual(body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null), cases, file.name);
    assert.match(body.paragraphs[0].text, new RegExp(`^Check 65${file.letter}: `));
    const withControl = cases.filter((c) => c.shape !== 'made-block');
    assert.deepEqual(body.contentControls.map((c) => c.tag), withControl.map((c) => `c65-${c.n}`), file.name);
    assert.ok(body.contentControls.every((c) => c.isShowingPlaceholder), file.name);
    const zip = unzipSync(bytes);
    const xml = strFromU8(zip['word/document.xml']);
    const named = [...xml.matchAll(/<w:docPart w:val="([^"]+)"\/>/g)].map((m) => m[1]);
    // every control names its entry but the bare-form one Jason types into, which names none (docx4j CR-063's question)
    assert.deepEqual(named, withControl.filter((c) => c.shape !== 'bare-bound').map((c) => (c.shape === 'custom' ? check65.CUSTOM_ENTRY : c.shape === 'missing' ? check65.MISSING_ENTRY : check65.DEFAULT_ENTRY)), file.name);
    assert.equal(cases.at(-1).shape, 'bare-bound');
    assert.ok(body.contentControls.at(-1).xmlMapping.isMapped, file.name);
    const glossary = Object.keys(zip).filter((k) => k.startsWith('word/glossary/'));
    if (file.glossary === 'none') assert.deepEqual(glossary, [], file.name);
    else {
      assert.deepEqual(glossary, ['word/glossary/document.xml']);
      assert.ok(pkg.getMainDocumentPart().glossaryDocumentPart, 'the glossary part is loaded');
      const names = [...strFromU8(zip['word/glossary/document.xml']).matchAll(/<w:docPartPr><w:name w:val="([^"]+)"\/>/g)].map((m) => m[1]);
      assert.deepEqual(names, [check65.DEFAULT_ENTRY, check65.CUSTOM_ENTRY]);
    }
  }
});

test("CR-002 section 63.7: check 65's snippet through the Word shim, on both files, gives Word's answers", async () => {
  const { runCheck } = await snippet65();
  const ours = {};
  for (const file of check65.CHECK65_FILES) {
    for (const mode of ['report', 'run']) {
      const pkg = await WordprocessingMLPackage.load(await check65.check65Bytes(file));
      await pkg.getBody();
      const result = await runCheck((f) => Word.run(pkg, f), mode);
      assert.equal(result.caseCount, check65.check65Cases(file).length);
      if (mode === 'report') for (const c of result.cases) assert.match(digest(c), c.shape === 'made-block' ? /^\d+ x-- - $/ : /^\d+ CPS - (PLAIN\d+A  PLAIN\d+B)?$/);
      else ours[file.letter] = result;
    }
  }
  // Word's, but 65b's made control (9), which shows its placeholder in the style here, where Word's glossary had lost it
  assert.deepEqual(Object.fromEntries(Object.entries(ours).map(([k, r]) => [k, r.cases.map(digest)])),
    { a: WORD_RUN_65.a, b: WORD_RUN_65.b.map((d) => (d.startsWith('9 ') ? '9 CPS - ' : d)) });
  // placeholderText reads the entry the control names (CR-002 section 63, D8), the custom control's among them; the setter
  // adds an entry of its own
  assert.equal(ours.b.cases[10].after.control.placeholderText, check65.CUSTOM_TEXT);
  assert.equal(ours.a.cases[5].result.placeholderText, 'Enter the name 6.');
});

const wordRun65 = async (file, mode) => JSON.parse(await readFile(join(here, 'fixtures', 'check65', `${file.name.replace('.docx', '')}-word365-${mode}.json`), 'utf8'));
const save65 = async (name) => {
  const { unzipSync, strFromU8 } = await import('fflate');
  const zip = unzipSync(await readFile(join(here, 'check65', name)));
  return Object.fromEntries(Object.entries(zip).map(([k, v]) => [k, strFromU8(v)]));
};
/** A case's control in a save of Word's: the w:sdt tagged c65-<n>, without Word's revision ids. */
const control65 = (xml, n) => [...xml.matchAll(/<w:sdt>[\s\S]*?<\/w:sdt>/g)].map((m) => m[0])
  .find((sdt) => sdt.includes(`<w:tag w:val="c65-${n}"/>`)).replace(/ (w:rsid\w*|w14:paraId|w14:textId)="[^"]*"/g, '');

test("check 65's run: Word's JSONs are of these files' cases, and its saves are Word's, as copied", async () => {
  for (const file of check65.CHECK65_FILES) {
    for (const mode of ['report', 'run']) {
      const word = await wordRun65(file, mode);
      assert.deepEqual([word.check, word.mode, word.host.version, word.caseCount], [65, mode, '16.0.20326.20232', check65.check65Cases(file).length]);
      assert.deepEqual(word.cases.map(({ n, op, shape, mode: m }) => ({ n, op, shape, mode: m })), check65.check65Cases(file));
    }
  }
  const manifest = JSON.parse(await readFile(join(here, 'check65', 'manifest.json'), 'utf8'));
  assert.equal(manifest.files.length, 5);
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await readFile(join(here, 'check65', name))).digest('hex'), sha1, name);
});

test("CR-002 section 63.8: Word 365's answers (check 65): the reference gives Word's form, only an entry in the glossary keeps it through a save, and a glossary without its styles loses the Placeholder Text style", async () => {
  const [a, b] = check65.CHECK65_FILES;
  // opened: as built; placeholderText reads the entry the control names, null where the glossary hasn't it
  for (const file of [a, b]) {
    for (const c of (await wordRun65(file, 'report')).cases) {
      assert.match(digest(c), c.shape === 'made-block' ? /^\d+ x-- - $/ : /^\d+ CPS - (PLAIN\d+A  PLAIN\d+B)?$/);
      if (!c.after.control) continue;
      const entry = file.glossary === 'none' || ['missing', 'bare-bound'].includes(c.shape) ? null : c.shape === 'custom' ? check65.CUSTOM_TEXT : PLACEHOLDER;
      assert.equal(c.after.control.placeholderText, entry, `${file.letter} case ${c.n}`);
    }
  }
  // the writes: Word's form in both files (check 64's), 65a's reference to no entry included
  const run = { a: await wordRun65(a, 'run'), b: await wordRun65(b, 'run') };
  assert.deepEqual({ a: run.a.cases.map(digest), b: run.b.cases.map(digest) }, WORD_RUN_65);
  for (const r of Object.values(run)) {
    assert.equal(r.cases.filter((c) => c.error).length, 0, 'Office JS refused nothing');
    assert.match(r.cases[2].after.ooxml, /<w:ins /);
    assert.doesNotMatch(r.cases[2].after.ooxml, /<w:del |<w:delText/);
    assert.deepEqual([r.cases[3].after.node, r.cases[4].after.node], ['NEW4', 'NEW5']);
    assert.equal(r.cases[5].result.placeholderText, 'Enter the name 6.');
  }
  // Word's save drops a reference to an entry no glossary holds: all of 65a's, and 65b's to the missing entry
  const openedA = await save65('65a-opened-word365.docx');
  assert.doesNotMatch(openedA['word/document.xml'], /<w:placeholder>/);
  assert.deepEqual(Object.keys(openedA).filter((k) => k.startsWith('word/glossary/')), []);
  const openedB = await save65('65b-opened-word365.docx');
  const named = (xml) => [...xml.matchAll(/<w:tag w:val="c65-(\d+)"\/>(?:(?!<\/w:sdtPr>)[\s\S])*?<w:docPart w:val="([^"]+)"\/>/g)].map((m) => `${m[1]} ${m[2] === check65.DEFAULT_ENTRY ? 'default' : m[2] === check65.CUSTOM_ENTRY ? 'custom' : m[2]}`);
  assert.deepEqual(named(openedB['word/document.xml']), ['1 default', '2 default', '3 default', '4 default', '5 default', '6 default', '7 default', '8 default', '10 custom', '11 custom']);
  // 65b's glossary kept, Word adding its other parts, but the entries' runs without the style its own styles lack
  assert.deepEqual(Object.keys(openedB).filter((k) => k.startsWith('word/glossary/')).sort(), ['word/glossary/_rels/document.xml.rels', 'word/glossary/document.xml', 'word/glossary/fontTable.xml', 'word/glossary/settings.xml', 'word/glossary/styles.xml', 'word/glossary/webSettings.xml']);
  assert.deepEqual([...openedB['word/glossary/document.xml'].matchAll(/<w:docPartBody>[\s\S]*?<w:r>(<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t>([^<]*)<\/w:t>/g)].map((m) => [m[1] ?? null, m[2]]), [[null, PLACEHOLDER], [null, check65.CUSTOM_TEXT]]);
  assert.doesNotMatch(openedB['word/glossary/styles.xml'], /w:styleId="PlaceholderText"/);
  // so the control Word makes from it shows its text unstyled (65b case 9), where 65a's, from the glossary Word made, has the style
  const runA = await save65('65a-run-word365.docx');
  const runB = await save65('65b-run-word365.docx');
  assert.match(control65(runA['word/document.xml'], 9), /<w:docPart w:val="DefaultPlaceholder_-1854013440"\/><\/w:placeholder><w:showingPlcHdr\/><\/w:sdtPr><w:sdtContent><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t>Click or tap here to enter text\.<\/w:t>/);
  assert.match(control65(runB['word/document.xml'], 9), /<w:docPart w:val="DefaultPlaceholder_-1854013440"\/><\/w:placeholder><w:showingPlcHdr\/><\/w:sdtPr><w:sdtContent><w:p><w:r><w:t>Click or tap here to enter text\.<\/w:t>/);
  assert.match(runA['word/glossary/styles.xml'], /<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"\/><w:basedOn w:val="DefaultParagraphFont"\/>/);
  assert.doesNotMatch(control65(runB['word/document.xml'], 13), /<w:placeholder>/);
  // text of a control's own (case 6) is a new entry: 32 hex digits, the default's docPartPr, its text in the style; the control names it
  for (const save of [runA, runB]) {
    const name = /<w:tag w:val="c65-6"\/><w:id w:val="65006"\/><w:placeholder><w:docPart w:val="([0-9A-F]{32})"\/><\/w:placeholder><w:showingPlcHdr\/>/.exec(control65(save['word/document.xml'], 6))[1];
    assert.match(save['word/glossary/document.xml'], new RegExp(`<w:docPartPr><w:name w:val="${name}"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:types><w:type w:val="bbPlcHdr"/></w:types><w:behaviors><w:behavior w:val="content"/></w:behaviors><w:guid w:val="\\{[0-9A-F-]{36}\\}"/></w:docPartPr><w:docPartBody><w:p [^>]*><w:r [^>]*><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Enter the name 6\\.</w:t></w:r></w:p></w:docPartBody>`));
    assert.match(control65(save['word/document.xml'], 6), /<w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t>Enter the name 6\.<\/w:t>/);
  }
  // in 65a Word made the glossary, the default entry beside case 6's
  assert.deepEqual([...runA['word/glossary/document.xml'].matchAll(/<w:docPartPr><w:name w:val="([^"]+)"/g)].map((m) => m[1] === check65.DEFAULT_ENTRY ? 'default' : 'own'), ['own', 'default']);
  // typed by hand into the bare-form bound control (65a case 10): the placeholder replaced, plain, and the node written
  assert.match(runA['customXml/item1.xml'], /<value10>TYPED10<\/value10>/);
  assert.match(control65(runA['word/document.xml'], 10), /<w:text\/><\/w:sdtPr><w:sdtEndPr\/><w:sdtContent><w:p><w:r><w:t>TYPED10<\/w:t><\/w:r><\/w:p><\/w:sdtContent>/);
  // in 65b the typing at step 8 went into case 10, so 14 is as built
  assert.match(control65(runB['word/document.xml'], 10), /<w:sdtContent><w:p><w:r><w:t>TYPED10<\/w:t><\/w:r><\/w:p><\/w:sdtContent>/);
  assert.match(runB['customXml/item1.xml'], /<value14\/>/);
  // typed into 14 once 65b's save was reopened: the node written, the placeholder replaced, and the control given an entry of
  // its own holding the placeholder it showed; in that session Word moved 7, 8, 9 and 13 to entries of their own too
  const typedB = await save65('65b-reopened-typed-word365.docx');
  assert.match(typedB['customXml/item1.xml'], /<value14>TYPED14<\/value14>/);
  const entry14 = /<w:id w:val="65014"\/><w:placeholder><w:docPart w:val="([0-9A-F]{32})"\/><\/w:placeholder><w:dataBinding [^>]*\/><w:text\/><\/w:sdtPr><w:sdtEndPr\/><w:sdtContent><w:p><w:r><w:t>TYPED14<\/w:t><\/w:r><\/w:p><\/w:sdtContent>/.exec(control65(typedB['word/document.xml'], 14))[1];
  assert.match(typedB['word/glossary/document.xml'], new RegExp(`<w:name w:val="${entry14}"/>[\\s\\S]*?<w:docPartBody><w:p [^>]*><w:pPr><w:pStyle w:val="${entry14}"/></w:pPr><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click or tap here to enter text\\.</w:t>`));
  const entryOf = (xml, n) => /<w:docPart w:val="([^"]+)"\/>/.exec(control65(xml, n))?.[1] ?? null;
  assert.deepEqual([1, 6, 7, 8, 9, 11, 13].map((n) => entryOf(runB['word/document.xml'], n) === entryOf(typedB['word/document.xml'], n)), [true, true, false, false, false, true, false]);
});

test("CR-002 section 63.8: the shim reads check 65's files as Word does, placeholderText too, and after the writes differs from Word's run only in 65b's made control", async () => {
  const { runCheck } = await snippet65();
  const view = (c) => ({ digest: digest(c), control: c.after.control && { type: c.after.control.type, text: c.after.control.text, paragraphs: c.after.control.paragraphs }, returned: c.result?.returned });
  const differing = {};
  const placeholderTexts = {};
  for (const file of check65.CHECK65_FILES) {
    for (const mode of ['report', 'run']) {
      const pkg = await WordprocessingMLPackage.load(await check65.check65Bytes(file));
      await pkg.getBody();
      const ours = await runCheck((f) => Word.run(pkg, f), mode);
      const word = await wordRun65(file, mode);
      differing[`${file.letter} ${mode}`] = word.cases.filter((c, i) => JSON.stringify(view(c)) !== JSON.stringify(view(ours.cases[i]))).map((c) => c.n);
      if (mode === 'report') placeholderTexts[file.letter] = word.cases.filter((c, i) => c.after.control && c.after.control.placeholderText !== ours.cases[i].after.control.placeholderText).map((c) => c.n);
    }
  }
  // until section 63, 2, 3, 5 and 9 of each. 65b's 9 is Word's made control showing the default entry unstyled, the style
  // its glossary lacked (section 63.8); this package writes the entry in the style
  assert.deepEqual(differing, { 'a report': [], 'a run': [], 'b report': [], 'b run': [9] });
  // the entry the control names, null where the glossary hasn't it, as Word reads it (until section 63, the runs)
  assert.deepEqual(placeholderTexts, { a: [], b: [] });
});

// Word 365's runs of check 65 (2026-10-25), as `digest` reads them (CR-002 section 63.8): check 64's answers in both files.
// 65b's made control (9) shows its placeholder without the style, which Word dropped from the glossary it read.
const WORD_RUN_65 = {
  a: [
    '1 C-- in NEW1',
    '2 C-- in |NEW2',
    '3 C-- in NEW3',
    '4 C-- in NEW4',
    '5 C-- in PLAIN5A NEW5 PLAIN5B',
    '6 CPS - ',
    '7 CPS - ',
    '8 CPS - ',
    '9 CPS - ',
    '10 CPS - ',
  ],
  b: [
    '1 C-- in NEW1',
    '2 C-- in |NEW2',
    '3 C-- in NEW3',
    '4 C-- in NEW4',
    '5 C-- in PLAIN5A NEW5 PLAIN5B',
    '6 CPS - ',
    '7 CPS - ',
    '8 CPS - ',
    '9 CP- - ',
    '10 C-- in NEW10',
    '11 CPS - ',
    '12 C-- in NEW12',
    '13 CPS - ',
    '14 CPS - ',
  ],
};

// --- check 66 (CR-002 section 63.10, for D9): how much glossary to write. 66a has the glossary document part and a styles
// part holding Placeholder Text (D9 (a)); 66b check 34's glossary, Word's six parts (D9 (b)); the same six cases in both

const snippet66 = async () => new Function(`${await readFile(join(here, 'fixtures', 'check66', 'check66-script-lab.js'), 'utf8')}\nreturn CHECK66;`)();

test("check 66's files and snippet agree: every case parses to the case it was built as; 66a's glossary is the document and styles parts, 66b's Word's", async () => {
  const { parseRegion } = await snippet66();
  const { unzipSync, strFromU8 } = await import('fflate');
  const cases = check66.check66Cases();
  assert.deepEqual(cases.map((c) => `${c.op} ${c.shape}`), ['open block', 'type block-bound', 'content-delete block-value', 'make made-block', 'set-placeholder-text block', 'type bare-bound']);
  const word = unzipSync(await readFile(check66.GLOSSARY_SOURCE));
  for (const file of check66.CHECK66_FILES) {
    const bytes = await check66.check66Bytes(file);
    const pkg = await WordprocessingMLPackage.load(bytes);
    const body = await pkg.getBody();
    assert.deepEqual(body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null), cases, file.name);
    assert.match(body.paragraphs[0].text, new RegExp(`^Check 66${file.letter}: `));
    const withControl = cases.filter((c) => c.shape !== 'made-block');
    assert.deepEqual(body.contentControls.map((c) => c.tag), withControl.map((c) => `c66-${c.n}`), file.name);
    assert.deepEqual(body.contentControls.map((c) => c.isShowingPlaceholder), withControl.map((c) => c.shape !== 'block-value'), file.name);
    const zip = unzipSync(bytes);
    // every control names the default entry but the bare-form one
    assert.deepEqual([...strFromU8(zip['word/document.xml']).matchAll(/<w:tag w:val="c66-(\d+)"\/>(?:(?!<\/w:sdtPr>)[\s\S])*?<w:docPart w:val="([^"]+)"\/>/g)].map((m) => `${m[1]} ${m[2]}`),
      ['1', '2', '3', '5'].map((n) => `${n} ${check66.DEFAULT_ENTRY}`));
    const glossary = Object.keys(zip).filter((k) => k.startsWith('word/glossary/')).sort();
    if (file.glossary === 'word') {
      assert.deepEqual(glossary, Object.keys(word).filter((k) => k.startsWith('word/glossary/')).sort());
      for (const k of glossary) assert.deepEqual(zip[k], word[k], k);
    } else {
      assert.deepEqual(glossary, ['word/glossary/_rels/document.xml.rels', 'word/glossary/document.xml', 'word/glossary/styles.xml']);
      assert.match(strFromU8(zip['word/glossary/document.xml']), /<w:docPartBody><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t xml:space="preserve">Click or tap here to enter text\.<\/w:t><\/w:r><\/w:p><\/w:docPartBody>/);
      // the styles are Word's, as check 34's glossary defines them
      const wordStyles = strFromU8(word['word/glossary/styles.xml']).replace(/<w:rsid [^>]*\/>/g, '');
      for (const style of check66.PLACEHOLDER_STYLES.match(/<w:style [\s\S]*?<\/w:style>/g)) assert.ok(wordStyles.includes(style), style);
    }
    const main = pkg.getMainDocumentPart();
    assert.ok(main.glossaryDocumentPart, 'the glossary part is loaded');
  }
});

test("CR-002 section 63.10: check 66's snippet through the Word shim, on both files, gives Word's answers", async () => {
  const { runCheck } = await snippet66();
  const ours = {};
  for (const file of check66.CHECK66_FILES) {
    for (const mode of ['report', 'run']) {
      const pkg = await WordprocessingMLPackage.load(await check66.check66Bytes(file));
      await pkg.getBody();
      const result = await runCheck((f) => Word.run(pkg, f), mode);
      assert.equal(result.caseCount, 6);
      ours[`${file.letter} ${mode}`] = result.cases.map(digest);
    }
  }
  // an emptied control (3) and a made one (4) show the default entry, in the style (CR-002 section 63, D8; until then,
  // no placeholder)
  const report = ['1 CPS - ', '2 CPS - ', '3 C-- - VALUE3', '4 x-- - ', '5 CPS - ', '6 CPS - '];
  const run = ['1 CPS - ', '2 CPS - ', '3 CPS - ', '4 CPS - ', '5 CPS - ', '6 CPS - '];
  assert.deepEqual(ours, { 'a report': report, 'a run': run, 'b report': report, 'b run': run });
});

const wordRun66 = async (file, mode) => JSON.parse(await readFile(join(here, 'fixtures', 'check66', `${file.name.replace('.docx', '')}-word365-${mode}.json`), 'utf8'));
const save66 = async (name) => {
  const { unzipSync, strFromU8 } = await import('fflate');
  return Object.fromEntries(Object.entries(unzipSync(await readFile(join(here, 'check66', name)))).map(([k, v]) => [k, strFromU8(v)]));
};
/** Every control in a save of Word's, in order: `<n> <flag><style> <entry> <text>`, the entry `default`, `own` (32 hex digits) or `-`. */
const controls66 = (xml) => [...xml.matchAll(/<w:sdt>[\s\S]*?<\/w:sdt>/g)].map(([sdt]) => {
  const entry = /<w:docPart w:val="([^"]+)"\/>/.exec(sdt)?.[1];
  const text = [.../<w:sdtContent>[\s\S]*<\/w:sdtContent>/.exec(sdt)[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');
  return `${/c66-(\d+)/.exec(sdt)[1]} ${sdt.includes('<w:showingPlcHdr/>') ? 'P' : '-'}${sdt.includes('<w:rStyle w:val="PlaceholderText"/>') ? 'S' : '-'} `
    + `${entry === undefined ? '-' : entry === check66.DEFAULT_ENTRY ? 'default' : /^[0-9A-F]{32}$/.test(entry) ? 'own' : entry} ${text}`;
});

test("check 66's run: Word's JSONs are of these files' cases, and its saves are Word's, as copied", async () => {
  for (const file of check66.CHECK66_FILES) {
    for (const mode of ['report', 'run']) {
      const word = await wordRun66(file, mode);
      assert.deepEqual([word.check, word.mode, word.host.version, word.caseCount], [66, mode, '16.0.20326.20232', 6]);
      assert.deepEqual(word.cases.map(({ n, op, shape, mode: m }) => ({ n, op, shape, mode: m })), check66.check66Cases());
    }
  }
  const manifest = JSON.parse(await readFile(join(here, 'check66', 'manifest.json'), 'utf8'));
  assert.equal(manifest.files.length, 8);
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await readFile(join(here, 'check66', name))).digest('hex'), sha1, name);
});

test("CR-002 section 63.11: Word 365's answers (check 66): a glossary of the document and styles parts does what Word's six parts do, through the run, a reopening and typing", async () => {
  const placeholder = /<w:docPartBody><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t>Click or tap here to enter text\.<\/w:t><\/w:r><\/w:p><\/w:docPartBody>/;
  const strip = (xml) => xml.replace(/ w:rsid\w*="[^"]*"|<w:rsid [^>]*\/>|<w:rsids>[\s\S]*?<\/w:rsids>/g, '');
  const runs = {};
  for (const file of check66.CHECK66_FILES) {
    // opened: placeholderText the entry's text, null for the bare-form control
    const report = await wordRun66(file, 'report');
    assert.deepEqual(report.cases.map(digest), ['1 CPS - ', '2 CPS - ', '3 C-- - VALUE3', '4 x-- - ', '5 CPS - ', '6 CPS - ']);
    assert.deepEqual(report.cases.map((c) => c.after.control?.placeholderText), [PLACEHOLDER, PLACEHOLDER, PLACEHOLDER, undefined, PLACEHOLDER, null]);
    // the run: the emptied control (3) and the made one (4) show the entry, in the style; 5 reads its own text
    const run = await wordRun66(file, 'run');
    runs[file.letter] = run.cases.map(digest);
    assert.deepEqual(runs[file.letter], ['1 CPS - ', '2 CPS - ', '3 CPS - ', '4 CPS - ', '5 CPS - ', '6 CPS - ']);
    assert.equal(run.cases.filter((c) => c.error).length, 0);
    assert.equal(run.cases[4].result.placeholderText, 'Enter the name 5.');
    // Word's saves: every reference kept through the run, a reopening and typing; nothing moved to an entry of its own
    const letter = file.letter;
    const saved = Object.fromEntries(await Promise.all(['opened', 'run', 'reopened', 'typed'].map(async (s) => [s, await save66(`66${letter}-${s}-word365.docx`)])));
    const shown = '1 PS default Click or tap here to enter text.';
    assert.deepEqual(controls66(saved.opened['word/document.xml']), [shown, '2 PS default Click or tap here to enter text.', '3 -- default VALUE3', '5 PS default Click or tap here to enter text.', '6 PS - Click or tap here to enter text.']);
    const afterRun = [shown, '2 PS default Click or tap here to enter text.', '3 PS default Click or tap here to enter text.', '4 PS default Click or tap here to enter text.', '5 PS own Enter the name 5.', '6 PS - Click or tap here to enter text.'];
    assert.deepEqual(controls66(saved.run['word/document.xml']), afterRun);
    assert.deepEqual(controls66(saved.reopened['word/document.xml']), afterRun);
    assert.deepEqual(controls66(saved.typed['word/document.xml']), [shown, '2 -- default TYPED2', ...afterRun.slice(2, 5), '6 -- - TYPED6']);
    assert.match(saved.typed['customXml/item1.xml'], /<value2>TYPED2<\/value2><value6>TYPED6<\/value6>/);
    // the glossary: the default entry in the style throughout, case 5's beside it, and the same full set of parts
    for (const s of Object.values(saved)) {
      assert.match(strip(s['word/glossary/document.xml']), placeholder);
      assert.deepEqual([...s['word/glossary/styles.xml'].matchAll(/w:styleId="([^"]+)"/g)].map((m) => m[1]), ['Normal', 'DefaultParagraphFont', 'TableNormal', 'NoList', 'PlaceholderText']);
      assert.deepEqual(Object.keys(s).filter((k) => k.startsWith('word/glossary/')).sort(), ['word/glossary/_rels/document.xml.rels', 'word/glossary/document.xml', 'word/glossary/fontTable.xml', 'word/glossary/settings.xml', 'word/glossary/styles.xml', 'word/glossary/webSettings.xml']);
    }
    assert.equal([...saved.typed['word/glossary/document.xml'].matchAll(/<w:docPart>/g)].length, 2);
  }
  // 66a and 66b alike; Word's glossary of 66a differs from 66b's only in the settings and web settings Word wrote for it
  assert.deepEqual(runs.a, runs.b);
  const [a, b] = await Promise.all(['66a-opened-word365.docx', '66b-opened-word365.docx'].map(save66));
  for (const part of ['word/glossary/document.xml', 'word/glossary/styles.xml', 'word/glossary/fontTable.xml', 'word/glossary/_rels/document.xml.rels']) assert.equal(strip(a[part]), strip(b[part]), part);
  assert.match(a['word/glossary/settings.xml'], /w:name="compatibilityMode" w:uri="http:\/\/schemas\.microsoft\.com\/office\/word" w:val="12"/);
  assert.match(b['word/glossary/settings.xml'], /w:name="compatibilityMode" w:uri="http:\/\/schemas\.microsoft\.com\/office\/word" w:val="15"/);
});

test("CR-002 section 63.11: the shim reads check 66's files as Word does, and after the writes too", async () => {
  const { runCheck } = await snippet66();
  const view = (c) => ({ digest: digest(c), control: c.after.control && { type: c.after.control.type, text: c.after.control.text, paragraphs: c.after.control.paragraphs }, returned: c.result?.returned });
  const differing = {};
  for (const file of check66.CHECK66_FILES) {
    for (const mode of ['report', 'run']) {
      const pkg = await WordprocessingMLPackage.load(await check66.check66Bytes(file));
      await pkg.getBody();
      const ours = await runCheck((f) => Word.run(pkg, f), mode);
      const word = await wordRun66(file, mode);
      differing[`${file.letter} ${mode}`] = word.cases.filter((c, i) => JSON.stringify(view(c)) !== JSON.stringify(view(ours.cases[i]))).map((c) => c.n);
    }
  }
  // until section 63, 3 and 4 of each: an emptied control and a made one showed no placeholder, where Word shows the entry
  assert.deepEqual(differing, { 'a report': [], 'a run': [], 'b report': [], 'b run': [] });
});

// --- CR-002 section 63.12: section 63 implemented. The checks' runs above are held to Word's; these are what the shim's
// reads can't show: the markup a tracked write leaves, the node a bound control's write reaches, and the glossary a save
// writes

/** A check's file run through the shim in `run` mode, its package kept: `{ pkg, result }`. */
async function shimRun(bytes, snippet, { customXml = false } = {}) {
  const pkg = await WordprocessingMLPackage.load(bytes);
  await pkg.getBody();
  if (customXml) await pkg.customXmlParts.load();
  const result = await snippet.runCheck((f) => Word.run(pkg, f), 'run');
  return { pkg, result };
}
const zipText = async (bytes) => {
  const { unzipSync, strFromU8 } = await import('fflate');
  return Object.fromEntries(Object.entries(unzipSync(bytes)).map(([k, v]) => [k, strFromU8(v)]));
};
/** A case's region in a document part's XML: what lies between its BEGIN and END paragraphs, revision attributes left out. */
const regionXml = (xml, n) => {
  const begin = xml.indexOf(`>BEGIN${n} `);
  const from = xml.indexOf('</w:p>', begin) + '</w:p>'.length;
  const to = xml.lastIndexOf('<w:p>', xml.indexOf(`>END${n}<`));
  return xml.slice(from, to).replace(/ (w:author|w:date|w16du:dateUtc|w:id)="[^"]*"/g, '').replace(/ xml:space="preserve"/g, '');
};

test("CR-002 section 63.17, D3: a Replace or delete over a whole control showing its placeholder takes it; tracked, Word's markers, the insertion first", async () => {
  const { pkg, result } = await shimRun(await check63Bytes(), await snippet63());
  assert.equal(result.cases.filter((c) => c.error).length, 0);
  const xml = (await zipText(await pkg.save()))['word/document.xml'];
  // tracked, the control between Word's markers, its placeholder gone unmarked and the flag cleared (check 63's run 2)
  const marked = (n, bound = '') => `<w:customXmlDelRangeStart/><w:sdt><w:sdtPr><w:alias w:val="Check 63 control ${n}"/><w:tag w:val="c63-${n}"/><w:id w:val="${63000 + n}"/>${bound}</w:sdtPr>`
    + '<w:sdtContent><w:customXmlDelRangeEnd/><w:customXmlDelRangeStart/></w:sdtContent></w:sdt><w:customXmlDelRangeEnd/>';
  // untracked the control goes: the paragraph's Replace (19), its Content delete (20), the control's Whole Replace (18)
  assert.equal(regionXml(xml, 19), '<w:p><w:r><w:t>NEW19</w:t></w:r></w:p>');
  assert.equal(regionXml(xml, 20), '<w:p/>');
  assert.equal(regionXml(xml, 18), '<w:p><w:r><w:t>PLAIN18A NEW18</w:t></w:r><w:r><w:t> PLAIN18B</w:t></w:r></w:p>');
  assert.equal(regionXml(xml, 12), '<w:p/>', 'a block control goes, its paragraph emptied in its place');
  // tracked (44 to 46): Word's markers around the control and inside it, its placeholder deleted, a Replace's insertion first
  assert.equal(regionXml(xml, 44), `<w:p><w:r><w:t>PLAIN44A </w:t></w:r><w:ins><w:r><w:t>NEW44</w:t></w:r></w:ins>${marked(44)}<w:r><w:t> PLAIN44B</w:t></w:r></w:p>`);
  assert.equal(regionXml(xml, 45), `<w:p><w:ins><w:r><w:t>NEW45</w:t></w:r></w:ins><w:del><w:r><w:delText>PLAIN45A </w:delText></w:r></w:del>${marked(45)}<w:del><w:r><w:delText> PLAIN45B</w:delText></w:r></w:del></w:p>`);
  assert.equal(regionXml(xml, 46), `<w:p><w:del><w:r><w:delText>PLAIN46A </w:delText></w:r></w:del>${marked(46)}<w:del><w:r><w:delText> PLAIN46B</w:delText></w:r></w:del></w:p>`);
  assert.equal(regionXml(xml, 38), '<w:customXmlDelRangeStart/><w:sdt><w:sdtPr><w:alias w:val="Check 63 control 38"/><w:tag w:val="c63-38"/><w:id w:val="63038"/></w:sdtPr>'
    + '<w:sdtContent><w:customXmlDelRangeEnd/><w:p/><w:customXmlDelRangeStart/></w:sdtContent></w:sdt><w:customXmlDelRangeEnd/>');
  // a tracked write into the control: only the insertion, plain, the placeholder gone unmarked (D2 (b), D5), in a bound
  // plain-text control too (28, 47)
  assert.equal(regionXml(xml, 28), '<w:sdt><w:sdtPr><w:alias w:val="Check 63 control 28"/><w:tag w:val="c63-28"/><w:id w:val="63028"/></w:sdtPr><w:sdtContent><w:p>'
    + '<w:ins><w:r><w:t>NEW28</w:t></w:r></w:ins></w:p></w:sdtContent></w:sdt>');
  assert.match(regionXml(xml, 47), /<w:sdtContent><w:p><w:ins><w:r><w:t>NEW47<\/w:t><\/w:r><\/w:ins><\/w:p><\/w:sdtContent>/);
  // a control's Content deleted in the bare form: the flag on over Word's filler, five spaces, untracked or not (D10)
  for (const n of [6, 32]) assert.match(regionXml(xml, n), /<w:showingPlcHdr\/><\/w:sdtPr><w:sdtContent><w:p><w:r><w:t> {5}<\/w:t><\/w:r><\/w:p><\/w:sdtContent>/, `case ${n}`);
  for (const n of [17, 43]) assert.match(regionXml(xml, n), /<w:showingPlcHdr\/><\/w:sdtPr><w:sdtContent><w:r><w:t> {5}<\/w:t><\/w:r><\/w:sdtContent>/, `case ${n}`);
  // accepted, the controls taken go and the text stays; rejected, they come back empty, their flag cleared as Word's markup has it
  const accepted = await WordprocessingMLPackage.load(await pkg.save());
  (await accepted.getBody()).acceptAll();
  const acceptedXml = (await zipText(await accepted.save()))['word/document.xml'];
  assert.equal(regionXml(acceptedXml, 45), '<w:p><w:r><w:t>NEW45</w:t></w:r></w:p>');
  assert.equal(regionXml(acceptedXml, 44).replace(/<\/w:t><\/w:r><w:r><w:t>/g, ''), '<w:p><w:r><w:t>PLAIN44A NEW44 PLAIN44B</w:t></w:r></w:p>');
  const rejected = await WordprocessingMLPackage.load(await pkg.save());
  (await rejected.getBody()).rejectAll();
  const restored = (await rejected.getBody()).contentControls.find((c) => c.tag === 'c63-45');
  assert.deepEqual([restored.isShowingPlaceholder, restored.text, restored.paragraphs[0].text], [false, '', 'PLAIN45A  PLAIN45B']);
});

test('CR-002 section 63.12, D6 (b): a paragraph put beside the only paragraph of a block control showing its placeholder goes outside it, the control left at block level', async () => {
  const { pkg } = await shimRun(await check63Bytes(), await snippet63());
  const xml = (await zipText(await pkg.save()))['word/document.xml'];
  // Word's save has the control run-level in its paragraph (section 63.2); here it stays as it was
  for (const [n, before] of [[10, false], [11, true], [36, false], [37, true]]) {
    const region = regionXml(xml, n);
    const control = /^(?:<w:p>.*?<\/w:p>)?(<w:sdt>.*<\/w:sdt>)/.exec(before ? region.replace(/^<w:p>(?:(?!<w:p>).)*?NEW\d+.*?<\/w:p>/, '') : region)?.[1];
    assert.match(control ?? '', /<w:sdtContent><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t>Click or tap here to enter text\.<\/w:t><\/w:r><\/w:p><\/w:sdtContent><\/w:sdt>/, `case ${n}`);
    assert.match(region, before ? new RegExp(`^<w:p>.*NEW${n}.*</w:p><w:sdt>`) : new RegExp(`</w:sdt><w:p>.*NEW${n}.*</w:p>$`), `case ${n}`);
  }
});

test('CR-002 section 63.12, D7: every text write into a bound control writes its node, in both forms, where the control stays', async () => {
  const { nodeOf } = await snippet63();
  // check 63's bare form: written, as Word's run 2 wrote them (run 1, in Design Mode, left every node empty)
  const bare = await shimRun(await check63Bytes(), await snippet63(), { customXml: true });
  const data63 = bare.pkg.customXmlParts.getByNamespace('urn:check63')[0].getXml();
  for (const n of [21, 22, 23, 24, 25, 47, 48, 49, 50, 51]) assert.equal(nodeOf(data63, n), `NEW${n}`, `check 63 case ${n}`);
  // where the write took the control (26, 52) nothing is bound any more, and the node stays as it was
  for (const n of [26, 52]) assert.equal(nodeOf(data63, n), '', `check 63 case ${n}`);
  // check 64, Word's form: every node written, a placeholder's empty node as much as a value's, as Word's were
  const word = await shimRun(await check64.check64Bytes(), await snippet64(), { customXml: true });
  const data64 = word.pkg.customXmlParts.getByNamespace('urn:check64')[0].getXml();
  for (let n = 43; n <= 52; n++) assert.equal(nodeOf(data64, n), `NEW${n}`, `check 64 case ${n}`);
});

test('CR-002 section 63.12, D8: placeholders in Word\'s form: the setter names an entry, an empty bound value shows the default entry, a made control shows it', async () => {
  const pkg = await WordprocessingMLPackage.load(await check63Bytes());
  const body = await pkg.getBody();
  const [block, run] = ['c63-1', 'c63-13'].map((tag) => body.contentControls.find((c) => c.tag === tag));
  // a control in the bare form names no entry, and reads null, as Word's (check 63)
  assert.deepEqual([block.placeholderEntry, block.placeholderText], [undefined, null]);
  // Word's default text names the default entry; any other a new entry of its own, 32 hex digits, as Word names one (check 65)
  block.placeholderText = PLACEHOLDER;
  run.placeholderText = 'Enter a name.';
  assert.equal(block.placeholderEntry, check65.DEFAULT_ENTRY);
  assert.match(run.placeholderEntry, /^[0-9A-F]{32}$/);
  assert.deepEqual([block.placeholderText, run.placeholderText, run.isShowingPlaceholder, run.text], [PLACEHOLDER, 'Enter a name.', true, '']);
  assert.match(await run.getXml(), /<w:placeholder><w:docPart w:val="[0-9A-F]{32}"\/><\/w:placeholder><w:showingPlcHdr\/><\/w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t>Enter a name\.<\/w:t><\/w:r><\/w:sdtContent>/);
  // set again, a new entry; the one named before stays in the glossary
  const first = run.placeholderEntry;
  run.placeholderText = 'Enter a name.';
  assert.notEqual(run.placeholderEntry, first);
  assert.equal(pkg.glossaryEntries.textOf(first), 'Enter a name.');
  // an empty bound value: the bound controls show the default entry, in Word's form, where docx4j's binding writes the
  // bare form (check 64's controls, as Word saves them)
  const bound = await WordprocessingMLPackage.load(await check63Bytes());
  await bound.getBody();
  await bound.customXmlParts.load();
  await bound.customXmlParts.applyBindings();
  const controls = (await bound.getBody()).contentControls.filter((c) => c.xmlMapping.isMapped);
  assert.equal(controls.length, 12);
  for (const c of controls) {
    assert.deepEqual([c.placeholderEntry, c.placeholderText, c.isShowingPlaceholder, c.text], [check65.DEFAULT_ENTRY, PLACEHOLDER, true, ''], c.tag);
    assert.match(await c.getXml(), /<w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t>Click or tap here to enter text\.<\/w:t>/, c.tag);
  }
  // a control made on an empty range shows the default entry; a kind Word gives another entry is made as before
  const made = await WordprocessingMLPackage.createPackage();
  const madeBody = await made.getBody();
  const point = madeBody.insertParagraph('Name: ', 'End').getRange('End').insertContentControl();
  const around = madeBody.insertParagraph('', 'End').insertContentControl('PlainText');
  const checkbox = madeBody.insertParagraph('', 'End').insertContentControl('CheckBox');
  assert.deepEqual([point.placeholderText, point.isShowingPlaceholder, around.placeholderText, around.isShowingPlaceholder], [PLACEHOLDER, true, PLACEHOLDER, true]);
  assert.deepEqual([checkbox.placeholderEntry, checkbox.isShowingPlaceholder], [undefined, false]);
  // a control around a paragraph holding text holds that text
  const filled = madeBody.insertParagraph('Kept.', 'End').insertContentControl();
  assert.deepEqual([filled.isShowingPlaceholder, filled.text], [false, 'Kept.']);
});

test('CR-002 section 63.12, D9: a placeholder written where the file has no glossary makes one on save, the document part and a styles part holding Placeholder Text; a glossary is otherwise added to, or left untouched', async () => {
  // check 63's file has no glossary
  const pkg = await WordprocessingMLPackage.load(await check63Bytes());
  const [block, run] = (await pkg.getBody()).contentControls.filter((c) => ['c63-1', 'c63-13'].includes(c.tag));
  block.placeholderText = PLACEHOLDER;
  run.placeholderText = 'Enter a name.';
  const saved = await pkg.save();
  const parts = await zipText(saved);
  assert.deepEqual(Object.keys(parts).filter((k) => k.startsWith('word/glossary/')).sort(), ['word/glossary/_rels/document.xml.rels', 'word/glossary/document.xml', 'word/glossary/styles.xml']);
  assert.match(parts['word/_rels/document.xml.rels'], /Target="glossary\/document\.xml" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/glossaryDocument"/);
  assert.match(parts['word/glossary/_rels/document.xml.rels'], /Target="styles\.xml" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/styles"/);
  assert.match(parts['[Content_Types].xml'], /<Override PartName="\/word\/glossary\/document\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document\.glossary\+xml"\/>/);
  assert.match(parts['[Content_Types].xml'], /<Override PartName="\/word\/glossary\/styles\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.styles\+xml"\/>/);
  // each entry as Word writes one (check 65): the default entry's docPartPr, a guid of its own, one run in the style
  const entry = (name, text) => new RegExp(`<w:docPart><w:docPartPr><w:name w:val="${name}"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>`
    + '<w:types><w:type w:val="bbPlcHdr"/></w:types><w:behaviors><w:behavior w:val="content"/></w:behaviors><w:guid w:val="\\{[0-9A-F-]{36}\\}"/></w:docPartPr>'
    + `<w:docPartBody><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${text.replace('.', '\\.')}</w:t></w:r></w:p></w:docPartBody></w:docPart>`);
  assert.match(parts['word/glossary/document.xml'], entry(check65.DEFAULT_ENTRY, PLACEHOLDER));
  assert.match(parts['word/glossary/document.xml'], entry(run.placeholderEntry, 'Enter a name.'));
  // the styles as Word defines them (check 66's: Word kept such a glossary's entries in their style)
  assert.match(parts['word/glossary/styles.xml'], /<w:style w:default="true" w:styleId="DefaultParagraphFont" w:type="character"><w:name w:val="Default Paragraph Font"\/>/);
  assert.match(parts['word/glossary/styles.xml'], /<w:style w:styleId="PlaceholderText" w:type="character"><w:name w:val="Placeholder Text"\/><w:basedOn w:val="DefaultParagraphFont"\/><w:uiPriority w:val="99"\/><w:unhideWhenUsed\/><w:rPr><w:color w:val="666666"\/><\/w:rPr><\/w:style>/);
  assert.match(parts['word/styles.xml'], /w:styleId="PlaceholderText"/, 'the main document has the style too');
  // reloaded, the controls read their entries; saved again, nothing is added twice
  const again = await WordprocessingMLPackage.load(saved);
  const [b2, r2] = (await again.getBody()).contentControls.filter((c) => ['c63-1', 'c63-13'].includes(c.tag));
  assert.deepEqual([b2.placeholderText, r2.placeholderText], [PLACEHOLDER, 'Enter a name.']);
  assert.equal([...(await zipText(await pkg.save()))['word/glossary/document.xml'].matchAll(/<w:docPart>/g)].length, 2);

  // 65b's glossary is the document part alone: an entry is added to it, and a styles part made, related from it
  const [, b] = check65.CHECK65_FILES;
  const minimal = await WordprocessingMLPackage.load(await check65.check65Bytes(b));
  (await minimal.getBody()).contentControls[0].placeholderText = 'Something else.';
  const minimalParts = await zipText(await minimal.save());
  assert.deepEqual([...minimalParts['word/glossary/document.xml'].matchAll(/<w:docPartPr><w:name w:val="([^"]+)"/g)].map((m) => m[1]).slice(0, 2), [check65.DEFAULT_ENTRY, check65.CUSTOM_ENTRY]);
  assert.equal([...minimalParts['word/glossary/document.xml'].matchAll(/<w:docPart>/g)].length, 3);
  assert.match(minimalParts['word/glossary/styles.xml'], /w:styleId="PlaceholderText"/);
  assert.match(minimalParts['word/glossary/_rels/document.xml.rels'], /relationships\/styles/);

  // Word's glossary (66b): read for the entries, untouched on save while nothing is written; an entry added leaves its styles part
  const [, wordFile] = check66.CHECK66_FILES;
  const source = await zipText(await check66.check66Bytes(wordFile));
  const untouched = await WordprocessingMLPackage.load(await check66.check66Bytes(wordFile));
  assert.equal((await untouched.getBody()).contentControls[0].placeholderText, PLACEHOLDER);
  const untouchedParts = await zipText(await untouched.save());
  // (a relationships part is written anew on every save)
  for (const k of Object.keys(source).filter((k) => k.startsWith('word/glossary/') && !k.includes('_rels'))) assert.equal(untouchedParts[k], source[k], k);
  (await untouched.getBody()).contentControls[0].placeholderText = 'Another.';
  const added = await zipText(await untouched.save());
  assert.equal(added['word/glossary/styles.xml'], source['word/glossary/styles.xml'], 'it holds Placeholder Text already');
  assert.match(added['word/glossary/document.xml'], /<w:t>Another\.<\/w:t>/);
});

// --- check 67 (CR-002 section 63.13): checks 58 and 63 gave the same untracked write into a control in the bare form, on
// the same markup, and Word's saves disagree: check 58's a plain run, check 63's the Placeholder Text style. Check 67 gives
// check 58's three writes in four phases (the view, and changeTrackingMode left alone or set Off) through both lookups,
// in three files (nothing else, a custom XML part, tracked changes)

const snippet67 = async () => new Function(`${await readFile(join(here, 'fixtures', 'check67', 'check67-script-lab.js'), 'utf8')}\nreturn CHECK67;`)();

test("check 67's files and snippet agree: every case parses to the case it was built as, each control in the bare form, showing its placeholder; the files differ only in what else they hold", async () => {
  const { parseRegion, PHASES } = await snippet67();
  const { unzipSync, strFromU8 } = await import('fflate');
  const cases = check67.check67Cases();
  assert.equal(cases.length, 24);
  assert.deepEqual(PHASES.map(([name]) => name), check67.PHASES);
  // check 58's writes: a block control's Replace and End, a run-level control's End; each phase through both lookups
  assert.deepEqual(cases.filter((c) => c.phase === 'simple-untouched').map((c) => `${c.op} ${c.shape} ${c.lookup}`), ['control-text-replace block tag', 'control-text-end block tag',
    'control-text-end run tag', 'control-text-replace block list', 'control-text-end block list', 'control-text-end run list']);
  for (const file of check67.CHECK67_FILES) {
    const bytes = await check67.check67Bytes(file);
    const pkg = await WordprocessingMLPackage.load(bytes);
    const body = await pkg.getBody();
    assert.deepEqual(body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null), cases, file.name);
    for (const c of cases) assert.equal(parseRegion(check67.beginText(c)).n, c.n);
    assert.match(body.paragraphs[0].text, new RegExp(`^Check 67${file.letter}: `));
    const controls = body.contentControls.filter((c) => c.tag !== 'c67-0');
    assert.deepEqual(controls.map((c) => c.tag), cases.map((c) => `c67-${c.n}`), file.name);
    assert.ok(controls.every((c) => c.isShowingPlaceholder && c.placeholderEntry === undefined && c.text === ''), file.name);
    const zip = unzipSync(bytes);
    const xml = strFromU8(zip['word/document.xml']);
    // checks 58's and 63's markup: the bare form, Word's placeholder text in the style
    assert.equal(xml.match(/<w:showingPlcHdr\/><\/w:sdtPr><w:sdtContent>(<w:p>)?<w:r><w:rPr><w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t xml:space="preserve">Click or tap here to enter text\.<\/w:t>/g).length, 24);
    assert.doesNotMatch(xml, /<w:placeholder>/);
    assert.match(strFromU8(zip['word/styles.xml']), /<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"\/><w:basedOn w:val="DefaultParagraphFont"\/><w:uiPriority w:val="99"\/><w:semiHidden\/><w:rPr><w:color w:val="666666"\/><\/w:rPr><\/w:style>/);
    // what else each holds: 67b a custom XML part and a bound control, 67c tracked changes in its first paragraph
    assert.equal(Object.keys(zip).some((k) => k.startsWith('customXml/')), file.letter === 'b', file.name);
    assert.equal(body.contentControls.some((c) => c.tag === 'c67-0'), file.letter === 'b', file.name);
    assert.equal(/<w:ins |<w:del /.test(xml), file.letter === 'c', file.name);
    if (file.letter === 'c') assert.deepEqual(body.getTrackedChanges().map((t) => [t.type, t.text]), [['Added', 'Inserted.'], ['Deleted', ' Deleted.']]);
  }
});

test("CR-002 section 63.13: check 67's snippet through the Word shim gives Word's answer, a plain run, in every phase, lookup and file", async () => {
  const { runCheck } = await snippet67();
  for (const file of check67.CHECK67_FILES) {
    const pkg = await WordprocessingMLPackage.load(await check67.check67Bytes(file));
    await pkg.getBody();
    const report = await runCheck((f) => Word.run(pkg, f), 'report');
    assert.equal(report.caseCount, 24);
    assert.deepEqual(report.cases.map(digest), check67.check67Cases().map((c) => `${c.n} CPS - ${c.shape === 'run' ? `PLAIN${c.n}A  PLAIN${c.n}B` : ''}`));
    for (const phase of check67.PHASES) {
      const result = await runCheck((f) => Word.run(pkg, f), phase);
      // each phase writes its own six cases only
      assert.deepEqual(result.cases.map((c) => c.n), check67.check67Cases().filter((c) => c.phase === phase).map((c) => c.n));
      assert.equal(result.cases.filter((c) => c.error).length, 0, `${file.letter} ${phase}`);
      // the text in the placeholder's place, plain, the flag cleared (D2 (b) since section 63.17; until then the style kept)
      assert.deepEqual(result.cases.map(digest), result.cases.map((c) => `${c.n} C-- in ${c.shape === 'run' ? `PLAIN${c.n}A FILL${c.n} PLAIN${c.n}B` : `FILL${c.n}`}`), `${file.letter} ${phase}`);
    }
    assert.equal(pkg.changeTrackingMode, 'Off');
  }
});

const wordRun67 = async (file) => JSON.parse(await readFile(join(here, 'fixtures', 'check67', `${file.name.replace('.docx', '')}-word365.json`), 'utf8'));
const save67 = async (name) => {
  const { unzipSync, strFromU8 } = await import('fflate');
  return Object.fromEntries(Object.entries(unzipSync(await readFile(join(here, 'check67', name)))).map(([k, v]) => [k, strFromU8(v)]));
};
/** A control in a save of Word's, by its tag, without Word's revision ids. */
const controlIn = (xml, tag) => [...xml.matchAll(/<w:sdt>[\s\S]*?<\/w:sdt>/g)].map((m) => m[0])
  .find((sdt) => sdt.includes(`<w:tag w:val="${tag}"/>`)).replace(/ (w:rsid\w*|w14:paraId|w14:textId)="[^"]*"/g, '');

test("check 67's run: Word's JSONs are of these files' cases, every step run in order, and its saves are Word's, as copied", async () => {
  for (const file of check67.CHECK67_FILES) {
    const word = await wordRun67(file);
    assert.deepEqual([word.check, word.host.version, Object.keys(word.steps)], [67, '16.0.20326.20232', ['report', ...check67.PHASES]]);
    assert.match(word.title, new RegExp(`^Check 67${file.letter}: `));
    assert.deepEqual(word.steps.report.cases.map(({ n, op, shape, phase, lookup }) => ({ n, op, shape, phase, lookup })), check67.check67Cases());
    for (const phase of check67.PHASES) {
      assert.deepEqual(word.steps[phase].cases.map((c) => c.n), check67.check67Cases().filter((c) => c.phase === phase).map((c) => c.n), `${file.letter} ${phase}`);
    }
  }
  const manifest = JSON.parse(await readFile(join(here, 'check67', 'manifest.json'), 'utf8'));
  assert.equal(manifest.files.length, 6);
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await readFile(join(here, 'check67', name))).digest('hex'), sha1, name);
});

test("CR-002 section 63.14: Word 365's answers (check 67): every write into a control in the bare form is a plain run, the flag cleared, in every phase, lookup and file, as check 58's were", async () => {
  for (const file of check67.CHECK67_FILES) {
    const word = await wordRun67(file);
    assert.deepEqual(word.steps.report.cases.map(digest), check67.check67Cases().map((c) => `${c.n} CPS - ${c.shape === 'run' ? `PLAIN${c.n}A  PLAIN${c.n}B` : ''}`));
    for (const phase of check67.PHASES) {
      const cases = word.steps[phase].cases;
      assert.equal(cases.filter((c) => c.error).length, 0, `${file.letter} ${phase}`);
      // no Placeholder Text style, no w:color auto: the text plain, as in Word's own form (check 64)
      assert.deepEqual(cases.map(digest), cases.map((c) => `${c.n} C-- in ${c.shape === 'run' ? `PLAIN${c.n}A FILL${c.n} PLAIN${c.n}B` : `FILL${c.n}`}`), `${file.letter} ${phase}`);
      for (const c of cases) assert.doesNotMatch(c.after.ooxml, /w:color w:val="auto"|PlaceholderText/, `${file.letter} case ${c.n}`);
    }
    // Word's save holds the same: one plain run in each control, the flag gone
    const saved = (await save67(`67${file.letter}-run-word365.docx`))['word/document.xml'];
    for (const c of check67.check67Cases()) {
      const plain = `<w:r><w:t>FILL${c.n}</w:t></w:r>`;
      assert.match(controlIn(saved, `c67-${c.n}`), new RegExp(`</w:sdtPr><w:sdtEndPr/><w:sdtContent>${c.shape === 'block' ? `<w:p>${plain}</w:p>` : plain}</w:sdtContent>`), `${file.letter} case ${c.n}`);
      assert.doesNotMatch(controlIn(saved, `c67-${c.n}`), /showingPlcHdr/, `${file.letter} case ${c.n}`);
    }
  }
  // check 58's save, the same writes (cases 20 to 22): the same plain runs, pinned here since section 63.13
  const { unzipSync, strFromU8 } = await import('fflate');
  const files = async (dir) => Object.fromEntries(Object.entries(unzipSync(await readFile(join(here, dir)))).map(([k, v]) => [k, strFromU8(v)]));
  const check58 = await files('check58/58-measures-word365.docx');
  for (const [n, block] of [[20, true], [21, true], [22, false]]) {
    const plain = `<w:r><w:t>FILL${n}</w:t></w:r>`;
    assert.match(controlIn(check58['word/document.xml'], `c58-${n}`), new RegExp(`<w:sdtContent>${block ? `<w:p>${plain}</w:p>` : plain}</w:sdtContent>`), `check 58 case ${n}`);
  }
  // only check 63's saves carry w:formsDesign (the schema's "Structured Document Tag Placeholder Text Should be Resaved"),
  // which its file didn't hold: Word wrote it in that session. Of the placeholder checks, 65b's run save carries it too
  const designed = [];
  for (const dir of ['check58', 'check63', 'check64', 'check65', 'check66', 'check67']) {
    for (const name of await readdir(join(here, dir))) {
      if (!name.endsWith('.docx')) continue;
      if (/<w:formsDesign\/>/.test((await files(`${dir}/${name}`))['word/settings.xml'])) designed.push(name);
    }
  }
  assert.deepEqual(designed.sort(), ['63-placeholders-word365.docx', '63-run-word365.docx', '65b-run-word365.docx']);
  assert.doesNotMatch(new TextDecoder().decode(unzipSync(await check63Bytes())['word/settings.xml']), /formsDesign/);
});

test("CR-002 section 63.14: the shim agrees with Word in all 72 of check 67's writes (until section 63.17 it kept the style in every one)", async () => {
  const { runCheck } = await snippet67();
  const view = (c) => ({ digest: digest(c), control: c.after.control && { type: c.after.control.type, text: c.after.control.text }, returned: c.result?.returned });
  for (const file of check67.CHECK67_FILES) {
    const word = await wordRun67(file);
    const pkg = await WordprocessingMLPackage.load(await check67.check67Bytes(file));
    await pkg.getBody();
    const differing = {};
    for (const step of ['report', ...check67.PHASES]) {
      const ours = await runCheck((f) => Word.run(pkg, f), step);
      differing[step] = word.steps[step].cases.filter((c, i) => JSON.stringify(view(c)) !== JSON.stringify(view(ours.cases[i]))).map((c) => c.n);
    }
    assert.deepEqual(differing, { report: [], ...Object.fromEntries(check67.PHASES.map((p) => [p, []])) }, file.name);
  }
});

// --- check 63, run 2 (CR-002 sections 63.15 and 63.16): the same file and snippet with Developer > Design Mode off. Run 1
// was in Design Mode (its saves carry w:formsDesign), so its bare-form answers were Design Mode's

const wordRun2 = async (mode) => JSON.parse(await readFile(join(here, 'fixtures', 'check63', `63-placeholders-word365-run2-${mode}.json`), 'utf8'));

test("check 63's run 2: Word's two JSONs are of this file's cases, run with Design Mode off", async () => {
  const cases = check63Cases();
  for (const mode of ['report', 'run']) {
    const word = await wordRun2(mode);
    assert.deepEqual([word.check, word.mode, word.host.version, word.caseCount], [63, mode, '16.0.20326.20232', 54]);
    assert.deepEqual(word.cases.map(({ n, op, shape, mode: m }) => ({ n, op, shape, ...(m ? { mode: m } : {}) })), cases);
  }
  // run 1's saves carry w:formsDesign, Design Mode's; run 2's don't
  const { unzipSync, strFromU8 } = await import('fflate');
  const designed = async (name) => /<w:formsDesign\/>/.test(strFromU8(unzipSync(await readFile(join(here, 'check63', name)))['word/settings.xml']));
  assert.deepEqual(await Promise.all(['63-placeholders-word365.docx', '63-run-word365.docx', '63-run2-opened-word365.docx', '63-run2-word365.docx'].map(designed)), [true, true, false, false]);
});

test("CR-002 section 63.16: Word 365's answers with Design Mode off (check 63, run 2): the bare form gives Word's form's answers; a Replace or delete over the whole control still takes it", async () => {
  const report = await wordRun2('report');
  const run = await wordRun2('run');
  // opened, Word reads the file as run 1 did
  assert.deepEqual(report.cases.map(digest), (await wordRun('report')).cases.map(digest));
  assert.equal(run.cases.filter((c) => c.error).length, 0, 'Office JS refused nothing');
  assert.deepEqual(run.cases.map(digest), WORD_RUN2);
  // a write into the control: plain, as in Word's form (check 64), where run 1 kept the style and added w:color auto
  for (const n of [1, 2, 4, 5, 7, 8, 9, 13, 14, 15, 16, 27, 28, 30, 31, 33, 34, 35, 39, 40, 41, 42]) {
    assert.doesNotMatch(run.cases[n - 1].after.ooxml, /PlaceholderText|w:color w:val="auto"/, `case ${n}`);
  }
  // tracked, only the insertion is recorded: the placeholder goes unmarked, as in Word's form
  for (const n of [27, 28, 29, 30, 31, 33, 34, 35, 39, 40, 41, 42, 47, 48, 49, 50, 51]) {
    assert.match(run.cases[n - 1].after.ooxml, /<w:ins /, `case ${n}`);
    assert.doesNotMatch(run.cases[n - 1].after.ooxml, /<w:del |<w:delText/, `case ${n}`);
  }
  // the control's insertParagraph clears the flag: the placeholder's paragraph left empty, the new one after it (3, 29)
  assert.match(run.cases[2].after.ooxml, /<\/w:sdtPr><w:sdtContent><w:p\/><w:p><w:r><w:t>NEW3<\/w:t><\/w:r><\/w:p><\/w:sdtContent>/);
  // a control's Content deleted: the flag stays, over a run of five spaces, untracked and tracked alike (6, 17, 32, 43)
  for (const n of [6, 17, 32, 43]) {
    assert.match(run.cases[n - 1].after.ooxml, /<w:showingPlcHdr\/><\/w:sdtPr><w:sdtContent>(<w:p>)?<w:r><w:t xml:space="preserve"> {5}<\/w:t><\/w:r>(<\/w:p>)?<\/w:sdtContent>/, `case ${n}`);
  }
  // a Replace or delete over the whole control takes it; tracked, Word's markers around it, nothing inside marked deleted
  // and the flag gone (38, 44 to 46, 52)
  for (const n of [38, 44, 45, 46, 52]) {
    const sdt = /<w:customXmlDelRangeStart [^>]*\/><w:sdt>([\s\S]*?)<\/w:sdt><w:customXmlDelRangeEnd [^>]*\/>/.exec(run.cases[n - 1].after.ooxml)?.[1];
    assert.ok(sdt, `case ${n}`);
    assert.doesNotMatch(sdt, /showingPlcHdr|<w:del |PlaceholderText/, `case ${n}`);
  }
  // every bound control's node holds the write (21 to 25, 47 to 51), as in Word's form, and the controls showed them on
  // reopening (Jason); where the write took the control (26, 52) the node stays empty
  for (const n of [21, 22, 23, 24, 25, 47, 48, 49, 50, 51]) assert.equal(run.cases[n - 1].after.node, `NEW${n}`, `case ${n}`);
  for (const n of [26, 52]) assert.equal(run.cases[n - 1].after.node, '', `case ${n}`);
  const { unzipSync, strFromU8 } = await import('fflate');
  const saved = strFromU8(unzipSync(await readFile(join(here, 'check63', '63-run2-word365.docx')))['customXml/item1.xml']);
  for (const n of [21, 22, 23, 24, 25, 47, 48, 49, 50, 51]) assert.ok(saved.includes(`<value${n}>NEW${n}</value${n}>`), `node ${n}`);
  // a paragraph beside a block control's only paragraph goes outside it, the control saved run-level, as in run 1 (10, 11)
  assert.match(run.cases[9].after.ooxml, /^<w:p>[^]*?<w:p><w:sdt>[^]*<\/w:sdt><\/w:p><w:p><w:r><w:t>NEW10<\/w:t><\/w:r><\/w:p>/);
});

test("CR-002 section 63.17: the shim gives check 63's run 2 in every write (section 63.12 followed run 1, and differed in 31)", async () => {
  const { runCheck } = await snippet63();
  const pkg = await WordprocessingMLPackage.load(await check63Bytes());
  await pkg.getBody();
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  const word = await wordRun2('run');
  const view = (c) => ({ digest: digest(c), control: c.after.control && { type: c.after.control.type, text: c.after.control.text, paragraphs: c.after.control.paragraphs }, returned: c.result?.returned });
  assert.deepEqual(word.cases.filter((c, i) => JSON.stringify(view(c)) !== JSON.stringify(view(ours.cases[i]))).map((c) => c.n), []);
  // and the markup where it tells: plain runs, only insertions tracked, Word's filler, nothing marked inside a taken control
  for (const n of [1, 2, 13, 14, 27, 28, 39, 40]) assert.doesNotMatch(ours.cases[n - 1].after.ooxml, /PlaceholderText|w:color w:val="auto"|<w:del /, `case ${n}`);
});

// Word 365's run 2 (2026-10-25, Design Mode off), case by case, as `digest` reads it (CR-002 section 63.16).
const WORD_RUN2 = [
  '1 C-- in NEW1',
  '2 C-- in NEW2',
  '3 C-- in |NEW3',
  '4 C-- in NEW4',
  '5 C-- in NEW5',
  '6 CP- - ',
  '7 C-- in NEW7',
  '8 C-- in NEW8',
  '9 C-- in NEW9',
  '10 CPS out |NEW10',
  '11 CPS out NEW11|',
  '12 x-- - ',
  '13 C-- in PLAIN13A NEW13 PLAIN13B',
  '14 C-- in PLAIN14A NEW14 PLAIN14B',
  '15 C-- in PLAIN15A NEW15 PLAIN15B',
  '16 C-- in PLAIN16A NEW16 PLAIN16B',
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
  '27 C-- in NEW27',
  '28 C-- in NEW28',
  '29 C-- in |NEW29',
  '30 C-- in NEW30',
  '31 C-- in NEW31',
  '32 CP- - ',
  '33 C-- in NEW33',
  '34 C-- in NEW34',
  '35 C-- in NEW35',
  '36 CPS out |NEW36',
  '37 CPS out NEW37|',
  '38 C-- - ',
  '39 C-- in PLAIN39A NEW39 PLAIN39B',
  '40 C-- in PLAIN40A NEW40 PLAIN40B',
  '41 C-- in PLAIN41A NEW41 PLAIN41B',
  '42 C-- in PLAIN42A NEW42 PLAIN42B',
  '43 CP- - PLAIN43A  PLAIN43B',
  '44 C-- out PLAIN44A NEW44 PLAIN44B',
  '45 C-- out NEW45',
  '46 C-- - ',
  '47 C-- in NEW47',
  '48 C-- in NEW48',
  '49 C-- in NEW49',
  '50 C-- in PLAIN50A NEW50 PLAIN50B',
  '51 C-- in PLAIN51A NEW51 PLAIN51B',
  '52 C-- out NEW52',
  '53 CP- - PLAIN53A  PLAIN53B',
  '54 CP- - ',
];

// --- CR-002 section 63.18: a placeholder shows its entry's content, and a control naming a missing entry is given Word's
// default, as docx4j's CR-063 has them (Jason, 2026-10-25: "ok do it")

test('CR-002 sections 63.18 and 63.21: a control shows a copy of the content of the entry it names, its runs as the glossary holds them, the control\'s own run properties not over them', async () => {
  const { unzipSync, zipSync, strToU8, strFromU8 } = await import('fflate');
  // 66a's glossary given an entry of two paragraphs, a bold run and a plain one; a centred block control in red and not
  // bold naming it, and a run-level one
  const [a] = check66.CHECK66_FILES;
  const zip = unzipSync(await check66.check66Bytes(a));
  const name = '6767ABCDEF0123456789ABCDEF012345';
  const entry = `<w:docPart><w:docPartPr><w:name w:val="${name}"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>`
    + '<w:types><w:type w:val="bbPlcHdr"/></w:types><w:behaviors><w:behavior w:val="content"/></w:behaviors></w:docPartPr><w:docPartBody>'
    + '<w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/><w:b/></w:rPr><w:t xml:space="preserve">Enter </w:t></w:r><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>the name,</w:t></w:r></w:p>'
    + '<w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>then the date.</w:t></w:r></w:p></w:docPartBody></w:docPart>';
  zip['word/glossary/document.xml'] = strToU8(strFromU8(zip['word/glossary/document.xml']).replace('</w:docParts>', `${entry}</w:docParts>`));
  zip['word/document.xml'] = strToU8(strFromU8(zip['word/document.xml']).replace('<w:sectPr>',
    `<w:sdt><w:sdtPr><w:rPr><w:b w:val="0"/><w:color w:val="FF0000"/></w:rPr><w:tag w:val="block"/><w:placeholder><w:docPart w:val="${name}"/></w:placeholder></w:sdtPr><w:sdtContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>VALUE</w:t></w:r></w:p></w:sdtContent></w:sdt>`
    + `<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="run"/><w:placeholder><w:docPart w:val="${name}"/></w:placeholder></w:sdtPr><w:sdtContent><w:r><w:t>VALUE</w:t></w:r></w:sdtContent></w:sdt></w:p><w:sectPr>`));
  const pkg = await WordprocessingMLPackage.load(zipSync(zip));
  const body = await pkg.getBody();
  const [block, run] = ['block', 'run'].map((tag) => body.contentControls.find((c) => c.tag === tag));
  assert.equal(block.placeholderText, 'Enter the name,\rthen the date.');
  block.showPlaceholder();
  run.showPlaceholder();
  // the block control: both paragraphs, each run as the entry holds it, the control's w:b off and red not over it, as Word
  // shows an entry (check 68, cases 21 to 24; until section 63.21 the control's properties went over it), the second
  // paragraph with a copy of the control's paragraph's properties (docx4j's binding); the run-level control: the first
  // paragraph's runs
  const bold = (rPr) => (rPr?.b === undefined ? 'unset' : rPr.b.val === false ? 'off' : 'on');
  const runs = (p) => p.runs.map((r) => [textOf(r.value), r.value.rPr?.rStyle?.val, bold(r.value.rPr), r.value.rPr?.color?.val]);
  assert.deepEqual(block.paragraphs.map(runs), [
    [['Enter ', 'PlaceholderText', 'on', undefined], ['the name,', 'PlaceholderText', 'unset', undefined]],
    [['then the date.', 'PlaceholderText', 'unset', undefined]],
  ]);
  assert.deepEqual(block.paragraphs.map((p) => p.p.pPr?.jc?.val), ['center', 'center']);
  // no placeholder text reads (a placeholder of two paragraphs reads its mark between them, '\r'; unmeasured in Word)
  assert.deepEqual([block.isShowingPlaceholder, block.text.replace(/\r/g, ''), block.placeholderEntry], [true, '', name]);
  assert.match(await run.getXml(), /<w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"\/><w:b\/><\/w:rPr><w:t xml:space="preserve">Enter <\/w:t><\/w:r><w:r><w:rPr><w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t>the name,<\/w:t><\/w:r><\/w:sdtContent>/);
  assert.deepEqual([run.isShowingPlaceholder, run.text, run.placeholderEntry], [true, '', name]);
  // the entry itself is left as it was
  assert.match(strFromU8(unzipSync(await pkg.save())['word/glossary/document.xml']), new RegExp(`<w:name w:val="${name}"/>`));
});

test('CR-002 section 63.18: a control naming an entry the glossary lacks is given Word\'s default entry when it shows its placeholder; until the glossary is read, its name is kept', async () => {
  const { unzipSync, strFromU8 } = await import('fflate');
  const [, b] = check65.CHECK65_FILES;
  const missing = (body) => body.contentControls.find((c) => c.placeholderEntry === check65.MISSING_ENTRY);
  const pkg = await WordprocessingMLPackage.load(await check65.check65Bytes(b));
  const control = missing(await pkg.getBody());
  assert.equal(control.placeholderText, null);
  control.showPlaceholder();
  // named the default entry, which 65b's glossary holds, and showing it; nothing written under the missing name
  assert.deepEqual([control.placeholderEntry, control.placeholderText, control.isShowingPlaceholder, control.text], [check65.DEFAULT_ENTRY, PLACEHOLDER, true, '']);
  const glossary = strFromU8(unzipSync(await pkg.save())['word/glossary/document.xml']);
  assert.deepEqual([...glossary.matchAll(/<w:docPartPr><w:name w:val="([^"]+)"/g)].map((m) => m[1]), [check65.DEFAULT_ENTRY, check65.CUSTOM_ENTRY]);
  // the glossary not yet read (no getBody, no Word.run): the reference is kept, and Word's default text shown
  const unread = await WordprocessingMLPackage.load(await check65.check65Bytes(b));
  const main = unread.getMainDocumentPart();
  await main.getContents();
  const kept = missing(main.body);
  kept.showPlaceholder();
  assert.deepEqual([kept.placeholderEntry, kept.isShowingPlaceholder], [check65.MISSING_ENTRY, true]);
  assert.match(await kept.getXml(), /<w:t>Click or tap here to enter text\.<\/w:t>/);
  // a date control naming a missing entry is given its own kind's default entry (section 63.22, docx4j CR-063's D8)
  const { unzipSync: unzip, zipSync, strToU8 } = await import('fflate');
  const zip = unzip(await check65.check65Bytes(b));
  zip['word/document.xml'] = strToU8(strFromU8(zip['word/document.xml']).replace('<w:sectPr>',
    `<w:sdt><w:sdtPr><w:tag w:val="date"/><w:placeholder><w:docPart w:val="${check65.MISSING_ENTRY}"/></w:placeholder><w:date/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>1 January</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr>`));
  const dated = await WordprocessingMLPackage.load(zipSync(zip));
  const date = (await dated.getBody()).contentControls.find((c) => c.tag === 'date');
  assert.deepEqual([date.type, date.defaultPlaceholderEntry?.name, date.showsWordsFormPlaceholder], ['DatePicker', 'DefaultPlaceholder_-1854013437', true]);
  date.showPlaceholder();
  assert.deepEqual([date.placeholderEntry, date.placeholderText, date.isShowingPlaceholder], ['DefaultPlaceholder_-1854013437', 'Click or tap to enter a date.', true]);
});

// --- check 68 (CR-002 section 63.19): the run properties a write into a placeholder takes, and a glossary entry under a
// control's own run properties. The placeholder run (or entry) italic, the control bold, so each case shows a property's
// source; and a Word's-form control holding a bold value, emptied, for what the entry Word brings back carries

const snippet68 = async () => new Function(`${await readFile(join(here, 'fixtures', 'check68', 'check68-script-lab.js'), 'utf8')}\nreturn CHECK68;`)();
/** A region's control after a case of check 68: `<n> <flag> <run>{<S style><b><i><c colour>} ...`. */
const look68 = (c) => {
  const sdt = /<w:sdt>([\s\S]*?)<\/w:sdt>/.exec(c.after.ooxml)?.[1] ?? '';
  const content = /<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/.exec(sdt)?.[1] ?? '';
  const runs = [...content.matchAll(/<w:r>(?:<w:rPr>([\s\S]*?)<\/w:rPr>)?<w:t[^>]*>([^<]*)<\/w:t><\/w:r>/g)].map(([, p = '', text]) =>
    `${text}{${p.includes('PlaceholderText') ? 'S' : ''}${/<w:b\/>/.test(p) ? 'b' : ''}${/<w:i\/>/.test(p) ? 'i' : ''}${/w:color/.test(p) ? 'c' : ''}}`);
  return `${c.n} ${sdt.includes('showingPlcHdr') ? 'P' : '-'} ${runs.join(' ')}`;
};

test("check 68's file and snippet agree: every case parses to the case it was built as; the bare and Word's forms, the run and control properties as built", async () => {
  const { parseRegion } = await snippet68();
  const { unzipSync, strFromU8 } = await import('fflate');
  const cases = check68.check68Cases();
  assert.equal(cases.length, 24);
  const bytes = await check68.check68Bytes();
  const pkg = await WordprocessingMLPackage.load(bytes);
  const body = await pkg.getBody();
  assert.deepEqual(body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null), cases);
  for (const c of cases) assert.equal(parseRegion(check68.beginText(c)).n, c.n);
  const controls = body.contentControls;
  assert.deepEqual(controls.map((c) => c.tag), cases.map((c) => `c68-${c.n}`));
  assert.deepEqual(controls.map((c) => c.placeholderEntry ?? null), cases.map((c) => (c.form === 'bare' ? null : c.form === 'value' && c.run === 'i' ? check68.CUSTOM_ENTRY : check68.DEFAULT_ENTRY)));
  assert.deepEqual(controls.map((c) => c.isShowingPlaceholder), cases.map((c) => c.form !== 'value'));
  assert.deepEqual(controls.map((c) => c.runProperties?.b !== undefined), cases.map((c) => c.control === 'b'));
  // the placeholder runs as built: the style, and italic or bold as the case says
  const xml = strFromU8(unzipSync(bytes)['word/document.xml']);
  for (const c of cases.filter((x) => x.form !== 'value')) {
    const extra = c.run === 'i' ? '<w:i/>' : c.run === 'b' ? '<w:b/>' : '';
    const sdt = /<w:sdt>[\s\S]*?<\/w:sdt>/g;
    const mine = [...xml.matchAll(sdt)].map((m) => m[0]).find((s) => s.includes(`<w:tag w:val="c68-${c.n}"/>`));
    assert.match(mine, new RegExp(`<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/>${extra}</w:rPr><w:t xml:space="preserve">Click or tap here to enter text\\.</w:t></w:r>`), `case ${c.n}`);
  }
  // the glossary: 66a's two parts, the custom entry's run italic
  assert.match(strFromU8(unzipSync(bytes)['word/glossary/document.xml']), new RegExp(`<w:name w:val="${check68.CUSTOM_ENTRY}"/>[\\s\\S]*<w:rStyle w:val="PlaceholderText"/><w:i/></w:rPr><w:t xml:space="preserve">Enter the name\\.</w:t>`));
  assert.ok(Object.keys(unzipSync(bytes)).includes('word/glossary/styles.xml'));
});

test("CR-002 sections 63.19 and 63.21: check 68's snippet through the Word shim: a write takes the control's own run properties, an emptied control's entry comes back as the glossary holds it", async () => {
  const { runCheck } = await snippet68();
  const pkg = await WordprocessingMLPackage.load(await check68.check68Bytes());
  await pkg.getBody();
  const report = await runCheck((f) => Word.run(pkg, f), 'report');
  assert.equal(report.caseCount, 24);
  const result = await runCheck((f) => Word.run(pkg, f), 'run');
  assert.equal(result.cases.filter((c) => c.error).length, 0);
  // until section 63.21 every write was plain and each entry came back under the control's bold
  assert.deepEqual(result.cases.map(look68), [
    ...Array.from({ length: 4 }, (_, i) => `${i + 1} - NEW${i + 1}{}`),
    ...Array.from({ length: 16 }, (_, i) => `${i + 5} - NEW${i + 5}{b}`),
    '21 P Click or tap here to enter text.{S}', '22 P Click or tap here to enter text.{S}', '23 P Enter the name.{Si}', '24 P Enter the name.{Si}',
  ]);
});

const wordRun68 = async (mode) => JSON.parse(await readFile(join(here, 'fixtures', 'check68', `68-placeholder-run-properties-word365-${mode}.json`), 'utf8'));

test("check 68's run: Word's JSONs are of this file's cases, run with Design Mode off, and its saves are Word's, as copied", async () => {
  for (const mode of ['report', 'run']) {
    const word = await wordRun68(mode);
    assert.deepEqual([word.check, word.mode, word.host.version, word.caseCount], [68, mode, '16.0.20326.20232', 24]);
    assert.deepEqual(word.cases.map(({ n, op, shape, form, run, control }) => ({ n, op, shape, form, run, control })), check68.check68Cases());
  }
  const manifest = JSON.parse(await readFile(join(here, 'check68', 'manifest.json'), 'utf8'));
  assert.equal(manifest.files.length, 2);
  const { unzipSync, strFromU8 } = await import('fflate');
  for (const { name, sha1 } of manifest.files) {
    const bytes = await readFile(join(here, 'check68', name));
    assert.equal(createHash('sha1').update(bytes).digest('hex'), sha1, name);
    assert.doesNotMatch(strFromU8(unzipSync(bytes)['word/settings.xml']), /formsDesign/, `${name}: Design Mode off`);
  }
});

test("CR-002 section 63.20: Word 365's answers (check 68): a write into a placeholder takes the control's own run properties and nothing of the placeholder's; an entry Word brings back is as the glossary holds it, the control's properties not over it", async () => {
  const run = await wordRun68('run');
  assert.equal(run.cases.filter((c) => c.error).length, 0, 'Office JS refused nothing');
  const expected = [
    // the bare form, the placeholder italic, the control with no properties: plain, the italic not taken
    '1 - NEW1{}', '2 - NEW2{}', '3 - NEW3{}', '4 - NEW4{}',
    // the control bold, in the bare form (its placeholder plain, or bold as docx4j's binder writes it) and in Word's: bold
    ...Array.from({ length: 16 }, (_, i) => `${i + 5} - NEW${i + 5}{b}`),
    // emptied, the entry in its own look, the control's bold not over it (Jason: "the placeholders in 21 to 24 do NOT")
    '21 P Click or tap here to enter text.{S}', '22 P Click or tap here to enter text.{S}', '23 P Enter the name.{Si}', '24 P Enter the name.{Si}',
  ];
  assert.deepEqual(run.cases.map(look68), expected);
  // Word's save holds the same, the controls' w:sdtPr/w:rPr bold kept
  const { unzipSync, strFromU8 } = await import('fflate');
  const saved = strFromU8(unzipSync(await readFile(join(here, 'check68', '68-run-word365.docx')))['word/document.xml']);
  assert.deepEqual(run.cases.map((c) => look68({ n: c.n, after: { ooxml: controlIn(saved, `c68-${c.n}`) } })), expected);
  for (const c of run.cases.filter((x) => x.control === 'b')) assert.match(controlIn(saved, `c68-${c.n}`), /<w:sdtPr><w:rPr><w:b\/><\/w:rPr>/, `case ${c.n}`);
});

test("CR-002 section 63.21: the shim gives check 68's answers in every case (until then it differed in 5 to 24)", async () => {
  const { runCheck } = await snippet68();
  const pkg = await WordprocessingMLPackage.load(await check68.check68Bytes());
  await pkg.getBody();
  await runCheck((f) => Word.run(pkg, f), 'report');
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  const word = await wordRun68('run');
  assert.deepEqual(word.cases.filter((c, i) => look68(c) !== look68(ours.cases[i])).map((c) => c.n), []);
});

// --- CR-002 section 63.22: Word's default placeholder entry for each kind (Jason's placeholders.docx; docx4j CR-063's D8)

test("CR-002 section 63.22: Word's default entry for each kind of control, as Word 365 made one of each (placeholders-word365.docx), is this package's", async () => {
  const manifest = JSON.parse(await readFile(join(here, 'placeholder-defaults', 'manifest.json'), 'utf8'));
  const bytes = await readFile(join(here, 'placeholder-defaults', manifest.files[0].name));
  assert.equal(createHash('sha1').update(bytes).digest('hex'), manifest.files[0].sha1);
  const pkg = await WordprocessingMLPackage.load(bytes);
  const controls = (await pkg.getBody()).contentControls;
  // every control Word made: its kind, the entry it names, and the entry's text, against this package's default for the kind
  const ours = controls.map((c) => [c.type, c.placeholderEntry ?? null, c.placeholderText, c.defaultPlaceholderEntry?.name ?? null, c.defaultPlaceholderEntry?.text ?? null]);
  assert.deepEqual(ours, [
    ['RichText', 'DefaultPlaceholder_-1854013440', 'Click or tap here to enter text.', 'DefaultPlaceholder_-1854013440', 'Click or tap here to enter text.'],
    ['PlainText', 'DefaultPlaceholder_-1854013440', 'Click or tap here to enter text.', 'DefaultPlaceholder_-1854013440', 'Click or tap here to enter text.'],
    ['Picture', null, null, null, null],
    ['BuildingBlockGallery', 'DefaultPlaceholder_-1854013436', 'Choose a building block.', 'DefaultPlaceholder_-1854013436', 'Choose a building block.'],
    ['CheckBox', null, null, null, null],
    ['ComboBox', 'DefaultPlaceholder_-1854013438', 'Choose an item.', 'DefaultPlaceholder_-1854013438', 'Choose an item.'],
    ['DropDownList', 'DefaultPlaceholder_-1854013438', 'Choose an item.', 'DefaultPlaceholder_-1854013438', 'Choose an item.'],
    ['DatePicker', 'DefaultPlaceholder_-1854013437', 'Click or tap to enter a date.', 'DefaultPlaceholder_-1854013437', 'Click or tap to enter a date.'],
    ['RepeatingSection', null, null, null, null],
    ['RepeatingSectionItem', 'DefaultPlaceholder_-1854013435', 'Enter any content that you want to repeat, including other content controls. You can also insert this control around table rows in order to repeat parts of a table.',
      'DefaultPlaceholder_-1854013435', 'Enter any content that you want to repeat, including other content controls. You can also insert this control around table rows in order to repeat parts of a table.'],
  ]);
  // each control Word made shows its entry in Placeholder Text, with the flag; the picture shows the flag over a stock picture
  for (const c of controls.filter((x) => x.placeholderEntry)) assert.deepEqual([c.isShowingPlaceholder, c.text], [true, ''], c.type);
});

test("CR-002 section 63.22: a date picker, a list and a text control each show their own kind's default entry, made or emptied; written into the glossary on save", async () => {
  const { unzipSync, strFromU8 } = await import('fflate');
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const made = Object.fromEntries(['DatePicker', 'DropDownList', 'ComboBox', 'PlainText', 'CheckBox'].map((kind) => [kind, body.insertParagraph('', 'End').insertContentControl(kind)]));
  assert.deepEqual(Object.entries(made).map(([kind, c]) => [kind, c.isShowingPlaceholder, c.placeholderEntry ?? null, c.placeholderText]), [
    ['DatePicker', true, 'DefaultPlaceholder_-1854013437', 'Click or tap to enter a date.'],
    ['DropDownList', true, 'DefaultPlaceholder_-1854013438', 'Choose an item.'],
    ['ComboBox', true, 'DefaultPlaceholder_-1854013438', 'Choose an item.'],
    ['PlainText', true, 'DefaultPlaceholder_-1854013440', 'Click or tap here to enter text.'],
    ['CheckBox', false, null, null],
  ]);
  assert.match(await made.DatePicker.getXml(), /<w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t>Click or tap to enter a date\.<\/w:t>/);
  // a date written, then its content deleted: the date picker's entry back (as Word's form does, check 66 case 3)
  made.DatePicker.insertText('25/10/2026', 'Replace');
  assert.deepEqual([made.DatePicker.isShowingPlaceholder, made.DatePicker.text], [false, '25/10/2026']);
  made.DatePicker.getRange('Content').delete();
  assert.deepEqual([made.DatePicker.isShowingPlaceholder, made.DatePicker.text, made.DatePicker.placeholderText], [true, '', 'Click or tap to enter a date.']);
  assert.match(await made.DatePicker.getXml(), /<w:t>Click or tap to enter a date\.<\/w:t>/);
  // setting a date picker's placeholder to its default text names its default entry
  made.DatePicker.placeholderText = 'Click or tap to enter a date.';
  assert.equal(made.DatePicker.placeholderEntry, 'DefaultPlaceholder_-1854013437');
  const glossary = strFromU8(unzipSync(await pkg.save())['word/glossary/document.xml']);
  assert.deepEqual([...glossary.matchAll(/<w:docPartPr><w:name w:val="([^"]+)"/g)].map((m) => m[1]).sort(),
    ['DefaultPlaceholder_-1854013437', 'DefaultPlaceholder_-1854013438', 'DefaultPlaceholder_-1854013440']);
  assert.match(glossary, /<w:name w:val="DefaultPlaceholder_-1854013437"\/>[\s\S]*?<w:t>Click or tap to enter a date\.<\/w:t>/);
  // a kind Word gives no entry, naming none, shows docx4j's run with no reference
  const picture = body.insertParagraph('', 'End').insertContentControl('Picture');
  picture.showPlaceholder();
  assert.deepEqual([picture.placeholderEntry, picture.isShowingPlaceholder], [undefined, true]);
  assert.match(await picture.getXml(), /<w:t>Click or tap here to enter text\.<\/w:t>/);
});

// --- check 69 (CR-002 section 63.23): where a paragraph's End is, when the paragraph ends in something with no width (a
// run-level control showing its placeholder, in either form; a hidden run) or in a control holding text, by every route
// to a paragraph's End (docx4j-python's CR-019 Phase C: its port wrote into the placeholder control)

const snippet69 = async () => new Function(`${await readFile(join(here, 'fixtures', 'check69', 'check69-script-lab.js'), 'utf8')}\nreturn CHECK69;`)();
/**
 * A case's paragraph after check 69, as `<n> <items> | <text>`: plain runs' text joined, a hidden run `{h:...}`, a run in
 * Placeholder Text `{ph}` (the default text) or `{S:...}`, a control `[...]` with `P ` where it shows its placeholder; and
 * Office JS's paragraph text.
 */
const look69 = (c) => {
  if (typeof c.after === 'string') return `${c.n} ${c.after}`;
  const ps = [...c.after.ooxml.matchAll(/<w:p(?: [^>]*)?>[\s\S]*?<\/w:p>/g)].map((m) => m[0]).filter((p) => !/>(BEGIN|END)\d+[ <]/.test(p));
  const items = (xml) => [...xml.matchAll(/<w:sdt>([\s\S]*?)<\/w:sdt>|<w:r(?: [^>]*)?>([\s\S]*?)<\/w:r>/g)].map(([, sdt, r]) => {
    if (sdt !== undefined) return `[${sdt.includes('<w:showingPlcHdr/>') ? 'P ' : ''}${items(/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/.exec(sdt)?.[1] ?? '')}]`;
    const text = [...r.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');
    if (/<w:vanish\/>/.test(r)) return `{h:${text}}`;
    if (/<w:rStyle w:val="PlaceholderText"\/>/.test(r)) return text === check69.PLACEHOLDER ? '{ph}' : `{S:${text}}`;
    return text;
  }).join('');
  return `${c.n} ${items(ps[ps.length - 1] ?? '')} | ${c.after.paragraph}`;
};

test("check 69's file and snippet agree: every case parses to the case it was built as; each paragraph's shape as built, the last case's paragraph the document's last", async () => {
  const { parseRegion } = await snippet69();
  const { unzipSync, strFromU8 } = await import('fflate');
  const cases = check69.check69Cases();
  assert.equal(cases.length, 49);
  const bytes = await check69.check69Bytes();
  const pkg = await WordprocessingMLPackage.load(bytes);
  const body = await pkg.getBody();
  assert.deepEqual(body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null), cases);
  for (const c of cases) assert.equal(parseRegion(check69.beginText(c)).n, c.n);
  // a control for each case of the four shapes that have one, Word's form naming the default entry
  const controls = body.contentControls;
  const controlled = cases.filter((c) => ['ph-word', 'ph-bare', 'control-text', 'ph-only'].includes(c.shape));
  assert.deepEqual(controls.map((c) => c.tag), controlled.map((c) => `c69-${c.n}`));
  assert.deepEqual(controls.map((c) => c.placeholderEntry ?? null), controlled.map((c) => (['ph-word', 'ph-only'].includes(c.shape) ? check69.DEFAULT_ENTRY : null)));
  assert.deepEqual(controls.map((c) => c.isShowingPlaceholder), controlled.map((c) => c.shape !== 'control-text'));
  // each paragraph as built: the cell cases' in a one-cell table, the last case's the document's last paragraph
  const xml = strFromU8(unzipSync(bytes)['word/document.xml']);
  for (const c of cases) {
    const at = xml.indexOf(`>BEGIN${c.n} `);
    const from = xml.indexOf('</w:p>', at) + '</w:p>'.length;
    const to = c.op === 'body-end' ? xml.indexOf('<w:sectPr>') : xml.lastIndexOf('<w:p>', xml.indexOf(`>END${c.n}<`));
    const region = xml.slice(from, to);
    assert.equal(region.includes('<w:tbl>'), c.op === 'cell-body-end', `case ${c.n}`);
    assert.ok(region.includes(check69.paragraphOf(c)), `case ${c.n}`);
  }
  assert.ok(xml.endsWith(`${check69.paragraphOf(cases[48])}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`));
  assert.ok(Object.keys(unzipSync(bytes)).includes('word/glossary/styles.xml'));
});

test("CR-002 section 63.27: check 69's snippet through the Word shim: a write at a paragraph's End goes after everything it holds, outside a trailing control; after a hidden run hidden, but by the End range's 'After'", async () => {
  const { runCheck } = await snippet69();
  const pkg = await WordprocessingMLPackage.load(await check69.check69Bytes());
  await pkg.getBody();
  const report = await runCheck((f) => Word.run(pkg, f), 'report');
  assert.equal(report.caseCount, 49);
  // before the run: a placeholder and a hidden run read as nothing, a control's value as its text
  assert.deepEqual(report.cases.filter((c) => c.op === 'paragraph-end').map(look69), [
    '1 TEXT1  | TEXT1 ', '9 TEXT9 [P {ph}] | TEXT9 ', '17 TEXT17 [P {ph}] | TEXT17 ', '25 TEXT25 {h:HIDDEN25} | TEXT25 ',
    '33 TEXT33 [VALUE33] | TEXT33 VALUE33', '41 [P {ph}] | ',
  ]);
  const result = await runCheck((f) => Word.run(pkg, f), 'run');
  assert.equal(result.cases.filter((c) => c.error).length, 0);
  // Word's answer in every case (D12 to D14); until section 63.27 the text went on the end of the last run with width,
  // before a placeholder control or a hidden run, and inside a control holding text
  assert.deepEqual(result.cases.map(look69), check69.check69Cases().map(word69));
  assert.deepEqual(result.cases.map((c) => c.result.returned), result.cases.map((c) => (c.shape === 'hidden' && c.op !== 'end-after' ? '' : `NEW${c.n}`)));
  for (const c of result.cases.filter((x) => x.after.control)) {
    assert.equal(c.after.showingPlaceholderText, c.shape !== 'control-text', `case ${c.n}`);
  }
});

const wordRun69 = async (mode) => JSON.parse(await readFile(join(here, 'fixtures', 'check69', `69-paragraph-end-word365-${mode}.json`), 'utf8'));
/** Word's answer for a check 69 case: NEW<n> after everything, outside a control; after a hidden run hidden, but by 'After'. */
const word69 = ({ n, op, shape }) => {
  const hidden = shape === 'hidden' && op !== 'end-after';
  return {
    plain: `${n} TEXT${n} NEW${n} | TEXT${n} NEW${n}`,
    'ph-word': `${n} TEXT${n} [P {ph}]NEW${n} | TEXT${n} NEW${n}`,
    'ph-bare': `${n} TEXT${n} [P {ph}]NEW${n} | TEXT${n} NEW${n}`,
    hidden: hidden ? `${n} TEXT${n} {h:HIDDEN${n}NEW${n}} | TEXT${n} ` : `${n} TEXT${n} {h:HIDDEN${n}}NEW${n} | TEXT${n} NEW${n}`,
    'control-text': `${n} TEXT${n} [VALUE${n}]NEW${n} | TEXT${n} VALUE${n}NEW${n}`,
    'ph-only': `${n} [P {ph}]NEW${n} | NEW${n}`,
  }[shape];
};
/** A case's region in a save's document part, revision and paragraph ids left out: BEGIN's paragraph to END's, or to the end. */
const region69 = (xml, n) => {
  const tidy = xml.replace(/ w:rsid\w*="[^"]*"| w14:(?:paraId|textId)="[^"]*"/g, '');
  const from = tidy.indexOf('</w:p>', tidy.indexOf(`>BEGIN${n} `)) + '</w:p>'.length;
  const end = tidy.indexOf(`>END${n}<`);
  return tidy.slice(from, end < 0 ? tidy.indexOf('<w:sectPr') : tidy.lastIndexOf('<w:p>', end));
};

test("check 69's run: Word's JSONs are of this file's cases, run with Design Mode off, and its saves are Word's, as copied", async () => {
  for (const mode of ['report', 'run']) {
    const word = await wordRun69(mode);
    assert.deepEqual([word.check, word.mode, word.host.version, word.caseCount], [69, mode, '16.0.20326.20232', 49]);
    assert.deepEqual(word.cases.map(({ n, op, shape }) => ({ n, op, shape })), check69.check69Cases());
  }
  const manifest = JSON.parse(await readFile(join(here, 'check69', 'manifest.json'), 'utf8'));
  assert.equal(manifest.files.length, 2);
  const { unzipSync, strFromU8 } = await import('fflate');
  for (const { name, sha1 } of manifest.files) {
    const bytes = await readFile(join(here, 'check69', name));
    assert.equal(createHash('sha1').update(bytes).digest('hex'), sha1, name);
    assert.doesNotMatch(strFromU8(unzipSync(bytes)['word/settings.xml']), /formsDesign/, `${name}: Design Mode off`);
  }
  // Word's save on opening holds every paragraph as built
  const opened = strFromU8(unzipSync(await readFile(join(here, 'check69', '69-opened-word365.docx')))['word/document.xml']);
  for (const c of check69.check69Cases()) {
    assert.equal(look69({ n: c.n, after: { ooxml: region69(opened, c.n), paragraph: '' } }), look69({ n: c.n, after: { ooxml: check69.paragraphOf(c), paragraph: '' } }), `case ${c.n}`);
  }
});

test("CR-002 section 63.24: Word 365's answers (check 69): a write at a paragraph's End goes after everything it holds, outside a trailing control, which keeps its placeholder; after a hidden run it is hidden, but by the End range's 'After'", async () => {
  const report = await wordRun69('report');
  assert.equal(report.cases.filter((c) => c.error).length, 0);
  const run = await wordRun69('run');
  assert.equal(run.cases.filter((c) => c.error).length, 0, 'Office JS refused nothing');
  // every route alike, but the End range's 'After' after a hidden run (case 27)
  assert.deepEqual(run.cases.map(look69), check69.check69Cases().map(word69));
  // the returned range reads the text, but where it was written hidden
  assert.deepEqual(run.cases.map((c) => c.result.returned), run.cases.map((c) => (c.shape === 'hidden' && c.op !== 'end-after' ? '' : `NEW${c.n}`)));
  // Word's save: NEW<n> a run of its own, with no properties but a hidden run's w:vanish; each control as it was
  const { unzipSync, strFromU8 } = await import('fflate');
  const saved = strFromU8(unzipSync(await readFile(join(here, 'check69', '69-run-word365.docx')))['word/document.xml']);
  for (const c of run.cases) {
    const region = region69(saved, c.n);
    const hidden = c.shape === 'hidden' && c.op !== 'end-after';
    assert.match(region, new RegExp(`<w:r>${hidden ? '<w:rPr><w:vanish/></w:rPr>' : ''}<w:t>NEW${c.n}</w:t></w:r></w:p>`), `case ${c.n}`);
    if (c.shape.startsWith('ph')) assert.match(region, /<w:showingPlcHdr\/><\/w:sdtPr><w:sdtEndPr\/><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"\/><\/w:rPr><w:t>Click or tap here to enter text\.<\/w:t><\/w:r><\/w:sdtContent><\/w:sdt><w:r>/, `case ${c.n}`);
  }
});

test("CR-002 section 63.27: the shim gives check 69's answers in every case (until then it differed in 41 of 49)", async () => {
  const { runCheck } = await snippet69();
  const pkg = await WordprocessingMLPackage.load(await check69.check69Bytes());
  await pkg.getBody();
  await runCheck((f) => Word.run(pkg, f), 'report');
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  const word = await wordRun69('run');
  assert.deepEqual(word.cases.filter((c, i) => look69(c) !== look69(ours.cases[i]) || c.result.returned !== ours.cases[i].result.returned).map((c) => c.n), []);
  // the saved markup: NEW<n> a run of its own after the control, its flag kept; hidden after a hidden run, but by 'After'
  const { unzipSync, strFromU8 } = await import('fflate');
  const saved = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
  for (const c of ours.cases) {
    const region = region69(saved, c.n);
    const hidden = c.shape === 'hidden' && c.op !== 'end-after';
    // the last run extended, but by 'After', which writes a run of its own with no properties (D14)
    if (c.shape === 'plain' && c.op !== 'end-after') assert.match(region, new RegExp(`<w:r><w:t>TEXT${c.n} NEW${c.n}</w:t></w:r></w:p>`), `case ${c.n}`);
    else if (hidden) assert.match(region, new RegExp(`<w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN${c.n}NEW${c.n}</w:t></w:r></w:p>`), `case ${c.n}`);
    else assert.match(region, new RegExp(`</w:(sdt|r)><w:r><w:t>NEW${c.n}</w:t></w:r></w:p>`), `case ${c.n}`);
  }
});

// --- check 70 (CR-002 section 63.25): where a paragraph's End is, when the paragraph ends in a bookmark, a comment's
// markers, a tracked deletion or insertion, a field or a text box; and which run's properties the text takes after a
// control (section 63.24's D13)

const snippet70 = async () => new Function(`${await readFile(join(here, 'fixtures', 'check70', 'check70-script-lab.js'), 'utf8')}\nreturn CHECK70;`)();
/**
 * A paragraph's items as tokens: a bookmark `<name>` ... `</b>`, a comment range `<c>` ... `</c>` and its reference
 * `{cref}`, a revision `[del:...]` or `[ins:...]`, a complex field `{f` instruction `i`, `|`, result, `}`, a simple field
 * `{fs:...}`, a text box `{box}`, a control `[...]` (`P ` showing its placeholder, `{ph}` the default text), a bold run
 * `*...*`, other run properties `{<rPr>:...}`.
 */
const items70 = (xml) => [...xml.matchAll(/<w:sdt>([\s\S]*?)<\/w:sdt>|<w:(del|ins) [^>]*>([\s\S]*?)<\/w:\2>|<w:fldSimple [^>]*>([\s\S]*?)<\/w:fldSimple>|<w:bookmarkStart [^>]*w:name="([^"]*)"[^>]*\/>|<w:bookmarkEnd [^>]*\/>|<w:commentRangeStart [^>]*\/>|<w:commentRangeEnd [^>]*\/>|<w:r(?: [^>]*)?>([\s\S]*?)<\/w:r>/g)]
  .map(([all, sdt, kind, revision, simple, bookmark, r]) => {
    if (sdt !== undefined) return `[${sdt.includes('<w:showingPlcHdr/>') ? 'P ' : ''}${items70(/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/.exec(sdt)?.[1] ?? '')}]`;
    if (kind !== undefined) return `[${kind}:${items70(revision)}]`;
    if (simple !== undefined) return `{fs:${items70(simple)}}`;
    if (bookmark !== undefined) return `<${bookmark}>`;
    if (all.startsWith('<w:bookmarkEnd')) return '</b>';
    if (all.startsWith('<w:commentRange')) return all.startsWith('<w:commentRangeStart') ? '<c>' : '</c>';
    const field = /<w:fldChar w:fldCharType="(begin|separate|end)"\/>/.exec(r)?.[1];
    if (field) return { begin: '{f', separate: '|', end: '}' }[field];
    if (/<w:instrText/.test(r)) return 'i';
    if (/<w:commentReference /.test(r)) return '{cref}';
    if (/<w:drawing>|<mc:AlternateContent>/.test(r)) return '{box}';
    const text = [...r.matchAll(/<w:(?:t|delText)(?: [^>]*)?>([^<]*)<\/w:(?:t|delText)>/g)].map((x) => x[1]).join('');
    const rPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(r)?.[1] ?? '';
    if (/PlaceholderText/.test(rPr)) return text === check70.PLACEHOLDER ? '{ph}' : `{S:${text}}`;
    if (rPr === '<w:b/>') return `*${text}*`;
    return rPr === '' ? text : `{${rPr}:${text}}`;
  }).join('');
/** A case after check 70: its paragraph (the one holding `TEXT<n> `, a text box's own left out) as tokens, and Office JS's text. */
const look70 = (c) => {
  if (typeof c.after === 'string') return `${c.n} ${c.after}`;
  const ps = [...c.after.ooxml.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, '<w:txbxContent/>').matchAll(/<w:p(?: [^>]*)?>[\s\S]*?<\/w:p>/g)].map((m) => m[0]);
  return `${c.n} ${items70(ps.find((p) => p.includes(`>TEXT${c.n} `)) ?? '')} | ${c.after.paragraph}`;
};

test("check 70's file and snippet agree: every case parses to the case it was built as; each paragraph's ending as built, with its comment, revision, field, text box or control", async () => {
  const { parseRegion } = await snippet70();
  const { unzipSync, strFromU8 } = await import('fflate');
  const cases = check70.check70Cases();
  assert.equal(cases.length, 88);
  const bytes = await check70.check70Bytes();
  const pkg = await WordprocessingMLPackage.load(bytes);
  const body = await pkg.getBody();
  assert.deepEqual(body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null), cases);
  for (const c of cases) assert.equal(parseRegion(check70.beginText(c)).n, c.n);
  // the controls of the D13 shapes, the placeholder in Word's form
  const controlled = cases.filter((c) => ['bold-ph', 'bold-control-text', 'control-bold-text'].includes(c.shape));
  assert.deepEqual(body.contentControls.map((c) => [c.tag, c.isShowingPlaceholder, c.placeholderEntry ?? null]),
    controlled.map((c) => [`c70-${c.n}`, c.shape === 'bold-ph', c.shape === 'bold-ph' ? check70.DEFAULT_ENTRY : null]));
  // each paragraph as built, the cell cases' in a one-cell table
  const files = unzipSync(bytes);
  const xml = strFromU8(files['word/document.xml']);
  for (const c of cases) {
    const from = xml.indexOf('</w:p>', xml.indexOf(`>BEGIN${c.n} `)) + '</w:p>'.length;
    const region = xml.slice(from, xml.lastIndexOf('<w:p>', xml.indexOf(`>END${c.n}<`)));
    assert.equal(region.includes('<w:tbl>'), c.op === 'cell-body-end', `case ${c.n}`);
    assert.ok(region.includes(check70.paragraphOf(c)), `case ${c.n}`);
  }
  assert.deepEqual(cases.map((c) => look70({ n: c.n, after: { ooxml: check70.paragraphOf(c), paragraph: '' } }).replace(/\d+/g, '#')).filter((s, i, all) => all.indexOf(s) === i), [
    '# <BM#>TEXT# </b> | ', '# TEXT# <PT#></b> | ', '# <c>TEXT# </c>{cref} | ', '# TEXT# [del:DELETED#] | ', '# TEXT# [ins:INSERTED#] | ',
    '# TEXT# {fi|FIELD#} | ', '# TEXT# {fs:FIELD#} | ', '# TEXT# {box} | ', '# *TEXT# *[P {ph}] | ', '# *TEXT# *[VALUE#] | ', '# TEXT# [*VALUE#*] | ',
  ]);
  // a comment for each comment case, and the revisions in the tracked cases, as the listing reads them
  assert.deepEqual([...strFromU8(files['word/comments.xml']).matchAll(/<w:comment w:id="(\d+)"/g)].map((m) => Number(m[1])), cases.filter((c) => c.shape === 'comment').map((c) => c.n));
  const changes = await pkg.getAllTrackedChanges();
  assert.deepEqual(changes.map((c) => [c.type, c.text, c.author]), cases.filter((c) => c.shape.startsWith('tracked'))
    .map((c) => (c.shape === 'tracked-del' ? ['Deleted', `DELETED${c.n}`, check70.AUTHOR] : ['Added', `INSERTED${c.n}`, check70.AUTHOR])));
});

test("CR-002 section 63.27: check 70's snippet through the Word shim: a write at a paragraph's End goes after everything it holds, plain, whatever the ending", async () => {
  const { runCheck } = await snippet70();
  const pkg = await WordprocessingMLPackage.load(await check70.check70Bytes());
  await pkg.getBody();
  const report = await runCheck((f) => Word.run(pkg, f), 'report');
  assert.equal(report.caseCount, 88);
  const result = await runCheck((f) => Word.run(pkg, f), 'run');
  assert.equal(result.cases.filter((c) => c.error).length, 0);
  // after every ending, by every route, plain (D12, D13); until section 63.27 the text went on the end of the last run
  // with width: inside a spanning bookmark, a comment's range and a field's result, before a point bookmark, a deletion
  // and a text box, bold after a bold run before a control, inside a control holding text
  const shapeOf = {
    'bookmark-span': (n) => `<BM${n}>TEXT${n} </b>NEW${n} | TEXT${n} NEW${n}`,
    'bookmark-point': (n) => `TEXT${n} <PT${n}></b>NEW${n} | TEXT${n} NEW${n}`,
    comment: (n) => `<c>TEXT${n} </c>{cref}NEW${n} | TEXT${n} NEW${n}`,
    'tracked-del': (n) => `TEXT${n} [del:DELETED${n}]NEW${n} | TEXT${n} NEW${n}`,
    'tracked-ins': (n) => `TEXT${n} [ins:INSERTED${n}]NEW${n} | TEXT${n} INSERTED${n}NEW${n}`,
    'field-complex': (n) => `TEXT${n} {fi|FIELD${n}}NEW${n} | TEXT${n} FIELD${n}NEW${n}`,
    'field-simple': (n) => `TEXT${n} {fs:FIELD${n}}NEW${n} | TEXT${n} FIELD${n}NEW${n}`,
    textbox: (n) => `TEXT${n} {box}NEW${n} | TEXT${n} NEW${n}`,
    'bold-ph': (n) => `*TEXT${n} *[P {ph}]NEW${n} | TEXT${n} NEW${n}`,
    'bold-control-text': (n) => `*TEXT${n} *[VALUE${n}]NEW${n} | TEXT${n} VALUE${n}NEW${n}`,
    'control-bold-text': (n) => `TEXT${n} [*VALUE${n}*]NEW${n} | TEXT${n} VALUE${n}NEW${n}`,
  };
  assert.deepEqual(result.cases.map(look70), check70.check70Cases().map((c) => `${c.n} ${shapeOf[c.shape](c.n)}`));
  assert.deepEqual(result.cases.map((c) => c.result.returned), result.cases.map((c) => `NEW${c.n}`));
  // the tracked insertion stays as it was, and the new text is outside it, untracked
  assert.equal((await pkg.getAllTrackedChanges()).filter((c) => /NEW/.test(c.text)).length, 0);
});

const wordRun70 = async (mode) => JSON.parse(await readFile(join(here, 'fixtures', 'check70', `70-paragraph-endings-word365-${mode}.json`), 'utf8'));
/** Word's answer for a check 70 case: NEW<n> after everything the paragraph holds; after a text box w:noProof by four routes. */
const word70 = ({ n, op, shape }) => {
  const after = (ending, text = '') => `${n} ${ending}NEW${n} | TEXT${n} ${text}NEW${n}`;
  const noProof = ['paragraph-end', 'end-before', 'content-end', 'cell-body-end'].includes(op);
  return {
    'bookmark-span': after(`<BM${n}>TEXT${n} </b>`),
    'bookmark-point': after(`TEXT${n} <PT${n}></b>`),
    comment: after(`<c>TEXT${n} </c>{cref}`),
    'tracked-del': after(`TEXT${n} [del:DELETED${n}]`, `DELETED${n}`),
    'tracked-ins': after(`TEXT${n} [ins:INSERTED${n}]`, `INSERTED${n}`),
    'field-complex': after(`TEXT${n} {fi|FIELD${n}}`, `FIELD${n}`),
    // Word reads a w:fldSimple as a complex field, and writes it so (its save on opening)
    'field-simple': after(`TEXT${n} {fi|FIELD${n}}`, `FIELD${n}`),
    textbox: noProof ? `${n} TEXT${n} {box}{<w:noProof/>:NEW${n}} | TEXT${n} NEW${n}` : after(`TEXT${n} {box}`),
    'bold-ph': after(`*TEXT${n} *[P {ph}]`),
    'bold-control-text': after(`*TEXT${n} *[VALUE${n}]`, `VALUE${n}`),
    'control-bold-text': after(`TEXT${n} [*VALUE${n}*]`, `VALUE${n}`),
  }[shape];
};
/** A case's region in a save of check 70, revision and paragraph ids left out. */
const region70 = (xml, n) => {
  const tidy = xml.replace(/ w:rsid\w*="[^"]*"| w14:(?:paraId|textId)="[^"]*"/g, '');
  const from = tidy.indexOf('</w:p>', tidy.indexOf(`>BEGIN${n} `)) + '</w:p>'.length;
  return tidy.slice(from, tidy.lastIndexOf('<w:p>', tidy.indexOf(`>END${n}<`)));
};
const items70Of = (c) => look70(c).split(' | ')[0];

test("check 70's run: Word's JSONs are of this file's cases, run with Design Mode off, and its saves are Word's, as copied", async () => {
  for (const mode of ['report', 'run']) {
    const word = await wordRun70(mode);
    assert.deepEqual([word.check, word.mode, word.host.version, word.caseCount], [70, mode, '16.0.20326.20232', 88]);
    assert.deepEqual(word.cases.map(({ n, op, shape }) => ({ n, op, shape })), check70.check70Cases());
  }
  const manifest = JSON.parse(await readFile(join(here, 'check70', 'manifest.json'), 'utf8'));
  assert.equal(manifest.files.length, 2);
  const { unzipSync, strFromU8 } = await import('fflate');
  for (const { name, sha1 } of manifest.files) {
    const bytes = await readFile(join(here, 'check70', name));
    assert.equal(createHash('sha1').update(bytes).digest('hex'), sha1, name);
    assert.doesNotMatch(strFromU8(unzipSync(bytes)['word/settings.xml']), /formsDesign/, `${name}: Design Mode off`);
  }
  // Word's save on opening holds every paragraph as built, but a w:fldSimple written as a complex field
  const opened = strFromU8(unzipSync(await readFile(join(here, 'check70', '70-opened-word365.docx')))['word/document.xml']);
  for (const c of check70.check70Cases()) {
    const built = items70Of({ n: c.n, after: { ooxml: check70.paragraphOf(c), paragraph: '' } });
    assert.equal(items70Of({ n: c.n, after: { ooxml: region70(opened, c.n), paragraph: '' } }), built.replace(/\{fs:(FIELD\d+)\}/, '{fi|$1}'), `case ${c.n}`);
  }
});

test("CR-002 section 63.26: Word 365's answers (check 70): a write at a paragraph's End goes after everything it holds, outside a bookmark, a comment's range and reference, a revision, a field and a text box; after a control, plain, whatever runs are about it", async () => {
  const run = await wordRun70('run');
  assert.equal(run.cases.filter((c) => c.error).length, 0, 'Office JS refused nothing');
  assert.deepEqual(run.cases.map(look70), check70.check70Cases().map(word70));
  assert.deepEqual(run.cases.map((c) => c.result.returned), run.cases.map((c) => `NEW${c.n}`));
  // Word's save holds the same; NEW<n> a run of its own, plain but after a text box (w:noProof, by four routes): not the
  // bold of the run before a control (Jason: "NEW65 is not bold; NEW73 is not bold"), nor a comment reference's style
  const { unzipSync, strFromU8 } = await import('fflate');
  const saved = strFromU8(unzipSync(await readFile(join(here, 'check70', '70-run-word365.docx')))['word/document.xml']);
  for (const c of run.cases) {
    const region = region70(saved, c.n);
    assert.equal(items70Of({ n: c.n, after: { ooxml: region, paragraph: '' } }), items70Of(c), `case ${c.n}`);
    assert.match(region, new RegExp(`<w:r>${/\{<w:noProof\/>:NEW/.test(look70(c)) ? '<w:rPr><w:noProof/></w:rPr>' : ''}<w:t>NEW${c.n}</w:t></w:r></w:p>`), `case ${c.n}`);
  }
});

test("CR-002 section 63.27: the shim writes where check 70's run wrote in every case (until then it differed in 80 of 88); what differs is a deletion's text read in All Markup, a simple field's form and w:noProof after a text box", async () => {
  const { runCheck } = await snippet70();
  const pkg = await WordprocessingMLPackage.load(await check70.check70Bytes());
  await pkg.getBody();
  await runCheck((f) => Word.run(pkg, f), 'report');
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  const word = await wordRun70('run');
  // where NEW<n> went: a simple field read as Word reads it, and Word's w:noProof after a text box left out (D13's departure)
  const placed = (c) => items70Of(c).replace(/\{fs:(FIELD\d+)\}/, '{fi|$1}').replace(/\{<w:noProof\/>:(NEW\d+)\}/, '$1');
  assert.deepEqual(word.cases.filter((c, i) => placed(c) !== placed(ours.cases[i])).map((c) => c.n), []);
  // what is left: Word's All Markup reads a deletion's text (check 53: paragraph.text follows Display for Review), a
  // w:fldSimple Word reads as a complex field, and the four routes giving w:noProof after a text box
  const differ = word.cases.filter((c, i) => look70(c) !== look70(ours.cases[i]));
  assert.deepEqual([...new Set(differ.map((c) => c.shape))], ['tracked-del', 'field-simple', 'textbox']);
  assert.deepEqual(differ.filter((c) => c.shape === 'textbox').map((c) => c.op), ['paragraph-end', 'end-before', 'content-end', 'cell-body-end']);
});

/** A package whose body is `body` (WordprocessingML, the w and r prefixes bound), its parts otherwise a created package's. */
async function withBody(body) {
  const { unzipSync, zipSync, strToU8, strFromU8 } = await import('fflate');
  const entries = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  const xml = strFromU8(entries['word/document.xml']);
  const root = /<w:document[^>]*>/.exec(xml)[0];
  entries['word/document.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${root}<w:body>${body}<w:sectPr/></w:body></w:document>`);
  const pkg = await WordprocessingMLPackage.load(zipSync(entries));
  await pkg.getBody();
  return pkg;
}
const firstParagraphXml = async (pkg) => {
  const { unzipSync, strFromU8 } = await import('fflate');
  const xml = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
  return /<w:body>(<w:p>[\s\S]*?<\/w:p>)/.exec(xml)[1].replace(/ xml:space="preserve"| w:(date|id)="[^"]*"/g, '');
};

test("CR-002 section 63.27: a write at a paragraph's End past a link, a smart tag's like, goes after it, plain; a proofing mark is no item, so the misspelt bold word before it lends its bold", async () => {
  const W_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const link = await withBody(`<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink w:anchor="x" xmlns:r="${W_REL}"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>here</w:t></w:r></w:hyperlink></w:p>`);
  link.body.paragraphs[0].insertText('.', 'End');
  assert.match(await firstParagraphXml(link), /<\/w:hyperlink><w:r><w:t>\.<\/w:t><\/w:r><\/w:p>$/);
  const proof = await withBody('<w:p><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:b/></w:rPr><w:t>Teh</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>');
  proof.body.paragraphs[0].insertText(' cat', 'End');
  assert.match(await firstParagraphXml(proof), /<w:r><w:rPr><w:b\/><\/w:rPr><w:t>Teh cat<\/w:t><\/w:r><w:proofErr w:type="spellEnd"\/><\/w:p>$/);
  // a paragraph with no runs: a run of its own, after a bookmark it holds
  const marks = await withBody('<w:p><w:bookmarkStart w:id="1" w:name="B"/><w:bookmarkEnd w:id="1"/></w:p>');
  assert.equal(marks.body.paragraphs[0].insertText('X', 'End').text, 'X');
  assert.match(await firstParagraphXml(marks), /<w:bookmarkEnd\/><w:r><w:t>X<\/w:t><\/w:r><\/w:p>$/);
});

test("CR-002 section 63.27: tracked, a write at a paragraph's End is a w:ins after everything the paragraph holds, or joins this author's own insertion that ends it", async () => {
  const pkg = await withBody('<w:p><w:r><w:t xml:space="preserve">TEXT </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="9"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click or tap here to enter text.</w:t></w:r></w:sdtContent></w:sdt></w:p>');
  pkg.changeTrackingMode = 'TrackAll';
  const p = pkg.body.paragraphs[0];
  assert.equal(p.insertText('NEW', 'End').text, 'NEW');
  assert.match(await firstParagraphXml(pkg), /<\/w:sdt><w:ins [^>]*><w:r><w:t>NEW<\/w:t><\/w:r><\/w:ins><\/w:p>$/);
  // the same author again: the insertion grows, in its run's look
  p.getRange('End').insertText('ER', 'Before');
  assert.match(await firstParagraphXml(pkg), /<\/w:sdt><w:ins [^>]*><w:r><w:t>NEWER<\/w:t><\/w:r><\/w:ins><\/w:p>$/);
  assert.deepEqual((await pkg.getAllTrackedChanges()).map((c) => [c.type, c.text]), [['Added', 'NEWER']]);
});

test("CR-002 section 63.27: a write at a paragraph's End holding a line break makes paragraphs after everything the paragraph holds, and the range runs across them; after a hidden run such text is visible", async () => {
  const pkg = await withBody('<w:p><w:r><w:t xml:space="preserve">TEXT </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:r><w:t>VALUE</w:t></w:r></w:sdtContent></w:sdt></w:p>');
  const range = pkg.body.paragraphs[0].insertText('A\nB', 'End');
  assert.deepEqual(pkg.body.paragraphs.map((p) => p.text), ['TEXT VALUEA', 'B']);
  assert.equal(range.text, 'A\rB');
  assert.equal(pkg.body.contentControls[0].text, 'VALUE');
  const hidden = await withBody('<w:p><w:r><w:t xml:space="preserve">TEXT </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN</w:t></w:r></w:p>');
  hidden.body.paragraphs[0].insertText('A\nB', 'End');
  assert.deepEqual(hidden.body.paragraphs.map((p) => p.text), ['TEXT A', 'B']);
  // without a break, after the hidden run, hidden (D13): the range reads nothing, as Word's
  const one = await withBody('<w:p><w:r><w:t xml:space="preserve">TEXT </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN</w:t></w:r></w:p>');
  assert.equal(one.body.paragraphs[0].insertText('A', 'End').text, '');
  assert.equal(one.body.paragraphs[0].text, 'TEXT ');
});

test("CR-002 section 63.27: a search hit that ends at a paragraph's end is no End route: text put after it goes where it went before, inside a trailing control's text", async () => {
  const pkg = await withBody('<w:p><w:r><w:t xml:space="preserve">TEXT </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:r><w:t>VALUE</w:t></w:r></w:sdtContent></w:sdt></w:p>');
  const [hit] = pkg.body.search('VALUE');
  hit.insertText('!', 'After');
  assert.equal(pkg.body.contentControls[0].text, 'VALUE!');
});
