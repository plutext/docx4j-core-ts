// CR-002 section 65: the measurement setters write a value only where Word 365 stores its attribute as written (docx4j
// CR-066 section 3.1, round 5: 16 bits, wrapped on reading, no error), and check 71, what Office JS's setters do there.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { unzipSync, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../dist/index.mjs';
import { Word } from '../dist/office-js/index.mjs';
import * as check71 from './fixtures/check71/build.mjs';
import * as check71b from './fixtures/check71b/build.mjs';
import { WORD_STORAGE } from '../dist/model/content/wordRange.mjs';

const here = dirname(fileURLToPath(import.meta.url));

/**
 * The range each setter writes within, in points, and the attribute it writes (CR-002 section 65): w:ind's left and right
 * signed 16-bit; its first line and hanging, one signed 16-bit field in Word's binary format (check 71), so 0 to 32767 and 0 to 32768;
 * a cell's width 0 to 32767 (signed, never negative). The spacing, line, list and font setters stop where Word 365's own do,
 * below storage (check 71b, section 65.9, D2): spacing 0 to 1584 pt (22 inches), a line 0.75 to 1584 pt under an auto rule,
 * a list's text indent short of 1584 pt either way, a font 1 to 1638 pt in half-points after rounding.
 */
const RANGES = {
  leftIndent: [-1638.4, 1638.35, 'w:ind/@w:left'],
  rightIndent: [-1638.4, 1638.35, 'w:ind/@w:right'],
  firstLineIndent: [-1638.4, 1638.35, 'w:ind/@w:'],   // w:firstLine, or w:hanging for a negative
  spaceBefore: [0, 1584, 'w:spacing/@w:before'],
  spaceAfter: [0, 1584, 'w:spacing/@w:after'],
  lineSpacing: [0.75, 1584, 'w:spacing/@w:line'],
  columnWidth: [0, 1638.35, 'w:tcW/@w:w'],
  listIndent: [-1583.95, 1583.95, 'w:ind/@w:left'],
  fontSize: [1, 1638, 'w:sz/@w:val'],
};
/** Whether this package writes a case's value: within its setter's range (a list case's label indent is 0; a font's size in half-points, rounded). */
const writes = (c) => {
  const v = c.setter === 'fontSize' ? Math.round(c.value * 2) / 2 : c.value;
  return v >= RANGES[c.setter][0] && v <= RANGES[c.setter][1];
};

const snippet71 = async () => new Function(`${await readFile(join(here, 'fixtures', 'check71', 'check71-script-lab.js'), 'utf8')}\nreturn CHECK71;`)();

test("check 71's file and snippet agree: every case parses to the case it was built as, its target as built", async () => {
  const { parseRegion } = await snippet71();
  const cases = check71.check71Cases();
  assert.equal(cases.length, 96);
  for (const c of cases) assert.deepEqual(parseRegion(check71.beginText(c)), c);
  const pkg = await WordprocessingMLPackage.load(await check71.check71Bytes());
  const body = await pkg.getBody();
  const regions = body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null);
  assert.deepEqual(regions, cases);
  for (const c of cases) assert.ok(body.paragraphs.some((p) => p.text === `TARGET${c.n}`), `case ${c.n}`);
  assert.equal(body.tables.length, cases.filter((c) => c.setter === 'columnWidth').length, "a table for each columnWidth case, none else");
});

test("check 71's snippet through the Word shim: report only changes nothing; the run sets each value within its attribute's range and refuses the rest with InvalidArgument, the target left as built", async () => {
  const { runCheck } = await snippet71();
  const pkg = await WordprocessingMLPackage.load(await check71.check71Bytes());
  await pkg.getBody();
  const before = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
  const report = await runCheck((f) => Word.run(pkg, f), 'report');
  assert.equal(report.caseCount, 96);
  assert.equal(strFromU8(unzipSync(await pkg.save())['word/document.xml']), before, 'report only changes nothing');
  const result = await runCheck((f) => Word.run(pkg, f), 'run');
  for (const c of result.cases) {
    const label = `case ${c.n}: ${c.setter} ${c.value}`;
    if (writes(c)) {
      assert.equal(c.result, 'set', label);
      if (c.setter === 'listIndent') assert.equal(c.after.read.leftIndent, c.value, label);
      else assert.equal(c.after.read, c.value, label);
    } else {
      assert.match(c.error, /^error: InvalidArgument: /, label);
      assert.match(c.error, /\(CR-002 section 65\)$/, label);
      if (c.setter !== 'listIndent') assert.ok(c.error.includes(RANGES[c.setter][2]) || /not a measurement/.test(c.error), label);
      // nothing written: the target holds no element a write would have made, but a cell's width as built
      const made = c.after.ooxml.document.filter((e) => !/^<w:numPr>/.test(e));
      assert.deepEqual(made, c.setter === 'columnWidth' ? ['<w:tcW w:type="dxa" w:w="2880"/>'] : [], label);
    }
  }
  const refused = result.cases.filter((c) => c.error).length;
  assert.equal(refused, result.cases.filter((c) => !writes(c)).length);
  assert.equal(refused, 80, 'the values past each range, of 96: 60 past storage, 20 within it but past Word\'s own limits (check 71b)');
});

test('CR-002 section 65: each setter refuses a value past the range Word 365 stores its attribute in, before anything changes, naming the range and what Word would read', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const p = body.insertParagraph('text', 'End');
  const xmlOf = () => JSON.stringify(p.element.value.pPr ?? null);
  const unchanged = xmlOf();
  assert.throws(() => { p.leftIndent = 400000; }, (e) => e.code === 'InvalidArgument'
    && e.message === 'Paragraph.leftIndent: 400000 pt is 8000000 twips in w:ind/@w:left, outside the range written there (-32768 to 32767 twips, -1638.4 to 1638.35 pt): Word 365 would read it as 4608 (CR-002 section 65)');
  assert.equal(xmlOf(), unchanged, 'no w:ind made');
  // the ends of each range written; one past refused
  p.leftIndent = 1638.35; p.leftIndent = -1638.4;
  assert.throws(() => { p.leftIndent = 1638.4; }, /would read it as -32768/);
  assert.throws(() => { p.rightIndent = -1638.45; }, /w:ind\/@w:right.*would read it as 32767/);
  p.firstLineIndent = 1638.35; p.firstLineIndent = -1638.4;
  assert.throws(() => { p.firstLineIndent = 1638.4; }, /w:ind\/@w:firstLine.*would read it as a hanging indent of 32768 /, "check 71 case 25: Word's own setter wrote it so");
  assert.throws(() => { p.firstLineIndent = -1638.45; }, /a hanging indent.*w:ind\/@w:hanging.*would read it as a first-line indent of 32767 /, 'check 71 case 33');
  // spacing: Word stores 0 to 65535, its own setter stops at 22 inches (check 71b), so the policy's reason, not Word's reading
  p.spaceBefore = 1584; p.spaceAfter = 0;
  assert.throws(() => { p.spaceBefore = 1584.05; }, /w:spacing\/@w:before.*\(0 to 31680 twips, 0 to 1584 pt\): Word's own setter stops at 22 inches, 1584 pt \(check 71b\); its storage runs to 65535, which its save writes back as -1 \(CR-002/);
  assert.throws(() => { p.spaceBefore = 3276.7; }, /stops at 22 inches/, 'within storage, past the policy');
  assert.throws(() => { p.spaceAfter = -1; }, /w:spacing\/@w:after.*would read it as 65516/, "past storage: Word's reading");
  // a line: Word's own setter writes an auto rule, 0.75 to 1584 pt (check 71b)
  p.lineSpacing = 1584;
  assert.deepEqual([p.element.value.pPr.spacing.line, p.element.value.pPr.spacing.lineRule], [31680, 'auto'], "as Word's setter writes it");
  assert.throws(() => { p.lineSpacing = 1584.05; }, /w:spacing\/@w:line.*\(15 to 31680 twips, 0.75 to 1584 pt\): Word's own setter takes 0.75 to 1584 pt/);
  assert.throws(() => { p.lineSpacing = 0.7; }, /Word's own setter takes 0.75 to 1584 pt/);
  assert.throws(() => { p.lineSpacing = 0; }, /Word's own setter takes/);
  assert.throws(() => { p.lineSpacing = 3276.75; }, /would read it as auto -1 /, 'an auto line of 65535 twips wraps to -1: past storage');
  assert.throws(() => { p.leftIndent = Number.NaN; }, (e) => e.code === 'InvalidArgument' && /NaN is not a measurement in points/.test(e.message));
  assert.throws(() => { p.spaceBefore = Number.POSITIVE_INFINITY; }, /Infinity is not a measurement/);
  assert.deepEqual([p.leftIndent, p.firstLineIndent, p.spaceBefore, p.spaceAfter, p.lineSpacing], [-1638.4, -1638.4, 1584, 0, 1584]);

  // a font's size: w:sz is 0 to 65535 half-points; Word's own setter takes 1 to 1638 pt, after rounding to a half-point (check 71b)
  const font = p.getRange('Whole').font;
  font.size = 1638; font.size = 1638.05; font.size = 0.9;
  assert.equal(p.element.value.content.find((c) => c.name?.localPart === 'r')?.value.rPr?.sz?.val ?? p.getRange('Whole').font.size * 2, 2, '0.9 pt rounds to 2 half-points, as Word wrote it');
  assert.throws(() => { font.size = 1638.4; }, /Font\.size: 1638.4 pt is 3277 half-points in w:sz\/@w:val.*\(2 to 3276 half-points, 1 to 1638 pt\): Word's own setter takes 1 to 1638 pt/);
  assert.throws(() => { font.size = 0.5; }, /Font\.size: 0.5 pt is 1 half-points.*Word's own setter takes 1 to 1638 pt/);
  assert.throws(() => { font.size = 32768; }, /Font\.size: 32768 pt is 65536 half-points in w:sz\/@w:val.*would read it as 0 /, 'past storage: what Word reads');
  assert.equal(font.size, 1);

  // a cell's width, in twips or in a percentage table's fiftieths; never negative
  const table = body.insertTable(1, 1, 'End', [['a']]);
  const cell = table.getCell(0, 0);
  cell.width = 1638.35;
  assert.throws(() => { cell.width = 1638.4; }, /TableCell\.width.*w:tcW\/@w:w.*would read it as -32768/);
  assert.throws(() => { cell.columnWidth = -1; }, /a cell's width is never negative/);
  assert.equal(cell.width, 1638.35);
  table.element.value.tblPr.tblW = { w: 5000, type: 'pct' };
  cell.element.value.tcPr.tcW = { w: 5000, type: 'pct' };
  assert.throws(() => { cell.width = 20000; }, /fiftieths of a percent in w:tcW\/@w:w/);

  // a list level's indents: its text's w:ind/@w:left, and the label's first-line or hanging indent relative to it
  const item = body.insertParagraph('item', 'End');
  const list = await item.startNewList();
  list.setLevelIndents(0, 1583.95, -1638.4); list.setLevelIndents(0, -1583.95, -1638.4); list.setLevelIndents(0, 1583.95, -1638.4);
  assert.throws(() => list.setLevelIndents(0, 1584, 0), /List\.setLevelIndents \(textIndent\): 1584 pt is 31680 twips in w:ind\/@w:left.*\(-31679 to 31679 twips, -1583.95 to 1583.95 pt\): Word's own setLevelIndents stops short of 22 inches/, "check 71b: Word's setLevelIndents refused 1584 and -1584, took 1583.95 and -1583.95");
  assert.throws(() => list.setLevelIndents(0, -1584, 0), /stops short of 22 inches/);
  assert.throws(() => list.setLevelIndents(0, 1638.4, 0), /List\.setLevelIndents \(textIndent\).*would read it as -32768/, 'past storage');
  assert.throws(() => list.setLevelIndents(0, 36, 1638.4), /\(bulletNumberPictureIndent\).*w:ind\/@w:firstLine/);
  assert.throws(() => list.setLevelIndents(0, 36, -1638.45), /a hanging bulletNumberPictureIndent.*w:ind\/@w:hanging/);
  assert.deepEqual([item.leftIndent, item.firstLineIndent], [1583.95, -1638.4]);
});

// --- check 71's run (Jason, 2026-10-26; Word 365 16.0.20326.20232, PC, Design Mode off)

const word71 = async () => JSON.parse(await readFile(join(here, 'fixtures', 'check71', '71-measurement-setters-word365-run.json'), 'utf8'));
const range = (from, to) => Array.from({ length: to - from + 1 }, (_, i) => from + i);
/** The elements a case changes in Word's save, as the snippet reads them from a range's OOXML. */
function savedElements(xml) {
  const out = new Map();
  for (const part of xml.split(/(?=<w:p\b(?:(?!<\/w:p>).)*?BEGIN\d+ )/)) {
    const m = /BEGIN(\d+) /.exec(part);
    if (!m) continue;
    const end = part.indexOf(`END${m[1]}<`);
    const region = end > 0 ? part.slice(0, end) : part;
    out.set(Number(m[1]), (region.match(/<w:ind\b[^>]*\/>|<w:spacing\b[^>]*\/>|<w:tcW\b[^>]*\/>|<w:sz\b[^>]*\/>/g) ?? []).map((e) => e.replace(/ w:rsid\w*="[^"]*"/g, '')));
  }
  return out;
}

test("check 71, Word 365's run: the indent setters and a cell's width wrap in 16 bits as Word reads a file, the spacing, line, list and font setters refuse past Word's limits with InvalidArgument; Word's save holds what the run read", async () => {
  const word = await word71();
  assert.deepEqual([word.check, word.mode, word.caseCount, word.host.version], [71, 'run', 96, '16.0.20326.20232']);
  assert.deepEqual(word.cases.map(({ n, setter, value }) => ({ n, setter, value })), check71.check71Cases());
  const refused = word.cases.filter((c) => c.error).map((c) => c.n);
  for (const c of word.cases) if (c.error) assert.equal(c.error, 'error: InvalidArgument: InvalidArgument', `case ${c.n}`);
  // refused: every spacing and line value here (each past 1584 pt or negative), the list's text indent from 1638 pt and
  // past -1638.4, a negative width, and a font past 1638 pt or at 0
  assert.deepEqual(refused, [...range(34, 66), 75, 76, 77, ...range(78, 85), 87, 88, ...range(90, 96)]);
  // what Word wrote where it wrote: the 16-bit field's reading of the twips, as WORD_STORAGE has it
  const twipsOf = (c) => Math.round(c.value * 20);
  for (const c of word.cases.filter((x) => !x.error && ['leftIndent', 'rightIndent', 'columnWidth'].includes(x.setter))) {
    const attribute = { leftIndent: 'w:left', rightIndent: 'w:right', columnWidth: 'w:w' }[c.setter];
    const stored = WORD_STORAGE.indLeft.reads(twipsOf(c));
    const found = c.after.ooxml.document.find((e) => e.includes(`${attribute}="`)) ?? null;
    assert.equal(found === null ? '0' : /="(-?\d+)"/.exec(found.slice(found.indexOf(attribute)))[1], stored, `case ${c.n}`);
  }
  // a first-line indent and a hanging one share one signed field: firstLine F as F, hanging H as -H
  const firstLine = Object.fromEntries(word.cases.filter((c) => c.setter === 'firstLineIndent').map((c) => [c.value, c.after.ooxml.document.join('')]));
  assert.deepEqual(firstLine, {
    1638: '<w:ind w:firstLine="32760"/>', 1638.35: '<w:ind w:firstLine="32767"/>', 1638.4: '<w:ind w:hanging="32768"/>',
    1639: '<w:ind w:hanging="32756"/>', 3276.7: '<w:ind w:hanging="2"/>', 3276.75: '<w:ind w:hanging="1"/>', 3276.8: '',
    400000: '<w:ind w:firstLine="4608"/>', '-1': '<w:ind w:hanging="20"/>', '-1638.4': '<w:ind w:hanging="32768"/>',
    '-1638.45': '<w:ind w:firstLine="32767"/>',
  });
  // and WORD_STORAGE reads each as Word stored it
  assert.deepEqual([WORD_STORAGE.indFirstLine.reads(32780), WORD_STORAGE.indFirstLine.reads(65534), WORD_STORAGE.indFirstLine.reads(8000000),
    WORD_STORAGE.indHanging.reads(20), WORD_STORAGE.indHanging.reads(32769)],
  ['a hanging indent of 32756', 'a hanging indent of 2', '4608', '20', 'a first-line indent of 32767'], 'cases 26, 27, 30, 31 and 33');
  // a list level: (-1, 0) as w:left="-20" w:firstLine="0", the label's indent relative to the text's (section 65.5)
  assert.deepEqual(word.cases.find((c) => c.n === 86).after.read, { leftIndent: -1, firstLineIndent: 0 });
  const save = await readFile(join(here, 'check71', '71-run-word365.docx'));
  const manifest = JSON.parse(await readFile(join(here, 'check71', 'manifest.json'), 'utf8'));
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await readFile(join(here, 'check71', name))).digest('hex'), sha1, name);
  const files = unzipSync(new Uint8Array(save));
  assert.match(strFromU8(files['word/numbering.xml']), /<w:ind w:left="-20" w:firstLine="0"\/>/);
  // Word's save holds what the run read from each region's OOXML (a font's w:sz in every run of the region but one)
  const saved = savedElements(strFromU8(files['word/document.xml']));
  for (const c of word.cases.filter((x) => x.setter !== 'fontSize' && x.setter !== 'listIndent')) {
    assert.deepEqual(saved.get(c.n), c.after.ooxml.document, `case ${c.n}`);
  }
});

test("check 71: where this package's setters differ from Word 365's, every case listed: it refuses what Word's indent and width setters wrap (a departure, D1); where Word's spacing, line, list and font setters refuse below their storage it refuses too, since check 71b (D2)", async () => {
  const word = await word71();
  const { runCheck } = await snippet71();
  const pkg = await WordprocessingMLPackage.load(await check71.check71Bytes());
  await pkg.getBody();
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  const wraps = [];
  const refusesBelow = [];
  for (const [i, c] of ours.cases.entries()) {
    const theirs = word.cases[i];
    if (Boolean(c.error) === Boolean(theirs.error)) continue;
    (c.error ? wraps : refusesBelow).push(c.n);
  }
  // Word writes a value its own reader then misreads; this package refuses it (section 65.5, D1)
  assert.deepEqual(wraps, [...range(3, 8), 11, ...range(14, 19), 22, ...range(25, 30), 33, ...range(69, 74)]);
  // Word refuses these within the storage range, at its own limits; so does this package since check 71b found them (section 65.9, D2):
  // until then it wrote cases 34 to 38, 45 to 49, 56 to 58, 78, 79, 87, 90 to 92 and 95
  assert.deepEqual(refusesBelow, []);
  for (const n of [...range(34, 38), ...range(45, 49), ...range(56, 58), 78, 79, 87, 90, 91, 92, 95]) assert.match(ours.cases[n - 1].error, /^error: InvalidArgument: /, `case ${n}`);
  // where both write, both write the same: the twips, and Word's own reading back
  for (const [i, c] of ours.cases.entries()) {
    const theirs = word.cases[i];
    if (c.error || theirs.error || c.setter === 'listIndent' || c.setter === 'fontSize') continue;
    // attributes in Word's order: the shim writes w:tcW's w:type first
    const inWordsOrder = (e) => e.replace(/ w:type="dxa" w:w="(-?\d+)"/, ' w:w="$1" w:type="dxa"');
    assert.deepEqual(c.after.ooxml.document.map(inWordsOrder), theirs.after.ooxml.document, `case ${c.n}`);
  }
});

// --- check 71b (CR-002 section 65.5 D2, 2026-10-27; Jason: "D1 and D2 as recommended"): where Word's own setters refuse, below storage

test("check 71b's file and snippet agree: 51 cases at and either side of Word's limits, check 71's snippet unchanged", async () => {
  const js71 = await readFile(join(here, 'fixtures', 'check71', 'check71-script-lab.js'), 'utf8');
  assert.equal(await readFile(join(here, 'fixtures', 'check71b', 'check71b-script-lab.js'), 'utf8'), js71, "the snippet is check 71's, a copy");
  const { parseRegion } = await snippet71();
  const cases = check71b.check71bCases();
  assert.equal(cases.length, 51);
  assert.deepEqual([...new Set(cases.map((c) => c.setter))], ['spaceBefore', 'spaceAfter', 'lineSpacing', 'listIndent', 'fontSize'], 'the setters check 71 found refusing below storage');
  for (const c of cases) assert.deepEqual(parseRegion(check71.beginText(c)), c);
  const pkg = await WordprocessingMLPackage.load(await check71b.check71bBytes());
  const body = await pkg.getBody();
  assert.deepEqual(body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null), cases);
  for (const c of cases) assert.ok(body.paragraphs.some((p) => p.text === `TARGET${c.n}`), `case ${c.n}`);
});

const word71b = async (mode) => JSON.parse(await readFile(join(here, 'fixtures', 'check71b', `71b-setter-limits-word365-${mode}.json`), 'utf8'));
/** Word's single-precision reading of a value (1583.94995117188 for 1583.95), as the value it is. */
const near = (a, b) => Math.abs(a - b) < 0.001;
const sameRead = (ours, theirs, label) => {
  if (typeof ours === 'object') { for (const k of Object.keys(theirs)) assert.ok(near(ours[k], theirs[k]), `${label}: ${k} ${ours[k]} vs ${theirs[k]}`); return; }
  assert.ok(near(ours, theirs), `${label}: ${ours} vs ${theirs}`);
};
/** Cases refused by Word: InvalidArgument at its own limits, below storage (section 65.9). */
const REFUSED_71B = [...range(5, 10), ...range(15, 20), ...range(21, 25), 30, 31, 32, 34, 35, 36, 37, 39, 40, 41, 42, 43, 50, 51];

test("check 71b, Word 365's run: its spacing and line setters take 0 to 1584 pt, the line under an auto rule from 0.75 pt, setLevelIndents a text indent short of 1584 pt either way, font.size 1 to 1638 pt after rounding; each refusal InvalidArgument; its save holds what the run read", async () => {
  const report = await word71b('report');
  const word = await word71b('run');
  const cases = check71b.check71bCases();
  for (const r of [report, word]) {
    assert.deepEqual([r.check, r.caseCount, r.host.version, r.cases.length], [71, 51, '16.0.20326.20232', 51]);
    assert.deepEqual(r.cases.map(({ n, setter, value }) => ({ n, setter, value })), cases);
  }
  // report only: the file as built, no element a case would write (a font case's run has no w:sz)
  for (const c of report.cases) { assert.equal(c.result, undefined, `case ${c.n}`); assert.deepEqual(c.after.ooxml.document, [], `case ${c.n}`); }
  // the run
  assert.deepEqual(word.cases.filter((c) => c.error).map((c) => c.n), REFUSED_71B);
  for (const c of word.cases) {
    if (c.error) assert.equal(c.error, 'error: InvalidArgument: InvalidArgument', `case ${c.n}`);
    else assert.equal(c.result, 'set', `case ${c.n}`);
  }
  const byCase = new Map(word.cases.map((c) => [c.n, c]));
  const written = (n) => byCase.get(n).after.ooxml.document;
  // spacing: 22 inches, 31680 twips, written; 31681 refused; a value equal to the style's (before 0) written as nothing
  assert.deepEqual([written(1), written(2), written(3), written(4), written(11), written(12), written(14)],
    [[], ['<w:spacing w:before="1"/>'], ['<w:spacing w:before="31679"/>'], ['<w:spacing w:before="31680"/>'], ['<w:spacing w:after="0"/>'], ['<w:spacing w:after="1"/>'], ['<w:spacing w:after="31680"/>']]);
  for (const n of [...range(5, 10), ...range(15, 20)]) assert.deepEqual(written(n), [], `case ${n}: nothing written`);
  // a line: 20 × points under w:lineRule="auto", 0.75 to 1584 pt; 0.7 and 1584.05 refused; the paragraph's 13.8 pt (the style's 276 auto) left
  assert.deepEqual([written(26), written(27), written(28), written(29)],
    [['<w:spacing w:line="15" w:lineRule="auto"/>'], ['<w:spacing w:line="20" w:lineRule="auto"/>'], ['<w:spacing w:line="31679" w:lineRule="auto"/>'], ['<w:spacing w:line="31680" w:lineRule="auto"/>']]);
  for (const n of [...range(21, 25), 30, 31, 32]) { assert.deepEqual(written(n), [], `case ${n}`); assert.ok(near(byCase.get(n).after.read, 13.8), `case ${n}: unchanged`); }
  // a list's text indent: 1583.95 and -1583.95 taken, read back as the paragraph's left indent; 1584 and -1584 refused, after
  // startNewList had put the paragraph in a new list with Word's default level (36 pt, hanging 18)
  sameRead(byCase.get(33).after.read, { leftIndent: 1583.95, firstLineIndent: 0 }, 'case 33');
  sameRead(byCase.get(38).after.read, { leftIndent: -1583.95, firstLineIndent: 0 }, 'case 38');
  for (const n of [34, 35, 36, 37, 39, 40, 41, 42]) {
    assert.deepEqual(byCase.get(n).after.read, { leftIndent: 36, firstLineIndent: -18 }, `case ${n}`);
    assert.match(written(n)[0], /^<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="\d+"\/><\/w:numPr>$/, `case ${n}: in the list startNewList made`);
  }
  // a font: half-points after rounding, 2 to 3276 (0.9 pt as 2, 1638.05 as 3276; 0.5 and 1638.4 refused); written to the run and the mark
  const sz = (n) => [...new Set(written(n))];
  assert.deepEqual([sz(44), sz(45), sz(46), sz(47), sz(48), sz(49)], [['<w:sz w:val="2"/>'], ['<w:sz w:val="2"/>'], ['<w:sz w:val="3"/>'], ['<w:sz w:val="3275"/>'], ['<w:sz w:val="3276"/>'], ['<w:sz w:val="3276"/>']]);
  assert.deepEqual([byCase.get(44).after.read, byCase.get(49).after.read, byCase.get(43).after.read, byCase.get(50).after.read], [1, 1638, 11, 11]);
  for (const n of [43, 50, 51]) assert.deepEqual(written(n), [], `case ${n}`);
  // Word's save, under its manifest, holds what the run read
  const manifest = JSON.parse(await readFile(join(here, 'check71b', 'manifest.json'), 'utf8'));
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await readFile(join(here, 'check71b', name))).digest('hex'), sha1, name);
  const files = unzipSync(new Uint8Array(await readFile(join(here, 'check71b', '71b-run-word365.docx'))));
  const saved = savedElements(strFromU8(files['word/document.xml']));
  for (const c of word.cases.filter((x) => x.setter !== 'fontSize' && x.setter !== 'listIndent')) assert.deepEqual(saved.get(c.n), c.after.ooxml.document, `case ${c.n}`);
  const numbering = strFromU8(files['word/numbering.xml']);
  assert.match(numbering, /<w:ind w:left="31679" w:firstLine="0"\/>/, 'case 33');
  assert.match(numbering, /<w:ind w:left="-31679" w:firstLine="0"\/>/, 'case 38');
  assert.equal((numbering.match(/<w:ind w:left="720" w:hanging="360"\/>/g) ?? []).length, 8, 'the eight lists whose indents were refused keep the default level');
});

test("check 71b's snippet through the Word shim, held to Word 365's run case for case (D2): the same cases refused, the same values read back and the same elements written, but spaceBefore 0, which Word writes as nothing where the style gives 0", async () => {
  const word = await word71b('run');
  const { runCheck } = await snippet71();
  const pkg = await WordprocessingMLPackage.load(await check71b.check71bBytes());
  await pkg.getBody();
  const before = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
  const report = await runCheck((f) => Word.run(pkg, f), 'report');
  assert.equal(report.caseCount, 51);
  assert.equal(strFromU8(unzipSync(await pkg.save())['word/document.xml']), before, 'report only changes nothing');
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  assert.deepEqual(ours.cases.filter((c) => c.error).map((c) => c.n), REFUSED_71B, "refused where Word refuses");
  const differ = [];
  const withoutNumId = (es) => es.map((e) => e.replace(/<w:numId w:val="\d+"\/>/, '<w:numId/>'));
  for (const [i, c] of ours.cases.entries()) {
    const theirs = word.cases[i];
    const label = `case ${c.n}: ${c.setter} ${c.value}`;
    if (c.error) {
      // the policy's reason within storage (check 71b), Word's reading past it (a negative spacing, 65535)
      assert.match(c.error, /^error: InvalidArgument: .*(\(check 71b\)|would read it as)/, label);
      assert.ok(writes(c) === false, label);
    } else {
      assert.equal(c.result, 'set', label);
      assert.ok(writes(c), label);
    }
    sameRead(c.after.read, theirs.after.read, label);
    // the elements: a list case's numId is the package's own count; a font's w:sz once here, in the run and the mark in Word's
    const [mine, wordsOwn] = c.setter === 'listIndent' ? [withoutNumId(c.after.ooxml.document), withoutNumId(theirs.after.ooxml.document)]
      : c.setter === 'fontSize' ? [[...new Set(c.after.ooxml.document)], [...new Set(theirs.after.ooxml.document)]]
        : [c.after.ooxml.document, theirs.after.ooxml.document];
    if (JSON.stringify(mine) !== JSON.stringify(wordsOwn)) differ.push([c.n, mine, wordsOwn]);
  }
  // Word writes nothing for a value equal to what the paragraph already has from its style (before 0); this package writes the value
  assert.deepEqual(differ, [[1, ['<w:spacing w:before="0"/>'], []]]);
});
