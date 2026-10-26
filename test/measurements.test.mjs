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
import { WORD_STORAGE } from '../dist/model/content/wordRange.mjs';

const here = dirname(fileURLToPath(import.meta.url));

/**
 * The range each setter writes within, in points, and the attribute it writes (CR-002 section 65): w:ind's left and right
 * signed 16-bit; its first line and hanging, one signed 16-bit field in Word's binary format (check 71), so 0 to 32767 and 0 to 32768;
 * w:spacing's before and after unsigned, 0xFFFF left out; an exact line 1 to 32768 (stored negated); a cell's width 0 to
 * 32767 (signed, never negative); w:sz 0 to 65535 half-points.
 */
const RANGES = {
  leftIndent: [-1638.4, 1638.35, 'w:ind/@w:left'],
  rightIndent: [-1638.4, 1638.35, 'w:ind/@w:right'],
  firstLineIndent: [-1638.4, 1638.35, 'w:ind/@w:'],   // w:firstLine, or w:hanging for a negative
  spaceBefore: [0, 3276.7, 'w:spacing/@w:before'],
  spaceAfter: [0, 3276.7, 'w:spacing/@w:after'],
  lineSpacing: [0.05, 1638.4, 'w:spacing/@w:line'],
  columnWidth: [0, 1638.35, 'w:tcW/@w:w'],
  listIndent: [-1638.4, 1638.35, 'w:ind/@w:left'],
  fontSize: [0, 32767.5, 'w:sz/@w:val'],
};
/** Whether this package writes a case's value: within its setter's range (a list case's label indent is 0). */
const writes = (c) => c.value >= RANGES[c.setter][0] && c.value <= RANGES[c.setter][1];

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
  assert.equal(refused, 60, 'the values past each range, of 96');
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
  p.spaceBefore = 3276.7; p.spaceAfter = 0;
  assert.throws(() => { p.spaceBefore = 3276.75; }, /w:spacing\/@w:before.*would read it as 65535/);
  assert.throws(() => { p.spaceAfter = -1; }, /w:spacing\/@w:after.*would read it as 65516/);
  p.lineSpacing = 1638.4;
  assert.throws(() => { p.lineSpacing = 3276.75; }, /would read it as atLeast 1 /, "an exact line of 65535 twips reads as at least 1 (CR-066's round 5)");
  assert.throws(() => { p.lineSpacing = 0; }, /would read it as atLeast 0 /);
  assert.throws(() => { p.leftIndent = Number.NaN; }, (e) => e.code === 'InvalidArgument' && /NaN is not a measurement in points/.test(e.message));
  assert.throws(() => { p.spaceBefore = Number.POSITIVE_INFINITY; }, /Infinity is not a measurement/);
  assert.deepEqual([p.leftIndent, p.firstLineIndent, p.spaceBefore, p.spaceAfter, p.lineSpacing], [-1638.4, -1638.4, 3276.7, 0, 1638.4]);

  // a font's size: w:sz is 0 to 65535 half-points; Word's UI stops at 1638 pt, its storage doesn't
  const font = p.getRange('Whole').font;
  font.size = 32767.5;
  assert.throws(() => { font.size = 32768; }, /Font\.size: 32768 pt is 65536 half-points in w:sz\/@w:val.*would read it as 0 /);
  assert.equal(font.size, 32767.5);

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
  list.setLevelIndents(0, 1638.35, -1638.4);
  assert.throws(() => list.setLevelIndents(0, 1638.4, 0), /List\.setLevelIndents \(textIndent\).*w:ind\/@w:left/);
  assert.throws(() => list.setLevelIndents(0, 36, 1638.4), /\(bulletNumberPictureIndent\).*w:ind\/@w:firstLine/);
  assert.throws(() => list.setLevelIndents(0, 36, -1638.45), /a hanging bulletNumberPictureIndent.*w:ind\/@w:hanging/);
  assert.deepEqual([item.leftIndent, item.firstLineIndent], [1638.35, -1638.4]);
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

test("check 71: where this package's setters differ from Word 365's, every case listed: it refuses what Word's indent and width setters wrap (a departure), and writes what Word's spacing, line, list and font setters refuse below their storage (Word's limits not yet measured)", async () => {
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
  // Word refuses these within the storage range, at its own limits (1584 pt for spacing, 1638 pt for a font, 0 refused):
  // where exactly is check 71b's question, so they stay written here until it answers (section 65.5, D2)
  assert.deepEqual(refusesBelow, [...range(34, 38), ...range(45, 49), ...range(56, 58), 78, 79, 87, 90, 91, 92, 95]);
  // where both write, both write the same: the twips, and Word's own reading back
  for (const [i, c] of ours.cases.entries()) {
    const theirs = word.cases[i];
    if (c.error || theirs.error || c.setter === 'listIndent' || c.setter === 'fontSize') continue;
    // attributes in Word's order: the shim writes w:tcW's w:type first
    const inWordsOrder = (e) => e.replace(/ w:type="dxa" w:w="(-?\d+)"/, ' w:w="$1" w:type="dxa"');
    assert.deepEqual(c.after.ooxml.document.map(inWordsOrder), theirs.after.ooxml.document, `case ${c.n}`);
  }
});
