// CR-001 Phase B step 3: list numbering (docx4j `org.docx4j.model.listnumbering`, CR-014).
//
// The parity test (test/parity.test.mjs) holds the whole model to docx4j's answers on 45
// documents.  This file is the unit net under it: the formatter table, and the docx4j tests
// whose documents or rules the parity fixtures do not reach - `LabelFormatterTest`,
// `NumberingRestartTest` (probe P8), `NumberingStoriesTest` (P7), `StartOverrideTest`,
// `IsLglTest`, `ListNumberIndTest`, `StyleLinkedLevelTest` and `DefaultStyleNumberedTest`,
// each named where it is ported.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import {
  WordprocessingMLPackage, NumberingDefinitionsPart, ZipPartStore,
  Emulator, NumberingState, NumberingStates, LevelDefinition,
  formatValue, formatterFor, register, resetWarnings,
  parseXml, unmarshalNode, setLogger,
} from '../dist/index.mjs';
import { fixture, fixturesDir, bytesEqual } from './helpers.mjs';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

// ---------------------------------------------------------------- the formatter table

/**
 * docx4j `LabelFormatterTest`, row for row: the `w:numFmt`, a counter value, the label.  The
 * boundaries are the interesting rows - where a formatter runs out (Roman 3999, the circled
 * digits at 20), where letters repeat (26/27/28, 52/53), the teens and hundreds of the English
 * words, the Hebrew exceptions for 15 and 16 - and a value a format cannot express gives the
 * decimal label (CR-014 phase 1).
 */
const FORMATTER_CASES = [
  ['decimal', 1, '1'], ['decimal', 100, '100'], ['decimalHalfWidth', 10, '10'],
  ['decimalZero', 1, '01'], ['decimalZero', 9, '09'], ['decimalZero', 10, '10'], ['decimalZero', 100, '100'],
  ['none', 5, ''], ['bullet', 5, '*'],

  ['upperRoman', 1, 'I'], ['upperRoman', 4, 'IV'], ['upperRoman', 1994, 'MCMXCIV'],
  ['upperRoman', 3999, 'MMMCMXCIX'], ['upperRoman', 4000, '4000'],
  ['lowerRoman', 1, 'i'], ['lowerRoman', 49, 'xlix'], ['lowerRoman', 0, '0'],

  ['lowerLetter', 1, 'a'], ['lowerLetter', 26, 'z'], ['lowerLetter', 27, 'aa'], ['lowerLetter', 28, 'bb'],
  ['lowerLetter', 52, 'zz'], ['lowerLetter', 53, 'aaa'], ['lowerLetter', 0, '0'],
  ['upperLetter', 1, 'A'], ['upperLetter', 26, 'Z'], ['upperLetter', 27, 'AA'], ['upperLetter', 28, 'BB'],

  ['ordinal', 1, '1st'], ['ordinal', 2, '2nd'], ['ordinal', 3, '3rd'], ['ordinal', 4, '4th'],
  ['ordinal', 11, '11th'], ['ordinal', 12, '12th'], ['ordinal', 13, '13th'], ['ordinal', 21, '21st'],
  ['ordinal', 22, '22nd'], ['ordinal', 23, '23rd'], ['ordinal', 101, '101st'], ['ordinal', 111, '111th'],
  ['ordinal', 112, '112th'], ['ordinal', 113, '113th'],

  ['cardinalText', 0, 'Zero'], ['cardinalText', 1, 'One'], ['cardinalText', 10, 'Ten'],
  ['cardinalText', 13, 'Thirteen'], ['cardinalText', 20, 'Twenty'], ['cardinalText', 21, 'Twenty-One'],
  ['cardinalText', 99, 'Ninety-Nine'], ['cardinalText', 100, 'One Hundred'],
  ['cardinalText', 101, 'One Hundred One'], ['cardinalText', 110, 'One Hundred Ten'],
  ['cardinalText', 1000, 'One Thousand'], ['cardinalText', 1001, 'One Thousand One'],
  ['cardinalText', 1234, 'One Thousand Two Hundred Thirty-Four'],
  ['cardinalText', 2000000, 'Two Million'],

  ['ordinalText', 1, 'First'], ['ordinalText', 2, 'Second'], ['ordinalText', 3, 'Third'],
  ['ordinalText', 4, 'Fourth'], ['ordinalText', 5, 'Fifth'], ['ordinalText', 8, 'Eighth'],
  ['ordinalText', 9, 'Ninth'], ['ordinalText', 11, 'Eleventh'], ['ordinalText', 12, 'Twelfth'],
  ['ordinalText', 13, 'Thirteenth'], ['ordinalText', 19, 'Nineteenth'], ['ordinalText', 20, 'Twentieth'],
  ['ordinalText', 21, 'Twenty-First'], ['ordinalText', 30, 'Thirtieth'], ['ordinalText', 42, 'Forty-Second'],
  ['ordinalText', 100, 'One Hundredth'], ['ordinalText', 101, 'One Hundred First'],
  ['ordinalText', 112, 'One Hundred Twelfth'], ['ordinalText', 120, 'One Hundred Twentieth'],
  ['ordinalText', 1000, 'One Thousandth'], ['ordinalText', 1001, 'One Thousand First'],

  ['hex', 1, '1'], ['hex', 10, 'A'], ['hex', 255, 'FF'], ['hex', 256, '100'],
  ['chicago', 1, '*'], ['chicago', 2, '†'], ['chicago', 3, '‡'], ['chicago', 4, '§'],
  ['chicago', 5, '**'], ['chicago', 8, '§§'], ['chicago', 9, '***'],
  ['numberInDash', 1, '- 1 -'], ['numberInDash', 12, '- 12 -'],

  ['decimalFullWidth', 1, '１'], ['decimalFullWidth', 10, '１０'],
  ['decimalFullWidth2', 7, '７'],
  ['thaiNumbers', 10, '๑๐'], ['hindiNumbers', 10, '१०'],

  ['russianLower', 1, 'а'], ['russianLower', 10, 'к'], ['russianLower', 28, 'я'],
  ['russianLower', 29, 'аа'], ['russianUpper', 1, 'А'], ['russianUpper', 30, 'ББ'],
  ['arabicAlpha', 1, 'أ'], ['arabicAlpha', 28, 'ي'], ['arabicAlpha', 29, 'أأ'],
  ['thaiLetters', 1, 'ก'], ['thaiLetters', 42, 'ฮ'], ['thaiLetters', 43, 'กก'],

  ['hebrew1', 1, 'א'], ['hebrew1', 10, 'י'], ['hebrew1', 11, 'יא'],
  ['hebrew1', 15, 'טו'], ['hebrew1', 16, 'טז'], ['hebrew1', 17, 'יז'],
  ['hebrew1', 20, 'כ'], ['hebrew1', 99, 'צט'], ['hebrew1', 100, 'ק'],
  ['hebrew1', 400, 'ת'], ['hebrew1', 500, 'תק'],
  ['hebrew1', 999, 'תתקצט'], ['hebrew1', 0, '0'],

  ['decimalEnclosedCircle', 1, '①'], ['decimalEnclosedCircle', 20, '⑳'],
  ['decimalEnclosedCircle', 21, '21'], ['decimalEnclosedCircleChinese', 2, '②'],
  ['chineseCounting', 1, '一'], ['chineseLegalSimplified', 1, '壹'],

  // no formatter registered: the decimal label
  ['japaneseCounting', 7, '7'], ['koreanDigital', 3, '3'],
];

test('LabelFormatterTest: the format table', () => {
  setLogger({ warn() { /* the fallbacks warn once each; not this test's subject */ } });
  try {
    const failures = [];
    for (const [numFmt, value, expected] of FORMATTER_CASES) {
      const actual = formatValue(numFmt, value);
      if (actual !== expected) failures.push(`${numFmt} ${value}: expected '${expected}' got '${actual}'`);
    }
    assert.deepEqual(failures, [], failures.join('\n'));
  } finally {
    setLogger();
  }
});

test('LabelFormatterTest: an undefined w:numFmt is decimal', () => {
  assert.equal(formatValue(undefined, 3), '3');
});

test('LabelFormatterTest: a formatter itself still says when a value is out of range', () => {
  assert.throws(() => formatterFor('upperRoman')(4000), RangeError);
  assert.throws(() => formatterFor('decimalEnclosedCircle')(21), RangeError);
  assert.throws(() => formatterFor('lowerLetter')(0), RangeError);
});

test('LabelFormatterTest: register replaces a formatter, and undefined removes it', () => {
  const before = formatterFor('koreanDigital');
  assert.equal(before, undefined);
  try {
    register('koreanDigital', (n) => `k${n}`);
    assert.equal(formatValue('koreanDigital', 3), 'k3');
  } finally {
    register('koreanDigital', before);
  }
  assert.equal(formatterFor('koreanDigital'), undefined);
});

test('the fallback to decimal warns once per format, naming where it first happened', () => {
  resetWarnings();
  const warnings = [];
  setLogger({ warn: (m) => warnings.push(m) });
  try {
    assert.equal(formatValue('upperRoman', 4000, 'numId 3 ilvl 0'), '4000');
    assert.equal(formatValue('upperRoman', 5000, 'numId 4 ilvl 1'), '5000');
    assert.equal(warnings.length, 1, warnings.join('\n'));
    assert.match(warnings[0], /upperRoman cannot express 4000/);
    assert.match(warnings[0], /first at numId 3 ilvl 0/);
  } finally {
    setLogger();
    resetWarnings();
  }
});

// ---------------------------------------------------------------- helpers

/** A package created from scratch with a numbering part of the given `w:numbering` body. */
async function packageWith(numberingBody, styles = []) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const ndp = new NumberingDefinitionsPart();
  ndp.setXml(`<w:numbering ${W}>${numberingBody}</w:numbering>`);
  pkg.getMainDocumentPart().addTargetPart(ndp);
  if (styles.length > 0) {
    const part = pkg.getMainDocumentPart().styleDefinitionsPart;
    const live = await part.getContents();
    live.style.push(...await unmarshalStyles(styles));
  }
  // the resolver was built by createPackage, before these parts existed
  await pkg.refresh();
  return pkg;
}

async function unmarshalStyles(fragments) {
  const element = await unmarshalNode(parseXml(`<w:styles ${W}>${fragments.join('')}</w:styles>`));
  return element.value.style;
}

async function pPrOf(xml) {
  const element = await unmarshalNode(parseXml(`<w:document ${W}><w:body><w:p>${xml}</w:p></w:body></w:document>`));
  return element.value.body.content[0].value.pPr;
}

/** The top-level `w:p` of a loaded document's body, as docx4j's `AbstractNumberingTest` walks it. */
function topLevelParagraphs(document) {
  return (document.body?.content ?? [])
    .map((item) => (item && typeof item === 'object' && 'value' in item ? item.value : item))
    .filter((value) => value?.TYPE_NAME === 'org_docx4j_wml.P');
}

/** A paragraph's text, from every `w:t` in it. */
function textOfParagraph(p) {
  let out = '';
  const walk = (value) => {
    if (Array.isArray(value)) { for (const item of value) walk(item); return; }
    if (typeof value !== 'object' || value === null) return;
    if ('value' in value && typeof value.name === 'object' && value.name !== null) { walk(value.value); return; }
    if (value.TYPE_NAME === 'org_docx4j_wml.Text') { out += value.value ?? ''; return; }
    for (const key of Object.keys(value)) {
      if (key === 'PARENT' || key === 'TYPE_NAME') continue;
      walk(value[key]);
    }
  };
  walk(p.content ?? []);
  return out;
}

/**
 * docx4j `AbstractNumberingTest.testNumbering`: walk the body's top-level paragraphs in order,
 * number each that has a `w:pPr`, and assert wherever the text carries `[expect]...[/expect]`.
 */
async function assertExpectedLabels(name) {
  const pkg = await WordprocessingMLPackage.load(await fixture(join('parity', name)));
  const emulator = await pkg.getNumberingEmulator();
  assert.ok(emulator, `${name} has a numbering part`);
  const document = await pkg.getMainDocumentPart().getContents();
  let assertions = 0;
  for (const p of topLevelParagraphs(document)) {
    if (p.pPr === undefined) continue;
    const result = emulator.getNumber(p.pPr);
    const match = /\[expect\](.*?)\[\/expect\]/s.exec(textOfParagraph(p));
    if (match === null) continue;
    assert.equal(result?.numString, match[1], `${name}: ${textOfParagraph(p).slice(0, 60)}`);
    assertions++;
  }
  assert.ok(assertions > 0, `${name}: no assertions were tested`);
  return assertions;
}

// ---------------------------------------------------------------- the docx4j oracle documents

test('StartOverrideTest: w:startOverride is spent on the w:num\'s first use in a story', async () => {
  // docx4j StartOverrideTest, startOverride.docx: two w:num over one w:abstractNum, the second
  // with <w:startOverride w:val="10"/>, and the labels Word shows written into the text
  const assertions = await assertExpectedLabels('startOverride.docx');
  assert.equal(assertions, 12);
});

test('IsLglTest: w:isLgl decimalises the levels a label inherits', async () => {
  // docx4j IsLglTest: Word's built-in Article/Section numbering, with and without w:isLgl at
  // ilvl 1 ("Section 1.01" against "Section I.01")
  await assertExpectedLabels('article-section-isLgl.docx');
  await assertExpectedLabels('article-section-NotIsLgl.docx');
});

test('ListNumberIndTest: the indent a level contributes, level first then its linked style', async () => {
  // docx4j ListNumberIndTest (org.docx4j.model.listnumbering, moved there by CR-014 phase 3),
  // over the seven flat OPC documents of its `ind/` directory, copied here unchanged.
  const cases = [
    ['abstract_style_with', 2880],      // the level's own w:ind, not the style's (17.1.0)
    ['abstract_style_without', 12],
    ['abstract_nostyle_ppr', 13],
    ['abstract_nostyle_noppr', undefined],
    ['override_nostyle_ppr', 23],
    ['abstract_style_ind_only', 11],    // no w:ind of its own: the linked style's (17.1.1)
    ['abstract_style_basedon', 31],     // ... found through w:basedOn (17.1.1)
  ];
  for (const [name, expected] of cases) {
    const xml = await readFile(join(fixturesDir, 'ind', `${name}.xml`), 'utf8');
    const pkg = await WordprocessingMLPackage.load(xml);
    await pkg.getPropertyResolver();
    const ind = pkg.numberingDefinitionsPart.getIndOf('1', '0');
    assert.equal(ind?.left, expected, name);
  }
});

// ---------------------------------------------------------------- w:lvlRestart (probe P8)

const RESTART_WALK = [0, 1, 2, 2, 1, 2, 0, 2];

/** Three levels, `%1.` / `%1.%2.` / `%1.%2.%3.`, with `w:lvlRestart` on level 2. */
function abstractNum(id, lvlRestart) {
  let out = `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="multilevel"/>`;
  const text = ['%1.', '%1.%2.', '%1.%2.%3.'];
  for (let ilvl = 0; ilvl < 3; ilvl++) {
    out += `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>`;
    if (ilvl === 2 && lvlRestart !== undefined) out += `<w:lvlRestart w:val="${lvlRestart}"/>`;
    out += `<w:lvlText w:val="${text[ilvl]}"/><w:lvlJc w:val="left"/>`
      + `<w:pPr><w:ind w:left="${720 + 720 * ilvl}" w:hanging="720"/></w:pPr></w:lvl>`;
  }
  return `${out}</w:abstractNum>`;
}

test('NumberingRestartTest: w:lvlRestart as Word labels it (CR-014 probe P8)', async () => {
  const pkg = await packageWith(
    abstractNum(80, 0) + abstractNum(81, 1) + abstractNum(82, undefined)
    + '<w:num w:numId="80"><w:abstractNumId w:val="80"/></w:num>'
    + '<w:num w:numId="81"><w:abstractNumId w:val="81"/></w:num>'
    + '<w:num w:numId="82"><w:abstractNumId w:val="82"/></w:num>');

  const walk = (numId) => RESTART_WALK
    .map((ilvl) => Emulator.getNumber(pkg, undefined, numId, String(ilvl)).numString)
    .join(' ');

  // list A, w:val 0: level 2 never restarts (Word golden, verbatim)
  assert.equal(walk('80'), '1. 1.1. 1.1.1. 1.1.2. 1.2. 1.2.3. 2. 2.1.4.');
  // list B, w:val 1: only a level-0 item restarts level 2 (Word golden, verbatim)
  assert.equal(walk('81'), '1. 1.1. 1.1.1. 1.1.2. 1.2. 1.2.3. 2. 2.1.1.');
  // no w:lvlRestart: any shallower level restarts it
  assert.equal(walk('82'), '1. 1.1. 1.1.1. 1.1.2. 1.2. 1.2.1. 2. 2.1.1.');
});

test('NumberingRestartTest: restartsAfter, level by level', () => {
  const levelWith = (lvlRestart) => {
    const level = new LevelDefinition('2');
    level.lvlRestart = lvlRestart;
    return level;
  };
  assert.equal(levelWith(undefined).lvlRestart, undefined);
  // no w:lvlRestart: every shallower level restarts a level-2 counter
  assert.equal(levelWith(undefined).restartsAfter(0), true);
  assert.equal(levelWith(undefined).restartsAfter(1), true);
  // 0: none does
  assert.equal(levelWith(0).restartsAfter(0), false);
  assert.equal(levelWith(0).restartsAfter(1), false);
  // 1: ilvl 0 only
  assert.equal(levelWith(1).restartsAfter(0), true);
  assert.equal(levelWith(1).restartsAfter(1), false);
  // 2: ilvl 0 and 1 (the same as none, for a level-2 counter)
  assert.equal(levelWith(2).restartsAfter(1), true);
});

// ---------------------------------------------------------------- stories (probe P7)

/** The numbered paragraphs directly in a content list (not inside a text box). */
function numbered(content) {
  return (content ?? [])
    .map((item) => (item && typeof item === 'object' && 'value' in item ? item.value : item))
    .filter((value) => value?.TYPE_NAME === 'org_docx4j_wml.P' && value.pPr?.numPr !== undefined);
}

function labels(emulator, paragraphs, state) {
  return paragraphs.map((p) => emulator.getNumber(p.pPr, state)?.numString ?? 'null').join(' ');
}

/** The first `w:txbxContent` under a content list. */
function textBoxContent(content) {
  let found;
  const walk = (value) => {
    if (found !== undefined) return;
    if (Array.isArray(value)) { for (const item of value) walk(item); return; }
    if (typeof value !== 'object' || value === null) return;
    if ('value' in value && typeof value.name === 'object' && value.name !== null) { walk(value.value); return; }
    if (value.TYPE_NAME === 'org_docx4j_wml.CTTxbxContent') { found = value; return; }
    for (const key of Object.keys(value)) {
      if (key === 'PARENT' || key === 'TYPE_NAME') continue;
      walk(value[key]);
    }
  };
  walk(content);
  return found;
}

test('NumberingStoriesTest: each story counts on its own (CR-014 probe P7)', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture(join('parity', 'numbering-stories.docx')));
  const main = pkg.getMainDocumentPart();
  const emulator = await pkg.getNumberingEmulator();
  const states = new NumberingStates();

  const document = await main.getContents();
  const body = numbered(document.body.content);
  assert.equal(body.length, 6);
  assert.equal(labels(emulator, body.slice(0, 3), states.forPart(main)), '1. 2. 3.');

  // the other stories, read between the body's two halves
  const header = main.headerParts[0];
  const footer = main.footerParts[0];
  assert.equal(states.forPart(header), states.forPart(footer), 'a header and a footer share one story');
  assert.equal(labels(emulator, numbered((await header.getContents()).content), states.forPart(header)), '1. 2. 3.');
  assert.equal(labels(emulator, numbered((await footer.getContents()).content), states.forPart(footer)), '4. 5. 6.');

  const footnotes = main.footnotesPart;
  const fn = [];
  for (const note of (await footnotes.getContents()).footnote ?? []) {
    if (note.type !== undefined) continue;   // separator, continuationSeparator
    fn.push(labels(emulator, numbered(note.content), states.forPart(footnotes)));
  }
  assert.equal(fn.join(' | '), '1. 2. 3. | 4. 5. 6.');

  const endnotes = main.endnotesPart;
  let en = '';
  for (const note of (await endnotes.getContents()).endnote ?? []) {
    if (note.type !== undefined) continue;
    en += labels(emulator, numbered(note.content), states.forPart(endnotes));
  }
  assert.equal(en, '1. 2. 3.');

  const comments = main.commentsPart;
  let cm = '';
  for (const comment of (await comments.getContents()).comment ?? []) {
    cm += labels(emulator, numbered(comment.content), states.forPart(comments));
  }
  assert.equal(cm, '1. 2. 3.');

  const txbx = textBoxContent(document.body.content);
  assert.ok(txbx, 'the document has a text box');
  assert.equal(labels(emulator, numbered(txbx.content), states.newStory()), '1. 2. 3.');

  // and the body carries on where it left off
  assert.equal(labels(emulator, body.slice(3, 6), states.forPart(main)), '4. 5. 6.');

  // none of that touched the numbering part's own default state
  assert.equal(labels(emulator, body, undefined), '1. 2. 3. 4. 5. 6.');
  assert.notEqual(states.main(), pkg.numberingDefinitionsPart.numberingState);
});

test('NumberingStoriesTest: peek takes nothing', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture(join('parity', 'numbering-stories.docx')));
  const emulator = await pkg.getNumberingEmulator();
  const body = numbered((await pkg.getMainDocumentPart().getContents()).body.content);
  const state = new NumberingState();
  assert.equal(Emulator.peek(pkg, body[0].pPr, state).numString, '1.');
  assert.equal(Emulator.peek(pkg, body[0].pPr, state).numString, '1.');
  assert.equal(Emulator.getNumber(pkg, body[0].pPr, state).numString, '1.');
  assert.equal(Emulator.peek(pkg, body[1].pPr, state).numString, '2.');
  assert.equal(Emulator.getNumber(pkg, body[1].pPr, state).numString, '2.');
  // against the part's default state too
  assert.equal(Emulator.peek(pkg, body[0].pPr).numString, '1.');
  assert.equal(Emulator.peek(pkg, body[0].pPr).numString, '1.');
  assert.equal(Emulator.getNumber(pkg, body[0].pPr).numString, '1.');
  assert.equal(Emulator.peek(pkg, body[0].pPr).numString, '2.');
});

test('NumberingStoriesTest: reset, and getEmulator(true), start every list again', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture(join('parity', 'numbering-stories.docx')));
  const emulator = await pkg.getNumberingEmulator();
  const body = numbered((await pkg.getMainDocumentPart().getContents()).body.content);
  const state = new NumberingState();
  assert.equal(labels(emulator, body, state), '1. 2. 3. 4. 5. 6.');
  state.reset();
  assert.equal(labels(emulator, body, state), '1. 2. 3. 4. 5. 6.');
  // getEmulator(true) does the same for the part's default state
  assert.equal(labels(emulator, body.slice(0, 3), undefined), '1. 2. 3.');
  const again = pkg.numberingDefinitionsPart.getEmulator(true);
  assert.equal(labels(again, body.slice(0, 3), undefined), '1. 2. 3.');
});

// ---------------------------------------------------------------- a level linked to a style

/** docx4j `StyleLinkedLevelTest`'s numbering, verbatim. */
const LINKED_NUMBERING =
  '<w:abstractNum w:abstractNumId="20"><w:multiLevelType w:val="hybridMultilevel"/>'
  + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
  + '<w:pStyle w:val="NumLinked"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>'
  + '<w:pPr><w:ind w:left="567" w:hanging="567"/></w:pPr></w:lvl>'
  + '</w:abstractNum>'
  + '<w:abstractNum w:abstractNumId="21"><w:multiLevelType w:val="hybridMultilevel"/>'
  + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
  + '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>'
  + '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
  + '</w:abstractNum>'
  // level 0 linked to NumLinkedInd0 - whose own w:ind is left 0, firstLine 0 - but stating an
  // indent of its own, which is the level's indent
  + '<w:abstractNum w:abstractNumId="22"><w:multiLevelType w:val="hybridMultilevel"/>'
  + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
  + '<w:pStyle w:val="NumLinkedInd0"/><w:lvlText w:val="›"/><w:lvlJc w:val="left"/>'
  + '<w:pPr><w:ind w:left="198" w:hanging="198"/></w:pPr></w:lvl>'
  + '</w:abstractNum>'
  // level 0 linked to NumLinkedInd0 and stating no indent of its own
  + '<w:abstractNum w:abstractNumId="23"><w:multiLevelType w:val="hybridMultilevel"/>'
  + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
  + '<w:pStyle w:val="NumLinkedInd0"/><w:lvlText w:val="›"/><w:lvlJc w:val="left"/>'
  + '</w:lvl>'
  + '</w:abstractNum>'
  + '<w:num w:numId="20"><w:abstractNumId w:val="20"/></w:num>'
  + '<w:num w:numId="21"><w:abstractNumId w:val="21"/></w:num>'
  + '<w:num w:numId="22"><w:abstractNumId w:val="22"/></w:num>'
  + '<w:num w:numId="23"><w:abstractNumId w:val="23"/></w:num>';

const numberedStyle = (id, ind) =>
  `<w:style w:type="paragraph" w:styleId="${id}">`
  + `<w:name w:val="${id}"/><w:basedOn w:val="Normal"/>`
  + '<w:pPr><w:numPr><w:numId w:val="20"/></w:numPr>'
  + `${ind ?? ''}</w:pPr></w:style>`;

async function linkedPackage() {
  return packageWith(LINKED_NUMBERING, [
    numberedStyle('NumLinked'),
    numberedStyle('NumLinkedInd0', '<w:ind w:left="0" w:firstLine="0"/>'),
    '<w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/><w:basedOn w:val="Normal"/></w:style>',
  ]);
}

const numberVia = (pkg, styleId, direct) =>
  Emulator.getNumber(pkg, styleId, '20', undefined, direct)?.numString;

test('StyleLinkedLevelTest: the linked style is numbered', async () => {
  const pkg = await linkedPackage();
  assert.equal(numberVia(pkg, 'NumLinked', false), '1.');
  assert.equal(numberVia(pkg, 'NumLinked', false), '2.');
});

test('StyleLinkedLevelTest: another style carrying the same w:numPr is not numbered or counted', async () => {
  const pkg = await linkedPackage();
  assert.equal(numberVia(pkg, 'NumLinked', false), '1.');
  assert.equal(numberVia(pkg, 'NumLinkedInd0', false), undefined, 'the level belongs to NumLinked');
  assert.equal(numberVia(pkg, 'NumLinked', false), '2.', 'and the paragraph it skipped was not counted');
});

test('StyleLinkedLevelTest: direct formatting is always numbered', async () => {
  const pkg = await linkedPackage();
  assert.equal(numberVia(pkg, 'Plain', true), '1.');
});

test('StyleLinkedLevelTest: an unlinked level is numbered through any style', async () => {
  const pkg = await linkedPackage();
  const result = Emulator.getNumber(pkg, 'Plain', '21', undefined, false);
  assert.ok(result, "numId 21's level names no style");
  assert.equal(result.numString, '1.');
});

test('StyleLinkedLevelTest: w:numId 0 drops the level indent with the label', async () => {
  const pkg = await linkedPackage();
  const resolver = pkg.propertyResolver;
  const off = resolver.getEffectivePPr(await pPrOf(
    '<w:pPr><w:pStyle w:val="NumLinked"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>'));
  if (off?.ind !== undefined) {
    assert.equal(off.ind.left, undefined, "the level's 567 twips must go with its label");
    assert.equal(off.ind.hanging, undefined);
  }
  // the control: without the w:numId 0 the level's indent stands
  const still = resolver.getEffectivePPr(await pPrOf('<w:pPr><w:pStyle w:val="NumLinked"/></w:pPr>')).ind;
  assert.ok(still, "the level's indent applies to a numbered paragraph");
  assert.equal(still.left, 567);
});

test('StyleLinkedLevelTest: a linked level stating an indent keeps its own', async () => {
  const pkg = await linkedPackage();
  const own = pkg.numberingDefinitionsPart.getIndOf('22', '0');
  assert.ok(own);
  assert.equal(own.left, 198, "the level's own w:ind, not the linked style's 0");
  assert.equal(own.hanging, 198);
  assert.equal(own.firstLine, undefined, "the linked style's w:firstLine must not leak in");

  const effective = pkg.propertyResolver.getEffectivePPr(
    await pPrOf('<w:pPr><w:numPr><w:numId w:val="22"/></w:numPr></w:pPr>')).ind;
  assert.ok(effective);
  assert.equal(effective.left, 198);
  assert.equal(effective.hanging, 198);
});

test('StyleLinkedLevelTest: a linked level stating no indent falls back to the style', async () => {
  const pkg = await linkedPackage();
  const ind = pkg.numberingDefinitionsPart.getIndOf('23', '0');
  assert.ok(ind, "the linked style's w:ind is the fallback");
  assert.equal(ind.left, 0);
  assert.equal(ind.firstLine, 0);
});

// ---------------------------------------------------------------- the default paragraph style

/** docx4j `DefaultStyleNumberedTest`: Normal carries a w:numPr for the default numbering's list 1. */
async function packageWithNumberedNormal() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const ndp = new NumberingDefinitionsPart();
  pkg.getMainDocumentPart().addTargetPart(ndp);
  await ndp.unmarshalDefaultNumbering();   // numId 1: decimal
  const styles = await pkg.getMainDocumentPart().styleDefinitionsPart.getContents();
  const normal = styles.style.find((s) => s._default === true && s.type === 'paragraph');
  normal.pPr ??= { TYPE_NAME: 'org_docx4j_wml.PPr' };
  normal.pPr.numPr = { TYPE_NAME: 'org_docx4j_wml.PPrBase.NumPr', numId: { TYPE_NAME: 'org_docx4j_wml.PPrBase.NumPr.NumId', val: 1 } };
  // a style based on nothing, so it inherits no numbering
  styles.style.push(...await unmarshalStyles(['<w:style w:type="paragraph" w:styleId="Unlinked"/>']));
  await pkg.refresh();
  return pkg;
}

test('DefaultStyleNumberedTest: paragraphs naming no style are numbered (CR-014 probe P6)', async () => {
  const pkg = await packageWithNumberedNormal();
  const noStyle = { TYPE_NAME: 'org_docx4j_wml.PPr' };
  assert.equal(Emulator.getNumber(pkg, noStyle).numString, '1.');
  assert.equal(Emulator.getNumber(pkg, noStyle).numString, '2.');
  assert.equal(Emulator.getNumber(pkg, noStyle).numString, '3.');
});

test('DefaultStyleNumberedTest: a style based on nothing is not', async () => {
  const pkg = await packageWithNumberedNormal();
  const pPr = await pPrOf('<w:pPr><w:pStyle w:val="Unlinked"/></w:pPr>');
  assert.equal(Emulator.getNumber(pkg, pPr), undefined);
  assert.equal(Emulator.getInd(pkg, 'Unlinked', undefined, undefined), undefined);
});

test('DefaultStyleNumberedTest: resolve says where the numbering came from', async () => {
  const pkg = await packageWithNumberedNormal();
  const emulator = pkg.numberingDefinitionsPart.getEmulator();

  let ref = emulator.resolve(undefined, undefined, undefined, false);
  assert.equal(ref.notNumbered, false);
  assert.equal(ref.numId, '1');
  assert.equal(ref.ilvl, '0');
  assert.equal(ref.direct, false);

  ref = emulator.resolve('Unlinked', undefined, undefined, false);
  assert.equal(ref.notNumbered, true);

  ref = emulator.resolve('Unlinked', '1', '2', true);
  assert.equal(ref.notNumbered, false);
  assert.equal(ref.direct, true);
  assert.equal(ref.ilvl, '2');
});

test('unmarshalDefaultNumbering gives docx4j\'s bullet and decimal sets', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const ndp = new NumberingDefinitionsPart();
  pkg.getMainDocumentPart().addTargetPart(ndp);
  const numbering = await ndp.unmarshalDefaultNumbering();
  assert.equal(numbering.abstractNum.length, 2);
  assert.equal(numbering.num.length, 2);
  await pkg.refresh();

  const definitions = ndp.definitions;
  // w:num 1 is the decimal set, w:num 2 the bullet one
  const decimal = definitions.list('1');
  assert.equal(decimal.level('0').numFmt, 'decimal');
  assert.equal(decimal.level('0').lvlText, '%1.');
  assert.equal(decimal.isBullet('0'), false);
  assert.equal(decimal.levels.size, 9, 'nine levels');

  const bullet = definitions.list('2');
  assert.equal(bullet.level('0').numFmt, 'bullet');
  assert.equal(bullet.isBullet('0'), true);
  assert.equal(bullet.level('0').font, 'Symbol');

  assert.equal(Emulator.getNumber(pkg, undefined, '1', '0').numString, '1.');
  assert.equal(Emulator.getNumber(pkg, undefined, '1', '0').numString, '2.');
  const first = Emulator.getNumber(pkg, undefined, '2', '0');
  assert.equal(first.isBullet, true);
  assert.equal(first.numFont, 'Symbol');
  assert.equal(first.bullet, first.numString);
});

// ---------------------------------------------------------------- numRefFor

test('numRefFor: a package with no numbering part, a missing pPr, and w:numId 0', async () => {
  const bare = await WordprocessingMLPackage.createPackage();
  assert.equal(Emulator.numRefFor(bare, {}).reason, 'no numbering part');
  assert.equal(Emulator.numRefFor(bare, undefined).reason, 'no pPr');

  const pkg = await linkedPackage();
  assert.equal(Emulator.numRefFor(pkg, undefined).reason, 'no pPr');

  // the paragraph's own w:numId 0 (docx4j 01d661547; ECMA-376 17.9.18)
  const direct = Emulator.numRefFor(pkg, await pPrOf(
    '<w:pPr><w:pStyle w:val="NumLinked"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>'));
  assert.equal(direct.notNumbered, true);
  assert.equal(direct.reason, "the paragraph's w:numId 0 turns numbering off");
  assert.equal(direct.numId, undefined);
  assert.equal(Emulator.getNumber(pkg, await pPrOf(
    '<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>')), undefined);

  // and one a style contributes
  const off = await packageWith(LINKED_NUMBERING, [
    '<w:style w:type="paragraph" w:styleId="NoNumbers"><w:name w:val="NoNumbers"/>'
    + '<w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:style>',
  ]);
  const viaStyle = Emulator.numRefFor(off, await pPrOf('<w:pPr><w:pStyle w:val="NoNumbers"/></w:pPr>'));
  assert.equal(viaStyle.notNumbered, true);
  assert.equal(viaStyle.reason, "style 'NoNumbers's w:numId 0 turns numbering off");

  // a resolved reference names the list, the level and where the numId came from
  const resolved = Emulator.numRefFor(pkg, await pPrOf('<w:pPr><w:pStyle w:val="NumLinked"/></w:pPr>'));
  assert.equal(resolved.notNumbered, false);
  assert.equal(resolved.numId, '20');
  assert.equal(resolved.ilvl, '0');
  assert.equal(resolved.direct, false);
  assert.equal(String(resolved), 'numId 20 ilvl 0 (from style)');
});

/**
 * docx4j `ParityAccessorsTest`'s three "and says why" cases, with its strings verbatim
 * (docx4j `VERSION_17_1_1` at `7fba7a150`): a `w:numId` 0, a `w:numId` naming no `w:num`, and
 * a level the `w:num`'s definition does not have.  All three are `notNumbered` with a reason
 * naming where the `w:numPr` came from, and `getNumber` answers nothing.  No golden reaches
 * the last two - a `w:numId` left dangling by Word's own re-save of an untouched
 * `mc:Fallback` (docx4j CR-021 §8.5) is exactly the case fixtures do not hold - so these are
 * the only guard on those strings.
 */
test('ParityAccessorsTest: numId 0, a dangling numId and a missing level say why', async () => {
  const pkg = await packageWithNumberedNormal();   // Normal carries w:numId 1, nine levels

  // numIdZeroTurnsNumberingOff: ECMA-376 17.9.18, and Normal's numbering does not come back
  const zero = await pPrOf('<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>');
  const zeroRef = Emulator.numRefFor(pkg, zero);
  assert.equal(zeroRef.notNumbered, true);
  assert.equal(zeroRef.reason, "the paragraph's w:numId 0 turns numbering off");
  assert.equal(Emulator.getNumber(pkg, zero), undefined);

  // danglingNumIdIsNotNumberedAndSaysWhy
  const dangling = await pPrOf('<w:pPr><w:numPr><w:numId w:val="99"/></w:numPr></w:pPr>');
  const danglingRef = Emulator.numRefFor(pkg, dangling);
  assert.equal(danglingRef.notNumbered, true);
  assert.equal(danglingRef.reason, "no w:num for numId 99 (the paragraph's own)");
  assert.equal(danglingRef.numId, undefined);
  assert.equal(Emulator.getNumber(pkg, dangling), undefined);

  // missingLevelIsNotNumberedAndSaysWhy
  const deep = await pPrOf('<w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="42"/></w:numPr></w:pPr>');
  const deepRef = Emulator.numRefFor(pkg, deep);
  assert.equal(deepRef.notNumbered, true);
  assert.equal(deepRef.reason, "no w:lvl 42 in w:num 1 (the paragraph's own)");
  assert.equal(Emulator.getNumber(pkg, deep), undefined);

  // nothing was counted on the way: all three answer before any counter is touched
  assert.equal(pkg.numberingDefinitionsPart.numberingState.isEmpty, true);
});

/** The same two, with the `w:numPr` coming from the paragraph's style: "(from style 'X')". */
test('ParityAccessorsTest: a dangling numId and a missing level from a style name it', async () => {
  const styled = await packageWith(
    '<w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
    + '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
    + '<w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num>',
    [
      '<w:style w:type="paragraph" w:styleId="Gone"><w:name w:val="Gone"/><w:basedOn w:val="Normal"/>'
      + '<w:pPr><w:numPr><w:numId w:val="99"/></w:numPr></w:pPr></w:style>',
      '<w:style w:type="paragraph" w:styleId="TooDeep"><w:name w:val="TooDeep"/><w:basedOn w:val="Normal"/>'
      + '<w:pPr><w:numPr><w:numId w:val="7"/><w:ilvl w:val="3"/></w:numPr></w:pPr></w:style>',
    ]);

  const gone = Emulator.numRefFor(styled, await pPrOf('<w:pPr><w:pStyle w:val="Gone"/></w:pPr>'));
  assert.equal(gone.notNumbered, true);
  assert.equal(gone.reason, "no w:num for numId 99 (from style 'Gone')");

  const tooDeep = Emulator.numRefFor(styled, await pPrOf('<w:pPr><w:pStyle w:val="TooDeep"/></w:pPr>'));
  assert.equal(tooDeep.notNumbered, true);
  assert.equal(tooDeep.reason, "no w:lvl 3 in w:num 7 (from style 'TooDeep')");
});

// ---------------------------------------------------------------- the round trip

test('a document whose numbering is only read is still saved byte for byte', async () => {
  const bytes = await fixture(join('parity', 'numbering-stories.docx'));
  const pkg = await WordprocessingMLPackage.load(bytes);
  const emulator = await pkg.getNumberingEmulator();
  const document = await pkg.getMainDocumentPart().getContents();
  const body = numbered(document.body.content);
  assert.equal(labels(emulator, body, new NumberingState()), '1. 2. 3. 4. 5. 6.');
  assert.equal(pkg.numberingDefinitionsPart.isUnmarshalled, false, 'the numbering part stays unread');

  const saved = await pkg.save();
  const src = new ZipPartStore(bytes);
  const out = new ZipPartStore(saved);
  for (const name of ['word/numbering.xml', 'word/styles.xml']) {
    assert.ok(bytesEqual(await src.load(name), await out.load(name)), `${name} byte-identical`);
  }
});
