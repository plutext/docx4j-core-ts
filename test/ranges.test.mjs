// CR-014, ranges across paragraphs. Phase 0 is check 52 (test/README.md), which also measures CR-002 section
// 55's D2 (Word's table rule, cases 86 to 95): its file and its Script Lab snippet must agree, since Word runs
// the snippet on the file and nothing can be fixed after the run. Every case's BEGIN paragraph parses with the
// snippet's own parser to the case the file was built from, and every token the snippet searches for is in the
// file once, as a whole word in upper case. Word 365's two saves of the run (test/check52/, outside fixtures/
// so that the weekly RevisionsHarness cuts no golden for them) hold this package's Accept All to Word's.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../dist/index.mjs';
import { Word, NotSupportedError } from '../dist/office-js/index.mjs';
import { check52Bytes, check52Cases } from './fixtures/check52/build.mjs';
import { check53Bytes, check53Cases } from './fixtures/check53/build.mjs';
import { check54Bytes, check54Cases } from './fixtures/check54/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));

/** The snippet's CHECK52, loaded as Script Lab runs it, with no page around it. */
async function snippet() {
  const source = await readFile(join(here, 'fixtures', 'check52', 'check52-script-lab.js'), 'utf8');
  return new Function(`${source}\nreturn CHECK52;`)();
}

test("check 52's file and snippet agree: every case parses, every token the snippet looks for is there once", async () => {
  const { parseRegion } = await snippet();
  const pkg = await WordprocessingMLPackage.load(await check52Bytes());
  await pkg.getBody();
  const texts = pkg.body.paragraphs.map((p) => p.text);
  const regions = texts.map(parseRegion).filter((r) => r !== null);
  const cases = check52Cases();
  assert.equal(cases.length, 95);
  assert.deepEqual(regions.map((r) => [r.n, r.kind, r.op, r.tracked]), cases.map((c) => [c.n, c.kind, c.op, c.tracking === 'tracked']));

  const count = (token) => texts.reduce((sum, text) => sum + (text.match(new RegExp(`(?<![A-Za-z0-9])${token}(?![A-Za-z0-9])`, 'g')) ?? []).length, 0);
  const tags = pkg.body.contentControls.map((c) => c.tag);
  for (const r of regions) {
    const tokens = [`BEGIN${r.n}`, `END${r.n}`];
    if (r.range) {
      const { from, to, toMark, whole, control, cell, paragraph } = r.range;
      tokens.push(...[from, to, toMark, whole, cell, paragraph].filter(Boolean));
      if (control) assert.equal(tags.filter((t) => t === control).length, 1, control);
    } else if (r.queries) {
      for (const q of r.queries) tokens.push(...q.split(/\^./).filter((s) => s !== ''));
    } else {
      assert.equal(r.kind, 'VIEWS');
    }
    for (const token of tokens) assert.equal(count(token), 1, `case ${r.n}: ${token}`);
  }
  // every kind of range, every operation and both tracking modes are there
  assert.deepEqual([...new Set(regions.filter((r) => r.range).map((r) => Object.keys(r.range).join('+')))].sort(),
    ['cell', 'control', 'from+to', 'from+toMark', 'paragraph', 'whole']);
  assert.deepEqual([...new Set(regions.map((r) => r.op))].sort(),
    ['bold', 'clear', 'comment', 'control-delete', 'delete', 'ooxml-replace', 'paragraph-after', 'paragraph-before', 'paragraph-delete',
      'read', 'replace', 'search', 'table-end', 'text-end', 'text-start']);
  // CR-002 section 55's cases: the token each names is in a table cell, the control's in one too
  for (const r of regions.filter((x) => x.kind.startsWith('TC'))) {
    const token = r.range.cell ?? r.range.paragraph;
    const paragraph = token ? pkg.body.paragraphs.find((p) => p.text.includes(token)) : pkg.body.contentControls.find((c) => c.tag === r.range.control).paragraphs[0];
    assert.ok(paragraph.parentTableCell, `case ${r.n}`);
  }
});

test("check 52's table for ooxml-replace is a flat OPC package this package reads: a table of one cell, nothing after it", async () => {
  const { TABLE_OOXML } = await snippet();
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  body.insertParagraph('before', 'End');
  const inserted = await body.insertOoxml(TABLE_OOXML, 'End');
  assert.equal(inserted.length, 1);
  assert.deepEqual(body.tables.map((t) => t.values), [[['NEWOOXML']]]);
});

test("check 52's run and its report after Accept All (Word 365, 2026-10-22) are of this file: the same 95 cases, each its kind, operation and tracking", async () => {
  for (const [name, mode] of [['52-ranges-word365-run.json', 'run'], ['52-ranges-word365-accepted.json', 'report']]) {
    const word = JSON.parse(await readFile(join(here, 'fixtures', 'check52', name), 'utf8'));
    assert.equal(word.mode, mode);
    assert.deepEqual(word.cases.map((c) => [c.n, c.kind, c.op, c.tracked]),
      check52Cases().map((c) => [c.n, c.kind, c.op, c.tracking === 'tracked']));
  }
});

const saves = join(here, 'check52');
const saveBytes = (name) => readFile(join(saves, name));

test("check 52's Word saves are Word's, as copied: both files listed, each SHA-1 matching", async () => {
  const manifest = JSON.parse(await readFile(join(saves, 'manifest.json'), 'utf8'));
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await saveBytes(name)).digest('hex'), sha1, name);
  assert.deepEqual(manifest.files.map((f) => f.name), ['52-ranges-word365.docx', '52-ranges-word365-accepted.docx']);
});

/** Each case's paragraphs between its BEGIN and END paragraphs: text, style, alignment, list, table nesting. */
async function regionsOf(pkg) {
  await pkg.getBody();
  const nesting = (p) => {
    let depth = 0;
    for (let cell = p.parentTableCell; cell; cell = cell.parentTable.parentTableCell) depth++;
    return depth;
  };
  const regions = new Map();
  let n = null;
  for (const p of pkg.body.paragraphs) {
    const begin = /^BEGIN(\d+) /.exec(p.text);
    if (begin) { n = Number(begin[1]); regions.set(n, []); continue; }
    if (n !== null && p.text === `END${n}`) { n = null; continue; }
    if (n !== null) regions.get(n).push({ text: p.text, style: p.styleBuiltIn, alignment: p.alignment, listItem: p.isListItem, nesting: nesting(p) });
  }
  return regions;
}

test("check 52: acceptAllRevisions on Word 365's save after the run gives Word's Accept All, region for region (CR-014 section 6.2)", async () => {
  const ours = await WordprocessingMLPackage.load(await saveBytes('52-ranges-word365.docx'));
  assert.ok((await ours.getAllTrackedChanges()).length > 0);
  await ours.acceptAllRevisions();
  assert.equal((await ours.getAllTrackedChanges()).length, 0);
  const mine = await regionsOf(ours);
  const word = await regionsOf(await WordprocessingMLPackage.load(await saveBytes('52-ranges-word365-accepted.docx')));
  assert.equal(word.size, 95);
  assert.deepEqual([...mine.keys()], [...word.keys()]);
  const differing = [...word.keys()].filter((n) => JSON.stringify(mine.get(n)) !== JSON.stringify(word.get(n)));
  // case 26 (CC, its content deleted under tracking): Word fills the emptied control with five spaces marked as its
  // placeholder, which reads as no text, as this package's empty control does (CR-012 section 18)
  assert.deepEqual(differing, []);
  // the joins: the joined paragraph in the properties of the paragraph the range ended in (HB, BH, LB, BL, TBL)
  assert.deepEqual([16, 17, 18, 19, 21].map((n) => word.get(n).map((p) => [p.text, p.style, p.alignment, p.listItem])), [
    [['HEAD16A  BODY16B', 'Normal', 'Centered', false]],
    [['BODY17A  HEAD17B', 'Heading1', 'Left', false]],
    [['LIST18A  BODY18B', 'Normal', 'Centered', false]],
    [['BODY19A  LIST19B', 'ListParagraph', 'Left', true]],
    [['HEAD21A  BODY21B', 'Normal', 'Centered', false]],
  ]);
});

// ------------------------------------------------------------------------------------------- phase 1, the reads

/** A package of check 52's file, or a Word save, with its font selector built, so that `font.name` reads the face. */
async function loaded(bytes) {
  const pkg = await WordprocessingMLPackage.load(bytes);
  await pkg.getBody();
  await pkg.getMainDocumentPart().getRunFontSelector();
  return pkg;
}

test("CR-014 phase 1: check 52's reads, run through the Word shim on the same file, are Word 365's, kind for kind", async () => {
  const { runCheck } = await snippet();
  const pkg = await loaded(await check52Bytes());
  const ours = await runCheck((f) => Word.run(pkg, f), 'reads');
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check52', '52-ranges-word365-run.json'), 'utf8'));
  assert.deepEqual(Object.keys(ours.reads), Object.keys(word.reads));
  for (const kind of Object.keys(word.reads)) {
    // `marks`, the hits of ^p within the range, since phase 3 too: every paragraph mark outside a table the range
    // holds, never a cell's end (CELL and the TC kinds find none)
    const mine = ours.reads[kind];
    const theirs = word.reads[kind];
    // Word gives the style 'Other', named '', to a paragraph after a nested table in a cell whose style isn't
    // Normal (TC2; check 53, cases 10 and 12), which is not followed. FN's 'Other', a range whose runs' character
    // styles differ, is Word's rule and followed (check 53, cases 5 and 8)
    if (kind === 'TC2') {
      assert.deepEqual([theirs.range.style, theirs.range.styleBuiltIn], ['', 'Other'], kind);
      assert.deepEqual([mine.range.style, mine.range.styleBuiltIn], ['List Paragraph', 'ListParagraph'], kind);
      mine.range = { ...mine.range, style: '', styleBuiltIn: 'Other' };
    }
    assert.deepEqual(mine, theirs, kind);
  }
});

test("CR-014 phase 1: check 52's report, run through the Word shim on Word's save after Accept All, is Word's, region for region", async () => {
  const { runCheck } = await snippet();
  const pkg = await loaded(await saveBytes('52-ranges-word365-accepted.docx'));
  const ours = await runCheck((f) => Word.run(pkg, f), 'report');
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check52', '52-ranges-word365-accepted.json'), 'utf8'));
  assert.equal(ours.cases.length, word.cases.length);
  const bold = new Set(word.cases.filter((c) => c.op === 'bold').map((c) => c.n));
  // getTextRanges, which the bold cases' report reads each word with (`words`), is not implemented; each paragraph's font is Word's (CR-002 section 59)
  const paragraphOf = ({ words, ...rest }) => rest;
  for (const [i, theirs] of word.cases.entries()) {
    const n = theirs.n;
    const mine = ours.cases[i].region;
    const expected = structuredClone(theirs.region);
    // Range.footnotes is not implemented
    assert.match(mine.footnotes, /^error: NotImplemented/, `${n}`);
    delete expected.footnotes;
    // getTextRanges, which the bold cases' report reads each word with, is not implemented
    if (bold.has(n)) {
      assert.match(mine.paragraphs, /^error: NotImplemented/, `${n}`);
      delete expected.paragraphs;
    } else {
      expected.paragraphs = expected.paragraphs.map(paragraphOf);
      mine.paragraphs = mine.paragraphs.map(paragraphOf);
    }
    if (expected.paragraphs) expected.paragraphs = listParagraphForOther(n, expected.paragraphs);
    // cases 13 and 26, a control emptied: Word marks it w:showingPlcHdr over a run of five spaces, which reads as
    // no text, here as in Word since CR-012 section 18 (CR-014 section 9.4)
    delete mine.footnotes;
    if (bold.has(n)) delete mine.paragraphs;
    assert.deepEqual(mine, expected, `case ${n} (${theirs.kind} ${theirs.op}${theirs.tracked ? ' tracked' : ''})`);
  }
});

test("CR-014 phase 1: the tracked changes a range across paragraphs lists, on Word's save after the run, are Word's but for the listing's own grouping", async () => {
  const { runCheck } = await snippet();
  const pkg = await loaded(await saveBytes('52-ranges-word365.docx'));
  const ours = await runCheck((f) => Word.run(pkg, f), 'report');
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check52', '52-ranges-word365-run.json'), 'utf8'));
  // a deletion's text and range are All Markup's in Word's run (deleted text shown), so a deletion is compared by its type
  const view = (changes) => changes.map((c) => [c.type, c.type === 'Deleted' ? '' : c.range]);
  const replaced = new Set([39, 40, 41, 42, 43, 46, 47, 48]);
  const compared = [];
  for (const [i, theirs] of word.cases.entries()) {
    const expected = theirs.region.trackedChanges;
    if (!Array.isArray(expected) || expected.length === 0) continue;
    const mine = view(ours.cases[i].region.trackedChanges);
    if (replaced.has(theirs.n)) {
      // a tracked Replace: Office JS lists the insertion alone, where the markup holds the deletion too, which is listed here
      assert.deepEqual(mine, [...view(expected), ['Deleted', '']], `case ${theirs.n}`);
    } else if (theirs.n === 54 || theirs.n === 55) {
      // Word lists formatting across a table's cells as one change; this package's listing (CR-002) a run at a time
      assert.equal(expected.length, 1, `case ${theirs.n}`);
      assert.equal(mine.length, 5, `case ${theirs.n}`);
    } else if (theirs.n === 87) {
      // a tracked table insertion at a cell's end: Word's one change is the table and the mark before it; this
      // package lists the mark alone, its range the mark (check 53), and not the nested table's rows
      assert.deepEqual(view(expected), [['Added', '\\u000dNEWCELL\\u000d\\u000a']]);
      assert.deepEqual(mine, [['Added', '\\u000d']]);
    } else {
      assert.deepEqual(mine, view(expected), `case ${theirs.n}`);
      compared.push(theirs.n);
    }
  }
  // case 1's deletion and insertion through a paragraph's mark, each one change spanning two paragraphs, among them
  assert.ok(compared.includes(1) && compared.length >= 19, compared.join(' '));
});

/**
 * Word reads a cell's paragraph after a nested table as style '' ('Other') in cases 89, 91 and 94, where its save has
 * it List Paragraph, as this package reads it: Word's reading taken as its markup says.
 */
function listParagraphForOther(n, paragraphs) {
  return paragraphs.map((p) => {
    if (p.styleBuiltIn !== 'Other') return p;
    assert.ok([89, 91, 94].includes(n), `${n}`);
    return { ...p, style: 'List Paragraph', styleBuiltIn: 'ListParagraph' };
  });
}

/** check52Bytes' paragraph holding a word, and a range over that word. */
const wordIn = (pkg, word) => pkg.body.search(word, { matchCase: true, matchWholeWord: true })[0];

test('CR-014 phase 1: a range across paragraphs, from expandTo, reads every paragraph; the writes phase 2 does not bring refuse', async () => {
  const pkg = await loaded(await check52Bytes());
  const range = wordIn(pkg, 'HEAD3B').expandTo(wordIn(pkg, 'BODY3A'));
  assert.equal(range.spansParagraphs, true);
  assert.equal(range.text, 'HEAD3B\rBODY3A');
  assert.deepEqual(range.paragraphs.map((p) => p.text), ['HEAD3A HEAD3B', 'BODY3A BODY3B']);
  assert.equal(range.endParagraph.text, 'BODY3A BODY3B');
  assert.deepEqual(['Whole', 'Content', 'Start', 'End', 'After'].map((l) => range.getRange(l).text), ['HEAD3B\rBODY3A', 'HEAD3B\rBODY3A', '', '', '']);
  assert.equal(range.getRange('End').paragraph.text, 'BODY3A BODY3B');
  // the font reads null where the runs differ: the heading's are bold, the body paragraph's italic
  assert.deepEqual([range.font.bold, range.font.italic, range.font.underline], [null, null, 'None']);
  // search within the span, a paragraph at a time
  assert.deepEqual(range.search('3').map((r) => r.text), ['3', '3']);
  assert.deepEqual(range.search('HEAD3A'), []);
  // refused before anything changes: a hyperlink across paragraphs (Word writes a field across them, CR-014
  // section 12), and a control on a range not holding its paragraphs whole, as Word refused it (check 54)
  const before = pkg.body.text;
  assert.throws(() => { range.hyperlink = 'https://example.com'; }, NotSupportedError);
  assert.throws(() => range.insertContentControl(), (e) => e.code === 'GeneralException');
  assert.equal(pkg.body.text, before);
  // the shim's getOoxml gives its two paragraphs' pieces (toApiScript writes the same blocks)
  const ooxml = await Word.run(pkg, async (context) => {
    const hit = context.document.body.search('HEAD3B', { matchCase: true })[0];
    const x = hit.expandTo(context.document.body.search('BODY3A', { matchCase: true })[0]).getOoxml();
    await context.sync();
    return x.value;
  });
  assert.deepEqual([...ooxml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]), ['HEAD3B', 'BODY3A']);
  assert.equal((ooxml.match(/<w:p\b/g) ?? []).length, 2);
  // expandTo from a paragraph's whole range takes its mark; ranges in different stories are refused
  const toMark = wordIn(pkg, 'HEAD15B').expandTo(wordIn(pkg, 'BODY15A').paragraphs[0].getRange('Whole'));
  assert.equal(toMark.text, 'HEAD15B\rBODY15A BODY15B\r');
  assert.equal(toMark.getRange('Content').text, 'HEAD15B\rBODY15A BODY15B');
  assert.equal(toMark.getRange('After').paragraph.text, 'NEXT15A NEXT15B');
  const note = await pkg.paragraphAt('footnote:1/0');
  assert.match(note.text, /NOTE12/);
  assert.throws(() => range.expandTo(note.getRange()), (e) => e.code === 'ItemNotFound' && /different stories/.test(e.message));
});

test('CR-014 phase 1: the views across paragraphs - a block control, a comment (D2), a tracked change, a body and a table - give one Range', async () => {
  const pkg = await loaded(await check52Bytes());
  const body = pkg.body;
  // a block control of three paragraphs: one Range, no longer a SpanningRange
  const control = body.contentControls.find((c) => c.tag === 'c52-13');
  assert.equal(control.getRange().text, 'HEAD13A HEAD13B\rBODY13A BODY13B\rLIST13A LIST13B\r');
  assert.equal(control.getRange('Content').text, 'HEAD13A HEAD13B\rBODY13A BODY13B\rLIST13A LIST13B');
  assert.equal(control.getRange().paragraphs.length, 3);
  // the comment from the body paragraph into the list item (case 1): one Range as Word's, getRanges() a paragraph at a time
  const [comment] = await body.getComments();
  assert.equal(comment.getRange().text, 'BODY1B\rLIST1A');
  assert.deepEqual(comment.getRanges().map((r) => r.text), ['BODY1B', 'LIST1A']);
  // the tracked insertion through the list item's mark: its range spans both paragraphs, as Word's (case 1)
  const inserted = body.getTrackedChanges().find((c) => c.type === 'Added');
  assert.equal(inserted.getRange().text, ' LIST1B\rPLAIN1A ');
  // a table's whole range ends each row \r\n (case 1), a cell's body's whole range its last paragraph's mark, \t
  const table = body.tables[0];
  assert.equal(table.getRange().text, 'T1R1C1A T1R1C1B\tT1R1C2A T1R1C2B\r\nT1R2C1A T1R2C1B\tT1R2C2A T1R2C2B\r\n');
  assert.equal(table.getRange('After').paragraph.text, 'END1');
  assert.equal(table.getCell(0, 0).body.getRange().text, 'T1R1C1A T1R1C1B\t');
  // the body's whole range is its text and the last mark
  assert.equal(body.getRange().text, `${body.text}\r`);
  assert.equal(body.getRange('Content').text, body.text);
  // the tables a range holds: one it crosses, none for a range inside a cell
  const across = wordIn(pkg, 'HEAD8B').expandTo(wordIn(pkg, 'BODY8A'));
  assert.deepEqual(across.tables.map((t) => t.rowCount), [2]);
  assert.deepEqual(table.getCell(0, 0).body.getRange().tables, []);
  assert.deepEqual(across.paragraphs.map((p) => p.tableNestingLevel), [0, 1, 1, 1, 1, 0]);
});

// ------------------------------------------------------------------------------- check 53 (phase 1's choices)

test("check 53's file and snippet agree: every case parses to its op, every word an op names is in the file once", async () => {
  const source = await readFile(join(here, 'fixtures', 'check53', 'check53-script-lab.js'), 'utf8');
  const { parseRegion } = new Function(`${source}\nreturn CHECK53;`)();
  const pkg = await WordprocessingMLPackage.load(await check53Bytes());
  await pkg.getBody();
  const texts = pkg.body.paragraphs.map((p) => p.text);
  const regions = texts.map(parseRegion).filter((r) => r !== null);
  assert.deepEqual(regions.map((r) => [r.n, `${r.op}${r.args.length ? ' ' : ''}${r.args.join(' ').toLowerCase()}`]),
    check53Cases().map((c) => [c.n, c.spec]));
  const count = (token) => texts.reduce((sum, text) => sum + (text.match(new RegExp(`(?<![A-Za-z0-9])${token}(?![A-Za-z0-9])`, 'g')) ?? []).length, 0);
  const tags = pkg.body.contentControls.map((c) => c.tag);
  for (const r of regions) {
    for (const token of [`BEGIN${r.n}`, `END${r.n}`]) assert.equal(count(token), 1, `case ${r.n}: ${token}`);
    if (r.op === 'control' || r.op === 'placeholder') assert.equal(tags.filter((t) => t === r.args[0]).length, 1, `case ${r.n}`);
    else for (const token of r.args) assert.equal(count(token), 1, `case ${r.n}: ${token}`);
  }
});

test("check 53 (Word 365, Simple Markup): the snippet through the Word shim reads as Word does, but for the differences listed, each outside CR-014 phase 1", async () => {
  const source = await readFile(join(here, 'fixtures', 'check53', 'check53-script-lab.js'), 'utf8');
  const { runCheck } = new Function(`${source}\nreturn CHECK53;`)();
  const pkg = await loaded(await check53Bytes());
  const ours = await runCheck((f) => Word.run(pkg, f), 'core-ts');
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check53', '53-reads-word365-simple.json'), 'utf8'));
  const all = JSON.parse(await readFile(join(here, 'fixtures', 'check53', '53-reads-word365-all.json'), 'utf8'));
  const flat = (o, p = '', out = {}) => {
    if (o && typeof o === 'object') for (const [k, v] of Object.entries(o)) flat(v, p ? `${p}.${k}` : k, out);
    else out[p] = o;
    return out;
  };
  const differing = {};
  for (const [i, theirs] of word.cases.entries()) {
    const [a, b] = [flat(ours.cases[i].reads), flat(theirs.reads)];
    const keys = [...new Set([...Object.keys(a), ...Object.keys(b)])].filter((k) => JSON.stringify(a[k]) !== JSON.stringify(b[k]));
    if (keys.length > 0) differing[theirs.n] = keys;
  }
  const changeRead = (n) => [`paragraphs.${n}.changes.0.type`, `paragraphs.${n}.changes.0.text`, `paragraphs.${n}.changes.0.range`, `paragraphs.${n}.changes.0.rangeParagraphs`,
    `paragraphs.${n}.wholeChanges.0.type`, `paragraphs.${n}.wholeChanges.0.text`, `paragraphs.${n}.wholeChanges.0.range`, `paragraphs.${n}.wholeChanges.0.rangeParagraphs`];
  // a control showing its placeholder reads as empty, as Word's (CR-012 section 18); its placeholderText reads the runs
  // here, where Word, with no glossary part, reads null; Word's showingPlaceholderText came back undefined
  const placeholder = ['control.placeholderText', 'showingPlaceholderText'];
  const styleOther = ['paragraph.style', 'paragraph.styleBuiltIn', 'whole.style', 'whole.styleBuiltIn'];
  assert.deepEqual(differing, {
    // Heading 1's theme colour (accent1, shaded) read through the theme, which is CR-008 phase 2's: #0F4761 in Word, the
    // w:val here. Every other colour, automatic ones #000000, and every font read where runs differ are Word's (CR-002 section 59)
    3: ['font.color'], 4: ['font.color'], 5: ['font.color'],
    // a paragraph after a nested table in a cell, whose style isn't Normal: Word's 'Other', not followed
    10: styleOther, 12: styleOther,
    // tracked changes in Simple Markup. A change's text here is docx4j's, as All Markup reads it, where Simple Markup
    // reads a deletion's as ''. Word joins the paragraphs a deleted mark ends, so its collapsed deletion touches two
    // paragraphs; and its paragraph listing gives each paragraph its part of a change (25), and a deleted
    // paragraph none (30). The listing and the readers are CR-002's and CR-012's
    25: ['region.0.text', 'region.0.rangeParagraphs', 'paragraphs.1.changes.0.text', 'paragraphs.1.wholeChanges.0.text', ...changeRead(2)],
    28: ['region.0.text', 'paragraphs.1.changes.0.text', 'paragraphs.1.wholeChanges.0.text'],
    30: ['region.0.text', ...changeRead(2)],
    // Range.footnotes is not implemented, so the snippet can't reach expandTo's ItemNotFound through the shim
    32: [''],
    34: placeholder, 35: placeholder,
  });
  // the colour across two paragraphs whose runs' colours differ: '' (MixedFont), as Word's
  assert.equal(flat(ours.cases[0].reads)['font.color'], flat(word.cases[0].reads)['font.color']);
  // the changes' text is All Markup's
  for (const n of [25, 28, 30]) assert.equal(ours.cases[n - 1].reads.region[0].text, all.cases[n - 1].reads.region[0].text, `case ${n}`);
});

// ------------------------------------------------------------------------------------------- phase 2, the writes

/** Check 52's snippet in its 'run' mode, through the Word shim on its own file, once for the tests below. */
let phase2;
const phase2Run = () => (phase2 ??= (async () => {
  const { runCheck } = await snippet();
  const pkg = await loaded(await check52Bytes());
  const result = await runCheck((f) => Word.run(pkg, f), 'run');
  return { result, saved: await pkg.save() };
})());

test("CR-014 phase 2: check 52's writes, run through the Word shim on the same file, leave each region as Word 365's run did", async () => {
  const { result: ours } = await phase2Run();
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check52', '52-ranges-word365-run.json'), 'utf8'));
  // under tracking Word's text is All Markup's, deleted text shown, and so are its font reads, which take the deleted
  // runs in; so a tracked case's paragraphs are compared without their text and font, and its text, tables and controls not
  const paragraph = ({ words, ...rest }) => rest;
  const view = (c) => {
    const r = c.region;
    const out = { error: c.error, result: c.result, range: c.range, comments: r.comments };
    // case 1 is read only, but its content holds tracked changes from the start
    const markup = c.tracked || c.n === 1;
    out.paragraphs = Array.isArray(r.paragraphs) ? listParagraphForOther(c.n, r.paragraphs).map((p) => (markup ? { ...paragraph(p), text: undefined, bold: undefined, italic: undefined, underline: undefined } : paragraph(p))) : r.paragraphs;
    out.changes = Array.isArray(r.trackedChanges) ? r.trackedChanges.map((x) => x.type) : r.trackedChanges;
    if (!markup) Object.assign(out, { text: r.text, tables: r.tables, controls: r.contentControls });
    return JSON.parse(JSON.stringify(out));
  };
  const differing = {};
  for (const [i, theirs] of word.cases.entries()) {
    const [a, b] = [view(ours.cases[i]), view(theirs)];
    const keys = [...new Set([...Object.keys(a), ...Object.keys(b)])].filter((k) => JSON.stringify(a[k]) !== JSON.stringify(b[k]));
    if (keys.length > 0) differing[theirs.n] = keys;
  }
  const each = (ns, keys) => Object.fromEntries(ns.map((n) => [n, keys]));
  assert.deepEqual(differing, {
    // a tracked Replace: Office JS lists the insertion alone, where the markup holds the deletion too, which this
    // package's listing (CR-002) lists; Word's save has it (the markup test below)
    ...each([39, 40, 41, 42, 43, 46, 47, 48], ['changes']),
    // the bold cases' report reads each word through getTextRanges, which is not implemented (the markup test holds
    // their w:b to Word's); Word lists formatting across a table's cells as one change, this package a run at a time
    ...each([49, 50, 51, 52, 53, 56], ['paragraphs']),
    ...each([54, 55], ['paragraphs', 'changes']),
    // on a control's content range Word's insertParagraph adds an empty paragraph as well, which is not copied (CR-014 section 6.2)
    ...each([60, 66], ['paragraphs', 'text', 'controls']),
    // insertComment (phase 3): the comments and their ranges are Word's, but the snippet calls getRange() at once
    // on what it returns, which is asynchronous here, so the script fails after the comment is in
    ...each([79, 80, 81, 82, 83, 84, 85], ['error', 'result']),
    // case 13, a control emptied: Word fills it with a bold run of five spaces (section 9.4), which reads bold; this
    // package's emptied paragraph reads its mark (CR-002 section 59)
    13: ['paragraphs'],
    // a table at a cell's end (CR-002 section 55.6): Word gives the new table's paragraph the cell paragraph's
    // properties, and so its numbering, which is insertTable's wherever it inserts and not followed (the markup test)
    ...each([86, 87], ['paragraphs']),
    // Office JS's Body.insertOoxml returns a Range; this package's the views of what went in (CR-002 phase I)
    94: ['result'],
  });
  // Word's refusals are this package's: a range starting or ending inside a table, deleted or replaced, tracked or not
  const refused = ours.cases.filter((c) => c.error === 'error: GeneralException').map((c) => c.n);
  assert.deepEqual(refused, [9, 10, 22, 23, 34, 35, 44, 45]);
  // ^p (phase 3), case 2's eleven queries: each paragraph's mark outside the table, a hit across two paragraphs'
  // mark, and nothing into or out of the table, at a cell's end, at a row's end, or for ^p^p
  assert.deepEqual(ours.cases[1].searches, word.cases[1].searches);
});

test("CR-014 phase 2: Accept All after check 52's writes through the shim gives Word 365's Accept All after its run, region for region", async () => {
  const { runCheck } = await snippet();
  const pkg = await loaded((await phase2Run()).saved);
  await pkg.acceptAllRevisions();
  const back = await loaded(await pkg.save());
  const ours = await runCheck((f) => Word.run(back, f), 'report');
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check52', '52-ranges-word365-accepted.json'), 'utf8'));
  const paragraph = ({ words, ...rest }) => rest;
  const differing = [];
  for (const [i, theirs] of word.cases.entries()) {
    // Range.footnotes is not implemented
    const view = ({ footnotes, ...r }) => ({ ...r, paragraphs: Array.isArray(r.paragraphs) ? listParagraphForOther(theirs.n, r.paragraphs).map(paragraph) : r.paragraphs });
    if (JSON.stringify(view(ours.cases[i].region)) !== JSON.stringify(view(theirs.region))) differing.push(theirs.n);
  }
  // every tracked deletion and replacement, accepted, is Word's, and every comment across paragraphs (phase 3);
  // what differs is listed in the test above: Word's bold filler in an emptied control (13, and 26 accepted),
  // getTextRanges (the bold cases), Word's extra empty paragraph in a control, the new table's paragraph in a cell
  // (CR-002 section 55.6); each cell's last paragraph accepted is Word's (86 to 95), and every paragraph's font
  // otherwise (CR-002 section 59)
  assert.deepEqual(differing, [13, 26, 49, 50, 51, 52, 53, 54, 55, 56, 60, 66, 86, 87]);
});

/**
 * A region's markup, a line a paragraph: its style, alignment and numbering, its mark's properties (`ins`, `del`,
 * `b`, `+change` for its w:rPrChange) and a w:pPrChange, then its runs, those with the same revision and
 * properties merged, each `ins:` or `del:` with its properties; a table's rows (`TR del`) and cells' paragraphs.
 * Ids, dates, rsids, paraIds, w:lastRenderedPageBreak and the complex-script twins (bCs, iCs, szCs) are left out.
 */
function markupOf(pkg) {
  const IGNORED = new Set(['TYPE_NAME', 'PARENT', 'rsidR', 'rsidRPr', 'rsidRDefault', 'rsidP', 'rsidDel', 'paraId', 'textId', 'bCs', 'iCs', 'szCs', 'lang']);
  const keys = (o) => (o ? Object.keys(o).filter((k) => !IGNORED.has(k) && o[k] !== undefined && k !== 'rPrChange').sort().join(',') + (o.rPrChange ? '+change' : '') : '');
  const runs = (items, revision = '') => (items ?? []).flatMap((el) => {
    const name = el.name?.localPart;
    if (name === 'ins' || name === 'del') return runs(el.value.customXmlOrSmartTagOrSdt, name);
    if (name === 'hyperlink') return runs(el.value.content, `${revision}link`);
    if (name !== 'r') return name?.startsWith('comment') || name?.startsWith('bookmark') ? [{ revision: name, rPr: '', text: '' }] : [];
    const text = (el.value.content ?? []).map((c) => {
      const n = c.name?.localPart;
      return n === 't' || n === 'delText' ? c.value.value ?? '' : n === 'lastRenderedPageBreak' ? '' : `<${n}>`;
    }).join('');
    return [{ revision, rPr: keys(el.value.rPr), text }];
  }).reduce((out, r) => {
    const last = out[out.length - 1];
    if (last && last.revision === r.revision && last.rPr === r.rPr) last.text += r.text; else out.push({ ...r });
    return out;
  }, []);
  const lines = (el, indent = '') => {
    const v = el.value;
    switch (el.name?.localPart) {
      case 'p': {
        const pPr = v.pPr ?? {};
        const props = [pPr.pStyle?.val, pPr.jc?.val, pPr.numPr ? 'num' : '', pPr.pPrChange ? 'pPrChange' : ''].filter(Boolean).join('/');
        const content = runs(v.content).map((r) => `${r.revision ? `${r.revision}:` : ''}${JSON.stringify(r.text)}${r.rPr ? `{${r.rPr}}` : ''}`).join(' ');
        return [`${indent}P[${props}|${keys(pPr.rPr)}] ${content}`];
      }
      case 'tbl': return (v.content ?? []).flatMap((tr) => [`${indent}TR ${keys(tr.value.trPr)}`,
        ...(tr.value.content ?? []).flatMap((tc) => (tc.value.content ?? []).flatMap((b) => lines(b, `${indent}  `)))]);
      case 'sdt': return [`${indent}SDT`, ...(v.sdtContent?.content ?? []).flatMap((b) => lines(b, `${indent}  `))];
      default: return [`${indent}<${el.name?.localPart}>`];
    }
  };
  const textOf = (el) => (el.name?.localPart === 'p' ? runs(el.value.content).map((r) => r.text).join('') : '');
  const regions = new Map();
  let n = null;
  for (const el of pkg.getMainDocumentPart().contents.body.content) {
    const begin = /^BEGIN(\d+) /.exec(textOf(el));
    if (begin) { n = Number(begin[1]); regions.set(n, []); continue; }
    if (n !== null && textOf(el) === `END${n}`) { n = null; continue; }
    if (n !== null) regions.get(n).push(...lines(el));
  }
  return regions;
}

test("CR-014 phase 2: the markup check 52's writes leave through the shim is Word 365's in its save after the run, region for region", async () => {
  const mine = markupOf(await loaded((await phase2Run()).saved));
  const word = markupOf(await loaded(await saveBytes('52-ranges-word365.docx')));
  const differing = [...word.keys()].filter((n) => mine.get(n).join('\n') !== word.get(n).join('\n'));
  // 1, read only: Word's save gave the comment reference's run w:u and w:sz of its own; 13: Word's filler for an
  // emptied control, a bold run of five spaces (section 9.4); 60 and 66: Word's extra
  // empty paragraph in a control; 67 and 68: Word gives a paragraph's mark its last run's formatting (italic) when
  // insertParagraph('After') goes after it, not followed (CR-014 section 10); 79 to 85: insertComment (phase 3),
  // whose markers are Word's but for the two differences below; 86 and 87: the new table's paragraph (below)
  assert.deepEqual(differing, [1, 13, 60, 66, 67, 68, 79, 80, 81, 82, 83, 84, 85, 86, 87]);
  // the comments across paragraphs (phase 3, section 11): the start marker before the range's first character,
  // the end and the reference run after its last, in its last paragraph, as Word's. Word gives the reference run
  // the text's size and, where text follows it, that text's italic, over the CommentReference style; this
  // package's reference run has the style alone, as CR-002's within one paragraph has. And on a control's content
  // range (82) Word put the start before the control's first paragraph, which reads the same (section 9.2)
  for (const n of [79, 80, 81, 82, 83, 84, 85]) {
    let expected = word.get(n).map((line) => line.replace(/"<commentReference>"\{(i,)?rStyle,sz\}/, '"<commentReference>"{rStyle}'));
    if (n === 82) expected = expected.filter((line) => line !== '  <commentRangeStart>').map((line) => line.replace('  P[Heading1|] "HEAD82A', '  P[Heading1|] commentRangeStart:"" "HEAD82A'));
    assert.deepEqual(mine.get(n), expected, `case ${n}`);
  }
  assert.deepEqual(word.get(13), ['SDT', '  P[ListParagraph/num|] "     "{b}']);
  assert.deepEqual(mine.get(13), ['SDT', '  P[ListParagraph/num|] ']);
  for (const n of [67, 68]) {
    assert.deepEqual(mine.get(n).map((line) => line.replace(/^P\[center\|\] "BODY/, 'P[center|i] "BODY')), word.get(n), `case ${n}`);
  }
  // a table at a cell's end (CR-002 section 55.6): the cell's paragraph split at its end, the first half's mark in its
  // last run's formatting and, tracked, inserted; the cell's own mark last, after the table, unmarked. Word gives the
  // new table's paragraph the cell paragraph's properties too, which is insertTable's and not followed
  for (const n of [86, 87]) {
    const besideTheTable = (lines) => lines.filter((_, k) => k !== 3);
    assert.deepEqual(besideTheTable(mine.get(n)), besideTheTable(word.get(n)), `case ${n}`);
  }
  assert.deepEqual([word.get(86)[3], mine.get(86)[3]], ['    P[ListParagraph/num|u] "NEWCELL"{u}', '    P[|] "NEWCELL"']);
  assert.deepEqual([word.get(87)[3], mine.get(87)[3]], ['    P[ListParagraph/num|ins,u+change] ins:"NEWCELL"{u}', '    P[|ins] ins:"NEWCELL"']);
  // what Word wrote, which this package now writes: a tracked deletion across paragraphs marks every mark but the
  // last deleted, with no w:pPrChange (case 16); a tracked Replace puts its w:ins before the w:del (case 39); an
  // untracked one gives its text the heading's look in a body paragraph (case 29); bold skips the heading's runs
  // and bolds a mark in the range, tracked with its w:rPrChange (case 56); a paragraph split tracked (case 70)
  assert.deepEqual(word.get(16), ['P[Heading1|del] "HEAD16A " del:"HEAD16B"', 'P[center|] del:"BODY16A"{i} " BODY16B"{i}']);
  assert.deepEqual(word.get(39), ['P[Heading1|del] "HEAD39A " ins:"NEWTEXT" del:"HEAD39B"', 'P[center|] del:"BODY39A"{i} " BODY39B"{i}']);
  assert.deepEqual(word.get(29), ['P[center|] "HEAD29A " "NEWTEXT"{b,color,rFonts,sz} " BODY29B"{i}']);
  assert.deepEqual(word.get(56).slice(1, 3), ['  P[Heading1|] "HEAD56A HEAD56B"', '  P[center|b+change] "BODY56A BODY56B"{b,i+change}']);
  assert.deepEqual(word.get(70).slice(1), ['P[center|i,ins] "BODY70A"{i}', 'P[center|i,ins] ins:"NEWPARA"{i}', 'P[center|] " BODY70B"{i}']);
});

test('CR-014 phase 2: delete and Replace across paragraphs, and the range a write leaves; Reject All after a tracked write gives back what was there', async () => {
  const pkg = await loaded(await check52Bytes());
  const region = (n) => { const ps = pkg.body.paragraphs; const i = ps.findIndex((p) => p.text.startsWith(`BEGIN${n} `)); return ps.slice(i + 1, ps.findIndex((p, k) => k > i && p.text === `END${n}`)); };
  const texts = (n) => region(n).map((p) => p.text);
  // untracked: the paragraphs join in the last one's properties, and the range collapses where they joined
  const range = wordIn(pkg, 'HEAD3B').expandTo(wordIn(pkg, 'BODY3A'));
  range.delete();
  assert.deepEqual(texts(3), ['HEAD3A  BODY3B']);
  assert.deepEqual([region(3)[0].alignment, region(3)[0].styleBuiltIn], ['Centered', 'Normal']);
  assert.deepEqual([range.spansParagraphs, range.text, range.start], [false, '', 7]);
  // Replace: the range covers the new text; '' is a deletion that keeps a last mark
  const replaced = wordIn(pkg, 'HEAD38B').expandTo(wordIn(pkg, 'BODY38A').paragraphs[0].getRange('Whole'));
  const inserted = replaced.insertText('NEW', 'Replace');
  assert.deepEqual([inserted.text, replaced.text, texts(38)], ['NEW', 'NEW', ['HEAD38A NEW', 'NEXT38A NEXT38B']]);
  wordIn(pkg, 'HEAD29B').expandTo(wordIn(pkg, 'BODY29A')).insertText('', 'Replace');
  assert.deepEqual(texts(29), ['HEAD29A  BODY29B']);
  // a footnote whose reference a deletion takes leaves the notes part on save (check 52 case 12)
  wordIn(pkg, 'HEAD12B').expandTo(wordIn(pkg, 'BODY12A')).delete();
  const notesOf = async (p) => (await p.getMainDocumentPart().footnotesPart.getContents()).footnote.filter((f) => f.type === undefined || f.type === 'normal').length;
  const notes = await notesOf(pkg);
  assert.equal(await notesOf(await WordprocessingMLPackage.load(await pkg.save())), notes - 1);

  // tracked: one deletion across the paragraphs; Reject All gives back the paragraphs, Accept All the untracked join
  const original = texts(16);
  pkg.changeTrackingMode = 'TrackAll';
  const tracked = wordIn(pkg, 'HEAD16B').expandTo(wordIn(pkg, 'BODY16A'));
  tracked.delete();
  assert.deepEqual(pkg.body.getTrackedChanges().filter((c) => /16/.test(c.text)).map((c) => [c.type, c.text]), [['Deleted', 'HEAD16B\rBODY16A']]);
  assert.deepEqual([tracked.spansParagraphs, tracked.text], [false, ''], 'collapsed where the deletion starts, as Simple Markup reads it');
  wordIn(pkg, 'HEAD49B').expandTo(wordIn(pkg, 'BODY49A')).font.bold = true;
  wordIn(pkg, 'HEAD57B').expandTo(wordIn(pkg, 'BODY57A')).insertParagraph('NEW', 'Before');
  wordIn(pkg, 'HEAD39B').expandTo(wordIn(pkg, 'BODY39A')).insertText('NEW', 'Replace');
  pkg.changeTrackingMode = 'Off';
  assert.deepEqual(texts(57), ['HEAD57A ', 'NEW', 'HEAD57B', 'BODY57A BODY57B']);
  const rejected = await loaded(await pkg.save());
  await rejected.rejectAllRevisions();
  const after = (n) => { const ps = rejected.body.paragraphs; const i = ps.findIndex((p) => p.text.startsWith(`BEGIN${n} `)); return ps.slice(i + 1, ps.findIndex((p, k) => k > i && p.text === `END${n}`)); };
  assert.deepEqual(after(16).map((p) => p.text), original);
  assert.deepEqual(after(57).map((p) => [p.text, p.styleBuiltIn]), [['HEAD57A HEAD57B', 'Heading1'], ['BODY57A BODY57B', 'Normal']]);
  assert.deepEqual(after(39).map((p) => p.text), ['HEAD39A HEAD39B', 'BODY39A BODY39B']);
  // the bold rejected: no run and no mark keeps a w:b or a change
  const xml = await rejected.getMainDocumentPart().getXml();
  const region49 = xml.slice(xml.indexOf('BEGIN49 '), xml.indexOf('END49'));
  assert.doesNotMatch(region49, /<w:b\/>|w:rPrChange/);
});

test('CR-014 phase 2: whole tables, a cell, the body and a control; what Word refuses is refused before anything changes', async () => {
  const pkg = await loaded(await check52Bytes());
  const body = pkg.body;
  // a range starting or ending inside a table: Word's GeneralException, tracked or not, and nothing changed
  const before = body.text;
  for (const mode of ['Off', 'TrackAll']) {
    pkg.changeTrackingMode = mode;
    for (const [from, to] of [['HEAD9B', 'T9R2C1A'], ['T10R1C1B', 'BODY10A']]) {
      const range = wordIn(pkg, from).expandTo(wordIn(pkg, to));
      assert.throws(() => range.delete(), (e) => e.code === 'GeneralException' && /inside a table/.test(e.message));
      assert.throws(() => range.insertText('X', 'Replace'), (e) => e.code === 'GeneralException');
    }
  }
  pkg.changeTrackingMode = 'Off';
  assert.equal(body.text, before);
  // Word applies a font and a paragraph to such a range (check 52 cases 51 and 59)
  const cell = wordIn(pkg, 'T51R1C1B').expandTo(wordIn(pkg, 'BODY51A'));
  cell.font.bold = true;
  assert.equal(cell.font.bold, true);
  // a range holding a whole block control: under tracking the control is deleted with Word's markers (check 54
  // case 15, CR-014 section 12), Accept All taking it and joining the paragraphs, Reject All giving them back;
  // untracked, the control goes as a table does and the paragraphs either side join
  const tracked = await loaded(await pkg.save());
  tracked.changeTrackingMode = 'TrackAll';
  wordIn(tracked, 'BEGIN26').expandTo(wordIn(tracked, 'END26')).delete();
  tracked.changeTrackingMode = 'Off';
  assert.match(await tracked.getMainDocumentPart().getXml(), /<w:customXmlDelRangeStart[^>]*\/><w:sdt>.*?<w:sdtContent><w:customXmlDelRangeEnd/);
  const [accepted, rejected] = [await loaded(await tracked.save()), await loaded(await tracked.save())];
  await accepted.acceptAllRevisions();
  await rejected.rejectAllRevisions();
  assert.equal(accepted.body.contentControls.filter((c) => c.tag === 'c52-26').length, 0);
  assert.equal(accepted.body.paragraphs.filter((p) => /BEGIN26|END26/.test(p.text)).length, 0);
  // the region as it was (Reject All rejects case 1's changes too, so the region alone is compared)
  const region26 = (doc) => { const t = doc.body.text; return t.slice(t.indexOf('BEGIN26 '), t.indexOf('END26')); };
  assert.equal(region26(rejected), region26(pkg));
  const around = wordIn(pkg, 'BEGIN26').expandTo(wordIn(pkg, 'END26'));
  const [controls, paragraphs] = [body.contentControls.length, body.paragraphs.length];
  around.delete();
  // BEGIN26, the control's three paragraphs and END26 become one empty paragraph
  assert.deepEqual([body.contentControls.length, body.paragraphs.length], [controls - 1, paragraphs - 4]);
  assert.equal(body.paragraphs.filter((p) => /BEGIN26|END26/.test(p.text)).length, 0);
  // into a block content control from outside it: the texts go and nothing joins across the control's edge; the
  // start's paragraph takes the end's properties (check 54 case 18)
  const control = body.contentControls.find((c) => c.tag === 'c52-13');
  const begin = wordIn(pkg, 'BEGIN13').paragraphs[0];
  wordIn(pkg, 'BEGIN13').expandTo(control.paragraphs[0].search('HEAD13A')[0]).delete();
  assert.deepEqual([begin.text, begin.styleBuiltIn], ['', 'Heading1']);
  assert.deepEqual([control.paragraphs[0].text, control.paragraphs[0].styleBuiltIn], [' HEAD13B', 'Heading1']);
  // a table's whole range takes the table; a cell's whole range keeps its last paragraph, emptied (CR-002 section 55)
  const tables = body.tables.length;
  body.tables[0].getRange().delete();
  assert.equal(body.tables.length, tables - 1);
  const cellBody = body.tables[0].getCell(0, 0).body;
  cellBody.getRange().delete();
  assert.deepEqual(cellBody.paragraphs.map((p) => p.text), ['']);
  // the control's content: one paragraph left, in its last paragraph's properties (check 52 case 13)
  control.getRange('Content').delete();
  assert.deepEqual(control.paragraphs.map((p) => [p.text, p.styleBuiltIn]), [['', 'ListParagraph']]);
  // the body's content: one empty paragraph, in the last one's properties
  body.getRange('Content').delete();
  assert.deepEqual(body.paragraphs.map((p) => p.text), ['']);
  assert.equal(body.tables.length, 0);
});

test("a range holding one table whole, nothing either side, deletes that table and not the block after it (docx4j-python's finding, 2026-10-23)", async () => {
  const table = '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>'
    + ['IN1', 'IN2'].map((t) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`).join('') + '</w:tr></w:tbl>';
  const shape = (body) => body.content.map((el) => (el.name.localPart === 'p' ? body.paragraphFor(el.value).text : el.name.localPart));
  for (const tracked of [false, true]) {
    const pkg = await WordprocessingMLPackage.createPackage();
    const body = await pkg.getBody();
    await body.insertXml(`<w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>${table}<w:p><w:r><w:t>AFTER</w:t></w:r></w:p><w:p><w:r><w:t>LAST</w:t></w:r></w:p>`, 'End');
    const range = body.tables[0].getRange('Whole');
    assert.equal(range.spansParagraphs, true);
    if (tracked) pkg.changeTrackingMode = 'TrackAll';
    range.delete();
    pkg.changeTrackingMode = 'Off';
    if (!tracked) { assert.deepEqual(shape(body), ['BEFORE', 'AFTER', 'LAST']); continue; }
    // tracked, its row deleted once; Accept All takes the table, the paragraphs either side as they were
    assert.deepEqual(shape(body), ['BEFORE', 'tbl', 'AFTER', 'LAST']);
    assert.equal(body.getTrackedChanges().length, 1);
    await pkg.acceptAllRevisions();
    assert.deepEqual(shape(body), ['BEFORE', 'AFTER', 'LAST']);
  }
});

test('CR-014 phase 2: insertParagraph splits a paragraph inside a link into two links, and refuses a split inside a content control; insertText at the ends; replaceText a paragraph at a time', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  await body.insertXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:jc w:val="center"/></w:pPr>'
    + '<w:hyperlink w:anchor="here"><w:r><w:t>linked text</w:t></w:r></w:hyperlink><w:r><w:t> after</w:t></w:r></w:p>', 'End');
  await body.insertXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>'
    + '<w:sdtContent><w:r><w:t>in a control</w:t></w:r></w:sdtContent></w:sdt></w:p>', 'End');
  body.insertParagraph('last one', 'End');
  const [linked, controlled] = body.paragraphs.filter((p) => p.text !== '');
  const range = linked.search('text')[0].expandTo(body.search('last')[0]);
  range.insertParagraph('NEW', 'Before');
  assert.deepEqual(body.paragraphs.filter((p) => p.text !== '').map((p) => [p.text, p.alignment]).slice(0, 3),
    [['linked ', 'Centered'], ['NEW', 'Centered'], ['text after', 'Centered']]);
  assert.deepEqual(body.paragraphs.filter((p) => p.text !== '').slice(0, 3).map((p) => p.getRange().hyperlink), ['#here', '', '#here']);
  assert.deepEqual([range.text, range.start], ['text after\rin a control\rlast', 0]);
  // a split inside an inline control would make two of it: refused, nothing changed
  const across = controlled.search('control')[0].expandTo(body.search('last')[0]);
  const xml = await pkg.getMainDocumentPart().getXml();
  assert.throws(() => across.insertParagraph('X', 'Before'), NotSupportedError);
  assert.equal(await pkg.getMainDocumentPart().getXml(), xml);
  // insertText at the start and the end; replaceText in each paragraph
  range.insertText('<', 'Start');
  range.insertText('>', 'End');
  assert.equal(range.text, 'text after\rin a control\rlast');
  assert.equal(body.text.split('\r').filter((t) => t !== '').join('|'), 'linked |NEW|<text after|in a control|last> one');
  assert.equal(range.replaceText('t', 'T'), 5);
  assert.equal(range.text, 'TexT afTer\rin a conTrol\rlasT');
});

// ------------------------------------------------------------------- check 54 (phase 2's choices, phase 3's controls)

test("check 54's file and snippet agree: every case parses to the case it was built as, every word its range names is in the file once", async () => {
  const source = await readFile(join(here, 'fixtures', 'check54', 'check54-script-lab.js'), 'utf8');
  const { parseRegion, runCheck } = new Function(`${source}\nreturn CHECK54;`)();
  const pkg = await loaded(await check54Bytes());
  const texts = pkg.body.paragraphs.map((p) => p.text);
  const regions = texts.map(parseRegion).filter((r) => r !== null);
  const cases = check54Cases();
  assert.equal(cases.length, 43);
  assert.deepEqual(regions.map((r) => [r.n, r.kind, r.op, r.tracked, r.spec]), cases.map((c) => [c.n, c.kind, c.op, c.tracking === 'tracked', c.spec]));
  const count = (token) => texts.reduce((sum, text) => sum + (text.match(new RegExp(`(?<![A-Za-z0-9])${token}(?![A-Za-z0-9])`, 'g')) ?? []).length, 0);
  for (const r of regions) {
    // the OWN case's last word is in the paragraph its run inserts first
    const tokens = [`BEGIN${r.n}`, `END${r.n}`, ...Object.values(r.range)].filter((token) => !token.startsWith('OWN'));
    for (const token of tokens) assert.equal(count(token), 1, `case ${r.n}: ${token}`);
  }
  // the snippet runs every case through the Word shim with no error of its own: what fails is this package's
  // refusals (NotImplemented) and getBookmarks, which the shim doesn't have
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  assert.equal(ours.cases.length, 43);
  const errors = new Set(ours.cases.flatMap((c) => [c.error, ...Object.values(typeof c.region === 'string' ? { region: c.region } : c.region)])
    .filter((v) => typeof v === 'string' && v.startsWith('error:')));
  assert.deepEqual([...errors].sort(), ['error: GeneralException', 'error: NotImplemented']);
});


// ------------------------------------------------------------------------------------- phase 3, comments and ^p

test('CR-014 phase 3: a comment across paragraphs, its range the range commented; ^p finds paragraph marks as Word does', async () => {
  const pkg = await loaded(await check52Bytes());
  const body = pkg.body;
  // a comment from a heading's middle across a table into a body paragraph: one Range, as Word's (check 52 case 80)
  const range = wordIn(pkg, 'HEAD80B').expandTo(wordIn(pkg, 'BODY80A'));
  const comment = await range.insertComment('NEW');
  assert.equal(comment.getRange().text, range.text);
  assert.deepEqual((await range.getComments()).map((c) => c.content), ['NEW']);
  // a range starting at a paragraph's end and one ending at the next paragraph's start: the markers stand there
  const atEnd = wordIn(pkg, 'HEAD3B').paragraphs[0].getRange('End').expandTo(wordIn(pkg, 'BODY3A'));
  assert.equal((await atEnd.insertComment('END')).getRange().text, '\rBODY3A');
  const toStart = wordIn(pkg, 'HEAD16B').expandTo(wordIn(pkg, 'BODY16A').paragraphs[0].getRange('Start'));
  assert.equal((await toStart.insertComment('START')).getRange().text, 'HEAD16B\r');
  const saved = await loaded(await pkg.save());
  assert.deepEqual((await saved.body.getComments()).filter((c) => ['NEW', 'END', 'START'].includes(c.content)).map((c) => c.getRange().text),
    ['\rBODY3A', 'HEAD16B\r', range.text]);

  // ^p: each paragraph's own mark the range holds, a hit across a mark, nothing at a cell's end or across a table's edge
  const region = wordIn(pkg, 'BEGIN2').paragraphs[0].getRange('Whole').expandTo(wordIn(pkg, 'END2').paragraphs[0].getRange('Whole'));
  assert.equal(region.search('^p').length, 6);
  assert.deepEqual(region.search('HEAD2B^pBODY2A').map((r) => [r.text, r.paragraphs.length]), [['HEAD2B\rBODY2A', 2]]);
  assert.deepEqual(['LIST2B^pT2R1C1A', 'T2R1C1B^p', 'T2R2C2B^pPLAIN2A', '^p^p'].map((q) => region.search(q).length), [0, 0, 0, 0]);
  const hit = region.search('^p')[1];
  assert.deepEqual([hit.text, hit.paragraph.text, hit.holdsEndMark], ['\r', 'HEAD2A HEAD2B', true]);
  // a paragraph's search reads its own mark; the body's every mark outside a table
  assert.deepEqual(wordIn(pkg, 'HEAD2A').paragraphs[0].search('^p').map((r) => r.text), ['\r']);
  assert.equal(body.search('B^pT', { matchCase: true }).length, 0, 'into a table from any paragraph before one');
  assert.ok(body.search('B^pE', { matchCase: true }).length > 80, 'into each END paragraph');
  assert.equal(body.search('HEAD2B^pBODY2A').length, 1);
  // a cell's paragraphs: a mark between two of them is a paragraph's; the cell's last isn't
  const cell = body.tables[0].getCell(0, 0).body;
  cell.paragraphs[0].insertParagraph('SECOND', 'After');
  assert.deepEqual(cell.getRange().search('^p').map((r) => r.paragraph.text), ['T1R1C1A T1R1C1B']);
  assert.deepEqual(cell.getRange().search('T1R1C1B^pSECOND').map((r) => r.text), ['T1R1C1B\rSECOND']);
  // an empty paragraph's mark after another mark: ^p^p
  wordIn(pkg, 'HEAD2B').paragraphs[0].insertParagraph('', 'After');
  assert.deepEqual(region.search('^p^p').map((r) => r.text), ['\r\r']);
  assert.deepEqual(region.search('HEAD2B^p^pBODY2A').map((r) => r.text), ['HEAD2B\r\rBODY2A']);
  // wildcards don't read ^p as a mark, but as the two characters: the twelve in case 2's BEGIN paragraph, which lists
  // the queries; a replacement whose query holds one is refused, unmeasured
  assert.deepEqual(region.search('^p', { matchWildcards: true }).map((r) => r.text), Array(12).fill('^p'));
  for (const replace of [() => body.replaceText('^p', ' '), () => region.replaceText('^p', ' '), () => region.paragraphs[0].replaceText('^p', ' ')]) {
    assert.throws(replace, (e) => e instanceof NotSupportedError && /\^p/.test(e.message));
  }
});

test('CR-014 phase 3: a comment\'s markers go beside runs holding nothing outside its range: a tab sharing a run is split off, within one paragraph and across two', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:tab/><w:t>c</w:t></w:r></w:p>';
  await body.insertXml(xml + xml, 'End');
  const [first, second] = body.paragraphs.filter((p) => p.text !== '');
  // within one paragraph: from 'b' to its end, where the run starts at 'a'
  const one = first.getRange('Start');
  one.start = 2;
  one.end = 5;
  assert.equal((await one.insertComment('ONE')).getRange().text, 'b\tc');
  // across two: from the first paragraph's 'c' to the second's 'a'
  const a = first.getRange('Start');
  a.start = a.end = 4;
  const across = a.expandTo(second.search('a')[0]);
  assert.equal(across.text, 'c\ra');
  assert.equal((await across.insertComment('TWO')).getRange().text, 'c\ra');
  assert.deepEqual([first.text, second.text], ['a\tb\tc', 'a\tb\tc']);
});

const check54 = (name) => readFile(join(here, 'fixtures', 'check54', name), 'utf8').then(JSON.parse);
const save54 = (name) => readFile(join(here, 'check54', name));

test("check 54's run and its report after Accept All (Word 365, 2026-10-22) are of this file, and its save is Word's, as copied", async () => {
  for (const [name, mode] of [['54-writes-word365-run.json', 'run'], ['54-writes-word365-accepted.json', 'report']]) {
    const word = await check54(name);
    assert.equal(word.mode, mode);
    assert.deepEqual(word.cases.map((c) => [c.n, c.kind, c.op, c.tracked]), check54Cases().map((c) => [c.n, c.kind, c.op, c.tracking === 'tracked']));
  }
  const manifest = JSON.parse(await readFile(join(here, 'check54', 'manifest.json'), 'utf8'));
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await save54(name)).digest('hex'), sha1, name);
  assert.deepEqual(manifest.files.map((f) => f.name), ['54-writes-word365.docx', '54-writes-word365-accepted.docx', '54-writes-word365-rejected.docx']);
});

/** Check 54's snippet in its 'run' mode, through the Word shim on its own file, once for the tests below. */
let phase3;
const phase3Run = () => (phase3 ??= (async () => {
  const source = await readFile(join(here, 'fixtures', 'check54', 'check54-script-lab.js'), 'utf8');
  const { runCheck } = new Function(`${source}\nreturn CHECK54;`)();
  const pkg = await loaded(await check54Bytes());
  const result = await runCheck((f) => Word.run(pkg, f), 'run');
  return { result, saved: await pkg.save(), runCheck };
})());

test("CR-014 phase 3: check 54's writes, run through the Word shim on the same file, leave each region as Word 365's run did", async () => {
  const { result: ours } = await phase3Run();
  const word = await check54('54-writes-word365-run.json');
  // as check 52's: under tracking Word's text is All Markup's, so a tracked case's paragraphs are compared without
  // their text, and its text, tables and controls not; a paragraph's font isn't reported; getBookmarks is not
  // implemented in the shim
  const view = (c) => {
    const r = c.region;
    const out = { error: c.error, result: c.result, range: c.range, comments: r.comments, hyperlink: r.hyperlink };
    out.paragraphs = Array.isArray(r.paragraphs) ? r.paragraphs.map((p) => (c.tracked ? { ...p, text: undefined } : p)) : r.paragraphs;
    out.changes = Array.isArray(r.trackedChanges) ? r.trackedChanges.map((x) => x.type) : r.trackedChanges;
    if (!c.tracked) Object.assign(out, { text: r.text, tables: r.tables, controls: r.contentControls });
    return JSON.parse(JSON.stringify(out));
  };
  const differing = {};
  for (const [i, theirs] of word.cases.entries()) {
    const [a, b] = [view(ours.cases[i]), view(theirs)];
    const keys = [...new Set([...Object.keys(a), ...Object.keys(b)])].filter((k) => JSON.stringify(a[k]) !== JSON.stringify(b[k]));
    if (keys.length > 0) differing[theirs.n] = keys;
  }
  const each = (ns, keys) => Object.fromEntries(ns.map((n) => [n, keys]));
  assert.deepEqual(differing, {
    // a tracked Replace within one paragraph: Office JS lists the insertion alone, this package the deletion too
    4: ['changes'],
    // an emptied list item in a cell reads Word's style 'Other' (CR-002 section 57's family)
    ...each([10, 11], ['paragraphs']),
    // Word lists a control's last paragraph's text and mark as two deletions, this package one (CR-002's listing)
    13: ['changes'],
    // a tracked Replace over a whole control: Office JS lists the insertion alone, this package the deletion too
    17: ['changes'],
    // tracked, Word's comment range reads the deleted word (All Markup)
    28: ['comments'],
    // Word's model marks the first paragraph's mark inserted when it inserts one after it, so deleting into the new
    // paragraph took its own mark back and joined them; this package marks the new one's (CR-002 section 13): the
    // same document once accepted or rejected (the Accept All test below)
    31: ['paragraphs'],
    // a style across paragraphs: Word also took away direct formatting (centring, numbering), not followed, and
    // listed one change
    32: ['paragraphs'], 33: ['paragraphs', 'changes'],
    // a hyperlink across paragraphs: Word wrote a HYPERLINK field across them, refused here
    34: ['error', 'hyperlink', 'paragraphs', 'result'],
    // insertOoxml is asynchronous here, so the snippet reads no text off what it returns
    35: ['result'],
    // a tracked control around paragraphs: Word lists its two insertion markers as changes, this package nothing
    43: ['changes'],
  });
});

test("CR-014 phase 3: Accept All after check 54's writes through the shim is Word 365's Accept All, in each region's report and markup", async () => {
  const { saved, runCheck } = await phase3Run();
  const pkg = await loaded(saved);
  await pkg.acceptAllRevisions();
  const back = await loaded(await pkg.save());
  const ours = await runCheck((f) => Word.run(back, f), 'report');
  const word = await check54('54-writes-word365-accepted.json');
  const reported = [];
  for (const [i, theirs] of word.cases.entries()) {
    // getBookmarks is not implemented in the shim
    const view = ({ bookmarks, ...r }) => r;
    if (JSON.stringify(view(ours.cases[i].region)) !== JSON.stringify(view(theirs.region))) reported.push(theirs.n);
  }
  const mine = markupOf(back);
  const words = markupOf(await loaded(await save54('54-writes-word365-accepted.docx')));
  const marked = [...words.keys()].filter((n) => mine.get(n).join('\n') !== words.get(n).join('\n'));
  // every case's result is Word's, untracked or accepted, the tracked deletions of a whole control among them (15,
  // 17), but for: a style's resets (32, 33) and a hyperlink across paragraphs (34), not followed; the reference
  // run's size (29, section 11.2); and Word's 'Other' for an emptied list item in a cell (10, 11, its report alone)
  assert.deepEqual(reported, [10, 11, 32, 33, 34]);
  assert.deepEqual(marked, [29, 32, 33, 34]);
  // Word's Accept All, which docx4j's resolution and this package's follow since docx4j CR-043 section 17 (29dbf5964,
  // CR-014 section 12.3): an emptied paragraph whose deleted mark had a table after it goes (9); a control's emptied
  // last paragraph keeps a bare mark (13); a paragraph whose deleted mark had a control after it takes the control's
  // first paragraph's properties, its content not joined (19)
  assert.equal(mine.get(9)[0], 'TR ');
  assert.equal(mine.get(13)[3], '  P[|] ');
  assert.equal(mine.get(19)[0], 'P[center|] "HEAD19A "');
});

test("check 54: rejectAllRevisions on Word 365's save after the run is Word's Reject All, region for region (docx4j CR-043 section 17)", async () => {
  const pkg = await loaded(await save54('54-writes-word365.docx'));
  await pkg.rejectAllRevisions();
  const mine = markupOf(await loaded(await pkg.save()));
  const word = markupOf(await loaded(await save54('54-writes-word365-rejected.docx')));
  assert.equal(word.size, 43);
  const differing = [...word.keys()].filter((n) => mine.get(n).join('\n') !== word.get(n).join('\n'));
  // the original everywhere, as Word's, every region's markup the same
  assert.deepEqual(differing, []);
});

test("docx4j CR-043 section 17's text-left probe: Accept All and Reject All are Word 365's, region for region, its case 4 aside", async () => {
  const dir = join(here, 's17');
  const manifest = JSON.parse(await readFile(join(dir, 'manifest.json'), 'utf8'));
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await readFile(join(dir, name))).digest('hex'), sha1, name);
  for (const [verb, save] of [['acceptAllRevisions', '17-text-left-accepted.docx'], ['rejectAllRevisions', '17-text-left-rejected.docx']]) {
    const pkg = await loaded(await readFile(join(dir, '17-text-left.docx')));
    await pkg[verb]();
    const mine = markupOf(await loaded(await pkg.save()));
    const word = markupOf(await loaded(await readFile(join(dir, save))));
    assert.deepEqual([...word.keys()], [1, 2, 3, 4, 5], verb);
    // case 4, a deleted mark on a cell's last paragraph, is no edit Word writes, and Word's Accept All dropped the
    // one-cell table; docx4j's resolution keeps the cell's paragraph, and so does this package's (docx4j's section 11)
    for (const n of [1, 2, 3, 5]) assert.deepEqual(mine.get(n), word.get(n), `${verb}, case ${n}`);
  }
  // Accept All: the text joins the table's first cell's first paragraph (1); a control's last paragraph keeps a bare
  // mark, its text left (2); an emptied paragraph before a control goes (3); text left before a control takes its first
  // paragraph's properties, its numbering too (5)
  const accepted = markupOf(await loaded(await readFile(join(dir, '17-text-left-accepted.docx'))));
  assert.deepEqual(accepted.get(1).slice(0, 2), ['TR ', '  P[|] "HEAD1A T1R1C1"']);
  assert.equal(accepted.get(3)[0], 'SDT');
});

test("CR-014 phase 3: the markup check 54's writes leave through the shim is Word 365's in its save after the run, region for region", async () => {
  const mine = markupOf(await loaded((await phase3Run()).saved));
  const word = markupOf(await loaded(await save54('54-writes-word365.docx')));
  const differing = [...word.keys()].filter((n) => mine.get(n).join('\n') !== word.get(n).join('\n'));
  // 28: the reference run's size, which Word's save gave the file's own reference run (as check 52's case 1); the
  // run itself is in the w:del, as Word's, since section 12.6; 29: the reference run's size; 31: Word's model of an inserted paragraph's mark; 32 and
  // 33: a style's resets; 34: the hyperlink field (CR-014 section 12)
  assert.deepEqual(differing, [28, 29, 31, 32, 33, 34]);
  // what Word wrote, which this package now writes: a whole control a tracked deletion took, its content deleted
  // between Word's customXmlDelRange markers (15); a control around paragraphs held whole, tracked (43)
  assert.deepEqual(word.get(15), ['P[Heading1|del] "HEAD15A " del:"HEAD15B HEAD15C"', '<customXmlDelRangeStart>', 'SDT', '  <customXmlDelRangeEnd>',
    '  P[|del] del:"PLAIN15A PLAIN15B"', '  <customXmlDelRangeStart>', '<customXmlDelRangeEnd>', 'P[center|] del:"BODY15A"{i} " BODY15B"{i}']);
  assert.deepEqual(word.get(43).slice(0, 3), ['<customXmlInsRangeStart>', 'SDT', '  <customXmlInsRangeEnd>']);
  // and both the same in the shim's run: all but those six are the markup of Word's run, case 4's tracked Replace
  // within one paragraph among them since it puts its w:ins first (Jason, 2026-10-23)
  assert.equal(word.size, 43);
});


test('CR-014 section 12.6: a deletion takes a comment\'s reference with it, and the comment goes, untracked and on Accept All; Reject All keeps it', async () => {
  const made = async () => {
    const pkg = await WordprocessingMLPackage.createPackage();
    const body = await pkg.getBody();
    const p = body.insertParagraph('A B C D', 'End');
    await p.search('B C')[0].insertComment('ON B C');
    return WordprocessingMLPackage.load(await pkg.save());
  };
  const commentsOf = async (bytes) => { const pkg = await loaded(bytes); return (await pkg.body.getComments()).map((c) => c.content); };
  // untracked: "C D" takes the reference after "C", and the comment goes on save
  const untracked = await made();
  await untracked.getBody();
  untracked.body.paragraphs.find((p) => p.text.startsWith('A')).search('C D')[0].delete();
  assert.deepEqual(await commentsOf(await untracked.save()), []);
  // tracked: the reference run in the w:del with the text, as Word's (check 54 case 28)
  const tracked = await made();
  await tracked.getBody();
  tracked.changeTrackingMode = 'TrackAll';
  tracked.body.paragraphs.find((p) => p.text.startsWith('A')).search('C D')[0].delete();
  tracked.changeTrackingMode = 'Off';
  assert.match(await tracked.getMainDocumentPart().getXml(), /<w:del [^>]*><w:r><w:rPr><w:rStyle w:val="CommentReference"\/><\/w:rPr><w:commentReference w:id="0"\/><\/w:r><w:r><w:delText xml:space="preserve"> D<\/w:delText>/);
  const saved = await tracked.save();
  const accepted = await loaded(saved);
  await accepted.acceptAllRevisions();
  assert.deepEqual(await commentsOf(await accepted.save()), []);
  const rejected = await loaded(saved);
  await rejected.rejectAllRevisions();
  assert.deepEqual(await commentsOf(await rejected.save()), ['ON B C']);
  // a replacement over the reference takes it too
  const empty = await made();
  await empty.getBody();
  const p = empty.body.paragraphs.find((q) => q.text.startsWith('A'));
  p.search('A B C D')[0].insertText('', 'Replace');
  assert.equal(p.text, '');
  assert.deepEqual(await commentsOf(await empty.save()), []);
});
