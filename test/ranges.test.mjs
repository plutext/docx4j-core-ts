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
  // case 26 (CC, its content deleted under tracking): Word fills the emptied control with five spaces, this package
  // leaves it empty; the paragraph is the same otherwise
  assert.deepEqual(differing, [26]);
  assert.deepEqual(word.get(26).map((p) => p.text), ['     ']);
  assert.deepEqual(mine.get(26), word.get(26).map((p) => ({ ...p, text: '' })));
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
    const { marks: ourMarks, ...mine } = ours.reads[kind];
    const { marks, ...theirs } = word.reads[kind];
    // ^p within the range is phase 3's: no hit here yet, where Word found each mark
    assert.deepEqual(ourMarks, [], kind);
    assert.ok(marks.length > 0 || kind.startsWith('TC') || kind === 'CELL', kind);
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
  // a paragraph's font reads its first run here (CR-002), where Office JS reads null where its runs differ
  const paragraphOf = ({ bold: b, italic, underline, words, ...rest }) => rest;
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
    // a paragraph after a nested table in a cell: Word's style 'Other' again (cases 89, 91, 94)
    for (const [k, p] of (expected.paragraphs ?? []).entries()) {
      if (p.styleBuiltIn === 'Other') {
        assert.ok([89, 91, 94].includes(n), `${n}`);
        assert.deepEqual([mine.paragraphs[k].styleBuiltIn, mine.paragraphs[k].style], ['ListParagraph', 'List Paragraph'], `${n}`);
        expected.paragraphs[k] = { ...p, style: 'List Paragraph', styleBuiltIn: 'ListParagraph' };
      }
    }
    // cases 13 and 26, a control emptied: Word marks it w:showingPlcHdr over a run of five spaces, and reads no
    // text there, where this package reads the spaces (CR-014 section 9.4)
    if (n === 13 || n === 26) {
      expected.text = expected.text.replace('\\u000d\\u000d', '\\u000d     \\u000d');
      expected.paragraphs[1].text = '     ';
      expected.contentControls[0].text = '     ';
    }
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

/** check52Bytes' paragraph holding a word, and a range over that word. */
const wordIn = (pkg, word) => pkg.body.search(word, { matchCase: true, matchWholeWord: true })[0];

test('CR-014 phase 1: a range across paragraphs, from expandTo, reads every paragraph and refuses every write until phase 2', async () => {
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
  // every write refused before anything changes
  const before = pkg.body.text;
  for (const write of [() => range.delete(), () => range.insertText('X', 'End'), () => { range.text = 'X'; }, () => { range.font.bold = true; },
    () => range.insertParagraph('X', 'After'), () => range.replaceText('3', '4'), () => range.insertContentControl(), () => { range.style = 'Normal'; }]) {
    assert.throws(write, NotSupportedError);
  }
  await assert.rejects(range.insertComment('X'), NotSupportedError);
  assert.equal(pkg.body.text, before);
  // and the one w:p a fragment of it would be (toApiScript, the shim's getOoxml)
  await assert.rejects(Word.run(pkg, async (context) => {
    const hit = context.document.body.search('HEAD3B', { matchCase: true })[0];
    hit.expandTo(context.document.body.search('BODY3A', { matchCase: true })[0]).getOoxml();
  }), NotSupportedError);
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
    // a colour reads '' for automatic and resolves theme colours another way here (Font's, CR-002 and CR-008): compared below where it is CR-014's
    const keys = [...new Set([...Object.keys(a), ...Object.keys(b)])].filter((k) => k !== 'font.color' && JSON.stringify(a[k]) !== JSON.stringify(b[k]));
    if (keys.length > 0) differing[theirs.n] = keys;
  }
  const changeRead = (n) => [`paragraphs.${n}.changes.0.type`, `paragraphs.${n}.changes.0.text`, `paragraphs.${n}.changes.0.range`, `paragraphs.${n}.changes.0.rangeParagraphs`,
    `paragraphs.${n}.wholeChanges.0.type`, `paragraphs.${n}.wholeChanges.0.text`, `paragraphs.${n}.wholeChanges.0.range`, `paragraphs.${n}.wholeChanges.0.rangeParagraphs`];
  const placeholder = ['control.text', 'control.placeholderText', 'control.paragraph', 'control.whole', 'showingPlaceholderText'];
  const styleOther = ['paragraph.style', 'paragraph.styleBuiltIn', 'whole.style', 'whole.styleBuiltIn'];
  assert.deepEqual(differing, {
    // a single paragraph's font reads its first run (CR-002), where Office JS reads '' or null where its runs differ
    8: ['font.italic'], 31: ['font.highlightColor'],
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
    // body.text in Simple Markup joins paragraphs at a deleted mark, drops a deleted paragraph and reads a control
    // showing its placeholder as empty (CR-012's readers)
    33: ['text.length', 'Whole.length', 'Content.length'],
    // a control showing its placeholder reads as empty in Word (CR-012's), and showingPlaceholderText is not implemented
    34: placeholder, 35: placeholder,
  });
  // the colour across two paragraphs whose runs' colours differ: '' (MixedFont), as Word's
  assert.equal(flat(ours.cases[0].reads)['font.color'], flat(word.cases[0].reads)['font.color']);
  // the changes' text is All Markup's
  for (const n of [25, 28, 30]) assert.equal(ours.cases[n - 1].reads.region[0].text, all.cases[n - 1].reads.region[0].text, `case ${n}`);
});
