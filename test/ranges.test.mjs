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
  // a write phase 2 doesn't bring is refused before anything changes: a control or a style (phase 3 brings
  // comments and controls on such a range)
  const before = pkg.body.text;
  for (const write of [() => range.insertContentControl(), () => { range.style = 'Normal'; }, () => { range.hyperlink = 'https://example.com'; }]) {
    assert.throws(write, NotSupportedError);
  }
  await assert.rejects(range.insertComment('X'), NotSupportedError);
  await assert.rejects(range.insertOoxml('<w:p/>', 'After'), NotSupportedError);
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

// ------------------------------------------------------------------------------------------- phase 2, the writes

/** Check 52's snippet in its 'run' mode, through the Word shim on its own file, once for the tests below. */
let phase2;
const phase2Run = () => (phase2 ??= (async () => {
  const { runCheck } = await snippet();
  const pkg = await loaded(await check52Bytes());
  const result = await runCheck((f) => Word.run(pkg, f), 'run');
  return { result, saved: await pkg.save() };
})());

/** Cases 86 to 95 are CR-002 section 55's (Word's table rule), which is not implemented: phase 2's tests stop at 85. */
const RANGE_CASES = 85;

test("CR-014 phase 2: check 52's writes, run through the Word shim on the same file, leave each region as Word 365's run did", async () => {
  const { result: ours } = await phase2Run();
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check52', '52-ranges-word365-run.json'), 'utf8'));
  // a paragraph's font reads its first run here (CR-002); under tracking Word's text is All Markup's, deleted text
  // shown, so a tracked case's paragraphs are compared without their text, and its text, tables and controls not
  const paragraph = ({ bold, italic, underline, words, ...rest }) => rest;
  const view = (c) => {
    const r = c.region;
    const out = { error: c.error, result: c.result, range: c.range, comments: r.comments };
    // case 1 is read only, but its content holds tracked changes from the start
    const markup = c.tracked || c.n === 1;
    out.paragraphs = Array.isArray(r.paragraphs) ? r.paragraphs.map((p) => (markup ? { ...paragraph(p), text: undefined } : paragraph(p))) : r.paragraphs;
    out.changes = Array.isArray(r.trackedChanges) ? r.trackedChanges.map((x) => x.type) : r.trackedChanges;
    if (!markup) Object.assign(out, { text: r.text, tables: r.tables, controls: r.contentControls });
    return JSON.parse(JSON.stringify(out));
  };
  const differing = {};
  for (const [i, theirs] of word.cases.entries()) {
    if (theirs.n > RANGE_CASES) break;
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
    // insertComment on a range across paragraphs is phase 3's; and the snippet calls getRange() at once on what it
    // returns, which is asynchronous here, so the one-paragraph case (83) fails in the script after the comment is in
    ...each([79, 80, 81, 82, 84, 85], ['error', 'comments', 'result']),
    83: ['error', 'result'],
  });
  // Word's refusals are this package's: a range starting or ending inside a table, deleted or replaced, tracked or not
  const refused = ours.cases.filter((c) => c.error === 'error: GeneralException').map((c) => c.n);
  assert.deepEqual(refused, [9, 10, 22, 23, 34, 35, 44, 45]);
});

test("CR-014 phase 2: Accept All after check 52's writes through the shim gives Word 365's Accept All after its run, region for region", async () => {
  const { runCheck } = await snippet();
  const pkg = await loaded((await phase2Run()).saved);
  await pkg.acceptAllRevisions();
  const back = await loaded(await pkg.save());
  const ours = await runCheck((f) => Word.run(back, f), 'report');
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check52', '52-ranges-word365-accepted.json'), 'utf8'));
  const paragraph = ({ bold, italic, underline, words, ...rest }) => rest;
  const differing = [];
  for (const [i, theirs] of word.cases.entries()) {
    if (theirs.n > RANGE_CASES) break;
    // Range.footnotes is not implemented, and a paragraph's font reads its first run here (CR-002)
    const view = ({ footnotes, ...r }) => ({ ...r, paragraphs: Array.isArray(r.paragraphs) ? r.paragraphs.map(paragraph) : r.paragraphs });
    if (JSON.stringify(view(ours.cases[i].region)) !== JSON.stringify(view(theirs.region))) differing.push(theirs.n);
  }
  // every tracked deletion and replacement, accepted, is Word's; what differs is listed in the test above:
  // getTextRanges (the bold cases), Word's extra empty paragraph in a control, comments (phase 3)
  assert.deepEqual(differing, [49, 50, 51, 52, 53, 54, 55, 56, 60, 66, 79, 80, 81, 82, 84, 85]);
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
  const differing = [...word.keys()].filter((n) => n <= RANGE_CASES && mine.get(n).join('\n') !== word.get(n).join('\n'));
  // 1, read only: Word's save gave the comment reference's run w:u and w:sz of its own; 13: Word's filler for an
  // emptied control, a bold run of five spaces (section 9.4); 60 and 66: Word's extra
  // empty paragraph in a control; 67 and 68: Word gives a paragraph's mark its last run's formatting (italic) when
  // insertParagraph('After') goes after it, not followed (CR-014 section 10); 79 to 85: insertComment, phase 3's
  assert.deepEqual(differing, [1, 13, 60, 66, 67, 68, 79, 80, 81, 82, 83, 84, 85]);
  assert.deepEqual(word.get(13), ['SDT', '  P[ListParagraph/num|] "     "{b}']);
  assert.deepEqual(mine.get(13), ['SDT', '  P[ListParagraph/num|] ']);
  for (const n of [67, 68]) {
    assert.deepEqual(mine.get(n).map((line) => line.replace(/^P\[center\|\] "BODY/, 'P[center|i] "BODY')), word.get(n), `case ${n}`);
  }
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
  // into a block content control from outside it: unmeasured, refused
  const control = body.contentControls.find((c) => c.tag === 'c52-13');
  const into = wordIn(pkg, 'BEGIN13').expandTo(control.paragraphs[0].getRange('Start'));
  assert.throws(() => into.delete(), NotSupportedError);
  // a range holding a whole block control: refused under tracking, where Word's markup for it isn't measured
  // (CR-014 section 10.3); untracked, the control goes as a table does and the paragraphs either side join
  const around = wordIn(pkg, 'BEGIN26').expandTo(wordIn(pkg, 'END26'));
  pkg.changeTrackingMode = 'TrackAll';
  assert.throws(() => around.delete(), (e) => e instanceof NotSupportedError && /holds a whole sdt/.test(e.message));
  pkg.changeTrackingMode = 'Off';
  assert.equal(body.text, before);
  const [controls, paragraphs] = [body.contentControls.length, body.paragraphs.length];
  around.delete();
  // BEGIN26, the control's three paragraphs and END26 become one empty paragraph
  assert.deepEqual([body.contentControls.length, body.paragraphs.length], [controls - 1, paragraphs - 4]);
  assert.equal(body.paragraphs.filter((p) => /BEGIN26|END26/.test(p.text)).length, 0);
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
