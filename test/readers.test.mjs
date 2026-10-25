// CR-012 phase 4, the readers (D3 to D6, section 6.2, section 12): Paragraph.text and search in Office JS's
// characters in Word's default view (Simple Markup, hidden text not shown), held to Word 365's answers on
// check 48's and 48b's files (test/README.md checks 48 and 48b); equations; text-box views at textbox:
// addresses, read-only (D4, D5); pkg.paragraphAt over notes and comments; outline() on the text model (D3).
// ED-009's branch editing is kept-branches.test.mjs's.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, strFromU8 } from 'fflate';
import { WordprocessingMLPackage, ZipPartStore, Texts, mathLetter } from '../dist/index.mjs';
import { check48Bytes } from './fixtures/check48/build.mjs';
import { check51Bytes } from './fixtures/check51/build.mjs';
import { fixture } from './helpers.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

/** Word's JSON writes every character outside printable ASCII as \uXXXX. */
const unescape = (s) => s.replace(/\\u([0-9a-fA-F]{4})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
const json = async (name) => JSON.parse(await readFile(join(here, 'fixtures', 'check48', name), 'utf8'));

/** Check 48b's five runs on check 48's file (A Simple Markup, B hidden shown, C All Markup, D No Markup, E Original), in order. */
async function views() {
  const t = await readFile(join(here, 'fixtures', 'check48', '48-text-views-word365.txt'), 'utf8');
  const out = [];
  for (let i = t.indexOf('{\n "label"'); i >= 0; i = t.indexOf('{\n "label"', i + 1)) {
    let depth = 0;
    let j = i;
    for (; j < t.length; j++) { if (t[j] === '{') depth++; else if (t[j] === '}' && --depth === 0) break; }
    out.push(JSON.parse(t.slice(i, j + 1)));
  }
  return out;
}

const load = async (which) => {
  const pkg = await WordprocessingMLPackage.load(await check48Bytes(which));
  await pkg.getBody();
  return pkg;
};

async function mini(bodyXml) {
  const make = await WordprocessingMLPackage.createPackage();
  make.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}" xmlns:m="${M}"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`);
  const pkg = await WordprocessingMLPackage.load(await make.save());
  await pkg.getBody();
  return pkg;
}

const xmlOf = async (pkg) => new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync('word/document.xml'));

// ------------------------------------------------------------- Paragraph.text and Body.text (D6, section 6.2)

test("check 48's and 48b's files: Paragraph.text is Office JS's in Word's default view, paragraph for paragraph", async () => {
  const [A] = await views();
  const pkg = await load('48');
  const paragraphs = pkg.body.paragraphs;
  assert.equal(paragraphs.length, A.paragraphs.length);
  paragraphs.forEach((p, i) => assert.equal(p.text, unescape(A.paragraphs[i].text), `check 48, paragraph ${i}`));
  // the characters, as check 48 measured them
  const text = (n) => paragraphs.find((p) => p.text.startsWith(`Case ${n}:`)).text.replace(/^Case \d+: [^.]*\. /, '');
  assert.equal(text(1).replace(/^.*?(FOOTA)/, '$1'), 'FOOTA\u0002FOOTB and ENDA\u0002ENDB');
  assert.equal(text(3), 'LINEA\u000BLINEB');
  assert.equal(text(4), 'PAGEA\u000CPAGEB');
  assert.equal(text(5), 'COLA\u000ECOLB');
  assert.equal(text(6), 'CRA\rCRB');
  assert.equal(text(7), 'TABA\tTABB0RIGHTB');
  assert.equal(text(8), 'HYPHA\u001EHYPHB\u001FHYPHC');
  assert.equal(text(9), 'SYMA(SYMB');
  assert.equal(text(10).slice(-8), 'HIDAHIDB', 'hidden text left out');
  assert.equal(text(11).slice(-20), 'REVAINSERTEDTEXTREVB', 'the accepted text: Simple Markup');
  assert.equal(text(16).slice(-8), 'PICAPICB');
  assert.ok(text(17).endsWith('MATHA \u{1D465}=1 MATHB'), "the equation as Word draws it, its x Unicode's mathematical italic");
  // 48b's take 2, in Word's defaults: every symbol `(`, and a run hidden through its character style left out
  const take2 = await json('48b-text-word365.json');
  const b = await load('48b');
  b.body.paragraphs.forEach((p, i) => assert.equal(p.text, unescape(take2.paragraphs[i].text), `check 48b, paragraph ${i}`));
});

test("check 48: a paragraph's 'Whole' range ends with its mark, \\r, or \\t for a cell's; Body.text is body.text's", async () => {
  const word = await json('48-text-word365.json');
  const pkg = await load('48');
  // check 48 ran in All Markup: only case 11 reads its deleted text there, and Simple Markup is the rule
  const simple = (s) => unescape(s).replace('REVADELETEDTEXTINSERTEDTEXTREVB', 'REVAINSERTEDTEXTREVB');
  pkg.body.paragraphs.forEach((p, i) => {
    assert.equal(p.getRange('Whole').text, simple(word.paragraphs[i].whole), `paragraph ${i}`);
    assert.equal(p.getRange('Content').text, p.text);
  });
  // its table's row ended \t\t in All Markup with hidden text shown; Word's default view ends a row \r\n (checks 30,
  // 31 and 51)
  assert.equal(pkg.body.text, simple(word.bodyText).replace('CELLA\tCELLB\t\tAFTERTABLE', 'CELLA\tCELLB\r\nAFTERTABLE'));
  assert.equal(pkg.body.getText({ view: 'original' }).includes('REVADELETEDTEXTREVB'), true);
  // two paragraphs in a cell: the first's mark \r, the last's the cell's end
  const two = await mini('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p><w:r><w:t>z</w:t></w:r></w:p>');
  assert.equal(two.body.text, 'a\rb\t\r\nz');
  assert.deepEqual(two.body.paragraphs.map((p) => p.getRange().text), ['a\r', 'b\t', '\t', 'z\r']);
});

test("check 51's file: cells, controls, comments and notes of several paragraphs, a note's own mark and the equations read as Word 365's", async () => {
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check51', '51-text-word365.json'), 'utf8'));
  const pkg = await WordprocessingMLPackage.load(await check51Bytes());
  const body = await pkg.getBody();
  // every paragraph, its whole range, and body.text: the equations' letters, minus signs and structures among them
  assert.equal(body.paragraphs.length, word.paragraphs.length);
  body.paragraphs.forEach((p, i) => {
    assert.equal(p.text, unescape(word.paragraphs[i].text), `paragraph ${i}`);
    assert.equal(p.getRange('Whole').text, unescape(word.paragraphs[i].whole), `paragraph ${i}, whole`);
  });
  assert.equal(body.text, unescape(word.bodyText));
  // the joiner is the paragraph mark, \r, as body.text has it
  const [table] = body.tables;
  assert.deepEqual(table.values, word.tables[0].values.map((row) => row.map(unescape)));
  table.rows[0].cells.forEach((cell, i) => {
    const measured = word.tables[0].cells[i];
    assert.deepEqual([cell.value, cell.body.text, cell.paragraphs.map((p) => p.text)], [unescape(measured.value), unescape(measured.bodyText), measured.paragraphs.map(unescape)]);
  });
  const [block, inline] = body.contentControls;
  assert.equal(block.text, unescape(word.contentControls[0].text));
  assert.deepEqual(inline.text, unescape(word.contentControls[1].text));
  assert.deepEqual([inline.getRange('Whole').text, inline.getRange('Content').text], [word.contentControls[1].whole, word.contentControls[1].content]);
  // a block control of several paragraphs: its range reads as Office JS's, all of them, one Range across them
  // (CR-014 phase 1; CR-012 section 17.12's SpanningRange until then); a hyperlink across its paragraphs refuses
  const range = block.getRange();
  assert.deepEqual([range.text, block.getRange('Content').text], [unescape(word.contentControls[0].whole), unescape(word.contentControls[0].content)]);
  assert.deepEqual(range.paragraphs.map((p) => p.text), word.contentControls[0].paragraphs);
  assert.deepEqual(range.search('CCP').map((r) => r.text), ['CCP', 'CCP', 'CCP']);
  assert.equal(range.font.bold, false, 'a read');
  assert.deepEqual([block.getRange('End').paragraph.text, block.getRange('End').start], ['CCP3', 4], "its end is the last paragraph's");
  const before = await pkg.getMainDocumentPart().getXml();
  assert.throws(() => { range.hyperlink = 'https://example.com'; }, (e) => e.name === 'NotSupportedError' && /CR-014/.test(e.message));
  assert.equal(await pkg.getMainDocumentPart().getXml(), before, 'nothing changed');
  // the control's own verbs act on all of it
  block.insertText('one', 'Replace');
  assert.equal(block.text, 'one');
  assert.equal((await body.getComments())[0].content, unescape(word.comments[0].content));
  // a note's own mark reads U+0002, as its reference does
  for (const [kind, measured] of [['footnote', word.footnotes[0]], ['endnote', word.endnotes[0]]]) {
    const first = await pkg.paragraphAt(`${kind}:1/0`);
    const paragraphs = first.parentBody.paragraphs;
    assert.deepEqual(paragraphs.map((p) => p.text), measured.paragraphs.map(unescape), kind);
    assert.equal(first.parentBody.text, unescape(measured.bodyText), kind);
    assert.equal(paragraphs[paragraphs.length - 1].getRange('Whole').text, unescape(measured.bodyWhole).split('\r').slice(-2).join('\r'), kind);
  }
});

test('hidden text: left out of Paragraph.text, a run\'s own w:vanish or its style\'s; an edit across it takes it; an insertion beside it does not go into it', async () => {
  const pkg = await mini('<w:p><w:r><w:t xml:space="preserve">see </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN</w:t></w:r><w:r><w:t xml:space="preserve"> here</w:t></w:r></w:p>');
  const p = pkg.body.paragraphs[0];
  assert.equal(p.text, 'see  here');
  assert.equal(p.getText({ view: 'original' }), 'see  here');
  assert.equal(p.search('HIDDEN').length, 0, "Word's default view: a hidden word isn't found");
  assert.equal(p.search('see  here').length, 1, 'and a query matches across it');
  p.insertText('X', 'End');
  p.getRange('Start').insertText('Y', 'Start');
  p.search(' here')[0].insertText('!', 'Start');
  assert.match(await xmlOf(pkg), /<w:t>Ysee !?<\/w:t>|<w:t xml:space="preserve">Ysee !?<\/w:t>/);
  assert.match(await xmlOf(pkg), /<w:vanish\/><\/w:rPr><w:t>HIDDEN<\/w:t>/, 'the hidden run as it was');
  p.search('see ! here')[0].delete();
  assert.equal(p.text, 'YX');
  assert.doesNotMatch(await xmlOf(pkg), /HIDDEN/, 'the hidden text inside the span went with it');
});

// ------------------------------------------------------------- search (section 6.2 question 3)

test("search: Word's own queries give Word's hits; Find codes, a picture's place, the hyphens", async () => {
  const [A] = await views();
  const take2 = await json('48b-text-word365.json');
  const check48 = await json('48-text-word365.json');
  for (const [label, which, searches, viewDependent] of [
    ['48b run A', '48', A.searches, []],
    ['48b take 2', '48b', take2.searches, []],
    // check 48 ran in All Markup with hidden text shown: the hidden and deleted words read otherwise there
    ['check 48', '48', check48.searches, ['HIDAHIDB', 'HIDDENTEXT', 'HIDAHIDDENTEXTHIDB', 'DELETEDTEXT', 'REVAINSERTEDTEXTREVB', 'REVADELETEDTEXTINSERTEDTEXTREVB']],
  ]) {
    const pkg = await load(which);
    const differ = [];
    for (const [query, hits] of Object.entries(searches)) {
      const mine = pkg.body.search(query).map((r) => r.text);
      if (JSON.stringify(mine) !== JSON.stringify(hits.map(unescape))) differ.push(query);
    }
    assert.deepEqual(differ, viewDependent, label);
  }
  // ^f and ^e both read U+0002, each matching its own reference
  const pkg = await load('48');
  assert.equal(pkg.body.search('FOOTA^eFOOTB').length, 0);
  assert.equal(pkg.body.search('ENDA^fENDB').length, 0);
  // a range's search reads its own span
  const p = pkg.body.paragraphs.find((q) => q.text.includes('HYPHA'));
  const range = p.search('HYPHB^-HYPHC')[0];
  assert.equal(range.text, 'HYPHB\u001FHYPHC');
  assert.equal(range.search('HYPHBHYPHC').length, 1);
  assert.equal(range.search('HYPHA').length, 0);
  // a regular expression (this package's extension) reads Paragraph.text itself
  assert.equal(p.search('HYPHB\\u001FHYPHC', { matchRegExp: true }).length, 1);
});

// ------------------------------------------------------------- equations

test("an equation reads as Word draws it, an edit takes it whole or not at all, and a tracked deletion takes it in Word's form", async () => {
  assert.equal(mathLetter('x', undefined, undefined), '\u{1D465}');
  assert.equal(mathLetter('h', undefined, undefined), 'ℎ', "italic h is the BMP's Planck constant");
  assert.equal(mathLetter('2', undefined, undefined), '2');
  assert.equal(mathLetter('x', 'p', undefined), 'x');
  assert.equal(mathLetter('x', 'b', undefined), '\u{1D431}');
  assert.equal(mathLetter('R', undefined, 'double-struck'), 'ℝ');
  assert.equal(mathLetter('α', undefined, undefined), '\u{1D6FC}', 'small Greek italic');
  assert.equal(mathLetter('Δ', undefined, undefined), '\u{1D6E5}', 'capital Greek italic too (check 51)');
  assert.equal(mathLetter('2', 'bi', undefined), '\u{1D7D0}', 'bold digits under bold italic (check 51)');
  const math = '<m:oMath><m:r><m:t>x</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>+1</m:t></m:r></m:oMath>';
  const make = () => mini(`<w:p><w:r><w:t xml:space="preserve">a </w:t></w:r>${math}<w:r><w:t xml:space="preserve"> b</w:t></w:r></w:p>`);
  const pkg = await make();
  const p = pkg.body.paragraphs[0];
  assert.equal(p.text, 'a \u{1D465}+1 b');
  assert.throws(() => p.search('+1 b')[0].delete(), (e) => e.name === 'NotSupportedError' && /cuts the equation/.test(e.message));
  assert.throws(() => p.splice(3, 3, 'z'), (e) => e.name === 'NotSupportedError' && /cuts the equation/.test(e.message), 'no insertion inside it');
  p.search('b')[0].font.bold = true;
  p.getRange('Whole').font.italic = true;
  assert.match(await xmlOf(pkg), /<m:oMath><m:r><m:t>x<\/m:t><\/m:r>/, "a format leaves the math runs' own");
  p.search('a \u{1D465}+1 ')[0].delete();
  assert.equal(p.text, 'b');
  assert.doesNotMatch(await xmlOf(pkg), /oMath/, 'taken whole');
  // tracked, a w:del in each math run, its m:rPr in it too (CR-002 sections 51 and 62.2 D2, check 57); accepted, it goes
  const tracked = await make();
  tracked.changeTrackingMode = 'TrackAll';
  tracked.body.paragraphs[0].getRange('Content').delete();
  assert.equal(tracked.body.paragraphs[0].text, '');
  assert.match(await xmlOf(tracked), /<m:oMath><m:r><w:del [^>]*><m:t>x<\/m:t><\/w:del><\/m:r><m:r><w:del [^>]*><m:rPr><m:sty m:val="p"\/><\/m:rPr><m:t>\+1<\/m:t><\/w:del><\/m:r><\/m:oMath>/);
  await tracked.acceptAllRevisions();
  assert.doesNotMatch(await xmlOf(tracked), /oMath/);
});

// ------------------------------------------------------------- text boxes (D4, D5)

/** Every .docx under a directory whose main part holds a text box. */
async function withTextBoxes(dir) {
  const out = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) out.push(...await withTextBoxes(path));
    else if (entry.name.endsWith('.docx')) {
      try {
        const files = unzipSync(await readFile(path), { filter: (f) => f.name === 'word/document.xml' });
        if (files['word/document.xml'] && strFromU8(files['word/document.xml']).includes('txbxContent')) out.push(path);
      } catch { /* not a zip */ }
    }
  }
  return out;
}

test("textbox: addresses resolve to the text model's text-box paragraphs, and addressOf gives them back (D4)", async () => {
  const files = [...await withTextBoxes(join(here, 'fixtures')), ...await withTextBoxes(join(here, 'text-extraction')), ...await withTextBoxes(join(here, 'mc-fun'))];
  let agreed = 0;
  const missed = [];
  for (const file of files) {
    let pkg;
    try { pkg = await WordprocessingMLPackage.load(await readFile(file)); await pkg.getBody(); } catch { continue; }
    let result;
    try { result = await Texts.of(pkg); } catch { continue; }
    const visit = async (blocks) => {
      for (const block of blocks) {
        if (block.kind === 'TABLE') { for (const row of block.rows) for (const cell of row.cells) await visit(cell.blocks); continue; }
        if (!block.textbox || !block.source || !block.address.startsWith('body')) continue;     // WordArt has no paragraph
        const p = await pkg.paragraphAt(block.address);
        if (p?.p === block.source && pkg.body.addressOf(p) === block.address && p.inTextBox) agreed++;
        else missed.push(`${relative(here, file)} ${block.address}`);
      }
    };
    await visit(result.blocks);
  }
  // every one, a box in a block-level mc:AlternateContent's selected branch too, since CR-013 phase 2
  assert.ok(agreed >= 132, `${agreed} text-box paragraphs agree`);
  assert.deepEqual(missed, []);
});

test("a text box's paragraph reads as any does and refuses every write (D5); an address naming nothing is undefined, never an error", async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('parity/numbering-stories.docx'));
  const body = await pkg.getBody();
  const p = body.paragraphAt('body/7/textbox:0/1');
  assert.equal(p.text, 'P21 text box item');
  assert.equal(p.inTextBox, true);
  assert.equal(body.paragraphs.includes(p), false);
  assert.equal(body.paragraphs.some((q) => q.p === p.p), false, "a text box's paragraphs aren't the body's (check 42)");
  assert.equal(typeof p.font.bold, 'boolean');
  assert.equal(p.search('item').length, 1);
  const before = await xmlOf(pkg);
  for (const [name, write] of Object.entries({
    'text': () => { p.text = 'x'; }, insertText: () => p.insertText('x', 'End'), 'search hit replaced': () => p.search('item')[0].insertText('x', 'Replace'),
    'range delete': () => p.getRange().delete(), 'font': () => { p.search('item')[0].font.bold = true; }, 'paragraph font': () => { p.font.italic = true; },
    'alignment': () => { p.alignment = 'Centered'; }, 'style': () => { p.styleId = 'Heading1'; }, 'delete': () => p.delete(),
    'insertParagraph': () => p.insertParagraph('x', 'After'), 'insertBreak': () => p.insertBreak('Line', 'End'),
    'content control': () => p.insertContentControl(), 'range content control': () => p.search('item')[0].insertContentControl(),
    'hyperlink': () => { p.search('item')[0].hyperlink = 'https://example.com'; }, 'list': () => p.detachFromList(),
  })) {
    assert.throws(write, (e) => e.name === 'NotSupportedError' && /CR-012 D5/.test(e.message), name);
  }
  await assert.rejects(p.insertComment('x'), (e) => e.name === 'NotSupportedError');
  assert.equal(await xmlOf(pkg), before, 'nothing changed');
  // addresses that name nothing
  for (const address of ['body/7/textbox:1/0', 'body/7/textbox:x/0', 'body/7/textbox:-1/0', 'body/7/textbox:0', 'body/textbox:0/0', 'textbox:0/0', 'body/0/textbox:0/0', 'body/7/textbox:0/9', 'body/7//1']) {
    assert.equal(body.elementAt(address), undefined, address);
    assert.equal(await pkg.paragraphAt(address), undefined, address);
  }
});

test("pkg.paragraphAt resolves footnote:, endnote: and comment: addresses, as the text model and Comment write them (D4)", async () => {
  const pkg = await load('48');
  assert.equal((await pkg.paragraphAt('footnote:1/0')).text, '\u0002 The footnote of case 1.', "the note's own mark U+0002, as its reference reads");
  assert.equal((await pkg.paragraphAt('endnote:1/0')).text, '\u0002 The endnote of case 1.');
  assert.equal((await pkg.paragraphAt('comment:0/0')).text, 'The comment of case 2.');
  const [comment] = await pkg.body.getComments();
  assert.equal(comment.paragraphs[0].p, (await pkg.paragraphAt('comment:0/0')).p);
  for (const address of ['footnote:9/0', 'endnote:1/5', 'comment:7/0', 'footnote:/0']) assert.equal(await pkg.paragraphAt(address), undefined, address);
  const result = await Texts.of(pkg);
  for (const block of result.blocks.filter((b) => b.story === 'FOOTNOTE' || b.story === 'ENDNOTE')) {
    assert.equal((await pkg.paragraphAt(block.address)).p, block.source, block.address);
  }
});

// ------------------------------------------------------------- outline (D3)

test('outline() reads the text model: its text, text boxes after their anchor at textbox: addresses, flagged; the options; the package unchanged', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('parity/numbering-stories.docx'));
  const body = await pkg.getBody();
  const outline = await body.outline();
  const i = outline.paragraphs.findIndex((p) => p.address === 'body/7');
  assert.deepEqual(outline.paragraphs.slice(i, i + 4).map((p) => [p.address, p.text, p.textbox ?? false]), [
    ['body/7', 'P23 A text box: ', false], ['body/7/textbox:0/0', 'P20 text box item', true],
    ['body/7/textbox:0/1', 'P21 text box item', true], ['body/7/textbox:0/2', 'P22 text box item', true],
  ]);
  for (const entry of outline.paragraphs) assert.ok(body.elementAt(entry.address), entry.address);
  // in a table, a box follows the table (REQ-518)
  const cells = await WordprocessingMLPackage.load(await fixture('parity/tables-textbox-in-cell.docx'));
  const boxes = (await (await cells.getBody()).outline()).paragraphs.filter((p) => p.textbox).map((p) => p.address);
  assert.deepEqual(boxes, ['body/0/0/0/1/textbox:0/0', 'body/0/1/0/1/textbox:0/0', 'body/1/textbox:0/0']);
  // the model's text and options: hidden text out unless asked for, the view asked for
  const doc = await load('48');
  const before = new ZipPartStore(await doc.save()).loadSync('word/footnotes.xml');
  const context = await Texts.contextOf(doc);
  const text = async (options, ctx) => (await doc.body.outline(options, ctx)).paragraphs.map((p) => p.text);
  const accepted = await text({}, context);
  assert.ok(accepted.some((t) => t.endsWith('HIDAHIDB')));
  assert.ok(accepted.some((t) => t.endsWith('REVAINSERTEDTEXTREVB')));
  assert.ok((await text({ hidden: true })).some((t) => t.endsWith('HIDAHIDDENTEXTHIDB')));
  assert.ok((await text({ view: 'original' })).some((t) => t.endsWith('REVADELETEDTEXTREVB')));
  await assert.rejects(doc.body.outline({ view: 'original' }, context), /TextContext read in the accepted view/);
  assert.deepEqual(new ZipPartStore(await doc.save()).loadSync('word/footnotes.xml'), before, 'the notes part read privately (REQ-204)');
  const whole = await doc.outline();
  assert.deepEqual(whole.paragraphs.map((p) => p.text), accepted);
  assert.deepEqual([whole.headers.length, whole.footers.length], [0, 0]);
});

test("outline(options, context, { blocks }) gives a block's entries as the whole outline does: joins, a table of contents, every view (section 17.13)", async () => {
  const read = async (path) => {
    const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile(join(here, path))));
    return { pkg, body: await pkg.getBody() };
  };
  // the files where a block read alone, with no more, differs from the whole: paragraphs a mark joins, a
  // table of contents begun in an earlier block, moves
  for (const path of ['text-extraction/revisions/cr043/cr043-base.docx', 'text-extraction/suite/phase0/extraction-probe-structure.docx',
    'text-extraction/suite/phase0/extraction-probe.docx', 'text-extraction/revisions/editor/revisions/revisions-word2010.docx',
    'text-extraction/revisions/editor/review/review-for-word.docx', 'text-extraction/revisions/python/artefact8/8-tracked-changes.docx']) {
    const { pkg, body } = await read(path);
    for (const options of [{}, { view: 'original' }, { view: 'markup' }, { toc: true }, { field_codes: true }]) {
      const context = await Texts.contextOf(pkg);
      const whole = await body.outline(options, context);
      const one = { paragraphs: [], tables: [] };
      for (let i = body.content.length - 1; i >= 0; i--) {                 // backwards: no block's reading leans on the one before
        const o = await body.outline(options, context, { blocks: [i] });
        one.paragraphs.unshift(...o.paragraphs);
        one.tables.unshift(...o.tables);
      }
      const label = `${path} ${JSON.stringify(options)}`;
      assert.deepEqual(one, whole, label);
      assert.deepEqual(await body.outline(options, context, { blocks: body.content.map((_, i) => i) }), whole, label);
    }
  }
  // a deleted mark: the entry holds both texts and names the second, which has none of its own
  const { pkg, body } = await read('text-extraction/suite/phase0/extraction-probe.docx');
  const context = await Texts.contextOf(pkg);
  assert.deepEqual(await body.outline({}, context, { blocks: [24] }), { paragraphs: [{ address: 'body/24', style: 'Normal', text: 'TC11_FIRST TC11_SECOND', joined: ['body/25'] }], tables: [] });
  assert.deepEqual(await body.outline({}, context, { blocks: [25] }), { paragraphs: [], tables: [] });
  // the caller's context numbers no note: a reading of the notes part after it still has the whole order
  assert.equal(context.footnotes.size, 0);
  // a table of contents' entry, inside a field begun in an earlier block, is left out unless asked for (D12)
  const structure = await read('text-extraction/suite/phase0/extraction-probe-structure.docx');
  assert.deepEqual((await structure.body.outline({}, undefined, { blocks: [5] })).paragraphs, []);
  assert.deepEqual((await structure.body.outline({ toc: true }, undefined, { blocks: [5] })).paragraphs.map((p) => p.text), ['HEAD55_TWO\t1']);
  // blocks by view, in document order whatever the order asked; a block that is not a top-level one refused
  const paragraph = body.paragraphs.find((p) => p.container === body.content && p.text !== '');
  const at = body.content.indexOf(paragraph.element);
  const table = body.elementAt(`body/${body.content.findIndex((e) => e.name.localPart === 'tbl')}`);
  const byView = await body.outline({}, context, { blocks: [table, paragraph] });
  assert.deepEqual(byView, await body.outline({}, context, { blocks: [at, body.content.indexOf(table.element)] }));
  assert.equal(byView.paragraphs[0].address, `body/${at}`);
  await assert.rejects(body.outline({}, context, { blocks: [body.content.length] }), /has no block/);
  const inCell = body.paragraphs.find((p) => p.container !== body.content);
  await assert.rejects(body.outline({}, context, { blocks: [inCell] }), /Not one of body's top-level blocks/);
});

// ------------------------------------------------------------- defects found on the way

test('Range.font: a read splits no run, and a write formats the span alone, not the rest of its run', async () => {
  const pkg = await mini('<w:p><w:r><w:t>flat</w:t><w:tab/><w:t>really</w:t></w:r></w:p>');
  const p = pkg.body.paragraphs[0];
  const before = await xmlOf(pkg);
  assert.equal(p.search('lat')[0].font.bold, false);
  assert.equal(await xmlOf(pkg), before, 'a read changed nothing');
  p.search('lat')[0].font.bold = true;
  assert.match(await xmlOf(pkg), /<w:r><w:t>f<\/w:t><\/w:r><w:r><w:rPr><w:b\/><w:bCs\/><\/w:rPr><w:t>lat<\/w:t><\/w:r><w:r><w:tab\/><w:t>really<\/w:t><\/w:r>/);
});

test("Paragraph.insertContentControl links the control to the paragraph's old parent, not into itself", async () => {
  const pkg = await mini('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
  const control = pkg.body.paragraphs[0].insertContentControl();
  assert.equal(control.sdt.PARENT, pkg.body.container);
  assert.equal(pkg.body.addressOf(control.paragraphs[0]), 'body/0/0');
});

// --- check 58 (2026-10-23): what CR-002 section 58, CR-012 sections 18 and 19 and CR-002 section 32's item 1 left to
// measure. Word runs the snippet on the file, so the two must agree; until Word's run is back the shim's answers are
// what this package gives.

const snippet58 = async () => new Function(`${await readFile(join(here, 'fixtures', 'check58', 'check58-script-lab.js'), 'utf8')}\nreturn CHECK58;`)();

test("check 58's file and snippet agree: every case parses to the case it was built as, each word a range looks for and each tag once", async () => {
  const { parseRegion, QUERIES } = await snippet58();
  const { check58Bytes, check58Cases } = await import('./fixtures/check58/build.mjs');
  const pkg = await WordprocessingMLPackage.load(await check58Bytes());
  const body = await pkg.getBody();
  const regions = body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null);
  const cases = check58Cases();
  assert.equal(regions.length, cases.length);
  assert.deepEqual(regions.map((r) => r.n), cases.map((c) => c.n));
  const ops = new Set(['changes', 'range', 'cell', 'placeholder', 'search', 'accept', 'reject', 'fill']);
  for (const r of regions) {
    assert.ok(ops.has(r.op), `case ${r.n}: ${r.op}`);
    if (r.op === 'range') for (const word of r.args) assert.equal(body.search(word, { matchCase: true, matchWholeWord: true }).length, 1, `case ${r.n}: ${word}`);
    if (r.op === 'placeholder' || r.op === 'fill') assert.equal(body.contentControls.filter((c) => c.tag === r.args[0]).length, 1, `case ${r.n}: ${r.args[0]}`);
    if (r.op === 'search') assert.equal(typeof QUERIES[r.args[0]], 'function', `case ${r.n}`);
    if (r.op === 'accept' || r.op === 'reject') assert.ok(['1', '2'].includes(r.args[0]), `case ${r.n}`);
    if (r.op === 'fill') assert.ok(['replace', 'end'].includes(r.args[1]), `case ${r.n}`);
  }
  // the controls' ids are distinct, and every control shows its placeholder
  const ids = body.contentControls.map((c) => c.id);
  assert.equal(new Set(ids).size, ids.length);
  assert.equal(body.contentControls.length, 6);
});

test("check 58's three JSONs (Word 365, 2026-10-23) are of this file, and its save is Word's, as copied", async () => {
  const { check58Cases } = await import('./fixtures/check58/build.mjs');
  const { createHash } = await import('node:crypto');
  const cases = check58Cases();
  for (const [mode, label] of [['read-simple', 'Simple Markup'], ['read-all', 'All Markup'], ['edit', 'All Markup']]) {
    const word = JSON.parse(await readFile(join(here, 'fixtures', 'check58', `58-measures-word365-${mode}.json`), 'utf8'));
    assert.deepEqual([word.label, word.mode], [label, mode.startsWith('read') ? 'read' : 'edit']);
    const expected = mode === 'edit' ? cases.filter((c) => /^(accept|reject|fill) /.test(c.spec)) : cases;
    assert.deepEqual(word.cases.map((c) => c.n), expected.map((c) => c.n), mode);
  }
  const manifest = JSON.parse(await readFile(join(here, 'check58', 'manifest.json'), 'utf8'));
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await readFile(join(here, 'check58', name))).digest('hex'), sha1, name);
});

test("check 58's snippet through the Word shim, held to Word 365's Simple Markup run and its Edit run, but for the differences listed", async () => {
  const { runCheck } = await snippet58();
  const { check58Bytes } = await import('./fixtures/check58/build.mjs');
  const { Word } = await import('../dist/office-js/index.mjs');
  const word = async (mode) => JSON.parse(await readFile(join(here, 'fixtures', 'check58', `58-measures-word365-${mode}.json`), 'utf8'));
  const pkg = await WordprocessingMLPackage.load(await check58Bytes());
  await pkg.getBody();
  const ours = { read: await runCheck((f) => Word.run(pkg, f), 'shim', 'read'), edit: await runCheck((f) => Word.run(pkg, f), 'shim', 'edit') };
  const theirs = { read: await word('read-simple'), edit: await word('edit') };
  // a listing as K1 pins it (CR-002 section 58.1): the region's changes and each paragraph's, its whole range's and its
  // content range's, by type. A change's text here is docx4j's, All Markup's, where Simple Markup reads a deletion's as ''
  const listing = (r) => ({ region: r.region.map((c) => c.type), paragraphs: r.paragraphs.map((p) => [p.changes, p.wholeChanges, p.contentChanges].map((l) => l.map((c) => c.type))) });
  // placeholderText reads null, as Word's, with no glossary part (CR-002 section 63, D8); Word's showingPlaceholderText came
  // back undefined, which JSON leaves out
  const control = (c) => c;
  const placeholder = (r) => ({ control: typeof r.control === 'object' ? control(r.control) : r.control });
  const views = {
    read: (c) => (['changes', 'accept', 'reject'].includes(c.op) ? listing(c.reads) : ['placeholder', 'fill'].includes(c.op) ? placeholder(c.reads) : c.reads),
    // the Edit run was in All Markup, so the region's text after it isn't compared; what was resolved and what is left are
    edit: (c) => (c.op === 'fill' ? { inserted: c.reads.inserted, after: placeholder(c.reads.after) }
      : { resolved: c.reads.resolved, changes: c.reads.after.changes.map((x) => [x.type, x.text]) }),
  };
  const differing = [];
  for (const mode of ['read', 'edit']) {
    assert.deepEqual(ours[mode].cases.map((c) => c.n), theirs[mode].cases.map((c) => c.n), mode);
    for (const [i, c] of theirs[mode].cases.entries()) {
      if (JSON.stringify(views[mode](c)) !== JSON.stringify(views[mode](ours[mode].cases[i]))) differing.push(`${mode} ${c.n}`);
    }
  }
  // K1: Word lists nothing in two deleted paragraphs in a row (2, and 3, whose region lists only the first author's
  // deletion), nor in a paragraph whose text alone is deleted (4); this package lists each in its paragraph. K3: Word's
  // search finds the placeholder's text, reading each hit as '' (12). K6: Word's classes match nothing in the emoji or
  // the Arabic word, where [!a-z ]@ finds them here (13); its mark, as Word's since K7
  assert.deepEqual(differing, ['read 2', 'read 3', 'read 4', 'read 12', 'read 13']);
  const wildcards = (r) => Object.fromEntries(Object.entries(r).filter(([q]) => q !== '[!a-z ]@'));
  assert.deepEqual(wildcards(ours.read.cases[12].reads), wildcards(theirs.read.cases[12].reads));
  assert.deepEqual(ours.read.cases[12].reads['[!a-z ]@'], [['\\ud83d\\ude00', '\\ud83d\\ude00', '\\u0643\\u064e\\u062a\\u064e\\u0628\\u064e', '\\u000d']]);
  const reads = (n) => ours.read.cases[n - 1].reads;
  assert.deepEqual(listing(reads(2)).paragraphs.slice(2, 4), [[['Deleted'], ['Deleted'], ['Deleted']], [[], [], []]]);
  assert.deepEqual(listing(theirs.read.cases[1].reads).paragraphs.slice(2, 4), [[[], [], []], [[], [], []]]);
  assert.deepEqual(reads(12)['Click or tap here to enter text.'], [[], [], ['Click or tap here to enter text.']]);
  assert.deepEqual(theirs.read.cases[11].reads['Click or tap here to enter text.'], [[''], [''], ['Click or tap here to enter text.']]);
  assert.deepEqual(reads(13)['?'], theirs.read.cases[12].reads['?'], '? a code point (K4), and the mark (K7)');
});

// --- check 59 (2026-10-23): where Word's wildcards match a paragraph mark (CR-002 section 32's K5)

test("check 59's file and snippet agree, and the snippet through the Word shim is Word 365's run: wildcards match a mark where ^p does (K7)", async () => {
  const { parseRegion, runCheck } = new Function(`${await readFile(join(here, 'fixtures', 'check59', 'check59-script-lab.js'), 'utf8')}\nreturn CHECK59;`)();
  const { check59Bytes, check59Cases } = await import('./fixtures/check59/build.mjs');
  const { Word } = await import('../dist/office-js/index.mjs');
  const pkg = await WordprocessingMLPackage.load(await check59Bytes());
  const body = await pkg.getBody();
  const regions = body.paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null);
  assert.deepEqual(regions, check59Cases().map((c) => ({ n: c.n, shape: c.shape })));
  const result = await runCheck((f) => Word.run(pkg, f), 'shim');
  assert.equal(result.cases.length, 5);
  assert.doesNotMatch(JSON.stringify(result), /error: /);
  // Word 365's run (2026-10-23), scope for scope: a paragraph's mark matched, never a cell's or a row's end, nor across a
  // table's edge, as `^p`'s; a control's range without its last paragraph's mark
  const word = JSON.parse(await readFile(join(here, 'fixtures', 'check59', '59-wildcard-marks-word365.json'), 'utf8'));
  assert.deepEqual(word.cases.map((c) => c.shape), result.cases.map((c) => c.shape));
  for (const [i, c] of word.cases.entries()) assert.deepEqual(result.cases[i].reads, c.reads, c.shape);
  assert.deepEqual(word.cases[1].reads.cell['?'], ['Q'], "a cell's end: no mark");
  assert.deepEqual(word.cases[3].reads.table['?'], ['Q', 'Z', 'W', 'X'], "cells' and rows' ends: no mark");
  assert.deepEqual(word.cases[4].reads.control['?'], ['Q', '\\u000d', 'Z'], "inside a control: its paragraphs' marks but the last, which its range doesn't hold");
  // a replacement where a hit holds a mark is refused before anything changes, as a plain ^p's
  const before = await body.getXml();
  assert.throws(() => body.replaceText('Q?Z', 'X', { matchWildcards: true }), /holds a paragraph mark/);
  assert.throws(() => body.paragraphs[2].replaceText('?', 'X', { matchWildcards: true }), /holds a paragraph mark/);
  assert.equal(await body.getXml(), before);
  assert.deepEqual([body.paragraphs[2].text, body.paragraphs[2].replaceText('Q', 'X', { matchWildcards: true }), body.paragraphs[2].text], ['Q', 1, 'X'], 'a hit holding no mark is replaced');
});

test("the matchers across marks are exported, for a caller's own segments (the editor's request, 2026-10-26; CR-014 section 15)", async () => {
  const root = await import('../dist/index.mjs');
  const model = await import('../dist/model/index.mjs');
  for (const name of ['wildcardMatchesAcross', 'paragraphMarkPieces', 'officeMatches', 'searchPattern']) {
    assert.equal(typeof root[name], 'function', name);
    assert.equal(model[name], root[name], name);
  }
  const { wildcardMatchesAcross, paragraphMarkPieces } = root;
  assert.deepEqual(paragraphMarkPieces('A^pB^p'), ['A', 'B', '']);
  assert.deepEqual(paragraphMarkPieces('A^^pB'), undefined, '^^ is an escaped caret');
  assert.deepEqual(paragraphMarkPieces('A^tB'), undefined);
  // segments of the caller's own, in officeMatches' shape: a caller's paragraphs, each followed by its mark
  const seg = (text, start = 0) => ({ text, start, item: { name: { localPart: 't' } } });
  const pieces = [
    { segments: [seg('XA'), seg('Q', 2)], from: 0, to: 3, length: 3, mark: true },
    { segments: [seg('ZB')], from: 0, to: 2, length: 2, mark: true },
    { segments: [seg('C')], from: 0, to: 1, length: 1, mark: false },
  ];
  const hits = wildcardMatchesAcross(pieces, 'Q?Z', {});
  assert.deepEqual(hits.map(({ first, start, last, end, mark }) => ({ first, start, last, end, mark })), [{ first: 0, start: 2, last: 1, end: 1, mark: false }]);
  assert.deepEqual(wildcardMatchesAcross(pieces, 'B^13', {}).map(({ first, start, last, end, mark }) => ({ first, start, last, end, mark })),
    [{ first: 1, start: 1, last: 1, end: 2, mark: true }], 'a hit ending with a mark holds it');
  // the same paragraphs searched by Body.search, whose wildcard search they are
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  for (const text of ['XAQ', 'ZB', 'C']) body.insertParagraph(text, 'End');
  const theirs = body.search('Q?Z', { matchWildcards: true });
  assert.deepEqual(theirs.map((r) => r.text), ['Q\rZ']);
  const own = body.paragraphs.map((p, j, all) => ({ segments: p.segments({ objects: true }), from: 0, to: p.text.length, length: p.text.length, mark: j < all.length - 1 }));
  assert.deepEqual(wildcardMatchesAcross(own, 'Q?Z', {}).map((m) => m.match[0]), ['Q\rZ']);
});

test("CR-002 section 64: a moved-from mark reads as nothing in Word's default view, as its runs do, and a moved-to mark is no mark in the original view (the editor's finding)", async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('revisions/revisions-word15.docx'));
  const body = await pkg.getBody();
  const [moveTo, moveFrom] = [body.paragraphs[14], body.paragraphs[16]];
  assert.ok(moveTo.element.value.pPr.rPr.moveTo && moveFrom.element.value.pPr.rPr.moveFrom, 'the fixture still holds the move');
  assert.deepEqual([moveFrom.text, moveFrom.markText, moveFrom.getRange('Whole').text], ['', '', '']);
  assert.equal(moveTo.markText, '\r');
  assert.match(body.text, /where it is\.\rThe third paragraph of the move/, 'no second mark where the paragraph moved from');
  assert.deepEqual(body.search('^p^p').map((r) => r.text), [], "until now the moved-from mark and the one before it read as ^p^p");
  assert.deepEqual(body.search('.^13T', { matchWildcards: true }).filter((r) => r.text.includes('\r\r')).map((r) => r.text), []);
  // the original view, Reject All's: the moved-from paragraph with its mark, the moved-to one without
  const original = body.getText({ view: 'original' });
  assert.equal(original.split('\r').length, body.text.split('\r').length, 'one mark for one, the move undone');
  assert.ok(original.includes(moveFrom.getText({ view: 'original' }) + '\r'));
  assert.equal(moveTo.getText({ view: 'original' }), '');
});

test("CR-002 section 64: a cell's paragraph finds its cell through nested row-level controls (a repeating section's items), and a wildcard search stays in the cell (the editor's finding)", async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('invoice2013.docx'));
  const body = await pkg.getBody();
  const tables = [];
  const walk = (b) => { for (const t of b.tables) { tables.push(t); for (const row of t.rows) for (const cell of row.cells) walk(cell.body); } };
  walk(body);
  let checked = 0;
  for (const table of tables) for (const row of table.rows) for (const cell of row.cells) for (const p of cell.body.paragraphs) {
    assert.equal(p.parentTableCell?.element.value, cell.element.value, JSON.stringify(p.text));
    checked++;
  }
  const item = body.paragraphs.find((p) => p.text === '   $120   ');
  assert.equal(item.markText, '\t');
  assert.ok(item.parentTableCell, 'the cell is found, four holders up: a repeating section item in its section');
  assert.equal(checked, 17, "every cell paragraph, those under the repeating section among them");
  // a hit never crosses a cell's end, nor a table's edge; until now one ran from that cell past the table, "\u0020\r\n\r\u2612"
  const hits = body.search('?^13?', { matchWildcards: true }).map((r) => r.text);
  assert.equal(hits.length, 9);
  assert.deepEqual(hits.filter((text) => text.includes('\n') || text.includes('\t')), []);
});

test('CR-012 section 19: a deleted mark reads as nothing in Word\'s default view, the paragraphs apart, a cell\'s end kept (check 58)', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const del = '<w:del w:id="1" w:author="A" w:date="2026-10-23T00:00:00Z"/>';
  await body.insertXml(`<w:p><w:pPr><w:rPr>${del}</w:rPr></w:pPr><w:r><w:t>ONE</w:t></w:r></w:p><w:p><w:r><w:t>TWO</w:t></w:r></w:p>`
    + `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:rPr>${del.replace('"1"', '"2"')}</w:rPr></w:pPr><w:r><w:t>CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`
    + '<w:p><w:r><w:t>LAST</w:t></w:r></w:p>', 'End');
  const [one, two] = body.paragraphs;
  assert.equal(one.markText, '');
  assert.equal(one.getRange('Whole').text, 'ONE');
  assert.equal(one.getRange('Whole').expandTo(two.getRange('Whole')).text, 'ONETWO\r');
  assert.equal(body.paragraphs.length, 4, 'the paragraphs stay apart');
  // a cell's last paragraph ends with the cell's \t whatever its mark (check 58 cases 9 and 10)
  assert.equal(body.paragraphs[2].markText, '\t');
  assert.equal(body.text, 'ONETWO\rCELL\r\nLAST');
  assert.equal(body.getText({ view: 'original' }), 'ONE\rTWO\rCELL\r\nLAST', 'the original view keeps a deleted mark');
});

test('CR-012 section 18: a control showing its placeholder reads as no text, is no text to search, and takes text in the placeholder\'s place (check 58)', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const placeholder = '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click here.</w:t></w:r>';
  const sdtPr = (id) => `<w:sdtPr><w:id w:val="${id}"/><w:showingPlcHdr/></w:sdtPr>`;
  await body.insertXml(`<w:sdt>${sdtPr(1)}<w:sdtContent><w:p>${placeholder}</w:p></w:sdtContent></w:sdt>`
    + `<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:sdt>${sdtPr(2)}<w:sdtContent>${placeholder}</w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> B</w:t></w:r></w:p>`, 'End');
  const [block, run] = body.contentControls;
  assert.deepEqual([block.text, run.text, body.paragraphs[1].text], ['', '', 'A  B']);
  // null, as Word reads a control naming no glossary entry (CR-002 section 63, D8; until then the runs were read)
  assert.deepEqual([block.placeholderText, run.placeholderText], [null, null]);
  assert.equal(body.search('Click').length, 0, 'K3: no text to search');
  assert.equal(body.text, '\rA  B');
  // text goes in the placeholder's place, at 'End' as at 'Replace', and the control shows it, plain, as Word's did (check
  // 58 cases 20 to 22; CR-002 section 63, D2 (b))
  block.insertText('FILLED', 'End');
  run.insertText('X', 'Start');
  assert.deepEqual([block.text, block.isShowingPlaceholder, run.text, run.isShowingPlaceholder], ['FILLED', false, 'X', false]);
  assert.equal(body.paragraphs[1].text, 'A X B');
  assert.ok([...block.paragraphs[0].runs, ...body.paragraphs[1].runs].every((r) => r.value.rPr === undefined), 'no Placeholder Text style, no colour');
});
