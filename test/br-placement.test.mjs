// CR-021: a w:br outside a w:r, read as Word reads it (the port of docx4j CR-073). docx4j's br-placement probe and Word 365's
// saves of it (test/br-placement/, copied unedited with a SHA-1 manifest): each probe loads, reports br.outside_run, reads as
// Word's save reads (Paragraph.text, the text model) and saves, unrefused, with Word's save's structure (docx4j's
// BrOutsideARunTest comparison: text, breaks, field characters and paragraph ends, each with its run's bold, italic and size).
// br2-3's bare w:instrText is a field code here, where Word shows it as text (CR-021 D5, docx4j's D6). Then docx4j's edge tests
// mirrored, the untouched and private reads, insertOoxml's two routes, and a caller's own hook.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, strFromU8 } from 'fflate';
import { WordprocessingMLPackage, MainDocumentPart, Texts, PlainText, repairBreaksOutsideRuns, mceKeep, parseXml } from '../dist/index.mjs';
import { breakTokens } from './helpers.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, 'br-placement');
const manifest = JSON.parse(await readFile(join(dir, 'manifest.json'), 'utf8'));
const bytesOf = async (name) => new Uint8Array(await readFile(join(dir, name)));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Each probe's warning: the story part holding the case, and docx4j's counts (its WARN's). */
const REPORTS = {
  'br-1-p-alone.docx': ['/word/document.xml', 1, 0],
  'br-2-p-between-runs.docx': ['/word/document.xml', 1, 0],
  'br-3-t-mixed.docx': ['/word/document.xml', 0, 1],
  'br-4-t-nested.docx': ['/word/document.xml', 0, 1],
  'br-5-hyperlink.docx': ['/word/document.xml', 1, 0],
  'br-6-cell.docx': ['/word/document.xml', 1, 0],
  'br-7-footer.docx': ['/word/footer1.xml', 1, 0],
  'br-8-p-page.docx': ['/word/document.xml', 1, 0],
  'br2-1-run-properties.docx': ['/word/document.xml', 8, 1],
  'br2-2-instrtext-field.docx': ['/word/document.xml', 0, 1],
  'br2-3-instrtext-bare-header.docx': ['/word/header1.xml', 0, 1],
  'br2-4-body-br.docx': ['/word/document.xml', 1, 0],
  'br2-5-body-br-page.docx': ['/word/document.xml', 1, 0],
};

/** D5: the bare w:instrText Word shows as text, a field code here. */
const BARE_CODE = 'br2-3-instrtext-bare-header.docx';

/** A package's story parts (main, headers, footers) as saved, by name. */
const storiesOf = (zip) => Object.fromEntries(Object.keys(zip).filter((n) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(n)).sort().map((n) => [n, strFromU8(zip[n])]));

async function read(name) {
  const pkg = await WordprocessingMLPackage.load(await bytesOf(name));
  await pkg.unmarshalAll();
  const paragraphs = (await pkg.getBody()).paragraphs.map((p) => p.text);
  const plain = PlainText.of(await Texts.of(pkg, { include: ['headers', 'footers'] }));
  return { pkg, paragraphs, plain };
}

test("the probe and Word 365's saves are as the manifest has them", async () => {
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await bytesOf(name)).digest('hex'), sha1, name);
  assert.deepEqual(Object.keys(manifest.saves), Object.keys(REPORTS));
  for (const [probe, save] of Object.entries(manifest.saves)) assert.ok(manifest.files.some((f) => f.name === probe) && manifest.files.some((f) => f.name === save), probe);
});

test('each probe loads with one br.outside_run warning, docx4j-python\'s fields and docx4j\'s message, and nothing misplaced; Word\'s save warns of nothing', async () => {
  for (const [name, [partName, breaks, textElements]] of Object.entries(REPORTS)) {
    const { pkg } = await read(name);
    const message = `${partName}: w:br outside a w:r: ${breaks} outside a run, ${textElements} text elements holding elements; read as Word reads them`;
    assert.deepEqual(pkg.warnings, [{ code: 'br.outside_run', partName, breaks, textElements, message }], name);
    assert.deepEqual(pkg.getPart(partName).warnings, pkg.warnings, `${name}: on its part`);
    assert.deepEqual(await pkg.checkContent(), [], `${name}: nothing misplaced`);
    assert.deepEqual((await read(manifest.saves[name])).pkg.warnings, [], `${name}: Word's save`);
  }
});

test('each probe reads as Word\'s save reads: Paragraph.text in Office JS\'s characters, and the text model with headers and footers but D5\'s bare field code', async () => {
  for (const name of Object.keys(REPORTS)) {
    const ours = await read(name);
    const word = await read(manifest.saves[name]);
    assert.deepEqual(ours.paragraphs, word.paragraphs, name);
    if (name !== BARE_CODE) assert.equal(ours.plain, word.plain, name);
  }
  // Word saves the bare w:instrText as w:t and shows it; here, as in docx4j, it is a field code, hidden by the text model: its
  // breaks are read, its text not
  const ours = await read(BARE_CODE);
  const word = await read(manifest.saves[BARE_CODE]);
  assert.equal(word.plain, ' \n I2 line one \n I2 line two \nI2 body text; the case is in the header\n');
  assert.equal(ours.plain, '\n\n\nI2 body text; the case is in the header\n');
  // the breaks in Office JS's characters, either side of the text (a case of each in the probe)
  assert.deepEqual((await read('br-3-t-mixed.docx')).paragraphs, ['BR3 before\u000bBR3 after']);
  assert.deepEqual((await read('br-8-p-page.docx')).paragraphs, ['BR8 before\u000c', 'BR8 after']);
  assert.deepEqual((await read('br2-4-body-br.docx')).paragraphs, ['B1 before', '\u000bB1 after']);
});

test('a save after unmarshalling is not refused, and each story part holds Word\'s save\'s structure, the run properties included (docx4j BrOutsideARunTest)', async () => {
  for (const name of Object.keys(REPORTS)) {
    const { pkg } = await read(name);
    const saved = storiesOf(unzipSync(await pkg.save()));
    assert.deepEqual(pkg.lastSave.misplaced, [], name);
    const word = storiesOf(unzipSync(await bytesOf(manifest.saves[name])));
    assert.deepEqual(Object.keys(saved), Object.keys(word), name);
    for (const [part, xml] of Object.entries(word)) {
      const tokens = breakTokens(saved[part]);
      if (name === BARE_CODE && part === 'word/header1.xml') {
        // D5: the field code's text and breaks whole, its element Word's w:t
        assert.ok(tokens.includes('instrText: I2 line one '), tokens.join(' | '));
        assert.deepEqual(tokens.map((t) => t.replace(/^instrText/, 't')), breakTokens(xml), `${name} ${part}`);
      } else assert.deepEqual(tokens, breakTokens(xml), `${name} ${part}`);
    }
  }
  // D1's run properties, read off br2-1: the mark's (sz28) while no run precedes the break, none after one, the bold run's in its w:t
  const tokens = breakTokens(storiesOf(unzipSync(await (await read('br2-1-run-properties.docx')).pkg.save()))['word/document.xml']);
  assert.ok(tokens.includes('br[sz28]'), 'R3 and R5: the mark\'s');
  assert.ok(tokens.includes('br[b sz40]'), 'R6: the bold run\'s');
  assert.equal(tokens.filter((t) => t === 'br').length, 6, 'R1, R2 (two), R4, R7 and R8: none');
});

test('untouched, the part is copied from the source byte for byte; a private read (the text model, checkContent) repairs its own tree and leaves the part untouched', async () => {
  for (const name of Object.keys(REPORTS)) {
    const source = storiesOf(unzipSync(await bytesOf(name)));
    const pkg = await WordprocessingMLPackage.load(await bytesOf(name));
    const plain = PlainText.of(await Texts.of(pkg, { include: ['headers', 'footers'] }));
    assert.equal(plain, (await read(name)).plain, `${name}: the private read sees the repair`);
    assert.deepEqual(await pkg.checkContent(), []);
    const part = pkg.getPart(REPORTS[name][0]);
    assert.equal(part.isUntouched, true, name);
    assert.deepEqual(storiesOf(unzipSync(await pkg.save())), source, `${name}: byte for byte`);
  }
});

// --- docx4j BrOutsideARunTest's edge tests, mirrored: the body's tokens after an unmarshal and a marshal

async function bodyOf(content) {
  const pkg = new WordprocessingMLPackage();
  const main = new MainDocumentPart();
  main.setXml(`<w:document xmlns:w="${W}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>${content}</w:body></w:document>`);
  pkg.addTargetPart(main);
  await main.getContents();
  const xml = await main.getXml();
  return { xml, tokens: breakTokens(xml), main };
}

test('a math run holds a w:br legally and keeps it while a stray one beside it is repaired; a w:br in a w:r is left as it is', async () => {
  const { xml, tokens, main } = await bodyOf('<w:p><w:r><w:t>a</w:t><w:br/></w:r><m:oMath><m:r><w:br/><m:t>x</m:t></m:r></m:oMath><w:br/><w:r><w:t>b</w:t></w:r></w:p>');
  assert.deepEqual(tokens, ['t:a', 'br', 'br', 't:b', '¶']);
  assert.match(xml, /<m:r><w:br\/><m:t>x<\/m:t><\/m:r>/, 'the math run\'s w:br');
  assert.deepEqual(main.warnings.map((w) => [w.breaks, w.textElements]), [[1, 0]]);
});

test('whitespace between a text element\'s children is kept only where xml:space preserves it', async () => {
  assert.deepEqual((await bodyOf('<w:p><w:r><w:t>  <w:br/></w:t></w:r></w:p>')).tokens, ['br', '¶']);
  assert.deepEqual((await bodyOf('<w:p><w:r><w:t xml:space="preserve">  <w:br/></w:t></w:r></w:p>')).tokens, ['t:  ', 'br', '¶']);
});

test('the mark\'s properties without what only a mark has (a tracked insertion of the mark)', async () => {
  const { xml, tokens } = await bodyOf('<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="a"/><w:b/></w:rPr></w:pPr><w:br/><w:r><w:t>x</w:t></w:r></w:p>');
  assert.deepEqual(tokens, ['br[b]', 't:x', '¶']);
  assert.match(xml, /<\/w:pPr><w:r><w:rPr><w:b\/><\/w:rPr><w:br\/><\/w:r>/);
});

test('deleted text keeps its break, and a stray w:br in a w:del gets a run there', async () => {
  assert.deepEqual((await bodyOf('<w:p><w:del w:id="1" w:author="a"><w:r><w:delText>x<w:br/>y</w:delText></w:r></w:del></w:p>')).tokens, ['delText:x', 'br', 'delText:y', '¶']);
  assert.deepEqual((await bodyOf('<w:p><w:del w:id="1" w:author="a"><w:r><w:delText>x</w:delText></w:r><w:br/><w:r><w:delText>y</w:delText></w:r></w:del></w:p>')).tokens,
    ['delText:x', 'br', 'delText:y', '¶']);
});

test('in w:body before a paragraph with no w:pPr: still its first runs, in order; before a table (unmeasured), a paragraph of its own', async () => {
  const { xml, tokens } = await bodyOf('<w:p><w:r><w:t>a</w:t></w:r></w:p><w:br/><w:br/><w:p><w:r><w:t>b</w:t></w:r></w:p>'
    + '<w:br w:type="page"/><w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  assert.deepEqual(tokens, ['t:a', '¶', 'br', 'br', 't:b', '¶', 'br(page)', '¶', 't:c', '¶']);
  assert.match(xml, /<w:p><w:r><w:br w:type="page"\/><\/w:r><\/w:p><w:tbl>/);
});

// --- this package's own

test('the pass alone: a part outside WordprocessingML and a clean part are left alone, with nothing reported', () => {
  const clean = parseXml(`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>a</w:t><w:br/></w:r></w:p></w:body></w:document>`);
  assert.deepEqual(repairBreaksOutsideRuns(clean, '/word/document.xml'), []);
  const other = parseXml('<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:w="' + W + '"><w:p><w:br/></w:p></p:sld>');
  assert.deepEqual(repairBreaksOutsideRuns(other), []);
  assert.equal(other.getElementsByTagNameNS(W, 'r').length, 0, 'unchanged');
  // a w:br in a block-level w:sdtContent is counted, as docx4j's WARN counts it, and left: its nearest block is the body
  const block = parseXml(`<w:document xmlns:w="${W}"><w:body><w:sdt><w:sdtContent><w:br/></w:sdtContent></w:sdt></w:body></w:document>`);
  assert.deepEqual(repairBreaksOutsideRuns(block).map((w) => [w.breaks, w.textElements, w.partName]), [[1, 0, undefined]]);
  assert.equal(block.getElementsByTagNameNS(W, 'r').length, 0);
});

test('two stray breaks before any run both take the mark\'s properties, decided on the tree as read; a nested w:t holding a break counts as docx4j counts it', async () => {
  const { tokens } = await bodyOf('<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:br/><w:br/><w:r><w:t>x</w:t></w:r><w:br/></w:p>');
  assert.deepEqual(tokens, ['br[sz28]', 'br[sz28]', 't:x', 'br', '¶']);
  const nested = await bodyOf('<w:p><w:r><w:t>a<w:t>b<w:br/>c</w:t>d</w:t></w:r></w:p>');
  assert.deepEqual(nested.tokens, ['t:ab', 'br', 't:cd', '¶']);
  assert.deepEqual(nested.main.warnings.map((w) => [w.breaks, w.textElements]), [[0, 2]]);
});

test('insertOoxml: a fragment with a w:br outside a run, into a created package (no hook), and a pkg:package whose w:t holds one', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  assert.equal(pkg.loadOptions.preprocessor, undefined, 'a created package has no hook');
  await pkg.body.insertOoxml('<w:p><w:r><w:t>a</w:t></w:r><w:br/><w:r><w:t>b</w:t></w:r></w:p>', 'End');
  await pkg.body.insertOoxml('<w:p><w:r><w:t>c<w:br/>d</w:t></w:r></w:p>', 'End');
  const texts = pkg.body.paragraphs.map((p) => p.text);
  assert.deepEqual(texts.slice(-2), ['a\u000bb', 'c\u000bd']);
  await pkg.save();
  assert.deepEqual(pkg.lastSave.misplaced, []);
  // the flat OPC route loads a package, whose parts repair as any
  const source = await WordprocessingMLPackage.createPackage();
  source.body.insertParagraph('MARKER', 'End');
  const flatOpc = (await source.saveFlatOpc()).replace(/<w:t>MARKER<\/w:t>/, '<w:t>e<w:br/>f</w:t>');
  assert.ok(flatOpc.includes('<w:t>e<w:br/>f</w:t>'));
  await pkg.body.insertOoxml(flatOpc, 'End');
  assert.equal(pkg.body.paragraphs.at(-1).text, 'e\u000bf');
});

test('a caller\'s own hook sees the repaired DOM, and its warnings follow the repair\'s', async () => {
  let inRuns;
  const pkg = await WordprocessingMLPackage.load(await bytesOf('br-2-p-between-runs.docx'), {
    preprocessor: (doc, part) => {
      inRuns = [...doc.getElementsByTagNameNS(W, 'br')].map((br) => br.parentNode.localName);
      return mceKeep(doc, part);
    },
  });
  await pkg.unmarshalAll();
  assert.deepEqual(inRuns, ['r']);
  assert.deepEqual(pkg.warnings.map((w) => w.code), ['br.outside_run']);
});
