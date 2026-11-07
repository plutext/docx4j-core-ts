// CR-026: an element Word doesn't know, read as Word 365 reads it (check 76, WD REQ-308 to REQ-316).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { PlainText, SaveLossException, Texts, WordprocessingMLPackage } from '../dist/index.mjs';
import { WML_ELEMENT_NAMES } from '../dist/opc/repair/wmlNames.generated.mjs';
import { readThroughUnknownElements } from '../dist/opc/repair/unknownElements.mjs';
import { parseXml, serializeXml } from '../dist/xml/dom.mjs';
import { check76Files } from './fixtures/check76/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const EXT = 'urn:docx4j:cr026';
const roots = `xmlns:w="${W}" xmlns:mc="${MC}" xmlns:ext="${EXT}"`;
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;

async function documentOf(blocks) {
  const base = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  return WordprocessingMLPackage.load(zipSync({ ...base, 'word/document.xml': strToU8(`<w:document ${roots}><w:body>${blocks}</w:body></w:document>`) }));
}
const skippedIn = (part) => part.warnings.filter((w) => w.code === 'part.skipped');
async function read(blocks) {
  const pkg = await documentOf(blocks);
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  return { pkg, main, xml: await main.getXml(), body: await pkg.getBody() };
}
const textOf = async (bytes) => {
  const pkg = await WordprocessingMLPackage.load(new Uint8Array(bytes));
  await pkg.unmarshalAll();
  return PlainText.of(await Texts.of(pkg, { include: ['headers', 'footers'], comments: true }));
};

test('the schema\'s names: 605, a name with no class among them', () => {
  assert.equal(WML_ELEMENT_NAMES.size, 605);
  assert.ok(WML_ELEMENT_NAMES.has('uniqueTag'));
  assert.ok(!WML_ELEMENT_NAMES.has('madeUp'));
});

// Route B: the runtime's skip for an element in a list, jsonix-CR-015, since the floor's raise. 76t71 reads without the held w:p's
// words, where Word keeps them and makes the link a field: REQ-316 requires nothing of it, so it's pinned as a known difference.
const KNOWN_DIFFERENCE = '76t71-held-paragraph-in-hyperlink';

test('check 76: every source Word saved reads as Word\'s save of it does, the text the same', async () => {
  const dir = join(here, 'check76');
  let held = 0;
  for (const f of await check76Files()) {
    const save = join(dir, `${f.name}-word365.docx`);
    if (!existsSync(save)) continue;
    if (f.name === KNOWN_DIFFERENCE) {
      assert.notEqual(await textOf(f.bytes), await textOf(readFileSync(save)), f.name);
      continue;
    }
    assert.equal(await textOf(f.bytes), await textOf(readFileSync(save)), f.name);
    held++;
  }
  assert.ok(held >= 70, `${held} held`);
});

test('D7: the tags go and what they held stands, one part.skipped for each element taken, at any depth', async () => {
  const { main, body } = await read(`<w:p>${run('a')}<w:madeUp><w:madeUp>${run('b')}</w:madeUp></w:madeUp>${run('c')}</w:p>`);
  assert.equal(body.paragraphs[0].text, 'abc');
  assert.deepEqual(skippedIn(main).map((w) => [w.kind, w.name, w.parentName, w.path]), [
    ['element', 'w:madeUp', 'w:p', '/w:document/w:body/w:p/w:madeUp'],
    ['element', 'w:madeUp', 'w:madeUp', '/w:document/w:body/w:p/w:madeUp/w:madeUp'],
  ]);
});

test('`76k`: an element of another namespace stops the walk, so a w:madeUp in one isn\'t read through', async () => {
  const pkg = await documentOf(`<w:p><w:pPr><ext:opaque><w:madeUp>${run('x')}</w:madeUp></ext:opaque></w:pPr>${run('a')}</w:p>`);
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  assert.equal((await pkg.getBody()).paragraphs[0].text, 'a');
  assert.deepEqual(skippedIn(main).map((w) => w.name), ['ext:opaque']);
});

test('a kept branch doesn\'t stop the walk (`76t44`, `76t45`), and a reader meets a DOM element in one without failing', async () => {
  const branch = (choice) => `<w:p>${run('a')}<mc:AlternateContent><mc:Choice xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" Requires="w14">${choice}</mc:Choice><mc:Fallback>${run('f')}</mc:Fallback></mc:AlternateContent></w:p>`;
  const through = await read(branch(`<w:madeUp>${run('b')}</w:madeUp>`));
  assert.equal(PlainText.of(await Texts.of(through.pkg)).trim(), 'ab');
  const held = await read(branch('<ext:opaque/>'));
  assert.equal(PlainText.of(await Texts.of(held.pkg)).trim(), 'a');
});

test('the same names under parents that make them something else are no host: the pass leaves what\'s in them', () => {
  const m = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
  for (const inner of [
    '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="a"><w:madeUp/></w:ins></w:rPr></w:pPr></w:p>',
    '<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="a"><w:madeUp/></w:del></w:rPr></w:pPr></w:p>',
    '<w:p><w:pPr><w:rPr><w:moveFrom w:id="1" w:author="a"><w:madeUp/></w:moveFrom></w:rPr></w:pPr></w:p>',
    '<w:p><w:pPr><w:rPr><w:moveTo w:id="1" w:author="a"><w:madeUp/></w:moveTo></w:rPr></w:pPr></w:p>',
    '<w:p><w:pPr><w:numPr><w:ins w:id="1" w:author="a"><w:madeUp/></w:ins></w:numPr></w:pPr></w:p>',
    '<w:tbl><w:tr><w:trPr><w:ins w:id="1" w:author="a"><w:madeUp/></w:ins></w:trPr></w:tr></w:tbl>',
    '<w:tbl><w:tr><w:trPr><w:del w:id="1" w:author="a"><w:madeUp/></w:del></w:trPr></w:tr></w:tbl>',
    `<w:p><m:r ${m}><w:ins w:id="1" w:author="a"><w:madeUp/></w:ins></m:r></w:p>`,
    `<w:p><m:r ${m}><m:ctrlPr><w:del w:id="1" w:author="a"><w:madeUp/></w:del></m:ctrlPr></m:r></w:p>`,
    '<w:sectPr><w:footnotePr><w:footnote w:id="0"><w:madeUp/></w:footnote></w:footnotePr></w:sectPr>',
    '<w:sectPr><w:endnotePr><w:endnote w:id="0"><w:madeUp/></w:endnote></w:endnotePr></w:sectPr>',
  ]) {
    const doc = parseXml(`<w:document ${roots}><w:body>${inner}</w:body></w:document>`);
    assert.deepEqual(readThroughUnknownElements(doc), [], inner);
    assert.match(serializeXml(doc), /w:madeUp/, inner);
  }
});

test('REQ-314: text straight in the element goes, reported, and a run beside it stands; none of it is CR-016\'s string in a run', async () => {
  const { main, body } = await read(`<w:p><w:r><w:madeUp>loose</w:madeUp><w:t>a</w:t></w:r><w:madeUp>gone${run('b')}</w:madeUp></w:p>`);
  assert.equal(body.paragraphs[0].text, 'ab');
  assert.deepEqual(skippedIn(main).map((w) => [w.kind, w.name]), [['element', 'w:madeUp'], ['text', 'loose'], ['element', 'w:madeUp'], ['text', 'gone']]);
  assert.deepEqual(main.warnings.filter((w) => w.code === 'target.misplaced'), []);
});

test('REQ-315: a w:br in the element is left for CR-021\'s rule 1, which gives it a run (`76t42`)', async () => {
  const { xml } = await read(`<w:p>${run('a')}<w:madeUp><w:br/></w:madeUp>${run('b')}</w:p>`);
  assert.match(xml, /<w:r>\s*<w:br\/>\s*<\/w:r>/);
});

test('D8, REQ-311: runs in a w:pPr\'s element go to the head in order, and a property beside them stays the runtime\'s to drop', async () => {
  const two = await read(`<w:p><w:pPr><w:madeUp>${run('1')}${run('2')}</w:madeUp></w:pPr>${run('a')}</w:p>`);
  assert.equal(two.body.paragraphs[0].text, '12a');
  assert.equal(skippedIn(two.main).length, 1);
  const mixed = await read(`<w:p><w:pPr><w:madeUp>${run('1')}<w:jc w:val="center"/></w:madeUp></w:pPr>${run('a')}</w:p>`);
  assert.equal(mixed.body.paragraphs[0].text, '1a');
  assert.equal(mixed.body.paragraphs[0].alignment, 'Left');
  assert.deepEqual(skippedIn(mixed.main).map((w) => w.parentClass), ['org_docx4j_wml.PPr'], 'once, by the read that skips it');
  const alone = await read(`<w:p><w:pPr><w:madeUp>${run('1')}</w:madeUp></w:pPr></w:p>`);
  assert.equal(alone.body.paragraphs[0].text, '1');
});

test('D9: a run under a w:pPrChange is read, and the change keeps the author and date it holds', async () => {
  const { xml, body } = await read(`<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="1" w:author="Someone" w:date="2026-11-05T00:00:00Z"><w:pPr><w:madeUp>${run('1')}</w:madeUp></w:pPr></w:pPrChange></w:pPr>${run('a')}</w:p>`);
  assert.equal(body.paragraphs[0].text, '1a');
  assert.match(xml, /w:author="Someone"/);
  assert.match(xml, /w:date="2026-11-05T00:00:00Z"/);
});

test('D10: a run in a w:sectPr in a w:pPr is read, and the w:sectPr goes, reported', async () => {
  const { main, xml, body } = await read(`<w:p><w:pPr><w:sectPr><w:madeUp>${run('1')}</w:madeUp><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>${run('a')}</w:p>`);
  assert.equal(body.paragraphs[0].text, '1a');
  assert.deepEqual(skippedIn(main).map((w) => w.name), ['w:madeUp', 'w:sectPr']);
  assert.doesNotMatch(xml.slice(0, xml.indexOf('</w:pPr>') + 1 || undefined), /<w:sectPr/);
});

const aligned = (jc, inner) => `<w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr>${inner}</w:p>`;
const texts = (body) => body.paragraphs.map((p) => [p.text, p.alignment]);
test('D11, REQ-316: a held w:p ends its host: the host takes its properties, else keeps its own, and what\'s left is empty with the host\'s', async () => {
  const both = await read(aligned('right', `${run('a')}<w:madeUp>${aligned('center', run('b'))}</w:madeUp>`));
  assert.deepEqual(texts(both.body), [['ab', 'Centered'], ['', 'Right']]);
  const plainHeld = await read(aligned('right', `${run('a')}<w:madeUp><w:p>${run('b')}</w:p></w:madeUp>`));
  assert.deepEqual(texts(plainHeld.body), [['ab', 'Right'], ['', 'Right']]);
  const two = await read(`<w:p>${run('a')}<w:madeUp><w:p>${run('b')}</w:p><w:p>${run('c')}</w:p></w:madeUp></w:p>`);
  assert.deepEqual(two.body.paragraphs.map((p) => p.text), ['ab', 'c', '']);
  const deep = await read(`<w:p>${run('a')}<w:madeUp><w:madeUp><w:p>${run('b')}</w:p></w:madeUp></w:madeUp></w:p>`);
  assert.deepEqual(deep.body.paragraphs.map((p) => p.text), ['ab', '']);
  assert.equal(skippedIn(deep.main).length, 2);
});

test('CR-025 section 2.2: a file read through, read and not changed, is copied with its tags; changed, its save is refused', async () => {
  const bytes = await (await documentOf(`<w:p><w:madeUp>${run('a')}</w:madeUp></w:p>`)).save();
  const reader = await WordprocessingMLPackage.load(bytes);
  await reader.getBody();
  assert.match(strFromU8(unzipSync(await reader.save())['word/document.xml']), /w:madeUp/);
  const writer = await WordprocessingMLPackage.load(bytes);
  await writer.getBody();
  writer.body.paragraphs[0].insertText(' more', 'End');
  await assert.rejects(writer.save(), SaveLossException);
  writer.checkLossOnSave = false;
  assert.doesNotMatch(strFromU8(unzipSync(await writer.save())['word/document.xml']), /w:madeUp/);
});

test('D12: insertOoxml refuses a caller\'s own fragment holding an unknown element, wherever it stands, and takes a clean one', async () => {
  const body = async () => { const pkg = await WordprocessingMLPackage.createPackage(); pkg.body.insertParagraph('x', 'End'); return pkg.body; };
  const w = `xmlns:w="${W}"`;
  for (const fragment of [
    `<w:p ${w}><w:madeUp><w:r><w:t>a</w:t></w:r></w:madeUp></w:p>`,
    `<w:p ${w}><w:r><w:t>a<w:madeUp/>b</w:t></w:r></w:p>`,
    `<w:p ${w} xmlns:ext="${EXT}"><w:pPr><ext:opaque/></w:pPr></w:p>`,
  ]) await assert.rejects((await body()).insertOoxml(fragment, 'End'), /insertOoxml refused/, fragment);
  await (await body()).insertOoxml(`<w:p ${w}><w:r><w:t>a</w:t></w:r></w:p>`, 'End');
});

test('D12: a flat OPC package holding one is refused too', async () => {
  const pkg = await documentOf(`<w:p><w:madeUp>${run('a')}</w:madeUp></w:p>`);
  const flat = await pkg.saveFlatOpc();
  const target = await WordprocessingMLPackage.createPackage();
  target.body.insertParagraph('x', 'End');
  await assert.rejects(target.body.insertOoxml(flat, 'End'), /insertOoxml refused/);
});
