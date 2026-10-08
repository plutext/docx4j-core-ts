// CR-005 section 8: HTML to the intermediate form (the editor's paste corpus, moved here with the
// converter on 2026-10-08) and that form to WordprocessingML through the engine's own parser.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { parseHTML } from 'linkedom';
import { WordprocessingMLPackage, contentOf, typeNameOf } from '../dist/index.mjs';
import { convertHtml, styleLookupOf, clipboardMarkup, isTerminalHtml, blocksToXml, needsOf } from '../dist/html/index.mjs';

const corpus = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'html');
const parser = (html) => parseHTML(html).document;
const options = { parser };

/** A capture as text: the clipboard's UTF-8 was read as Windows-1252 and written as UTF-8 again on the reference machine (README). */
async function capture(name) {
  let text = (await readFile(join(corpus, `${name}.html`))).toString('utf8').replace(/^﻿/, '');
  if (/[ÂÃ][\u0080-¿€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ]/.test(text)) text = fromCp1252(text);
  return text;
}
const CP1252 = { '€': 0x80, '‚': 0x82, 'ƒ': 0x83, '„': 0x84, '…': 0x85, '†': 0x86, '‡': 0x87, 'ˆ': 0x88, '‰': 0x89, 'Š': 0x8a, '‹': 0x8b, 'Œ': 0x8c, 'Ž': 0x8e, '‘': 0x91, '’': 0x92, '“': 0x93, '”': 0x94, '•': 0x95, '–': 0x96, '—': 0x97, '˜': 0x98, '™': 0x99, 'š': 0x9a, '›': 0x9b, 'œ': 0x9c, 'ž': 0x9e, 'Ÿ': 0x9f };
function fromCp1252(text) {
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); bytes[i] = c < 0x100 ? c : (CP1252[text[i]] ?? 0x3f); }
  return Buffer.from(bytes).toString('utf8');
}

/** The lookup a document gives: its paragraph and character styles by id and name, and whether it has Word's table grid. */
async function lookupFor(pkg) {
  const styles = await pkg.getMainDocumentPart().styleDefinitionsPart.readContents();
  const entries = [];
  for (const s of styles.style ?? []) {
    if (!s.styleId) continue;
    if (s.type !== undefined && s.type !== 'paragraph' && s.type !== 'character') continue;
    if (s.hidden !== undefined && s.hidden.val !== false) continue;
    entries.push({ id: s.styleId, name: s.name?.val ?? s.styleId });
  }
  return styleLookupOf(entries, (await pkg.getPropertyResolver()).getStyle('TableGrid') ? ['TableGrid'] : []);
}
const samples = async () => WordprocessingMLPackage.load(await readFile(join(corpus, 'paste-samples.docx')));

const NAMES = ['headings', 'inline', 'bullets', 'numbering', 'table', 'dropped', 'all'];
for (const name of NAMES) {
  test(`${name}.html converts as pinned (the editor's expectations, unchanged by the move)`, async () => {
    const result = convertHtml(await capture(name), await lookupFor(await samples()), options);
    const path = join(corpus, `${name}.expected.json`);
    const json = JSON.stringify(result, null, 2) + '\n';
    if (process.env['UPDATE_HTML']) { await writeFile(path, json); return; }
    assert.deepEqual(JSON.parse(json), JSON.parse(await readFile(path, 'utf8')));
  });
}

test("the CF_HTML header is skipped; a terminal's HTML is told from Word's", async () => {
  const text = await capture('inline');
  assert.ok(/^Version:/.test(text), 'the capture carries the CF_HTML header');
  assert.doesNotMatch(clipboardMarkup(text), /^Version:|StartHTML:/);   // the header gone, the fragment Word marked kept
  assert.match(clipboardMarkup(text), /^\s*<h1>/);
  const konsole = (await readFile(join(corpus, 'terminal', 'konsole-26.04.html'))).toString('utf8');
  assert.equal(isTerminalHtml(konsole, options), true);
  assert.equal(isTerminalHtml(text, options), false);
});

test('every capture builds: the fragment parses to as many blocks as the form has, paragraphs and tables alike, and inserts into a new document', async () => {
  for (const name of NAMES) {
    const { blocks } = convertHtml(await capture(name), await lookupFor(await samples()), options);
    const xml = blocksToXml(blocks, { numId: () => '1', relId: () => 'rId9' });
    const elements = await contentOf(xml);
    assert.equal(elements.length, blocks.length, name);
    elements.forEach((el, i) => assert.equal(typeNameOf(el), blocks[i].kind === 'table' ? 'org_docx4j_wml.Tbl' : 'org_docx4j_wml.P', `${name} block ${i}`));
    const pkg = await WordprocessingMLPackage.createPackage();
    await (await pkg.getMainDocumentPart().getBody()).insertXml(xml, 'End');
    const again = await WordprocessingMLPackage.load(await pkg.save());
    const paragraphs = (await again.getMainDocumentPart().getBody()).paragraphs;
    const expected = blocks.filter((b) => b.kind === 'paragraph').map((b) => b.inlines.map((i) => ('text' in i ? i.text : 'tab' in i ? '\t' : '\n')).join(''));
    const got = paragraphs.map((p) => p.text).filter((t, i) => i < expected.length || t !== '');
    for (const [i, text] of expected.entries()) assert.ok(got.includes(text), `${name}: paragraph ${i} ${JSON.stringify(text)} present after the round trip`);
  }
});

test('a list gets its numbering, a link its hyperlink in Word\'s look, and neither without the document\'s answer', async () => {
  const lookup = await lookupFor(await samples());
  const html = '<ul><li>one</li><li>two</li></ul><p>See <a href="https://example.org/x">the site</a> and <a href="mailto:a@b.c">mail</a>.</p>';
  const { blocks } = convertHtml(html, lookup, options);
  assert.deepEqual(needsOf(blocks).links, ['https://example.org/x', 'mailto:a@b.c']);
  assert.equal(needsOf(blocks).lists.length, 1);
  assert.equal(needsOf(blocks).lists[0].bullet, true);
  const resolved = blocksToXml(blocks, { numId: () => '7', relId: (a) => (a.startsWith('https') ? 'rId5' : undefined) });
  assert.match(resolved, /<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="7"\/><\/w:numPr>/);
  assert.match(resolved, /<w:hyperlink r:id="rId5"><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t>the site<\/w:t><\/w:r><\/w:hyperlink>/);
  assert.doesNotMatch(resolved, /mail<\/w:t><\/w:r><\/w:hyperlink>/);   // the mail link, unresolved, keeps its text plain
  const plain = blocksToXml(blocks);
  assert.doesNotMatch(plain, /numPr|hyperlink/);
  assert.equal((await contentOf(resolved)).length, 3);
});

test('a cell spanning rows holds its place below with vMerge continuations; a cell spanning columns has its gridSpan; a header row repeats', async () => {
  const lookup = await lookupFor(await samples());
  const html = '<table><tr><th rowspan="2">a</th><th colspan="2">b</th></tr><tr><td>c</td><td>d</td></tr><tr><td>e</td><td>f</td><td>g</td></tr></table>';
  const { blocks } = convertHtml(html, lookup, options);
  assert.equal(blocks[0].kind, 'table');
  const xml = blocksToXml(blocks);
  const rows = xml.split('<w:tr>').slice(1);
  assert.equal(rows.length, 3);
  assert.match(rows[0], /<w:vMerge w:val="restart"\/>/);
  assert.match(rows[0], /<w:gridSpan w:val="2"\/>/);
  assert.match(rows[1], /^<w:tc><w:tcPr><w:tcW w:w="\d+" w:type="dxa"\/><w:vMerge\/><\/w:tcPr><w:p\/><\/w:tc>/);
  assert.equal((rows[1].match(/<w:tc>/g) ?? []).length, 3);
  assert.equal((rows[2].match(/<w:tc>/g) ?? []).length, 3);
  assert.doesNotMatch(rows[2], /vMerge/);
  const pkg = await WordprocessingMLPackage.createPackage();
  await (await pkg.getMainDocumentPart().getBody()).insertXml(xml, 'End');
  const again = await WordprocessingMLPackage.load(await pkg.save());
  assert.equal((await again.getMainDocumentPart().getBody()).tables.length, 1);
});
