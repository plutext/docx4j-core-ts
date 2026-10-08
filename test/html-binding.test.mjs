// CR-005 section 8.2: escaped XHTML bound through the tag's od:xpath entry, converted by the html
// module and written as the control's content; no w:dataBinding is involved.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseHTML } from 'linkedom';
import { WordprocessingMLPackage, tagParamsOf, xpathsEntriesOf, Namespaces } from '../dist/index.mjs';
import { fixture } from './helpers.mjs';

const parser = (html) => parseHTML(html).document;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const escape = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** A new document with one control bound to a data node holding `markup`, escaped, through an XPaths entry x1. */
async function template(markup, { run = false, entry = 'x1' } = {}) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const p = body.insertParagraph('placeholder', 'End');
  // A run-level control is made over a text range (a paragraph's own insertContentControl wraps the paragraph, as Office JS's does).
  const control = run ? body.search('placeholder')[0].insertContentControl('RichText') : body.insertContentControl('RichText');
  const data = pkg.customXmlParts.add(`<data><body>${escape(markup)}</body><other>x</other></data>`);
  pkg.customXmlParts.add(`<xpaths xmlns="http://opendope.org/xpaths"><xpath id="x1"><dataBinding xmlns:w="${W}" w:storeItemID="${data.id}" w:xpath="/data/body" w:prefixMappings=""/></xpath></xpaths>`);
  control.tag = `od:xpath=${entry}&od:ContentType=application/xhtml+xml`;
  return { pkg, body, control };
}

test('the tag and the XPaths part are read', async () => {
  assert.deepEqual([...tagParamsOf('od:xpath=x1&od:ContentType=application/xhtml+xml')], [['od:xpath', 'x1'], ['od:ContentType', 'application/xhtml+xml']]);
  const { pkg } = await template('<p>x</p>');
  const entries = xpathsEntriesOf(pkg.customXmlParts);
  assert.equal(entries.size, 1);
  assert.equal(entries.get('x1').xpath, '/data/body');
  assert.ok(entries.get('x1').storeItemID.length > 10);
});

test('a block-level control takes the markup\'s blocks: a heading in its style, bold, a link in Word\'s look with its relationship, a list with a definition; saved and reloaded; no data binding', async () => {
  const markup = '<h1>Title</h1><p>Hello <b>bold</b> and <a href="https://example.org/x">a link</a>.</p><ul><li>one</li><li>two</li></ul>';
  const { pkg, control } = await template(markup);
  assert.equal(control.xmlMapping.isMapped, false);
  const result = await pkg.customXmlParts.applyBindings({ html: { parser } });
  assert.equal(result.bound, 1);
  assert.equal(result.updated, 1);
  assert.equal(result.skipped, 0);
  assert.equal(result.notes, undefined);
  const texts = control.paragraphs.map((p) => p.text);
  assert.deepEqual(texts, ['Title', 'Hello bold and a link.', 'one', 'two']);
  assert.equal(control.paragraphs[0].style, 'Heading 1');
  assert.ok(control.paragraphs[2].listItem, 'the list paragraph is numbered');
  assert.equal(control.paragraphs[2].listItem.level, 0);
  const again = await WordprocessingMLPackage.load(await pkg.save());
  const main = again.getMainDocumentPart();
  const body = await main.getBody();
  const bound = body.contentControls[0];
  assert.equal(bound.tag, 'od:xpath=x1&od:ContentType=application/xhtml+xml');
  assert.deepEqual(bound.paragraphs.map((p) => p.text), texts);
  const xml = await main.getXml();
  assert.match(xml, /<w:hyperlink r:id="rId\d+"><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t>a link<\/w:t><\/w:r><\/w:hyperlink>/);
  assert.match(xml, /<w:pStyle w:val="Heading1"\/>/);
  assert.match(xml, /<w:b\/>/);
  assert.doesNotMatch(xml, /w:dataBinding/);
  const rels = await main.getRelationshipsPart().getXml();
  assert.match(rels, /<Relationship [^>]*Target="https:\/\/example.org\/x"[^>]*\/>/);
  assert.match(rels, /<Relationship [^>]*TargetMode="External"[^>]*Type="[^"]*\/hyperlink"/);
  assert.ok(Namespaces.HYPERLINK);
});

test('a run-level control takes the first paragraph\'s inline content; the blocks after it are noted', async () => {
  const { pkg, body, control } = await template('<p>First <i>one</i></p><p>Second</p><table><tr><td>c</td></tr></table>', { run: true });
  const before = body.paragraphs.length;
  const result = await pkg.customXmlParts.applyBindings({ html: { parser } });
  assert.equal(result.updated, 1);
  assert.equal(control.text, 'First one');
  assert.equal(body.paragraphs.length, before);
  assert.ok(result.notes.some((n) => /blocks after the first, in a run-level control \(2\)/.test(n)), JSON.stringify(result.notes));
  const again = await WordprocessingMLPackage.load(await pkg.save());
  assert.equal((await again.getBody()).contentControls[0].text, 'First one');
});

test('left and noted: an entry the XPaths part lacks; no parser where the runtime has no DOMParser', async () => {
  const missing = await template('<p>x</p>', { entry: 'x9' });
  const r1 = await missing.pkg.customXmlParts.applyBindings({ html: { parser } });
  assert.equal(r1.skipped, 1);
  assert.match(r1.notes[0], /od:xpath=x9: no such entry/);
  const bare = await template('<p>x</p>');
  const r2 = await bare.pkg.customXmlParts.applyBindings();
  assert.equal(r2.skipped, 1);
  assert.match(r2.notes[0], /needs an HTML parser/);
  assert.equal(bare.control.text, 'placeholder');
});

test('Word check 37 (2026-10-08): both Words keep the XHTML-bound controls, their content and tags, and add no w:dataBinding', async () => {
  for (const name of ['37b-word2010.docx', '37b-bound-word15.docx']) {
    const pkg = await WordprocessingMLPackage.load(await fixture(`check37/${name}`));
    const main = pkg.getMainDocumentPart();
    const controls = (await main.getBody()).contentControls;
    assert.equal(controls.length, 3, name);
    const block = controls.find((c) => c.tag === 'od:xpath=x1&od:ContentType=application/xhtml+xml');
    assert.deepEqual(block.paragraphs.map((p) => p.text).slice(0, 2), ['Delivery terms', 'Goods are delivered within 14 days of the order, see the terms.'], name);
    assert.equal(block.paragraphs[0].style, 'Heading 1', name);
    assert.equal(block.tables.length, 1, name);
    const inline = controls.find((c) => c.tag === 'od:xpath=x2&od:ContentType=application/xhtml+xml');
    assert.equal(inline.text, 'First paragraph in italics', name);
    const xml = await main.getXml();
    assert.equal((xml.match(/w:dataBinding/g) ?? []).length, 1, `${name}: the text-bound control's binding alone`);
  }
});
