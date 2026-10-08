// The files for Word check 37 (CR-005 section 8.2; docx4j-ts-editor ED-005 proposal 44 part 5): what
// Word 2010 and Word 15 show for a control bound to escaped XHTML before and after the engine's bind.
// Run from the repository root after `npm run build`:
//
//   node test/fixtures/check37/build.mjs
//
// writes into this directory:
//   37a-template.docx   the template as authored: a block-level control and a run-level one, each tagged
//                       od:xpath=xN&od:ContentType=application/xhtml+xml with no w:dataBinding, holding
//                       placeholder text; a plain text control with a w:dataBinding to the same markup node,
//                       for contrast (Word shows the escaped markup as text there)
//   37b-bound.docx      the same after applyBindings({ html: { parser } }): the markup converted into the
//                       two controls (a heading, bold, a link, a list, a table; the run-level one the
//                       first paragraph's inlines)
// Open each in Word 2010 and Word 15, look, save as 37a-word2010.docx, 37a-word15.docx, 37b-word2010.docx,
// 37b-word15.docx beside them (README.md says what to look for).
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseHTML } from 'linkedom';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const escape = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const BLOCK = '<h1>Delivery terms</h1><p>Goods are delivered <b>within 14 days</b> of the order, see <a href="https://example.org/terms">the terms</a>.</p><ul><li>Standard delivery is free.</li><li>Express delivery costs extra.</li></ul><table><tr><th>Service</th><th>Days</th></tr><tr><td>Standard</td><td>14</td></tr><tr><td>Express</td><td>2</td></tr></table>';
const INLINE = '<p>First paragraph <i>in italics</i></p><p>A second paragraph, which a run-level control cannot hold.</p>';

async function template() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  body.insertParagraph('Check 37: escaped XHTML bound through od:xpath entries (no w:dataBinding on the first two controls).', 'End');
  body.insertParagraph('1. A block-level control (should show the converted blocks after the bind):', 'End');
  const blockHolder = body.insertParagraph('XHTML placeholder (block-level)', 'End');
  const block = blockHolder.insertContentControl('RichText');
  body.insertParagraph('2. A run-level control (should show the first paragraph\'s text after the bind): [', 'End');
  const inlineHolder = body.paragraphs[body.paragraphs.length - 1];
  inlineHolder.insertText('XHTML placeholder (run-level)', 'End');
  const inline = body.search('XHTML placeholder (run-level)')[0].insertContentControl('RichText');
  inlineHolder.insertText('] end of the paragraph.', 'End');
  body.insertParagraph('3. For contrast, a plain text control with a w:dataBinding to the same node (Word shows the markup as text):', 'End');
  body.insertParagraph('text binding placeholder', 'End');
  const textControl = body.search('text binding placeholder')[0].insertContentControl('PlainText');
  const data = pkg.customXmlParts.add(`<data><body>${escape(BLOCK)}</body><inline>${escape(INLINE)}</inline></data>`);
  pkg.customXmlParts.add(`<xpaths xmlns="http://opendope.org/xpaths"><xpath id="x1"><dataBinding xmlns:w="${W}" w:storeItemID="${data.id}" w:xpath="/data/body" w:prefixMappings=""/></xpath><xpath id="x2"><dataBinding xmlns:w="${W}" w:storeItemID="${data.id}" w:xpath="/data/inline" w:prefixMappings=""/></xpath></xpaths>`);
  block.tag = 'od:xpath=x1&od:ContentType=application/xhtml+xml';
  block.title = 'XHTML block';
  inline.tag = 'od:xpath=x2&od:ContentType=application/xhtml+xml';
  inline.title = 'XHTML inline';
  await pkg.customXmlParts.load();
  textControl.xmlMapping.setMapping('/data/body', '', data);
  textControl.title = 'Text binding';
  return pkg;
}

const a = await template();
await writeFile(join(here, '37a-template.docx'), await a.save());
const b = await template();
const result = await b.customXmlParts.applyBindings({ html: { parser: (h) => parseHTML(h).document } });
console.log('bound:', JSON.stringify(result));
await writeFile(join(here, '37b-bound.docx'), await b.save());
console.log('written 37a-template.docx and 37b-bound.docx in', here);
