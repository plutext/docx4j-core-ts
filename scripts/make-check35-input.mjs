// The input of test/README.md check 35 (the editor's request of 2026-10-03, at Jason's word): rich text
// content controls, unbound, to map in Word 15's XML Mapping Pane and through Word 2010's VBA, beside
// empty paragraphs for controls the pane inserts already mapped, and a rich text control bound with
// `w:dataBinding` as the engine and the editor bind one (1f), its text not its node's. The custom XML
// part is <data><name>Ann</name><rich1/><rich2/><bound>Bea</bound></data>.
//
//   npm run build && node scripts/make-check35-input.mjs
//
// writes test/fixtures/revisions/check35/35-input.docx.
//
// Since CR-002 section 40 (2026-10-04) `setMapping` writes `w15:dataBinding` on a rich text control, so
// this script no longer reproduces 1f's `w:dataBinding`: the committed 35-input.docx is what Word was given.
import { mkdir, writeFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from '../dist/index.mjs';

const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
const body = pkg.body;
const part = pkg.customXmlParts.add('<data><name>Ann</name><rich1/><rich2/><bound>Bea</bound></data>');
const para = (text) => body.insertParagraph(text, 'End');

para('Check 35: rich text content controls and custom XML mapping. The custom XML part is '
  + '<data><name>Ann</name><rich1/><rich2/><bound>Bea</bound></data>. Do each case as test/README.md check 35 says.');

// 1a: a run-level rich text control titled "1a" around "quick brown", "brown" bold.
const a = para('1a. The quick brown fox.');
a.search('brown')[0].font.bold = true;
a.search('quick brown')[0].insertContentControl('RichText').title = '1a';

// 1b: a block-level rich text control titled "1b" holding two paragraphs, one word bold.
para('1b. A block-level rich text control follows, holding two paragraphs:');
const b = para('First paragraph, with a bold word.');
b.search('bold')[0].font.bold = true;
const control = b.insertContentControl('RichText');
control.title = '1b';
control.insertParagraph('Second paragraph.', 'End');

// 1c and 1d: an empty paragraph each, for a rich text control the pane inserts already mapped.
para('1c. The empty paragraph below:');
para('');
para('1d. The empty paragraph below:');
para('');

// 1f: a run-level rich text control titled "1f", bound with `w:dataBinding` to /data/bound ("Bea"), as
// `XmlMapping.setMapping` binds one, then given the text "Old text", so that what Word shows on open
// says whether it honours the binding.
const f = para('1f. A bound rich text control: here.');
const bound = f.search('here')[0].insertContentControl('RichText');
bound.title = '1f';
await pkg.customXmlParts.load();
if (!bound.xmlMapping.setMapping('/data/bound', undefined, part)) throw new Error('1f: not mapped');
bound.insertText('Old text', 'Replace');
// The engine's edit wrote "Old text" to the node too; the node keeps "Bea".
const node = part.selectSingleNode('/data/bound');
node.nodeValue = 'Bea';
node.ownerPart.touch();

para('After.');

const dir = new URL('../test/fixtures/revisions/check35/', import.meta.url);
await mkdir(dir, { recursive: true });
await writeFile(new URL('35-input.docx', dir), await pkg.save());
