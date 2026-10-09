// CR-002 section 41: pkg.styles.ensure falls back to Word's built-in styles as docx4j's KnownStyles.xml
// carries them, for an id its other sources lack (the editor's Markdown headings 5 and 6, the HTML code
// styles, Quote), with what each is based on and linked to, and without the list numbering the file
// left on the headings.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { WordprocessingMLPackage } from '../dist/index.mjs';

const styleIn = (styles, id) => (styles.style ?? []).find((s) => s.styleId === id);

test('the known styles: the eleven the editor carries, and the rest of the 164, splice; what no source has still throws', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const eleven = ['Heading5', 'Heading6', 'Heading5Char', 'Heading6Char', 'Quote', 'QuoteChar', 'IntenseQuote', 'IntenseQuoteChar', 'ListParagraph', 'HTMLCode', 'HTMLPreformatted', 'HTMLPreformattedChar'];
  const added = await pkg.styles.ensure(eleven);
  assert.deepEqual([...added].sort(), [...eleven].sort());
  for (const id of ['Heading7', 'TOCHeading', 'BookTitle', 'ListBullet', 'Index1', 'TOC9', 'Bibliography']) assert.ok((await pkg.styles.ensure(id)).includes(id), id);
  assert.deepEqual(await pkg.styles.ensure('Subtitle'), [], 'a default the created document has already');
  const styles = await pkg.getMainDocumentPart().styleDefinitionsPart.getContents();
  for (const id of [...eleven, 'Heading7', 'TOCHeading', 'BookTitle', 'ListBullet', 'Index1', 'TOC9', 'Bibliography', 'Subtitle']) assert.ok(styleIn(styles, id), id);
});

test('ensure splices a heading the defaults lack, with its linked character style, and without the file\'s w:numPr; the document then styles a paragraph with it, saved and reloaded', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const p = body.insertParagraph('A fifth-level heading', 'End');
  assert.deepEqual(await pkg.styles.ensure('Heading5'), ['Heading5Char', 'Heading5'], 'the linked style first, then the heading; Normal is there already');
  const styles = await pkg.getMainDocumentPart().styleDefinitionsPart.getContents();
  const heading = styleIn(styles, 'Heading5');
  assert.equal(heading.name.val, 'heading 5');
  assert.equal(heading.basedOn.val, 'Normal');
  assert.equal(heading.link.val, 'Heading5Char');
  assert.equal(heading.pPr.numPr, undefined, 'the list numbering the file left on it is dropped');
  assert.equal(heading.pPr.outlineLvl.val, 4);
  assert.equal(heading.rPr.color.val, '243F60', 'the file\'s Heading 5: the major theme face in the accent colour');
  p.styleId = 'Heading5';
  assert.equal(p.style, 'Heading 5', 'the display name, as Office JS reports it');
  assert.deepEqual(await pkg.styles.ensure(['Heading5', 'Heading6']), ['Heading6Char', 'Heading6'], 'already there, nothing again');
  const again = await WordprocessingMLPackage.load(await pkg.save());
  const reloaded = (await again.getBody()).paragraphs[0];
  assert.equal(reloaded.styleId, 'Heading5');
  const xml = await again.getMainDocumentPart().styleDefinitionsPart.getXml();
  const spliced = /<w:style [^>]*w:styleId="Heading5"[^>]*>.*?<\/w:style>/s.exec(xml);
  assert.ok(spliced, 'the heading is in the saved part');
  assert.doesNotMatch(spliced[0], /<w:numPr>/, 'and carries no list');
});

test('the HTML code styles and Quote: a character style on the default paragraph font; a paragraph style with its Char; the document\'s own definition wins', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  assert.deepEqual(await pkg.styles.ensure(['HTMLCode', 'HTMLPreformatted', 'Quote']), ['HTMLCode', 'HTMLPreformattedChar', 'HTMLPreformatted', 'QuoteChar', 'Quote']);
  const styles = await pkg.getMainDocumentPart().styleDefinitionsPart.getContents();
  assert.equal(styleIn(styles, 'HTMLCode').type, 'character');
  assert.equal(styleIn(styles, 'HTMLCode').basedOn.val, 'DefaultParagraphFont');
  assert.equal(styleIn(styles, 'Quote').name.val, 'Quote');
  // the document's own Quote stays as it is
  styleIn(styles, 'Quote').name.val = 'My Quote';
  assert.deepEqual(await pkg.styles.ensure('Quote'), []);
  assert.equal(styleIn(styles, 'Quote').name.val, 'My Quote');
  // a list style comes without its numbering, which the file does not carry either
  assert.deepEqual(await pkg.styles.ensure('ListBullet'), ['ListBullet']);
  assert.equal(styleIn(styles, 'ListBullet').pPr.numPr, undefined);
  await assert.rejects(pkg.styles.ensure('NoSuchStyle'), /No definition to splice for the style NoSuchStyle: .*KnownStyles\.xml/);
});
