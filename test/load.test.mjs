import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  OpcPackage, WordprocessingMLPackage, PresentationMLPackage, SpreadsheetMLPackage,
  MainDocumentPart, StyleDefinitionsPart, HeaderPart, FooterPart, ImagePart, ChartPart, EmbeddedPackagePart,
  DefaultXmlPart, CommentsExtensiblePart, CustomXmlDataStoragePart, CustomXmlDataStoragePropertiesPart, DocPropsCorePart, DocPropsExtendedPart,
  CommentsExtendedPart, PeoplePart, MainPresentationPart, SlidePart, WorkbookPart, WorksheetPart, SharedStringsPart,
  ZipPartStore, ContentTypes, Namespaces, Docx4JException, PartUnmarshalException,
} from '../dist/index.mjs';
import { fixture } from './helpers.mjs';

test('load a Word-saved docx: parts, classes, shortcuts', async () => {
  const bytes = await fixture('loadAndSave.docx');
  const pkg = await OpcPackage.load(bytes);
  assert.ok(pkg instanceof WordprocessingMLPackage);
  const main = pkg.getMainDocumentPart();
  assert.ok(main instanceof MainDocumentPart);
  assert.equal(main.partName.name, '/word/document.xml');
  assert.ok(main.styleDefinitionsPart instanceof StyleDefinitionsPart);
  assert.ok(main.themePart, 'theme');
  assert.ok(main.commentsPart, 'comments');
  assert.ok(main.commentsExtendedPart instanceof CommentsExtendedPart);
  assert.ok(main.peoplePart instanceof PeoplePart);
  assert.ok(main.footnotesPart && main.endnotesPart);
  assert.equal(main.headerParts.length, 3);
  assert.ok(main.headerParts.every((h) => h instanceof HeaderPart));
  assert.equal(main.footerParts.length, 3);
  assert.ok(main.footerParts.every((f) => f instanceof FooterPart));
  assert.ok(pkg.docPropsCorePart instanceof DocPropsCorePart);
  assert.ok(pkg.docPropsExtendedPart instanceof DocPropsExtendedPart);

  const image = pkg.getPart('/word/media/image1.png');
  assert.ok(image instanceof ImagePart);
  assert.equal(image.contentType, ContentTypes.IMAGE_PNG);
  assert.equal(image.relationshipType, Namespaces.IMAGE);
  assert.ok(pkg.getPart('/word/media/image2.svg') instanceof ImagePart);
  assert.ok(pkg.getPart('/word/charts/chart1.xml') instanceof ChartPart);
  assert.ok(pkg.getPart('/word/embeddings/Microsoft_Excel_Worksheet.xlsx') instanceof EmbeddedPackagePart);
  // unknown relationship types keep their XML as a DOM part
  const ext = pkg.getPart('/word/commentsExtensible.xml');
  assert.ok(ext instanceof CommentsExtensiblePart, 'typed since objects 0.1.3');
  assert.equal(pkg.getMainDocumentPart().commentsExtensiblePart, ext);
  assert.equal(ext.relationshipType, Namespaces.COMMENTS_EXTENSIBLE);
  assert.ok(pkg.getPart('/docMetadata/LabelInfo.xml') instanceof DefaultXmlPart);
  // custom XML by itemId
  const cx = pkg.getPart('/customXml/item1.xml');
  assert.ok(cx instanceof CustomXmlDataStoragePart);
  assert.ok(cx.relationshipsPart.getPart(cx.relationshipsPart.list[0]) instanceof CustomXmlDataStoragePropertiesPart);
  assert.equal(pkg.customXmlDataStorageParts.size, 1);
  assert.ok(pkg.customXmlDataStorageParts.get(cx.itemId) === cx);
  assert.match(cx.itemId, /^\{[0-9a-f-]+\}$/);
  // every zip entry that is not a rels part or content types is a part
  const store = new ZipPartStore(bytes);
  const names = [...store.partNames()].filter((n) => n !== '[Content_Types].xml' && !n.endsWith('.rels'));
  for (const n of names) assert.ok(pkg.getPart('/' + n), `part ${n} loaded`);
  assert.equal(pkg.parts.size, names.length);
  // nothing unmarshalled by load
  for (const p of pkg.parts) if (p.isUnmarshalled !== undefined) assert.equal(p.isUnmarshalled, false, p.partName.name);
  // contents on demand
  const doc = await main.getContents();
  assert.equal(doc.TYPE_NAME, 'org_docx4j_wml.Document');
  assert.ok(doc.body.content.length > 0);
  assert.equal(main.isUnmarshalled, true);
  assert.equal(main.contents, doc);
});

test('load: main part with a non-standard name (Word Online)', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('HelloWordOnline.docx'));
  assert.equal(pkg.getMainDocumentPart().partName.name, '/word/document22.xml');
  assert.ok(pkg.getMainDocumentPart().styleDefinitionsPart);
  const doc = await pkg.getMainDocumentPart().getContents();
  assert.equal(doc.TYPE_NAME, 'org_docx4j_wml.Document');
});

test('load: headers without rels parts, hyperlink rels skipped', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('header-no-rels.docx'));
  const main = pkg.getMainDocumentPart();
  assert.equal(main.headerParts.length, 3);
  for (const h of main.headerParts) assert.equal(h.relationshipsPart, undefined);
  const dupe = await WordprocessingMLPackage.load(await fixture('hyperlink_dupe.docx'));
  assert.ok(dupe.getMainDocumentPart());
  assert.equal(dupe.customXmlDataStorageParts.size, 1);
});

test('load: pptx and xlsx', async () => {
  const ppt = await OpcPackage.load(await fixture('loadAndSave.pptx'));
  assert.ok(ppt instanceof PresentationMLPackage);
  const pres = ppt.getMainPresentationPart();
  assert.ok(pres instanceof MainPresentationPart);
  assert.equal(pres.slideParts.length, 3);
  assert.ok(pres.slideParts[0] instanceof SlidePart);
  assert.ok(pres.themePart);
  const sld = await pres.slideParts[0].getContents();
  assert.equal(sld.TYPE_NAME, 'org_pptx4j_pml.Sld');

  const xlsx = await OpcPackage.load(await fixture('loadAndSave.xlsx'));
  assert.ok(xlsx instanceof SpreadsheetMLPackage);
  const wb = xlsx.getWorkbookPart();
  assert.ok(wb instanceof WorkbookPart);
  assert.ok(wb.sharedStringsPart instanceof SharedStringsPart);
  assert.ok(wb.stylesPart && wb.themePart && wb.calcChainPart);
  assert.equal(wb.worksheetParts.length, 1);
  assert.ok(wb.worksheetParts[0] instanceof WorksheetPart);
  const ws = await wb.worksheetParts[0].getContents();
  assert.equal(ws.TYPE_NAME, 'org_xlsx4j_sml.Worksheet');
  assert.ok(xlsx.getPart('/xl/drawings/vmlDrawing1.vml'));
  assert.ok(xlsx.getPart('/xl/media/image1.png') instanceof ImagePart);
});

test('load: the wrong package class throws', async () => {
  const xlsx = await fixture('loadAndSave.xlsx');
  await assert.rejects(() => WordprocessingMLPackage.load(xlsx), Docx4JException);
});

// CR-002 section 27 (item 9 of the E4 release): a part that cannot be unmarshalled names itself, and
// carries the location jsonix 3.4.0 found rather than leaving a caller to parse the message. The
// editor's parts panel shows a person what its save would refuse; "which part" is the first thing
// they need and the runtime cannot know it.
test('an unmarshal failure names the part, with the location and line', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const styles = pkg.getMainDocumentPart().styleDefinitionsPart;
  styles.setXml((await styles.getXml()).replace('<w:sz w:val="22" />', '<w:sz w:val="big" />'));

  const error = await styles.getContents().then(() => undefined, (e) => e);
  assert.ok(error instanceof PartUnmarshalException, `expected PartUnmarshalException, got ${error?.constructor?.name}`);
  assert.equal(error.partName, '/word/styles.xml');
  assert.match(error.message, /^\/word\/styles\.xml: /, 'the part comes first in the message');
  assert.match(error.message, /Argument \[NaN\] must be an integer/, 'and the runtime\'s own message follows');
  assert.equal(error.location, 'w:sz/@w:val', 'the attribute, not cut at its own prefix separator');
  assert.equal(typeof error.line, 'number');
  assert.ok(error.line > 0);
  assert.equal(typeof error.column, 'number');
  // the value as written, and where: the runtime's own message says NaN
  assert.ok(error.message.startsWith(`/word/styles.xml: w:sz/@w:val="big" (line ${error.line}, column ${error.column}): `), error.message);
  assert.ok(error.cause, 'the runtime error is the cause');

  // readContents reports the same way, since a private read fails for the same reasons
  const second = await styles.readContents().then(() => undefined, (e) => e);
  assert.ok(second instanceof PartUnmarshalException);
  assert.equal(second.partName, '/word/styles.xml');
});

// ECMA-376 Transitional's ST_OnOff admits "on" and "off" beside true, false, 1 and 0, and pandoc writes
// them: w:tblHeader w:val="on" on a table's header row, w:wordWrap w:val="off" in its Source Code style.
// @docx4j/jsonix 3.4.0 accepted only the xsd:boolean spellings, so the document part threw
// PartUnmarshalException at w:tblHeader/@w:val and the file could not be opened; 3.4.1 (jsonix-CR-007)
// reads them, and @docx4j/generated-objects-ts 0.3.2, this package's floor, requires it. They are
// written back as true and false.
test('a document whose on/off values are "on" and "off" loads, reads them as booleans and saves', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('onoff/pandoc-code.docx'));
  await pkg.unmarshalAll();
  const read = async (p) => {
    const body = await p.getBody();
    const header = body.tables[0].rows[0].tr.trPr.cnfStyleOrDivIdOrGridBefore.find((i) => i.name.localPart === 'tblHeader');
    const style = p.getMainDocumentPart().styleDefinitionsPart.contents.style.find((s) => s.styleId === 'SourceCode');
    return { tblHeader: header.value.val, wordWrap: style.pPr.wordWrap.val };
  };
  assert.deepEqual(await read(pkg), { tblHeader: true, wordWrap: false });

  const saved = await pkg.save();
  const again = await WordprocessingMLPackage.load(saved);
  await again.unmarshalAll();
  assert.deepEqual(await read(again), { tblHeader: true, wordWrap: false });
  assert.match(await again.getMainDocumentPart().getXml(), /<w:tblHeader w:val="true"\s*\/>/);
  assert.match(await again.getMainDocumentPart().styleDefinitionsPart.getXml(), /<w:wordWrap w:val="false"\s*\/>/);
});

// Not well-formed: the parser's error names no part, so the part's does (Node: xmldom's locator gives
// the line and column; a browser's message says them itself, the case below).
test('a part that is not well-formed names itself, with the line and column', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const styles = pkg.getMainDocumentPart().styleDefinitionsPart;
  styles.setXml((await styles.getXml()).replace('<w:sz w:val="22" />', '<w:sz w:val="22" <'));
  const error = await styles.getContents().then(() => undefined, (e) => e);
  assert.ok(error instanceof PartUnmarshalException, `expected PartUnmarshalException, got ${error?.constructor?.name}`);
  assert.equal(error.partName, '/word/styles.xml');
  assert.match(error.message, /^\/word\/styles\.xml: Malformed XML \(line \d+, column \d+\): Unescaped '<' not allowed in attributes values$/);
  assert.equal(typeof error.line, 'number');
  assert.equal(typeof error.column, 'number');
  assert.equal(error.location, undefined, 'no element: the parser stopped before one');
  assert.equal(styles.isUnmarshalled, false);
});

// Jason, 2026-10-17: a styles part that cannot be read refuses the document, since nothing in it can be
// given its formatting, and the error says that first, then what was wrong where. Put right, it opens.
test('a document whose styles part cannot be read is refused, saying why; put right, it opens', async () => {
  const made = await WordprocessingMLPackage.createPackage();
  const good = await made.getMainDocumentPart().styleDefinitionsPart.getXml();
  made.getMainDocumentPart().styleDefinitionsPart.setXml(good.replace('<w:sz w:val="22" />', '<w:sz w:val="big" />'));
  const pkg = await WordprocessingMLPackage.load(await made.save());

  const error = await pkg.getBody().then(() => undefined, (e) => e);
  assert.ok(error instanceof PartUnmarshalException, `expected PartUnmarshalException, got ${error?.constructor?.name}`);
  assert.match(error.message, /^The document cannot be opened: its styles part could not be read, and every paragraph's formatting depends on it\. \/word\/styles\.xml: w:sz\/@w:val="big" \(line \d+, column \d+\): Argument \[NaN\] must be an integer/);
  assert.equal(error.partName, '/word/styles.xml');
  assert.equal(error.location, 'w:sz/@w:val');
  assert.equal(typeof error.line, 'number');
  assert.ok(error.cause instanceof PartUnmarshalException, "the part's own error is the cause");
  await assert.rejects(() => pkg.getPropertyResolver(), (e) => e.message === error.message, 'the resolver says the same');

  // the failure is not kept: the part put right, the document opens
  pkg.getMainDocumentPart().styleDefinitionsPart.setXml(good);
  const body = await pkg.getBody();
  assert.ok(body.paragraphs.length >= 0);
  assert.ok(pkg.propertyResolver);
});

// CR-002 section 33, found by the editor's E4.b: a browser's DOMParser does not throw on malformed
// XML. Chromium hands back the document up to the error with an XHTML parsererror first in the root,
// Firefox a parsererror root in its own namespace, and parseXml passed either on, so a part with a
// stray `<` unmarshalled truncated. Node's xmldom throws, so the browsers are stood in for through
// Jsonix.DOM.use with parsers that return what theirs do.
test('parseXml refuses the parsererror document a browser returns for malformed XML', async () => {
  const { Jsonix } = await import('@docx4j/generated-objects-ts');
  const { parseXml, WordprocessingMLPackage } = await import('../dist/index.mjs');
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const XmldomParser = Jsonix.DOM.getImplementation().DOMParser;
  const returning = (xml) => ({
    DOMParser: class { parseFromString() { return new XmldomParser().parseFromString(xml, 'application/xml'); } },
  });
  const chromium = `<w:settings xmlns:w="${W}"><parsererror xmlns="http://www.w3.org/1999/xhtml" style="display: block">`
    + '<h3>This page contains the following errors:</h3><div style="font-family:monospace">error on line 3 at column 19: '
    + 'attributes construct error\n</div><h3>Below is a rendering of the page up to the first error.</h3></parsererror>'
    + '<w:zoom w:percent="100"/></w:settings>';
  const firefox = '<parsererror xmlns="http://www.mozilla.org/newlayout/xml/parsererror.xml">XML Parsing Error: not well-formed\n'
    + 'Location: about:blank\nLine Number 3, Column 19:<sourcetext>  &lt;w:bad &lt;</sourcetext></parsererror>';

  // made before a stand-in parser is installed: creating a package parses its default parts
  const pkg = await WordprocessingMLPackage.createPackage();
  const settings = pkg.getMainDocumentPart().documentSettingsPart;
  settings.setXml(`<w:settings xmlns:w="${W}"><w:zoom w:percent="100"/><w:bad <</w:settings>`);
  try {
    Jsonix.DOM.use(returning(chromium));
    assert.throws(() => parseXml('<a/>'), /^Docx4JException: Malformed XML: error on line 3 at column 19: attributes construct error$/);
    // and through a part, which is where the content was lost
    await assert.rejects(() => settings.getContents(), /Malformed XML: error on line 3 at column 19/);
    assert.equal(settings.isUnmarshalled, false, 'nothing was unmarshalled');

    Jsonix.DOM.use(returning(firefox));
    assert.throws(() => parseXml('<a/>'), /Malformed XML: XML Parsing Error: not well-formed Location: about:blank Line Number 3, Column 19:/);
  } finally {
    Jsonix.DOM.use(null);
  }
  // a document that parsed is untouched, an element of that name in another namespace included
  assert.equal(parseXml('<a xmlns="urn:x"><parsererror/></a>').documentElement.localName, 'a');
});
