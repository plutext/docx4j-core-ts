import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  OpcPackage, WordprocessingMLPackage, PresentationMLPackage, SpreadsheetMLPackage,
  MainDocumentPart, StyleDefinitionsPart, HeaderPart, FooterPart, ImagePart, ChartPart, EmbeddedPackagePart,
  DefaultXmlPart, CommentsExtensiblePart, CustomXmlDataStoragePart, CustomXmlDataStoragePropertiesPart, DocPropsCorePart, DocPropsExtendedPart,
  CommentsExtendedPart, PeoplePart, MainPresentationPart, SlidePart, WorkbookPart, WorksheetPart, SharedStringsPart,
  ZipPartStore, ContentTypes, Namespaces, Docx4JException,
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
