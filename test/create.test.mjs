import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OpcPackage, WordprocessingMLPackage, ImagePart, HeaderPart, ZipPartStore, ContentTypes, Namespaces, PartName, XML_DECLARATION } from '../dist/index.mjs';
import { fixture, bytesEqual, plain } from './helpers.mjs';

test('createPackage: a new docx saves and reloads', async () => {
  const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'LETTER', landscape: true });
  const main = pkg.getMainDocumentPart();
  assert.equal(main.partName.name, '/word/document.xml');
  assert.ok(main.styleDefinitionsPart);
  assert.ok(main.documentSettingsPart);
  assert.ok(pkg.docPropsCorePart && pkg.docPropsExtendedPart);
  main.contents.body.content.push({ name: { namespaceURI: Namespaces.NS_WORD12, localPart: 'p' }, value: { TYPE_NAME: 'org_docx4j_wml.P', content: [{ name: { namespaceURI: Namespaces.NS_WORD12, localPart: 'r' }, value: { TYPE_NAME: 'org_docx4j_wml.R', content: [{ name: { namespaceURI: Namespaces.NS_WORD12, localPart: 't' }, value: { TYPE_NAME: 'org_docx4j_wml.Text', value: 'Hello, core-ts' } }] } }] } });
  const bytes = await pkg.save();
  const store = new ZipPartStore(bytes);
  const names = [...store.partNames()];
  assert.equal(names[0], '[Content_Types].xml');
  assert.deepEqual(new Set(names), new Set(['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/_rels/document.xml.rels', 'word/styles.xml', 'word/settings.xml', 'docProps/core.xml', 'docProps/app.xml']));
  const ct = new TextDecoder().decode(store.loadSync('[Content_Types].xml'));
  assert.ok(ct.includes(`PartName="/word/document.xml" ContentType="${ContentTypes.WORDPROCESSINGML_DOCUMENT}"`));
  assert.ok(ct.includes(`PartName="/word/styles.xml"`));
  const back = await WordprocessingMLPackage.load(bytes);
  const doc = await back.getMainDocumentPart().getContents();
  assert.equal(doc.body.sectPr.pgSz.w, 15840);
  assert.equal(doc.body.sectPr.pgSz.h, 12240);
  assert.equal(doc.body.sectPr.pgSz.orient, 'landscape');
  assert.equal(doc.body.sectPr.pgMar.top, 1440);
  assert.equal(doc.body.content[0].value.content[0].value.content[0].value.value, 'Hello, core-ts');
  const styles = await back.getMainDocumentPart().styleDefinitionsPart.getContents();
  assert.ok(styles.style.some((s) => s.styleId === 'Normal'));
  assert.ok(styles.style.some((s) => s.styleId === 'Heading1'));
  const rels = back.getMainDocumentPart().relationshipsPart;
  assert.equal(rels.getRelationshipByType(Namespaces.STYLES).target, 'styles.xml');
  assert.equal(rels.getRelationshipByType(Namespaces.SETTINGS).target, 'settings.xml');
  assert.equal(back.relationshipsPart.getRelationshipByType(Namespaces.DOCUMENT).target, 'word/document.xml');
  assert.equal(back.relationshipsPart.getRelationshipByType(Namespaces.PROPERTIES_CORE).target, 'docProps/core.xml');
});

test('addTargetPart: image with a rels entry and a default content type', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('HelloWordOnline.docx'));
  const main = pkg.getMainDocumentPart();
  const png = await (await OpcPackage.load(await fixture('loadAndSave.docx'))).getPart('/word/media/image1.png').getBytes();
  const image = new ImagePart('/word/media/image1.png', ContentTypes.IMAGE_PNG);
  image.setBytes(png);
  const rel = main.addTargetPart(image);
  assert.equal(rel.target, 'media/image1.png');
  assert.equal(rel.type, Namespaces.IMAGE);
  assert.match(rel.id, /^rId\d+$/);
  assert.ok(pkg.getPart('/word/media/image1.png') === image);
  assert.equal(image.package, pkg);
  assert.ok(main.relationshipsPart.getRel('/word/media/image1.png') === rel);
  assert.ok(main.relationshipsPart.isATarget(PartName.of('/word/media/image1.png')));
  // same name again: rename (docx4j appends a counter to the proposed name: image1 -> image12)
  const image2 = new ImagePart('/word/media/image1.png', ContentTypes.IMAGE_PNG);
  image2.setBytes(png);
  const rel2 = main.addTargetPart(image2, 'RENAME_IF_NAME_EXISTS');
  assert.equal(image2.partName.name, '/word/media/image12.png');
  assert.equal(rel2.target, 'media/image12.png');
  assert.notEqual(rel2.id, rel.id);
  // reuse
  const image3 = new ImagePart('/word/media/image1.png', ContentTypes.IMAGE_PNG);
  const rel3 = main.addTargetPart(image3, 'REUSE_EXISTING');
  assert.equal(rel3, rel);
  // a header with typed contents
  const header = new HeaderPart('/word/header1.xml');
  header.setContents({ content: [] });
  const hrel = main.addTargetPart(header);
  assert.equal(hrel.type, Namespaces.HEADER);

  const bytes = await pkg.save();
  const back = await WordprocessingMLPackage.load(bytes);
  const img = back.getPart('/word/media/image1.png');
  assert.ok(img instanceof ImagePart);
  assert.equal(img.contentType, ContentTypes.IMAGE_PNG);
  assert.ok(bytesEqual(await img.getBytes(), png));
  assert.ok(back.getPart('/word/media/image12.png'));
  assert.equal(back.contentTypeManager.getContentType('/word/media/image1.png'), ContentTypes.IMAGE_PNG);
  assert.equal(back.getMainDocumentPart().headerParts.length, 1);
  const hdr = await back.getMainDocumentPart().headerParts[0].getContents();
  assert.equal(hdr.TYPE_NAME, 'org_docx4j_wml.Hdr');
  const store = new ZipPartStore(bytes);
  const hdrXml = new TextDecoder().decode(store.loadSync('word/header1.xml'));
  assert.ok(hdrXml.startsWith(XML_DECLARATION));
});

test('removePart removes the part, its relationship and its own targets', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('loadAndSave.docx'));
  const main = pkg.getMainDocumentPart();
  const chart = pkg.getPart('/word/charts/chart1.xml');
  assert.ok(chart.relationshipsPart.size > 0);
  const embedded = chart.relationshipsPart.getPart(chart.relationshipsPart.list[0]);
  const removed = main.relationshipsPart.removePart(chart.partName);
  assert.ok(removed.some((n) => n.equals('/word/charts/chart1.xml')));
  assert.equal(pkg.getPart('/word/charts/chart1.xml'), undefined);
  assert.equal(pkg.getPart(embedded.partName), undefined, 'the chart-only targets went too');
  assert.equal(main.relationshipsPart.getRel('/word/charts/chart1.xml'), undefined);
  const back = await OpcPackage.load(await pkg.save());
  assert.equal(back.getPart('/word/charts/chart1.xml'), undefined);
  assert.equal(back.parts.size, pkg.parts.size);
  assert.deepEqual(plain(back.getMainDocumentPart().relationshipsPart.list), plain(main.relationshipsPart.list));
});
