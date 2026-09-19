import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OpcPackage, WordprocessingMLPackage, PresentationMLPackage, SpreadsheetMLPackage, ImagePart, HeaderPart, SlideLayoutPart, ZipPartStore, ContentTypes, Namespaces, PartName, XML_DECLARATION, createSlideSize } from '../dist/index.mjs';
import { fixture, bytesEqual, plain } from './helpers.mjs';

test('createPackage: a new docx saves and reloads', async () => {
  const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'LETTER', landscape: true });
  const main = pkg.getMainDocumentPart();
  assert.equal(main.partName.name, '/word/document.xml');
  assert.ok(main.styleDefinitionsPart);
  assert.ok(main.documentSettingsPart);
  // the theme part, as Word puts one in every document it creates (CR-001 section 14.6)
  assert.ok(main.themePart);
  assert.ok(pkg.docPropsCorePart && pkg.docPropsExtendedPart);
  main.contents.body.content.push({ name: { namespaceURI: Namespaces.NS_WORD12, localPart: 'p' }, value: { TYPE_NAME: 'org_docx4j_wml.P', content: [{ name: { namespaceURI: Namespaces.NS_WORD12, localPart: 'r' }, value: { TYPE_NAME: 'org_docx4j_wml.R', content: [{ name: { namespaceURI: Namespaces.NS_WORD12, localPart: 't' }, value: { TYPE_NAME: 'org_docx4j_wml.Text', value: 'Hello, core-ts' } }] } }] } });
  const bytes = await pkg.save();
  const store = new ZipPartStore(bytes);
  const names = [...store.partNames()];
  assert.equal(names[0], '[Content_Types].xml');
  assert.deepEqual(new Set(names), new Set(['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/_rels/document.xml.rels', 'word/styles.xml', 'word/settings.xml', 'word/theme/theme1.xml', 'docProps/core.xml', 'docProps/app.xml']));
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
  assert.equal(rels.getRelationshipByType(Namespaces.THEME).target, 'theme/theme1.xml');
  assert.equal(back.relationshipsPart.getRelationshipByType(Namespaces.DOCUMENT).target, 'word/document.xml');
  assert.equal(back.relationshipsPart.getRelationshipByType(Namespaces.PROPERTIES_CORE).target, 'docProps/core.xml');
});

test('createPackage: the settings part carries compatibilityMode 15, and a loaded document\'s is untouched', async () => {
  // Without it Word 365 opens the document in compatibility mode (CR-001 section 17).
  const pkg = await WordprocessingMLPackage.createPackage();
  const settings = pkg.getMainDocumentPart().documentSettingsPart;
  assert.deepEqual(settings.contents.compat.compatSetting.map((s) => [s.name, s.uri, s.val]), [
    ['compatibilityMode', 'http://schemas.microsoft.com/office/word', '15'],
    ['overrideTableStyleFontSizeAndJustification', 'http://schemas.microsoft.com/office/word', '1'],
    ['enableOpenTypeFeatures', 'http://schemas.microsoft.com/office/word', '1'],
    ['doNotFlipMirrorIndents', 'http://schemas.microsoft.com/office/word', '1'],
    ['differentiateMultirowTableHeaders', 'http://schemas.microsoft.com/office/word', '1'],
    ['useWord2013TrackBottomHyphenation', 'http://schemas.microsoft.com/office/word', '1'],
  ]);
  const xml = new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync('word/settings.xml'));
  assert.match(xml, /<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http:\/\/schemas.microsoft.com\/office\/word" w:val="15"\/>/);

  // the option, and docx4j's own behaviour back
  const at14 = await WordprocessingMLPackage.createPackage({ compatibilityMode: '14' });
  assert.equal(at14.getMainDocumentPart().documentSettingsPart.contents.compat.compatSetting[0].val, '14');
  const none = await WordprocessingMLPackage.createPackage({ compatibilityMode: '' });
  assert.deepEqual(none.getMainDocumentPart().documentSettingsPart.contents.compat.compatSetting.map((s) => s.name), ['overrideTableStyleFontSizeAndJustification', 'enableOpenTypeFeatures', 'doNotFlipMirrorIndents', 'differentiateMultirowTableHeaders', 'useWord2013TrackBottomHyphenation']);

  // an existing document's settings part is left exactly as it was
  const original = await fixture('loadAndSave.docx');
  const loaded = await WordprocessingMLPackage.load(original);
  const saved = new ZipPartStore(await loaded.save());
  assert.ok(bytesEqual(new ZipPartStore(original).loadSync('word/settings.xml'), saved.loadSync('word/settings.xml')));
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

test('PresentationMLPackage.createPackage: a new pptx saves and reloads', async () => {
  const pkg = await PresentationMLPackage.createPackage({ slideSize: 'SCREEN16x9' });
  const pp = pkg.getMainPresentationPart();
  assert.equal(pp.partName.name, '/ppt/presentation.xml');
  assert.equal(pkg.slideMasterParts.length, 1);
  assert.equal(pkg.slideLayoutParts.length, 1);
  assert.equal(pkg.slideParts.length, 1);
  assert.ok(pkg.themePart);
  assert.equal(pp.contents.sldSz.cx, 9144000);
  assert.equal(pp.contents.sldSz.cy, 5143500);
  assert.equal(pp.contents.sldSz.type, 'screen16x9');
  assert.equal(pp.contents.notesSz.cx, 6858000);

  const bytes = await pkg.save();
  const store = new ZipPartStore(bytes);
  // docx4j's part set (PresentationMLPackage.createPackage) plus the one slide
  assert.deepEqual(new Set(store.partNames()), new Set([
    '[Content_Types].xml', '_rels/.rels',
    'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels',
    'ppt/slideMasters/slideMaster1.xml', 'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    'ppt/slideLayouts/slideLayout1.xml', 'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    'ppt/theme/theme1.xml',
    'ppt/slides/slide1.xml', 'ppt/slides/_rels/slide1.xml.rels',
  ]));
  const ct = new TextDecoder().decode(store.loadSync('[Content_Types].xml'));
  for (const [name, type] of [
    ['/ppt/presentation.xml', ContentTypes.PRESENTATIONML_MAIN],
    ['/ppt/slideMasters/slideMaster1.xml', ContentTypes.PRESENTATIONML_SLIDE_MASTER],
    ['/ppt/slideLayouts/slideLayout1.xml', ContentTypes.PRESENTATIONML_SLIDE_LAYOUT],
    ['/ppt/slides/slide1.xml', ContentTypes.PRESENTATIONML_SLIDE],
    ['/ppt/theme/theme1.xml', ContentTypes.OFFICEDOCUMENT_THEME],
  ]) assert.ok(ct.includes(`PartName="${name}" ContentType="${type}"`), name);

  const back = await PresentationMLPackage.load(bytes);
  const slides = await back.getSlideParts();
  assert.equal(slides.length, 1);
  assert.equal(slides[0].partName.name, '/ppt/slides/slide1.xml');
  // the slide is on the layout, the layout on the master, and both master and presentation on the theme
  assert.ok(slides[0].slideLayoutPart instanceof SlideLayoutPart);
  const master = back.slideMasterParts[0];
  assert.equal(master.slideLayoutParts.length, 1);
  assert.equal(slides[0].slideLayoutPart.slideMasterPart, master);
  assert.ok(master.themePart);
  assert.equal(back.themePart, master.themePart, 'one theme part, related from both');
  // the sldIdLst and sldMasterIdLst entries point at those relationships
  const presentation = await back.getMainPresentationPart().getContents();
  assert.equal(presentation.sldIdLst.sldId.length, 1);
  assert.ok(presentation.sldIdLst.sldId[0].id >= 256 && presentation.sldIdLst.sldId[0].id <= 2147483647);
  assert.ok(presentation.sldMasterIdLst.sldMasterId[0].id > 2147483648);
  const sld = await slides[0].getContents();
  assert.equal(sld.cSld.spTree.nvGrpSpPr.cNvPr.id, 1);

  // the slide's placeholders: PowerPoint shows "Click to add title" only where the slide itself
  // has them (CR-001 section 17). Empty on a created slide, matching the layout's.
  const shapes = sld.cSld.spTree.spOrGrpSpOrGraphicFrame;
  assert.deepEqual(shapes.map((s) => [s.nvSpPr.nvPr.ph.type, s.nvSpPr.nvPr.ph.idx]), [['title', undefined], [undefined, 1]]);
  assert.deepEqual(shapes.map((s) => s.txBody.p.length), [1, 1]);
  assert.equal(shapes[0].txBody.p[0].egTextRun, undefined, 'empty, so PowerPoint shows the prompt');
  assert.ok(shapes[0].nvSpPr.cNvSpPr.spLocks.noGrp);
  assert.equal(shapes[0].spPr.xfrm, undefined, 'the geometry is the layout\'s');

  // and the layout defines them, with the geometry the slide inherits
  const layout = await slides[0].slideLayoutPart.getContents();
  const layoutShapes = layout.cSld.spTree.spOrGrpSpOrGraphicFrame;
  assert.deepEqual(layoutShapes.map((s) => [s.nvSpPr.nvPr.ph.type, s.nvSpPr.nvPr.ph.idx]), [['title', undefined], [undefined, 1]]);
  assert.equal(layoutShapes[0].spPr.xfrm.off.x, 685800);
  assert.ok(layoutShapes[0].spPr.xfrm.ext.cx > 0 && layoutShapes[1].spPr.xfrm.ext.cy > 0);
  assert.equal(layoutShapes[0].txBody.p[0].egTextRun[0].t, 'Click to edit Master title style');

  // a second slide, with text in its placeholders
  const added = await back.addSlide({ title: 'The title', body: 'First line\nSecond line' });
  assert.equal(added.partName.name, '/ppt/slides/slide2.xml');
  assert.equal(back.slideParts.length, 2);
  const addedShapes = added.contents.cSld.spTree.spOrGrpSpOrGraphicFrame;
  assert.equal(addedShapes[0].txBody.p[0].egTextRun[0].t, 'The title');
  assert.deepEqual(addedShapes[1].txBody.p.map((p) => p.egTextRun[0].t), ['First line', 'Second line']);
  const again = await PresentationMLPackage.load(await back.save());
  assert.equal((await again.getSlideParts()).length, 2);
  const reloaded = await (await again.getSlideParts())[1].getContents();
  assert.equal(reloaded.cSld.spTree.spOrGrpSpOrGraphicFrame[0].txBody.p[0].egTextRun[0].t, 'The title');
});

test('createSlideSize: docx4j\'s well-known sizes, portrait swapping the two', () => {
  assert.deepEqual({ ...createSlideSize('SCREEN4x3') }, { TYPE_NAME: 'org_pptx4j_pml.Presentation.SldSz', cx: 9144000, cy: 6858000, type: 'screen4x3' });
  assert.deepEqual({ ...createSlideSize('A4', false) }, { TYPE_NAME: 'org_pptx4j_pml.Presentation.SldSz', cx: 6858000, cy: 9906000, type: 'A4' });
  assert.throws(() => createSlideSize('B4JIS'), /No support for slide size B4JIS/);
});

test('SpreadsheetMLPackage.createPackage: a new xlsx saves and reloads', async () => {
  const pkg = await SpreadsheetMLPackage.createPackage();
  const wb = pkg.getWorkbookPart();
  assert.equal(wb.partName.name, '/xl/workbook.xml');
  // docx4j adds the one bookViews/workbookView, without which Excel 2010 could crash on print
  assert.equal(wb.contents.bookViews.workbookView.length, 1);
  const sheet = pkg.createWorksheetPart('Sheet1');
  assert.equal(sheet.partName.name, '/xl/worksheets/sheet1.xml');
  assert.deepEqual(plain(sheet.contents.sheetData), {});
  const entry = wb.contents.sheets.sheet[0];
  assert.equal(entry.name, 'Sheet1');
  assert.equal(entry.sheetId, 1);
  assert.equal(entry.id, sheet.sourceRelationship.id);

  const bytes = await pkg.save();
  const store = new ZipPartStore(bytes);
  assert.deepEqual(new Set(store.partNames()), new Set([
    '[Content_Types].xml', '_rels/.rels',
    'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet1.xml',
  ]));
  const ct = new TextDecoder().decode(store.loadSync('[Content_Types].xml'));
  assert.ok(ct.includes(`PartName="/xl/workbook.xml" ContentType="${ContentTypes.SPREADSHEETML_WORKBOOK}"`));
  assert.ok(ct.includes(`PartName="/xl/worksheets/sheet1.xml" ContentType="${ContentTypes.SPREADSHEETML_WORKSHEET}"`));
  assert.ok(new TextDecoder().decode(store.loadSync('xl/worksheets/sheet1.xml')).includes('<sheetData/>'));

  const back = await SpreadsheetMLPackage.load(bytes);
  const sheets = await back.getWorksheetParts();
  assert.equal(sheets.length, 1);
  assert.equal(sheets[0].partName.name, '/xl/worksheets/sheet1.xml');
  assert.equal(sheets[0].workbookPart, back.getWorkbookPart());
  assert.equal(back.getWorkbookPart().getWorksheet(0), sheets[0]);
  // a second sheet, and one inserted first
  const second = back.createWorksheetPart('Data');
  assert.equal(second.partName.name, '/xl/worksheets/sheet2.xml');
  back.createWorksheetPart('Cover', 0);
  assert.deepEqual(back.getWorkbookPart().contents.sheets.sheet.map((s) => s.name), ['Cover', 'Sheet1', 'Data']);
  assert.deepEqual(back.worksheetParts.map((p) => p.partName.name), ['/xl/worksheets/sheet3.xml', '/xl/worksheets/sheet1.xml', '/xl/worksheets/sheet2.xml']);
  const again = await SpreadsheetMLPackage.load(await back.save());
  assert.deepEqual((await again.getWorksheetParts()).map((p) => p.partName.name), ['/xl/worksheets/sheet3.xml', '/xl/worksheets/sheet1.xml', '/xl/worksheets/sheet2.xml']);
});

test('the xlsx fixture workbook: Excel\'s x15 absPath mc:AlternateContent has no Fallback', async () => {
  // docx4j CR-021's finding, pinned here: Excel writes
  //   <mc:AlternateContent><mc:Choice Requires="x15"><x15ac:absPath url="..."/></mc:Choice></mc:AlternateContent>
  // with no mc:Fallback. x15 is not a namespace the object model has a module for, so the
  // preprocessor of CR-001 section 5.6 finds no understood Choice and no Fallback, and drops the
  // element (ECMA-376 Part 3 10.2.1). The untouched part still round-trips byte for byte.
  const bytes = await fixture('loadAndSave.xlsx');
  const source = new ZipPartStore(bytes);
  const raw = new TextDecoder().decode(source.loadSync('xl/workbook.xml'));
  assert.ok(raw.includes('<mc:Choice Requires="x15">') && raw.includes('absPath'));
  assert.ok(!raw.includes('mc:Fallback'));

  const pkg = await SpreadsheetMLPackage.load(bytes);
  const workbook = await pkg.getWorkbookPart().getContents();
  assert.equal(workbook.alternateContent, undefined, 'the unresolvable Choice is dropped');
  assert.equal(workbook.sheets.sheet.length, 1);
  assert.equal(workbook.ignorable, 'x15 xr xr6 xr10 xr2');
  const marshalled = new TextDecoder().decode(await pkg.getWorkbookPart().getBytes());
  assert.ok(!marshalled.includes('absPath'), 'a re-marshalled workbook loses it, as Word and Excel do on open');

  const untouched = await SpreadsheetMLPackage.load(bytes);
  const saved = new ZipPartStore(await untouched.save());
  assert.ok(bytesEqual(source.loadSync('xl/workbook.xml'), saved.loadSync('xl/workbook.xml')));
});
