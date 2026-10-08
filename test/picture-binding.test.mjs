// CR-005 section 9: picture bindings. (a) Word's picture content control, w:sdtPr/w:picture with a
// w:dataBinding; (b) od:Handler=picture on a rich text control's tag, bound through its od:xpath
// entry, with and without a width parameter. The node holds base64 image data.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { WordprocessingMLPackage, addImage, base64Decode, pictureWidthOf, tagParamsOf } from '../dist/index.mjs';
import { pngOf, bytesEqual } from './helpers.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const EMU_PER_POINT = 12700;
/** 4 x 3 px at 96 dpi: 38100 x 28575 EMU. */
const SMALL = pngOf(4, 3, [255, 0, 0]);
/** The bound image, 6 x 3 px: 57150 x 28575 EMU. */
const BLUE = pngOf(6, 3, [0, 0, 255]);
/** 400 x 3 px: 3810000 x 28575 EMU, 6000 twips wide. */
const WIDE = pngOf(400, 3, [0, 128, 0]);
/** An SVG document, which has no pixel size. */
const SVG = Buffer.from('<?xml version="1.0"?>\n<!-- a mark -->\n<svg xmlns="http://www.w3.org/2000/svg" width="40" height="30"><rect width="40" height="30" fill="blue"/></svg>').toString('base64');

/** A paragraph holding a floating picture (wp:anchor), as Word writes one, over an image relationship. */
function anchoredParagraphXml(relId, cx, cy, id) {
  return `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`
    + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${id}" name="Floating picture"/>`
    + `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`
    + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Floating picture"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`
    + `</wp:anchor></w:drawing></w:r></w:p>`;
}

/** A new document with a data part holding the images and an XPaths part with entries x1 (blue), x2 (wide), x3 (empty), x4 (not an image). */
async function document() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const data = pkg.customXmlParts.add(`<data><logo>${BLUE}</logo><wide>${WIDE}</wide><empty></empty><text>not an image</text><svg>${SVG}</svg></data>`);
  const entry = (id, xpath) => `<xpath id="${id}"><dataBinding xmlns:w="${W}" w:storeItemID="${data.id}" w:xpath="${xpath}" w:prefixMappings=""/></xpath>`;
  pkg.customXmlParts.add(`<xpaths xmlns="http://opendope.org/xpaths">${entry('x1', '/data/logo')}${entry('x2', '/data/wide')}${entry('x3', '/data/empty')}${entry('x4', '/data/text')}${entry('x5', '/data/svg')}</xpaths>`);
  await pkg.customXmlParts.load();
  return { pkg, body, data };
}

const imageParts = (pkg) => [...pkg.parts.values()].filter((p) => p.partName.name.startsWith('/word/media/')).map((p) => p.partName.name).sort();

async function bytesOfBlip(pkg, control) {
  const xml = await control.getXml();
  const relId = /r:embed="(rId\d+)"/.exec(xml)[1];
  const part = pkg.getMainDocumentPart().relationshipsPart.getPart(relId);
  return part.getBytes();
}

test('the tag\'s width parameter is read: absent keeps the drawing, auto and N scale a new one', () => {
  assert.deepEqual(pictureWidthOf(tagParamsOf('od:xpath=x1&od:Handler=picture')), { kind: 'keep' });
  assert.deepEqual(pictureWidthOf(tagParamsOf('od:xpath=x1&od:Handler=picture&width=auto')), { kind: 'auto' });
  assert.deepEqual(pictureWidthOf(tagParamsOf('od:xpath=x1&od:Handler=picture&width=4500')), { kind: 'twips', twips: 4500 });
  assert.deepEqual(pictureWidthOf(tagParamsOf('od:xpath=x1&od:Handler=picture&width=wide')), { kind: 'auto' });
});

test('(a) a picture content control with a w:dataBinding: the blip points at a new image part, the drawing keeps its extent; saved and reloaded', async () => {
  const { pkg, body, data } = await document();
  const p = body.insertParagraph('', 'End');
  const before = p.insertInlinePictureFromBase64(SMALL, 'End', { name: 'placeholder.png' });
  const control = p.insertContentControl('Picture');
  control.title = 'Logo';
  assert.equal(control.type, 'Picture');
  // setMapping applies the binding as it writes it (CR-002 section 39)
  assert.equal(control.xmlMapping.setMapping('/data/logo', '', data), true);
  const picture = control.pictureContentControl.inlinePicture;
  assert.ok(bytesEqual(await picture.imagePart().getBytes(), base64Decode(BLUE)), 'the picture shows the node\'s image');
  assert.equal(picture.width, before.width, 'the extent is the template\'s');
  assert.equal(picture.height, before.height);
  assert.equal(picture.element, before.element, 'the drawing itself is kept');
  assert.deepEqual(imageParts(pkg), ['/word/media/image1.png', '/word/media/image2.png'], 'the template image stays in the package');

  const result = await pkg.customXmlParts.applyBindings();
  assert.deepEqual(result, { bound: 1, updated: 1, skipped: 0 });
  const again = await WordprocessingMLPackage.load(await pkg.save());
  const reloaded = (await again.getBody()).contentControls[0];
  assert.equal(reloaded.type, 'Picture');
  assert.equal(reloaded.xmlMapping.xpath, '/data/logo', 'the XPath as given (CR-002 section 39)');
  assert.ok(bytesEqual(await reloaded.pictureContentControl.inlinePicture.imagePart().getBytes(), base64Decode(BLUE)));
  assert.match(await reloaded.getXml(), /<w:picture\/>/);
});

test('(a) the reverse direction leaves a picture control alone (unmeasured until Word check 38)', async () => {
  const { pkg, body, data } = await document();
  const p = body.insertParagraph('', 'End');
  p.insertInlinePictureFromBase64(SMALL, 'End');
  p.insertContentControl('Picture').xmlMapping.setMapping('/data/logo', '', data);
  const result = await pkg.customXmlParts.updateFromContentControls();
  assert.deepEqual(result, { bound: 1, updated: 0, skipped: 1 });
  assert.equal(data.selectSingleNode('/data/logo').text, BLUE);
});

test('(b) od:Handler=picture without width: a floating picture keeps its anchor, wrapping, position and extent; only r:embed changes', async () => {
  const { pkg, body } = await document();
  const main = pkg.getMainDocumentPart();
  const template = addImage(main, SMALL, body.container, undefined, { name: 'template.png' });
  const [p] = await body.insertXml(anchoredParagraphXml(template.relId, 38100, 28575, 7), 'End');
  const control = p.insertContentControl('RichText');
  control.tag = 'od:xpath=x1&od:Handler=picture';
  const before = await control.getXml();
  const result = await pkg.customXmlParts.applyBindings();
  assert.deepEqual(result, { bound: 1, updated: 1, skipped: 0 });
  const after = await control.getXml();
  assert.notEqual(after, before);
  assert.equal(after.replace(/r:embed="rId\d+"/, 'r:embed="X"'), before.replace(/r:embed="rId\d+"/, 'r:embed="X"'), 'everything but the embed is as it was');
  assert.match(after, /<wp:anchor [^>]*behindDoc="(0|false)"[^>]*>/);
  assert.match(after, /<wp:positionH relativeFrom="column"><wp:posOffset>914400<\/wp:posOffset><\/wp:positionH>/);
  assert.match(after, /<wp:wrapSquare wrapText="bothSides"\/>/);
  assert.match(after, /<wp:extent cx="38100" cy="28575"\/>/);
  assert.doesNotMatch(after, /w:dataBinding/);
  assert.ok(bytesEqual(await bytesOfBlip(pkg, control), base64Decode(BLUE)));
  const again = await WordprocessingMLPackage.load(await pkg.save());
  const reloaded = (await again.getBody()).contentControls[0];
  assert.equal(reloaded.tag, 'od:xpath=x1&od:Handler=picture');
  assert.ok(bytesEqual(await bytesOfBlip(again, reloaded), base64Decode(BLUE)));
});

test('(b) width=N: the content becomes an inline picture at its natural size, scaled down to N twips; width=auto the natural size', async () => {
  const { pkg, body } = await document();
  body.insertParagraph('scaled', 'End');
  const scaled = body.search('scaled')[0].insertContentControl('RichText');
  scaled.tag = 'od:xpath=x1&od:Handler=picture&width=20';
  body.insertParagraph('natural', 'End');
  const natural = body.search('natural')[0].insertContentControl('RichText');
  natural.tag = 'od:xpath=x1&od:Handler=picture&width=auto';
  const block = body.insertParagraph('block', 'End').insertContentControl('RichText');
  block.tag = 'od:xpath=x2&od:Handler=picture&width=4500';
  block.paragraphs[0].alignment = 'Centered';
  const result = await pkg.customXmlParts.applyBindings();
  assert.deepEqual(result, { bound: 3, updated: 3, skipped: 0 });
  // 20 twips = 12700 EMU: 57150 x 28575 scaled to 12700 x 6350
  const small = scaled.paragraphs[0].inlinePictures[0];
  assert.equal(small.width * EMU_PER_POINT, 12700);
  assert.equal(small.height * EMU_PER_POINT, 6350);
  assert.equal(scaled.text, '', 'the text the control held is gone');
  const whole = natural.paragraphs[0].inlinePictures[0];
  assert.equal(whole.width * EMU_PER_POINT, 57150);
  assert.equal(whole.height * EMU_PER_POINT, 28575);
  // 6000 twips wide into 4500: 2857500 EMU, the height following
  const fitted = block.paragraphs[0].inlinePictures[0];
  assert.equal(fitted.width * EMU_PER_POINT, 4500 * 635);
  assert.equal(fitted.height * EMU_PER_POINT, Math.round((28575 * 4500 * 635) / 3810000));
  assert.equal(block.paragraphs[0].alignment, 'Centered', 'a block-level control keeps its paragraph\'s properties');
  assert.ok(bytesEqual(await fitted.imagePart().getBytes(), base64Decode(WIDE)));
  const again = await WordprocessingMLPackage.load(await pkg.save());
  const controls = (await again.getBody()).contentControls;
  assert.equal(controls.length, 3);
  assert.equal(controls[2].paragraphs[0].inlinePictures[0].width * EMU_PER_POINT, 4500 * 635);
});

test('(b) in a table cell, width=N fits the cell\'s width less its margins when that is narrower: Word\'s 108-twip defaults, the table\'s w:tblCellMar, the cell\'s own w:tcMar', async () => {
  const { pkg, body } = await document();
  const cell = (tcPr, tblPr = '') => `<w:tbl><w:tblPr>${tblPr}<w:tblW w:w="1000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/>${tcPr}</w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
  await body.insertXml(cell(''), 'End');
  await body.insertXml(cell('', '<w:tblCellMar><w:left w:w="200" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>'), 'End');
  await body.insertXml(cell('<w:tcMar><w:left w:w="50" w:type="dxa"/><w:right w:w="50" w:type="dxa"/></w:tcMar>', '<w:tblCellMar><w:left w:w="200" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>'), 'End');
  for (const range of body.search('cell')) range.insertContentControl('RichText').tag = 'od:xpath=x2&od:Handler=picture&width=4500';
  const result = await pkg.customXmlParts.applyBindings();
  assert.deepEqual(result, { bound: 3, updated: 3, skipped: 0 });
  const widths = body.tables.map((t) => Math.round(t.getCell(0, 0).body.paragraphs[0].inlinePictures[0].width * EMU_PER_POINT));
  assert.deepEqual(widths, [(1000 - 216) * 635, (1000 - 300) * 635, (1000 - 100) * 635]);
});

test('an SVG value replaces a picture (the drawing kept, the part image/svg+xml) but cannot be placed as a new one', async () => {
  const { pkg, body, data } = await document();
  const p = body.insertParagraph('', 'End');
  p.insertInlinePictureFromBase64(SMALL, 'End');
  const control = p.insertContentControl('Picture');
  assert.equal(control.xmlMapping.setMapping('/data/svg', '', data), true);
  const picture = control.pictureContentControl.inlinePicture;
  assert.equal(picture.imageFormat, 'Svg');
  assert.equal(picture.imagePart().partName.name, '/word/media/image1.svg');
  assert.ok(bytesEqual(await picture.imagePart().getBytes(), base64Decode(SVG)));
  body.insertParagraph('inline', 'End');
  const inline = body.search('inline')[0].insertContentControl('RichText');
  inline.tag = 'od:xpath=x5&od:Handler=picture&width=auto';
  const result = await pkg.customXmlParts.applyBindings();
  assert.deepEqual(result, { bound: 2, updated: 1, skipped: 1, notes: [result.notes[0]] });
  assert.match(result.notes[0], /od:xpath=x5: not an image this package reads \(A SVG image has no pixel size/);
  assert.equal(inline.text, 'inline');
  const again = await WordprocessingMLPackage.load(await pkg.save());
  assert.equal((await again.getBody()).contentControls[0].pictureContentControl.inlinePicture.imageFormat, 'Svg');
});

test('left and noted: an empty node, a node that is not an image, a missing entry, a tag-bound control with no picture to replace', async () => {
  const { pkg, body } = await document();
  const make = (text, tag) => { body.insertParagraph(text, 'End'); const c = body.search(text)[0].insertContentControl('RichText'); c.tag = tag; return c; };
  const empty = make('empty', 'od:xpath=x3&od:Handler=picture&width=auto');
  const text = make('text', 'od:xpath=x4&od:Handler=picture&width=auto');
  const missing = make('missing', 'od:xpath=x9&od:Handler=picture');
  const noBlip = make('noblip', 'od:xpath=x1&od:Handler=picture');
  const result = await pkg.customXmlParts.applyBindings();
  assert.equal(result.bound, 4);
  assert.equal(result.updated, 0);
  assert.equal(result.skipped, 4);
  assert.match(result.notes[0], /od:xpath=x3: the node is empty/);
  assert.match(result.notes[1], /od:xpath=x4: not an image this package reads \(Unsupported image format/);
  assert.match(result.notes[2], /od:xpath=x9: no such entry/);
  assert.match(result.notes[3], /od:xpath=x1: the control holds no a:blip/);
  for (const [control, label] of [[empty, 'empty'], [text, 'text'], [missing, 'missing'], [noBlip, 'noblip']]) assert.equal(control.text, label);
  assert.deepEqual(imageParts(pkg), [], 'no image part was added');
});

test('(a) a picture control whose node is not an image is left, with a note naming the control', async () => {
  const { pkg, body, data } = await document();
  const p = body.insertParagraph('', 'End');
  p.insertInlinePictureFromBase64(SMALL, 'End');
  const control = p.insertContentControl('Picture');
  control.title = 'Logo';
  assert.equal(control.xmlMapping.setMapping('/data/text', '', data), true, 'the node exists, so the binding is written');
  const result = await pkg.customXmlParts.applyBindings();
  assert.deepEqual(result, { bound: 1, updated: 0, skipped: 1, notes: [result.notes[0]] });
  assert.match(result.notes[0], /^Logo: not an image this package reads/);
  assert.ok(bytesEqual(await control.pictureContentControl.inlinePicture.imagePart().getBytes(), base64Decode(SMALL)));
});
