// The files for Word check 38 (CR-005 section 9; docx4j-ts-editor ED-005 proposal 45): what Word
// 2010 and Word 15 do with a picture content control mapped to a base64 node, and what they show for
// od:Handler=picture controls bound by the engine. Run from the repository root after `npm run build`:
//
//   node test/fixtures/check38/build.mjs
//
// writes into this directory:
//   38a-picture-template.docx   Word's picture content control (w:sdtPr/w:picture) with a w:dataBinding to
//                               /data/logo, a blue 300 x 150 image; the control shows a red 200 x 100 placeholder
//   38a-picture-bound.docx      the same after the engine's bind: the control shows the blue image, the drawing
//                               (its extent, the placeholder's) kept
//   38b-handler-template.docx   five rich text controls tagged od:xpath=..&od:Handler=picture, no w:dataBinding:
//                               a floating red placeholder (no width), two text placeholders (width=4500 and
//                               width=auto), one in a 3000-twip table cell (width=4500), and an inline red
//                               placeholder bound to an SVG node (no width)
//   38b-handler-bound.docx      the same after applyBindings(): the floating picture shows the blue image at the
//                               same place and size; 2 to 4 hold a new inline picture of the wide green image
//                               (800 x 200, 12000 twips) scaled to 4500 twips, the text width, and the cell's width
//                               less its margins; 5's blip points at an image/svg+xml part
// Open each in Word 2010 and Word 15, look, save beside them as README.md says.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage, addImage, find } from '../../../dist/index.mjs';
import { pngOf } from '../../helpers.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
/** The templates' placeholder: 200 x 100 px at 96 dpi, 2.08 x 1.04 in. */
const RED = pngOf(200, 100, [220, 40, 40]);
/** /data/logo: 300 x 150 px, 3.125 x 1.56 in (4500 x 2250 twips). */
const BLUE = pngOf(300, 150, [40, 70, 220]);
/** /data/wide: 800 x 200 px, 8.33 x 2.08 in (12000 x 3000 twips), wider than the page's text. */
const WIDE = pngOf(800, 200, [40, 160, 60]);
/** /data/svg: an orange circle on a grey square, as SVG; Word 2016 and later read SVG parts, Word 2010 and 15 do not. */
const SVG = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="0 0 200 100"><rect width="200" height="100" fill="#dddddd"/><circle cx="100" cy="50" r="40" fill="#f08020"/></svg>').toString('base64');

/** A paragraph holding a floating picture (wp:anchor), as Word writes one, over an image relationship. */
function anchoredParagraphXml(relId, cx, cy, id) {
  return `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`
    + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>2743200</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${id}" name="Floating picture" descr="the floating placeholder"/>`
    + `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`
    + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Floating picture"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`
    + `</wp:anchor></w:drawing></w:r></w:p>`;
}

/** 38a: Word's picture content control, mapped. `setMapping` applies the binding as it writes it, so the template puts the placeholder back afterwards. */
async function pictureControl({ bound }) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  body.insertParagraph(`Check 38a${bound ? ', bound by the engine' : ''}: a picture content control (w:sdtPr/w:picture) with a w:dataBinding to a node holding a blue 300 x 150 image. ${bound ? 'The control shows the blue image, in the red placeholder\'s drawing.' : 'The control shows a red 200 x 100 placeholder; the node holds blue.'}`, 'End');
  const p = body.insertParagraph('', 'End');
  const placeholder = p.insertInlinePictureFromBase64(RED, 'End', { name: 'placeholder.png', altTextDescription: 'the red placeholder' });
  const placeholderRelId = placeholder.relId;
  const control = p.insertContentControl('Picture');
  control.title = 'Logo';
  control.tag = 'logo';
  const data = pkg.customXmlParts.add(`<data><logo>${BLUE}</logo></data>`);
  await pkg.customXmlParts.load();
  control.xmlMapping.setMapping('/data/logo', '', data);
  if (!bound) for (const blip of find(control.sdt.sdtContent, 'org_docx4j_dml.CTBlip')) blip.embed = placeholderRelId;
  body.insertParagraph('After the control.', 'End');
  return pkg;
}

/** 38b: od:Handler=picture controls, bound through the XPaths part's entries. */
async function handlerControls() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const main = pkg.getMainDocumentPart();
  const data = pkg.customXmlParts.add(`<data><logo>${BLUE}</logo><wide>${WIDE}</wide><svg>${SVG}</svg></data>`);
  const entry = (id, xpath) => `<xpath id="${id}"><dataBinding storeItemID="${data.id}" xpath="${xpath}" prefixMappings=""/></xpath>`;
  pkg.customXmlParts.add(`<xpaths xmlns="http://opendope.org/xpaths">${entry('x1', '/data/logo')}${entry('x2', '/data/wide')}${entry('x3', '/data/svg')}</xpaths>`);
  await pkg.customXmlParts.load();
  body.insertParagraph('Check 38b: od:Handler=picture controls, bound through od:xpath entries (no w:dataBinding on any of them).', 'End');
  body.insertParagraph('1. A floating red placeholder (200 x 100) in a rich text control, no width parameter. After the bind, the same floating picture shows the blue 300 x 150 image, at the same place and size:', 'End');
  const template = addImage(main, RED, body.container, undefined, { name: 'placeholder.png' });
  const [anchored] = await body.insertXml(anchoredParagraphXml(template.relId, 1905000, 952500, 50), 'End');
  anchored.insertText('Text in the paragraph that holds the floating picture, wrapping around it on both sides; enough of it to show the wrap, so this sentence goes on a little longer than it needs to.', 'End');
  const floating = anchored.insertContentControl('RichText');
  floating.tag = 'od:xpath=x1&od:Handler=picture';
  floating.title = 'Floating picture';
  body.insertParagraph('2. width=4500: a wide green image (800 x 200, 12000 twips) scaled down to 4500 twips, 3.125 inches:', 'End');
  const fixed = body.insertParagraph('picture placeholder (width=4500)', 'End').insertContentControl('RichText');
  fixed.tag = 'od:xpath=x2&od:Handler=picture&width=4500';
  fixed.title = 'width=4500';
  body.insertParagraph('3. width=auto: the same image scaled down to the text width:', 'End');
  const auto = body.insertParagraph('picture placeholder (width=auto)', 'End').insertContentControl('RichText');
  auto.tag = 'od:xpath=x2&od:Handler=picture&width=auto';
  auto.title = 'width=auto';
  body.insertParagraph('4. In a table cell 3000 twips (2.08 inches) wide, width=4500: the image fits the cell less its margins (108 twips each side: 2784 twips, 1.93 inches):', 'End');
  await body.insertXml(`<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>cell placeholder</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`, 'End');
  const cell = body.search('cell placeholder')[0].insertContentControl('RichText');
  cell.tag = 'od:xpath=x2&od:Handler=picture&width=4500';
  cell.title = 'in a cell';
  body.insertParagraph('5. An inline red placeholder bound to an SVG node, no width: after the bind the same drawing shows an orange circle on grey, if this Word reads an SVG part pointed at directly by a:blip (no PNG fallback is written):', 'End');
  const svgHolder = body.insertParagraph('', 'End');
  svgHolder.insertInlinePictureFromBase64(RED, 'End', { name: 'svg-placeholder.png' });
  const svg = svgHolder.insertContentControl('RichText');
  svg.tag = 'od:xpath=x3&od:Handler=picture';
  svg.title = 'SVG, no width';
  body.insertParagraph('After the controls.', 'End');
  return pkg;
}

await writeFile(join(here, '38a-picture-template.docx'), await (await pictureControl({ bound: false })).save());
await writeFile(join(here, '38a-picture-bound.docx'), await (await pictureControl({ bound: true })).save());
await writeFile(join(here, '38b-handler-template.docx'), await (await handlerControls()).save());
const bound = await handlerControls();
const result = await bound.customXmlParts.applyBindings();
console.log('38b bound:', JSON.stringify(result));
await writeFile(join(here, '38b-handler-bound.docx'), await bound.save());
console.log('written the four files in', here);
