// The file for Word check 46 (CR-002 section 52): what Word draws for an SVG picture after this
// package's picture binding has swapped its image. loadAndSave.docx's SVG icon three times, two of them
// bound by this package to a node holding a solid red PNG (the binding's keep mode, CR-005 section 9).
// Run from the repository root after `npm run build`:
//
//   node test/fixtures/check46/build.mjs
//
// writes 46-svg-binding.docx into this directory (test/README.md check 46 says what to do with it).
import { readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage, addImagePart, base64Encode, deepCopy, linkParents } from '../../../dist/index.mjs';
import { pngOf } from '../../helpers.mjs';

const here = dirname(fileURLToPath(import.meta.url));

const RED = pngOf(96, 96, [220, 0, 0]);
const src = await WordprocessingMLPackage.load(new Uint8Array(await readFile(join(here, '..', 'loadAndSave.docx'))));
const [icon] = (await src.getBody()).inlinePictures;
const pkg = await WordprocessingMLPackage.createPackage();
const body = await pkg.getBody();
const main = pkg.getMainDocumentPart();
const png = addImagePart(main, base64Encode(await icon.imagePart().getBytes()));
const svg = addImagePart(main, base64Encode(await icon.svgImagePart().getBytes()));
const data = pkg.customXmlParts.add(`<data><red>${RED}</red></data>`);
pkg.customXmlParts.add(`<xpaths xmlns="http://opendope.org/xpaths"><xpath id="x1"><dataBinding storeItemID="${data.id}" xpath="/data/red" prefixMappings=""/></xpath></xpaths>`);
await pkg.customXmlParts.load();

let id = 1;
/** A paragraph holding the label, then a copy of the icon's drawing on this package's two image parts. */
function iconParagraph(label) {
  body.insertParagraph(label, 'End');
  const p = body.insertParagraph('', 'End');
  const drawing = deepCopy(icon.element);
  const inline = drawing.value.anchorOrInline[0];
  inline.docPr.id = id;
  const pic = inline.graphic.graphicData.any[0].value;
  pic.nvPicPr.cNvPr.id = id++;
  const blip = pic.blipFill.blip;
  blip.embed = png.relId;
  blip.extLst.ext[0].any.value.embed = svg.relId;
  const run = { name: { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localPart: 'r' }, value: { TYPE_NAME: 'org_docx4j_wml.R', content: [drawing] } };
  (p.p.content ??= []).push(run);
  linkParents(run, p.p);
  return p;
}

iconParagraph('46a. The icon, not bound (Word\'s SVG picture: a:blip on its PNG, asvg:svgBlip on its SVG):');
const b = iconParagraph('46b. od:Handler=picture, bound by this package to a red PNG (no w:dataBinding):');
const handler = b.insertContentControl('RichText');
handler.tag = 'od:xpath=x1&od:Handler=picture';
const c = iconParagraph('46c. A picture content control with w:dataBinding, bound by this package to the same red PNG:');
const control = c.insertContentControl('Picture');
console.log('setMapping', control.xmlMapping.setMapping('/data/red', '', data));
console.log('applyBindings', JSON.stringify(await pkg.customXmlParts.applyBindings()));
body.insertParagraph('End of check 46.', 'End');
const out = await pkg.save();
await writeFile(join(here, '46-svg-binding.docx'), out);

const back = await WordprocessingMLPackage.load(out);
const bb = await back.getBody();
for (const p of bb.inlinePictures) console.log(p.relId, p.svgRelId, p.imagePart().contentType, p.imageFormat);
const xml = await back.getMainDocumentPart().getXml();
for (const m of xml.matchAll(/<a:blip .*?<\/a:blip>|<a:blip [^>]*\/>/g)) console.log(m[0].replace(/xmlns:asvg="[^"]*" /, ''));
console.log((xml.match(/w:dataBinding/g) ?? []).length, 'w:dataBinding');
