// The files for Word check 47 (CR-002 section 52): which form an SVG picture's binding should write.
// Run from the repository root after `npm run build`:
//
//   node test/fixtures/check47/build.mjs
//
// writes into this directory (test/README.md check 47 says what to do with each):
//   47-bindings.docx       two picture content controls with a w:dataBinding, left for Word to bind:
//                          47a, its node an SVG (a blue circle); 47c, its node a red PNG, the control's
//                          picture at 30% opacity (a:alphaModFix). Each control's picture is a green square.
//   47-svgblip-only.docx   loadAndSave.docx's SVG icon as Word writes it (the baseline), then the same
//                          icon with an a:blip holding only the asvg:svgBlip, no r:embed (47b)
import { readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage, addImagePart, base64Encode, deepCopy, linkParents } from '../../../dist/index.mjs';
import { pngOf } from '../../helpers.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const GREEN = pngOf(96, 96, [0, 160, 0]);
const RED = pngOf(96, 96, [220, 0, 0]);
const BLUE_SVG = Buffer.from('<svg viewBox="0 0 96 96" xmlns="http://www.w3.org/2000/svg"><circle cx="48" cy="48" r="40" fill="#0050DC"/></svg>').toString('base64');

/** A picture content control mapped to `xpath`, its picture a green square left as it is (not bound here). */
function boundControl(body, data, label, xpath) {
  body.insertParagraph(label, 'End');
  const p = body.insertParagraph('', 'End');
  const picture = p.insertInlinePictureFromBase64(GREEN, 'End', { width: 72 });
  const blip = picture.inline.graphic.graphicData.any[0].value.blipFill.blip;
  const green = blip.embed;
  const control = p.insertContentControl('Picture');
  control.xmlMapping.setMapping(xpath, '', data);
  blip.embed = green;                // setMapping binds at once; the picture goes back to the green square
  return blip;
}

// --- 47-bindings.docx
{
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const data = pkg.customXmlParts.add(`<data><svg>${BLUE_SVG}</svg><red>${RED}</red></data>`);
  await pkg.customXmlParts.load();
  boundControl(body, data, '47a. A picture control bound to an SVG (a blue circle); its picture here a green square:', '/data/svg');
  const blip = boundControl(body, data, '47c. A picture control bound to a red PNG; its picture here a green square at 30% opacity (a:alphaModFix):', '/data/red');
  blip.alphaBiLevelOrAlphaCeilingOrAlphaFloor = [{ TYPE_NAME: 'org_docx4j_dml.CTAlphaModulateFixedEffect', amt: 30000 }];
  body.insertParagraph('End of 47-bindings.', 'End');
  await writeFile(join(here, '47-bindings.docx'), await pkg.save());
}

// --- 47-svgblip-only.docx
{
  const src = await WordprocessingMLPackage.load(new Uint8Array(await readFile(join(here, '..', 'loadAndSave.docx'))));
  const [icon] = (await src.getBody()).inlinePictures;
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const main = pkg.getMainDocumentPart();
  const png = addImagePart(main, base64Encode(await icon.imagePart().getBytes()));
  const svg = addImagePart(main, base64Encode(await icon.svgImagePart().getBytes()));
  let id = 1;
  const iconParagraph = (label, embed) => {
    body.insertParagraph(label, 'End');
    const p = body.insertParagraph('', 'End');
    const drawing = deepCopy(icon.element);
    const inline = drawing.value.anchorOrInline[0];
    inline.docPr.id = id;
    const pic = inline.graphic.graphicData.any[0].value;
    pic.nvPicPr.cNvPr.id = id++;
    const blip = pic.blipFill.blip;
    if (embed === undefined) delete blip.embed; else blip.embed = embed;
    blip.extLst.ext[0].any.value.embed = svg.relId;
    const run = { name: { namespaceURI: W, localPart: 'r' }, value: { TYPE_NAME: 'org_docx4j_wml.R', content: [drawing] } };
    (p.p.content ??= []).push(run);
    linkParents(run, p.p);
  };
  iconParagraph('Baseline. The icon as Word writes it: a:blip on its PNG, asvg:svgBlip on its SVG:', png.relId);
  iconParagraph('47b. The icon with an a:blip holding only the asvg:svgBlip, no r:embed:', undefined);
  body.insertParagraph('End of 47-svgblip-only.', 'End');
  await writeFile(join(here, '47-svgblip-only.docx'), await pkg.save());
}
