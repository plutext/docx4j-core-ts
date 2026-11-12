// How the ctpp-* inputs were built (CR-005 section 10.2), by this package at 0.8.0 with the section 10.2 export, run
// against dist/. Kept as the record: the committed .docx are the bytes the bridge ran, so never write over them. A
// rebuild differs from them only in the control's w:id and the store's ID, which are random.
//   node test/picture-fit/build.mjs <an empty directory>
// Each is a block-level picture control mapped to /data/pic, its node then given a picture without binding, so Word
// fills the control when it opens the file:
//   ctpp-00: this package's form (flag) over a 2:1 picture, the editor's pdu6 rebuilt
//   ctpp-01: this package's form (flag) over a 1:2 picture
//   ctpp-02: an unflagged 1270000-square drawing of a red square picture over a 2:1 picture
//   ctpp-03: the same over a 1:2 picture
//   ctpp-04: an unflagged 1270000 x 635000 drawing of a red 2:1 picture over a 1:2 picture (the confirming case)
import { writeFileSync, existsSync, readdirSync } from 'node:fs';
import { WordprocessingMLPackage } from '../../dist/index.mjs';
import { pngOf } from '../helpers.mjs';

const out = process.argv[2];
if (!out || !existsSync(out) || readdirSync(out).length > 0) {
  console.error('usage: node test/picture-fit/build.mjs <an empty directory>');
  process.exit(1);
}
const WIDE = pngOf(120, 60, [0, 0, 255]);    // blue, 2:1
const TALL = pngOf(60, 120, [0, 160, 0]);    // green, 1:2
const SQUARE = pngOf(64, 64, [255, 0, 0]);   // red, ctpp-02's and ctpp-03's own picture
const OWN_WIDE = pngOf(128, 64, [255, 0, 0]); // red, 2:1, ctpp-04's own picture

async function build(name, label, nodePicture, own) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const data = pkg.customXmlParts.add('<data><pic></pic></data>');
  await pkg.customXmlParts.load();
  body.insertParagraph(`BEGIN ${name}: ${label}`, 'End');
  const node = () => data.selectNodes('/data/pic')[0];
  let control;
  if (!own) {
    // this package's form: mapped over the empty node, the flag and docx4j's placeholder picture at 1270000 EMU square
    control = body.insertParagraph('', 'End').insertContentControl('Picture');
    control.xmlMapping.setMapping('/data/pic', '', data);
  } else {
    // a picture of its own, mapped while the node holds the same bytes, so no flag
    const p = body.insertParagraph('', 'End');
    p.insertInlinePictureFromBase64(own.picture, 'End', { width: own.width, height: own.height });
    control = p.insertContentControl('Picture');
    node().text = own.picture;
    control.xmlMapping.setMapping('/data/pic', '', data);
  }
  control.title = name;
  // the node's picture written without binding, so the file holds the drawing as built and Word fills it on open
  node().text = nodePicture;
  body.insertParagraph(`END ${name}`, 'End');
  writeFileSync(`${out}/${name}.docx`, await pkg.save());
}
await build('ctpp-00-form-wide', 'our placeholder form over a 2:1 picture (control: the editor pdu6)', WIDE);
await build('ctpp-01-form-tall', 'our placeholder form over a 1:2 picture', TALL);
await build('ctpp-02-plain-wide', 'an unflagged square picture over a 2:1 picture', WIDE, { picture: SQUARE, width: 100, height: 100 });
await build('ctpp-03-plain-tall', 'an unflagged square picture over a 1:2 picture', TALL, { picture: SQUARE, width: 100, height: 100 });
await build('ctpp-04-wide-box-tall', 'an unflagged 2:1 picture (100 x 50 pt) over a 1:2 picture', TALL, { picture: OWN_WIDE, width: 100, height: 50 });
