// CR-002 phase C: tables, inline pictures, insertOoxml and content controls.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  WordprocessingMLPackage, Table, TableRow, TableCell, InlinePicture, ContentControl, ImagePart,
  imageInfoOf, naturalSizeEmu, ZipPartStore, Namespaces, base64Decode,
} from '../dist/index.mjs';
import { fixture, bytesEqual } from './helpers.mjs';

/** A 4 x 3 pixel PNG at 96 dpi (pHYs 3780 px/m), made with node:zlib; see test/README.md. */
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAAEElEQVR4nGN4YGAARww4OQAs9g8Bfw7KYwAAAABJRU5ErkJggg==';
/** A 2 x 2 GIF87a. */
const GIF = 'R0lGODdhAgACAIAAAAAAAP///ywAAAAAAgACAAACAkQBADs=';
/** A 2 x 2 24-bit BMP at 3780 px/m (96 dpi). */
const BMP = 'Qk1GAAAAAAAAADYAAAAoAAAAAgAAAAIAAAABABgAAAAAABAAAADEDgAAxA4AAAAAAAAAAAAAQEBAQEBAQEBAQEBAQEBAQA==';
/** A JPEG header: SOI, a JFIF APP0 at 300 dpi, SOF0 of 8 x 6, EOI. */
const JPEG = '/9j/4AAQSkZJRgABAQEBLAEsAAD/wAARCAAGAAgDAAAAAAAAAAAAAAD/2Q==';

const EMU_PER_POINT = 12700;

test('table: insertTable, values, rows and cells, addRows, deleteRows, header rows, round trip', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const table = body.insertTable(2, 3, 'End', [['a', 'b', 'c'], ['d', 'e', 'f']]);
  assert.ok(table instanceof Table);
  assert.equal(table.rowCount, 2);
  assert.deepEqual(table.values, [['a', 'b', 'c'], ['d', 'e', 'f']]);
  assert.equal(table.rows.length, 2);
  assert.ok(table.rows[0] instanceof TableRow);
  assert.equal(table.rows[0].cellCount, 3);
  // the grid is the section's text width (A4 less 2.54 cm margins)
  assert.deepEqual(table.columnWidths().reduce((a, b) => a + b, 0), 11907 - 1440 - 1440);

  const cell = table.getCell(1, 1);
  assert.ok(cell instanceof TableCell);
  assert.equal(cell.text, 'e');
  assert.equal(cell.rowIndex, 1);
  assert.equal(cell.cellIndex, 1);
  assert.equal(cell.parentTable.element, table.element);
  cell.value = 'E';
  assert.equal(table.values[1][1], 'E');

  table.style = 'TableGrid';
  assert.equal(table.styleId, 'TableGrid');
  assert.equal(table.style, 'Table Grid');
  assert.equal(table.styleBuiltIn, 'TableGrid');
  assert.throws(() => { table.styleBuiltIn = 'Other'; }, RangeError);
  table.headerRowCount = 1;
  assert.equal(table.headerRowCount, 1);
  assert.ok(table.rows[0].isHeader);
  assert.ok(!table.rows[1].isHeader);

  const added = table.addRows('End', 1, [['g', 'h', 'i']]);
  assert.equal(added.length, 1);
  assert.equal(added[0].rowIndex, 2);
  assert.deepEqual(table.values[2], ['g', 'h', 'i']);
  table.rows[1].insertRows('Before', 1, [['x', 'y', 'z']]);
  assert.deepEqual(table.values.map((r) => r[0]), ['a', 'x', 'd', 'g']);
  table.deleteRows(1, 1);
  assert.deepEqual(table.values.map((r) => r[0]), ['a', 'd', 'g']);

  assert.equal(body.tables.length, 1);
  assert.deepEqual(body.paragraphs.map((p) => p.text), ['a', 'b', 'c', 'd', 'E', 'f', 'g', 'h', 'i']);

  const back = await WordprocessingMLPackage.load(await pkg.save());
  const t2 = (await back.getBody()).tables[0];
  assert.deepEqual(t2.values, [['a', 'b', 'c'], ['d', 'E', 'f'], ['g', 'h', 'i']]);
  assert.equal(t2.styleId, 'TableGrid');
  assert.equal(t2.headerRowCount, 1);
  t2.delete();
  assert.equal((await back.getBody()).tables.length, 0);
});

test('table cell: a body of its own, addresses, parentTableCell', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  body.insertParagraph('intro', 'End');
  const table = body.insertTable(2, 2, 'End', [['a', 'b'], ['c', 'd']]);
  const cell = table.getCell(1, 0);
  cell.insertParagraph('second', 'End');
  assert.deepEqual(cell.paragraphs.map((p) => p.text), ['c', 'second']);
  cell.insertText('!', 'End');
  assert.equal(cell.text, 'c\nsecond!');

  const inCell = cell.paragraphs[0];
  assert.equal(inCell.parentTableCell?.cellIndex, 0);
  assert.equal(inCell.parentTableCell?.rowIndex, 1);
  assert.equal(body.paragraphs[0].parentTableCell, undefined);

  // the cell's body is addressed from the cell, so its addresses carry on from the document's
  const address = body.addressOf(cell);
  assert.equal(address, 'body/1/1/0');
  assert.equal(cell.body.prefix, address);
  assert.equal(cell.body.paragraphAt(`${address}/1`)?.text, 'second!');
  assert.equal(body.elementAt(`${address}/1`)?.element.value.content[0].value.content[0].value.value, 'second!');
});

test('image headers: PNG, JPEG, GIF and BMP sizes and resolutions', () => {
  const png = imageInfoOf(base64Decode(PNG));
  assert.deepEqual([png.contentType, png.widthPx, png.heightPx, png.dpiX, png.dpiY], ['image/png', 4, 3, 96, 96]);
  assert.deepEqual(naturalSizeEmu(png), { cx: 38100, cy: 28575 });
  const jpeg = imageInfoOf(base64Decode(JPEG));
  assert.deepEqual([jpeg.contentType, jpeg.widthPx, jpeg.heightPx, jpeg.dpiX, jpeg.dpiY], ['image/jpeg', 8, 6, 300, 300]);
  const gif = imageInfoOf(base64Decode(GIF));
  assert.deepEqual([gif.contentType, gif.widthPx, gif.heightPx, gif.dpiX], ['image/gif', 2, 2, 96]);
  const bmp = imageInfoOf(base64Decode(BMP));
  assert.deepEqual([bmp.contentType, bmp.widthPx, bmp.heightPx, bmp.dpiX], ['image/bmp', 2, 2, 96]);
  assert.throws(() => imageInfoOf(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])), /Unsupported image format/);
});

test('inline picture: the image part, the relationship, the drawing, and the saved docx', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const picture = pkg.body.insertInlinePictureFromBase64(PNG, 'End', { altTextDescription: 'a red rectangle', name: 'red.png' });
  assert.ok(picture instanceof InlinePicture);
  assert.equal(picture.width, 38100 / EMU_PER_POINT);
  assert.equal(picture.height, 28575 / EMU_PER_POINT);
  assert.equal(picture.imageFormat, 'Png');
  assert.equal(picture.altTextDescription, 'a red rectangle');

  const main = pkg.getMainDocumentPart();
  const rel = main.relationshipsPart.getRelationshipById(picture.relId);
  assert.equal(rel.type, Namespaces.IMAGE);
  assert.equal(rel.target, 'media/image1.png');
  const part = pkg.getPart('/word/media/image1.png');
  assert.ok(part instanceof ImagePart);
  assert.equal(part.contentType, 'image/png');
  assert.ok(bytesEqual(await part.getBytes(), base64Decode(PNG)));
  assert.equal(picture.imagePart(), part);

  const xml = await main.getXml();
  assert.match(xml, /<w:drawing>/);
  assert.match(xml, /<wp:inline/);
  assert.match(xml, new RegExp(`<a:blip r:embed="${picture.relId}"`));
  assert.match(xml, /<wp:extent cx="38100" cy="28575"\/>/);
  assert.match(xml, /uri="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/picture"/);

  picture.width = 72;
  assert.equal(picture.width, 72);
  assert.equal(picture.height, 54, 'the aspect ratio is kept');

  const bytes = await pkg.save();
  const names = [...new ZipPartStore(bytes).partNames()];
  assert.ok(names.includes('word/media/image1.png'));
  const contentTypes = new TextDecoder().decode(new ZipPartStore(bytes).loadSync('[Content_Types].xml'));
  assert.match(contentTypes, /Extension="png" ContentType="image\/png"|PartName="\/word\/media\/image1.png"/);

  const back = await WordprocessingMLPackage.load(bytes);
  const body = await back.getBody();
  assert.equal(body.inlinePictures.length, 1);
  const again = body.inlinePictures[0];
  assert.equal(again.imageFormat, 'Png');
  assert.equal(again.altTextDescription, 'a red rectangle');
  assert.equal(again.width, 72);
  assert.equal(await again.getBase64(), PNG);
  assert.equal(await again.getBase64ImageSrc(), PNG);
  again.delete();
  assert.equal(body.inlinePictures.length, 0);
});

test('inline picture: a second image gets its own part, and a paragraph takes one too', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const p = pkg.body.insertParagraph('caption: ', 'End');
  const first = p.insertInlinePictureFromBase64(GIF, 'End');
  const second = pkg.body.insertInlinePictureFromBase64(BMP, 'End');
  assert.equal(first.imagePart().partName.name, '/word/media/image1.gif');
  assert.equal(second.imagePart().partName.name, '/word/media/image1.bmp', 'names are per extension, as docx4j getNewPartName');
  assert.notEqual(first.relId, second.relId);
  assert.equal(p.inlinePictures.length, 1);
  assert.equal(p.text, 'caption: ');
  // wp:docPr ids are unique in the document
  const ids = pkg.body.inlinePictures.map((q) => q.inline.docPr.id);
  assert.equal(new Set(ids).size, ids.length);
  await WordprocessingMLPackage.load(await pkg.save());
});

test('insertOoxml: a pkg:package with an image is copied in with fresh names and ids', async () => {
  const source = await WordprocessingMLPackage.createPackage();
  source.body.insertParagraph('from the other package', 'End');
  source.body.insertTable(1, 2, 'End', [['x', 'y']]);
  source.body.insertInlinePictureFromBase64(PNG, 'End', { altTextDescription: 'carried over' });
  const flatOpc = await source.saveFlatOpc();

  const target = await WordprocessingMLPackage.createPackage();
  // an image of its own first, so that the copy has to take the next free name and id
  target.body.insertInlinePictureFromBase64(GIF, 'End');
  target.body.insertParagraph('before', 'End');
  const inserted = await target.body.insertOoxml(flatOpc, 'End');
  assert.equal(inserted.length, 3);

  const copied = target.getPart('/word/media/image1.png');
  assert.ok(copied instanceof ImagePart, 'the image part was copied under a free name');
  assert.ok(bytesEqual(await copied.getBytes(), base64Decode(PNG)));
  const pictures = target.body.inlinePictures;
  assert.equal(pictures.length, 2);
  const carried = pictures[1];
  assert.equal(carried.altTextDescription, 'carried over');
  assert.equal(carried.imagePart(), copied, 'r:embed was rewritten to the new relationship');
  assert.notEqual(carried.relId, pictures[0].relId);
  assert.equal(target.body.tables.length, 1);
  assert.deepEqual(target.body.tables[0].values, [['x', 'y']]);
  assert.match(target.body.text, /before\nfrom the other package/);

  const back = await WordprocessingMLPackage.load(await target.save());
  const body = await back.getBody();
  assert.equal(body.inlinePictures.length, 2);
  assert.equal(await body.inlinePictures[1].getBase64(), PNG);
});

test('insertOoxml: a bare fragment, Replace, and merging one paragraph at range level', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const inserted = await body.insertOoxml('<w:p><w:r><w:t>from a fragment</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>', 'End');
  assert.equal(inserted.length, 2);
  assert.equal(body.text, 'from a fragment\ncell');
  assert.equal(body.tables.length, 1);

  const p = body.paragraphs[0];
  await p.insertOoxml('<w:p><w:r><w:t> and more</w:t></w:r></w:p>', 'End');
  assert.equal(p.text, 'from a fragment and more');
  const [range] = p.search('fragment');
  await range.insertOoxml('<w:p><w:r><w:t>XML </w:t></w:r></w:p>', 'Before');
  assert.equal(p.text, 'from a XML fragment and more');
  await p.insertOoxml('<w:p><w:r><w:t>a block</w:t></w:r></w:p>', 'Before');
  assert.equal(body.paragraphs[0].text, 'a block');

  await body.insertOoxml('<w:p><w:r><w:t>only this</w:t></w:r></w:p>', 'Replace');
  assert.equal(body.text, 'only this');
  await WordprocessingMLPackage.load(await pkg.save());
});

test('content controls: the four forms of w:sdt in invoice.docx', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('invoice.docx'));
  const body = await pkg.getBody();
  const controls = body.contentControls;
  assert.ok(controls[0] instanceof ContentControl);
  assert.deepEqual(controls.map((c) => c.form), ['Block', 'Row', 'Cell', 'Cell', 'Block', 'Block']);
  assert.deepEqual(controls.map((c) => c.title), ['Customer name', 'Repeat', 'Description', 'Price', 'Payment instructions', 'leave out']);
  assert.deepEqual(controls.map((c) => c.type), ['PlainText', 'RichText', 'PlainText', 'PlainText', 'RichText', 'RichText']);
  assert.match(controls[0].tag, /^od:xpath=x1/);
  assert.equal(controls[0].text, 'Joe Bloggs');
  assert.equal(controls[2].text, 'apples');
  assert.equal(controls[1].tables.length, 1, 'a row control reports the table it is in');
  assert.ok(controls[0].id > 0);

  // a run-level control (Word types it CTSdtCell here) reports its paragraphs and a range
  const description = controls[2];
  assert.equal(description.paragraphs.length, 1);
  assert.equal(description.paragraphs[0].parentTableCell?.cellIndex, 0);

  // block control: text, insertText, search and a range
  const customer = controls[0];
  assert.equal(customer.getRange().text, 'Joe Bloggs');
  assert.deepEqual(customer.search('Bloggs').map((r) => r.text), ['Bloggs']);
  customer.insertText('Jane Doe', 'Replace');
  assert.equal(customer.text, 'Jane Doe');
  customer.title = 'Client name';
  customer.tag = 'client';
  assert.equal(customer.title, 'Client name');
  assert.equal(customer.tag, 'client');

  // the nested controls of the repeat row
  assert.deepEqual(controls[1].contentControls.map((c) => c.title), ['Description', 'Price']);

  // delete, with and without the content
  const before = body.paragraphs.length;
  customer.delete(true);
  assert.equal(body.contentControls.length, 5);
  assert.equal(body.paragraphs.length, before, 'the paragraph stayed');
  assert.match(body.text, /Jane Doe/);
  const payment = body.contentControls.find((c) => c.title === 'Payment instructions');
  payment.delete(false);
  assert.equal(body.contentControls.length, 4);
  assert.ok(!body.text.includes('Please remit'));

  await WordprocessingMLPackage.load(await pkg.save());
});

test('content controls: untouched parts still round trip byte for byte', async () => {
  const bytes = await fixture('invoice.docx');
  const pkg = await WordprocessingMLPackage.load(bytes);
  const body = await pkg.getBody();
  assert.equal(body.contentControls.length, 6);
  const saved = await pkg.save();
  const before = new ZipPartStore(bytes);
  const after = new ZipPartStore(saved);
  for (const name of ['word/styles.xml', 'word/theme/theme1.xml', 'customXml/item1.xml', 'word/comments.xml']) {
    assert.ok(bytesEqual(before.loadSync(name), after.loadSync(name)), `${name} is byte identical`);
  }
});
