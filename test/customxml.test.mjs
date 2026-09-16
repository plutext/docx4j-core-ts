// CR-002 phase E: custom XML parts, the XPath engine, XML mapping, the typed content controls
// and insertContentControl. The invoice fixture is docx4j's OpenDoPE sample: bound controls at
// block, row and cell level, a repeat and two conditions.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  WordprocessingMLPackage, ZipPartStore, ContentControl, CustomXmlPart, CustomXmlNode,
  CustomXmlDataStoragePart, CustomXmlDataStoragePropertiesPart, canonicalXPathOf, parsePrefixMappings, formatPrefixMappings,
  sdtProperty,
} from '../dist/index.mjs';
import { fixture, bytesEqual } from './helpers.mjs';

/** The customer data part of the invoice. */
const DATA_ITEM_ID = '{8B049945-9DFE-4726-9DE9-CF5691E53858}';

async function invoice() {
  return WordprocessingMLPackage.load(await fixture('invoice.docx'));
}

test('custom XML parts: items, ids, namespaces, schema refs and the document element', async () => {
  const pkg = await invoice();
  const parts = await pkg.customXmlParts.load();
  assert.equal(parts.length, 4);
  assert.ok(parts[0] instanceof CustomXmlPart);

  // getItem ignores case and braces (docx4j keys the map on the lower-cased id)
  const data = pkg.customXmlParts.getItem(DATA_ITEM_ID);
  assert.ok(data);
  assert.equal(pkg.customXmlParts.getItem(DATA_ITEM_ID.toLowerCase()), data);
  assert.equal(pkg.customXmlParts.getItem(DATA_ITEM_ID.replace(/[{}]/g, '')), data);
  assert.equal(pkg.customXmlParts.getItem('{00000000-0000-0000-0000-000000000000}'), undefined);

  assert.equal(data.namespaceUri, '');
  assert.equal(data.builtIn, false);
  assert.equal(data.schemaCollection.length, 0);
  assert.equal(data.part instanceof CustomXmlDataStoragePart, true);

  // the OpenDoPE side parts are in their own namespaces and carry ds:schemaRef entries
  const xpaths = pkg.customXmlParts.getByNamespace('http://opendope.org/xpaths');
  assert.equal(xpaths.length, 1);
  assert.ok(xpaths[0].schemaCollection.includes('http://opendope.org/xpaths'));

  const root = data.documentElement;
  assert.ok(root instanceof CustomXmlNode);
  assert.equal(root.baseName, 'invoice');
  assert.equal(root.nodeType, 'Element');
  assert.ok(root.hasChildNodes());
  assert.ok(data.getXml().startsWith('<invoice>'));
});

test('XPath: the three bindings resolve to their nodes, and a node reports Word\'s canonical xpath', async () => {
  const pkg = await invoice();
  await pkg.customXmlParts.load();
  const body = await pkg.getBody();
  const bound = body.contentControls.filter((c) => c.xmlMapping.isMapped);
  assert.equal(bound.length, 3);
  assert.deepEqual(bound.map((c) => c.xmlMapping.xpath), [
    '/invoice[1]/customer[1]/name[1]',
    '/invoice[1]/items/item[1]/name',
    '/invoice[1]/items/item[1]/price',
  ]);
  assert.deepEqual(bound.map((c) => c.xmlMapping.customXmlNode.text), ['Joe Bloggs', 'apples', '$20']);
  for (const control of bound) assert.equal(control.xmlMapping.customXmlPart.id, DATA_ITEM_ID);

  // the canonical form Word writes: every step positioned, prefixes ns0, ns1, ...
  const name = bound[1].xmlMapping.customXmlNode;
  assert.equal(name.xpath, '/invoice[1]/items[1]/item[1]/name[1]');
  assert.equal(name.prefixMappings, '');
  assert.equal(name.parentNode.baseName, 'item');
  assert.equal(name.ownerPart.id, DATA_ITEM_ID);

  // selectNodes and selectSingleNode, from the part and from a node
  const data = pkg.customXmlParts.getItem(DATA_ITEM_ID);
  assert.equal(data.selectNodes('/invoice/items/item').length, 3);
  assert.equal(data.selectSingleNode('/invoice/items/item[2]/name').text, 'bananas');
  assert.equal(data.documentElement.selectNodes('items/item/price').map((n) => n.text).join(','), '$20,$30,$40');
  assert.equal(data.selectSingleNode('/invoice/nothing'), undefined);

  // a namespaced part needs the prefixes, as Word's prefixMappings string gives them
  const xpaths = pkg.customXmlParts.getByNamespace('http://opendope.org/xpaths')[0];
  const nodes = xpaths.selectNodes("/ns15:xpaths/ns15:xpath", "xmlns:ns15='http://opendope.org/xpaths'");
  assert.equal(nodes.length, 6);
  assert.equal(nodes[0].namespaceUri, 'http://opendope.org/xpaths');
  assert.equal(canonicalXPathOf(nodes[0].node).xpath, '/ns0:xpaths[1]/ns0:xpath[1]');
  assert.equal(canonicalXPathOf(nodes[0].node).prefixMappings, "xmlns:ns0='http://opendope.org/xpaths'");
});

test('prefix mappings: Word\'s string form parses and formats', () => {
  const parsed = parsePrefixMappings(`xmlns:ns0='http://a' xmlns:ns1="http://b"`);
  assert.deepEqual(parsed, { ns0: 'http://a', ns1: 'http://b' });
  assert.equal(formatPrefixMappings(parsed), `xmlns:ns0='http://a' xmlns:ns1='http://b'`);
  assert.deepEqual(parsePrefixMappings(undefined), {});
});

test('applyBindings pushes the custom XML into the controls, updateFromContentControls writes it back', async () => {
  const pkg = await invoice();
  await pkg.customXmlParts.load();
  const data = pkg.customXmlParts.getItem(DATA_ITEM_ID);
  data.selectSingleNode('/invoice/customer/name').text = 'Alice Smith';
  data.selectSingleNode('/invoice/items/item[1]/price').text = '$21';

  const applied = await pkg.customXmlParts.applyBindings();
  assert.deepEqual(applied, { bound: 3, updated: 3, skipped: 0 });
  const body = await pkg.getBody();
  const bound = body.contentControls.filter((c) => c.xmlMapping.isMapped);
  assert.deepEqual(bound.map((c) => c.text), ['Alice Smith', 'apples', '$21']);

  // and back the other way: an edit made through the paragraph, not through the control, so that
  // nothing writes through on its own
  bound[1].paragraphs[0].insertText('pears', 'Replace');
  const updated = await pkg.customXmlParts.updateFromContentControls();
  assert.equal(updated.bound, 3);
  assert.equal(data.selectSingleNode('/invoice/items/item[1]/name').text, 'pears');

  // the round trip survives a save and a reload
  const again = await WordprocessingMLPackage.load(await pkg.save());
  await again.customXmlParts.load();
  const reloaded = (await again.getBody()).contentControls.filter((c) => c.xmlMapping.isMapped);
  assert.deepEqual(reloaded.map((c) => c.text), ['Alice Smith', 'pears', '$21']);
  assert.equal(reloaded[0].xmlMapping.customXmlNode.text, 'Alice Smith');
});

test('a bound control\'s insertText writes through to the custom XML node (the 2026-09-16 Word finding)', async () => {
  const pkg = await invoice();
  await pkg.customXmlParts.load();
  const control = (await pkg.getBody()).contentControls[0];
  control.insertText('Jane Doe', 'Replace');
  assert.equal(control.text, 'Jane Doe');
  // Word refreshes a bound control from the part on open, so the part has to hold the new value
  assert.equal(control.xmlMapping.customXmlNode.text, 'Jane Doe');
  assert.ok(pkg.customXmlParts.getItem(DATA_ITEM_ID).getXml().includes('Jane Doe'));

  const again = await WordprocessingMLPackage.load(await pkg.save());
  await again.customXmlParts.load();
  const reloaded = (await again.getBody()).contentControls[0];
  assert.equal(reloaded.xmlMapping.customXmlNode.text, 'Jane Doe');
});

test('custom XML parts that are only read stay byte for byte after a save', async () => {
  const bytes = await fixture('invoice.docx');
  const pkg = await WordprocessingMLPackage.load(bytes);
  await pkg.customXmlParts.load();                       // parses every custom XML part's DOM
  const body = await pkg.getBody();
  assert.equal(body.contentControls[0].xmlMapping.customXmlNode.text, 'Joe Bloggs');
  const saved = await pkg.save();
  const before = new ZipPartStore(bytes);
  const after = new ZipPartStore(saved);
  for (const name of ['customXml/item1.xml', 'customXml/item3.xml', 'customXml/itemProps3.xml', 'customXml/itemProps1.xml']) {
    assert.ok(bytesEqual(await before.load(name), await after.load(name)), `${name} changed`);
  }
});

test('customXmlParts.add: a new part, its properties part, an itemID and a mapping that survives a save', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const paragraph = pkg.body.insertParagraph('Hello', 'End');
  const control = paragraph.insertContentControl('PlainText');
  control.tag = 'greeting';

  const part = pkg.customXmlParts.add('<greeting xmlns="http://example.com/g"><to>World</to></greeting>');
  await pkg.customXmlParts.load();
  assert.equal(part.part.partName.name, '/customXml/item1.xml');
  assert.match(part.id, /^\{[0-9A-F-]{36}\}$/);
  assert.equal(part.namespaceUri, 'http://example.com/g');
  assert.deepEqual(part.schemaCollection, ['http://example.com/g']);
  const props = pkg.getPart('/customXml/itemProps1.xml');
  assert.ok(props instanceof CustomXmlDataStoragePropertiesPart);
  // Word drops a custom XML part the main document part does not relate to
  assert.ok(pkg.getMainDocumentPart().relationshipsPart.getRel('/customXml/item1.xml'));

  assert.equal(control.xmlMapping.isMapped, false);
  const ok = control.xmlMapping.setMapping('/ns0:greeting[1]/ns0:to[1]', "xmlns:ns0='http://example.com/g'", part);
  assert.equal(ok, true);
  assert.equal(control.xmlMapping.storeItemID, part.id);
  assert.equal(control.xmlMapping.customXmlNode.text, 'World');
  // Word returns false, and changes nothing, when the XPath selects nothing
  assert.equal(control.xmlMapping.setMapping('/ns0:greeting[1]/ns0:missing[1]', "xmlns:ns0='http://example.com/g'"), false);
  assert.equal(control.xmlMapping.xpath, '/ns0:greeting[1]/ns0:to[1]');

  await pkg.customXmlParts.applyBindings();
  assert.equal(control.text, 'World');

  const again = await WordprocessingMLPackage.load(await pkg.save());
  await again.customXmlParts.load();
  const reloaded = (await again.getBody()).contentControls[0];
  assert.equal(reloaded.xmlMapping.xpath, '/ns0:greeting[1]/ns0:to[1]');
  assert.equal(reloaded.xmlMapping.customXmlPart.id, part.id);
  assert.equal(reloaded.xmlMapping.customXmlNode.text, 'World');

  // setMappingByNode writes the node's canonical path and prefixes
  const node = reloaded.xmlMapping.customXmlPart.documentElement.childElements[0];
  assert.equal(reloaded.xmlMapping.setMappingByNode(node), true);
  assert.equal(reloaded.xmlMapping.xpath, '/ns0:greeting[1]/ns0:to[1]');
  assert.equal(reloaded.xmlMapping.prefixMappings, "xmlns:ns0='http://example.com/g'");

  reloaded.xmlMapping.delete();
  assert.equal(reloaded.xmlMapping.isMapped, false);
});

test('CustomXmlNode: reading, editing, inserting and deleting nodes and attributes', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const part = pkg.customXmlParts.add('<order id="7"><lines><line>a</line></lines></order>');
  await pkg.customXmlParts.load();

  const root = part.documentElement;
  assert.equal(root.baseName, 'order');
  assert.equal(root.attributes.length, 1);
  assert.equal(root.attributes[0].baseName, 'id');
  assert.equal(root.attributes[0].nodeType, 'Attribute');
  assert.equal(root.attributes[0].nodeValue, '7');
  assert.equal(root.attributes[0].xpath, '/order[1]/@id');
  assert.equal(root.childElements.length, 1);
  assert.equal(root.firstChild.baseName, 'lines');
  assert.equal(root.firstChild.parentNode.baseName, 'order');

  const lines = part.selectSingleNode('/order/lines');
  lines.appendChildNode('<line>b</line>');
  lines.appendChildNode('line', '', 'Element', 'c');            // Office JS's four-argument form
  assert.deepEqual(lines.childElements.map((n) => n.text), ['a', 'b', 'c']);
  assert.equal(lines.childElements[2].xpath, '/order[1]/lines[1]/line[3]');

  lines.insertNodeBefore('<line>z</line>', lines.firstChild);
  assert.deepEqual(lines.childElements.map((n) => n.text), ['z', 'a', 'b', 'c']);
  lines.removeChild(lines.firstChild);
  lines.replaceChildNode(lines.childElements[2], '<line>C</line>');
  assert.deepEqual(lines.childElements.map((n) => n.text), ['a', 'b', 'C']);
  lines.childElements[1].delete();
  assert.deepEqual(lines.childElements.map((n) => n.text), ['a', 'C']);

  part.insertElement('/order/lines', '<line>d</line>');
  part.updateElement('/order/lines/line[1]', '<line>A</line>');
  part.deleteElement('/order/lines/line[2]');
  assert.deepEqual(part.selectNodes('/order/lines/line').map((n) => n.text), ['A', 'd']);

  part.insertAttribute('/order', 'status', 'draft');
  assert.equal(part.selectSingleNode('/order').node.getAttribute('status'), 'draft');
  part.updateAttribute('/order', 'status', 'final');
  assert.equal(part.selectSingleNode('/order').node.getAttribute('status'), 'final');
  part.deleteAttribute('/order', 'status');
  assert.equal(part.selectSingleNode('/order').node.getAttribute('status'), null);

  part.setXml('<order><lines/></order>');
  assert.equal(part.getXml(), '<order><lines/></order>');
  assert.equal(part.documentElement.hasChildNodes(), true);

  // delete() takes the part, its properties part and the relationship with it
  const partName = part.part.partName.name;
  part.delete();
  assert.equal(pkg.getPart(partName), undefined);
  assert.equal(pkg.getPart('/customXml/itemProps1.xml'), undefined);
  assert.equal(pkg.customXmlDataStorageParts.size, 0);
});

test('insertContentControl on a body, a paragraph and a range', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const first = body.insertParagraph('The quick brown fox', 'End');
  body.insertParagraph('Second', 'End');

  // a range wraps the runs it covers, splitting them at the boundaries
  const [range] = first.search('brown');
  const runControl = range.insertContentControl('PlainText');
  assert.ok(runControl instanceof ContentControl);
  assert.equal(runControl.form, 'Run');
  assert.equal(runControl.type, 'PlainText');
  assert.equal(runControl.text, 'brown');
  assert.equal(first.text, 'The quick brown fox');
  assert.ok(runControl.id > 0);

  // a paragraph wraps itself
  const paragraphControl = body.paragraphs[1].insertContentControl('DatePicker');
  assert.equal(paragraphControl.form, 'Block');
  assert.equal(paragraphControl.type, 'DatePicker');
  assert.equal(paragraphControl.text, 'Second');
  assert.notEqual(paragraphControl.id, runControl.id);
  assert.equal(body.content.length, 2);

  // a body wraps everything it holds
  const bodyControl = body.insertContentControl('Group');
  assert.equal(body.content.length, 1);
  assert.equal(bodyControl.type, 'Group');
  assert.equal(bodyControl.text, 'The quick brown fox\nSecond');
  assert.ok(bodyControl.groupContentControl);
  assert.equal(body.contentControls.length, 3);

  // and all of it survives a round trip
  const again = await WordprocessingMLPackage.load(await pkg.save());
  const controls = (await again.getBody()).contentControls;
  assert.deepEqual(controls.map((c) => c.type), ['Group', 'PlainText', 'DatePicker']);
  assert.equal(controls[1].text, 'brown');
});

test('typed content controls: the properties Office JS hangs off a control', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;

  const checkboxParagraph = body.insertParagraph('', 'End');
  const checkbox = checkboxParagraph.insertContentControl('CheckBox');
  assert.equal(checkbox.type, 'CheckBox');
  const box = checkbox.checkboxContentControl;
  assert.ok(box);
  assert.equal(box.isChecked, false);
  assert.equal(box.checkedSymbol, '☒');
  assert.equal(box.uncheckedSymbol, '☐');
  box.isChecked = true;
  assert.equal(box.isChecked, true);
  assert.equal(checkbox.text, '☒');          // Word shows the glyph, in the checkbox font
  assert.equal(checkbox.datePickerContentControl, undefined);

  const listParagraph = body.insertParagraph('', 'End');
  const dropDown = listParagraph.insertContentControl('DropDownList');
  const list = dropDown.dropDownListContentControl;
  assert.ok(list);
  assert.deepEqual(list.listItems, []);
  list.addListItem('Apples', 'apples');
  list.addListItem('Pears', 'pears');
  assert.deepEqual(list.listItems.map((i) => [i.displayText, i.value]), [['Apples', 'apples'], ['Pears', 'pears']]);
  list.listItems[1].delete();
  assert.equal(list.listItems.length, 1);
  list.deleteAllListItems();
  assert.equal(list.listItems.length, 0);
  assert.equal(dropDown.comboBoxContentControl, undefined);

  const dateParagraph = body.insertParagraph('1/01/2026', 'End');
  const date = dateParagraph.insertContentControl('DatePicker').datePickerContentControl;
  assert.ok(date);
  assert.equal(date.dateFormat, 'd/MM/yyyy');
  date.dateDisplayFormat = 'yyyy-MM-dd';
  assert.equal(date.dateFormat, 'yyyy-MM-dd');
  date.dateDisplayLocale = 'en-AU';
  assert.equal(date.dateDisplayLocale, 'en-AU');
  assert.equal(date.dateStorageFormat, 'dateTime');
  date.fullDate = new Date(Date.UTC(2026, 0, 1));
  assert.equal(date.fullDate.getUTCFullYear(), 2026);

  const sectionParagraph = body.insertParagraph('Item', 'End');
  const section = sectionParagraph.insertContentControl('RepeatingSection');
  const repeating = section.repeatingSectionContentControl;
  assert.ok(repeating);
  assert.equal(repeating.allowInsertDeleteSection, true);
  repeating.allowInsertDeleteSection = false;
  assert.equal(repeating.allowInsertDeleteSection, false);
  repeating.sectionTitle = 'Lines';
  assert.equal(repeating.sectionTitle, 'Lines');
  assert.deepEqual(repeating.items, []);

  // the w:sdtPr properties Office JS reports on every control
  const control = body.paragraphs[0].insertContentControl('PlainText');
  assert.equal(control.appearance, 'BoundingBox');
  control.appearance = 'Tags';
  assert.equal(control.appearance, 'Tags');
  assert.equal(control.color, '');
  control.color = '#FF0000';
  assert.equal(control.color, '#FF0000');
  assert.equal(control.cannotDelete, false);
  assert.equal(control.cannotEdit, false);
  control.cannotDelete = true;
  assert.equal(sdtProperty(control.sdt.sdtPr, 'lock').value.val, 'sdtLocked');
  control.cannotEdit = true;
  assert.equal(sdtProperty(control.sdt.sdtPr, 'lock').value.val, 'sdtContentLocked');
  control.cannotDelete = false;
  assert.equal(sdtProperty(control.sdt.sdtPr, 'lock').value.val, 'contentLocked');
  assert.equal(control.removeWhenEdited, false);
  control.removeWhenEdited = true;
  assert.equal(control.removeWhenEdited, true);

  const empty = body.insertParagraph('', 'End').insertContentControl('PlainText');
  assert.equal(empty.placeholderText, '');
  empty.placeholderText = 'Your name';
  assert.equal(empty.isShowingPlaceholder, true);
  assert.equal(empty.placeholderText, 'Your name');
  assert.throws(() => { control.placeholderText = 'no'; }, /holds content/);

  // everything above is real w:sdtPr markup after a round trip
  const again = await WordprocessingMLPackage.load(await pkg.save());
  const reloaded = (await again.getBody()).contentControls;
  const kinds = reloaded.map((c) => c.type);
  assert.ok(kinds.includes('CheckBox'));
  assert.ok(kinds.includes('DropDownList'));
  assert.ok(kinds.includes('RepeatingSection'));
  const reloadedCheckbox = reloaded.find((c) => c.type === 'CheckBox');
  assert.equal(reloadedCheckbox.checkboxContentControl.isChecked, true);
  assert.equal(reloadedCheckbox.text, '☒');
});

test('typed content controls on a document Word wrote (docx4j\'s invoice2013)', async () => {
  const bytes = await fixture('invoice2013.docx');
  const pkg = await WordprocessingMLPackage.load(bytes);
  await pkg.customXmlParts.load();
  const controls = (await pkg.getBody()).contentControls;
  const byType = new Map();
  for (const control of controls) if (!byType.has(control.type)) byType.set(control.type, control);
  const kinds = [...byType.keys()].join(', ');
  assert.ok(byType.has('CheckBox'), kinds);
  assert.ok(byType.has('DatePicker'), kinds);
  assert.ok(byType.has('Picture'), kinds);
  assert.ok(byType.has('RepeatingSection'), kinds);

  const checkbox = byType.get('CheckBox');
  assert.equal(typeof checkbox.checkboxContentControl.isChecked, 'boolean');
  assert.equal(checkbox.checkboxContentControl.font, 'MS Gothic');
  assert.equal(checkbox.title, '/invoice[1]/VAT[1]/@applies');

  const date = byType.get('DatePicker').datePickerContentControl;
  assert.equal(date.dateFormat, 'd MMMM yyyy');
  assert.equal(date.fullDate.getUTCFullYear(), 2015);

  // a picture control reports the picture it holds
  assert.ok(byType.get('Picture').pictureContentControl);

  // a w15 repeating section: its items are the w15:repeatingSectionItem controls inside it
  const section = byType.get('RepeatingSection').repeatingSectionContentControl;
  assert.ok(section);
  assert.equal(section.allowInsertDeleteSection, true);
  assert.ok(section.items.length >= 1);
  assert.equal(section.items[0].type, 'RepeatingSectionItem');
  const added = section.insertItemAfter(section.items.length - 1);
  assert.equal(added.type, 'RepeatingSectionItem');
  assert.equal(section.items.length, 2);

  // the bindings of the document all name its one data part
  const bound = controls.filter((c) => c.xmlMapping.isMapped);
  assert.ok(bound.length >= 15);
  assert.equal(bound[0].xmlMapping.customXmlPart.id, '{5D7BA57F-1E52-4637-9F82-2D4025768D4F}');
  const data = pkg.customXmlParts.items[0];
  data.selectSingleNode('/invoice/customer/company').text = 'Acme Pty Ltd';
  const applied = await pkg.customXmlParts.applyBindings();
  assert.ok(applied.updated > 0);
  // the picture control and the repeating sections are left alone: a picture binding is deferred
  // and a container would be flattened by a value (CR-002 section 12)
  assert.ok(applied.skipped >= 3);
  assert.equal(controls.find((c) => c.title === '/invoice[1]/customer[1]/company[1]').text, 'Acme Pty Ltd');
  // the checkbox shows Word's glyph, the date is formatted with its w:dateFormat
  assert.equal(checkbox.checkboxContentControl.isChecked, true);
  assert.equal(checkbox.text, '☒');
  assert.equal(byType.get('DatePicker').text, '29 January 2015');
  // and the repeat still holds its items, with their own controls
  const stillThere = (await pkg.getBody()).contentControls.find((c) => c.type === 'RepeatingSection');
  assert.ok(stillThere.repeatingSectionContentControl.items.length >= 2);

  // the reverse writes the value, not the rendered glyph or the formatted date
  checkbox.checkboxContentControl.isChecked = false;
  const back = await pkg.customXmlParts.updateFromContentControls();
  assert.ok(back.updated >= 1);
  assert.equal(checkbox.xmlMapping.customXmlNode.text, 'false');
  assert.equal(byType.get('DatePicker').xmlMapping.customXmlNode.text, '2015-01-29T00:00:00');

  // reading a document does not rewrite the parts it did not touch
  const readOnly = await WordprocessingMLPackage.load(bytes);
  await readOnly.customXmlParts.load();
  const saved = await readOnly.save();
  const before = new ZipPartStore(bytes);
  const after = new ZipPartStore(saved);
  for (const name of ['word/styles.xml', 'word/document.xml', 'customXml/item1.xml']) {
    assert.ok(bytesEqual(await before.load(name), await after.load(name)), `${name} changed`);
  }
});
