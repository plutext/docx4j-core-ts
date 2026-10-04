// CR-002 phase E: custom XML parts, the XPath engine, XML mapping, the typed content controls
// and insertContentControl. The invoice fixture is docx4j's OpenDoPE sample: bound controls at
// block, row and cell level, a repeat and two conditions.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  WordprocessingMLPackage, ZipPartStore, ContentControl, CustomXmlPart, CustomXmlNode,
  CustomXmlDataStoragePart, CustomXmlDataStoragePropertiesPart, canonicalXPathOf, parsePrefixMappings, formatPrefixMappings,
  sdtProperty, DefaultXPathEngine, parseXml,
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

// CR-002 section 21: readiness is a property of the document, not of the global.
//
// A mixed environment — a global `document` with `evaluate` (jsdom) while the XML is parsed by
// another DOM (xmldom, which the runtime uses in Node) — used to make DefaultXPathEngine decide
// "native" once from the global and then never load the `xpath` package, so a select over the
// parsed tree had no evaluator. Found by the docx4j-ts-editor session under vitest with jsdom
// (ED-003 section 11.8) once jsonix 3.3.0's `node` export condition took effect.
test('the XPath engine readies for the document it is given, not the global one', async () => {
  const engine = new DefaultXPathEngine();
  const document_ = globalThis.document;
  // a global that claims to evaluate, as jsdom's does
  globalThis.document = { evaluate: () => { throw new Error('the host DOM cannot read this tree'); } };
  try {
    const fresh = new DefaultXPathEngine();
    assert.equal(fresh.isReady, true, 'the global makes it look ready');

    // ready() with a node of the actually-parsed document loads the engine's own XPath instead
    const parsed = parseXml('<a><b>one</b><b>two</b></a>');
    await fresh.ready(parsed);
    assert.deepEqual(fresh.select('/a/b', parsed).map((n) => n.textContent), ['one', 'two'],
      'and then the parsed tree is readable');

    // without that, the failure says why rather than answering nonsense
    const blind = new DefaultXPathEngine();
    await blind.ready();
    assert.throws(() => blind.select('/a/b', parsed), /parsed by another one|not ready/);
  } finally {
    if (document_ === undefined) delete globalThis.document;
    else globalThis.document = document_;
  }
  // unchanged where the parser and the evaluator are the same DOM: plain Node
  await engine.ready(parseXml('<a><b>one</b></a>'));
  assert.equal(engine.isReady, true);
});

// CR-002 section 39: what mapping does to the control's text (test/README.md check 34, 2026-10-03)
test('setMapping puts the node\'s value into the control, as Word and Office JS do (check 34)', async () => {
  const make = async () => {
    const pkg = await WordprocessingMLPackage.createPackage();
    await pkg.customXmlParts.load();
    pkg.customXmlParts.add('<data><name>Ann</name><empty/></data>');
    const p = pkg.body.insertParagraph('The quick brown fox.', 'End');
    p.search('brown')[0].font.bold = true;
    return pkg;
  };
  const xmlOf = async (pkg) => (await pkg.getMainDocumentPart().getXml()).replace(/ w14:\w+="[^"]*"/g, '');
  // 2a: a control over "quick brown" (its second run bold), then the mapping: "Ann", the part unchanged
  const a = await make();
  const over = a.body.paragraphs[0].search('quick brown')[0].insertContentControl('PlainText');
  assert.equal(over.xmlMapping.setMapping('/data/name'), true);
  assert.equal(over.text, 'Ann');
  assert.equal(a.body.paragraphs[0].text, 'The Ann fox.');
  assert.match(await xmlOf(a), /<w:sdtPr><w:id w:val="\d+"\/><w:text\/><w:dataBinding w:storeItemID="\{[0-9A-F-]+\}" w:xpath="\/data\/name"\/><\/w:sdtPr><w:sdtContent><w:r><w:t>Ann<\/w:t><\/w:r><\/w:sdtContent>/);
  assert.equal(a.customXmlParts.items[0].getXml().includes('<name>Ann</name>'), true, 'the node keeps its value');
  // 2b: at a caret, and a control already holding the value
  const b = await make();
  const caret = b.body.paragraphs[0].search('quick')[0].getRange('End').insertContentControl('PlainText');
  caret.xmlMapping.setMapping('/data/name');
  assert.equal(caret.text, 'Ann');
  assert.equal(b.body.paragraphs[0].text, 'The quickAnn brown fox.');
  const c = await make();
  const ann = c.body.paragraphs[0].search('quick brown')[0].insertContentControl('PlainText');
  ann.insertText('Ann', 'Replace');
  ann.xmlMapping.setMapping('/data/name');
  assert.equal(ann.text, 'Ann');
  // 1i: mapped to an empty node, the control shows its placeholder
  const d = await make();
  const empty = d.body.paragraphs[0].getRange('End').insertContentControl('PlainText');
  assert.equal(empty.xmlMapping.setMapping('/data/empty'), true);
  assert.equal(empty.isShowingPlaceholder, true);
  // a path that selects nothing: false, and the control untouched
  const e = await make();
  const kept = e.body.paragraphs[0].search('quick brown')[0].insertContentControl('PlainText');
  assert.equal(kept.xmlMapping.setMapping('/data/missing'), false);
  assert.equal(kept.text, 'quick brown');
  // setMappingByNode the same
  const f = await make();
  const byNode = f.body.paragraphs[0].search('quick brown')[0].insertContentControl('PlainText');
  const node = f.customXmlParts.items[0].selectSingleNode('/data/name');
  assert.equal(byNode.xmlMapping.setMappingByNode(node), true);
  assert.equal(byNode.text, 'Ann');
});

// CR-002 section 40, item 1 (test/README.md check 35, 2026-10-03): Word maps a rich text control with
// w15:dataBinding; w:dataBinding on one makes it a plain text control to Word 15 and Word 2010.
test('setMapping writes w15:dataBinding on a rich text control and w:dataBinding on the typed kinds (check 35)', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  await pkg.customXmlParts.load();
  const part = pkg.customXmlParts.add('<data><name>Ann</name><other>Bea</other></data>');
  const p = pkg.body.insertParagraph('The quick brown fox.', 'End');
  const rich = p.search('quick')[0].insertContentControl('RichText');
  const plain = p.search('fox')[0].insertContentControl('PlainText');
  assert.equal(rich.xmlMapping.setMapping('/data/name'), true);
  assert.equal(plain.xmlMapping.setMapping('/data/other'), true);
  // the node's text goes into the rich text control, as into a plain text one (1a)
  assert.equal(rich.text, 'Ann');
  assert.equal(rich.xmlMapping.isMapped, true);
  assert.equal(rich.xmlMapping.xpath, '/data/name');
  assert.equal(rich.xmlMapping.usesW15, true);
  assert.equal(plain.xmlMapping.usesW15, false);
  // mapping again replaces the binding, never leaving two
  assert.equal(rich.xmlMapping.setMappingByNode(part.selectSingleNode('/data/other')), true);
  assert.equal(rich.text, 'Bea');
  const xml = await pkg.getMainDocumentPart().getXml();
  assert.equal(xml.match(/<w15:dataBinding /g).length, 1);
  assert.match(xml, /<w15:dataBinding w:storeItemID="\{[0-9A-F-]+\}" w:xpath="\/data\[1\]\/other\[1\]"\/>/);
  assert.equal(xml.match(/<w:dataBinding /g).length, 1, 'the plain text control alone');
  // item 2: the w15 element makes the prefix ignorable, declared on the root
  assert.match(xml, /<w:document [^>]*xmlns:w15="http:\/\/schemas.microsoft.com\/office\/word\/2012\/wordml"[^>]*mc:Ignorable="w15"/);

  // a rich text control bound with w:dataBinding (0.3.1 and earlier, 1f) is rebound in w15, in place
  const reloaded = await WordprocessingMLPackage.load(await pkg.save());
  await reloaded.customXmlParts.load();
  const again = (await reloaded.getBody()).contentControls[0];
  assert.equal(again.xmlMapping.xpath, '/data[1]/other[1]', 'read back from w15:dataBinding');
  again.xmlMapping.delete();
  assert.equal(again.xmlMapping.isMapped, false);
  again.putProperty({ name: { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localPart: 'dataBinding' },
    value: { TYPE_NAME: 'org_docx4j_wml.CTDataBinding', xpath: '/data/name', storeItemID: again.parentBody.package_.customXmlParts.items[0].id } });
  assert.equal(again.xmlMapping.isMapped, true);
  assert.equal(again.xmlMapping.setMapping('/data/name'), true);
  const rebound = await reloaded.getMainDocumentPart().getXml();
  assert.equal(rebound.match(/<w15:dataBinding /g).length, 1);
  assert.equal(rebound.match(/<w:dataBinding /g).length, 1, 'only the plain text control keeps w:dataBinding');
  again.xmlMapping.delete();
  assert.equal(again.xmlMapping.isMapped, false, 'delete removes the w15 binding');
});

// CR-002 section 40, item 1: a node holding a rich text control's content as Flat OPC is never put
// into a control as text, nor overwritten with a control's text. 35a-word15.docx is Word's save of
// check 35: `name` holds a whole pkg:package, 1a is mapped to it with w15:dataBinding, and the inert
// plain text control Word nested in 1a (1e) is bound to the same node with w:dataBinding - the one
// Word 2010 fills with the package as raw XML.
test('a node holding Flat OPC is not put into a control as text (check 35)', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('revisions/check35/35a-word15.docx'));
  await pkg.customXmlParts.load();
  const body = await pkg.getBody();
  const all = [];
  const collect = (controls) => { for (const c of controls) { all.push(c); collect(c.contentControls); } };
  collect(body.contentControls);
  const onName = all.filter((c) => c.xmlMapping.isMapped && /name/.test(c.xmlMapping.xpath));
  assert.ok(onName.length >= 2, '1a and the control nested in it');
  assert.ok(onName.some((c) => c.type === 'PlainText'), 'the nested control is plain text, which applyBindings would fill');
  const value = onName[0].xmlMapping.customXmlNode.text;
  assert.match(value, /<pkg:package /);
  const before = onName.map((c) => c.text);
  await pkg.customXmlParts.applyBindings();
  assert.deepEqual(onName.map((c) => c.text), before, 'no control took the package as text');
  for (const c of all) assert.doesNotMatch(c.text, /pkg:package/);
  // nor is the package overwritten with a control's text
  await pkg.customXmlParts.updateFromContentControls();
  assert.equal(onName[0].xmlMapping.customXmlNode.text, value);
});

// CR-002 section 40, item 2: Word 2010 opens a part holding a w15 element only when w15 is ignorable.
test('a w15 element the engine writes lists w15 in mc:Ignorable (check 35)', async () => {
  const rootOf = async (pkg) => /<w:document [^>]*>/.exec(await pkg.getMainDocumentPart().getXml())[0];
  const section = await WordprocessingMLPackage.createPackage();
  assert.doesNotMatch(await rootOf(section), /Ignorable/);
  section.body.insertParagraph('Row', 'End').insertContentControl('RepeatingSection');
  assert.match(await rootOf(section), /xmlns:w15="[^"]*"[^>]*mc:Ignorable="w15"/);
  for (const set of [(c) => { c.appearance = 'Tags'; }, (c) => { c.color = '#FF0000'; }]) {
    const pkg = await WordprocessingMLPackage.createPackage();
    const control = pkg.body.insertParagraph('Text', 'End').insertContentControl('PlainText');
    assert.doesNotMatch(await rootOf(pkg), /Ignorable/, 'a control with no w15 element declares nothing');
    set(control);
    assert.match(await rootOf(pkg), /mc:Ignorable="w15"/);
  }
  // beside the tracker's w16du, and once
  const tracked = await WordprocessingMLPackage.createPackage();
  tracked.body.insertParagraph('Kept.', 'End');
  await tracked.setChangeTrackingMode('TrackAll');
  const control = tracked.body.paragraphs[0].insertContentControl('PlainText');
  control.appearance = 'Tags';
  control.color = '#00FF00';
  const ignorable = /mc:Ignorable="([^"]*)"/.exec(await rootOf(tracked))[1].split(' ');
  assert.deepEqual([...ignorable].sort(), ['w15', 'w16du']);
});

// CR-002 section 40, item 3: the PlaceholderText definition, from current Word (35a-word15.docx).
test('styles.ensure splices PlaceholderText, and a placeholder the engine writes brings it at save (check 35)', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  assert.deepEqual(await pkg.styles.ensure('PlaceholderText'), ['PlaceholderText']);
  assert.deepEqual(await pkg.styles.ensure('PlaceholderText'), []);
  assert.match(await pkg.getMainDocumentPart().styleDefinitionsPart.getXml(),
    /<w:style w:styleId="PlaceholderText" w:type="character"><w:name w:val="Placeholder Text"\/><w:basedOn w:val="DefaultParagraphFont"\/><w:uiPriority w:val="99"\/><w:unhideWhenUsed\/><w:rPr><w:color w:val="666666"\/><\/w:rPr><\/w:style>/);

  const stylesOf = async (bytes) => (await WordprocessingMLPackage.load(bytes)).getMainDocumentPart().styleDefinitionsPart.getXml();
  // a control mapped to an empty node shows the placeholder: the style arrives with the save
  const mapped = await WordprocessingMLPackage.createPackage();
  await mapped.customXmlParts.load();
  mapped.customXmlParts.add('<data><empty/></data>');
  const control = mapped.body.insertParagraph('', 'End').getRange('End').insertContentControl('PlainText');
  control.xmlMapping.setMapping('/data/empty');
  assert.equal(control.isShowingPlaceholder, true);
  assert.match(await stylesOf(await mapped.save()), /w:styleId="PlaceholderText"/);
  // so does one set through placeholderText
  const set = await WordprocessingMLPackage.createPackage();
  set.body.insertParagraph('', 'End').getRange('End').insertContentControl('PlainText').placeholderText = 'Type a name';
  assert.match(await stylesOf(await set.save()), /w:styleId="PlaceholderText"/);
  // and a document with no placeholder gets none
  const none = await WordprocessingMLPackage.createPackage();
  none.body.insertParagraph('Text', 'End').insertContentControl('PlainText');
  assert.doesNotMatch(await stylesOf(await none.save()), /w:styleId="PlaceholderText"/);
});

// CR-002 section 40, item 2, for content a caller brings: objects CR-008 (0.3.1) lists every Office
// extension namespace a marshalled part uses in mc:Ignorable. Word 2010 refuses a document whose
// w15:appearance is not declared (test/README.md check 36), so this is what makes inserted content
// openable there; before 0.3.1 insertOoxml and insertXml declared nothing.
test('content brought in by insertOoxml has its extension namespaces declared ignorable (check 36)', async () => {
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"';
  const ignorableOf = async (pkg) => /mc:Ignorable="([^"]*)"/.exec(/<w:document [^>]*>/.exec(await pkg.getMainDocumentPart().getXml())[0])?.[1] ?? null;
  // an element (check 36a's content) and an attribute (36c's)
  const control = await WordprocessingMLPackage.createPackage();
  control.body.insertParagraph('Kept.', 'End');
  assert.equal(await ignorableOf(control), null, 'nothing to declare yet');
  await control.body.insertOoxml(`<w:sdt ${NS}><w:sdtPr><w:id w:val="5"/><w15:appearance w15:val="tags"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>In a control</w:t></w:r></w:p></w:sdtContent></w:sdt>`, 'End');
  assert.equal(await ignorableOf(control), 'w15');
  assert.match(await control.getMainDocumentPart().getXml(), /<w15:appearance w15:val="tags"\/>/);
  const revision = await WordprocessingMLPackage.createPackage();
  await revision.body.insertOoxml(`<w:p ${NS}><w:ins w:id="1" w:author="A" w:date="2026-10-04T10:00:00Z" w16du:dateUtc="2026-10-04T00:00:00Z"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>`, 'End');
  assert.equal(await ignorableOf(revision), 'w16du');
  // and through a save
  const again = await WordprocessingMLPackage.load(await control.save());
  assert.equal(await ignorableOf(again), 'w15');
});
