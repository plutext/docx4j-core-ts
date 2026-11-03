// CR-020: bound controls as Word repairs and formats them, the port of docx4j CR-058. docx4j's probes and Word 365's
// saves of them (test/store-items/, with a SHA-1 manifest): a w:storeItemID naming no part bound to the part its XPath
// fits, by every reader and the bind, the bind writing that part's ID back as Word's save does (docx4j's
// StoreItemRepairTest); a bound date's am/pm and AM/PM as Word shows them, and the document's default language where
// w:lid is absent (docx4j's BoundDateFormatTest); and checks 73 and 73b's snippet through the Word shim, held to Word's
// runs (test/fixtures/check73/, check73b/).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, strFromU8 } from 'fflate';
import { parseHTML } from 'linkedom';
import {
  WordprocessingMLPackage, Texts, PlainText, StyleDefinitionsPart, XmlPart, firstStep, formatDate, selectEntryNode,
} from '../dist/index.mjs';
import { Word } from '../dist/office-js/index.mjs';
import { swappedRelationships } from './fixtures/check73b/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, 'store-items');
const manifest = JSON.parse(await readFile(join(dir, 'manifest.json'), 'utf8'));
const bytesOf = async (name) => new Uint8Array(await readFile(join(dir, name)));
const runOf = async (check, name) => JSON.parse(await readFile(join(here, 'fixtures', check, name), 'utf8'));

const PART1 = '{B1D0F00D-0000-4000-8000-000000000001}';
const PART2 = '{B1D0F00D-0000-4000-8000-000000000002}';
const STALE = '{B1D0F00D-0000-4000-8000-0000000000FF}';
const ZONE = 'Australia/Sydney';                                     // Jason's machine at the run: UTC+11 in October

/** A probe (or its bytes) loaded, its parts parsed, `configure` given the package, its body read, and bound unless `bind` is false. */
async function open(source, { bind = true, configure } = {}) {
  const pkg = await WordprocessingMLPackage.load(typeof source === 'string' ? await bytesOf(source) : source);
  configure?.(pkg);
  await pkg.customXmlParts.load();
  const body = await pkg.getBody();
  const result = bind ? await pkg.customXmlParts.applyBindings() : undefined;
  const control = (tag) => {
    const found = body.contentControls.find((c) => c.tag === tag);
    assert.ok(found, `control ${tag}`);
    return found;
  };
  return { pkg, body, result, control };
}

/** Each control as a reader sees it: what it shows and its binding. */
const controlsOf = (body) => body.contentControls.map((c) => ({
  tag: c.tag, text: c.text, storeItemID: c.xmlMapping.storeItemID, isMapped: c.xmlMapping.isMapped, showingPlaceholder: c.isShowingPlaceholder,
}));

/** The `w:storeItemID` of each control in a saved document.xml, by tag. */
const savedIds = async (pkg) => {
  const xml = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
  return Object.fromEntries([...xml.matchAll(/<w:sdtPr>[\s\S]*?<\/w:sdtPr>/g)].map(([pr]) => [/<w:tag w:val="([^"]*)"/.exec(pr)?.[1], /w:storeItemID="([^"]*)"/.exec(pr)?.[1]]));
};

const plainText = async (source) => PlainText.of(await Texts.of(await WordprocessingMLPackage.load(typeof source === 'string' ? await bytesOf(source) : source), { view: 'accepted' }));

test('the probes and Word 365\'s saves are as the manifest has them', async () => {
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await bytesOf(name)).digest('hex'), sha1, name);
  assert.deepEqual(Object.keys(manifest.saves), ['bind-plain.docx', 'bind-plain-two-parts.docx', 'bind-plain-no-part.docx', 'typed-controls-dates.docx', 'typed-controls-twelve.docx']);
  for (const [probe, save] of Object.entries(manifest.saves)) assert.ok(manifest.files.some((f) => f.name === probe) && manifest.files.some((f) => f.name === save), probe);
});

// --- the store item resolver (docx4j StoreItemRepairTest) ---------------------------------------------------------------

test("StoreItemRepairTest: bind-plain's B7, whose w:storeItemID names no part, is bound to the one part its XPath fits, and gets that part's ID, as Word's save has it", async () => {
  const { pkg, result, control } = await open('bind-plain.docx');
  assert.equal(control('B1').text, 'B1_XML_VALUE');
  assert.equal(control('B7').text, 'B7_XML_VALUE');
  assert.equal(control('B7').xmlMapping.storeItemID, PART1, 'the ID written back (D2)');
  const ids = await savedIds(pkg);
  assert.equal(ids.B1, PART1);
  assert.equal(ids.B7, PART1);
  // D4: the repair said, the control and the part named
  assert.deepEqual(result.notes, [`B7: store item ${STALE} names no custom XML part; /customXml/item1.xml ${PART1} fits /p:root[1]/p:b7[1] and is bound, its ID written into the control, as Word repairs it`]);
});

test('StoreItemRepairTest: the text model shows the value Word shows, and writes nothing', async () => {
  const text = await plainText('bind-plain.docx');
  assert.ok(text.includes('B7_XML_VALUE') && text.includes('B1_XML_VALUE'), text);
  assert.ok(!text.includes('B7_STALE_SURFACE'), text);
});

test('StoreItemRepairTest: where two parts fit, the first in part-name order whose XPath selects a node (D7, check 73b): B7 part 1, B11 part 2', async () => {
  const { pkg, result, control } = await open('bind-plain-two-parts.docx');
  assert.equal(control('B1').text, 'B1_XML_VALUE');
  assert.equal(control('B7').text, 'B7_XML_VALUE_PART1');
  assert.equal(control('B11').text, 'B11_XML_VALUE_PART2');
  assert.deepEqual(await savedIds(pkg), { B1: PART1, B7: PART1, B11: PART2 });
  assert.equal(result.notes.length, 2);
  assert.match(result.notes[0], /^B7: .*\/customXml\/item1\.xml .*, the first of the 2 that fit \/p:root\[1\]\/p:b7\[1\] to select a node, is bound/);
  assert.match(result.notes[1], /^B11: .*\/customXml\/item2\.xml .*, the first of the 2 that fit \/p:root\[1\]\/p:b11\[1\] to select a node, is bound/);
  const text = await plainText('bind-plain-two-parts.docx');
  assert.ok(text.includes('B7_XML_VALUE_PART1') && text.includes('B11_XML_VALUE_PART2'), text);
});

test('StoreItemRepairTest: where no part fits, the control is left with its surface and its ID, and the rest binds', async () => {
  const { pkg, result, control } = await open('bind-plain-no-part.docx');
  assert.equal(control('B1').text, 'B1_XML_VALUE');
  assert.equal(control('B7').text, 'B7_STALE_SURFACE');
  assert.equal(control('B12').text, 'B12_STALE_SURFACE');
  assert.deepEqual(await savedIds(pkg), { B1: PART1, B7: STALE, B12: STALE });
  assert.deepEqual(result, {
    bound: 3, updated: 1, skipped: 2, notes: [
      `B7: store item ${STALE} names no custom XML part, and none fits /p:root[1]/p:b7[1]; left as it is`,
      `B12: store item ${STALE} names no custom XML part, and none fits /p:nothing[1]/p:b7[1]; left as it is`,
    ],
  });
  const text = await plainText('bind-plain-no-part.docx');
  assert.ok(text.includes('B1_XML_VALUE') && text.includes('B7_STALE_SURFACE') && text.includes('B12_STALE_SURFACE'), text);
  assert.ok(!text.includes('B7_XML_VALUE'), text);
});

test("each store-item probe after applyBindings is Word's save, control by control: what it shows, its w:storeItemID, isMapped; B1 of bind-plain is Jason's typing", async () => {
  for (const probe of ['bind-plain.docx', 'bind-plain-two-parts.docx', 'bind-plain-no-part.docx']) {
    const mine = controlsOf((await open(probe)).body);
    const word = controlsOf((await open(manifest.saves[probe], { bind: false })).body);
    if (probe === 'bind-plain.docx') {
      // the share's README, step 3: something typed into B1 before the save, which Word wrote to the part too
      assert.equal(word[0].text, 'B1 typed something');
      mine.shift();
      word.shift();
    }
    assert.deepEqual(mine, word, probe);
  }
});

test("the text model of each store-item probe is the text model of Word's save, but B1's typing and B8: a control without w:text reads its surface (REQ-509, docx4j's rule) where Word shows the value", async () => {
  const lines = async (source) => (await plainText(source)).split('\n');
  for (const probe of ['bind-plain.docx', 'bind-plain-two-parts.docx', 'bind-plain-no-part.docx']) {
    const mine = await lines(probe);
    const word = await lines(manifest.saves[probe]);
    assert.equal(mine.length, word.length, probe);
    const differ = mine.flatMap((line, i) => (line === word[i] ? [] : [[line, word[i]]]));
    assert.deepEqual(differ, probe !== 'bind-plain.docx' ? [] : [
      ['B1 inline, stale surface: B1_XML_VALUE (end)', 'B1 inline, stale surface: B1 typed something (end)'],
      // Word reads a control bound by w:dataBinding as plain text, w:richText or not (CR-002 section 61 H4), and its save
      // gives B8 w:text; the text model reads a binding only where w:text is (CR-012 REQ-509, docx4j's BoundValues)
      ['B8 rich-text control bound to a plain value: B8_RICH_STALE_SURFACE (end)', 'B8 rich-text control bound to a plain value: B10_XML_VALUE (end)'],
    ], probe);
  }
});

test('XmlMapping: customXmlPart is the part the binding resolves to, storeItemID the ID as stored until applyBindings; isMapped false where the XPath selects nothing (D8, check 73b\'s B4) or nothing fits', async () => {
  const plain = await open('bind-plain.docx', { bind: false });
  const b7 = plain.control('B7').xmlMapping;
  assert.equal(b7.storeItemID, STALE, 'D3: as stored');
  assert.equal(b7.customXmlPart.id, PART1, 'D3: the resolved part, as Office JS gives it');
  assert.equal(b7.customXmlNode.text, 'B7_XML_VALUE');
  assert.equal(b7.isMapped, true);
  const b4 = plain.control('B4').xmlMapping;
  assert.equal(b4.customXmlPart.id, PART1);
  assert.equal(b4.customXmlNode, undefined);
  assert.equal(b4.isMapped, false, 'Word: false (check 73b)');
  assert.ok(b4.dataBinding, 'the control still carries its binding');

  const two = await open('bind-plain-two-parts.docx', { bind: false });
  assert.equal(two.control('B7').xmlMapping.customXmlPart.id, PART1);
  assert.equal(two.control('B11').xmlMapping.customXmlPart.id, PART2);

  const none = await open('bind-plain-no-part.docx', { bind: false });
  for (const tag of ['B7', 'B12']) {
    const mapping = none.control(tag).xmlMapping;
    assert.equal(mapping.customXmlPart, undefined, tag);
    assert.equal(mapping.customXmlNode, undefined, tag);
    assert.equal(mapping.isMapped, false, tag);
    assert.equal(mapping.storeItemID, STALE, tag);
  }
  assert.equal(none.control('B1').xmlMapping.isMapped, true);
});

test('updateFromContentControls writes a stale control\'s text to the part it resolves to, and leaves its ID (D2)', async () => {
  const { pkg, control } = await open('bind-plain-two-parts.docx', { bind: false });
  const result = await pkg.customXmlParts.updateFromContentControls();
  assert.equal(result.updated, 3);
  const [part1, part2] = pkg.customXmlParts.items;
  assert.equal(part1.selectSingleNode('/p:root[1]/p:b7[1]', "xmlns:p='urn:probe:bind'").text, 'B7_STALE_SURFACE');
  assert.equal(part2.selectSingleNode('/p:root[1]/p:b7[1]', "xmlns:p='urn:probe:bind'").text, 'B7_XML_VALUE_PART2');
  assert.equal(part2.selectSingleNode('/p:root[1]/p:b11[1]', "xmlns:p='urn:probe:bind'").text, 'B11_STALE_SURFACE');
  assert.equal(control('B7').xmlMapping.storeItemID, STALE);
  assert.equal(control('B11').xmlMapping.storeItemID, STALE);
});

test('no reader writes: a probe read by the text model, the parts and every mapping saves as it saved before, byte for byte', async () => {
  for (const probe of ['bind-plain.docx', 'bind-plain-two-parts.docx', 'bind-plain-no-part.docx']) {
    const before = unzipSync(await (await WordprocessingMLPackage.load(await bytesOf(probe))).save());
    const pkg = await WordprocessingMLPackage.load(await bytesOf(probe));
    await Texts.of(pkg);
    await pkg.customXmlParts.load();
    for (const part of pkg.parts) if (part instanceof XmlPart && !part.partName.name.endsWith('.rels')) assert.equal(part.isUnmarshalled, false, part.partName.name);
    const after = unzipSync(await pkg.save());
    assert.deepEqual(Object.keys(after).sort(), Object.keys(before).sort(), probe);
    for (const name of Object.keys(before)) assert.deepEqual(after[name], before[name], `${probe} ${name}`);

    // the content API's reads: the main part is unmarshalled, so it is written again, as it is after any read of the body
    const read = await WordprocessingMLPackage.load(await bytesOf(probe));
    await read.getBody();
    const unmarshalled = unzipSync(await read.save());
    const mapped = (await open(probe, { bind: false })).body;
    for (const c of mapped.contentControls) void [c.xmlMapping.customXmlPart, c.xmlMapping.customXmlNode, c.xmlMapping.isMapped, c.xmlMapping.storeItemID];
    const afterMappings = unzipSync(await mapped.package_.save());
    for (const name of Object.keys(unmarshalled)) assert.deepEqual(afterMappings[name], unmarshalled[name], `${probe} ${name} after the mappings`);
  }
});

test('a stale ID is resolved again once the parts change: a part added that fits, then removed', async () => {
  const { pkg, control } = await open('bind-plain-no-part.docx', { bind: false });
  assert.equal(control('B12').xmlMapping.customXmlPart, undefined);
  const added = pkg.customXmlParts.add("<p:nothing xmlns:p='urn:probe:bind'><p:b7>B12_NEW_PART</p:b7></p:nothing>");
  await pkg.customXmlParts.load();
  assert.equal(control('B12').xmlMapping.customXmlPart, added);
  assert.equal(control('B12').xmlMapping.customXmlNode.text, 'B12_NEW_PART');
  assert.equal(control('B7').xmlMapping.customXmlPart, undefined, 'B7 is in another namespace still');
  pkg.customXmlParts.removePart(added);
  assert.equal(control('B12').xmlMapping.customXmlPart, undefined);
  assert.equal(control('B12').xmlMapping.storeItemID, STALE, 'removePart leaves a binding it did not name');
});

test("the resolver's first step: docx4j's firstStep, the binding's prefixes, else the objects package's; no named first step fits every part", () => {
  const p = "xmlns:p='urn:probe:bind'";
  assert.deepEqual(firstStep('/p:root[1]/p:b7[1]', p), { namespaceURI: 'urn:probe:bind', localName: 'root' });
  assert.deepEqual(firstStep(' p:root/p:b7', p), { namespaceURI: 'urn:probe:bind', localName: 'root' });
  assert.deepEqual(firstStep('/root[1]/b7', ''), { namespaceURI: '', localName: 'root' });
  assert.deepEqual(firstStep('/w:document/w:body', ''), { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localName: 'document' });
  assert.deepEqual(firstStep('/q:root', ''), { namespaceURI: '', localName: 'root' }, 'a prefix nothing maps');
  for (const xpath of ['//p:b7', '/*', '/p:*', '/@a', '/child::p:root', '.', '..', '', undefined]) assert.equal(firstStep(xpath, p), undefined, String(xpath));
});

test('an od:xpath entry whose storeItemID names no part is read from the part its XPath fits; the XHTML pass binds it, and says so', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  body.insertParagraph('placeholder', 'End');
  const control = body.insertContentControl('RichText');
  control.tag = 'od:xpath=x1&od:ContentType=application/xhtml+xml';
  const data = pkg.customXmlParts.add('<data><body>&lt;p&gt;Bound &lt;b&gt;here&lt;/b&gt;&lt;/p&gt;</body></data>');
  pkg.customXmlParts.add(`<xpaths xmlns="http://opendope.org/xpaths"><xpath id="x1"><dataBinding storeItemID="${STALE}" xpath="/data/body" prefixMappings=""/></xpath></xpaths>`);
  await pkg.customXmlParts.load();
  const selected = selectEntryNode('x1', pkg.customXmlParts);
  assert.equal(selected.node.text, '<p>Bound <b>here</b></p>');
  assert.equal(selected.repair, `od:xpath=x1: store item ${STALE} names no custom XML part; ${data.part.partName.name} ${data.id} fits /data/body and is used`);
  const result = await pkg.customXmlParts.applyBindings({ html: { parser: (html) => parseHTML(html).document } });
  assert.equal(control.text, 'Bound here');
  assert.deepEqual(result.notes, [selected.repair]);
});

// --- bound dates (docx4j BoundDateFormatTest) ------------------------------------------------------------------------------

/** Each date control's text after applyBindings, by tag. */
async function dates(probe, configure) {
  const { body } = await open(probe, { configure });
  return Object.fromEntries(body.contentControls.map((c) => [c.tag, c.text]));
}

const converted = (pkg) => { pkg.dateTimeZone = ZONE; pkg.boundDatesConvertedAsWord = true; };

test("BoundDateFormatTest: Word's formats, nominally (no zone): am/pm and AM/PM the upper-case names; a value that is no date as stored; the empty value the placeholder", async () => {
  const shown = await dates('typed-controls-dates.docx');
  assert.deepEqual(shown, {
    DF1: 'Thu 22 Oct 2026',
    DF2: '3:05 PM',                                                   // am/pm, upper case as Word; was "3:05 a5/p5"
    DF3: '12:05 PM',                                                  // AM/PM at noon; was "12:05 A10/P10"
    DF4: '15:05',
    DF5: '22 Oktober 2026',
    DF6: 'Thursday, 22 October 2026',
    DF7: '22 October 2026',                                           // no w:lid and no document default language: en-US
    DF8: 'not-a-date',
    DF9: '2026-13-45',
    DF10: '',                                                         // the placeholder: Office JS's text of a control showing it
    DF11: '2026-10-22',
    DF12: 'Day 22 of October',
    DF13: 'Okt. 2026',
  });
  const { control } = await open('typed-controls-dates.docx');
  assert.equal(control('DF10').isShowingPlaceholder, true);
  assert.deepEqual(await dates('typed-controls-dates.docx', (pkg) => { pkg.dateTimeZone = ZONE; }), shown, 'a zone alone converts nothing');
});

test("BoundDateFormatTest: in Word's zone, converted as Word converts (CR-018), every date control is Word's save", async () => {
  for (const probe of ['typed-controls-dates.docx', 'typed-controls-twelve.docx']) {
    const mine = controlsOf((await open(probe, { configure: converted })).body);
    const word = controlsOf((await open(manifest.saves[probe], { bind: false })).body);
    assert.deepEqual(mine, word, probe);
  }
  const shown = await dates('typed-controls-dates.docx', converted);
  assert.deepEqual([shown.DF1, shown.DF2, shown.DF3, shown.DF4, shown.DF11], ['Thu 22 Oct 2026', '2:05 AM', '11:05 PM', '02:05', '2026-10-22']);
});

test("BoundDateFormatTest: the twelve o'clock probe, nominally and in Word's zone (Word's six)", async () => {
  assert.deepEqual(await dates('typed-controls-twelve.docx'), {
    TW1: '1:05 AM', TW2: '01:05 PM', TW3: '12:05 PM 22 Oct', TW4: '12:05 PM 22 Oct', TW5: '12:05 PM 22 Oct', TW6: '01:05',
  });
  assert.deepEqual(await dates('typed-controls-twelve.docx', converted), {
    TW1: '12:05 PM', TW2: '12:05 AM', TW3: '12:05 PM 22 Oct', TW4: '12:05 PM 22 Oct', TW5: '4:05 AM 23 Oct', TW6: '12:05',
  });
});

test('BoundDateFormatTest: without a w:lid the document default language is used (D5), the run\'s not consulted; a w:lid still wins', async () => {
  const germanDefaults = (pkg) => {
    const styles = new StyleDefinitionsPart();
    styles.setXml('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="de-DE"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>');
    pkg.getMainDocumentPart().addTargetPart(styles);
  };
  const shown = await dates('typed-controls-dates.docx', germanDefaults);
  assert.equal(shown.DF7, '22 Oktober 2026', 'no w:lid, run fr-FR: the document\'s German');
  assert.equal(shown.DF5, '22 Oktober 2026', 'its own w:lid de-DE');
  assert.equal(shown.DF6, 'Thursday, 22 October 2026', 'its own w:lid en-AU');
});

test('formatDate: am/pm and AM/PM in any mix of cases are the locale\'s day period upper-cased; A/P and a/p stay literal', () => {
  const at = (h, m) => new Date(Date.UTC(2026, 9, 22, h, m));
  assert.equal(formatDate('h:mm am/pm', at(15, 5)), '3:05 PM');
  assert.equal(formatDate('h:mm AM/PM', at(12, 5)), '12:05 PM');
  assert.equal(formatDate('hh:mm aM/Pm', at(0, 5)), '12:05 AM');
  assert.equal(formatDate('h:mm am/pm', at(9, 0), 'en-AU'), '9:00 AM', 'CLDR\'s en-AU names are am and pm');
  assert.equal(formatDate('h:mm am/pm', at(21, 0), 'de-DE'), '9:00 PM');
  assert.equal(formatDate('h A/P a/p', at(15, 5)), '3 A/P a/p');
  assert.equal(formatDate("'am/pm' h", at(15, 5)), 'am/pm 3', 'quoted text is literal');
});

// --- checks 73 and 73b through the Word shim ------------------------------------------------------------------------

const check73b = new Function(`${await readFile(join(here, 'fixtures', 'check73b', 'check73b-script-lab.js'), 'utf8')}\nreturn CHECK73B;`)();

/** Check 73b's snippet through the shim on a probe after applyBindings, as Word reads it after the open. */
async function shimRun(source) {
  const { pkg } = await open(source);
  return check73b.runCheck((fn) => Word.run(pkg, fn));
}

/** A run's controls as compared: the shim's ids are its own. */
const comparable = (run) => run.controls.map(({ id, ...rest }) => rest);

/** Each control's `w:storeItemID` in a run's body OOXML, by tag. */
const bodyIds = (run) => Object.fromEntries(run.bodyBindings.map((b) => [b.tag, b.dataBinding.map((d) => /w:storeItemID="([^"]*)"/.exec(d)[1])]));

test("check 73b's snippet through the shim, after applyBindings, is Word's run after the open: each control's text, mapping, customXmlPart and customXmlNode; the differences listed", async () => {
  const twoParts = await bytesOf('bind-plain-two-parts.docx');
  const cases = [
    ['bind-plain.docx', '73b-bind-plain-word365-run.json'],
    ['bind-plain-two-parts.docx', '73b-two-parts-1-word365-run.json'],
    ['bind-plain-two-parts.docx', '73b-two-parts-2-word365-run.json'],
    ['bind-plain-two-parts.docx', '73b-two-parts-3-word365-run.json'],
    [swappedRelationships(twoParts), '73b-two-parts-swapped-word365-run.json'],
  ];
  for (const [source, runName] of cases) {
    const word = await runOf('check73b', runName);
    const shim = await shimRun(source);
    const mine = comparable(shim);
    const theirs = comparable(word);
    // B4's customXmlNode: Word throws GeneralException; the shim has no node to give, and the snippet's load fails on undefined
    const b4 = theirs.findIndex((c) => c.tag === 'B4');
    if (b4 >= 0) {
      assert.match(theirs[b4].customXmlNode, /^error: GeneralException/);
      assert.match(mine[b4].customXmlNode, /^error: /);
      mine[b4].customXmlNode = theirs[b4].customXmlNode;
    }
    // B8, a run-level control with no kind bound by w:dataBinding: plain text to Word, as a block one is (CR-002 section 61
    // H4), where this package's rule covers block controls only; CR-020 section 8.1, CR-002 section 61.7
    const b8 = theirs.findIndex((c) => c.tag === 'B8');
    if (b8 >= 0) {
      assert.deepEqual([theirs[b8].type, mine[b8].type], ['PlainText', 'RichText']);
      mine[b8].type = theirs[b8].type;
    }
    assert.deepEqual(mine, theirs, runName);
    assert.deepEqual(bodyIds(shim), bodyIds(word), `${runName}: the IDs in the body's OOXML`);
    // listed: the parts collection, which the shim lacks; and the lookups, of which Word's body.getContentControls() alone finds these controls
    assert.match(shim.parts, /^error: NotImplemented/);
    assert.deepEqual(word.lookups.map((l) => l.count > 0), [false, false, true, false, false], runName);
    assert.deepEqual(shim.lookups.map((l) => l.via), word.lookups.map((l) => l.via));
    assert.equal(shim.lookups[2].count, word.lookups[2].count, runName);
  }
});

test("check 73b's snippet through the shim before applyBindings: the stale controls' parts and nodes are Word's already; the text and the body's OOXML are the file's, where Word's open has repaired them", async () => {
  const shim = await check73b.runCheck(async (fn) => Word.run((await open('bind-plain.docx', { bind: false })).pkg, fn));
  const word = await runOf('check73b', '73b-bind-plain-word365-run.json');
  const b7 = (run) => run.controls.find((c) => c.tag === 'B7');
  assert.deepEqual([b7(shim).customXmlPart, b7(shim).customXmlNode, b7(shim).mapping], [b7(word).customXmlPart, b7(word).customXmlNode, b7(word).mapping]);
  assert.equal(b7(shim).text, 'B7_STALE_SURFACE');
  assert.equal(b7(word).text, 'B7_XML_VALUE');
  assert.deepEqual(bodyIds(shim).B7, [STALE]);
  assert.deepEqual(bodyIds(word).B7, [PART1]);
});

test("check 73's runs: the IDs Word's body OOXML held after the open are the shim's after applyBindings, but two-parts' B7, which check 73's open gave part 2 (section 5.1, unexplained)", async () => {
  for (const [probe, runName] of [
    ['bind-plain.docx', '73-bind-plain-word365-run.json'],
    ['bind-plain-two-parts.docx', '73-bind-plain-two-parts-word365-run.json'],
    ['bind-plain-no-part.docx', '73-bind-plain-no-part-word365-run.json'],
  ]) {
    const word = bodyIds(await runOf('check73', runName));
    const shim = bodyIds(await shimRun(probe));
    if (probe === 'bind-plain-two-parts.docx') {
      assert.deepEqual(word.B7, [PART2]);
      assert.deepEqual(shim.B7, [PART1], 'check 73b\'s four fresh opens, and Word\'s save of 23 October');
      delete word.B7;
      delete shim.B7;
    }
    assert.deepEqual(shim, word, runName);
  }
});
