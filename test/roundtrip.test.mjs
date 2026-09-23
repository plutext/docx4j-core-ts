import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OpcPackage, WordprocessingMLPackage, SpreadsheetMLPackage, ZipPartStore, FlatOpcPartStore, MemoryPartSink, XmlPart, RelationshipsPart, unmarshalPackage, ContentTypes } from '../dist/index.mjs';
import { fixture, bytesEqual, plain } from './helpers.mjs';

async function checkRoundTrip(name, touch) {
  const original = await fixture(name);
  const pkg = await OpcPackage.load(original);
  const touched = touch ? await touch(pkg) : [];
  const saved = await pkg.save();
  const before = new ZipPartStore(original);
  const after = new ZipPartStore(saved);
  const beforeNames = new Set(before.partNames());
  const afterNames = [...after.partNames()];
  assert.equal(afterNames[0], '[Content_Types].xml', 'content types first');
  assert.deepEqual(new Set(afterNames), beforeNames, 'same entries');
  for (const n of afterNames) {
    if (n === '[Content_Types].xml' || n.endsWith('.rels') || touched.includes('/' + n)) continue;
    assert.ok(bytesEqual(before.loadSync(n), after.loadSync(n)), `${n} byte-identical`);
  }
  const reloaded = await OpcPackage.load(saved);
  assert.equal(reloaded.constructor, pkg.constructor);
  assert.equal(reloaded.parts.size, pkg.parts.size);
  // relationships survive re-marshalling
  for (const part of pkg.parts) {
    const other = reloaded.getPart(part.partName);
    assert.ok(other, `${part.partName} reloaded`);
    assert.equal(other.constructor, part.constructor, `${part.partName} class`);
    assert.equal(other.contentType, part.contentType, `${part.partName} content type`);
    if (part.relationshipsPart) {
      assert.deepEqual(plain(other.relationshipsPart.list), plain(part.relationshipsPart.list), `${part.partName} rels`);
    }
  }
  assert.deepEqual(plain(reloaded.relationshipsPart.list), plain(pkg.relationshipsPart.list));
  return { pkg, reloaded, saved };
}

test('round trip: untouched docx is byte-identical part by part', async () => {
  await checkRoundTrip('loadAndSave.docx');
  await checkRoundTrip('HelloWordOnline.docx');
  await checkRoundTrip('header-no-rels.docx');
  await checkRoundTrip('hyperlink_dupe.docx');
});

test('round trip: untouched pptx and xlsx', async () => {
  await checkRoundTrip('loadAndSave.pptx');
  await checkRoundTrip('loadAndSave.xlsx');
});

test('round trip: resolving fonts reads the theme, settings and font table without touching them', async () => {
  // CR-001 Phase B step 4: the selector and the mapper read those parts with readContents(),
  // as the PropertyResolver reads the styles part, so a document whose fonts a caller merely
  // asked about still saves byte for byte.  Only the main document part is unmarshalled, by
  // the names walk.
  for (const name of ['invoice.docx', 'HelloWordOnline.docx']) {
    await checkRoundTrip(name, async (pkg) => {
      const main = pkg.getMainDocumentPart();
      const mapper = await main.getFontMapper();
      assert.ok(mapper.size > 0, `${name}: nothing mapped`);
      assert.ok((await main.fontsInUse()).size > 0, `${name}: no fonts in use`);
      assert.ok(main.themePart, `${name} has a theme part`);
      assert.equal(main.themePart.isUnmarshalled, false, 'the theme part stays untouched');
      assert.equal(main.fontTablePart.isUnmarshalled, false, 'the font table stays untouched');
      assert.equal(main.styleDefinitionsPart.isUnmarshalled, false, 'the styles part stays untouched');
      return [main.partName.name];
    });
  }
});

test('round trip: an unmarshalled part is re-marshalled and deep-equal after reload', async () => {
  const { pkg, reloaded } = await checkRoundTrip('loadAndSave.docx', async (pkg) => {
    await pkg.getMainDocumentPart().getContents();
    await pkg.getMainDocumentPart().styleDefinitionsPart.getContents();
    return ['/word/document.xml', '/word/styles.xml'];
  });
  const doc1 = pkg.getMainDocumentPart().contents;
  const doc2 = await reloaded.getMainDocumentPart().getContents();
  assert.deepEqual(plain(doc2), plain(doc1));
  const st1 = pkg.getMainDocumentPart().styleDefinitionsPart.contents;
  const st2 = await reloaded.getMainDocumentPart().styleDefinitionsPart.getContents();
  assert.deepEqual(plain(st2), plain(st1));
  // a change made it through
  const text = (await reloaded.getMainDocumentPart().getXml());
  assert.ok(text.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'));
});

test('round trip: flat OPC out and back', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('loadAndSave.docx'));
  const flat = await pkg.saveFlatOpc();
  assert.ok(flat.startsWith('<?xml version="1.0" standalone="yes"?><?mso-application progid="Word.Document"?><pkg:package'));
  // the objects package can type it
  const typed = await unmarshalPackage(flat);
  const parts = typed.value.part;
  assert.equal(parts.length, pkg.parts.size + [...pkg.parts].filter((p) => p.relationshipsPart && p.relationshipsPart.size > 0).length + 1);
  const docPart = parts.find((p) => p.name === '/word/document.xml');
  assert.equal(docPart.contentType, ContentTypes.WORDPROCESSINGML_DOCUMENT);
  // typed when the objects package can unmarshal the part as is, else DOM: both carry the root name
  const any = docPart.xmlData.any;
  assert.equal(any.localName ?? any.name.localPart, 'document');
  // and this package can load it
  const back = await OpcPackage.load(flat);
  assert.ok(back instanceof WordprocessingMLPackage);
  assert.equal(back.parts.size, pkg.parts.size);
  const store = new FlatOpcPartStore(flat);
  assert.equal(store.contentTypeOf('word/media/image1.png'), ContentTypes.IMAGE_PNG);
  const img1 = await pkg.getPart('/word/media/image1.png').getBytes();
  const img2 = await back.getPart('/word/media/image1.png').getBytes();
  assert.ok(bytesEqual(img1, img2), 'binary part survives base64');
  const doc = await back.getMainDocumentPart().getContents();
  assert.deepEqual(plain(doc), plain(await pkg.getMainDocumentPart().getContents()));
  // flat -> zip -> flat again
  const zipped = await back.save();
  const again = await OpcPackage.load(zipped);
  assert.equal(again.parts.size, pkg.parts.size);
  // typed PackageElement as a source
  const fromTyped = await OpcPackage.load(new FlatOpcPartStore(typed));
  assert.equal(fromTyped.parts.size, pkg.parts.size);
});

test('round trip: memory sink and unmarshalAll', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('HelloWordOnline.docx'));
  await pkg.unmarshalAll();
  for (const p of pkg.parts) if (p instanceof XmlPart && !(p instanceof RelationshipsPart)) assert.ok(p.isUnmarshalled, p.partName.name);
  const store = await pkg.saveTo(new MemoryPartSink());
  assert.ok(store.has('word/document22.xml'));
  assert.equal(store.contentTypes.get('word/document22.xml'), ContentTypes.WORDPROCESSINGML_DOCUMENT);
  const back = await OpcPackage.load(store);
  assert.equal(back.parts.size, pkg.parts.size);
});

test('a re-marshalled part declares every prefix its mc:Ignorable names', async () => {
  // Excel's xl/workbook.xml carries mc:Ignorable="x15 xr xr6 xr10 xr2" and xr:revisionPtr, whose
  // xr6/xr10 attributes and whose xr2:uid the object model does not bind; dropping them dropped
  // their declarations too, and an mc:Ignorable naming an undeclared prefix is what Excel offers
  // to repair (CR-001 section 17). XmlPart re-declares them from the source root.
  // the prefixes a root's mc:Ignorable names, each of which it must also declare; [] when the
  // marshalled root has no mc:Ignorable (the model binds the attribute on w:document, w:settings
  // and CT_Workbook, but not on CT_Worksheet, which therefore loses it - lossy, but not a repair
  // trigger, since nothing is left undeclared)
  const declarations = (xml, tag) => {
    const at = xml.indexOf(tag);
    assert.ok(at >= 0, `${tag} in the saved part`);
    const root = xml.slice(at, xml.indexOf('>', at) + 1);
    const ignorable = /mc:Ignorable="([^"]*)"/.exec(root);
    if (!ignorable) return [];
    const prefixes = ignorable[1].split(/\s+/).filter((p) => p !== '');
    for (const p of prefixes) assert.ok(root.includes(`xmlns:${p}="`), `xmlns:${p} declared on ${tag}: ${root}`);
    return prefixes;
  };
  const decode = (bytes) => new TextDecoder().decode(bytes);

  const xlsx = await SpreadsheetMLPackage.load(await fixture('loadAndSave.xlsx'));
  const workbookPart = xlsx.getWorkbookPart();
  await workbookPart.getContents();
  const sheet = xlsx.worksheetParts[0];
  await sheet.getContents();
  const saved = new ZipPartStore(await xlsx.save());
  assert.deepEqual(declarations(decode(saved.loadSync('xl/workbook.xml')), '<workbook'), ['x15', 'xr', 'xr6', 'xr10', 'xr2']);
  declarations(decode(saved.loadSync(sheet.partName.name.slice(1))), '<worksheet');

  const docx = await WordprocessingMLPackage.load(await fixture('loadAndSave.docx'));
  await docx.getMainDocumentPart().getContents();
  await docx.getMainDocumentPart().documentSettingsPart.getContents();
  const savedDocx = new ZipPartStore(await docx.save());
  assert.ok(declarations(decode(savedDocx.loadSync('word/document.xml')), '<w:document').includes('w14'));
  assert.ok(declarations(decode(savedDocx.loadSync('word/settings.xml')), '<w:settings').length > 0);
});

// docx4j-core-tests SettingsDocIdOrderTest (its CR for 17.2.1, dae2dfc8b): CT_Settings declares
// w14:docId before w15:chartTrackingRefBased, which is the order Word desktop writes.
// HelloWordOnline.docx carries Word Online's order, the other way round, so re-marshalling it is
// what shows the schema's order rather than the input's. CR-001 section 17.7.
test('a re-marshalled settings part puts w14:docId before w15:chartTrackingRefBased', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('HelloWordOnline.docx'));
  const settings = [...pkg.parts.values()].find((p) => p.partName.name.endsWith('settings.xml'));
  const source = new TextDecoder().decode(new ZipPartStore(await fixture('HelloWordOnline.docx')).loadSync('word/settings.xml'));
  assert.deepEqual(source.match(/<w1[45]:(docId|chartTrackingRefBased)/g),
    ['<w15:chartTrackingRefBased', '<w14:docId', '<w15:docId'], 'the input is Word Online-ordered');

  await settings.getContents();
  const out = new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync('word/settings.xml'));
  assert.deepEqual(out.match(/<w1[45]:(docId|chartTrackingRefBased)/g),
    ['<w14:docId', '<w15:chartTrackingRefBased', '<w15:docId']);
});
