import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OpcPackage, WordprocessingMLPackage, ZipPartStore, FlatOpcPartStore, MemoryPartSink, XmlPart, RelationshipsPart, unmarshalPackage, ContentTypes } from '../dist/index.mjs';
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
