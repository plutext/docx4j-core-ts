// CR-025 section 8: a relationships part in another namespace, or in none, on its root or on any Relationship, is refused when it
// is read (D10), where its children in a wrong namespace loaded and a save lost the parts they named with nothing said. What else
// such a part's read drops is on pkg.warnings (D12). The rels-namespace round is in test/rels-namespace/, docx4j's copy, with a
// SHA-1 manifest; Word's saves of it show Word reading every relationship by its local name after its repair, which
// LoadOptions.repair: 'word' does (CR-031 section 2.3).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strToU8, strFromU8, unzipSync } from 'fflate';
import { OpcPackage, Texts, PlainText, InvalidFormatException, PartUnmarshalException, RelationshipsNamespaceException } from '../dist/index.mjs';
import { storedZip } from './helpers.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, 'rels-namespace');
const manifest = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8'));
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const STYLES = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const WRONG = 'urn:docx4j:wrong';

// each file's refusal: the part, and the namespace found (null for none)
const REFUSED = {
  'rn-02-plain-wrong': ['/word/_rels/document.xml.rels', WRONG],
  'rn-03-rich-wrong': ['/word/_rels/document.xml.rels', WRONG],
  'rn-04-plain-no-namespace': ['/word/_rels/document.xml.rels', null],
  'rn-05-plain-root-wrong': ['/word/_rels/document.xml.rels', WRONG],
  'rn-06-plain-children-wrong': ['/word/_rels/document.xml.rels', WRONG],
  'rn-07-plain-one-wrong': ['/word/_rels/document.xml.rels', WRONG],
  'rn-09-package-rels-wrong': ['/_rels/.rels', WRONG],
};
const CONTROL_OF = { 'rn-03-rich-wrong': 'rn-01-control-rich' };

const bytesOf = (name) => new Uint8Array(readFileSync(join(dir, name)));
const load = async (bytes, options) => {
  const pkg = await OpcPackage.load(bytes, options);
  await pkg.unmarshalAll();
  return pkg;
};
const textOf = async (pkg) => PlainText.of(await Texts.of(pkg, { include: ['headers', 'footers'], comments: true })).replace(/\s+/g, ' ').trim();
const docx = (documentRels, { styles = true } = {}) => storedZip([
  ['[Content_Types].xml', `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN}"/><Override PartName="/word/styles.xml" ContentType="${STYLES}"/></Types>`],
  ['_rels/.rels', `${XML}<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
  ['word/document.xml', `${XML}<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>WORDS</w:t></w:r></w:p></w:body></w:document>`],
  ['word/_rels/document.xml.rels', `${XML}${documentRels}`],
  ...(styles ? [['word/styles.xml', `${XML}<w:styles xmlns:w="${W}"/>`]] : []),
].map(([name, value]) => [name, strToU8(value)]));
const STYLES_REL = `<Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/>`;

test('CR-025 section 8: the round is as the manifest has it', () => {
  for (const f of manifest.files) assert.equal(createHash('sha1').update(bytesOf(f.name)).digest('hex'), f.sha1, f.name);
  assert.equal(manifest.files.length, 20);
});

test('CR-025 D10: each form is refused at the load, naming the part and the namespace found', async () => {
  for (const [stem, [partName, namespace]] of Object.entries(REFUSED)) {
    await assert.rejects(OpcPackage.load(bytesOf(`${stem}.docx`)), (e) => e instanceof RelationshipsNamespaceException
      && e instanceof InvalidFormatException && e.code === 'package.relationships_namespace' && e.partName === partName
      && e.namespace === namespace && e.message.includes(partName), stem);
  }
});

test('CR-025 D10: Word reads every relationship after its repair, so each save reads as its control does', async () => {
  for (const stem of Object.keys(REFUSED)) {
    const control = CONTROL_OF[stem] ?? 'rn-00-control-plain';
    assert.equal(await textOf(await load(bytesOf(`${stem}-word365.docx`))), await textOf(await load(bytesOf(`${control}.docx`))), stem);
  }
});

test('CR-025 D10: the controls and a package with no relationships part load and save as before', async () => {
  for (const stem of ['rn-00-control-plain', 'rn-01-control-rich', 'rn-08-rich-no-part']) {
    const source = bytesOf(`${stem}.docx`);
    const pkg = await load(source);
    assert.deepEqual(pkg.warnings, [], stem);
    if (stem !== 'rn-08-rich-no-part') assert.deepEqual(Object.keys(unzipSync(await pkg.save())).sort(), Object.keys(unzipSync(source)).sort(), stem);
  }
});

test('CR-025 D10: a flat OPC package is refused the same way', async () => {
  const flat = await (await OpcPackage.load(bytesOf('rn-01-control-rich.docx'))).saveFlatOpc();
  const i = flat.indexOf('<Relationship ', flat.indexOf('pkg:name="/word/_rels/document.xml.rels"'));
  assert.ok(i > 0);
  const moved = `${flat.slice(0, i)}<x:Relationship xmlns:x="${WRONG}" ${flat.slice(i + '<Relationship '.length)}`;
  await assert.rejects(OpcPackage.load(moved), (e) => e instanceof RelationshipsNamespaceException && e.partName === '/word/_rels/document.xml.rels');
  assert.equal(typeof (await OpcPackage.load(flat)), 'object');
});

test('CR-025 D12: what else a relationships part\'s read drops is on pkg.warnings, and a save is not refused for it', async () => {
  const pkg = await load(docx(`<Relationships xmlns="${PR}" xmlns:x="${WRONG}"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml" x:note="kept?"/><Other/><x:Other/></Relationships>`));
  const skipped = pkg.warnings.filter((w) => w.code === 'part.skipped');
  assert.deepEqual(skipped.map((w) => [w.partName, w.kind, w.name]).sort(), [
    ['/word/_rels/document.xml.rels', 'attribute', 'x:note'],
    ['/word/_rels/document.xml.rels', 'element', 'Other'],
    ['/word/_rels/document.xml.rels', 'element', 'x:Other'],
  ]);
  // the styles part it names is read, so nothing the relationships name is lost, and the save goes through
  assert.ok(pkg.getMainDocumentPart().styleDefinitionsPart);
  const saved = unzipSync(await pkg.save());
  assert.ok(saved['word/styles.xml']);
  assert.doesNotMatch(strFromU8(saved['word/_rels/document.xml.rels']), /Other|note/);
});

test('CR-025 section 8: a root with another local name is left to the runtime, as before', async () => {
  await assert.rejects(load(docx(`<x:Rels xmlns:x="${WRONG}">${STYLES_REL}</x:Rels>`)), (e) => e instanceof PartUnmarshalException && !(e instanceof RelationshipsNamespaceException));
});

test('CR-025 D12: the package relationships part\'s own is on pkg.warnings too', async () => {
  const entries = unzipSync(docx(`<Relationships xmlns="${PR}">${STYLES_REL}</Relationships>`));
  entries['_rels/.rels'] = strToU8(`${XML}<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/><Other/></Relationships>`);
  const pkg = await load(storedZip(Object.entries(entries)));
  assert.deepEqual(pkg.warnings.filter((w) => w.code === 'part.skipped').map((w) => [w.partName, w.name]), [['/_rels/.rels', 'Other']]);
});

// how many Relationship elements a relationships part holds, in any namespace or none
const relationshipsIn = (xml) => (xml.match(/<(?:[\w.-]+:)?Relationship[\s/>]/g) ?? []).length;
const zipName = (partName) => partName.slice(1);

test('CR-031 section 2.3: under repair word or all each form reads as Word\'s repair reads it, every relationship by local name, said once, and a save writes the right namespace', async () => {
  for (const [stem, [partName, namespace]] of Object.entries(REFUSED)) {
    const source = bytesOf(`${stem}.docx`);
    const count = relationshipsIn(strFromU8(unzipSync(source)[zipName(partName)]));
    const control = await textOf(await load(bytesOf(`${CONTROL_OF[stem] ?? 'rn-00-control-plain'}.docx`)));
    for (const repair of ['word', 'all']) {
      const pkg = await load(source, { repair });
      assert.equal(await textOf(pkg), control, `${stem} ${repair}: the control's text, as Word's save of it reads`);
      assert.deepEqual(pkg.relationshipsNamespaces.map((w) => [w.code, w.partName, w.namespace, w.count]),
        [['package.relationships_namespace', partName, namespace, count]], `${stem} ${repair}`);
      assert.equal(pkg.warnings.filter((w) => w.code === 'package.relationships_namespace').length, 1, stem);
      assert.equal(pkg.repaired, 'word', stem);
      const result = await Texts.of(pkg);
      assert.deepEqual(result.warnings.filter((w) => w.kind === 'package.relationships_namespace').map((w) => [w.selector, w.detail]), [[partName, count]], stem);
      assert.equal(result.metadata.repaired, 'word', stem);
      // the save loses nothing and writes the right namespace, so it loads with nothing said
      const saved = await pkg.save();
      assert.deepEqual(Object.keys(unzipSync(saved)).sort(), Object.keys(unzipSync(source)).sort(), `${stem}: every part kept`);
      const rels = strFromU8(unzipSync(saved)[zipName(partName)]);
      assert.equal(relationshipsIn(rels), count, stem);
      assert.doesNotMatch(rels, new RegExp(WRONG.replace(/[.:]/g, '\\$&')), stem);
      const again = await load(saved);
      assert.deepEqual(again.relationshipsNamespaces, [], stem);
      assert.equal(again.repaired, 'none', stem);
      assert.equal(await textOf(again), control, `${stem}: read back`);
    }
  }
});

test('CR-031 section 2.3: the controls and a flat OPC package under repair word', async () => {
  for (const stem of ['rn-00-control-plain', 'rn-01-control-rich']) {
    const pkg = await load(bytesOf(`${stem}.docx`), { repair: 'word' });
    assert.deepEqual(pkg.relationshipsNamespaces, [], stem);
    assert.equal(pkg.repaired, 'none', stem);
  }
  const flat = await (await OpcPackage.load(bytesOf('rn-01-control-rich.docx'))).saveFlatOpc();
  const i = flat.indexOf('<Relationship ', flat.indexOf('pkg:name="/word/_rels/document.xml.rels"'));
  const moved = `${flat.slice(0, i)}<x:Relationship xmlns:x="${WRONG}" ${flat.slice(i + '<Relationship '.length)}`;
  const pkg = await OpcPackage.load(moved, { repair: 'word' });
  assert.deepEqual(pkg.relationshipsNamespaces.map((w) => [w.partName, w.namespace]), [['/word/_rels/document.xml.rels', WRONG]]);
  assert.equal(await textOf(await (async () => { await pkg.unmarshalAll(); return pkg; })()), await textOf(await load(bytesOf('rn-01-control-rich.docx'))));
});
