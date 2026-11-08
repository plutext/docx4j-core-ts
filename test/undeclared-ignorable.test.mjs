// CR-030: a prefix in mc:Ignorable with no declaration in scope names nothing, is read as though the list didn't hold it, and is
// said once a part (mce.undeclared_prefix, WD REQ-323). The portfolio's round is in test/undeclared-ignorable/, copied unedited with
// a SHA-1 manifest: each file held to Word's text by both load routes, the warning to where a prefix has no declaration, and the
// saves to REQ-323's outcome, a part written anew holding no such prefix and one copied keeping it.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { OpcPackage, Texts, PlainText, parseXml, preprocessAlternateContent, SaveLossException } from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, 'undeclared-ignorable');
const manifest = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8'));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const ROUTES = [{}, { mcePreprocess: true }];
const PARTS = ['word/document.xml', 'word/header1.xml', 'word/footnotes.xml', 'word/styles.xml', 'word/settings.xml'];

// Each file's parts whose mc:Ignorable lists a prefix with no declaration in scope, read from the files by hand.
const UNDECLARED = {
  'ui-00-control': {},
  'ui-01-main-unknown': { '/word/document.xml': 'zz9' },
  'ui-02-main-known': { '/word/document.xml': 'w14' },
  'ui-03-main-several': { '/word/document.xml': 'w15 wp14' },
  'ui-04-main-lower-known': { '/word/document.xml': 'w14' },
  'ui-05-main-lower-unknown': { '/word/document.xml': 'zz9' },
  'ui-06-main-declared-unknown': {},
  'ui-07-main-leading-space': {},
  'ui-10-header': { '/word/header1.xml': 'zz9' },
  'ui-11-footnotes': { '/word/footnotes.xml': 'zz9' },
  'ui-12-styles': { '/word/styles.xml': 'zz9' },
  'ui-13-settings': { '/word/settings.xml': 'zz9' },
  'ui-14-all-parts-known': Object.fromEntries(PARTS.map((p) => [`/${p}`, 'w14'])),
};

const bytesOf = (name) => new Uint8Array(readFileSync(join(dir, name)));
const norm = (text) => text.replace(/\s+/g, ' ').trim();
const load = async (bytes, options = {}) => {
  const pkg = await OpcPackage.load(bytes, options);
  await pkg.unmarshalAll();
  return pkg;
};
const read = async (pkg) => Texts.of(pkg, { include: ['headers', 'footers'], comments: true });
const byPart = (warnings) => Object.fromEntries(warnings.filter((w) => w.code === 'mce.undeclared_prefix').map((w) => [w.partName, w.prefixes.join(' ')]));
const rootOf = (xml) => {
  const doc = parseXml(xml);
  return doc.documentElement;
};
const undeclaredOn = (root) => (root.getAttributeNS(MC, 'Ignorable') ?? '').trim().split(/\s+/).filter((p) => p !== '' && !root.lookupNamespaceURI(p));

test('CR-030: the files are as the manifest has them', () => {
  for (const f of manifest.files) assert.equal(createHash('sha1').update(bytesOf(f.name)).digest('hex'), f.sha1, f.name);
  assert.deepEqual(Object.keys(manifest.stems).sort(), Object.keys(UNDECLARED).sort());
});

test('CR-030: mce.undeclared_prefix once a part, with its prefixes, by both load routes, on the part and in the text model', async () => {
  for (const [stem, expected] of Object.entries(UNDECLARED)) {
    for (const options of ROUTES) {
      const pkg = await load(bytesOf(`${stem}.docx`), options);
      assert.deepEqual(byPart(pkg.warnings), expected, `${stem} ${JSON.stringify(options)}`);
      const result = await read(pkg);
      const said = Object.fromEntries(result.warnings.filter((w) => w.kind === 'mce.undeclared_prefix').map((w) => [w.selector, w.detail]));
      assert.deepEqual(said, expected, `${stem}: the text model`);
    }
  }
});

test('CR-030: each file Word opened reads with Word\'s text, by both load routes', async () => {
  for (const [stem, { saved }] of Object.entries(manifest.stems)) {
    if (!saved) continue;
    const word = norm(PlainText.of(await read(await load(bytesOf(`${stem}-word365.docx`)))));
    for (const options of ROUTES) {
      assert.equal(norm(PlainText.of(await read(await load(bytesOf(`${stem}.docx`), options)))), word, `${stem} ${JSON.stringify(options)}`);
    }
  }
});

test('CR-030: a prefix declared only below the list names nothing there, so its element isn\'t ignorable (ui-05), and one declared on the root is (ui-06)', async () => {
  const flags = async (stem) => (await load(bytesOf(`${stem}.docx`))).warnings.filter((w) => w.code === 'part.skipped').map((w) => [w.name, w.ignorable]);
  assert.deepEqual(await flags('ui-05-main-lower-unknown'), [['zz9:x', false]]);
  assert.deepEqual(await flags('ui-06-main-declared-unknown'), [['zz9:x', true]]);
});

test('CR-030: a part copied from its source keeps the prefix, and a part written anew holds no undeclared one', async () => {
  for (const stem of Object.keys(UNDECLARED)) {
    const source = bytesOf(`${stem}.docx`);
    const from = unzipSync(source);
    const untouched = unzipSync(await (await OpcPackage.load(source)).save());
    for (const p of PARTS) assert.equal(strFromU8(untouched[p]), strFromU8(from[p]), `${stem} ${p}: copied`);
    const pkg = await load(source);
    const saved = unzipSync(await pkg.save());
    // ui-05 and ui-06 hold an element the read dropped, so their main part, read and not changed, is copied (CR-025 section 2.2)
    const skipped = new Set(pkg.warnings.filter((w) => w.code === 'part.skipped').map((w) => w.partName));
    assert.deepEqual([...skipped], ['ui-05-main-lower-unknown', 'ui-06-main-declared-unknown'].includes(stem) ? ['/word/document.xml'] : [], stem);
    for (const p of PARTS) {
      if (skipped.has(`/${p}`)) assert.equal(strFromU8(saved[p]), strFromU8(from[p]), `${stem} ${p}: copied`);
      else assert.deepEqual(undeclaredOn(rootOf(strFromU8(saved[p]))), [], `${stem} ${p}: written anew`);
    }
  }
});

test('CR-030: ui-05 and ui-06 changed: ui-05 (Word refuses it) CR-025\'s gate refuses, ui-06\'s ignorable element Word\'s save drops too; either way written anew with no undeclared prefix', async () => {
  for (const stem of ['ui-05-main-lower-unknown', 'ui-06-main-declared-unknown']) {
    const pkg = await load(bytesOf(`${stem}.docx`));
    pkg.getMainDocumentPart().body.insertParagraph('EDIT', 'End');
    if (stem.startsWith('ui-05')) {
      await assert.rejects(pkg.save(), (e) => e instanceof SaveLossException && e.losses[0].partName === '/word/document.xml', stem);
      pkg.checkLossOnSave = false;
    }
    const saved = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
    assert.match(saved, /EDIT/, stem);
    assert.deepEqual(undeclaredOn(rootOf(saved)), [], stem);
  }
});

test('CR-030: an unknown prefix goes from a part written anew, and a known one is declared', async () => {
  const ignorable = async (stem) => rootOf(strFromU8(unzipSync(await (await load(bytesOf(`${stem}.docx`))).save())['word/document.xml']));
  assert.equal((await ignorable('ui-01-main-unknown')).getAttributeNS(MC, 'Ignorable') || '', '');
  const known = await ignorable('ui-03-main-several');
  assert.equal(known.getAttributeNS(MC, 'Ignorable'), 'w14 w15 wp14');
  assert.equal(known.lookupNamespaceURI('w15'), 'http://schemas.microsoft.com/office/word/2012/wordml');
});

test('CR-030: the list is read where it stands: below the root, with a leading space, and a declaration on a descendant out of scope', () => {
  const warned = (xml) => preprocessAlternateContent(parseXml(xml), { resolve: false, partName: '/p' }).warnings
    .filter((w) => w.code === 'mce.undeclared_prefix').map((w) => w.prefixes);
  assert.deepEqual(warned(`<w:document xmlns:w="${W}" xmlns:mc="${MC}"><w:body><w:p mc:Ignorable="a b"/><w:p mc:Ignorable="b c"/></w:body></w:document>`), [['a', 'b', 'c']]);
  assert.deepEqual(warned(`<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:a="urn:a" mc:Ignorable=" a"/>`), []);
  assert.deepEqual(warned(`<w:document xmlns:w="${W}" xmlns:mc="${MC}" mc:Ignorable="a"><w:body xmlns:a="urn:a"/></w:document>`), [['a']]);
  assert.deepEqual(warned(`<w:document xmlns:w="${W}" xmlns:mc="${MC}"/>`), []);
});

test('CR-030: a list inside a branch is read where it stands, so a resolving load says nothing of a prefix its mc:Choice declares', () => {
  const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const xml = (listed) => `<w:document xmlns:w="${W}" xmlns:mc="${MC}"><w:body><mc:AlternateContent><mc:Choice Requires="w14" xmlns:w14="${W14}">`
    + `<w:p mc:Ignorable="${listed}"/></mc:Choice></mc:AlternateContent></w:body></w:document>`;
  for (const resolve of [false, true]) {
    const warned = (listed) => preprocessAlternateContent(parseXml(xml(listed)), { resolve, partName: '/p' }).warnings
      .filter((w) => w.code === 'mce.undeclared_prefix').map((w) => w.prefixes);
    assert.deepEqual(warned('w14'), [], `resolve ${resolve}`);
    assert.deepEqual(warned('w14 zz'), [['zz']], `resolve ${resolve}`);
  }
});

test('CR-030: in nested branches a list is read where it stood in the source, not again once an outer branch has moved', () => {
  const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
  const xml = `<w:document xmlns:w="${W}" xmlns:mc="${MC}"><w:body><mc:AlternateContent><mc:Choice Requires="w14" xmlns:w14="${W14}">`
    + `<mc:AlternateContent><mc:Choice Requires="w15" xmlns:w15="${W15}"><w:p mc:Ignorable="w14 w15"/></mc:Choice></mc:AlternateContent>`
    + `</mc:Choice></mc:AlternateContent></w:body></w:document>`;
  const result = preprocessAlternateContent(parseXml(xml), { partName: '/p' });
  assert.equal(result.resolved, 2);
  assert.deepEqual(result.warnings.filter((w) => w.code === 'mce.undeclared_prefix'), []);
});
