// The Node-only container (CR-001 Phase C: docx4j's UnzippedPartStore) and OpcPackage.clone().
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm, readFile, writeFile, mkdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { OpcPackage, WordprocessingMLPackage, SpreadsheetMLPackage, ZipPartStore } from '../dist/index.mjs';
import { DirectoryPartStore, DirectoryPartSink } from '../dist/node/index.mjs';
import { fixture, bytesEqual, plain } from './helpers.mjs';

async function temp(name) {
  return mkdtemp(join(tmpdir(), `core-ts-${name}-`));
}

test('DirectoryPartSink then DirectoryPartStore: zip to directory to zip reproduces every part', async () => {
  const dir = await temp('dir');
  try {
    const bytes = await fixture('loadAndSave.docx');
    const pkg = await OpcPackage.load(bytes);
    assert.equal(await pkg.saveTo(new DirectoryPartSink(dir)), dir);
    // the parts are files under the directory, [Content_Types].xml among them
    assert.ok((await readFile(join(dir, '[Content_Types].xml'), 'utf8')).startsWith('<?xml'));
    assert.ok((await readFile(join(dir, 'word', 'document.xml'))).length > 0);

    const store = await DirectoryPartStore.open(dir);
    const names = [...store.partNames()];
    assert.equal(names[0], '[Content_Types].xml', 'listed first, as a zip written by Word has it');
    assert.ok(store.has('word/document.xml') && store.has('/word/document.xml'));
    assert.equal(store.has('word/nothing.xml'), false);
    const source = new ZipPartStore(bytes);
    assert.equal(store.size('word/document.xml'), source.loadSync('word/document.xml').length);
    await assert.rejects(() => store.load('word/nothing.xml'), /No part/);

    const back = await WordprocessingMLPackage.load(store);
    assert.equal(back.parts.size, pkg.parts.size);
    const out = new ZipPartStore(await back.save());
    for (const name of source.partNames()) {
      if (name === '[Content_Types].xml' || name.endsWith('.rels')) continue;  // both are always regenerated
      assert.ok(bytesEqual(source.loadSync(name), out.loadSync(name)), `${name} byte-identical`);
    }
    // the relationships parts are re-marshalled by any save, zip to zip included, so they are
    // compared as trees
    assert.deepEqual(plain(back.getMainDocumentPart().relationshipsPart.list), plain(pkg.getMainDocumentPart().relationshipsPart.list));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('DirectoryPartStore: a package edited through the directory saves back to it', async () => {
  const dir = await temp('edit');
  try {
    await (await WordprocessingMLPackage.createPackage()).saveTo(new DirectoryPartSink(dir));
    const pkg = await WordprocessingMLPackage.load(await DirectoryPartStore.open(dir));
    (await pkg.getBody()).insertParagraph('Edited in place', 'End');
    await pkg.saveTo(new DirectoryPartSink(dir));
    const again = await WordprocessingMLPackage.load(await DirectoryPartStore.open(dir));
    assert.equal((await again.getBody()).text.trim(), 'Edited in place');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('DirectoryPartStore.open: a missing directory is a Docx4JException', async () => {
  await assert.rejects(() => DirectoryPartStore.open(join(tmpdir(), 'core-ts-no-such-directory')), /Cannot read package directory/);
});

test('DirectoryPartSink: a part name that climbs out of the directory is refused', async () => {
  const dir = await temp('slip');
  try {
    const sink = new DirectoryPartSink(dir);
    sink.put('../escaped.xml', new Uint8Array([1]));
    await assert.rejects(() => sink.finish(), /outside the target directory/);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('DirectoryPartStore: a directory that is not a package fails to load', async () => {
  const dir = await temp('empty');
  try {
    await mkdir(join(dir, 'word'), { recursive: true });
    await writeFile(join(dir, 'word', 'document.xml'), '<x/>');
    const store = await DirectoryPartStore.open(dir);
    await assert.rejects(() => OpcPackage.load(store));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('clone: an independent package of the same class', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('loadAndSave.docx'));
  pkg.author = { name: 'Jane Doe', initials: 'JD' };
  const body = await pkg.getBody();
  const textBefore = body.text;
  const copy = await pkg.clone();
  assert.ok(copy instanceof WordprocessingMLPackage);
  assert.notEqual(copy, pkg);
  assert.equal(copy.parts.size, pkg.parts.size);
  assert.equal(copy.author.name, 'Jane Doe', 'the settings that are not in the package come too');

  // editing the clone leaves the original alone, and the other way round
  (await copy.getBody()).insertParagraph('Only in the clone', 'End');
  assert.equal(body.text, textBefore);
  assert.ok((await copy.getBody()).text.includes('Only in the clone'));
  body.insertParagraph('Only in the original', 'End');
  assert.ok(!(await copy.getBody()).text.includes('Only in the original'));

  // a part nobody unmarshalled is still the bytes it was loaded with
  const source = new ZipPartStore(await fixture('loadAndSave.docx'));
  const out = new ZipPartStore(await copy.save());
  assert.ok(bytesEqual(source.loadSync('word/styles.xml'), out.loadSync('word/styles.xml')));
  assert.equal(copy.getPart('/word/styles.xml').isUnmarshalled, false);
});

test('clone: the subclasses keep their class and their shortcuts', async () => {
  const wml = await WordprocessingMLPackage.createPackage();
  wml.body.insertParagraph('Hello', 'End');
  const wmlCopy = await wml.clone();
  assert.ok(wmlCopy instanceof WordprocessingMLPackage);
  assert.equal((await wmlCopy.getBody()).text.trim(), 'Hello');
  assert.ok(wmlCopy.getMainDocumentPart().styleDefinitionsPart);
  assert.equal(wmlCopy.fonts.defaultTheme, wml.fonts.defaultTheme);

  const xlsx = await SpreadsheetMLPackage.load(await fixture('loadAndSave.xlsx'));
  const xlsxCopy = await xlsx.clone();
  assert.ok(xlsxCopy instanceof SpreadsheetMLPackage);
  assert.equal((await xlsxCopy.getWorksheetParts()).length, 1);

  const generic = await OpcPackage.load(await fixture('loadAndSave.pptx'));
  const genericCopy = await generic.clone();
  assert.equal(genericCopy.constructor, generic.constructor);
  assert.equal(genericCopy.parts.size, generic.parts.size);
});
