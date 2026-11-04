// CR-024: a package holding two entries of one name is refused (WD REQ-206), OPC part names compared without case, and
// read from the first entry of each name under LoadOptions.duplicateEntries: 'first', as Word 365 reads it after its
// repair (check 75, three rounds; Word's saves in test/check75/ with a SHA-1 manifest). docx4j-python's CR-032 has the
// same cases. Each zip is built here by storedZip, since fflate's writers refuse a repeated name; the first copy holds
// FIRST and the second SECOND.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, strToU8, strFromU8 } from 'fflate';
import {
  WordprocessingMLPackage, OpcPackage, MemoryPartStore, ZipPartStore, DuplicateEntryException, InvalidFormatException,
  Docx4JException, Texts, PlainText, TextOptions, TextOptionError, TEXT_OPTION_NAMES,
} from '../dist/index.mjs';
import { DirectoryPartStore } from '../dist/node/index.mjs';
import { storedZip } from './helpers.mjs';
import { check75Files } from './fixtures/check75/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

const types = (overrides = '') => `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
  + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
  + '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
  + `<Override PartName="/word/document.xml" ContentType="${MAIN}"/>${overrides}</Types>`;
const rels = (...relationships) => `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
  + relationships.map(([id, type, target]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('') + '</Relationships>';
const packageRels = (target = 'word/document.xml') => rels(['rId1', 'officeDocument', target]);
const documentXml = (text, sectPr = '') => `${XML}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p>${sectPr}</w:body></w:document>`;
const zip = (...entries) => storedZip(entries.map(([name, value]) => [name, typeof value === 'string' ? strToU8(value) : value]));
/** A package's parts: [Content_Types].xml, the package relationships and the main part, then `rest`, in that order. */
const docx = (main, ...rest) => zip(['[Content_Types].xml', types()], ['_rels/.rels', packageRels()], ['word/document.xml', documentXml(main)], ...rest);

const text = async (pkg, options) => PlainText.of(await Texts.of(pkg, options)).trim();
const first = { duplicateEntries: 'first' };

test('CR-024: the main part twice is refused, naming it, before anything is read', async () => {
  const bytes = docx('FIRST', ['word/document.xml', documentXml('SECOND')]);
  await assert.rejects(WordprocessingMLPackage.load(bytes), (e) => {
    assert.ok(e instanceof DuplicateEntryException && e instanceof InvalidFormatException);
    assert.equal(e.code, 'package.duplicate_entry');
    assert.deepEqual(e.entries, { 'word/document.xml': 2 });
    assert.equal(e.message, "Duplicate entry word/document.xml in this package: 2 entries (REQ-206). Load with { duplicateEntries: 'first' } to read the first");
    return true;
  });
  // 'refuse' is the default, by name too
  await assert.rejects(WordprocessingMLPackage.load(bytes, { duplicateEntries: 'refuse' }), DuplicateEntryException);
  // the store keeps the first, lists each name once and counts the rest
  const store = new ZipPartStore(bytes);
  assert.deepEqual([...store.partNames()], ['[Content_Types].xml', '_rels/.rels', 'word/document.xml']);
  assert.deepEqual([...store.duplicates()], [['word/document.xml', 2]]);
  assert.match(strFromU8(await store.load('word/document.xml')), /FIRST/);
});

test("CR-024: under duplicateEntries: 'first' the first copy is read, one warning a name, and a save writes only that copy", async () => {
  const pkg = await WordprocessingMLPackage.load(docx('FIRST', ['word/document.xml', documentXml('SECOND')]), first);
  assert.equal(await text(pkg), 'FIRST');
  assert.deepEqual(pkg.duplicateEntries.map(({ code, entry, count }) => ({ code, entry, count })),
    [{ code: 'package.duplicate_entry', entry: 'word/document.xml', count: 2 }]);
  assert.equal(pkg.warnings[0].code, 'package.duplicate_entry');
  const saved = unzipSync(await pkg.save());
  assert.deepEqual(Object.keys(saved).filter((n) => /document\.xml$/i.test(n)), ['word/document.xml']);
  assert.match(strFromU8(saved['word/document.xml']), /FIRST/);
  // the save loads with no duplicate and no warning
  assert.deepEqual((await WordprocessingMLPackage.load(await pkg.save())).duplicateEntries, []);
  // three copies: count 3
  const three = await WordprocessingMLPackage.load(docx('FIRST', ['word/document.xml', documentXml('SECOND')], ['word/document.xml', documentXml('THIRD')]), first);
  assert.equal(three.duplicateEntries[0].count, 3);
  assert.equal(await text(three), 'FIRST');
});

test('CR-024: a case variant is the same name, in either order; the first copy is read under its own spelling (check 75a)', async () => {
  const decoySecond = docx('FIRST', ['Word/Document.xml', documentXml('DECOY')]);
  const decoyFirst = zip(['[Content_Types].xml', types()], ['_rels/.rels', packageRels()], ['Word/Document.xml', documentXml('DECOY')], ['word/document.xml', documentXml('SECOND')]);
  await assert.rejects(WordprocessingMLPackage.load(decoySecond), (e) => e instanceof DuplicateEntryException && e.entries['word/document.xml'] === 2);
  await assert.rejects(WordprocessingMLPackage.load(decoyFirst), (e) => e instanceof DuplicateEntryException && e.entries['Word/Document.xml'] === 2);
  assert.equal(await text(await WordprocessingMLPackage.load(decoySecond, first)), 'FIRST');
  // Word reads the first entry where the relationship names the second spelling (check 75a): so does 'first'
  const pkg = await WordprocessingMLPackage.load(decoyFirst, first);
  assert.equal(await text(pkg), 'DECOY');
  assert.equal(pkg.duplicateEntries[0].entry, 'Word/Document.xml');
});

test('CR-024: every entry is covered: [Content_Types].xml and the relationships parts, the first copy read', async () => {
  // [Content_Types].xml twice, the second making the main part a template's
  const ct = zip(['[Content_Types].xml', types()], ['[Content_Types].xml', types().replace('document.main+xml', 'template.main+xml')],
    ['_rels/.rels', packageRels()], ['word/document.xml', documentXml('FIRST')]);
  await assert.rejects(WordprocessingMLPackage.load(ct), (e) => e instanceof DuplicateEntryException && e.entries['[Content_Types].xml'] === 2);
  assert.equal((await WordprocessingMLPackage.load(ct, first)).getMainDocumentPart().contentType, MAIN);
  // _rels/.rels twice, the second naming another main part: the first's, and an untouched save keeps it
  const other = `<Override PartName="/word/other.xml" ContentType="${MAIN}"/>`;
  const packageTwice = zip(['[Content_Types].xml', types(other)], ['_rels/.rels', packageRels()], ['_rels/.rels', packageRels('word/other.xml')],
    ['word/document.xml', documentXml('FIRST')], ['word/other.xml', documentXml('OTHER')]);
  await assert.rejects(WordprocessingMLPackage.load(packageTwice), DuplicateEntryException);
  const pkg = await WordprocessingMLPackage.load(packageTwice, first);
  assert.equal(await text(pkg), 'FIRST');
  const saved = unzipSync(await pkg.save());
  assert.ok(saved['word/document.xml'] && !saved['word/other.xml']);
  // word/_rels/document.xml.rels twice, each naming its own header: the first's header
  const header = (t) => `${XML}<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:hdr>`;
  const headerTypes = '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    + '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';
  const partTwice = zip(['[Content_Types].xml', types(headerTypes)], ['_rels/.rels', packageRels()],
    ['word/document.xml', documentXml('BODY', '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>')],
    ['word/_rels/document.xml.rels', rels(['rIdH', 'header', 'header1.xml'])], ['word/_rels/document.xml.rels', rels(['rIdH', 'header', 'header2.xml'])],
    ['word/header1.xml', header('HEADER_FIRST')], ['word/header2.xml', header('HEADER_SECOND')]);
  await assert.rejects(WordprocessingMLPackage.load(partTwice), (e) => e instanceof DuplicateEntryException && e.entries['word/_rels/document.xml.rels'] === 2);
  assert.match(await text(await WordprocessingMLPackage.load(partTwice, first), { include: ['headers'] }), /HEADER_FIRST/);
});

test('CR-024: a media part twice and an orphan twice are refused; under the option the media part\'s first copy is saved and the orphan is not carried', async () => {
  const one = Uint8Array.from([1, 2, 3]);
  const two = Uint8Array.from([4, 5, 6, 7]);
  const media = zip(['[Content_Types].xml', types()], ['_rels/.rels', packageRels()], ['word/document.xml', documentXml('BODY')],
    ['word/_rels/document.xml.rels', rels(['rIdImg', 'image', 'media/image1.png'])], ['word/media/image1.png', one], ['word/media/image1.png', two],
    ['word/orphan.xml', documentXml('O1')], ['word/orphan.xml', documentXml('O2')]);
  await assert.rejects(WordprocessingMLPackage.load(media), (e) => {
    assert.deepEqual(e.entries, { 'word/media/image1.png': 2, 'word/orphan.xml': 2 });
    assert.match(e.message, /^Duplicate entries in this package: word\/media\/image1\.png, 2 entries; word\/orphan\.xml, 2 entries \(REQ-206\)/);
    return true;
  });
  const pkg = await WordprocessingMLPackage.load(media, first);
  assert.deepEqual(pkg.duplicateEntries.map((d) => d.entry), ['word/media/image1.png', 'word/orphan.xml']);
  const saved = unzipSync(await pkg.save());
  assert.deepEqual(saved['word/media/image1.png'], one);
  assert.equal(saved['word/orphan.xml'], undefined);
});

test('CR-024: a relationship whose target differs from its entry only in case finds its part, the names being one', async () => {
  const pkg = await WordprocessingMLPackage.load(zip(['[Content_Types].xml', types()], ['_rels/.rels', packageRels()], ['word/Document.xml', documentXml('FOUND')]));
  assert.equal(await text(pkg), 'FOUND');
  assert.deepEqual(pkg.duplicateEntries, []);
});

test('CR-024: the flat OPC, directory and memory stores refuse a duplicate and read the first under the option', async () => {
  // flat OPC: two pkg:part of one name
  const part = (name, ct, xml) => `<pkg:part pkg:name="${name}" pkg:contentType="${ct}"><pkg:xmlData>${xml.replace(XML, '')}</pkg:xmlData></pkg:part>`;
  const flat = '<?xml version="1.0" standalone="yes"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
    + part('/_rels/.rels', 'application/vnd.openxmlformats-package.relationships+xml', packageRels())
    + part('/word/document.xml', MAIN, documentXml('FIRST')) + part('/word/document.xml', MAIN, documentXml('SECOND')) + '</pkg:package>';
  await assert.rejects(WordprocessingMLPackage.load(flat), (e) => e instanceof DuplicateEntryException && e.entries['word/document.xml'] === 2);
  assert.equal(await text(await WordprocessingMLPackage.load(flat, first)), 'FIRST');
  // the same fragment to insertOoxml is refused, its own load's default
  const host = await WordprocessingMLPackage.createPackage();
  await assert.rejects((await host.getBody()).insertOoxml(flat, 'End'), DuplicateEntryException);
  // memory: the constructor's entries
  const entries = [['[Content_Types].xml', strToU8(types())], ['_rels/.rels', strToU8(packageRels())],
    ['word/document.xml', strToU8(documentXml('FIRST'))], ['Word/document.xml', strToU8(documentXml('SECOND'))]];
  await assert.rejects(OpcPackage.load(new MemoryPartStore(entries)), DuplicateEntryException);
  assert.equal(await text(await OpcPackage.load(new MemoryPartStore(entries), first)), 'FIRST');
  // a directory: Word/ beside word/ (a case-sensitive file system; elsewhere the two are one directory and nothing is held twice)
  const dir = await mkdtemp(join(tmpdir(), 'cr024-'));
  try {
    for (const [name, bytes] of entries) {
      await mkdir(dirname(join(dir, name)), { recursive: true });
      await writeFile(join(dir, name), bytes);
    }
    const store = await DirectoryPartStore.open(dir);
    if (store.duplicates().size > 0) {
      await assert.rejects(OpcPackage.load(store), DuplicateEntryException);
      assert.equal(store.duplicates().get([...store.duplicates().keys()][0]), 2);
    }
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('CR-024: the text model carries package.duplicate_entry, its detail the count; duplicate_entries is the WD\'s option name', async () => {
  const pkg = await WordprocessingMLPackage.load(docx('FIRST', ['word/document.xml', documentXml('SECOND')]), first);
  const r = await Texts.of(pkg);
  const w = r.warnings.find((x) => x.kind === 'package.duplicate_entry');
  assert.equal(w.selector, 'word/document.xml');
  assert.equal(w.detail, 2);
  assert.deepEqual(w.toJSON(), { kind: 'package.duplicate_entry', address: undefined, selector: 'word/document.xml', detail: 2 });
  assert.equal(r.warnings.filter((x) => x.kind === 'package.duplicate_entry').length, 1);
  // a warning with no count leaves detail out of its JSON
  assert.ok(!('detail' in (await Texts.of(await WordprocessingMLPackage.load(docx('PLAIN'))).then((x) => x.warnings[0]?.toJSON() ?? {}))));
  // the option by name: its default, its values, its place before profile, and a value outside them refused
  assert.equal(new TextOptions().duplicateEntries, 'refuse');
  assert.equal(TextOptions.of({ duplicate_entries: 'FIRST' }).duplicateEntries, 'first');
  assert.equal(new TextOptions().toMap().duplicate_entries, 'refuse');
  assert.equal(TEXT_OPTION_NAMES.indexOf('warnings') + 1, TEXT_OPTION_NAMES.indexOf('duplicate_entries'));
  assert.equal(TEXT_OPTION_NAMES.at(-1), 'profile');
  assert.throws(() => TextOptions.of({ duplicate_entries: 'last' }), (e) => e instanceof TextOptionError && e.statement === 'REQ-206');
  // the load's option refuses a value outside its two
  await assert.rejects(WordprocessingMLPackage.load(docx('FIRST'), { duplicateEntries: 'last' }), (e) => e instanceof Docx4JException && /REQ-206/.test(e.message));
});

test("CR-024: Word check 75's files and Word 365's saves (test/check75, SHA-1 manifest)", async () => {
  const dir = join(here, 'check75');
  const manifest = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8'));
  for (const f of manifest.files) assert.equal(createHash('sha1').update(readFileSync(join(dir, f.name))).digest('hex'), f.sha1, f.name);
  const files = new Map((await check75Files()).map((f) => [f.name, f.bytes]));
  // 75a: an ASCII case pair, refused here; Word reads the first entry, HEADER 1, where the relationship names the second
  const a = files.get('75a-ascii-pair');
  await assert.rejects(WordprocessingMLPackage.load(a), (e) => e instanceof DuplicateEntryException && e.entries['word/headerA.xml'] === 2);
  const headerOf = async (pkg) => (await text(pkg, { include: ['headers'] })).split('\n').find((l) => l.startsWith('HEADER'));
  const ours = await headerOf(await WordprocessingMLPackage.load(a, first));
  const words = await headerOf(await WordprocessingMLPackage.load(readFileSync(join(dir, '75a-ascii-pair-word365.docx'))));
  assert.match(ours, /^HEADER 1: the first entry/);
  assert.equal(ours, words);
  // 75j: the writer's control, one ASCII header, opened cleanly by Word, loads here with no duplicate
  assert.deepEqual((await WordprocessingMLPackage.load(files.get('75j-writer-ascii-only'))).duplicateEntries, []);
  // the non-ASCII files hold no case pair under the key (raw É and é lower-case to one; percent-encoded, %C3%89 and %C3%A9
  // stay two, %c3%a9 and %C3%A9 are one), and Word's saves of them hold no header: Word drops a non-ASCII part name
  await assert.rejects(WordprocessingMLPackage.load(files.get('75c-nonascii-pair')), DuplicateEntryException);
  assert.deepEqual((await WordprocessingMLPackage.load(files.get('75g-encoded-pair'))).duplicateEntries, []);
  await assert.rejects(WordprocessingMLPackage.load(files.get('75h-hex-case-pair')), DuplicateEntryException);
  for (const name of ['75b-nonascii-single', '75c-nonascii-pair', '75f-encoded-single', '75g-encoded-pair', '75k-fflate-encoded-single']) {
    const saved = unzipSync(readFileSync(join(dir, `${name}-word365.docx`)));
    assert.ok(!Object.keys(saved).some((n) => /header/i.test(n)), name);
  }
});
