// CR-024: a package holding two entries of one name is refused (WD REQ-206), OPC part names compared without case, and
// read from the first entry of each name under LoadOptions.repair: 'word' (CR-031 D1, which folded duplicateEntries into it), as Word 365 reads it after its
// repair (check 75, three rounds; Word's saves in test/check75/ with a SHA-1 manifest). docx4j-python's CR-032 has the
// same cases. Each zip is built here by storedZip, since fflate's writers refuse a repeated name; the first copy holds
// FIRST and the second SECOND. Section 9's are at the end: more than one officeDocument relationship, REQ-206's statement of
// WD 0.100, refused, and read from the last relationship's target under the same switch (docx4j's package-wiring probes
// and Word 365's first-open saves of them in test/package-wiring/, with a SHA-1 manifest; python's CR-032 section 11).
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
  WordprocessingMLPackage, SpreadsheetMLPackage, OpcPackage, MemoryPartStore, ZipPartStore, DuplicateEntryException, MultipleMainPartsException,
  ExternalMainPartException, InvalidFormatException, Docx4JException, Texts, PlainText, TextOptions, TextOptionError, TEXT_OPTION_NAMES,
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
  + relationships.map(([id, type, target, mode]) => `<Relationship Id="${id}" Type="${type.includes('://') ? type : `${R}/${type}`}" Target="${target}"${mode ? ` TargetMode="${mode}"` : ''}/>`).join('') + '</Relationships>';
const packageRels = (target = 'word/document.xml') => rels(['rId1', 'officeDocument', target]);
const documentXml = (text, sectPr = '') => `${XML}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p>${sectPr}</w:body></w:document>`;
const zip = (...entries) => storedZip(entries.map(([name, value]) => [name, typeof value === 'string' ? strToU8(value) : value]));
/** A package's parts: [Content_Types].xml, the package relationships and the main part, then `rest`, in that order. */
const docx = (main, ...rest) => zip(['[Content_Types].xml', types()], ['_rels/.rels', packageRels()], ['word/document.xml', documentXml(main)], ...rest);

const text = async (pkg, options) => PlainText.of(await Texts.of(pkg, options)).trim();
const first = { repair: 'word' };

test('CR-024: the main part twice is refused, naming it, before anything is read', async () => {
  const bytes = docx('FIRST', ['word/document.xml', documentXml('SECOND')]);
  await assert.rejects(WordprocessingMLPackage.load(bytes), (e) => {
    assert.ok(e instanceof DuplicateEntryException && e instanceof InvalidFormatException);
    assert.equal(e.code, 'package.duplicate_entry');
    assert.deepEqual(e.entries, { 'word/document.xml': 2 });
    assert.equal(e.message, "Duplicate entry word/document.xml in this package: 2 entries (REQ-206). Load with { repair: 'word' } to read the first, as Word does (REQ-207)");
    return true;
  });
  // 'refuse' is the default, by name too
  await assert.rejects(WordprocessingMLPackage.load(bytes, { repair: 'refuse' }), DuplicateEntryException);
  // the store keeps the first, lists each name once and counts the rest
  const store = new ZipPartStore(bytes);
  assert.deepEqual([...store.partNames()], ['[Content_Types].xml', '_rels/.rels', 'word/document.xml']);
  assert.deepEqual([...store.duplicates()], [['word/document.xml', 2]]);
  assert.match(strFromU8(await store.load('word/document.xml')), /FIRST/);
});

test("CR-024: under repair: 'word' the first copy is read, one warning a name, and a save writes only that copy", async () => {
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

test('CR-024: the text model carries package.duplicate_entry, its detail the count; repair is the load\'s option, duplicate_entries not offered (CR-031 D1)', async () => {
  const pkg = await WordprocessingMLPackage.load(docx('FIRST', ['word/document.xml', documentXml('SECOND')]), first);
  const r = await Texts.of(pkg);
  const w = r.warnings.find((x) => x.kind === 'package.duplicate_entry');
  assert.equal(w.selector, 'word/document.xml');
  assert.equal(w.detail, 2);
  assert.deepEqual(w.toJSON(), { kind: 'package.duplicate_entry', address: undefined, selector: 'word/document.xml', detail: 2 });
  assert.equal(r.warnings.filter((x) => x.kind === 'package.duplicate_entry').length, 1);
  // a warning with no count leaves detail out of its JSON
  assert.ok(!('detail' in (await Texts.of(await WordprocessingMLPackage.load(docx('PLAIN'))).then((x) => x.warnings[0]?.toJSON() ?? {}))));
  // repair by name: its default, its values, its place before profile, and a value outside them refused
  assert.equal(new TextOptions().repair, 'refuse');
  assert.equal(TextOptions.of({ repair: 'WORD' }).repair, 'word');
  assert.equal(TextOptions.of({ repair: 'all' }).repair, 'all');
  assert.equal(new TextOptions().toMap().repair, 'refuse');
  assert.equal(TEXT_OPTION_NAMES.indexOf('warnings') + 1, TEXT_OPTION_NAMES.indexOf('repair'));
  assert.equal(TEXT_OPTION_NAMES.at(-1), 'profile');
  assert.throws(() => TextOptions.of({ repair: 'first' }), (e) => e instanceof TextOptionError && e.statement === 'REQ-207' && e.reason === 'invalid');
  // duplicate_entries, REQ-206's MAY, isn't offered: refused at any value, naming repair: 'word'
  for (const value of ['first', 'refuse']) {
    assert.throws(() => TextOptions.of({ duplicate_entries: value }), (e) => e instanceof TextOptionError && e.statement === 'REQ-206'
      && e.reason === 'not_implemented' && /repair: 'word'/.test(e.message));
  }
  // the load's option refuses a value outside its three
  for (const repair of ['first', 'WORD', 'none', true]) {
    await assert.rejects(WordprocessingMLPackage.load(docx('FIRST'), { repair }), (e) => e instanceof Docx4JException && /REQ-207/.test(e.message), String(repair));
  }
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

// --- section 9: more than one officeDocument relationship -----------------------------------------------------------

const STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument';
const PRESENTATION = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const override = (name, type = MAIN) => `<Override PartName="/${name}" ContentType="${type}"/>`;
const stem = (name) => name.split('/').pop().split('.')[0].toUpperCase();
const officeDocuments = (targets) => targets.map((target, i) => [`rId${i + 1}`, 'officeDocument', target]);
/** A package with an officeDocument relationship to each of `targets`, in that order, each part holding its own name's stem as its words. */
const mains = (targets, { relationships = officeDocuments(targets), overrides = '', rest = [] } = {}) => zip(
  ['[Content_Types].xml', types(targets.filter((t) => t !== 'word/document.xml').map((t) => override(t)).join('') + overrides)],
  ['_rels/.rels', rels(...relationships)],
  ...[...new Set(targets)].map((t) => [t, documentXml(stem(t))]),
  ...rest,
);
const TWO = ['word/document.xml', 'word/other.xml'];
const refusedNaming = (targets) => (e) => {
  assert.ok(e instanceof MultipleMainPartsException && e instanceof InvalidFormatException);
  assert.equal(e.code, 'package.multiple_main_parts');
  assert.deepEqual(e.targets, targets);
  return true;
};

test('CR-024 section 9: two officeDocument relationships are refused, naming the targets, in their order', async () => {
  await assert.rejects(WordprocessingMLPackage.load(mains(TWO)), (e) => {
    refusedNaming(['/word/document.xml', '/word/other.xml'])(e);
    assert.equal(e.message, "2 officeDocument relationships in this package: /word/document.xml, /word/other.xml (REQ-206). Load with { repair: 'word' } to read the last, as a fresh Word does (REQ-207)");
    return true;
  });
  await assert.rejects(WordprocessingMLPackage.load(mains([...TWO].reverse()), { repair: 'refuse' }), refusedNaming(['/word/other.xml', '/word/document.xml']));
  // what the statement counts: a transitional and a strict one are two, and so are two naming one target
  const relationships = [['rId1', 'officeDocument', TWO[0]], ['rId2', STRICT, TWO[1]]];
  await assert.rejects(WordprocessingMLPackage.load(mains(TWO, { relationships })), refusedNaming(['/word/document.xml', '/word/other.xml']));
  const same = mains(['word/document.xml', 'word/document.xml']);
  await assert.rejects(WordprocessingMLPackage.load(same), refusedNaming(['/word/document.xml', '/word/document.xml']));
  // every package type, the rule being the load's (measured on Word alone)
  const sheet = unzipSync(await (await SpreadsheetMLPackage.createPackage()).save());
  const workbook = strFromU8(sheet['_rels/.rels']).match(/Target="([^"]+)"[^>]*officeDocument|officeDocument"[^>]*Target="([^"]+)"/).slice(1).find(Boolean);
  const twice = strFromU8(sheet['_rels/.rels']).replace('</Relationships>', `<Relationship Id="rIdTwice" Type="${R}/officeDocument" Target="${workbook}"/></Relationships>`);
  await assert.rejects(OpcPackage.load(storedZip(Object.entries({ ...sheet, '_rels/.rels': strToU8(twice) }))), MultipleMainPartsException);
});

test("CR-024 section 9: under repair: 'word' the last relationship's target is read, whatever the entries or the Overrides", async () => {
  assert.equal(await text(await WordprocessingMLPackage.load(mains(TWO), first)), 'OTHER');
  assert.equal(await text(await WordprocessingMLPackage.load(mains([...TWO].reverse()), first)), 'DOCUMENT');
  // the relationships' order decides, not the entries'
  const reordered = zip(['[Content_Types].xml', types(override('word/other.xml'))], ['_rels/.rels', rels(...officeDocuments(TWO))],
    ['word/other.xml', documentXml('OTHER')], ['word/document.xml', documentXml('DOCUMENT')]);
  assert.equal(await text(await WordprocessingMLPackage.load(reordered, first)), 'OTHER');
  // three: the order decides and they're counted, the part named as the record has it
  const three = await WordprocessingMLPackage.load(mains(['word/document.xml', 'word/zeta.xml', 'word/beta.xml']), first);
  assert.equal(await text(three), 'BETA');
  assert.deepEqual(three.warnings, [{
    code: 'package.multiple_main_parts', partName: '/word/beta.xml', count: 3,
    message: "3 officeDocument relationships in this package (REQ-206); the last's target read, /word/beta.xml",
  }]);
  assert.equal(three.multipleMainParts, three.warnings[0]);
  // a strict one counts under the switch too
  const strict = await WordprocessingMLPackage.load(mains(TWO, { relationships: [['rId1', STRICT, TWO[0]], ['rId2', 'officeDocument', TWO[1]]] }), first);
  assert.equal(await text(strict), 'OTHER');
  assert.equal(strict.multipleMainParts.count, 2);
});

test('CR-031 D3 and D4: an External officeDocument relationship beside a part refused at refuse and word, read past at all; alone, refused at every value', async () => {
  const external = (id) => [id, 'officeDocument', 'http://example.com/a?b=1', 'External'];
  const part = (id, target = TWO[0]) => [id, 'officeDocument', target];
  // beside one part, either way round: refused at refuse and word, naming the URL and the part (python's fields)
  for (const relationships of [[part('rId1'), external('rId2')], [external('rId1'), part('rId2')]]) {
    for (const repair of [undefined, 'refuse', 'word']) {
      await assert.rejects(OpcPackage.load(mains([TWO[0]], { relationships }), { repair }), (e) => e instanceof ExternalMainPartException
        && e instanceof InvalidFormatException && e.code === 'package.external_main_part' && e.urls.join() === 'http://example.com/a?b=1'
        && e.targets.join() === '/word/document.xml' && /repair: 'all'/.test(e.message), String(repair));
    }
    // at all the part is read, the External relationship dropped and said, and a save writes none
    const pkg = await OpcPackage.load(mains([TWO[0]], { relationships }), { repair: 'all' });
    assert.ok(pkg instanceof WordprocessingMLPackage);
    assert.equal(pkg.getMainPart(), pkg.getMainDocumentPart());
    assert.equal(await text(pkg), 'DOCUMENT');
    assert.deepEqual(pkg.warnings.map((w) => [w.code, w.partName, w.urls]), [['package.external_main_part', '/word/document.xml', ['http://example.com/a?b=1']]]);
    assert.equal(pkg.externalMainPart, pkg.warnings[0]);
    assert.equal(pkg.repaired, 'own');
    const result = await Texts.of(pkg);
    assert.deepEqual(result.warnings.filter((w) => w.kind === 'package.external_main_part').map((w) => w.toJSON()),
      [{ kind: 'package.external_main_part', address: undefined, selector: '/word/document.xml', detail: 'http://example.com/a?b=1' }]);
    assert.equal(result.metadata.repaired, 'own');
    const saved = unzipSync(await pkg.save());
    assert.doesNotMatch(strFromU8(saved['_rels/.rels']), /External/);
    const again = await OpcPackage.load(storedZip(Object.entries(saved)));
    assert.deepEqual(again.warnings, []);
    assert.equal(again.repaired, 'none');
  }
  // beside two parts: refused for the External one first, as docx4j has it; at all both repairs, Word's last-relationship first
  const three = [part('rId1'), part('rId2', TWO[1]), external('rId3')];
  for (const repair of [undefined, 'word']) {
    await assert.rejects(WordprocessingMLPackage.load(mains(TWO, { relationships: three }), { repair }),
      (e) => e instanceof ExternalMainPartException && e.targets.join() === '/word/document.xml,/word/other.xml', String(repair));
  }
  const pkg = await WordprocessingMLPackage.load(mains(TWO, { relationships: three }), { repair: 'all' });
  assert.equal(await text(pkg), 'OTHER');
  assert.equal(pkg.multipleMainParts.count, 2);
  assert.deepEqual(pkg.warnings.map((w) => [w.code, w.partName]), [['package.multiple_main_parts', '/word/other.xml'], ['package.external_main_part', '/word/other.xml']]);
  // alone it names no main part: refused at every value (REQ-321), naming the URL, where it loaded as a bare package
  const plain = (id) => [id, 'officeDocument', 'http://example.com/a.docx', 'External'];
  for (const relationships of [[plain('rId1')], [plain('rId1'), plain('rId2')]]) {
    for (const repair of [undefined, 'word', 'all']) {
      await assert.rejects(OpcPackage.load(mains([], { relationships }), { repair }), (e) => e instanceof InvalidFormatException
        && !(e instanceof ExternalMainPartException) && /REQ-321/.test(e.message) && e.message.includes('http://example.com/a.docx'), String(repair));
    }
  }
});

test('CR-024 section 9: every route to the main part agrees, the part read chooses the package\'s class, and the other part is not loaded', async () => {
  const pkg = await OpcPackage.load(mains(TWO), first);
  assert.equal(pkg.getMainDocumentPart().partName.name, '/word/other.xml');
  assert.equal(pkg.getMainPart(), pkg.getMainDocumentPart());
  assert.equal(pkg.getPart('/word/document.xml'), undefined);
  assert.equal(pkg.relationshipsPart.getRelationshipsByType(`${R}/officeDocument`).length, 1);
  // the class comes from the part read: the first relationship's part is a presentation's
  const mixed = zip(['[Content_Types].xml', types(override('ppt/presentation.xml', PRESENTATION))],
    ['_rels/.rels', rels(['rId1', 'officeDocument', 'ppt/presentation.xml'], ['rId2', 'officeDocument', 'word/document.xml'])],
    ['ppt/presentation.xml', `${XML}<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`],
    ['word/document.xml', documentXml('WORDS')]);
  const read = await OpcPackage.load(mixed, first);
  assert.ok(read instanceof WordprocessingMLPackage);
  assert.equal(await text(read), 'WORDS');
});

test('CR-024 section 9: a save holds one main part, and what only the other reached goes with it; the save reloads with nothing said', async () => {
  const otherRels = rels(['rId1', 'header', 'header9.xml']);
  const header = `${XML}<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>HEADER</w:t></w:r></w:p></w:hdr>`;
  const bytes = mains([...TWO].reverse(), { rest: [['word/_rels/other.xml.rels', otherRels], ['word/header9.xml', header]] });
  const pkg = await WordprocessingMLPackage.load(bytes, first);
  const saved = await pkg.save();
  const entries = unzipSync(saved);
  assert.deepEqual(Object.keys(entries).sort(), ['[Content_Types].xml', '_rels/.rels', 'word/document.xml']);
  assert.equal(strFromU8(entries['_rels/.rels']).match(/officeDocument"/g).length, 1);
  assert.ok(!strFromU8(entries['[Content_Types].xml']).includes('/word/other.xml'));
  const again = await WordprocessingMLPackage.load(saved);
  assert.deepEqual(again.warnings, []);
  assert.equal(await text(again), 'DOCUMENT');
  // clone() goes through a save, so the copy holds the one main part and has nothing to say
  const copy = await pkg.clone();
  assert.equal(await text(copy), 'DOCUMENT');
  assert.equal(copy.multipleMainParts, undefined);
  // two relationships to one target: the part stays, and keeps its type
  const same = await WordprocessingMLPackage.load(mains(['word/document.xml', 'word/document.xml']), first);
  assert.equal(same.multipleMainParts.count, 2);
  assert.match(strFromU8(unzipSync(await same.save())['[Content_Types].xml']), /\/word\/document\.xml/);
  // a dropped target the kept part still reaches by a relationship of its own: it stays, and keeps its type
  const reached = mains(TWO, { rest: [['word/_rels/other.xml.rels', rels(['rId1', 'subDocument', 'document.xml'])]] });
  const held = await WordprocessingMLPackage.load(reached, first);
  assert.equal(held.getPart('/word/document.xml').contentType, MAIN);
  assert.match(strFromU8(unzipSync(await held.save())['[Content_Types].xml']), /\/word\/document\.xml/);
});

test('CR-024 section 9: the text model raises package.multiple_main_parts once, and a duplicate entry comes first; one relationship says nothing', async () => {
  const r = await Texts.of(await WordprocessingMLPackage.load(mains(TWO), first));
  assert.deepEqual(r.warnings.map((w) => w.toJSON()), [{ kind: 'package.multiple_main_parts', address: undefined, selector: '/word/other.xml', detail: 2 }]);
  // both at once: the entries are refused first, nothing having been read, and under the switch each is said, the entries first
  const both = mains(TWO, { rest: [['word/other.xml', documentXml('SECOND')]] });
  await assert.rejects(WordprocessingMLPackage.load(both), DuplicateEntryException);
  const pkg = await WordprocessingMLPackage.load(both, first);
  assert.deepEqual(pkg.warnings.map((w) => w.code), ['package.duplicate_entry', 'package.multiple_main_parts']);
  assert.deepEqual((await Texts.of(pkg)).warnings.map((w) => w.kind), ['package.duplicate_entry', 'package.multiple_main_parts']);
  assert.equal(await text(pkg), 'OTHER');
  // no case: one relationship, and another relationship type beside it, which the statement doesn't count
  const one = await WordprocessingMLPackage.load(mains([TWO[0]]));
  assert.deepEqual(one.warnings, []);
  assert.equal(one.multipleMainParts, undefined);
  const graphicFrame = [['rId1', 'officeDocument', TWO[0]], ['rId2', 'http://schemas.microsoft.com/office/2006/relationships/graphicFrameDoc', 'word/other.xml']];
  const beside = zip(['[Content_Types].xml', types(override('word/other.xml', 'application/xml'))], ['_rels/.rels', rels(...graphicFrame)],
    ['word/document.xml', documentXml('DOCUMENT')], ['word/other.xml', '<other/>']);
  assert.equal(await text(await WordprocessingMLPackage.load(beside)), 'DOCUMENT');
});

test('CR-024 section 9: flat OPC is refused and read as a zip is, and a fragment to insertOoxml is refused', async () => {
  const part = (name, ct, xml) => `<pkg:part pkg:name="${name}" pkg:contentType="${ct}"><pkg:xmlData>${xml.replace(XML, '')}</pkg:xmlData></pkg:part>`;
  const flat = '<?xml version="1.0" standalone="yes"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
    + part('/_rels/.rels', 'application/vnd.openxmlformats-package.relationships+xml', rels(...officeDocuments(TWO)))
    + part('/word/document.xml', MAIN, documentXml('DOCUMENT')) + part('/word/other.xml', MAIN, documentXml('OTHER')) + '</pkg:package>';
  await assert.rejects(WordprocessingMLPackage.load(flat), refusedNaming(['/word/document.xml', '/word/other.xml']));
  assert.equal(await text(await WordprocessingMLPackage.load(flat, first)), 'OTHER');
  const host = await WordprocessingMLPackage.createPackage();
  await assert.rejects((await host.getBody()).insertOoxml(flat, 'End'), MultipleMainPartsException);
});

test("CR-024 section 9: docx4j's package-wiring probes read as Word 365's first-open saves (test/package-wiring, SHA-1 manifest)", async () => {
  const dir = join(here, 'package-wiring');
  const manifest = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8'));
  for (const f of manifest.files) assert.equal(createHash('sha1').update(readFileSync(join(dir, f.name))).digest('hex'), f.sha1, f.name);
  const load = (name, options) => WordprocessingMLPackage.load(readFileSync(join(dir, name)), options);
  const cases = {
    'two-main-rels-document-first.docx': ['/word/document.xml', '/word/other.xml', 'MAIN_OTHER_XML'],
    'two-main-rels-other-first.docx': ['/word/other.xml', '/word/document.xml', 'MAIN_DOCUMENT_XML'],
    'two-main-rels-zip-other-first.docx': ['/word/document.xml', '/word/other.xml', 'MAIN_OTHER_XML'],
    'two-main-rels-types-other-first.docx': ['/word/document.xml', '/word/other.xml', 'MAIN_OTHER_XML'],
    'two-main-rels-names-swapped.docx': ['/word/zeta.xml', '/word/alpha.xml', 'MAIN_ALPHA_XML'],
  };
  for (const [name, [firstTarget, lastTarget, words]] of Object.entries(cases)) {
    await assert.rejects(load(name), refusedNaming([firstTarget, lastTarget]), name);
    const pkg = await load(name, first);
    assert.equal(await text(pkg), words, name);
    assert.equal(await text(await load(manifest.saves[name])), words, `${name}: Word's save`);
    assert.deepEqual(pkg.warnings.map((w) => [w.code, w.partName, w.count]), [['package.multiple_main_parts', lastTarget, 2]], name);
    const entries = Object.keys(unzipSync(await pkg.save()));
    assert.deepEqual(entries.filter((n) => n.startsWith('word/')), [lastTarget.substring(1)], name);
  }
  // a part no relationship names is no case, its Override last or not: the related part read, by either option, as Word reads it silently
  for (const [name, words] of [['one-main-rel-other-override-last.docx', 'MAIN_DOCUMENT_XML'], ['one-main-rel-document-override-last.docx', 'MAIN_OTHER_XML']]) {
    for (const options of [undefined, first]) {
      const pkg = await load(name, options);
      assert.equal(await text(pkg), words, name);
      assert.deepEqual(pkg.warnings, [], name);
    }
    assert.equal(await text(await load(manifest.saves[name])), words, `${name}: Word's save`);
  }
});
