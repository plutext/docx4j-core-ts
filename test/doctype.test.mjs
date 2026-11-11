// CR-033: a DOCTYPE read by this package, never by the platform's parser (the portfolio's docs/dtd_handling_policy.md; Jason,
// 2026-11-11: "refuse where the open xml requires refusal; for other parts, guard against malicious entity expansion and other
// attacks", then decisions 1 to 4). In a part ECMA-376 Part 2 defines (the content types, a relationships part, the core
// properties, a signature part) a DOCTYPE refuses the package, or under repair 'word' is read as Word 365 repairs it; elsewhere
// the part is read with its internal entities expanded and its attribute defaults applied within limits, nothing outside the
// package fetched, and `xml.doctype` says so with Word's reading. A save after an edit writes every such part without its DOCTYPE.
// Word's two rounds, dtd-word-check and dtd2-word-check (docx4j-python's measurement through the portfolio's bridge, 2026-11-10),
// are in test/doctype/, copied unedited with a SHA-1 manifest.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { OpcPackage, Texts, PlainText, parseXml, DoctypeException, PartUnmarshalException, readDoctype, DOCTYPE_LIMITS } from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, 'doctype');
const manifest = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8'));
const quiet = () => { const e = console.error; const w = console.warn; console.error = () => {}; console.warn = () => {}; return () => { console.error = e; console.warn = w; }; };
const bytesOf = (name) => new Uint8Array(readFileSync(join(dir, name)));
const entriesOf = (bytes) => unzipSync(bytes);
const lines = (t) => t.replace(/^﻿/, '').split(/\r\n|\r|\n/).map((l) => l.trim()).filter((l) => l !== '');
const cases = readdirSync(dir).filter((n) => n.endsWith('.docx') && !n.endsWith('.word.docx')).map((n) => n.replace(/\.docx$/, '')).sort();
const resultOf = (n) => JSON.parse(readFileSync(join(dir, `${n}.result.json`), 'utf8'));

async function load(name, repair) {
  const restore = quiet();
  try {
    return await OpcPackage.load(bytesOf(`${name}.docx`), repair === undefined ? {} : { repair });
  } finally { restore(); }
}

async function textOf(pkg) {
  const restore = quiet();
  try { return PlainText.of(await Texts.of(pkg)); } finally { restore(); }
}

/** A copy of a case with one entry's text changed. */
function withEntry(name, entry, change) {
  const entries = entriesOf(bytesOf(`${name}.docx`));
  entries[entry] = strToU8(change(strFromU8(entries[entry])));
  return zipSync(entries);
}

// Where each case's DOCTYPE is, as the rounds built them, and whether Part 2 forbids one there.
const FORBIDDEN = {
  'dtd-14-core-bare': '/docProps/core.xml',
  'dtd-16-content-types-bare': '/[Content_Types].xml',
  'dtd-17-package-rels-bare': '/_rels/.rels',
  'dtd-18-document-rels-bare': '/word/_rels/document.xml.rels',
  'dtd2-91-signature': '/_xmlsignatures/sig1.xml',
};

test('CR-033: the files are as the manifest has them', () => {
  for (const [name, sha1] of Object.entries(manifest.files)) assert.equal(createHash('sha1').update(bytesOf(name)).digest('hex'), sha1, name);
  assert.equal(Object.keys(manifest.files).length, 80);
});

test('a part Part 2 defines: the package refused at the default; under repair \'word\' the content types and _rels/.rels still, the rest read as Word repairs them', async () => {
  for (const [name, partName] of Object.entries(FORBIDDEN)) {
    await assert.rejects(load(name), (e) => e instanceof DoctypeException && e instanceof PartUnmarshalException && e.partName === partName && e.code === 'package.doctype' && /Part 2 §6\.2\.5/.test(e.message), name);
  }
  for (const name of ['dtd-16-content-types-bare', 'dtd-17-package-rels-bare']) {
    for (const repair of ['word', 'all']) await assert.rejects(load(name, repair), DoctypeException, `${name} ${repair}`);
    // Word refuses both after its repair prompt
    assert.equal(resultOf(name).opened, 'refused');
  }
  for (const repair of ['word', 'all']) {
    // the core properties read as absent, as Word reads them (no title)
    const core = await load('dtd-14-core-bare', repair);
    assert.equal(core.docPropsCorePart, undefined);
    assert.equal(core.repaired, 'word');
    assert.deepEqual(core.warnings.filter((w) => w.code === 'xml.doctype').map((w) => [w.partName, w.repaired, w.word]), [['/docProps/core.xml', true, 'reads the part as absent']]);
    const saved = entriesOf(await core.save());
    assert.equal(saved['docProps/core.xml'], undefined);
    assert.doesNotMatch(strFromU8(saved['[Content_Types].xml']), /core\.xml/);
    // a signature part read as it is, as Word opens such a file
    const sig = await load('dtd2-91-signature', repair);
    assert.equal(sig.repaired, 'none');
    assert.deepEqual(sig.warnings.filter((w) => w.code === 'xml.doctype').map((w) => [w.partName, w.repaired, w.word]), [['/_xmlsignatures/sig1.xml', false, 'opens it']]);
    assert.equal(resultOf('dtd2-91-signature').opened, 'silent');
  }
});

test('a document relationships part under repair \'word\': every relationship lost, as Word\'s repair loses them, its text Word\'s and its save Word\'s parts', async () => {
  const pkg = await load('dtd-18-document-rels-bare', 'word');
  assert.equal(pkg.repaired, 'word');
  assert.deepEqual(pkg.relationshipsNotRead.map((w) => [w.partName, w.count]), [['/word/_rels/document.xml.rels', 0]]);
  assert.deepEqual(pkg.warnings.filter((w) => w.code === 'xml.doctype').map((w) => [w.partName, w.repaired, w.word]), [['/word/_rels/document.xml.rels', true, 'loses every relationship']]);
  assert.deepEqual(lines(await textOf(pkg)), lines(readFileSync(join(dir, 'dtd-18-document-rels-bare.txt'), 'latin1')));
  // the save writes the repaired form, as CR-032's does under repair 'word': without the parts only the lost relationships named
  const saved = Object.keys(entriesOf(await pkg.save()));
  for (const gone of ['word/styles.xml', 'word/settings.xml', 'word/header1.xml', 'customXml/item1.xml']) assert.ok(!saved.includes(gone), gone);
  assert.ok(!Object.keys(entriesOf(bytesOf('dtd-18-document-rels-bare.word.docx'))).includes('word/header1.xml'));
});

test('every other part: read at every value of repair with xml.doctype naming it, and Word\'s reading of the file said, as the rounds measured it', async () => {
  const expected = { 'refused': 'refuses the file', 'repaired': 'reads the part as absent' };
  let checked = 0;
  for (const name of cases) {
    if (name in FORBIDDEN || /control/.test(name)) continue;
    for (const repair of ['refuse', 'word']) {
      const pkg = await load(name, repair);
      await textOf(pkg);
      const restore = quiet();
      try { await pkg.customXmlParts.load(); } finally { restore(); }
      const doctypes = pkg.warnings.filter((w) => w.code === 'xml.doctype');
      // a glossary and an embedded package aren't read by the text model, so their DOCTYPE isn't met
      if (/glossary|embedded/.test(name)) { assert.deepEqual(doctypes, [], name); continue; }
      assert.equal(doctypes.length, 1, `${name}: ${doctypes.map((w) => w.partName)}`);
      assert.equal(doctypes[0].repaired, false);
      assert.equal(doctypes[0].word, expected[resultOf(name).opened], `${name}: Word ${resultOf(name).opened}`);
      assert.equal(pkg.repaired, 'none');
      checked++;
    }
  }
  assert.equal(checked, 2 * 17);
});

test('the internal subset as XML 1.0 has a non-validating processor read it: an entity expanded (dtd-06, dtd-11), a default applied (dtd-04), the external subset not read (dtd-07)', async () => {
  const six = await load('dtd-06-document-entity-used');
  assert.match(await textOf(six), /^DTD_BEGIN ENTITY_EXPANDED$/m);
  assert.deepEqual(six.warnings.filter((w) => w.code === 'xml.doctype').map((w) => [w.holds, w.expanded]), [[['entities'], 1]]);
  // a save after unmarshalling writes the expansion, never the reference escaped (`&amp;bp;`, before CR-033), and no DOCTYPE
  await six.unmarshalAll();
  const saved = strFromU8(entriesOf(await six.save())['word/document.xml']);
  assert.match(saved, /DTD_BEGIN ENTITY_EXPANDED/);
  assert.doesNotMatch(saved, /DOCTYPE|&amp;bp;/);

  const four = await load('dtd-04-document-attribute-default');
  const doc = await four.getMainDocumentPart().getContents();
  const runs = JSON.stringify(doc.body.content);
  assert.match(runs, /"b":\{[^}]*"val":false/);
  assert.equal(four.warnings.find((w) => w.code === 'xml.doctype').defaulted, 1);

  const eleven = await load('dtd-11-styles-entity-used');
  const styles = await eleven.getMainDocumentPart().styleDefinitionsPart.getContents();
  assert.equal(styles.style.find((s) => s.styleId === 'StyleOne').pPr.jc.val, 'center');

  const seven = await load('dtd-07-document-external-dtd');
  await textOf(seven);
  const w = seven.warnings.find((x) => x.code === 'xml.doctype');
  assert.deepEqual(w.holds, ['external']);
  assert.deepEqual(w.notRead, ['SYSTEM "http://example.invalid/wordprocessingml.dtd"']);
});

// ------------------------------------------------------------------------------------------------ the safeguards

const R = (subset, body) => `<?xml version="1.0"?>\n<!DOCTYPE r [${subset}]>\n<r>${body}</r>`;
const malformedBecause = (pattern) => (e) => /^Malformed XML: /.test(e.message) && pattern.test(e.message);

test('the parser never sees a DOCTYPE: the text it is given has none, the lines kept; entities in content and in attributes, markup in replacement text', () => {
  const reading = readDoctype(R('<!ENTITY e "v"><!ENTITY q \'say "hi" <x>\'><!ENTITY m "<i>x</i>">', '<a k="&e;" s="&q;">&e;&m;&nope;</a>'));
  assert.doesNotMatch(reading.text, /DOCTYPE|ENTITY/);
  assert.equal(reading.text.split('\n').length, 3);
  const doc = parseXml(R('<!ENTITY e "v"><!ENTITY q \'say "hi" <x>\'><!ENTITY m "<i>x</i>">', '<a k="&e;" s="&q;">&e;&m;&nope;</a>'));
  assert.equal(doc.doctype, null);
  const a = doc.getElementsByTagName('a')[0];
  assert.equal(a.getAttribute('k'), 'v');
  assert.equal(a.getAttribute('s'), 'say "hi" <x>');
  assert.equal(a.getElementsByTagName('i')[0].textContent, 'x');
  // an entity no declaration names reads as its own text, as xmldom read it before
  assert.equal(a.textContent, 'vx&nope;');
  // a character reference in an entity value is expanded at the declaration, its result read at the use (XML 1.0 §4.5)
  assert.equal(parseXml(R('<!ENTITY amp2 "&#38;#38;">', '&amp2;')).documentElement.textContent, '&');
  // a parse error after the DOCTYPE keeps the source's line
  assert.throws(() => parseXml(R('<!ENTITY e "v">', '<a></b>')), (e) => e.locator?.lineNumber === 3);
});

test('expansion is bounded: nesting past 8, references past 10,000, replacement text past 1 MiB, an entity naming itself, each refused as XML not well formed', () => {
  assert.deepEqual({ ...DOCTYPE_LIMITS }, { depth: 8, references: 10_000, replacement: 1024 * 1024 });
  // the billion laughs: ten levels of ten
  const laughs = ['<!ENTITY lol0 "lol">', ...Array.from({ length: 9 }, (_, i) => `<!ENTITY lol${i + 1} "${`&lol${i};`.repeat(10)}">`)].join('');
  assert.throws(() => parseXml(R(laughs, '&lol9;')), malformedBecause(/nested more than 8 deep/));
  // shallow but wide: 2 levels, 100 x 200 references
  assert.throws(() => parseXml(R(`<!ENTITY x "x"><!ENTITY h "${'&x;'.repeat(100)}">`, '&h;'.repeat(200))), malformedBecause(/more than 10000 entity references/));
  // few references, long text
  assert.throws(() => parseXml(R(`<!ENTITY big "${'y'.repeat(4096)}">`, '&big;'.repeat(300))), malformedBecause(/replacement text/));
  assert.throws(() => parseXml(R('<!ENTITY a "&b;"><!ENTITY b "&a;">', '&a;')), malformedBecause(/refers to itself/));
  // within the limits: 7 levels of 3
  const fine = ['<!ENTITY n0 "z">', ...Array.from({ length: 7 }, (_, i) => `<!ENTITY n${i + 1} "${`&n${i};`.repeat(3)}">`)].join('');
  assert.equal(parseXml(R(fine, '&n7;')).documentElement.textContent.length, 3 ** 7);
});

test('a part over a limit is refused as a part not well formed is, naming the part and the limit', async () => {
  const laughs = ['<!ENTITY lol0 "lol">', ...Array.from({ length: 9 }, (_, i) => `<!ENTITY lol${i + 1} "${`&lol${i};`.repeat(10)}">`)].join('');
  const bytes = withEntry('dtd-06-document-entity-used', 'word/document.xml', (x) => x.replace(/<!DOCTYPE[^\]]*\]>/, `<!DOCTYPE w:document [${laughs}]>`).replace('&bp;', '&lol9;'));
  const restore = quiet();
  try {
    const pkg = await OpcPackage.load(bytes);
    await assert.rejects(pkg.getMainDocumentPart().getContents(), (e) => e instanceof PartUnmarshalException && e.partName === '/word/document.xml' && /nested more than 8 deep/.test(e.message));
  } finally { restore(); }
});

test('nothing outside the text is read: an external entity reads as nothing, a parameter entity stops later declarations, an external subset is named and left', () => {
  const external = readDoctype(R('<!ENTITY secret SYSTEM "file:///etc/passwd"><!ENTITY pub PUBLIC "-//x//y" "http://example.invalid/e">', 'a&secret;b&pub;c'));
  assert.deepEqual(external.summary.holds, ['entities', 'external']);
  assert.deepEqual(external.summary.notRead, ['&secret;', '&pub;']);
  assert.equal(parseXml(R('<!ENTITY secret SYSTEM "file:///etc/passwd">', 'a&secret;b')).documentElement.textContent, 'ab');
  // XML 1.0 §5.1: after a reference to a parameter entity that isn't read, entity and attribute-list declarations aren't processed
  const doc = parseXml(R('<!ENTITY before "1"><!ENTITY % pe SYSTEM "http://example.invalid/pe">%pe;<!ENTITY after "2"><!ATTLIST r k CDATA "d">', '&before;&after;'));
  assert.equal(doc.documentElement.textContent, '1&after;');
  assert.equal(doc.documentElement.hasAttribute('k'), false);
  const withSubset = readDoctype('<!DOCTYPE r SYSTEM "http://example.invalid/r.dtd" [<!ENTITY e "v">]><r>&e;</r>');
  assert.deepEqual(withSubset.summary.notRead, ['SYSTEM "http://example.invalid/r.dtd"']);
  assert.equal(parseXml('<!DOCTYPE r SYSTEM "http://example.invalid/r.dtd" [<!ENTITY e "v">]><r>&e;</r>').documentElement.textContent, 'v');
});

test('attribute defaults: a literal or #FIXED value applied where the attribute is absent, in its namespace; a stated value kept; a default namespace declaration not applied', () => {
  const doc = parseXml('<!DOCTYPE w:r [<!ATTLIST w:b w:val CDATA "0" w:x CDATA #FIXED "f" w:y CDATA #IMPLIED><!ATTLIST w:r xmlns:z CDATA "urn:z">]>'
    + '<w:r xmlns:w="urn:w"><w:b/><w:b w:val="1"/></w:r>');
  const [first, second] = [...doc.getElementsByTagName('w:b')];
  assert.equal(first.getAttributeNS('urn:w', 'val'), '0');
  assert.equal(first.getAttributeNS('urn:w', 'x'), 'f');
  assert.equal(first.hasAttribute('w:y'), false);
  assert.equal(second.getAttributeNS('urn:w', 'val'), '1');
  assert.equal(doc.documentElement.hasAttribute('xmlns:z'), false);
});

// ------------------------------------------------------------------------------------------------ the save (D4)

test('a save after an edit writes every part that held a DOCTYPE without it; an untouched save passes each through', async () => {
  for (const [name, entry] of [['dtd-13-header-bare', 'word/header1.xml'], ['dtd-10-styles-bare', 'word/styles.xml'], ['dtd-15-store-bare', 'customXml/item1.xml'], ['dtd2-22-footnotes', 'word/footnotes.xml']]) {
    const source = strFromU8(entriesOf(bytesOf(`${name}.docx`))[entry]);
    assert.match(source, /<!DOCTYPE/, name);
    const untouched = await load(name);
    const passed = entriesOf(await untouched.save());
    assert.equal(strFromU8(passed[entry]), source, `${name}: passed through`);
    assert.deepEqual(untouched.lastSave.doctypes, []);
    const edited = await load(name);
    const body = await edited.getMainDocumentPart().getBody();
    body.insertParagraph('EDITED', 'End');
    const written = entriesOf(await edited.save());
    assert.doesNotMatch(strFromU8(written[entry]), /<!DOCTYPE/, `${name}: written afresh`);
    assert.deepEqual(edited.lastSave.doctypes, [`/${entry}`], name);
    // and it reads as it did: the same text, the edit beside it
    const restore = quiet();
    try {
      const again = await OpcPackage.load(await edited.save());
      assert.equal((await textOf(again)).replace('EDITED\n', ''), await textOf(await load(name)));
      assert.equal(again.warnings.filter((w) => w.code === 'xml.doctype').length, 0);
    } finally { restore(); }
  }
  // the main part's own DOCTYPE goes with any marshal of it, and is listed
  const main = await load('dtd-01-document-bare');
  (await main.getMainDocumentPart().getBody()).insertParagraph('EDITED', 'End');
  assert.doesNotMatch(strFromU8(entriesOf(await main.save())['word/document.xml']), /<!DOCTYPE/);
  assert.deepEqual(main.lastSave.doctypes, ['/word/document.xml']);
  // a signature part is left as it came: its bytes are what it signs
  const sig = await load('dtd2-91-signature', 'word');
  (await sig.getMainDocumentPart().getBody()).insertParagraph('EDITED', 'End');
  assert.match(strFromU8(entriesOf(await sig.save())['_xmlsignatures/sig1.xml']), /<!DOCTYPE/);
  assert.deepEqual(sig.lastSave.doctypes, []);
});

