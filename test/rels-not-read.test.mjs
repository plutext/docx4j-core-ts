// CR-032: a relationships part that isn't well formed, read as Word's repair reads it (WD 0.76's ruling on WD 0.73's item (i);
// docx4j CR-075 D5, D6 and D8, released in 17.3.6). At every value of repair the relationships before the break are read, a part
// only a lost one named is absent, and part.not_read names the relationships part with the count kept. A save is refused unless
// the load's repair is 'word' or 'all', which write that repaired form as Word's save does, or checkLossOnSave is false.
// `_rels/.rels` without the main part's relationship among those kept is refused at the load. docx4j's broken-part files and
// tables are in test/rels-not-read/, with a SHA-1 manifest.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strToU8, strFromU8, unzipSync } from 'fflate';
import { OpcPackage, Texts, PlainText, PartUnmarshalException, SaveLossException } from '../dist/index.mjs';
import { storedZip } from './helpers.mjs';
import { table, verify } from './text-extraction/conformance.mjs';
import { textsAdapter } from './text-extraction/adapter.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, 'rels-not-read');
const manifest = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8'));
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const INTACT = 'broken-part-intact';
const DOC = 'broken-part-document-rels';
const PKG = 'broken-part-package-rels';
// the relationships part each reads short, and the relationships kept: the three before the break, the main part's alone
const NOT_READ = { [DOC]: ['/word/_rels/document.xml.rels', 3], [PKG]: ['/_rels/.rels', 1] };
// what only a lost relationship named, as docx4j's README and Word's repair have it
const LOST = {
  [DOC]: ['/word/comments.xml', '/word/commentsExtended.xml'],
  [PKG]: ['/docProps/app.xml', '/docProps/core.xml', '/docProps/custom.xml'],
};
const CASE = { [DOC]: '119', [PKG]: '120' };
const REPAIRS = ['refuse', 'word', 'all'];

const bytesOf = (name) => new Uint8Array(readFileSync(join(dir, name)));
const load = (name, options) => OpcPackage.load(bytesOf(`${name}.docx`), options);
const partsOf = (pkg) => [...pkg.parts].map((p) => p.partName.name).sort();
const entriesOf = (bytes) => Object.fromEntries(Object.entries(unzipSync(bytes)).map(([k, v]) => [k, Buffer.from(v).toString('hex')]));

/** The intact twin with entries replaced, as a stored zip. */
const intactWith = (replaced) => {
  const entries = unzipSync(bytesOf(`${INTACT}.docx`));
  for (const [name, text] of Object.entries(replaced)) entries[name] = strToU8(text);
  return storedZip(Object.entries(entries));
};
const intactRels = () => strFromU8(unzipSync(bytesOf(`${INTACT}.docx`))['word/_rels/document.xml.rels']);
const intactPackageRels = () => strFromU8(unzipSync(bytesOf(`${INTACT}.docx`))['_rels/.rels']);

test('CR-032: the files are as the manifest has them', () => {
  for (const f of manifest.files) assert.equal(createHash('sha1').update(bytesOf(f.name)).digest('hex'), f.sha1, f.name);
  assert.equal(manifest.files.length, 9);
});

test('CR-032 D1: each file reads the relationships before the break, at every value of repair, and a part only a lost one named is absent', async () => {
  const intact = partsOf(await load(INTACT));
  const control = await load(INTACT);
  assert.deepEqual(control.relationshipsNotRead, [], 'the intact twin: nothing read short');
  for (const name of [DOC, PKG]) {
    for (const repair of REPAIRS) {
      const at = `${name} ${repair}`;
      const pkg = await load(name, { repair });
      const [partName, count] = NOT_READ[name];
      assert.equal(pkg.relationshipsNotRead.length, 1, at);
      const w = pkg.relationshipsNotRead[0];
      assert.deepEqual([w.code, w.partName, w.count], ['part.not_read', partName, count], at);
      assert.match(w.reason, /^Malformed XML \(line \d+, column \d+\): /, at);
      assert.equal(typeof w.line, 'number', at);
      assert.equal(typeof w.column, 'number', at);
      assert.ok(pkg.warnings.includes(w), `${at}: on pkg.warnings`);
      assert.equal(pkg.repaired, 'none', `${at}: not a repair the option makes; docx4j doesn't count it either`);
      assert.deepEqual(partsOf(pkg), intact.filter((p) => !LOST[name].includes(p)), `${at}: the parts`);
      for (const lost of LOST[name]) assert.equal(pkg.contentTypeManager.getOverrideContentType(lost), undefined, `${at}: ${lost}'s Override gone`);
    }
  }
});

test('CR-032 D1: docx4j\'s tables, cases 119 and 120: the tokens in the three views, and part.not_read in each', async () => {
  const engine = textsAdapter();
  const tokens = table(join(dir, 'expected-tokens.tsv'));
  const suite = tokens.map((r) => r.token);
  const findings = table(join(dir, 'expected-findings.tsv'));
  let checked = 0;
  for (const name of [DOC, PKG]) {
    const rows = tokens.filter((r) => r.case === CASE[name]);
    assert.equal(rows.length, { [DOC]: 19, [PKG]: 9 }[name], name);
    for (const view of ['accepted', 'original', 'markup']) {
      const out = await engine.extract(join(dir, `${name}.docx`), view, 'default', 'txt');
      for (const row of rows) {
        verify(out.text, row.token, row[view], `${row.case}:${row.token}:${view}`, suite);
        checked++;
      }
      const finding = findings.filter((f) => f.case === CASE[name]);
      assert.equal(finding.length, 1, name);
      const key = `${finding[0].channel}\t${finding[0].kind}\t${finding[0].selector}`;
      assert.equal(out.findings[key], finding[0].expected, `${name} ${view}: ${key}`);
    }
  }
  assert.equal(checked, 3 * 28, 'twenty-eight rows, three views');
});

test('CR-032: the tables\' with_option column is Word\'s text, and this package reads Word\'s profile alike where it can (no fields: update here)', async () => {
  const engine = textsAdapter();
  const tokens = table(join(dir, 'expected-tokens.tsv'));
  const suite = tokens.map((r) => r.token);
  for (const name of [DOC, PKG]) {
    const word = readFileSync(join(dir, `${name}-word365-utf8.txt`), 'utf8').replace(/^﻿/, '');
    const rows = tokens.filter((r) => r.case === CASE[name]);
    for (const row of rows) verify(word, row.token, row.with_option, `${row.case}:${row.token}: Word's text`, suite);
    // the profile without fields: update, which this package doesn't implement (CR-012 phase 5a): every row not about a field
    const out = await engine.extract(join(dir, `${name}.docx`), 'accepted', 'profile: word-text; list_labels: true', 'txt');
    for (const row of rows.filter((r) => !/REQ-506/.test(r.statements))) {
      verify(out.text, row.token, row.with_option, `${row.case}:${row.token}: profile word-text`, suite);
    }
  }
});

test('CR-032 D2: a save is refused at the default, the part and the count named, and nothing is written', async () => {
  for (const name of [DOC, PKG]) {
    const pkg = await load(name);
    await assert.rejects(pkg.save(), (e) => {
      assert.ok(e instanceof SaveLossException, name);
      assert.equal(e.code, 'save.content_lost');
      const [partName, count] = NOT_READ[name];
      assert.deepEqual(e.losses, [{ partName, kind: 'relationships', count }]);
      assert.match(e.message, new RegExp(`${partName.replace(/[./]/g, '\\$&')} isn't well formed, and only the ${count} relationships? before the break w(?:as|ere) read`));
      assert.match(e.message, /repair: 'word'/);
      assert.match(e.message, /checkLossOnSave = false/);
      return true;
    });
    assert.equal(pkg.lastSave, undefined, `${name}: nothing written`);
    await assert.rejects(pkg.clone(), SaveLossException, `${name}: clone() goes through a save, as CR-025 has it`);
  }
});

test('CR-032 D2: under repair word and all, and with checkLossOnSave off, the save writes the repaired form, the same each way', async () => {
  for (const name of [DOC, PKG]) {
    const off = await load(name);
    off.checkLossOnSave = false;
    const written = await off.save();
    assert.deepEqual(off.lastSave.relationshipsNotRead.map((w) => [w.partName, w.count]), [NOT_READ[name]], name);
    const entries = entriesOf(written);
    for (const repair of ['word', 'all']) {
      const pkg = await load(name, { repair });
      const bytes = await pkg.save();
      assert.deepEqual(pkg.lastSave.relationshipsNotRead.map((w) => [w.partName, w.count]), [NOT_READ[name]], `${name} ${repair}`);
      assert.deepEqual(entriesOf(bytes), entries, `${name} ${repair}: the same save`);
      const clone = await pkg.clone();
      assert.deepEqual(clone.relationshipsNotRead, [], `${name} ${repair}: the clone reads its repaired form`);
    }
    // the save reads at the default with nothing said: the relationships kept, well formed; no part or Override of what was lost
    const again = await OpcPackage.load(written);
    assert.deepEqual(again.relationshipsNotRead, [], name);
    const source = unzipSync(bytesOf(`${name}.docx`));
    const saved = unzipSync(written);
    for (const lost of LOST[name]) {
      assert.equal(saved[lost.slice(1)], undefined, `${name}: ${lost} not written`);
      assert.ok(!strFromU8(saved['[Content_Types].xml']).includes(`PartName="${lost}"`), `${name}: no Override for ${lost}`);
    }
    const rels = NOT_READ[name][0].slice(1);
    assert.equal((strFromU8(saved[rels]).match(/<Relationship /g) ?? []).length, NOT_READ[name][1], `${name}: ${rels} as kept`);
    assert.deepEqual(saved['word/document.xml'], source['word/document.xml'], `${name}: the main part, untouched, copied byte for byte`);
    assert.equal(PlainText.of(await Texts.of(again)), PlainText.of(await Texts.of(await load(name))), `${name}: the save reads as the short read`);
  }
});

test('CR-032 D2: a caller who puts the relationships back says so with setContents, and the save goes', async () => {
  const pkg = await load(DOC);
  const rp = pkg.getMainDocumentPart().relationshipsPart;
  rp.setContents(rp.contents);
  assert.deepEqual(pkg.relationshipsNotRead, []);
  await pkg.save();
  assert.deepEqual(pkg.lastSave.relationshipsNotRead, []);
});

test('CR-032 D3: _rels/.rels without the main part\'s relationship among those kept is refused at the load, naming the part and the parse error', async () => {
  const rels = intactPackageRels();
  const cut = rels.slice(0, rels.indexOf('<Relationship ') + 20);
  for (const repair of REPAIRS) {
    await assert.rejects(OpcPackage.load(intactWith({ '_rels/.rels': cut }), { repair }), (e) => {
      assert.ok(e instanceof PartUnmarshalException, repair);
      assert.equal(e.partName, '/_rels/.rels');
      assert.match(e.message, /^\/_rels\/\.rels: Malformed XML \(line \d+, column \d+\): .*; the main part's relationship isn't among the 0 before the break/);
      assert.ok(e.cause instanceof PartUnmarshalException);
      assert.equal(typeof e.line, 'number');
      return true;
    });
  }
  // the main part's relationship second, after one kept: still refused
  const reordered = rels.replace(/(<Relationship Id="rId1"[^>]*\/>)(<Relationship Id="rId2"[^>]*\/>)/, '$2$1');
  const second = reordered.slice(0, reordered.indexOf('<Relationship Id="rId1"') + 10);
  await assert.rejects(OpcPackage.load(intactWith({ '_rels/.rels': second })), /isn't among the 1 before the break/);
});

test('CR-032 D1: the shapes, as docx4j reads them', async () => {
  const rels = intactRels();
  const head = rels.slice(0, rels.indexOf('<Relationship '));
  const all = [...rels.matchAll(/<Relationship [^>]*\/>/g)].map((m) => m[0]);
  const read = async (text, options) => {
    const pkg = await OpcPackage.load(intactWith({ 'word/_rels/document.xml.rels': text }), options);
    const rp = pkg.getMainDocumentPart().relationshipsPart;
    return { pkg, ids: rp.list.map((r) => r.id), w: pkg.relationshipsNotRead[0] };
  };

  // cut inside the root's start tag: nothing kept, nothing the main part names read
  let r = await read(rels.slice(0, rels.indexOf('<Relationships') + 20));
  assert.deepEqual([r.ids, r.w.count], [[], 0]);
  assert.deepEqual(partsOf(r.pkg), ['/docProps/app.xml', '/docProps/core.xml', '/docProps/custom.xml', '/word/document.xml']);

  // the second relationship giving its Target twice, the rest intact: the first alone, before the break. (xmldom reads an undefined
  // entity, an unquoted value or a lone '&' through, where docx4j's StAX reader stops: what breaks a part is the parser's, CR-032 section 5)
  r = await read(`${head}${all[0]}${all[1].replace('Target="numbering.xml"', 'Target="numbering.xml" Target="again.xml"')}${all.slice(2).join('')}</Relationships>`);
  assert.deepEqual([r.ids, r.w.count], [['rId5'], 1]);
  r = await read(`${head}${all[0]}${all[1]}${all[2].replace('Target="settings.xml"', 'Target="set<tings.xml"')}${all.slice(3).join('')}</Relationships>`);
  assert.deepEqual([r.ids, r.w.count], [['rId5', 'rId6'], 2], "an unescaped '<' in the third's value");

  // a Relationship without a Target before the break is passed over, as neither reader could use it
  r = await read(`${head}${all[0].replace(/ Target="[^"]*"/, '')}${all[1]}${all[2]}`);
  assert.deepEqual([r.ids, r.w.count], [['rId6', 'rId7'], 2]);

  // a comment holding a Relationship's text, and a Target holding '>', before the break
  const gt = `<Relationship Id="rId20" Type="${R}/hyperlink" Target="https://example.com/a&gt;b?x=1>2" TargetMode="External"/>`;
  r = await read(`${head}<!-- <Relationship Id="rId99" Type="x" Target="y"/> -->${all[0]}${gt}${all[1]}`);
  assert.deepEqual([r.ids, r.w.count], [['rId5', 'rId20', 'rId6'], 3]);
  const link = r.pkg.getMainDocumentPart().relationshipsPart.getRelationshipById('rId20');
  assert.deepEqual([link.target, link.targetMode], ['https://example.com/a>b?x=1>2', 'External']);

  // what follows the root's end tag breaks it: every relationship is before the break
  r = await read(`${rels}<junk>`);
  assert.deepEqual([r.ids, r.w.count], [['rId5', 'rId6', 'rId7', 'rId8', 'rId9'], 5]);

  // relationships in another namespace, or none, before a break: kept by their local names, and saved in the right one
  const wrong = `<?xml version="1.0"?><r:Relationships xmlns:r="urn:docx4j:wrong">${all.slice(0, 2).map((x) => x.replace('<Relationship ', '<r:Relationship ')).join('')}<Relationship Id="rId7" Type="${R}/settings" Target="settings.xml"/>`;
  r = await read(wrong);
  assert.deepEqual([r.ids, r.w.count], [['rId5', 'rId6', 'rId7'], 3], 'at the default too: a break, not a namespace refusal');
  r = await read(wrong, { repair: 'word' });
  assert.deepEqual([r.ids, r.w.count], [['rId5', 'rId6', 'rId7'], 3]);
  const saved = strFromU8(unzipSync(await r.pkg.save())['word/_rels/document.xml.rels']);
  assert.ok(saved.includes(`<Relationships xmlns="${PR}">`), saved);
  assert.ok(!saved.includes('urn:docx4j:wrong'), saved);

  // a part that parses to its end keeps every relationship and says nothing, whatever its other faults
  r = await read(rels);
  assert.deepEqual([r.ids.length, r.w], [5, undefined]);
});

test('CR-032 D1: the text model raises part.not_read, the relationships part its selector and the count its detail', async () => {
  for (const name of [DOC, PKG]) {
    const result = await Texts.of(await load(name));
    const w = result.warnings.filter((x) => x.kind === 'part.not_read');
    assert.deepEqual(w.map((x) => [x.partName, x.detail, x.address]), [[...NOT_READ[name], undefined]], name);
    assert.equal(result.metadata.repaired, 'none', name);
  }
});
