// CR-031: the repair option, WD REQ-207 and REQ-208. `refuse`, the default, keeps every refusal; `word` follows Word's repair where
// the WD states one; `all` does that and makes this package's own repairs, one today: an External officeDocument relationship beside
// an internal one dropped (REQ-206's NOTE). The x-ext round is in test/repair/, the share's files copied unedited with a SHA-1
// manifest that also records Word 365's two opens of each: a repair offered, then refused. CR-024's two readings under `word` are
// held in duplicate-entries.test.mjs, and the relationships part in a wrong namespace in relationships-namespace.test.mjs.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import {
  OpcPackage, WordprocessingMLPackage, Texts, PlainText, TextOptions, TextResult, ExternalMainPartException, InvalidFormatException, REPAIR_MODES,
} from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, 'repair');
const manifest = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8'));
const URL = 'http://example.org/x.docx';
const STEMS = ['x-ext-first-source', 'x-ext-last-source', 'x-ext-first-saved', 'x-ext-last-saved'];

const bytesOf = (name) => new Uint8Array(readFileSync(join(dir, name)));
const textOf = async (pkg) => PlainText.of(await Texts.of(pkg)).replace(/\s+/g, ' ').trim();

test('CR-031: the x-ext round is as the manifest has it, and Word refused every open', () => {
  for (const f of manifest.files) assert.equal(createHash('sha1').update(bytesOf(f.name)).digest('hex'), f.sha1, f.name);
  assert.deepEqual(manifest.files.map((f) => f.name), [...STEMS.map((s) => `${s}.docx`),
    ...STEMS.flatMap((s) => [`ctra-${s}.docx`, `ctra-${s}-word365.docx`, `ctra-${s}-word365.txt`])]);
  for (const stem of STEMS) {
    for (const run of ['a', 'b']) assert.deepEqual(manifest.word365[stem][run].dialogs, ['repair', 'error'], `${stem} ${run}`);
    for (const run of ['a', 'b']) assert.equal(manifest.word365[stem][run].opened, 'refused', `${stem} ${run}`);
  }
});

test('CR-031 D3: each x-ext file is refused at refuse and word, first or last alike, naming the URL and the part', async () => {
  for (const stem of STEMS) {
    for (const repair of [undefined, 'refuse', 'word']) {
      await assert.rejects(OpcPackage.load(bytesOf(`${stem}.docx`), { repair }), (e) => e instanceof ExternalMainPartException
        && e instanceof InvalidFormatException && e.code === 'package.external_main_part' && e.urls.join() === URL
        && e.targets.join() === '/word/document.xml' && e.message.includes(URL) && /repair: 'all'/.test(e.message), `${stem} ${repair}`);
    }
  }
});

test('CR-031 section 2.2: at all each x-ext file reads its internal part, the External relationship dropped and said, repaired own, and a save writes none', async () => {
  for (const stem of STEMS) {
    const pkg = await OpcPackage.load(bytesOf(`${stem}.docx`), { repair: 'all' });
    assert.ok(pkg instanceof WordprocessingMLPackage, stem);
    assert.equal(pkg.getMainDocumentPart().partName.name, '/word/document.xml');
    assert.deepEqual(pkg.externalMainPart && [pkg.externalMainPart.code, pkg.externalMainPart.partName, pkg.externalMainPart.urls],
      ['package.external_main_part', '/word/document.xml', [URL]], stem);
    assert.deepEqual(pkg.warnings.filter((w) => w.code.startsWith('package.')).map((w) => w.code), ['package.external_main_part'], stem);
    assert.equal(pkg.repaired, 'own', stem);
    assert.equal(pkg.relationshipsPart.list.filter((r) => r.targetMode === 'External').length, 0, stem);
    const result = await Texts.of(pkg);
    assert.deepEqual(result.warnings.filter((w) => w.kind === 'package.external_main_part').map((w) => [w.selector, w.detail]), [['/word/document.xml', URL]]);
    assert.equal(result.metadata[TextResult.REPAIRED], 'own', stem);
    const text = await textOf(pkg);
    assert.ok(text.length > 0, stem);
    const saved = await pkg.save();
    assert.doesNotMatch(strFromU8(unzipSync(saved)['_rels/.rels']), /External|example\.org/, stem);
    // the save is an ordinary package: it loads at the default with nothing said, and reads the same
    const again = await OpcPackage.load(saved);
    assert.deepEqual(again.warnings, [], stem);
    assert.equal(again.repaired, 'none', stem);
    assert.equal(await textOf(again), text, stem);
  }
});

test('CR-031 D7: Word opened every save repair all writes with no message, its text the reading here, and its own save reads at the default with nothing said', async () => {
  for (const stem of STEMS) {
    // the round holds today's save: the parts Word opened are what repair: 'all' writes now (the zip's own times aside)
    const ours = unzipSync(await (await OpcPackage.load(bytesOf(`${stem}.docx`), { repair: 'all' })).save());
    const sent = unzipSync(bytesOf(`ctra-${stem}.docx`));
    assert.deepEqual(Object.keys(ours), Object.keys(sent), stem);
    for (const name of Object.keys(sent)) assert.equal(strFromU8(ours[name]), strFromU8(sent[name]), `${stem} ${name}: the save sent`);
    for (const run of ['a', 'b']) assert.deepEqual(manifest.d7.word365[stem][run], { opened: 'silent', dialogs: [], word: '16.0.20326.20232' }, `${stem} ${run}`);
    const word = readFileSync(join(dir, `ctra-${stem}-word365.txt`), 'utf8').replace(/\s+/g, ' ').trim();
    assert.equal(word, 'DOCUMENT', stem);
    assert.equal(await textOf(await OpcPackage.load(bytesOf(`ctra-${stem}.docx`))), word, `${stem}: the save reads as Word reads it`);
    const saved = await OpcPackage.load(bytesOf(`ctra-${stem}-word365.docx`));
    assert.deepEqual(saved.warnings, [], stem);
    assert.equal(await textOf(saved), word, `${stem}: Word's save`);
  }
});

test('CR-031 section 2.4: repaired is none where nothing was repaired, at every value, and the result says it whatever the options', async () => {
  const plain = await (await WordprocessingMLPackage.createPackage()).save();
  for (const repair of REPAIR_MODES) {
    const pkg = await OpcPackage.load(plain, { repair });
    assert.equal(pkg.repaired, 'none', repair);
    assert.equal((await Texts.of(pkg, { warnings: false, metadata: true })).metadata.repaired, 'none', repair);
  }
  assert.deepEqual([...REPAIR_MODES], ['refuse', 'word', 'all']);
});

test('CR-031 D6: repair leaves the save\'s own gate alone, at every value', async () => {
  const plain = await (await WordprocessingMLPackage.createPackage()).save();
  for (const repair of [undefined, ...REPAIR_MODES]) assert.equal((await OpcPackage.load(plain, { repair })).checkLossOnSave, true, String(repair));
});

test('CR-031: the text model takes repair by the WD\'s name, as the suite\'s adapter passes it to the load', () => {
  assert.equal(TextOptions.of({}).repair, 'refuse');
  for (const value of REPAIR_MODES) assert.equal(TextOptions.of({ repair: value }).repair, value);
});
