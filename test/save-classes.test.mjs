// CR-025 section 10: what a read skipped is classed by the Word file that measured its shape. `dropped`, Word's own save drops it
// too, so a save of the changed part writes it without and lists it on lastSave.dropped, never refused; `lost`, Word's save keeps
// it or Word won't open the file or nobody measured it, so the save is refused unless checkLossOnSave is off (D14, D15). Held to
// Word check 76's own verdicts: every file Word opened with no message holds no lost item but 76t71, where Word keeps words this
// package drops (CR-026's known difference), and every file Word refused that holds an item holds a lost one.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { OpcPackage, WordprocessingMLPackage, SaveLossException } from '../dist/index.mjs';
import { check76Files } from './fixtures/check76/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const EXT = 'urn:docx4j:cr025:ext';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

// Word 365's refusals in check 76 (test/check76/results.txt): round 1 76e; round 2 "p q r s v x y" and 76o; round 6 76t28; round 8
// 76t46 to 76t49; rounds 9 and 10 76t53, 76t54, 76t55, 76t57, 76t59. Every other file opened with no message.
const REFUSED = new Set(['76e', '76o', '76p', '76q', '76r', '76s', '76v', '76x', '76y', '76t28', '76t46', '76t47', '76t48', '76t49',
  '76t53', '76t54', '76t55', '76t57', '76t59']);
/** Word opened it with no message and kept words this package drops: lost is the class (CR-026 section 5.2). */
const WORD_KEEPS = new Set(['76t71']);
const stemOf = (name) => name.split('-')[0];
const skippedOf = (pkg) => pkg.warnings.filter((w) => w.code === 'part.skipped');

test('CR-025 section 10: the refusals listed are the bridge\'s where it ran (rounds 6 to 11)', () => {
  const dir = join(here, 'check76');
  const results = readdirSync(dir).filter((n) => n.endsWith('-bridge-result.json'));
  assert.ok(results.length >= 45);
  for (const name of results) {
    const { opened } = JSON.parse(readFileSync(join(dir, name), 'utf8'));
    assert.equal(opened === 'refused', REFUSED.has(stemOf(name)), name);
  }
});

test('CR-025 section 10: check 76\'s classes hold to Word\'s verdict on each file', async () => {
  let files = 0;
  for (const f of await check76Files()) {
    files++;
    const stem = stemOf(f.name);
    const pkg = await OpcPackage.load(f.bytes);
    await pkg.unmarshalAll();
    const items = skippedOf(pkg);
    for (const w of items) {
      assert.ok(w.onSave === 'dropped' || w.onSave === 'lost', `${f.name}: ${w.name}`);
      if (w.onSave === 'dropped') assert.ok(typeof w.measured === 'string' && w.measured.length > 0, `${f.name}: a dropped item names what decides it`);
    }
    const lost = items.filter((w) => w.onSave === 'lost');
    if (REFUSED.has(stem)) assert.ok(items.length === 0 || lost.length > 0, `${f.name}: Word refuses it, so what the read skipped is lost`);
    else if (WORD_KEEPS.has(stem)) assert.ok(lost.length > 0, `${f.name}: Word keeps what this package drops`);
    else assert.deepEqual(lost.map((w) => w.name), [], `${f.name}: Word opens it with no message and its save drops the same`);
  }
  assert.equal(files, 98);
});

/** A package whose main part is `body`, with `ext` declared on the root and listed in mc:Ignorable when `ignorable`. */
async function packageOf(body, { ignorable = false, more = '' } = {}) {
  const base = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  const root = `<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:ext="${EXT}" xmlns:w14="${W14}"${ignorable ? ' mc:Ignorable="ext"' : ''}${more}>`;
  base['word/document.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${root}<w:body>${body}<w:p><w:r><w:t>END</w:t></w:r></w:p></w:body></w:document>`);
  const pkg = await WordprocessingMLPackage.load(zipSync(base));
  await pkg.unmarshalAll();
  return pkg;
}
const classes = (pkg) => skippedOf(pkg).map((w) => [w.name, w.namespace, w.onSave, w.measured ?? null]);

test('CR-025 section 10.3: each row of the table, its class and what decides it', async () => {
  const rows = [
    // an element in a namespace Word doesn't understand: ignorable, Word's save drops it, wherever it stands; not, Word refuses
    ['<w:p><ext:opaque/><w:r><w:t>a</w:t></w:r></w:p>', { ignorable: true }, [['ext:opaque', EXT, 'dropped', '76b']]],
    ['<w:tbl><ext:opaque/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>', { ignorable: true }, [['ext:opaque', EXT, 'dropped', 'mc:Ignorable']]],
    ['<w:p><ext:opaque/><w:r><w:t>a</w:t></w:r></w:p>', {}, [['ext:opaque', EXT, 'lost', '76e']]],
    ['<w:tbl><ext:opaque/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>', {}, [['ext:opaque', EXT, 'lost', null]]],
    // attributes: one in such a namespace, ignorable or not, and an unknown w: one, Word's save drops (76h, 76i, 76j)
    ['<w:p ext:note="1"><w:r><w:t>a</w:t></w:r></w:p>', { ignorable: true }, [['ext:note', EXT, 'dropped', '76h']]],
    ['<w:p ext:note="1"><w:r><w:t>a</w:t></w:r></w:p>', {}, [['ext:note', EXT, 'dropped', '76i']]],
    ['<w:p w:madeUp="1"><w:r><w:t>a</w:t></w:r></w:p>', {}, [['w:madeUp', W, 'dropped', '76j']]],
    // an unknown w: element: read through, its tags Word's reading; a property in w:pPr or w:rPr Word's save drops; a run in a
    // w:rPr or a w:tcPr Word refuses
    ['<w:p><w:madeUp><w:r><w:t>a</w:t></w:r></w:madeUp></w:p>', {}, [['w:madeUp', W, 'dropped', 'REQ-308']]],
    ['<w:p><w:pPr><w:madeUp><w:jc w:val="center"/></w:madeUp></w:pPr><w:r><w:t>a</w:t></w:r></w:p>', {}, [['w:madeUp', W, 'dropped', '76g']]],
    ['<w:p><w:r><w:rPr><w:madeUp><w:b/></w:madeUp></w:rPr><w:t>a</w:t></w:r></w:p>', {}, [['w:madeUp', W, 'dropped', '76t6']]],
    ['<w:p><w:r><w:rPr><w:madeUp><w:r><w:t>b</w:t></w:r></w:madeUp></w:rPr><w:t>a</w:t></w:r></w:p>', {}, [['w:madeUp', W, 'lost', '76t28']]],
    ['<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:madeUp><w:r><w:t>b</w:t></w:r></w:madeUp></w:tcPr><w:p/></w:tc></w:tr></w:tbl>', {}, [['w:madeUp', W, 'lost', '76t48']]],
    // a held w:p with a run after it, which the pass leaves whole: Word refuses it (76t53)
    ['<w:p><w:r><w:t>a</w:t></w:r><w:madeUp><w:p><w:r><w:t>b</w:t></w:r></w:p></w:madeUp><w:r><w:t>c</w:t></w:r></w:p>', {}, [['w:madeUp', W, 'lost', null]]],
    // in an understood namespace that the model lacks (python's 1b): Word's save keeps it
    ['<w:p><w14:madeUp/><w:r><w:t>a</w:t></w:r></w:p>', {}, [['w14:madeUp', W14, 'lost', null]]],
  ];
  for (const [body, options, expected] of rows) assert.deepEqual(classes(await packageOf(body, options)), expected, body);
});

test('CR-025 section 10.3: text an unknown w: element holds is absent in Word\'s reading (REQ-314), so dropped', async () => {
  const text = classes(await packageOf('<w:p><w:madeUp>words</w:madeUp><w:r><w:t>a</w:t></w:r></w:p>'));
  assert.deepEqual(text.find(([name]) => name === 'words'), ['words', null, 'dropped', 'REQ-314']);
  assert.ok(text.every(([, , onSave]) => onSave === 'dropped'));
});

test('CR-025 section 10.3: a changed part holding only what Word\'s save drops is written without it, listed, never refused', async () => {
  const pkg = await packageOf('<w:p><ext:opaque/><w:r><w:t>a</w:t></w:r></w:p>', { ignorable: true });
  (await pkg.getBody()).insertParagraph('ADDED', 'End');
  const saved = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
  assert.match(saved, /ADDED/);
  assert.doesNotMatch(saved, /ext:opaque/);
  assert.deepEqual(pkg.lastSave.dropped.map((w) => w.name), ['ext:opaque']);
  assert.deepEqual(pkg.lastSave.lost, []);
  // and it reloads with nothing skipped
  const again = await OpcPackage.load(await pkg.save());
  await again.unmarshalAll();
  assert.deepEqual(skippedOf(again), []);
});

test('CR-025 section 10.3: a part holding both classes is refused for its lost items alone; with the check off each goes to its list', async () => {
  const body = '<w:p ext:note="1"><w:r><w:t>a</w:t></w:r></w:p><w:p><ext:opaque/><w:r><w:t>b</w:t></w:r></w:p>';
  const pkg = await packageOf(body);
  assert.deepEqual(classes(pkg).map(([name, , onSave]) => [name, onSave]), [['ext:note', 'dropped'], ['ext:opaque', 'lost']]);
  (await pkg.getBody()).insertParagraph('ADDED', 'End');
  await assert.rejects(pkg.save(), (e) => e instanceof SaveLossException && e.losses.length === 1 && e.losses[0].count === 1
    && /onSave lost/.test(e.message));
  pkg.checkLossOnSave = false;
  const saved = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
  assert.doesNotMatch(saved, /ext:/);
  assert.deepEqual(pkg.lastSave.dropped.map((w) => w.name), ['ext:note']);
  assert.deepEqual(pkg.lastSave.lost.map((w) => w.name), ['ext:opaque']);
});

test('CR-025 section 10.3: a part read and not changed is copied byte for byte, whatever its items\' classes', async () => {
  const body = '<w:p ext:note="1"><w:r><w:t>a</w:t></w:r></w:p><w:p><ext:opaque/><w:r><w:t>b</w:t></w:r></w:p>';
  const pkg = await packageOf(body);
  const before = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
  assert.match(before, /ext:note="1"/);
  assert.match(before, /<ext:opaque\/>/);
  assert.deepEqual(pkg.lastSave.dropped, []);
  assert.deepEqual(pkg.lastSave.lost, []);
});

test('CR-025 section 10.3: a w:document inside a w:document, which Word refuses (nd-10), is lost', async () => {
  const base = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  base['word/document.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:document><w:body><w:p/></w:body></w:document><w:body><w:p><w:r><w:t>a</w:t></w:r></w:p></w:body></w:document>`);
  const pkg = await WordprocessingMLPackage.load(zipSync(base));
  await pkg.unmarshalAll();
  assert.deepEqual(classes(pkg), [['w:document', W, 'lost', 'nd-10']]);
});
