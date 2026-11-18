// CR-028 phase 2: a child given more than once read as Word 365 reads it (WD REQ-322), through the objects package's CR-022
// (0.9.0: wordOverwritten as the runtime's onOverwrittenElement on every facade read), held to Word's saves of docx4j's files.
// test/repeated-containers/ is docx4j-core-tests/src/test/resources/repeated/ at the docx4j 17.3.6 release (tag docx4j-17.3.6,
// 4afcf4be8), all 326 files copied unedited, with manifest.json's SHA-1 of each: the rounds rc-, rt-, sm-, ft-, nd-, op- and
// cr079-phase0 (c79-), each source beside Word's save (<name>-word365.docx) and its Save As text (<name>-word365.txt).
// The comparisons are docx4j's RepeatedChildrenTest's, which ports docx4j-python's test_repeated.py, so the three engines are held
// to one reading; then this package's own: its readers (the text model, Body.text, Paragraph and Font) on each source against the
// same readers on Word's save, every load route, the save and its gate, and a write.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { OpcPackage, Texts, PlainText, TextOptions, parseXml } from '../dist/index.mjs';

const dir = new URL('./repeated-containers/', import.meta.url);
const bytes = (name) => new Uint8Array(readFileSync(new URL(`${name}.docx`, dir)));
const hasSave = (name) => existsSync(new URL(`${name}-word365.docx`, dir));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SOURCES = readdirSync(dir).filter((f) => f.endsWith('.docx') && !f.endsWith('-word365.docx')).map((f) => f.slice(0, -5)).sort();
const names = (prefix) => SOURCES.filter((n) => n.startsWith(prefix));
/** `mc:AlternateContent` as the main part's root: CR-028 phase 1, not built, so read empty and refused a write (D10 as worded). */
const MC_ROOTS = ['rt-50-mc-root-three-in-choice', 'rt-51-mc-root-three-in-fallback', 'sm-52-branch-orient-first'];
/** A w:body in a w:body, with a w:sectPr of its own: the page widths of this package's save (D12, open). */
const NESTED_SECTIONS = { 'nd-22-body-middle-with-sectPr': ['15840', '11906'], 'nd-24-body-three-deep': ['12240', '15840', '11906'] };

test('the files are docx4j 17.3.6\'s repeated/, unedited', () => {
  const manifest = JSON.parse(readFileSync(new URL('manifest.json', dir), 'utf8'));
  const files = readdirSync(dir).filter((f) => f !== 'manifest.json' && f !== 'README.md').sort();
  assert.deepEqual(manifest.files.map((f) => f.name), files);
  for (const { name, sha1 } of manifest.files) {
    assert.equal(createHash('sha1').update(readFileSync(new URL(name, dir))).digest('hex'), sha1, name);
  }
  assert.equal(files.length, 326);
  assert.equal(SOURCES.length, 112);
  // Word refused the five that nest a w:document (nd-10 to nd-14), so they have no save; D12 leaves them as they read
  assert.deepEqual(SOURCES.filter((n) => !hasSave(n)), ['nd-10-document-only-child', 'nd-11-document-after-body',
    'nd-12-document-before-body', 'nd-13-document-in-body', 'nd-14-document-three-deep']);
});

// --- docx4j's helpers, on this package's DOM ------------------------------------------------------------

/** This package's save, the main part read (docx4j's `saved`): the gate on, so a save that would lose content throws. */
async function saved(name) {
  const pkg = await OpcPackage.load(bytes(name));
  await pkg.getMainDocumentPart().getContents();
  return pkg.save();
}

function entry(zip, name) {
  const found = unzipSync(zip)[name];
  assert.ok(found, name);
  return strFromU8(found);
}

const root = (zip) => parseXml(entry(zip, 'word/document.xml')).documentElement;
const all = (e, local) => Array.from(e.getElementsByTagNameNS(W, local));
const children = (e) => Array.from(e.childNodes).filter((n) => n.nodeType === 1);

/** An attribute's value as Word writes an on/off one. */
function spelled(local, value) {
  if (local !== 'val') return value;
  if (value === 'true' || value === 'on') return '1';
  if (value === 'false' || value === 'off') return '0';
  return value;
}

function attributes(e) {
  const out = {};
  for (const a of Array.from(e.attributes)) {
    const local = a.localName ?? a.nodeName;
    if (local.startsWith('rsid') || a.prefix === 'xmlns' || a.nodeName === 'xmlns') continue;
    out[local] = spelled(local, a.value);
  }
  return sorted(out);
}

const sorted = (o) => Object.fromEntries(Object.entries(o).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)));

function target(e, words) {
  const found = all(e, 'p').find((p) => p.textContent.includes(words));
  assert.ok(found, `no paragraph holding ${words}`);
  return found;
}

/** python's target_properties: what the paragraph's w:pPr and w:rPr hold, each child's attributes, the last of a name. */
function targetProperties(e, words) {
  const out = {};
  const holders = [];
  const collect = (n) => {
    for (const c of children(n)) {
      if (c.namespaceURI === W && (c.localName === 'pPr' || c.localName === 'rPr')) holders.push(c);
      collect(c);
    }
  };
  collect(target(e, words));
  for (const properties of holders) for (const child of children(properties)) out[child.localName] = attributes(child);
  // an on/off property stated off reads as one not stated, where nothing beneath sets it: Word's save writes nothing (rc-31, sm-80)
  for (const onOff of ['b', 'vanish']) if (JSON.stringify(out[onOff]) === '{"val":"0"}') delete out[onOff];
  return sorted(out);
}

/** python's _stated: what a w:sectPr, w:tblPr, w:pPr or w:rPr states, with the defaults Word's save adds or leaves out. */
function stated(element) {
  const facts = {};
  const section = element.localName === 'sectPr';
  const table = element.localName === 'tblPr';
  let headers = 0;
  for (const child of children(element)) {
    const name = child.localName;
    if (name === 'headerReference') { headers++; continue; }
    if (table && (name === 'tblCellMar' || name === 'tblLook')) continue; // Word's own defaults
    const said = attributes(child);
    if (said.orient === 'portrait') delete said.orient; // the default, which Word's save leaves out (sm-07)
    if (name === 'cols' && !('space' in said)) said.space = '720'; // likewise the other way (sm-32)
    facts[name] = sorted(said);
  }
  if (section) {
    if (!('cols' in facts)) facts.cols = { space: '720' };
    facts.headerReference = { count: String(headers) };
  }
  return sorted(facts);
}

function merged(e) {
  const sections = all(e, 'sectPr').map(stated);
  const tables = all(e, 'tblPr').map(stated);
  const paragraph = [];
  const runs = [];
  for (const c of children(target(e, 'TARGET TEXT'))) {
    if (c.localName === 'pPr' && children(c).length > 0) paragraph.push(stated(c));
    if (c.localName === 'r') {
      for (const rPr of children(c)) {
        if (rPr.localName !== 'rPr' || children(rPr).length === 0) continue;
        const facts = stated(rPr);
        if (JSON.stringify(facts) !== '{"vanish":{"val":"0"}}') runs.push(facts);
      }
    }
  }
  return { sections, tables, paragraph, runs };
}

const pageWidths = (e) => all(e, 'sectPr').flatMap((s) => children(s).filter((c) => c.localName === 'pgSz').map((c) => c.getAttributeNS(W, 'w')));

function numPrOf(zip, words) {
  const numPr = all(target(root(zip), words), 'numPr')[0];
  return Object.fromEntries(children(numPr).map((c) => [c.localName, c.getAttributeNS(W, 'val')]));
}

function references(zip) {
  const rels = parseXml(entry(zip, 'word/_rels/document.xml.rels')).documentElement;
  const targets = new Map(children(rels).map((r) => [r.getAttribute('Id'), r.getAttribute('Target')]));
  return all(root(zip), 'sectPr').map((section) => children(section)
    .filter((ref) => ref.localName === 'headerReference' || ref.localName === 'footerReference')
    .map((ref) => {
      const words = parseXml(entry(zip, `word/${targets.get(ref.getAttributeNS(R, 'id'))}`)).documentElement.textContent.trim();
      return `${ref.localName} ${ref.getAttributeNS(W, 'type')} ${words}`;
    }).sort());
}

/** A copy of a file with its main part, or another part, replaced. */
function withPart(name, part, xml) {
  const entries = unzipSync(bytes(name));
  entries[part] = strToU8(xml);
  return zipSync(entries);
}

// --- docx4j's RepeatedChildrenTest ----------------------------------------------------------------------

const PROPERTIES = ['rc-20-pPr-values', 'rc-21-pPr-merge', 'rc-22-rPr-values', 'rc-23-rPr-merge', 'rc-30-jc', 'rc-31-b-on-off',
  'rc-32-b-off-on', 'rc-33-sz', 'rt-20-pPr-values', 'rt-21-pPr-merge', 'rt-22-rPr-values', 'rt-23-rPr-merge', 'rt-30-jc',
  'rt-31-b-on-off-on', 'rt-32-b-off-on-off', 'rt-33-sz'];

test('a property given twice is what Word\'s save holds', async () => {
  for (const name of PROPERTIES) {
    assert.deepEqual(targetProperties(root(await saved(name)), 'TARGET TEXT'), targetProperties(root(bytes(`${name}-word365`)), 'TARGET TEXT'), name);
  }
  // cr079-phase0: a w:pPr's w:rPr merges too, and a w:pPrChange in the first copy survives
  assert.deepEqual(targetProperties(root(await saved('c79-40-ppr-nested-rpr')), 'MARKED WORDS'),
    targetProperties(root(bytes('c79-40-ppr-nested-rpr-word365')), 'MARKED WORDS'));
  assert.equal(all(target(root(await saved('c79-42-ppr-pprchange')), 'CHANGED WORDS'), 'pPrChange').length, 1);
});

test('a w:numPr is applied to one object: the first\'s w:ilvl with the second\'s w:numId (c79-41)', async () => {
  for (const zip of [await saved('c79-41-ppr-numpr'), bytes('c79-41-ppr-numpr-word365')]) {
    assert.deepEqual(numPrOf(zip, 'LISTED WORDS'), { ilvl: '0', numId: '2' });
  }
});

const plainText = async (zip) => PlainText.of(await Texts.of(await OpcPackage.load(zip), TextOptions.of({}))).trim();

test('the words are the words Word read, by the text model; the mc roots (phase 1) pinned', async () => {
  const files = ['rc-', 'rt-', 'nd-2', 'c79-3'].flatMap(names).filter(hasSave);
  assert.equal(files.length, 45);
  for (const name of files) {
    const ours = await plainText(bytes(name));
    const theirs = await plainText(bytes(`${name}-word365`));
    if (MC_ROOTS.includes(name)) {
      // CR-028 section 3.1 isn't built: the part reads empty, where Word reads the branch it takes
      assert.equal(ours, '', name);
      assert.notEqual(theirs, '', name);
    } else {
      assert.equal(ours, theirs, name);
    }
  }
});

test('the text model reads every other source with a save as it reads Word\'s, headers, footers and list labels included', async () => {
  const options = () => TextOptions.of({ include: ['headers', 'footers'], list_labels: true });
  const everything = async (zip) => PlainText.of(await Texts.of(await OpcPackage.load(zip), options())).trim();
  const files = SOURCES.filter(hasSave).filter((n) => !MC_ROOTS.includes(n));
  assert.equal(files.length, 104);
  const failed = [];
  for (const name of files) if (await everything(bytes(name)) !== await everything(bytes(`${name}-word365`))) failed.push(name);
  assert.deepEqual(failed, []);
});

test('a save holds Word\'s sections: the joined bodies\', a nested body\'s where it stands last, a paragraph\'s where it is', async () => {
  let compared = 0;
  for (const name of ['rc-1', 'rt-1', 'nd-2'].flatMap(names)) {
    const ours = root(await saved(name));
    const theirs = root(bytes(`${name}-word365`));
    if (pageWidths(ours).length === 0) continue; // a file that states no section has none, where Word's save writes a default
    if (name in NESTED_SECTIONS) {
      // CR-028 D12, open: a nested w:body is kept as it stands, its w:sectPr with it, where Word applies each nested body's
      // w:sectPr to the last as an earlier copy (nd-22's landscape kept on the outer body's A4 size) and saves one
      assert.deepEqual(pageWidths(ours), NESTED_SECTIONS[name], name);
      assert.deepEqual(pageWidths(theirs), ['11906'], name);
      continue;
    }
    assert.deepEqual(pageWidths(ours), pageWidths(theirs), name);
    assert.deepEqual(all(ours, 'sectPr').map(stated), all(theirs, 'sectPr').map(stated), name);
    compared++;
  }
  assert.ok(compared >= 8, String(compared));
});

test('a table\'s width stays the later copy\'s (sm-61), and none where the later copy states none', async () => {
  const widths = all(root(await saved('sm-61-tblpr-attributes')), 'tblW');
  assert.deepEqual(attributes(widths.at(-1)), { type: 'auto' });
  const table = '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblPr><w:jc w:val="center"/></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const pkg = await OpcPackage.load(withPart('rc-00-control', 'word/document.xml', `<w:document xmlns:w="${W}"><w:body>${table}<w:p/></w:body></w:document>`));
  const tbl = (await pkg.getMainDocumentPart().getContents()).body.content[0].value;
  assert.equal(tbl.tblPr.tblW, undefined);
  assert.equal(tbl.tblPr.jc.val, 'center');
});

test('a save holds what Word\'s save holds: the sectpr-merge round (sm-61 left out, as docx4j and python leave it)', async () => {
  const failed = [];
  for (const name of names('sm-').filter((n) => !n.startsWith('sm-61') && !MC_ROOTS.includes(n))) {
    const ours = merged(root(await saved(name)));
    const theirs = merged(root(bytes(`${name}-word365`)));
    if (JSON.stringify(ours) !== JSON.stringify(theirs)) failed.push(`${name}\n  core-ts ${JSON.stringify(ours)}\n  Word    ${JSON.stringify(theirs)}`);
  }
  assert.deepEqual(failed, []);
});

test('a child given twice in one parent merges as two copies of the parent do (op-*)', async () => {
  for (const name of ['op-10-pgsz-in-one-sectpr', 'op-11-pgmar-in-one-sectpr', 'op-12-cols-in-one-sectpr']) {
    assert.deepEqual(all(root(await saved(name)), 'sectPr').map(stated), all(root(bytes(`${name}-word365`)), 'sectPr').map(stated), name);
  }
  for (const zip of [await saved('op-20-numpr-in-one-ppr'), bytes('op-20-numpr-in-one-ppr-word365')]) {
    const numPr = all(root(zip), 'numPr')[0];
    assert.deepEqual(Object.fromEntries(children(numPr).map((c) => [c.localName, c.getAttributeNS(W, 'val')])), { ilvl: '1', numId: '2' });
  }
  const markRPr = (zip) => {
    let found;
    for (const pPr of all(root(zip), 'pPr')) for (const c of children(pPr)) if (c.localName === 'rPr') found = stated(c);
    return found;
  };
  const ours = markRPr(await saved('op-21-rpr-in-one-ppr'));
  assert.deepEqual(ours, markRPr(bytes('op-21-rpr-in-one-ppr-word365')));
  assert.ok('b' in ours, JSON.stringify(ours));
});

test('only the first copy\'s references stand, of any w:type, across joined bodies too; a first copy with none leaves none', async () => {
  const files = names('ft-');
  assert.equal(files.length, 12);
  for (const name of files) assert.deepEqual(references(await saved(name)), references(bytes(`${name}-word365`)), name);
});

test('a part only a later copy\'s reference named stays in the package, where Word\'s save drops it', async () => {
  const parts = (zip) => Object.keys(unzipSync(zip)).filter((k) => /^word\/(header|footer)\d*\.xml$/.test(k)).sort();
  for (const [name, ours, word] of [['ft-02-footer-later-only', ['word/footer1.xml'], []],
    ['ft-03-footer-both', ['word/footer1.xml', 'word/footer2.xml'], ['word/footer1.xml']],
    ['sm-41-header-last', ['word/header1.xml', 'word/header2.xml'], []]]) {
    assert.deepEqual(parts(await saved(name)), ours, name);
    assert.deepEqual(parts(bytes(`${name}-word365`)), word, name);
  }
});

// --- departures from docx4j's test ----------------------------------------------------------------------

test('a footer\'s paragraph keeps its first w:pPr whole (objects CR-022 D7 (c), Word\'s hf-* rounds), where docx4j keeps the last', async () => {
  const footer = `<w:ftr xmlns:w="${W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:pPr><w:ind w:left="720"/></w:pPr><w:r><w:t>FOOT</w:t></w:r></w:p></w:ftr>`;
  const pkg = await OpcPackage.load(withPart('ft-01-footer-first-only', 'word/footer1.xml', footer));
  const pPr = (await pkg.parts.get('/word/footer1.xml').getContents()).content[0].value.pPr;
  assert.equal(pPr.jc.val, 'center');
  assert.equal(pPr.ind, undefined);
});

test('nothing is said of a join yet (CR-028 D8, the code open with docx4j and python; docx4j logs one WARN a part)', async () => {
  const pkg = await OpcPackage.load(bytes('rt-21-pPr-merge'));
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  assert.deepEqual(main.warnings, []);
  assert.deepEqual(pkg.warnings, []);
});

// --- this package's own ---------------------------------------------------------------------------------

/**
 * What this package's readers report of the paragraph holding `words`: Paragraph's and Font's effective reads. Word's save
 * adds a styles part with its own defaults, so Word's main part is read in the source's package, beside the source's parts.
 */
async function readers(zip, words) {
  const pkg = await OpcPackage.load(zip);
  const body = await pkg.getBody();
  const p = body.paragraphs.find((x) => x.text.includes(words));
  assert.ok(p, words);
  const f = p.font;
  return { alignment: p.alignment, leftIndent: p.leftIndent, rightIndent: p.rightIndent, firstLineIndent: p.firstLineIndent,
    spaceBefore: p.spaceBefore, spaceAfter: p.spaceAfter, bold: f.bold, italic: f.italic, size: f.size, name: f.name, color: f.color };
}

test('Paragraph and Font read each source as they read Word\'s save', async () => {
  for (const name of PROPERTIES) {
    const words = withPart(name, 'word/document.xml', entry(bytes(`${name}-word365`), 'word/document.xml'));
    assert.deepEqual(await readers(bytes(name), 'TARGET TEXT'), await readers(words, 'TARGET TEXT'), name);
  }
});

test('Body.text reads each source as it reads Word\'s save; an mc root refused (D10 as worded)', async () => {
  for (const name of SOURCES.filter(hasSave)) {
    const ours = await OpcPackage.load(bytes(name));
    const theirs = await OpcPackage.load(bytes(`${name}-word365`));
    if (MC_ROOTS.includes(name)) {
      await assert.rejects(ours.getBody(), name);
      continue;
    }
    assert.equal((await ours.getBody()).text, (await theirs.getBody()).text, name);
  }
});

test('every source: unchanged since its read, nothing skipped, a save unrefused that reads back alike', async () => {
  for (const name of SOURCES.filter(hasSave).filter((n) => !MC_ROOTS.includes(n))) {
    const pkg = await OpcPackage.load(bytes(name));
    await pkg.unmarshalAll();
    // CR-028 section 3.2's third trap: the read and isUnchangedSinceRead's re-read both join
    for (const part of pkg.parts.values()) {
      if (typeof part.isUnchangedSinceRead === 'function') assert.equal(await part.isUnchangedSinceRead(), true, `${name} ${part.partName}`);
    }
    assert.deepEqual(pkg.warnings.filter((w) => w.code === 'part.skipped'), [], name);
    const text = await plainText(bytes(name));
    const once = await pkg.save();
    assert.equal(await plainText(once), text, name);
    // the save writes one container, so a second read and save changes nothing more
    const again = await OpcPackage.load(once);
    await again.unmarshalAll();
    assert.equal(entry(await again.save(), 'word/document.xml'), entry(once, 'word/document.xml'), name);
  }
});

test('every load route reads alike: getContents, unmarshalAll, flat OPC, mcePreprocess, and a resolved Texts.of', async () => {
  const expected = await plainText(bytes('rc-10-body-word365'));
  assert.equal(await plainText(bytes('rc-10-body')), expected);
  const resolved = await OpcPackage.load(bytes('rc-10-body'), { mcePreprocess: true });
  assert.equal(PlainText.of(await Texts.of(resolved, TextOptions.of({}))).trim(), expected);
  const all = await OpcPackage.load(bytes('rc-10-body'));
  await all.unmarshalAll();
  assert.equal(PlainText.of(await Texts.of(all, TextOptions.of({}))).trim(), expected);
  const flat = await OpcPackage.load(await (await OpcPackage.load(bytes('rc-10-body'))).saveFlatOpc());
  assert.equal(PlainText.of(await Texts.of(flat, TextOptions.of({}))).trim(), expected);
  for (const name of ['rc-11-sdtContent', 'rc-12-txbxContent', 'rt-13-txbxContent-four']) {
    const flatOf = await OpcPackage.load(await (await OpcPackage.load(bytes(name))).saveFlatOpc());
    assert.equal(PlainText.of(await Texts.of(flatOf, TextOptions.of({}))).trim(), await plainText(bytes(`${name}-word365`)), name);
  }
});

test('a write through Body on two joined bodies is kept, with both bodies\' words (rc-10)', async () => {
  const pkg = await OpcPackage.load(bytes('rc-10-body'));
  const body = await pkg.getBody();
  const before = body.text;
  body.insertParagraph('ADDED WORDS', 'End');
  const zip = await pkg.save();
  const text = await plainText(zip);
  assert.ok(text.endsWith('ADDED WORDS'), text);
  assert.equal(text, `${await plainText(bytes('rc-10-body-word365'))}\nADDED WORDS`);
  assert.equal(all(root(zip), 'body').length, 1);
  assert.ok(before.length > 0);
});

test('insertOoxml reads a fragment\'s repeated w:pPr as one, merged (CR-028 section 3.5)', async () => {
  const pkg = await OpcPackage.load(bytes('rc-00-control'));
  const body = await pkg.getBody();
  // two paragraphs, since one is placed inline in the paragraph it goes into, as Word places it (CR-005 section 11.7)
  const [p] = await body.insertOoxml('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:pPr><w:ind w:left="720"/></w:pPr><w:r><w:t>INSERTED</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>AFTER</w:t></w:r></w:p>', 'After', body.paragraphs[0]);
  assert.equal(p.text, 'INSERTED');
  assert.equal(p.p.pPr.jc.val, 'center');
  assert.equal(p.p.pPr.ind.left, 720);
  assert.equal(p.alignment, 'Centered');
  assert.equal(p.leftIndent, 36);
});
