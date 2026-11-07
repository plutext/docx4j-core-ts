// CR-028 D10 and D11: a main document part whose root isn't w:document, read as WD REQ-318 to REQ-320 have it. Word's files are in
// test/main-part-root/, copied unedited with a SHA-1 manifest: each root Word reads is held to Word's text by both load routes and
// written as a w:document, each Word refuses is refused, and any other reads as an empty document with part.unexpected_root, its
// writes refused and a save copying it as written.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strToU8, strFromU8, unzipSync } from 'fflate';
import {
  OpcPackage, Texts, PlainText, Docx4JException, PartUnmarshalException, SaveLossException, parseXml, readMainPartRoot,
} from '../dist/index.mjs';
import { storedZip } from './helpers.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, 'main-part-root');
const manifest = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8'));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const ROUTES = [{}, { mcePreprocess: true }];

// What each root reads as, by its statement: REQ-319 read, REQ-318 refused, REQ-320 empty. Held here, not taken from the pass.
const READ = ['or-31-madeup', 'or-32-body', 'mr-01-madeup-p', 'mr-10-two-p', 'mr-11-p-sectpr', 'mr-12-tbl', 'mr-13-sdt', 'mr-15-body',
  'mr-16-empty', 'mr-17-nested', 'mr-18-p-properties', 'mr-30-sdtcontent-p', 'mr-32-p-root', 'mr-50-madeup-p-styles', 'mu-12-foo-holding-p',
  'rq-10-madeup-ac', 'rq-20-sdtcontent-empty'];
const REFUSED = ['or-12-footnotes', 'or-13-endnotes', 'or-15-footnotes-bare', 'mr-03-foo-document', 'mr-14-run', 'mr-20-other-namespace-p',
  'mr-21-no-namespace-p', 'mr-31-tc-p', 'mu-10-madeup-one-document', 'mu-11-madeup-two-documents', 'rw-10-foo-two-documents',
  'rw-11-zfoo-two-documents', 'rw-12-wfoo-two-documents', 'rw-13-foo-one-document', 'rq-11-madeup-sdt-runs', 'rq-21-sdtcontent-rows',
  'rq-22-sdtcontent-row-sdt'];
const EMPTY = { 'or-10-hdr': 'w:hdr', 'or-11-ftr': 'w:ftr', 'or-14-comments': 'w:comments', 'or-30-styles': 'w:styles', 'mr-22-ignorable-p': 'z:madeUp' };

const bytesOf = (name) => new Uint8Array(readFileSync(join(dir, name)));
const norm = (text) => text.replace(/\s+/g, ' ').trim();
const mainXml = (bytes) => strFromU8(unzipSync(bytes)['word/document.xml']);
const load = async (bytes, options = {}) => {
  const pkg = await OpcPackage.load(bytes, options);
  await pkg.unmarshalAll();
  return pkg;
};
const textOf = async (pkg) => {
  const result = await Texts.of(pkg, { include: ['headers', 'footers'], comments: true });
  return { text: norm(PlainText.of(result)), warnings: result.warnings };
};
const docx = (root) => storedZip([
  ['[Content_Types].xml', `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN}"/></Types>`],
  ['_rels/.rels', `${XML}<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
  ['word/document.xml', `${XML}${root}`],
].map(([name, value]) => [name, strToU8(value)]));
const run = (text) => `<w:r><w:t>${text}</w:t></w:r>`;

test('CR-028 D11: the files are as the manifest has them, and the table names every one', () => {
  for (const f of manifest.files) assert.equal(createHash('sha1').update(bytesOf(f.name)).digest('hex'), f.sha1, f.name);
  const stems = Object.keys(manifest.stems).filter((s) => s !== 'rq-00-control').sort();
  assert.deepEqual([...READ, ...REFUSED, ...Object.keys(EMPTY)].sort(), stems);
  // the statement follows Word: each refused root is one Word refused, and each other one Word opened
  for (const s of REFUSED) assert.equal(manifest.stems[s].word, 'refused', s);
  for (const s of [...READ, ...Object.keys(EMPTY)]) assert.notEqual(manifest.stems[s].word, 'refused', s);
});

test('CR-028 D11: REQ-319, each root Word reads gives Word\'s text by both load routes, and nothing more is said of it', async () => {
  for (const stem of READ) {
    const word = await textOf(await load(bytesOf(`${stem}-word365.docx`)));
    for (const options of ROUTES) {
      const pkg = await load(bytesOf(`${stem}.docx`), options);
      const { text, warnings } = await textOf(pkg);
      assert.equal(text, word.text, `${stem} ${JSON.stringify(options)}`);
      assert.equal(pkg.getMainDocumentPart().unexpectedRoot, undefined, stem);
      assert.ok(!warnings.some((w) => w.kind === 'part.unexpected_root'), stem);
      assert.equal(norm(pkg.body.text), norm((await load(bytesOf(`${stem}-word365.docx`))).body.text), `${stem}: Body.text`);
    }
  }
});

test('CR-028 D11: REQ-319, a root read is written as a w:document, and a write through Body is kept', async () => {
  for (const stem of READ) {
    const pkg = await load(bytesOf(`${stem}.docx`));
    const skipped = pkg.getMainDocumentPart().skippedByRead;
    pkg.body.insertParagraph('ADDED', 'End');
    if (skipped.length > 0) {
      // what the read dropped (a w:sectPr standing in the root, an unknown element read through) is CR-025's to guard
      await assert.rejects(pkg.save(), SaveLossException, stem);
      pkg.checkLossOnSave = false;
    }
    const saved = await pkg.save();
    assert.match(mainXml(saved), /^<\?xml[^>]*\?>\s*<w:document[\s>]/, stem);
    assert.match((await textOf(await load(saved))).text, /ADDED$/, stem);
  }
});

test('CR-028 D11: REQ-318, each root Word refuses is refused at the read, by both load routes, naming the root', async () => {
  for (const stem of REFUSED) {
    for (const options of ROUTES) {
      const pkg = await OpcPackage.load(bytesOf(`${stem}.docx`), options);
      await assert.rejects(pkg.unmarshalAll(), (e) => e instanceof PartUnmarshalException && /REQ-318/.test(e.message)
        && e.message.includes(manifest.stems[stem].root), `${stem} ${JSON.stringify(options)}`);
      await assert.rejects(Texts.of(await OpcPackage.load(bytesOf(`${stem}.docx`), options)), PartUnmarshalException, stem);
    }
  }
});

test('CR-028 D11: REQ-320, any other root reads as an empty document with part.unexpected_root, on the part and in the text model', async () => {
  for (const [stem, root] of Object.entries(EMPTY)) {
    for (const options of ROUTES) {
      const pkg = await load(bytesOf(`${stem}.docx`), options);
      const { text, warnings } = await textOf(pkg);
      assert.equal(text, '', stem);
      assert.deepEqual(pkg.warnings.filter((w) => w.code === 'part.unexpected_root').map((w) => [w.partName, w.root]), [['/word/document.xml', root]], stem);
      assert.deepEqual(warnings.filter((w) => w.kind === 'part.unexpected_root').map((w) => [w.selector, w.detail]), [['/word/document.xml', root]], stem);
      assert.equal(pkg.getMainDocumentPart().unexpectedRoot, root, stem);
    }
    // where Word saved it, it shows the same nothing; or-14's comment Word keeps through its repair, and the statement doesn't
    if (existsSync(join(dir, `${stem}-word365.docx`)) && stem !== 'or-14-comments') {
      assert.equal((await textOf(await load(bytesOf(`${stem}-word365.docx`)))).text, '', stem);
    }
  }
});

test('CR-028 D10: on such a root a write through Body is refused, a save copies the part as written, and a changed one is refused', async () => {
  for (const stem of Object.keys(EMPTY)) {
    const source = bytesOf(`${stem}.docx`);
    const pkg = await load(source);
    assert.throws(() => pkg.body, (e) => e instanceof Docx4JException && /REQ-320/.test(e.message), stem);
    assert.equal(mainXml(await pkg.save()), mainXml(source), `${stem}: copied as written`);
    pkg.getMainDocumentPart().contents.body.content = [{ name: { namespaceURI: W, localPart: 'p' }, value: { TYPE_NAME: 'org_docx4j_wml.P' } }];
    await assert.rejects(pkg.save(), SaveLossException, stem);
  }
});

test('CR-028 D11: contents a caller sets are a document like any other', async () => {
  const pkg = await load(bytesOf('or-10-hdr.docx'));
  const main = pkg.getMainDocumentPart();
  main.setContents({ TYPE_NAME: 'org_docx4j_wml.Document', body: { TYPE_NAME: 'org_docx4j_wml.Body' } });
  assert.equal(main.unexpectedRoot, undefined);
  pkg.body.insertParagraph('SET', 'End');
  assert.equal((await textOf(await load(await pkg.save()))).text, 'SET');
});

test('CR-028 D11: the control and every ordinary root pass untouched', async () => {
  const source = bytesOf('rq-00-control.docx');
  const pkg = await load(source);
  assert.deepEqual(pkg.warnings, []);
  assert.equal((await textOf(pkg)).text, (await textOf(await load(bytesOf('rq-00-control-word365.docx')))).text);
  const doc = parseXml(mainXml(source));
  const before = doc.documentElement;
  assert.deepEqual(readMainPartRoot(doc, '/word/document.xml'), []);
  assert.equal(doc.documentElement, before);
});

test('CR-028 D11: the roots REQ-318 names that no file here holds are refused too', () => {
  const roots = [
    `<w:txbxContent xmlns:w="${W}"><w:p/></w:txbxContent>`,
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"/>`,
    `<w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml"/>`,
    `<w:endnotes xmlns:w="${W}"/>`,
    `<w:madeUp xmlns:w="${W}"><w:document/></w:madeUp>`,
    `<z:madeUp xmlns:z="urn:z" xmlns:w="${W}"><w:document/></z:madeUp>`,
  ];
  for (const xml of roots) assert.throws(() => readMainPartRoot(parseXml(xml), '/word/document.xml'), PartUnmarshalException, xml);
});

test('CR-028 D11: a w:sectPr standing in an unknown root has no effect (mr-11), and one in a w:body it holds is read (mr-15)', async () => {
  const sectionOf = async (stem) => (await load(bytesOf(`${stem}.docx`))).getMainDocumentPart().contents.body.sectPr;
  assert.equal(await sectionOf('mr-11-p-sectpr'), undefined);
  assert.notEqual(await sectionOf('mr-15-body'), undefined);
});

test('CR-028 D11: a root in a namespace Word understands that isn\'t w:\'s, holding a w:p, is any other root, not one refused', async () => {
  const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const pkg = await load(docx(`<w14:madeUp xmlns:w14="${W14}" xmlns:w="${W}"><w:p>${run('UNDERSTOOD')}</w:p></w14:madeUp>`));
  assert.equal((await textOf(pkg)).text, '');
  assert.equal(pkg.getMainDocumentPart().unexpectedRoot, 'w14:madeUp');
});

test('CR-028 D11: a block w:sdt as the root is a body holding it, and one at run level is any other root', async () => {
  const block = await load(docx(`<w:sdt xmlns:w="${W}"><w:sdtContent><w:p>${run('BLOCK')}</w:p></w:sdtContent></w:sdt>`));
  assert.equal((await textOf(block)).text, 'BLOCK');
  const runs = await load(docx(`<w:sdt xmlns:w="${W}"><w:sdtContent>${run('RUN')}</w:sdtContent></w:sdt>`));
  assert.equal((await textOf(runs)).text, '');
  assert.equal(runs.getMainDocumentPart().unexpectedRoot, 'w:sdt');
});

test('CR-028 D11: an unknown root holding a w:body reads that body, what stands beside it said to go', async () => {
  const pkg = await load(docx(`<w:madeUp xmlns:w="${W}"><w:p>${run('BESIDE')}</w:p><w:body><w:p>${run('BODY')}</w:p></w:body></w:madeUp>`));
  assert.equal((await textOf(pkg)).text, 'BODY');
  assert.deepEqual(pkg.getMainDocumentPart().skippedByRead.map((w) => [w.kind, w.name]), [['element', 'w:p']]);
});

test('CR-028 D11: the root\'s declarations and mc:Ignorable go onto the w:document, so nothing the root bound is lost', async () => {
  const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const pkg = await load(docx(`<w:body xmlns:w="${W}" xmlns:mc="${MC}" xmlns:w14="${W14}" mc:Ignorable="w14"><w:p w14:paraId="1A2B3C4D">${run('IDS')}</w:p></w:body>`));
  assert.deepEqual(pkg.warnings, []);
  const saved = mainXml(await pkg.save());
  assert.match(saved, /<w:document[^>]*mc:Ignorable="w14"/);
  assert.match(saved, /w14:paraId="1A2B3C4D"/);
  assert.doesNotMatch(saved, /<w:body[^>]*Ignorable/);
});

test('CR-028 D11: mc:AlternateContent as the root is REQ-317\'s, and left as it was', () => {
  const doc = parseXml(`<mc:AlternateContent xmlns:mc="${MC}" xmlns:w="${W}"><mc:Fallback><w:document><w:body/></w:document></mc:Fallback></mc:AlternateContent>`);
  const before = doc.documentElement;
  assert.deepEqual(readMainPartRoot(doc, '/word/document.xml'), []);
  assert.equal(doc.documentElement, before);
});
