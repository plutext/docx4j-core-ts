import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import { join } from 'node:path';
import { OpcPackage, mceKeep, mcePreprocess, WordprocessingMLPackage, HeaderPart, parseXml, resolveAlternateContent, preprocessAlternateContent, createMcePreprocessor, UNDERSTOOD_NAMESPACES, KEPT_HOSTS, MCE_NS, ZipPartStore, FlatOpcPartStore } from '../dist/index.mjs';
import { fixture, fixturesDir } from './helpers.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

test('resolveAlternateContent: the first understood mc:Choice wins, else the fallback', () => {
  const xml = `<w:p xmlns:w="${W}" xmlns:mc="${MCE_NS}" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:x="urn:unknown"><w:r><mc:AlternateContent><mc:Choice Requires="x"><w:t>unknown</w:t></mc:Choice><mc:Choice Requires="wps x"><w:t>partly</w:t></mc:Choice><mc:Choice Requires="wps"><w:t>choice</w:t></mc:Choice><mc:Fallback><w:t>fallback</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>`;
  let doc = parseXml(xml);
  assert.equal(resolveAlternateContent(doc), 1);
  assert.equal(doc.getElementsByTagNameNS(MCE_NS, 'AlternateContent').length, 0);
  let t = doc.getElementsByTagNameNS(W, 't');
  assert.equal(t.length, 1);
  assert.equal(t[0].textContent, 'choice');
  assert.equal(t[0].parentNode.localName, 'r');

  doc = parseXml(xml);
  resolveAlternateContent(doc, new Set());
  t = doc.getElementsByTagNameNS(W, 't');
  assert.equal(t.length, 1);
  assert.equal(t[0].textContent, 'fallback');

  // nested: inner resolved first
  const nested = `<w:p xmlns:w="${W}" xmlns:mc="${MCE_NS}" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><mc:AlternateContent><mc:Choice Requires="wps"><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>inner</w:t></mc:Choice></mc:AlternateContent></w:r></mc:Choice></mc:AlternateContent></w:p>`;
  doc = parseXml(nested);
  assert.equal(resolveAlternateContent(doc), 2);
  assert.equal(doc.getElementsByTagNameNS(W, 't')[0].textContent, 'inner');
  // no fallback and nothing understood: kept whole since CR-010 phase 1 (ISO removes it)
  doc = parseXml(`<w:p xmlns:w="${W}" xmlns:mc="${MCE_NS}" xmlns:x="urn:unknown"><mc:AlternateContent><mc:Choice Requires="x"><w:r/></mc:Choice></mc:AlternateContent></w:p>`);
  assert.equal(resolveAlternateContent(doc), 0);
  assert.equal(doc.getElementsByTagNameNS(MCE_NS, 'AlternateContent').length, 1);
  assert.equal(doc.getElementsByTagNameNS(W, 'r').length, 1);
  assert.ok(UNDERSTOOD_NAMESPACES.has('http://schemas.microsoft.com/office/word/2010/wordprocessingShape'));
});

test('a Word 2010 flat OPC package with mc:AlternateContent in a header', async () => {
  const flat = await readFile(join(fixturesDir, 'mc-alternate-content-header.xml'), 'utf8');
  // resolved (mcePreprocess: true): the wps choice is taken, a w:drawing, not the VML fallback
  const pkg = await OpcPackage.load(flat, { mcePreprocess: true });
  assert.ok(pkg instanceof WordprocessingMLPackage);
  const headers = pkg.getMainDocumentPart().headerParts;
  assert.ok(headers.length > 0);
  const header = headers.find((h) => h instanceof HeaderPart);
  const hdr = await header.getContents();
  assert.equal(hdr.TYPE_NAME, 'org_docx4j_wml.Hdr');
  const xml = await header.getXml();
  assert.ok(!xml.includes('mc:AlternateContent'));
  assert.ok(xml.includes('<w:drawing>') || xml.includes('<w:drawing '), 'choice branch');
  assert.ok(!xml.includes('<w:pict'), 'fallback branch dropped');
  // the whole package saves as a zip and reloads
  const zip = await pkg.save();
  const back = await OpcPackage.load(zip);
  assert.equal(back.parts.size, pkg.parts.size);

  // kept, the default since CR-010 phase 4: the header writes both branches back
  const kept = await OpcPackage.load(flat);
  const keptHeader = kept.getMainDocumentPart().headerParts.find((h) => h instanceof HeaderPart);
  await keptHeader.getContents();
  const keptXml = await keptHeader.getXml();
  assert.ok(keptXml.includes('<mc:AlternateContent>') && keptXml.includes('<mc:Choice Requires="wps">'), 'kept whole');
  assert.ok(keptXml.includes('<w:drawing>') || keptXml.includes('<w:drawing '), 'the Choice');
  assert.ok(keptXml.includes('<w:pict'), 'and the Fallback');
  assert.equal((await OpcPackage.load(await kept.save())).parts.size, kept.parts.size);
});

test('the fallback branch when nothing is understood; preprocessing off', async () => {
  const flat = await readFile(join(fixturesDir, 'mc-alternate-content-header.xml'), 'utf8');
  const fallbackPkg = await OpcPackage.load(flat, { preprocessor: createMcePreprocessor({ understood: new Set() }) });
  const header = fallbackPkg.getMainDocumentPart().headerParts.find((h) => h instanceof HeaderPart);
  await header.getContents();
  const xml = await header.getXml();
  assert.ok(xml.includes('<w:pict'), 'fallback branch');
  assert.ok(!xml.includes('<w:drawing'));

  // Off: since the CR-021 regeneration (objects 0.1.5) the model admits mc:AlternateContent in
  // the WordprocessingML hosts, so the part unmarshals with the element typed and both branches
  // kept, and a save writes both back. Since objects 0.4.0 (its CR-009, CR-010 phase 2) what the
  // branches hold is typed too; the content API reads kept branches since CR-010 phase 3, and
  // keeping is the default since phase 4, so the option and the default load are the same.
  const raw = await OpcPackage.load(await fixture('loadAndSave.docx'), { mcePreprocess: false });
  const contents = await raw.getMainDocumentPart().getContents();
  const alternates = [];
  const walk = (v) => {
    if (!v || typeof v !== 'object' || v.nodeType) return;
    if (v.TYPE_NAME === 'org_docx4j_mce.AlternateContent') alternates.push(v);
    for (const k of Object.keys(v)) if (k !== 'PARENT') walk(v[k]);
  };
  walk(contents);
  assert.equal(alternates.length, 1);
  assert.equal(alternates[0].choice.length, 1);
  assert.equal(alternates[0].choice[0].requires, 'wps');
  // Typed: the branches' lax wildcard types only elements the context declares globally, and
  // objects 0.4.0 (its CR-009) declares w:drawing and w:pict globally, as docx4j does. Until then
  // they were DOM. The same release binds w14:anchorId on VML shapes (docx4j CR-048), so the
  // Fallback's v:shapetype keeps it through a save.
  const drawing = alternates[0].choice[0].any[0];
  assert.equal(drawing.name.localPart, 'drawing');
  assert.equal(drawing.value.TYPE_NAME, 'org_docx4j_wml.Drawing');
  assert.equal(drawing.value.anchorOrInline[0].TYPE_NAME, 'org_docx4j_dml_wordprocessingDrawing.Anchor');
  const pict = alternates[0].fallback.any[0];
  assert.equal(pict.name.localPart, 'pict');
  assert.equal(pict.value.TYPE_NAME, 'org_docx4j_wml.Pict');
  const shapetype = pict.value.anyAndAny.find((item) => item.value?.TYPE_NAME === 'org_docx4j_vml.CTShapetype');
  assert.equal(shapetype.value.anchorId, '7969C50B');
  const source = new TextDecoder().decode(new ZipPartStore(await fixture('loadAndSave.docx')).loadSync('word/document.xml'));
  const saved = new TextDecoder().decode(new ZipPartStore(await raw.save()).loadSync('word/document.xml'));
  for (const tag of ['<mc:AlternateContent', '<mc:Choice', '<mc:Fallback', '<w:drawing', '<w:pict']) {
    assert.equal(saved.split(tag).length, source.split(tag).length, `${tag} kept`);
  }
  assert.ok(saved.includes('w14:anchorId="7969C50B"'), 'the v:shapetype keeps w14:anchorId');
});

test('a re-marshalled document declares the conventional prefixes and the mc:Ignorable ones', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('loadAndSave.docx'));
  await pkg.getMainDocumentPart().getContents();
  const store = new ZipPartStore(await pkg.save());
  const xml = new TextDecoder().decode(store.loadSync('word/document.xml'));
  assert.ok(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'));
  const root = xml.slice(xml.indexOf('<w:document'), xml.indexOf('>', xml.indexOf('<w:document')) + 1);
  const ignorable = /mc:Ignorable="([^"]*)"/.exec(root)[1].split(' ');
  assert.ok(ignorable.includes('w14'));
  for (const p of ignorable) assert.ok(root.includes(`xmlns:${p}="`), `xmlns:${p} declared on the root`);
  assert.ok(!/\bp\d+:/.test(xml), 'no generated prefixes');
  assert.ok(!xml.includes('xmlns:xml='));
  assert.ok(xml.includes('xml:space="preserve"'));
  assert.ok(xml.includes('w14:paraId='));
  const rels = new TextDecoder().decode(store.loadSync('word/_rels/document.xml.rels'));
  assert.ok(rels.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship '));
});

// UNDERSTOOD_NAMESPACES is hand-kept (`src/opc/mce/understood.mts`); this is what announces the
// next objects regeneration, so that a module added there is weighed here rather than silently
// left out of the Choice rule. Every module's defaultElementNamespaceURI must be in the set - and
// the set may hold more, which the file marks and explains.
test('every namespace the object model has a module for is understood', async () => {
  const dir = new URL('../node_modules/@docx4j/generated-objects-ts/modules/', import.meta.url);
  const missing = [];
  for (const file of (await readdir(dir)).sort()) {
    if (!file.endsWith('.mjs') || file.includes('.el.') || file.includes('.factory.')) continue;
    const source = await readFile(new URL(file, dir), 'utf8');
    // the compiler writes it as a JS string literal, with the slashes escaped
    const ns = /defaultElementNamespaceURI:\s*'((?:[^'\\]|\\.)*)'/.exec(source)?.[1]?.replace(/\\\//g, '/');
    if (ns && !UNDERSTOOD_NAMESPACES.has(ns)) missing.push(`${file}: ${ns}`);
  }
  assert.deepEqual(missing, [], `add these to UNDERSTOOD_NAMESPACES, or record why not:\n${missing.join('\n')}`);
});

// CR-009 section 8: Office gates a chartex chart's frame with a later chartex namespace than the
// chart's own - Requires="cx1" (2015/9/8) for a waterfall or a histogram, "cx2" (2015/10/21) for a
// funnel - though nothing in the Choice is in it. Before they were understood the Fallback was
// taken, and a save that re-marshalled the part wrote the chart as its fallback picture (Word,
// PowerPoint) or as Excel's "not available" box. docx4j's three waterfalls (its charts/ fixtures).
const elementForm = (n) => {
  if (n.nodeType === 3) return n.nodeValue.trim() ? JSON.stringify(n.nodeValue) : '';
  if (n.nodeType !== 1) return '';
  const attrs = [...n.attributes].filter((a) => a.name !== 'xmlns' && !a.name.startsWith('xmlns:'))
    .map((a) => `{${a.namespaceURI ?? ''}}${a.localName}=${JSON.stringify(a.value)}`).sort();
  return `<{${n.namespaceURI}}${n.localName} ${attrs.join(' ')}>${[...n.childNodes].map(elementForm).join('')}</>`;
};
const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
for (const [file, name] of [['waterfall.docx', 'word/document.xml'], ['waterfall.xlsx', 'xl/drawings/drawing1.xml'], ['waterfall.pptx', 'ppt/slides/slide1.xml']]) {
  test(`a chartex frame behind Requires="cx1" keeps its chart through a save: charts/${file}`, async () => {
    const bytes = await fixture(`charts/${file}`);
    const source = parseXml(new TextDecoder().decode(new ZipPartStore(bytes).loadSync(name)));
    const choice = source.getElementsByTagNameNS(MCE_NS, 'Choice')[0];
    assert.equal(choice.getAttribute('Requires'), 'cx1');
    const frame = [...choice.childNodes].find((c) => c.nodeType === 1);

    // resolved (mcePreprocess: true), the Choice is taken; kept (the default since CR-010 phase 4),
    // the element is written back whole, the Fallback's picture with it
    for (const resolved of [true, false]) {
      const pkg = await OpcPackage.load(bytes, { mcePreprocess: resolved });
      await pkg.unmarshalAll();
      const store = new ZipPartStore(await pkg.save());
      const saved = parseXml(new TextDecoder().decode(store.loadSync(name)));
      // canonical (below) leaves aside the lexical rewrites, xsd:boolean's 1 written true among them
      const elements = [...saved.getElementsByTagNameNS(MCE_NS, 'AlternateContent')].map(canonical);
      if (resolved) assert.equal(elements.length, 0, 'the branch was chosen on load');
      else assert.deepEqual(elements, [...source.getElementsByTagNameNS(MCE_NS, 'AlternateContent')].map(canonical), 'kept whole');
      const frames = [...saved.getElementsByTagNameNS(frame.namespaceURI, frame.localName)].map(elementForm);
      assert.ok(frames.includes(elementForm(frame)), `the Choice's ${frame.nodeName} is written back whole`);
      const chart = saved.getElementsByTagNameNS(CX, 'chart');
      assert.equal(chart.length, 1, 'cx:chart, once: the Fallback holds none');
      const rels = store.loadSync(name.replace(/([^/]+)$/, '_rels/$1.rels'));
      const id = chart[0].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      assert.match(new TextDecoder().decode(rels), new RegExp(`Id="${id}"[^>]*chartEx|chartEx[^>]*Id="${id}"`), 'its relationship still names the chartEx part');
    }
  });
}

test('a funnel\'s frame, behind Requires="cx2", is taken too', () => {
  const doc = parseXml(`<w:r xmlns:w="${W}" xmlns:mc="${MCE_NS}" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"><mc:AlternateContent><mc:Choice Requires="cx2"><w:drawing><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="${CX}"><cx:chart xmlns:cx="${CX}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId4"/></a:graphicData></a:graphic></w:drawing></mc:Choice><mc:Fallback><w:drawing><w:t>picture</w:t></w:drawing></mc:Fallback></mc:AlternateContent></w:r>`);
  assert.equal(resolveAlternateContent(doc), 1);
  assert.equal(doc.getElementsByTagNameNS(CX, 'chart').length, 1);
  assert.equal(doc.getElementsByTagNameNS(W, 't').length, 0);
});

// CR-010 phase 1 (Jason, 2026-10-17: D3, phase 1 now, with its two additions). The element is kept
// whole under c:chartSpace and the SpreadsheetML workbook, where promoting a Choice lost it outright
// (CR-001 sections 21 and 21.1): a chart's c14:style went, and with it the Fallback's c:style.
test('a chart part keeps its c14:style mc:AlternateContent, and its c:style Fallback, through a save', async () => {
  const bytes = await fixture('loadAndSave.docx');
  const name = 'word/charts/chart1.xml';
  const source = parseXml(new TextDecoder().decode(new ZipPartStore(bytes).loadSync(name)));
  const want = [...source.getElementsByTagNameNS(MCE_NS, 'AlternateContent')].map(elementForm);
  assert.equal(want.length, 1);
  assert.equal(source.getElementsByTagNameNS(MCE_NS, 'AlternateContent')[0].parentNode.localName, 'chartSpace');
  const pkg = await OpcPackage.load(bytes);
  await pkg.unmarshalAll();
  const chart = [...pkg.parts].find((p) => p.partName.name === '/' + name);
  const kept = chart.contents.alternateContent;
  assert.equal(kept.choice[0].requires, 'c14');
  assert.ok(kept.fallback, 'the Fallback is in the tree');
  const saved = parseXml(new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync(name)));
  assert.deepEqual([...saved.getElementsByTagNameNS(MCE_NS, 'AlternateContent')].map(elementForm), want, 'written back whole');
  assert.equal(saved.getElementsByTagNameNS('http://schemas.microsoft.com/office/drawing/2007/8/2/chart', 'style')[0].getAttribute('val'), '102');
  assert.equal(saved.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/chart', 'style')[0].getAttribute('val'), '2');
  assert.equal(pkg.warnings.filter((w) => w.partName === '/' + name).length, 0, 'nothing resolved there, so nothing to warn about');
});

test('KEPT_HOSTS, and keptHosts: an empty set resolves everywhere, as before', () => {
  assert.deepEqual([...KEPT_HOSTS].sort(), [
    '{http://schemas.openxmlformats.org/drawingml/2006/chart}chartSpace',
    '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}workbook',
  ]);
  const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="${MCE_NS}"><mc:AlternateContent><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent></c:chartSpace>`;
  assert.deepEqual(preprocessAlternateContent(parseXml(xml)), { resolved: 0, kept: 1, warnings: [] });
  const doc = parseXml(xml);
  assert.equal(preprocessAlternateContent(doc, { keptHosts: new Set() }).resolved, 1);
  assert.equal(doc.documentElement.firstChild.localName, 'style');
  assert.equal(doc.documentElement.firstChild.namespaceURI, 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart');
});

// The text extraction paper's D4 item 4 (decided 2026-10-17): a reader takes the first Choice of an
// element nothing can be selected from. Load does not promote it - its content is in no namespace
// the model understands, and a w:r refuses it, so the part would not open - and does not remove it,
// as ISO would: it keeps it whole, for a reader over kept branches (CR-010 phase 3) to read.
test('an element nothing can be selected from is kept whole, and a save writes it back', async () => {
  const box = '<zz:box><w:txbxContent><w:p><w:r><w:t>Hello box</w:t></w:r></w:p></w:txbxContent></zz:box>';
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}" xmlns:mc="${MCE_NS}" xmlns:zz="urn:example:zz"><w:body><w:p><w:r><w:t>Before </w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="zz">${box}</mc:Choice></mc:AlternateContent></w:r><w:r><w:t>after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const back = await WordprocessingMLPackage.load(await pkg.save());
  const body = await back.getBody();
  assert.equal(body.paragraphs[0].text, 'Before after', 'the content API does not read a kept branch yet');
  body.paragraphs[0].insertText('!', 'End');
  const xml = new TextDecoder().decode(new ZipPartStore(await back.save()).loadSync('word/document.xml'));
  assert.ok(xml.includes('<mc:AlternateContent><mc:Choice Requires="zz">'), 'kept');
  assert.ok(xml.includes('Hello box'));
  assert.ok(xml.includes('after!'), 'and the edit');
});

// The text extraction paper's section 4.1 item 6 (decided 2026-10-17), in its load-time form: the
// preprocessor compares the text of the branch it takes with the branches it drops, and records a
// warning on the part when they differ.
test('mce.branches_differ: a text box whose Fallback says something else', async () => {
  const wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const run = (choice, fallback) => `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><w:txbxContent><w:p><w:r><w:t>${choice}</w:t></w:r></w:p></w:txbxContent></w:drawing></mc:Choice><mc:Fallback><w:pict><w:txbxContent><w:p><w:r><w:t>${fallback}</w:t></w:r></w:p></w:txbxContent></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}" xmlns:mc="${MCE_NS}" xmlns:wps="${wps}"><w:body><w:p>${run('Current text', 'Stale  text')}${run('Same', ' Same ')}</w:p><w:sectPr/></w:body></w:document>`);
  const back = await WordprocessingMLPackage.load(await pkg.save());
  assert.deepEqual(back.warnings, [], 'nothing parsed yet');
  await back.getBody();
  const warnings = back.getMainDocumentPart().warnings;
  assert.equal(warnings.length, 1, 'whitespace aside, the second box\'s branches agree');
  const [w] = warnings;
  assert.equal(w.code, 'mce.branches_differ');
  assert.equal(w.partName, '/word/document.xml');
  assert.equal(w.selected, 'Choice(wps)');
  assert.deepEqual(w.branches, [{ branch: 'Choice(wps)', text: 'Current text' }, { branch: 'Fallback', text: 'Stale text' }]);
  assert.equal(typeof w.line, 'number');
  assert.equal(typeof w.column, 'number');
  assert.match(w.message, /^\/word\/document\.xml \(line \d+, column \d+\): .*Choice\(wps\) was read \("Current text"\), not Fallback "Stale text"$/);
  assert.deepEqual(back.warnings, warnings, 'the package gathers its parts\'');
  back.getMainDocumentPart().setContents(back.getMainDocumentPart().contents);
  assert.deepEqual(back.getMainDocumentPart().warnings, [], 'contents set: the source\'s warnings no longer apply');
});

test('mce.branches_differ: a private read reports too, and an untouched part stays untouched', async () => {
  // docx4j's waterfall: Excel's Fallback is its "not available" box, with text; the cx1 frame has none.
  const pkg = await OpcPackage.load(await fixture('charts/waterfall.xlsx'));
  const drawing = [...pkg.parts].find((p) => p.partName.name === '/xl/drawings/drawing1.xml');
  await drawing.readContents();
  assert.ok(drawing.isUntouched);
  assert.equal(drawing.warnings.length, 1);
  assert.equal(drawing.warnings[0].selected, 'Choice(cx1)');
  assert.equal(drawing.warnings[0].branches[0].text, '');
  assert.match(drawing.warnings[0].branches[1].text, /^This chart isn't available in your version of Excel/);
});

// CR-010 phase 0 (section 11): the survey of sections 2, 10.2 and 10.4 as a test, over this
// repository's fixtures; the release checks add docx4j's charts directory, which is not here.
// Loaded with every mc:AlternateContent kept, unmarshalled and saved, each part that held the
// element writes every one back whole, and what a kept branch holds in WordprocessingML is typed.
// Elements are compared by names, attributes and children, leaving aside prefix declarations and
// the three lexical rewrites the model makes anywhere (section 10.4): xsd:boolean's 1 and 0
// written true and false, o:gfxdata's base64 without its line breaks, and whitespace-only text.

const O_NS = 'urn:schemas-microsoft-com:office:office';
const XML_PACKAGE_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';

function canonical(node) {
  if (node.nodeType === 3 || node.nodeType === 4) return node.nodeValue.trim() ? JSON.stringify(node.nodeValue) : '';
  if (node.nodeType !== 1) return '';
  const attributes = [...node.attributes].filter((a) => a.name !== 'xmlns' && !a.name.startsWith('xmlns:')).map((a) => {
    const value = a.namespaceURI === O_NS && a.localName === 'gfxdata' ? a.value.replace(/\s+/g, '')
      : a.value === 'true' ? '1' : a.value === 'false' ? '0' : a.value;
    return `{${a.namespaceURI ?? ''}}${a.localName}=${JSON.stringify(value)}`;
  }).sort();
  return `<{${node.namespaceURI}}${node.localName} ${attributes.join(' ')}>${[...node.childNodes].map(canonical).join('')}</>`;
}

/** Every zip package and flat OPC package under a directory, sorted. */
async function packagesUnder(dir) {
  const out = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) out.push(...await packagesUnder(path));
    else if (/\.(docx|xlsx|pptx|xml)$/i.test(entry.name)) out.push(path);
  }
  return out.sort();
}

/** Adds, by qualified name, what the branches of each mc:AlternateContent under `value` hold: typed or DOM. */
function branchContents(value, into, seen = new Set()) {
  if (!value || typeof value !== 'object' || seen.has(value) || value.nodeType) return;
  seen.add(value);
  if (value.TYPE_NAME === 'org_docx4j_mce.AlternateContent') {
    for (const branch of [...(value.choice ?? []), ...(value.fallback ? [value.fallback] : [])]) {
      for (const item of branch.any ?? []) {
        const key = item.nodeType ? `DOM {${item.namespaceURI}}${item.localName}` : `typed {${item.name?.namespaceURI}}${item.name?.localPart}`;
        into.set(key, (into.get(key) ?? 0) + 1);
      }
    }
  }
  for (const key of Object.keys(value)) if (key !== 'PARENT') branchContents(value[key], into, seen);
}

test('kept branches: each fixture part holding mc:AlternateContent writes every one back whole, its WordprocessingML typed (CR-010 phase 0)', async () => {
  let parts = 0;
  const held = new Map();
  for (const path of await packagesUnder(fixturesDir)) {
    const flat = path.endsWith('.xml');
    const input = flat ? await readFile(path, 'utf8') : new Uint8Array(await readFile(path));
    if (flat && !input.includes(XML_PACKAGE_NS)) continue;
    const source = flat ? new FlatOpcPartStore(input) : new ZipPartStore(input);
    const expected = new Map();
    for (const name of source.partNames()) {
      if (!name.endsWith('.xml') || name.includes('_rels/') || name === '[Content_Types].xml') continue;
      const text = new TextDecoder().decode(await source.load(name));
      if (!text.includes('AlternateContent')) continue;
      const elements = [...parseXml(text).getElementsByTagNameNS(MCE_NS, 'AlternateContent')];
      if (elements.length > 0) expected.set(name, elements.map(canonical));
    }
    if (expected.size === 0) continue;
    const pkg = await OpcPackage.load(input, { mcePreprocess: false });
    await pkg.unmarshalAll();
    for (const part of pkg.parts) if (part.isUnmarshalled) branchContents(part.contents, held);
    const saved = new ZipPartStore(await pkg.save());
    for (const [name, want] of expected) {
      parts++;
      const got = [...parseXml(new TextDecoder().decode(saved.loadSync(name))).getElementsByTagNameNS(MCE_NS, 'AlternateContent')].map(canonical);
      assert.deepEqual(got, want, `${path.slice(fixturesDir.length + 1)} ${name}: every mc:AlternateContent back whole`);
    }
  }
  // 28 on 2026-10-18 (27 in the zips, 1 in the flat OPC header): a floor, so the walk cannot pass by finding nothing
  assert.ok(parts >= 28, `${parts} parts hold mc:AlternateContent`);
  // objects 0.4.0 (its CR-009, CR-010 phase 2) types w:drawing and w:pict in a branch; what stays
  // DOM is outside WordprocessingML (DrawingML frames and shapes, c:style, the x14 controls)
  assert.deepEqual([...held.keys()].filter((key) => key.startsWith(`DOM {${W}}`)), []);
  assert.ok((held.get(`typed {${W}}drawing`) ?? 0) > 0 && (held.get(`typed {${W}}pict`) ?? 0) > 0, JSON.stringify([...held]));
});

// CR-010 phase 4 (Jason, 2026-10-18: "implement CR-010 phase 4"): load keeps every mc:AlternateContent
// by default, as docx4j does since its CR-021, and `mcePreprocess: true` resolves as before. The
// warning channel stays: the default pass, `mceKeep`, reports what resolving would and changes nothing.

/** Every org_docx4j_mce.AlternateContent under a value (a generic walk, every branch). */
function alternatesIn(value, out = [], seen = new Set()) {
  if (!value || typeof value !== 'object' || value.nodeType || seen.has(value)) return out;
  seen.add(value);
  if (value.TYPE_NAME === 'org_docx4j_mce.AlternateContent') out.push(value);
  for (const key of Object.keys(value)) if (key !== 'PARENT') alternatesIn(value[key], out, seen);
  return out;
}

test('phase 4: load keeps mc:AlternateContent by default, and mcePreprocess: true resolves it', async () => {
  const bytes = await fixture('loadAndSave.docx');
  const source = new TextDecoder().decode(new ZipPartStore(bytes).loadSync('word/document.xml'));
  for (const [options, kept] of [[undefined, true], [{ mcePreprocess: false }, true], [{ mcePreprocess: true }, false]]) {
    const pkg = await WordprocessingMLPackage.load(bytes, options);
    assert.equal(pkg.loadOptions.preprocessor, kept ? mceKeep : mcePreprocess, JSON.stringify(options));
    const main = pkg.getMainDocumentPart();
    assert.equal(alternatesIn(await main.getContents()).length, kept ? 1 : 0, JSON.stringify(options));
    const saved = new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync('word/document.xml'));
    assert.equal(saved.split('<mc:Fallback>').length - 1, kept ? source.split('<mc:Fallback>').length - 1 : 0, JSON.stringify(options));
    // the content API answers the same either way (phase 3)
    const body = await pkg.getBody();
    assert.deepEqual([body.paragraphs.length, body.inlinePictures.length], [41, 1], JSON.stringify(options));
  }
  // a clone loads as its original did
  const resolved = await WordprocessingMLPackage.load(bytes, { mcePreprocess: true });
  assert.equal(alternatesIn(await (await resolved.clone()).getMainDocumentPart().getContents()).length, 0);
  assert.equal(alternatesIn(await (await (await WordprocessingMLPackage.load(bytes)).clone()).getMainDocumentPart().getContents()).length, 1);
});

test('phase 4: the keeping pass reports what resolving reports, and leaves the DOM as it was', () => {
  const wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const box = (text) => `<w:drawing><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></w:drawing>`;
  // an element whose branches differ, holding a second in its Choice whose branches differ too, and a
  // third in its Fallback, which neither pass enters (resolving drops it)
  const nestedInChoice = `<mc:AlternateContent><mc:Choice Requires="wps">${box('inner new')}</mc:Choice><mc:Fallback>${box('inner old')}</mc:Fallback></mc:AlternateContent>`;
  const nestedInFallback = `<mc:AlternateContent><mc:Choice Requires="wps">${box('x')}</mc:Choice><mc:Fallback>${box('y')}</mc:Fallback></mc:AlternateContent>`;
  const xml = `<w:document xmlns:w="${W}" xmlns:mc="${MCE_NS}" xmlns:wps="${wps}"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps">${box('outer new')}${nestedInChoice}</mc:Choice><mc:Fallback>${box('outer old')}${nestedInFallback}</mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>`;
  const resolving = preprocessAlternateContent(parseXml(xml), { partName: '/word/document.xml' });
  const doc = parseXml(xml);
  const before = canonical(doc.documentElement);
  const keeping = preprocessAlternateContent(doc, { resolve: false, partName: '/word/document.xml' });
  assert.deepEqual([resolving.resolved, resolving.kept], [2, 0]);
  assert.deepEqual([keeping.resolved, keeping.kept], [0, 2], 'the elements met: the outer one and the one in its selected branch');
  assert.equal(keeping.warnings.length, 2);
  assert.deepEqual(keeping.warnings, resolving.warnings);
  assert.equal(canonical(doc.documentElement), before, 'the DOM is untouched');
  assert.deepEqual(mceKeep(parseXml(xml)), resolving.warnings.map((w) => ({ ...w, partName: undefined, message: w.message.replace('/word/document.xml', 'a part') })));
  // KEPT_HOSTS and the nothing-selectable rule keep an element in both passes, and neither reports it
  const chart = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="${MCE_NS}"><mc:AlternateContent><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent></c:chartSpace>`;
  assert.deepEqual(preprocessAlternateContent(parseXml(chart), { resolve: false }), { resolved: 0, kept: 1, warnings: [] });
});

/** The fixture packages (zip or flat OPC) with a part holding mc:AlternateContent: path and input. */
async function packagesHolding() {
  const out = [];
  for (const path of await packagesUnder(fixturesDir)) {
    const flat = path.endsWith('.xml');
    const input = flat ? await readFile(path, 'utf8') : new Uint8Array(await readFile(path));
    if (flat && !input.includes(XML_PACKAGE_NS)) continue;
    const source = flat ? new FlatOpcPartStore(input) : new ZipPartStore(input);
    for (const name of source.partNames()) {
      if (name.endsWith('.xml') && new TextDecoder().decode(await source.load(name)).includes('AlternateContent')) {
        out.push({ path, input });
        break;
      }
    }
  }
  return out;
}

test('phase 4: over the fixtures, a kept load reports the warnings a resolved load does', async () => {
  let compared = 0;
  const counts = {};
  for (const { path, input } of await packagesHolding()) {
    // every XML part read privately, so no part is touched; kept by the option, which the default is
    // (the first test above), since before phase 4 the option meant no pass at all, and no warning
    const loads = [];
    for (const options of [{ mcePreprocess: true }, { mcePreprocess: false }]) {
      const pkg = await OpcPackage.load(input, options);
      for (const part of pkg.parts) if (typeof part.readContents === 'function') await part.readContents();
      loads.push(pkg.warnings);
    }
    const where = path.slice(fixturesDir.length + 1);
    assert.deepEqual(loads[1], loads[0], where);
    compared++;
    for (const w of loads[0]) counts[w.selected] = (counts[w.selected] ?? 0) + 1;
  }
  // 16 packages and 8 warnings on 2026-10-18 (a14 5, cx1, sle15 and tsle 1 each): floors, so the
  // comparison cannot pass by finding nothing
  assert.ok(compared >= 16, `${compared} packages hold the element`);
  assert.ok(Object.values(counts).reduce((a, b) => a + b, 0) >= 8, JSON.stringify(counts));
});

test('phase 4: insertOoxml into a loaded document writes an mc:AlternateContent as given', async () => {
  const fragment = `<w:p xmlns:w="${W}" xmlns:mc="${MCE_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:r><w:t>before</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t>new</w:t></mc:Choice><mc:Fallback><w:t>old</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>`;
  for (const [options, kept] of [[undefined, true], [{ mcePreprocess: true }, false]]) {
    const pkg = await WordprocessingMLPackage.load(await fixture('loadAndSave.docx'), options);
    await (await pkg.getBody()).insertOoxml(fragment, 'End');
    const saved = new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync('word/document.xml'));
    assert.equal(saved.includes('<mc:Choice Requires="w14"><w:t>new</w:t></mc:Choice><mc:Fallback><w:t>old</w:t></mc:Fallback>'), kept, JSON.stringify(options));
    assert.equal(saved.includes('>old</w:t>'), kept, 'resolved, the Fallback is dropped');
    assert.ok(saved.includes('>new</w:t>'));
  }
});

// A dependency defect phase 4 exposed. @docx4j/jsonix 3.4.1's Base64Binary.decode assigned an
// undeclared `input` (jsonix-factory.js line 4489). Node loads jsonix as sloppy CommonJS, so the
// assignment made a global and the suites passed; a strict ESM bundle (the editor's Vite build,
// vitest's jsdom) threw ReferenceError, and the part failed to unmarshal. Kept, a Word text box's VML
// Fallback is unmarshalled, and its v:shape's o:gfxdata (xsd:base64Binary) is decoded; resolved, the
// Fallback is dropped first. Found by the editor against phase 4 on 2026-10-18; fixed in jsonix 3.4.2
// (231679e, three `var`s), when this test, which had asserted the leak from the five fixtures below,
// failed as meant (CR-010 section 14.4). A global created here is a ReferenceError in a strict consumer.
test('phase 4: no load leaks a global, which a strict-mode consumer would get as a ReferenceError (jsonix 3.4.2)', async () => {
  const leaks = {};
  const holding = await packagesHolding();
  for (const { path, input } of holding) {
    for (const [mode, options] of [['resolved', { mcePreprocess: true }], ['kept', undefined]]) {
      const before = new Set(Object.getOwnPropertyNames(globalThis));
      const pkg = await OpcPackage.load(input, options);
      await pkg.unmarshalAll();
      await pkg.save();
      for (const name of Object.getOwnPropertyNames(globalThis)) {
        if (before.has(name)) continue;
        (leaks[`${mode} ${name}`] ??= []).push(path.slice(fixturesDir.length + 1));
        delete globalThis[name];
      }
    }
  }
  // packages whose kept load leaked `input` on 3.4.1 are among those walked
  const names = holding.map(({ path }) => path.slice(fixturesDir.length + 1));
  for (const name of ['loadAndSave.docx', 'mc-alternate-content-header.xml', 'revisions/cr043/cr043-tracked.docx']) assert.ok(names.includes(name), name);
  assert.deepEqual(leaks, {});
});

// A kept mc:Choice's Requires names prefixes that its content need not use (waterfall's Requires="cx1"
// over cx: content, Excel's Requires="x14" over x:controls). @docx4j/generated-objects-ts 0.4.1's marshal
// removed each root declaration no element or attribute used, so 10 of these fixtures' 33 were written
// undeclared; 0.4.2 counts them as used, as docx4j CR-049 does (JaxbXmlPart's mcChoiceNamespaces: "These
// must be declared, or Word can't open the docx"). This test asserted the 0.4.1 defect until the range
// was raised to ^0.4.2 (CR-010 section 14.8). The resolved load writes no such Choice.
test('phase 4: a re-marshalled mc:Choice keeps the declaration its Requires names', async () => {
  const undeclared = { kept: {}, resolved: {} };
  let written = 0;
  for (const { path, input } of await packagesHolding()) {
    for (const [load, options] of [['kept', undefined], ['resolved', { mcePreprocess: true }]]) {
      const pkg = await OpcPackage.load(input, options);
      await pkg.unmarshalAll();
      const marshalled = new Set([...pkg.parts].filter((part) => part.isUnmarshalled).map((part) => part.partName.name));
      const saved = new ZipPartStore(await pkg.save());
      for (const name of saved.partNames()) {
        if (!marshalled.has(name.startsWith('/') ? name : `/${name}`)) continue;
        const xml = new TextDecoder().decode(saved.loadSync(name));
        if (!xml.includes('Requires')) continue;
        for (const choice of parseXml(xml).getElementsByTagNameNS(MCE_NS, 'Choice')) {
          for (const prefix of (choice.getAttribute('Requires') ?? '').split(/\s+/).filter(Boolean)) {
            if (load === 'kept') written++;
            if (choice.lookupNamespaceURI(prefix)) continue;
            const key = `${path.slice(fixturesDir.length + 1)} ${name.replace(/^\//, '')} ${prefix}`;
            undeclared[load][key] = (undeclared[load][key] ?? 0) + 1;
          }
        }
      }
    }
  }
  assert.ok(written >= 30, `${written} Requires prefixes written by the kept load`);
  assert.deepEqual(undeclared.resolved, {});
  assert.deepEqual(undeclared.kept, {});
});
