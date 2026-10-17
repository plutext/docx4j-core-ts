import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import { join } from 'node:path';
import { OpcPackage, WordprocessingMLPackage, HeaderPart, parseXml, resolveAlternateContent, preprocessAlternateContent, createMcePreprocessor, UNDERSTOOD_NAMESPACES, KEPT_HOSTS, MCE_NS, ZipPartStore } from '../dist/index.mjs';
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
  const pkg = await OpcPackage.load(flat);
  assert.ok(pkg instanceof WordprocessingMLPackage);
  const headers = pkg.getMainDocumentPart().headerParts;
  assert.ok(headers.length > 0);
  const header = headers.find((h) => h instanceof HeaderPart);
  const hdr = await header.getContents();
  assert.equal(hdr.TYPE_NAME, 'org_docx4j_wml.Hdr');
  // the wps choice was taken: a w:drawing, not the VML fallback
  const xml = await header.getXml();
  assert.ok(!xml.includes('mc:AlternateContent'));
  assert.ok(xml.includes('<w:drawing>') || xml.includes('<w:drawing '), 'choice branch');
  assert.ok(!xml.includes('<w:pict'), 'fallback branch dropped');
  // the whole package saves as a zip and reloads
  const zip = await pkg.save();
  const back = await OpcPackage.load(zip);
  assert.equal(back.parts.size, pkg.parts.size);
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
  // kept as DOM, and a save writes both back: nothing is lost, but the content API cannot see a
  // w:drawing inside a DOM branch, which is why resolving on load stays the default (CR-001
  // section 17.5).
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
  // DOM, not typed: the objects mapping declares w:drawing and w:pict only within their parents,
  // and the branches' lax wildcard types global elements alone. objects CR-009 (proposed
  // 2026-10-17, for CR-010 phase 2) declares them globally, as docx4j does; these two assertions
  // fail when it arrives, and then assert the typed form instead.
  assert.equal(alternates[0].choice[0].any[0].nodeType, 1);
  assert.equal(alternates[0].choice[0].any[0].nodeName, 'w:drawing');
  assert.equal(alternates[0].fallback.any[0].nodeType, 1);
  assert.equal(alternates[0].fallback.any[0].nodeName, 'w:pict');
  const source = new TextDecoder().decode(new ZipPartStore(await fixture('loadAndSave.docx')).loadSync('word/document.xml'));
  const saved = new TextDecoder().decode(new ZipPartStore(await raw.save()).loadSync('word/document.xml'));
  for (const tag of ['<mc:AlternateContent', '<mc:Choice', '<mc:Fallback', '<w:drawing', '<w:pict']) {
    assert.equal(saved.split(tag).length, source.split(tag).length, `${tag} kept`);
  }
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

    const pkg = await OpcPackage.load(bytes);
    await pkg.unmarshalAll();
    const store = new ZipPartStore(await pkg.save());
    const saved = parseXml(new TextDecoder().decode(store.loadSync(name)));
    assert.equal(saved.getElementsByTagNameNS(MCE_NS, 'AlternateContent').length, 0, 'the branch was chosen on load');
    const frames = [...saved.getElementsByTagNameNS(frame.namespaceURI, frame.localName)].map(elementForm);
    assert.ok(frames.includes(elementForm(frame)), `the Choice's ${frame.nodeName} is written back whole`);
    const chart = saved.getElementsByTagNameNS(CX, 'chart');
    assert.equal(chart.length, 1, 'cx:chart, not the Fallback');
    const rels = store.loadSync(name.replace(/([^/]+)$/, '_rels/$1.rels'));
    const id = chart[0].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    assert.match(new TextDecoder().decode(rels), new RegExp(`Id="${id}"[^>]*chartEx|chartEx[^>]*Id="${id}"`), 'its relationship still names the chartEx part');
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
