// CR-005 section 11: Word fragments. A control tagged od:progid=Word.Document, bound through its od:xpath entry to a node
// holding a Word document as escaped Flat OPC, and Word's own rich text mapped by w15:dataBinding to such a node, filled with
// the document's content as Word 365 resolves docx4j's altChunk (route 'resolve'), or with that altChunk (route 'altChunk').
// The goldens are Word 365's saves of the editor's ED-005 proposal 46 check (test/fragments/, copied unedited from the editor's
// c71ebd6 with a SHA-1 manifest): 46a, docx4j's altChunk in the control, which Word resolved; 46b, the same with w:matchSrc;
// 46c, the template as authored, whose w15-mapped control Word filled from the node. Each was run in compatibility mode 15
// (-word365) and 12 (-word365-mode12).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { OpcPackage, WordprocessingMLPackage, Texts, PlainText, TextOptions, parseXml, applyFlatOpcBindingTo, isFragmentBound, isW15FlatOpcMapped } from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, 'fragments');
const manifest = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8'));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const AF = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk';
const quiet = () => { const e = console.error; const w = console.warn; console.error = () => {}; console.warn = () => {}; return () => { console.error = e; console.warn = w; }; };

const bytesOf = (name) => new Uint8Array(readFileSync(join(dir, name)));
const entriesOf = (bytes) => Object.fromEntries(Object.entries(unzipSync(bytes)).map(([k, v]) => [k, v]));
const textOf = (entries, name) => (entries[name] ? strFromU8(entries[name]) : undefined);

/** The template run through applyBindings with the fragment options, saved: the package and the saved entries. */
async function bound(name, fragments = {}) {
  const restore = quiet();
  try {
    const pkg = await OpcPackage.load(bytesOf(name));
    const result = await pkg.customXmlParts.applyBindings({ fragments });
    const saved = await pkg.save();
    return { pkg, result, saved, entries: entriesOf(saved) };
  } finally { restore(); }
}

const elements = (parent, local) => [...parent.getElementsByTagNameNS(W, local)];
const childrenNamed = (parent, local) => [...parent.childNodes].filter((c) => c.nodeType === 1 && c.namespaceURI === W && c.localName === local);
const wAttr = (e, name) => e?.getAttributeNS(W, name) ?? null;

/** An element as a string with its attributes sorted, less the rsids, w14 ids and `w:noProof` Word adds. */
function canon(e) {
  if (!e) return '';
  // an on/off value as Word writes it (this package's marshal writes `true` where Word writes `1`)
  const value = (v) => (v === 'true' ? '1' : v === 'false' ? '0' : v);
  const attrs = [...e.attributes].filter((a) => !/rsid|^w14:|^xmlns/.test(a.name)).map((a) => `${a.localName}=${value(a.value)}`).sort().join(',');
  const kids = [...e.childNodes].filter((c) => c.nodeType === 1 && c.localName !== 'noProof').map(canon).join('');
  return `<${e.localName}${attrs ? ' ' + attrs : ''}>${kids}`;
}

/** A control's `w:sdtContent` by its alias. */
function contentOfControl(doc, alias) {
  const sdt = elements(doc, 'sdt').find((s) => wAttr(elements(childrenNamed(s, 'sdtPr')[0], 'alias')[0], 'val') === alias);
  return sdt ? childrenNamed(sdt, 'sdtContent')[0] : undefined;
}

/** What a list reference reads as: its level's format and text, through the package's numbering. */
function listOf(entries) {
  const numbering = parseXml(textOf(entries, 'word/numbering.xml') ?? `<w:numbering xmlns:w="${W}"/>`);
  return (numPr) => {
    if (!numPr) return '';
    const numId = wAttr(elements(numPr, 'numId')[0], 'val');
    const ilvl = wAttr(elements(numPr, 'ilvl')[0], 'val') ?? '0';
    const num = elements(numbering, 'num').find((n) => wAttr(n, 'numId') === numId);
    const abstractId = wAttr(elements(num, 'abstractNumId')[0], 'val');
    const abstract = elements(numbering, 'abstractNum').find((a) => wAttr(a, 'abstractNumId') === abstractId);
    const lvl = elements(abstract, 'lvl').find((l) => wAttr(l, 'ilvl') === ilvl);
    return `${wAttr(elements(lvl, 'numFmt')[0], 'val')} ${wAttr(elements(lvl, 'lvlText')[0], 'val')}`;
  };
}

/** A footnote reference as its note's text. */
function notesOf(entries) {
  const notes = parseXml(textOf(entries, 'word/footnotes.xml') ?? `<w:footnotes xmlns:w="${W}"/>`);
  return (id) => {
    const note = elements(notes, 'footnote').find((n) => wAttr(n, 'id') === id);
    return note ? elements(note, 't').map((t) => t.textContent).join('') : `(no footnote ${id})`;
  };
}

/** Each paragraph in a control's content: its text, its properties less `w:numPr`, the list it reads in, each run's properties, its notes and pictures. */
function paragraphsOf(entries, alias) {
  const doc = parseXml(textOf(entries, 'word/document.xml'));
  const content = contentOfControl(doc, alias);
  assert.ok(content, `a control ${alias}`);
  const list = listOf(entries);
  const note = notesOf(entries);
  return elements(content, 'p').map((p) => {
    const pPr = childrenNamed(p, 'pPr')[0];
    const numPr = pPr ? childrenNamed(pPr, 'numPr')[0] : undefined;
    return {
      text: elements(p, 't').map((t) => t.textContent).join(''),
      pPr: pPr ? [...pPr.childNodes].filter((c) => c.nodeType === 1 && c.localName !== 'numPr').map(canon).join('') : '',
      list: list(numPr),
      runs: elements(p, 'r').map((r) => canon(childrenNamed(r, 'rPr')[0])).filter((x) => x !== '' && x !== '<rPr>'),
      notes: elements(p, 'footnoteReference').map((ref) => note(wAttr(ref, 'id'))),
      pictures: elements(p, 'drawing').length,
    };
  });
}

/** The content's tables, inner controls and pictures' bytes. */
function blocksOf(entries, alias) {
  const doc = parseXml(textOf(entries, 'word/document.xml'));
  const content = contentOfControl(doc, alias);
  const rels = parseXml(textOf(entries, 'word/_rels/document.xml.rels'));
  const target = (id) => [...rels.getElementsByTagNameNS(PR, 'Relationship')].find((r) => r.getAttribute('Id') === id)?.getAttribute('Target');
  return {
    tables: elements(content, 'tbl').map((t) => ({ borders: canon(elements(t, 'tblBorders')[0]), cells: elements(t, 'tc').map((tc) => elements(tc, 't').map((x) => x.textContent).join('')) })),
    controls: elements(content, 'sdt').map((s) => ({ tag: wAttr(elements(s, 'tag')[0], 'val'), id: wAttr(elements(s, 'id')[0], 'val') })),
    pictures: [...content.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip')].map((b) => createHash('sha1').update(entries[`word/${target(b.getAttributeNS(R, 'embed'))}`]).digest('hex')),
  };
}

/** A style's definition by id, as `canon` writes it, less what Word adds to a style it keeps (`w:uiPriority`). */
function styleOf(entries, id) {
  const style = elements(parseXml(textOf(entries, 'word/styles.xml')), 'style').find((s) => wAttr(s, 'styleId') === id);
  return style ? canon(style).replace(/<uiPriority val=\d+>/, '') : undefined;
}

const sectionsOf = (entries) => elements(parseXml(textOf(entries, 'word/document.xml')), 'sectPr').map((s) => canon(elements(s, 'pgSz')[0]));

test('CR-005 section 11: the files are as the manifest has them', () => {
  for (const [name, sha1] of Object.entries(manifest.files)) assert.equal(createHash('sha1').update(bytesOf(name)).digest('hex'), sha1, name);
  assert.equal(Object.keys(manifest.files).length, 10);
});

test('the controls each form names: od:progid=Word.Document with an od:xpath entry, and rich text mapped by w15:dataBinding to Flat OPC', async () => {
  const restore = quiet();
  try {
    const pkg = await OpcPackage.load(bytesOf('46c-template.docx'));
    await pkg.customXmlParts.load();
    const controls = (await pkg.getMainDocumentPart().getBody()).contentControlsIn('every');
    assert.deepEqual(controls.map((c) => [c.title, isFragmentBound(c), isW15FlatOpcMapped(c)]), [
      ['Word fragment', true, false], ['Plain text mapped', false, false], ['Rich text mapped', false, true],
    ]);
  } finally { restore(); }
});

for (const mode of ['', '-mode12']) {
  test(`route 'resolve': the od:progid control holds what Word 365 resolved docx4j's altChunk to (46a${mode}), paragraph for paragraph`, async () => {
    const word = entriesOf(bytesOf(`46a-word365${mode}.docx`));
    const { entries, result } = await bound('46c-template.docx', { w15FlatOpc: 'leave' });
    assert.deepEqual(paragraphsOf(entries, 'Word fragment'), paragraphsOf(word, 'Word fragment'));
    assert.deepEqual(blocksOf(entries, 'Word fragment'), blocksOf(word, 'Word fragment'));
    // the text pass counted the two mapped controls and left them; the fragment pass bound the tagged one
    assert.deepEqual({ ...result }, { bound: 3, updated: 1, skipped: 2 });
  });

  test(`route 'resolve': Word's w15-mapped rich text holds what Word 365 filled it with on open (46c${mode})`, async () => {
    const word = entriesOf(bytesOf(`46c-word365${mode}.docx`));
    const { entries, result } = await bound('46c-template.docx');
    assert.deepEqual(paragraphsOf(entries, 'Rich text mapped'), paragraphsOf(word, 'Rich text mapped'));
    // the inner control's w:id: Word left the od:progid control, so the w15 control's copy kept 201; here the od:progid control's
    // copy took 201 first, and this one's is made fresh
    const ours = blocksOf(entries, 'Rich text mapped');
    assert.deepEqual(ours.controls, [{ tag: 'inner', id: '202' }]);
    assert.deepEqual({ ...ours, controls: undefined }, { ...blocksOf(word, 'Rich text mapped'), controls: undefined });
    assert.deepEqual(blocksOf(word, 'Rich text mapped').controls, [{ tag: 'inner', id: '201' }]);
    assert.deepEqual({ ...result }, { bound: 3, updated: 2, skipped: 1 });
  });

  test(`keepSourceFormatting: the fragment's styles as direct formatting, as Word's w:matchSrc (46b${mode})`, async () => {
    const word = entriesOf(bytesOf(`46b-word365${mode}.docx`));
    const { entries } = await bound('46c-template.docx', { w15FlatOpc: 'leave', keepSourceFormatting: true });
    assert.deepEqual(paragraphsOf(entries, 'Word fragment'), paragraphsOf(word, 'Word fragment'));
    assert.deepEqual(blocksOf(entries, 'Word fragment'), blocksOf(word, 'Word fragment'));
  });
}

test('the reconciliation is Word\'s: the document\'s Heading 1, the styles it lacks copied, the fragment\'s list a new one, the document\'s list unbroken, the section dropped', async () => {
  const word = entriesOf(bytesOf('46a-word365.docx'));
  const template = entriesOf(bytesOf('46c-template.docx'));
  const { entries, pkg } = await bound('46c-template.docx', { w15FlatOpc: 'leave' });
  // a style both define is the document's; one the document lacks is copied with its definition
  assert.equal(styleOf(entries, 'Heading1'), styleOf(template, 'Heading1'));
  for (const id of ['FragmentNote', 'FootnoteReference']) assert.equal(styleOf(entries, id), styleOf(word, id), id);
  // Word links a character style it makes to a copied paragraph style (FootnoteTextChar); the engine needn't (proposal 46, part 2)
  assert.equal(styleOf(entries, 'FootnoteText'), styleOf(word, 'FootnoteText').replace('<link val=FootnoteTextChar>', ''));
  assert.equal(styleOf(entries, 'FootnoteTextChar'), undefined);
  assert.ok(styleOf(word, 'FootnoteTextChar'));
  // the lists as read: the fragment's I and II, the document's item three still 3 (Word's text, through the text model)
  const labels = async (bytes) => {
    const restore = quiet();
    try {
      const o = TextOptions.of({}); o.set('profile', 'word-text'); o.set('list_labels', 'true');
      return PlainText.of(await Texts.of(await OpcPackage.load(bytes), o)).split('\n').filter((l) => /list, item/.test(l) && !l.startsWith('<')).map((l) => l.trim());
    } finally { restore(); }
  };
  const ours = await labels(await pkg.save());
  assert.deepEqual(ours, await labels(bytesOf('46a-word365.docx')));
  assert.deepEqual(ours.map((l) => l.split('\t')[0]), ['1.', '2.', 'I.', 'II.', '3.']);
  // one section, the document's A4: the fragment's landscape A5 is dropped, as Word drops it
  assert.deepEqual(sectionsOf(entries), sectionsOf(word));
});

test('the save reads back as written: no warning, no misplaced content, the fragment\'s text in the text model', async () => {
  const { saved } = await bound('46c-template.docx');
  const restore = quiet();
  try {
    const again = await OpcPackage.load(saved);
    await again.unmarshalAll();
    assert.deepEqual(again.warnings.map((w) => w.code), []);
    assert.deepEqual((await again.checkContent()).map((f) => f.rule), []);
    const text = PlainText.of(await Texts.of(again));
    assert.match(text, /The fragment's footnote\./);
    // Known difference, reported (a WD question, REQ-509): the text model reads the plain text control mapped to the same node
    // as the node's text, the escaped package, where Word 365 shows its placeholder (46c (2), both runs). So the fragment's
    // heading is read three times: twice resolved, once inside that text.
    assert.equal(text.split('Fragment heading').length - 1, 3);
    assert.match(text, /^<\?xml[^\n]*<pkg:package/m);
  } finally { restore(); }
});

test('route \'altChunk\': docx4j\'s form, as 46a holds it: an application/xml part holding the node\'s Flat OPC, an aFChunk relationship, the chunk in the control; w:matchSrc as 46b holds it', async () => {
  for (const [keepSourceFormatting, golden] of [[false, '46a-altchunk.docx'], [true, '46b-altchunk-matchsrc.docx']]) {
    const docx4j = entriesOf(bytesOf(golden));
    const { entries, result } = await bound('46c-template.docx', { route: 'altChunk', keepSourceFormatting, w15FlatOpc: 'leave' });
    const doc = parseXml(textOf(entries, 'word/document.xml'));
    const chunk = elements(contentOfControl(doc, 'Word fragment'), 'altChunk');
    assert.equal(chunk.length, 1);
    assert.equal(canon(contentOfControl(doc, 'Word fragment')).replace(/id=rId\w+/, 'id=…'), canon(contentOfControl(parseXml(textOf(docx4j, 'word/document.xml')), 'Word fragment')).replace(/id=rId\w+/, 'id=…'));
    const rels = [...parseXml(textOf(entries, 'word/_rels/document.xml.rels')).getElementsByTagNameNS(PR, 'Relationship')];
    const rel = rels.find((r) => r.getAttribute('Id') === chunk[0].getAttributeNS(R, 'id'));
    assert.equal(rel.getAttribute('Type'), AF);
    assert.equal(rel.getAttribute('Target'), 'chunk1.xml');
    // the chunk is the node's text, as docx4j writes it; this template's node begins with the XML declaration already
    assert.equal(textOf(entries, 'word/chunk1.xml'), textOf(docx4j, 'word/chunk1.xml'));
    const types = textOf(entries, '[Content_Types].xml');
    assert.ok(/PartName="\/word\/chunk1\.xml" ContentType="application\/xml"/.test(types) || /Extension="xml" ContentType="application\/xml"/.test(types));
    // nothing reconciled: Word does it on open
    assert.equal(styleOf(entries, 'FragmentNote'), undefined);
    assert.deepEqual({ ...result }, { bound: 3, updated: 1, skipped: 2 });
  }
});

test('Word\'s w15 form: its mapping kept by default, as Word 365 keeps it; w15Mapping \'drop\' removes the w15:dataBinding alone; w15FlatOpc \'leave\' leaves the control', async () => {
  const mapped = (entries) => {
    const doc = parseXml(textOf(entries, 'word/document.xml'));
    const sdt = elements(doc, 'sdt').find((s) => elements(s, 'alias').some((a) => wAttr(a, 'val') === 'Rich text mapped'));
    return { binding: sdt.getElementsByTagNameNS('http://schemas.microsoft.com/office/word/2012/wordml', 'dataBinding').length, id: wAttr(elements(sdt, 'id')[0], 'val'), text: elements(childrenNamed(sdt, 'sdtContent')[0], 't').map((t) => t.textContent).join('') };
  };
  const word = mapped(entriesOf(bytesOf('46c-word365.docx')));
  const kept = mapped((await bound('46c-template.docx')).entries);
  assert.equal(kept.binding, 1);
  assert.equal(word.binding, 1);
  assert.equal(kept.text, word.text);
  const dropped = mapped((await bound('46c-template.docx', { w15Mapping: 'drop' })).entries);
  assert.deepEqual({ ...dropped, text: undefined }, { binding: 0, id: '103', text: undefined });
  assert.equal(dropped.text, word.text);
  const left = mapped((await bound('46c-template.docx', { w15FlatOpc: 'leave' })).entries);
  assert.deepEqual(left, { binding: 1, id: '103', text: '(3) rich text control placeholder' });
  // Word 365 leaves the plain text control mapped to the same node with its placeholder, and so does the text pass
  const plain = (entries) => elements(contentOfControl(parseXml(textOf(entries, 'word/document.xml')), 'Plain text mapped'), 't').map((t) => t.textContent).join('');
  assert.equal(plain((await bound('46c-template.docx')).entries), plain(entriesOf(bytesOf('46c-word365.docx'))));
});

// ------------------------------------------------------------------------------------------------ shapes no file holds

const FRAGMENT = (() => {
  const data = textOf(entriesOf(bytesOf('46c-template.docx')), 'customXml/item1.xml');
  return parseXml(data).getElementsByTagName('fragment')[0].textContent;
})();
const escape = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** A new document with `blocks` (WordprocessingML with `{control}` for the bound control's tag) and a data node holding `fragment`. */
async function made(blocks, fragment = FRAGMENT) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const data = pkg.customXmlParts.add(`<data><fragment>${escape(fragment)}</fragment><text>plain</text></data>`);
  pkg.customXmlParts.add(`<xpaths xmlns="http://opendope.org/xpaths"><xpath id="x1"><dataBinding storeItemID="${data.id}" xpath="/data/fragment" prefixMappings=""/></xpath><xpath id="x2"><dataBinding storeItemID="${data.id}" xpath="/data/text" prefixMappings=""/></xpath><xpath id="x3"><dataBinding storeItemID="${data.id}" xpath="/data/none" prefixMappings=""/></xpath></xpaths>`);
  await body.insertOoxml(blocks.replace(/\{control\}/g, 'od:xpath=x1&amp;od:progid=Word.Document'), 'End');
  return { pkg, body };
}

const tagOf = (entry) => `od:xpath=${entry}&amp;od:progid=Word.Document`;

test('placement: a block control in a cell ends the cell with a paragraph; a cell-level control\'s cell takes the content; a run-level control is left, with a note', async () => {
  const restore = quiet();
  try {
    const cell = (inner) => `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${inner}</w:tr></w:tbl>`;
    const { pkg } = await made(
      cell(`<w:tc><w:sdt><w:sdtPr><w:alias w:val="block in cell"/><w:tag w:val="{control}"/><w:id w:val="11"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc>`)
      + cell(`<w:sdt><w:sdtPr><w:alias w:val="cell"/><w:tag w:val="{control}"/><w:id w:val="12"/></w:sdtPr><w:sdtContent><w:tc><w:p><w:r><w:t>y</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt>`)
      + `<w:p><w:sdt><w:sdtPr><w:alias w:val="run"/><w:tag w:val="{control}"/><w:id w:val="13"/></w:sdtPr><w:sdtContent><w:r><w:t>z</w:t></w:r></w:sdtContent></w:sdt></w:p>`);
    const result = await pkg.customXmlParts.applyBindings();
    assert.deepEqual({ bound: result.bound, updated: result.updated, skipped: result.skipped }, { bound: 3, updated: 2, skipped: 1 });
    assert.ok(result.notes.some((n) => /run-level control is left/.test(n)), result.notes.join('\n'));
    const doc = parseXml(textOf(entriesOf(await pkg.save()), 'word/document.xml'));
    const tcs = elements(doc, 'tc').filter((tc) => tc.parentNode.localName === 'tr' || tc.parentNode.localName === 'sdtContent');
    for (const tc of tcs.slice(0, 2)) {
      const last = [...tc.childNodes].filter((c) => c.nodeType === 1).pop();
      assert.equal(last.localName === 'p' || last.localName === 'sdt', true);
    }
    assert.ok(elements(doc, 't').some((t) => t.textContent === 'z'));
    // the cell-level control keeps its cell, and the cell holds the fragment
    const cellControl = elements(doc, 'sdt').find((s) => elements(s, 'alias').some((a) => wAttr(a, 'val') === 'cell'));
    const tc = childrenNamed(childrenNamed(cellControl, 'sdtContent')[0], 'tc')[0];
    assert.ok(elements(tc, 't').some((t) => /Fragment heading/.test(t.textContent)));
    assert.equal([...tc.childNodes].filter((c) => c.nodeType === 1).pop().localName, 'p');
  } finally { restore(); }
});

test('what the document doesn\'t take is dropped with a note: a section break inside the fragment, comments; ids made fresh only where they collide', async () => {
  const restore = quiet();
  try {
    const fragment = FRAGMENT
      .replace('<w:p><w:pPr><w:pStyle w:val="FragmentNote"/></w:pPr>', '<w:p><w:pPr><w:pStyle w:val="FragmentNote"/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:bookmarkStart w:id="0" w:name="taken"/><w:bookmarkStart w:id="7" w:name="free"/><w:commentRangeStart w:id="0"/>')
      .replace('A paragraph in FragmentNote, a style the document lacks (green, italic).</w:t></w:r>', 'A paragraph in FragmentNote, a style the document lacks (green, italic).</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:bookmarkEnd w:id="0"/><w:bookmarkEnd w:id="7"/>');
    assert.notEqual(fragment, FRAGMENT);
    const { pkg } = await made(`<w:p><w:bookmarkStart w:id="0" w:name="mine"/><w:r><w:t>mine</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`
      + `<w:sdt><w:sdtPr><w:id w:val="201"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>the document's 201</w:t></w:r></w:p></w:sdtContent></w:sdt>`
      + `<w:sdt><w:sdtPr><w:alias w:val="fragment"/><w:tag w:val="{control}"/><w:id w:val="99"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`, fragment);
    const result = await pkg.customXmlParts.applyBindings();
    assert.equal(result.updated, 1);
    const notes = result.notes.join('\n');
    assert.match(notes, /dropped section breaks inside the fragment \(1\)/);
    assert.match(notes, /dropped comments \(their ranges and references\) \(3\)/);
    const doc = parseXml(textOf(entriesOf(await pkg.save()), 'word/document.xml'));
    const content = contentOfControl(doc, 'fragment');
    assert.equal(elements(content, 'sectPr').length, 0);
    assert.equal(elements(content, 'commentRangeStart').length + elements(content, 'commentReference').length, 0);
    // bookmark 0 is the document's: the fragment's 'taken' moves, its end with it; 7 collides with nothing and stays
    const starts = Object.fromEntries(elements(content, 'bookmarkStart').map((b) => [wAttr(b, 'name'), wAttr(b, 'id')]));
    assert.equal(starts.free, '7');
    assert.notEqual(starts.taken, '0');
    assert.deepEqual(elements(content, 'bookmarkEnd').map((b) => wAttr(b, 'id')).sort(), [starts.taken, '7'].sort());
    // the inner control's 201 collides with the document's and is made fresh; its tag is kept and it isn't bound
    const inner = elements(content, 'sdt').find((s) => elements(s, 'tag').some((t) => wAttr(t, 'val') === 'inner'));
    assert.notEqual(wAttr(elements(inner, 'id')[0], 'val'), '201');
  } finally { restore(); }
});

test('a fragment in a header: its parts go beside the header, a note\'s reference is dropped with a note (a header holds no notes)', async () => {
  const restore = quiet();
  try {
    const entries = entriesOf(bytesOf('46c-template.docx'));
    const tag = tagOf('x1');
    const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${W}" xmlns:r="${R}"><w:sdt><w:sdtPr><w:alias w:val="header fragment"/><w:tag w:val="${tag}"/><w:id w:val="301"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt></w:hdr>`;
    const doc = textOf(entries, 'word/document.xml').replace('<w:sectPr>', '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>');
    const rels = textOf(entries, 'word/_rels/document.xml.rels').replace('</Relationships>', '<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>');
    const types = textOf(entries, '[Content_Types].xml').replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
    const bytes = zipSync({ ...entries, 'word/document.xml': strToU8(doc), 'word/_rels/document.xml.rels': strToU8(rels), '[Content_Types].xml': strToU8(types), 'word/header1.xml': strToU8(header) });
    const pkg = await OpcPackage.load(bytes);
    const result = await pkg.customXmlParts.applyBindings({ fragments: { w15FlatOpc: 'leave' } });
    assert.equal(result.updated, 2);
    assert.match(result.notes.join('\n'), /dropped notes, in a header or a footer \(1\)/);
    const saved = entriesOf(await pkg.save());
    const hdr = parseXml(textOf(saved, 'word/header1.xml'));
    assert.equal(elements(hdr, 'footnoteReference').length, 0);
    assert.ok(elements(hdr, 't').some((t) => /Fragment heading/.test(t.textContent)));
    const hdrRels = [...parseXml(textOf(saved, 'word/_rels/header1.xml.rels')).getElementsByTagNameNS(PR, 'Relationship')];
    assert.ok(hdrRels.some((r) => /image$/.test(r.getAttribute('Type'))));
  } finally { restore(); }
});

test('what is left, with a note: no entry, a node holding no Flat OPC, a package that can\'t be read; and applyFlatOpcBindingTo on one control', async () => {
  const restore = quiet();
  try {
    const control = (alias, entry) => `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:tag w:val="${tagOf(entry)}"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>${alias}</w:t></w:r></w:p></w:sdtContent></w:sdt>`;
    const { pkg } = await made(control('no entry', 'x9') + control('text node', 'x2') + control('no node', 'x3'));
    const result = await pkg.customXmlParts.applyBindings();
    assert.deepEqual({ bound: result.bound, updated: result.updated, skipped: result.skipped }, { bound: 3, updated: 0, skipped: 3 });
    assert.match(result.notes.join('\n'), /od:xpath=x9: no such entry/);
    assert.match(result.notes.join('\n'), /od:xpath=x2: the node holds no Flat OPC package/);
    assert.match(result.notes.join('\n'), /od:xpath=x3: \/data\/none selects nothing/);
    const broken = await made(control('broken', 'x1'), FRAGMENT.replace('<w:body>', '<w:body><w:p>'));
    await broken.pkg.customXmlParts.load();
    const outcome = await applyFlatOpcBindingTo((await broken.pkg.getMainDocumentPart().getBody()).contentControlsIn('every')[0], broken.pkg.customXmlParts, broken.pkg);
    assert.equal(outcome.done, false);
    assert.match(outcome.note, /can't be read/);
    // one control alone, the pass's own entry point
    const one = await made(control('one', 'x1'));
    await one.pkg.customXmlParts.load();
    const done = await applyFlatOpcBindingTo((await one.pkg.getMainDocumentPart().getBody()).contentControlsIn('every')[0], one.pkg.customXmlParts, one.pkg);
    assert.equal(done.done, true);
  } finally { restore(); }
});

test('a document without notes, numbering or a fragment style: the parts made as needed, Word\'s separators in the notes part', async () => {
  const restore = quiet();
  try {
    const { pkg } = await made(`<w:sdt><w:sdtPr><w:alias w:val="f"/><w:tag w:val="{control}"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`);
    assert.equal(pkg.getMainDocumentPart().footnotesPart, undefined);
    const result = await pkg.customXmlParts.applyBindings();
    assert.equal(result.updated, 1);
    const saved = entriesOf(await pkg.save());
    const notes = parseXml(textOf(saved, 'word/footnotes.xml'));
    assert.deepEqual(elements(notes, 'footnote').map((n) => [wAttr(n, 'id'), wAttr(n, 'type')]), [['-1', 'separator'], ['0', 'continuationSeparator'], ['1', null]]);
    assert.ok(styleOf(saved, 'FragmentNote'));
    // the created document's numbering part takes the fragment's list as a new definition
    const lists = elements(parseXml(textOf(saved, 'word/numbering.xml')), 'num').length;
    assert.ok(lists >= 1);
  } finally { restore(); }
});


test('a style the fragment makes its default, named only there, is copied without w:default: the document keeps one default paragraph style', async () => {
  const restore = quiet();
  try {
    const fragment = FRAGMENT
      .replace('w:default="1" w:styleId="Normal"', 'w:default="1" w:styleId="FragBody"').replace(/<w:basedOn w:val="Normal"\/>/g, '<w:basedOn w:val="FragBody"/>')
      .replace('<w:p><w:r><w:t xml:space="preserve">A footnote follows', '<w:p><w:pPr><w:pStyle w:val="FragBody"/></w:pPr><w:r><w:t xml:space="preserve">A footnote follows');
    assert.ok(fragment.includes('w:styleId="FragBody"') && fragment.includes('<w:pStyle w:val="FragBody"/>'));
    const { pkg } = await made(`<w:sdt><w:sdtPr><w:alias w:val="f"/><w:tag w:val="{control}"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`, fragment);
    assert.equal((await pkg.customXmlParts.applyBindings()).updated, 1);
    const styles = elements(parseXml(textOf(entriesOf(await pkg.save()), 'word/styles.xml')), 'style');
    const body = styles.find((s) => wAttr(s, 'styleId') === 'FragBody');
    assert.ok(body, 'FragBody copied');
    assert.equal(wAttr(body, 'default'), null);
    assert.equal(styles.filter((s) => wAttr(s, 'type') === 'paragraph' && ['1', 'true'].includes(wAttr(s, 'default'))).length, 1);
  } finally { restore(); }
});

test('what a run added is reported (created), so that a caller binding again can remove what the run before made: lists, styles, parts', async () => {
  const restore = quiet();
  try {
    const pkg = await OpcPackage.load(bytesOf('46c-template.docx'));
    await pkg.customXmlParts.load();
    const control = (await pkg.getMainDocumentPart().getBody()).contentControlsIn('every').find((c) => isFragmentBound(c));
    const first = await applyFlatOpcBindingTo(control, pkg.customXmlParts, pkg);
    assert.equal(first.done, true);
    assert.deepEqual({ ...first.created, styleIds: [...first.created.styleIds].sort(), parts: [...first.created.parts].sort() },
      { numIds: ['2'], abstractNumIds: ['1'], styleIds: ['FootnoteReference', 'FootnoteText', 'FragmentNote'], parts: ['/word/footnotes.xml', '/word/media/image1.png'] });
    // again, as a refresh would: a new list and a new picture part; the styles are the document's now, the notes part is there
    const second = await applyFlatOpcBindingTo(control, pkg.customXmlParts, pkg);
    assert.deepEqual(second.created, { numIds: ['3'], abstractNumIds: ['2'], styleIds: [], parts: ['/word/media/image2.png'] });
    // the altChunk route adds the chunk's part alone
    const other = await OpcPackage.load(bytesOf('46c-template.docx'));
    await other.customXmlParts.load();
    const tagged = (await other.getMainDocumentPart().getBody()).contentControlsIn('every').find((c) => isFragmentBound(c));
    const chunk = await applyFlatOpcBindingTo(tagged, other.customXmlParts, other, { route: 'altChunk' });
    assert.deepEqual(chunk.created, { numIds: [], abstractNumIds: [], styleIds: [], parts: ['/word/chunk1.xml'] });
  } finally { restore(); }
});

// ------------------------------------------------------------------------------------------------ insertOoxml (section 11.6)

// Word 365's own insertOoxml of the same package (the portfolio's bridge round insertooxml-46, Office JS through the check verb,
// read back straight after the insert), in test/fragments/io46/ with a SHA-1 manifest.
const io46 = join(dir, 'io46');
const io46Manifest = JSON.parse(readFileSync(join(io46, 'manifest.json'), 'utf8'));
const wordRun = (c) => JSON.parse(readFileSync(join(io46, `${c}.check-result.json`), 'utf8'));
/** The package a case's script gave insertOoxml, as the script holds it. */
const flatOf = (c) => JSON.parse(/const FLAT = (".*?");\n/s.exec(readFileSync(join(io46, `${c}.check-script.js`), 'utf8'))[1]);

/** Each paragraph as Office JS's snapshot has it: text, style name, list string. */
async function snapshotOf(pkg) {
  const body = await pkg.getMainDocumentPart().getBody();
  return body.paragraphs.map((p) => ({ text: p.text, style: p.style, list: p.isListItem ? p.listItem.listString : null }));
}
const wordSnapshot = (r) => r.after.paragraphs.map((p) => ({ text: p.text, style: p.style, list: p.list?.string ?? null }));

async function io46Template() {
  const restore = quiet();
  try { return await OpcPackage.load(new Uint8Array(readFileSync(join(io46, 'io46-template-no-w15.docx')))); } finally { restore(); }
}

/** Word's paragraphs set beside ours: the plain text control (2), which Word filled from its Flat OPC node at the check's open, as ours. */
const besideOurs = (word, ours) => word.map((p, i) => (p.text.startsWith('<?xml') ? ours[i] : p));

test('insertOoxml: the round\'s files are as its manifest has them; the three cases that ran opened silently', () => {
  for (const [name, sha1] of Object.entries(io46Manifest.files)) assert.equal(createHash('sha1').update(new Uint8Array(readFileSync(join(io46, name)))).digest('hex'), sha1, name);
  for (const c of ['io46-1-before', 'io46-3-replace-control', 'io46-4-before-bare']) {
    assert.equal(wordRun(c).inserted, true, c);
    assert.equal(JSON.parse(readFileSync(join(io46, `${c}.result.json`), 'utf8')).opened, 'silent', c);
  }
  assert.equal(wordRun('io46-2-twice').inserted, false);
});

test('insertOoxml Before a paragraph reconciles as Word 365\'s does (io46-1): the fragment\'s styles copied, its list a new one, its footnote, one section', async () => {
  const word = wordRun('io46-1-before');
  const pkg = await io46Template();
  const body = await pkg.getMainDocumentPart().getBody();
  const at = body.paragraphs.find((p) => p.text === 'After the control.');
  await at.getRange('Whole').insertOoxml(flatOf('io46-1-before'), 'Before');
  const ours = await snapshotOf(pkg);
  // the fragment's last paragraph joins the paragraph it went Before, as Word joins it (section 11.7)
  assert.equal(ours.at(-1).text, "Last paragraph of the fragment. Its section is landscape A5; the document's is portrait A4.After the control.");
  assert.deepEqual(ours, besideOurs(wordSnapshot(word), ours));
  // the styles: the document's Heading 1, the fragment's own copied (Word also links a FootnoteTextChar, as in 46a)
  const styles = await pkg.getMainDocumentPart().styleDefinitionsPart.getContents();
  const ids = styles.style.map((s) => s.styleId);
  for (const id of word.after.ooxml.styleIds.filter((s) => s !== 'FootnoteTextChar')) assert.ok(ids.includes(id) || ['DefaultParagraphFont', 'TableNormal', 'NoList'].includes(id), id);
  assert.equal(ids.includes('FragmentNote') && ids.includes('FootnoteText') && ids.includes('FootnoteReference'), true);
  // the lists: the document's numId 1 decimal, the fragment's a new one, upper Roman, as Word's nums have them
  const numbering = await pkg.getMainDocumentPart().numberingDefinitionsPart.getContents();
  const formatOf = (numId) => {
    const num = numbering.num.find((n) => String(n.numId) === String(numId));
    return numbering.abstractNum.find((a) => String(a.abstractNumId) === String(num.abstractNumId.val)).lvl[0].numFmt.val;
  };
  assert.deepEqual(word.after.ooxml.nums.map((n) => n.lvl0.numFmt), ['decimal', 'upperRoman']);
  assert.deepEqual([formatOf(1), formatOf(2)], ['decimal', 'upperRoman']);
  // the footnote, and the inner control with its tag and id
  const notes = await pkg.getMainDocumentPart().footnotesPart.getContents();
  assert.deepEqual(notes.footnote.filter((n) => !n.type || n.type === 'normal').map((n) => JSON.stringify(n).includes("The fragment's footnote.")), [true]);
  assert.deepEqual(word.after.footnotes, ["\u0002 The fragment's footnote."]);
  const controls = body.contentControlsIn('every').map((c) => [c.id, c.tag]);
  for (const c of word.after.contentControls) assert.ok(controls.some(([id, tag]) => id === c.id && tag === c.tag), JSON.stringify(c));
  assert.equal(word.after.sections, 1);
});

test('insertOoxml Replace in a control\'s paragraph holds what Word\'s ContentControl.insertOoxml Replace holds (io46-3), paragraph for paragraph', async () => {
  const word = wordRun('io46-3-replace-control');
  const pkg = await io46Template();
  const body = await pkg.getMainDocumentPart().getBody();
  const control = body.contentControlsIn('every').find((c) => c.title === 'Word fragment');
  await control.paragraphs[0].insertOoxml(flatOf('io46-3-replace-control'), 'Replace');
  const ours = await snapshotOf(pkg);
  assert.deepEqual(ours, besideOurs(wordSnapshot(word), ours));
});

test('insertOoxml of a package without its styles and numbering parts (io46-4): a style it doesn\'t define reads as the default, its list a new one over Word\'s default', async () => {
  const word = wordRun('io46-4-before-bare');
  const flat = flatOf('io46-4-before-bare');
  assert.doesNotMatch(flat, /\/word\/styles\.xml|\/word\/numbering\.xml/);
  const pkg = await io46Template();
  const body = await pkg.getMainDocumentPart().getBody();
  await body.paragraphs.find((p) => p.text === 'After the control.').getRange('Whole').insertOoxml(flat, 'Before');
  const ours = await snapshotOf(pkg);
  assert.deepEqual(ours, besideOurs(wordSnapshot(word), ours));
  // the heading's Heading1, which the package doesn't define, is gone, though the document defines one
  assert.equal(ours.find((p) => p.text.startsWith('Fragment heading')).style, 'Normal');
  const styles = (await pkg.getMainDocumentPart().styleDefinitionsPart.getContents()).style.map((s) => s.styleId);
  assert.equal(styles.includes('FragmentNote'), false);
  // the list: a new numId over Word's nine decimal levels, as Word made it, the document's list unbroken
  const numbering = await pkg.getMainDocumentPart().numberingDefinitionsPart.getContents();
  const added = numbering.abstractNum.at(-1);
  assert.deepEqual(added.lvl.map((l) => [l.numFmt.val, l.lvlText.val, Number(l.pPr.ind.left), Number(l.pPr.ind.hanging)]),
    Array.from({ length: 9 }, (_, l) => ['decimal', `%${l + 1}.`, 720 * (l + 1), 720]));
  assert.deepEqual(word.after.ooxml.nums.map((n) => [n.numId, n.lvl0.numFmt]), [['1', 'decimal'], ['2', 'decimal']]);
});

test('insertOoxml of a bare fragment: nothing reconciled, nothing to reconcile, and placed as a package is (one paragraph inline at the body\'s End)', async () => {
  const pkg = await io46Template();
  const body = await pkg.getMainDocumentPart().getBody();
  const before = (await pkg.getMainDocumentPart().styleDefinitionsPart.getContents()).style.length;
  await body.insertOoxml('<w:p><w:pPr><w:pStyle w:val="Unknown"/></w:pPr><w:r><w:t>bare</w:t></w:r></w:p>', 'End');
  assert.equal((await pkg.getMainDocumentPart().styleDefinitionsPart.getContents()).style.length, before);
  assert.equal(body.paragraphs.at(-1).text, 'After the control.bare');
});

// ------------------------------------------------------------------------------------------------ the join (section 11.7)

// Word 365's insertOoxml of several paragraphs into a paragraph (the portfolio's bridge round insertooxml-join, eleven Office JS
// runs on a hand-written template, read back straight after), in test/fragments/ij/ with a SHA-1 manifest. Each case's script
// names its call (CASE, TARGET) and holds its package (OOXML); the same call is made here on the same template.
const ij = join(dir, 'ij');
const ijManifest = JSON.parse(readFileSync(join(ij, 'manifest.json'), 'utf8'));
const ijCases = Object.keys(ijManifest.files).filter((f) => f.endsWith('.check-result.json')).map((f) => f.replace('.check-result.json', '')).sort();
const ijScript = (c) => {
  const js = readFileSync(join(ij, `${c}.check-script.js`), 'utf8');
  const constant = (name) => JSON.parse(new RegExp(`const ${name} = (.*?);\\n`, 's').exec(js)[1]);
  return { target: constant('TARGET'), ooxml: constant('OOXML') };
};
const W_DECLARED = `xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;

/** A body's blocks as the join is read: each paragraph's properties and its runs, adjacent runs alike in properties as one (Word's form). */
function blocksOfBody(xml) {
  const doc = parseXml(xml.includes('<w:document') ? xml : `<x ${W_DECLARED}>${xml}</x>`);
  const body = elements(doc, 'body')[0];
  const out = [];
  for (const block of [...body.childNodes].filter((c) => c.nodeType === 1)) {
    if (block.localName === 'sectPr') continue;
    if (block.localName === 'tbl') { out.push({ table: elements(block, 't').map((t) => t.textContent) }); continue; }
    const runs = [];
    for (const r of elements(block, 'r')) {
      const rPr = canon(childrenNamed(r, 'rPr')[0]);
      const text = elements(r, 't').map((t) => t.textContent).join('');
      if (runs.length > 0 && runs.at(-1).rPr === rPr) runs.at(-1).text += text; else runs.push({ rPr, text });
    }
    out.push({ pPr: canon(childrenNamed(block, 'pPr')[0]), runs: runs.filter((r) => r.text !== '') });
  }
  // getOoxml ends a body with an empty paragraph of its own
  while (out.length > 0 && out.at(-1).runs?.length === 0 && out.at(-1).pPr === '') out.pop();
  return out;
}

async function ijRun(c) {
  const { target, ooxml } = ijScript(c);
  const restore = quiet();
  try {
    const pkg = await OpcPackage.load(new Uint8Array(readFileSync(join(ij, 'ij-template.docx'))));
    const body = await pkg.getMainDocumentPart().getBody();
    const p = target.para === undefined ? undefined : body.paragraphs.find((x) => x.text === target.para);
    let returned;
    if (target.kind === 'body') returned = await body.insertOoxml(ooxml, target.loc);
    else if (target.kind === 'range') returned = await p.getRange('Whole').insertOoxml(ooxml, target.loc);
    else if (target.kind === 'paragraph') returned = await p.insertOoxml(ooxml, target.loc);
    else returned = await p.search(target.find, { matchCase: true, matchWholeWord: true })[0].insertOoxml(ooxml, target.loc);
    return { pkg, returned, blocks: blocksOfBody(strFromU8(entriesOf(await pkg.save())['word/document.xml'])) };
  } finally { restore(); }
}

test('the join\'s round: its files as the manifest has them, every check opened silently, Paragraph.insertOoxml Before and After refused by Office JS', () => {
  for (const [name, sha1] of Object.entries(ijManifest.files)) assert.equal(createHash('sha1').update(new Uint8Array(readFileSync(join(ij, name)))).digest('hex'), sha1, name);
  assert.equal(ijCases.length, 11);
  for (const c of ijCases) assert.equal(JSON.parse(readFileSync(join(ij, `${c}.result.json`), 'utf8')).opened, 'silent', c);
  for (const c of ['ij04-para-before-p1', 'ij05-para-after-p3']) assert.match(JSON.stringify(JSON.parse(readFileSync(join(ij, `${c}.check-result.json`), 'utf8')).errors), /InvalidArgument/, c);
});

for (const c of ijCases.filter((x) => !/ij04|ij05/.test(x))) {
  test(`insertOoxml joins as Word 365's does (${c}): each paragraph's properties, its runs and theirs, block for block`, async () => {
    const word = JSON.parse(readFileSync(join(ij, `${c}.check-result.json`), 'utf8'));
    const { blocks } = await ijRun(c);
    assert.deepEqual(blocks, blocksOfBody(word.after.ooxml.body));
  });
}

test('Paragraph.insertOoxml Before and After, which Office JS refuses: this package\'s own, the package\'s blocks whole beside the paragraph', async () => {
  for (const [c, loc] of [['ij04-para-before-p1', 'Before'], ['ij05-para-after-p3', 'After']]) {
    const { blocks } = await ijRun(c);
    const word = JSON.parse(readFileSync(join(ij, `${c}.check-result.json`), 'utf8'));
    assert.equal(word.after?.count ?? word.before.count, word.before.count, `${c}: Word left the document as it was`);
    assert.equal(blocks.length, 5 + 3, `${c} ${loc}: three paragraphs beside the five`);
  }
});
