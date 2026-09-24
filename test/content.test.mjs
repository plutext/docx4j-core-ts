import { test } from 'node:test';
import assert from 'node:assert/strict';
import { WordprocessingMLPackage, OpcPackage, Paragraph, Range, wml as parseWml, textOf, find, ZipPartStore } from '../dist/index.mjs';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { fixture } from './helpers.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

test('body: insertParagraph, text, paragraphs, docx4j aliases, round trip', async () => {
  const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
  const body = pkg.body;
  const p1 = body.insertParagraph('Hello World', 'End');
  assert.ok(p1 instanceof Paragraph);
  assert.equal(p1.text, 'Hello World');
  assert.equal(p1.style, 'Normal');
  const p0 = body.insertParagraph('Title', 'Start');
  p0.styleBuiltIn = 'Heading 1';
  assert.equal(p0.styleId, 'Heading1');
  assert.equal(p0.style, 'Heading 1');
  assert.equal(p0.styleBuiltIn, 'Heading1');
  assert.equal(p1.styleBuiltIn, 'Normal');
  body.addStyledParagraphOfText('Heading2', 'Sub');
  body.addParagraphOfText('Tab\there');
  assert.deepEqual(body.paragraphs.map((p) => p.text), ['Title', 'Hello World', 'Sub', 'Tab\there']);
  assert.equal(body.text, 'Title\nHello World\nSub\nTab\there');
  assert.equal(body.paragraphs[3].p.content[0].value.content.length, 3, 'tab became w:tab');
  // PARENT is linked on inserted content
  assert.equal(p1.p.PARENT, pkg.getMainDocumentPart().contents.body);
  assert.equal(p1.p.content[0].value.PARENT, p1.p);
  const back = await WordprocessingMLPackage.load(await pkg.save());
  const b2 = await back.getBody();
  assert.deepEqual(b2.paragraphs.map((p) => [p.text, p.styleId]), [['Title', 'Heading1'], ['Hello World', 'Normal'], ['Sub', 'Heading2'], ['Tab\there', 'Normal']]);
  assert.equal(b2.paragraphs[3].text, 'Tab\there');
});

test('paragraph: insertText, insertParagraph, delete, alignment, indents, outline level', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const p = body.insertParagraph('middle', 'End');
  p.insertText('start ', 'Start');
  p.insertText(' end', 'End');
  assert.equal(p.text, 'start middle end');
  assert.equal(p.runs.length, 1, 'text extends the existing run');
  const r = p.insertText('replaced', 'Replace');
  assert.ok(r instanceof Range);
  assert.equal(p.text, 'replaced');
  assert.equal(r.text, 'replaced');
  const before = p.insertParagraph('before', 'Before');
  const after = p.insertParagraph('after', 'After');
  assert.deepEqual(body.paragraphs.map((q) => q.text), ['before', 'replaced', 'after']);
  assert.equal(before.index, 0);
  p.delete();
  assert.deepEqual(body.paragraphs.map((q) => q.text), ['before', 'after']);
  after.alignment = 'Centered';
  assert.equal(after.alignment, 'Centered');
  assert.equal(after.p.pPr.jc.val, 'center');
  after.leftIndent = 36;
  after.firstLineIndent = -18;
  after.spaceAfter = 6;
  assert.equal(after.p.pPr.ind.left, 720);
  assert.equal(after.p.pPr.ind.hanging, 360);
  assert.equal(after.firstLineIndent, -18);
  assert.equal(after.p.pPr.spacing.after, 120);
  assert.equal(after.spaceAfter, 6);
  after.outlineLevel = 2;
  assert.equal(after.p.pPr.outlineLvl.val, 1);
  assert.equal(before.outlineLevel, 10);
  before.text = 'set text';
  assert.equal(before.text, 'set text');
  // CR-001 Phase B step 2: reads are effective, and an absent w:jc resolves to Left as in Word
  assert.equal(before.alignment, 'Left');
  assert.equal(before.formatting({ direct: true }).alignment, 'Unknown', 'the direct read still says nothing is stated');
  assert.equal(after.formatting({ direct: true }).alignment, 'Centered');
});

test('font: reads direct formatting, writes to runs, splits runs for a range', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const p = pkg.body.insertParagraph('plain bold plain', 'End');
  const [range] = p.search('bold');
  range.font.bold = true;
  range.font.italic = true;
  range.font.size = 14;
  range.font.color = '#FF0000';
  range.font.highlightColor = 'yellow';
  range.font.underline = 'Single';
  assert.equal(p.runs.length, 3, 'run split at both ends');
  assert.equal(p.text, 'plain bold plain');
  const bold = p.runs[1].value;
  assert.deepEqual(Object.keys(bold.rPr).filter((k) => k !== 'TYPE_NAME').sort(), ['b', 'bCs', 'color', 'highlight', 'i', 'iCs', 'sz', 'szCs', 'u']);
  assert.equal(bold.rPr.sz.val, 28);
  assert.equal(bold.rPr.color.val, 'FF0000');
  assert.equal(bold.rPr.highlight.val, 'yellow');
  assert.equal(range.font.bold, true);
  assert.equal(range.font.size, 14);
  assert.equal(range.font.color, '#FF0000');
  assert.equal(range.font.highlightColor, '#FFFF00');
  assert.equal(range.font.underline, 'Single');
  assert.equal(p.runs[0].value.rPr, undefined, 'other runs untouched');
  assert.equal(p.font.bold, false, 'first run is not bold');
  p.font.name = 'Arial';
  assert.ok(p.runs.every((r) => r.value.rFonts === undefined && r.value.rPr.rFonts.ascii === 'Arial'));
  range.font.bold = false;
  assert.equal(bold.rPr.b, undefined);
  const xml = await p.getXml();
  assert.ok(xml.includes('<w:i/>'));
  // saved text survives the split
  const back = await WordprocessingMLPackage.load(await pkg.save());
  assert.equal((await back.getBody()).paragraphs[0].text, 'plain bold plain');
});

test('search and replace across runs, options, wildcards', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const p = body.insertParagraph('The quick brown fox jumps over the lazy dog', 'End');
  p.search('quick')[0].font.bold = true;   // splits runs: 'The ' | 'quick' | ' brown fox...'
  assert.equal(p.runs.length, 3);
  const across = p.search('quick brown');
  assert.equal(across.length, 1, 'match spanning two runs');
  assert.equal(across[0].text, 'quick brown');
  across[0].insertText('slow red', 'Replace');
  assert.equal(p.text, 'The slow red fox jumps over the lazy dog');
  assert.equal(p.runs[1].value.rPr.b !== undefined, true, 'replacement keeps the first run formatting');
  assert.deepEqual(body.search('the').map((r) => r.text), ['The', 'the']);
  assert.deepEqual(body.search('the', { matchCase: true }).map((r) => r.text), ['the']);
  assert.deepEqual(body.search('fox', { matchWholeWord: true }).length, 1);
  assert.deepEqual(body.search('ox', { matchWholeWord: true }).length, 0);
  assert.deepEqual(body.search('<s*d>', { matchWildcards: true }).map((r) => r.text), ['slow red']);
  for (const r of body.search('dog')) r.insertText('cat', 'Replace');
  assert.equal(p.text, 'The slow red fox jumps over the lazy cat');
  // delete a range and insert around it
  const [lazy] = p.search(' lazy');
  lazy.delete();
  assert.equal(p.text, 'The slow red fox jumps over the cat');
  const [cat] = p.search('cat');
  cat.insertText('big ', 'Before');
  cat.insertText('!', 'After');
  assert.equal(p.text, 'The slow red fox jumps over the big cat!');
  assert.equal(cat.text, 'cat');
});

test('insertXml and insertElement: fragments, targets, validation, PARENT', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const a = body.insertParagraph('A', 'End');
  const inserted = await body.insertXml('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>From XML</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>', 'After', a);
  assert.equal(inserted.length, 2);
  assert.ok(inserted[0] instanceof Paragraph);
  assert.equal(inserted[0].styleId, 'Heading1');
  assert.equal(inserted[0].p.PARENT, pkg.getMainDocumentPart().contents.body);
  assert.equal(body.tables.length, 1);
  assert.deepEqual(body.paragraphs.map((p) => p.text), ['A', 'From XML', 'cell']);
  // a run-level fragment is rejected by insertElement, and the message says why
  const [run] = await parseWml('<w:r><w:t>x</w:t></w:r>');
  assert.equal(run.name.localPart, 'r');
  assert.throws(() => body.insertElement(run, 'End'), /cannot hold r \(org_docx4j_wml\.R\)/);
  // el-built content through addObject
  body.addObject(el.p({ content: [el.r({ content: [el.t({ value: 'from el' })] })] }));
  assert.equal(body.paragraphs.at(-1).text, 'from el');
  assert.equal(body.paragraphs.at(-1).p.PARENT, pkg.getMainDocumentPart().contents.body);
  // insert before by address
  body.insertElement(el.p({ content: [el.r({ content: [el.t({ value: 'first' })] })] }), 'Before', 'body/0');
  assert.equal(body.paragraphs[0].text, 'first');
  // w14 attributes and mc in a fragment parse
  const [withId] = await parseWml(`<w:p w14:paraId="1A2B3C4D"><w:r><w:t xml:space="preserve"> spaced</w:t></w:r></w:p>`);
  assert.equal(withId.value.paraId, '1A2B3C4D');
  assert.equal(textOf(withId.value), ' spaced');
  const back = await WordprocessingMLPackage.load(await pkg.save());
  assert.deepEqual((await back.getBody()).paragraphs.map((p) => p.text), ['first', 'A', 'From XML', 'cell', 'from el']);
});

test('addresses, paraIds and outline on a loaded document', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('loadAndSave.docx'));
  const body = await pkg.getBody();
  const outline = body.outline();
  assert.ok(outline.paragraphs.length > 5);
  assert.ok(outline.paragraphs.every((p) => /^body\/\d+/.test(p.address)));
  assert.ok(outline.paragraphs.some((p) => p.paraId), 'Word wrote paraIds');
  const third = outline.paragraphs[2];
  const p = body.paragraphAt(third.address);
  assert.equal(p.text, third.text);
  assert.equal(body.addressOf(p), third.address);
  assert.equal(body.paragraphAt({ paraId: third.paraId }).text, third.text);
  assert.equal(body.paragraphAt(`w14:${third.paraId.toLowerCase()}`).text, third.text);
  const byText = body.paragraphAt({ contains: third.text.slice(0, 8) });
  assert.ok(byText);
  // a new paragraph gets a paraId because the document uses them
  const added = p.insertParagraph('added', 'After');
  assert.match(added.paraId, /^[0-9A-F]{8}$/);
  assert.equal(body.addressOf(added), third.address.replace(/\d+$/, (n) => String(Number(n) + 1)));
  // tables get addresses and their paragraphs are nested under them
  if (outline.tables.length) {
    const t = outline.tables[0];
    assert.ok(outline.paragraphs.some((q) => q.address.startsWith(t.address + '/')));
  }
  // headers and footers through the package
  const doc = await pkg.outline();
  assert.equal(doc.headers.length, 3);
  assert.equal(doc.footers.length, 3);
  assert.ok(doc.headers.every((h) => /^rId\d+$/.test(h.relId)));
  const hp = doc.headers.find((h) => h.paragraphs.length > 0);
  if (hp) {
    const hpara = await pkg.paragraphAt(hp.paragraphs[0].address);
    assert.equal(hpara.text, hp.paragraphs[0].text);
  }
  // edits through views are saved
  p.insertText(' [edited]', 'End');
  const back = await WordprocessingMLPackage.load(await pkg.save());
  assert.ok((await back.getBody()).paragraphAt(third.address).text.endsWith(' [edited]'));
  assert.equal(new ZipPartStore(await pkg.save()).has('word/document.xml'), true);
});

test('body of a header; insertText Replace; clear; find', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('loadAndSave.docx'));
  const main = pkg.getMainDocumentPart();
  const header = main.headerParts[0];
  const hb = await header.getBody();
  const n = hb.paragraphs.length;
  hb.insertParagraph('Header line', 'End');
  assert.equal(hb.paragraphs.length, n + 1);
  assert.ok(hb.addressOf(hb.paragraphs.at(-1)).startsWith('header:rId'));
  const body = await pkg.getBody();
  const r = body.insertText('Only this', 'Replace');
  assert.equal(body.paragraphs.length, 1);
  assert.equal(r.text, 'Only this');
  assert.ok(main.contents.body.sectPr, 'section properties kept');
  body.clear();
  assert.equal(body.paragraphs.length, 0);
  assert.equal(find(main.contents, 'org_docx4j_wml.P').length, 0);
  const back = await OpcPackage.load(await pkg.save());
  const hb2 = await back.getMainDocumentPart().headerParts[0].getBody();
  assert.equal(hb2.paragraphs.at(-1).text, 'Header line');
});

test('run mapping: applyRunOptions / readRunOptions round trip, underline true, bad highlight', async () => {
  const { applyRunOptions, readRunOptions } = await import('../dist/index.mjs');
  const all = { bold: true, italic: true, underline: 'Double', strikeThrough: true, doubleStrikeThrough: true, subscript: true, superscript: false, name: 'Arial', size: 11.5, color: '#00FF00', highlightColor: '#FFFF00', style: 'Strong' };
  const rPr = applyRunOptions({}, all);
  assert.deepEqual(readRunOptions(rPr), all);
  assert.equal(rPr.sz.val, 23);
  assert.equal(rPr.highlight.val, 'yellow');
  assert.equal(readRunOptions(applyRunOptions({}, { underline: true })).underline, 'Single');
  assert.equal(readRunOptions(applyRunOptions(rPr, { underline: false })).underline, 'None');
  assert.equal(readRunOptions(applyRunOptions({}, { color: 'auto' })).color, 'auto');
  assert.equal(readRunOptions(applyRunOptions(rPr, { bold: false, name: '', color: '', highlightColor: null, style: '' })).bold, false);
  assert.deepEqual(Object.keys(rPr).filter((k) => ['b', 'bCs', 'rFonts', 'color', 'highlight', 'rStyle'].includes(k)), []);
  assert.throws(() => applyRunOptions({}, { highlightColor: '#123456' }), /Not a highlight colour/);
  assert.deepEqual(readRunOptions(undefined), { bold: false, italic: false, underline: 'None', strikeThrough: false, doubleStrikeThrough: false, subscript: false, superscript: false, name: '', size: 0, color: '', highlightColor: null, style: '' });
});

test('wml: the wrapper for content controls comes from the first decisive descendant', async () => {
  const tn = (els) => els.map((e) => e.value.TYPE_NAME);
  const sdtR = '<w:sdt><w:sdtPr><w:tag w:val="x"/></w:sdtPr><w:sdtContent><w:r><w:t>run</w:t></w:r></w:sdtContent></w:sdt>';
  assert.deepEqual(tn(await parseWml(sdtR)), ['org_docx4j_wml.SdtRun']);
  assert.equal((await parseWml(sdtR))[0].value.sdtContent.TYPE_NAME, 'org_docx4j_wml.CTSdtContentRun');
  assert.deepEqual(tn(await parseWml(sdtR, { wrapper: 'body' })), ['org_docx4j_wml.SdtBlock']);
  const sdtP = '<w:sdt><w:sdtContent><w:p><w:r><w:t>para</w:t></w:r></w:p></w:sdtContent></w:sdt>';
  assert.deepEqual(tn(await parseWml(sdtP)), ['org_docx4j_wml.SdtBlock']);
  // nested: the outer decision types the inner control from its container
  const nestedP = `<w:sdt><w:sdtContent>${sdtP}</w:sdtContent></w:sdt>`;
  const [outerP] = await parseWml(nestedP);
  assert.equal(outerP.value.TYPE_NAME, 'org_docx4j_wml.SdtBlock');
  assert.equal(outerP.value.sdtContent.content[0].value.TYPE_NAME, 'org_docx4j_wml.SdtBlock');
  const nestedR = `<w:sdt><w:sdtContent>${sdtR}</w:sdtContent></w:sdt>`;
  const [outerR] = await parseWml(nestedR);
  assert.equal(outerR.value.TYPE_NAME, 'org_docx4j_wml.SdtRun');
  assert.equal(outerR.value.sdtContent.content[0].value.TYPE_NAME, 'org_docx4j_wml.SdtRun');
  // a bookmark before the run is not decisive
  const withBookmark = '<w:sdt><w:sdtContent><w:bookmarkStart w:id="0" w:name="b"/><w:r><w:t>run</w:t></w:r><w:bookmarkEnd w:id="0"/></w:sdtContent></w:sdt>';
  assert.deepEqual(tn(await parseWml(withBookmark)), ['org_docx4j_wml.SdtRun']);
  // a row control
  const sdtTr = '<w:sdt><w:sdtContent><w:tr><w:tc><w:p/></w:tc></w:tr></w:sdtContent></w:sdt>';
  assert.deepEqual(tn(await parseWml(sdtTr)), ['org_docx4j_wml.CTSdtRow']);
  // nothing decisive: block
  assert.deepEqual(tn(await parseWml('<w:sdt><w:sdtContent/></w:sdt>')), ['org_docx4j_wml.SdtBlock']);
  assert.deepEqual(tn(await parseWml('<w:sdt><w:sdtPr><w:tag w:val="x"/></w:sdtPr></w:sdt>')), ['org_docx4j_wml.SdtBlock']);
  assert.deepEqual(tn(await parseWml('<w:sdt><w:sdtContent><w:sdt><w:sdtContent/></w:sdt></w:sdtContent></w:sdt>')), ['org_docx4j_wml.SdtBlock']);
  // customXml
  assert.deepEqual(tn(await parseWml('<w:customXml w:element="e"><w:r><w:t>x</w:t></w:r></w:customXml>')), ['org_docx4j_wml.CTCustomXmlRun']);
  assert.deepEqual(tn(await parseWml('<w:customXml w:element="e"><w:p/></w:customXml>')), ['org_docx4j_wml.CTCustomXmlBlock']);
  // the other rows still hold: siblings, run content, a global element
  assert.deepEqual(tn(await parseWml('<w:p/><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>')), ['org_docx4j_wml.P', 'org_docx4j_wml.Tbl']);
  assert.deepEqual(tn(await parseWml('<w:t>x</w:t><w:tab/>')), ['org_docx4j_wml.Text', 'org_docx4j_wml.R.Tab']);
  const [settings] = await parseWml('<w:settings><w:zoom w:percent="100"/></w:settings>');
  assert.equal(settings.value.TYPE_NAME, 'org_docx4j_wml.CTSettings');
  await assert.rejects(() => parseWml('<w:settings/><w:settings/>'));
});

test('text inside a tracked insertion is read and searchable; deleted text is not', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const [p] = await body.insertXml('<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r><w:ins w:id="1" w:author="a" w:date="2026-09-10T00:00:00Z"><w:r><w:t>inserted</w:t></w:r></w:ins><w:del w:id="2" w:author="a" w:date="2026-09-10T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>', 'End');
  assert.equal(p.text, 'kept inserted');
  assert.equal(p.runs.length, 2);
  const [hit] = body.search('inserted');
  assert.ok(hit);
  hit.font.bold = true;
  assert.equal(p.p.content[1].value.customXmlOrSmartTagOrSdt[0].value.rPr.b !== undefined, true);
  hit.insertText('added', 'Replace');
  assert.equal(p.text, 'kept added');
});

test('style is the display name, styleBuiltIn the Word.Style value, styleId the id', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const p = body.insertParagraph('x', 'End');
  // without the styles part unmarshalled: derived from the id
  p.styleBuiltIn = 'Toc1';
  assert.equal(p.styleId, 'TOC1', "Word's spelling of the id");
  assert.equal(p.style, 'TOC 1');
  assert.equal(p.styleBuiltIn, 'Toc1');
  p.style = 'Comment Text';
  assert.equal(p.styleId, 'CommentText');
  assert.equal(p.styleBuiltIn, 'Other');
  p.style = 'Heading 2';
  assert.equal(p.styleId, 'Heading2');
  p.style = 'Heading3';                           // an id is accepted too
  assert.equal(p.style, 'Heading 3');
  // with the styles part unmarshalled: names come from w:name, custom ones included
  const styles = await pkg.getMainDocumentPart().styleDefinitionsPart.getContents();
  styles.style.push({ type: 'paragraph', styleId: 'MyStyle', name: { val: 'My Style' }, basedOn: { val: 'Normal' } });
  p.style = 'my style';
  assert.equal(p.styleId, 'MyStyle');
  assert.equal(p.style, 'My Style');
  assert.equal(p.styleBuiltIn, 'Other');
  p.style = 'heading 1';                          // the stored name of a built-in
  assert.equal(p.styleId, 'Heading1');
  assert.equal(p.style, 'Heading 1');
  assert.throws(() => { p.styleBuiltIn = 'Other'; }, RangeError);
  assert.equal(p.getRange().styleId, 'Heading1');
});

// CR-002 section 20: Office JS Range.hyperlink. `address#location` separates the address from a
// location within it, so "#name" is a link inside the document (w:anchor) and an address is an
// external relationship of the range's part.
test('Range.hyperlink: an external link over a span, with its relationship', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const paragraph = pkg.body.insertParagraph('see the docs here', 'End');
  assert.equal(paragraph.getRange().hyperlink, '', 'no hyperlink to begin with');

  const span = paragraph.search('the docs')[0];
  span.hyperlink = 'https://docx4j.org/';
  assert.equal(span.hyperlink, 'https://docx4j.org/');
  assert.equal(paragraph.text, 'see the docs here', 'the text is untouched');

  const store = new ZipPartStore(await pkg.save());
  const xml = new TextDecoder().decode(store.loadSync('word/document.xml'));
  const id = /<w:hyperlink[^>]*r:id="([^"]+)"/.exec(xml)?.[1];
  assert.ok(id, 'a w:hyperlink with an r:id is written');
  // w:history is what Word and docx4j both write on a hyperlink they make. The value is `true`
  // where they write `1`: both are ST_OnOff, and this package marshals every boolean that way -
  // a re-marshalled Word document turns its own w:history="1" into "true" too.
  assert.match(xml, /<w:hyperlink[^>]*w:history="(1|true)"/);
  // w:history is what Word and docx4j both write on a hyperlink they make. The value is `true`
  // where they write `1`: both are ST_OnOff, and this package marshals every boolean that way -
  // a re-marshalled Word document turns its own w:history="1" into "true" too.
  assert.match(xml, /<w:hyperlink[^>]*w:history="(1|true)"/);
  const rels = new TextDecoder().decode(store.loadSync('word/_rels/document.xml.rels'));
  assert.ok(rels.includes(`Id="${id}"`) && rels.includes('Target="https://docx4j.org/"')
    && rels.includes('TargetMode="External"'), 'and an external relationship for it');

  // the runs the link covers are exactly the span's
  const back = await WordprocessingMLPackage.load(await pkg.save());
  await back.getMainDocumentPart().getContents();
  const reloaded = back.body.paragraphs[0];
  assert.equal(reloaded.search('the docs')[0].hyperlink, 'https://docx4j.org/');
  assert.equal(reloaded.search('here')[0].hyperlink, '', 'outside the span there is none');
});

test('Range.hyperlink: a location inside the document, and removal', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const paragraph = pkg.body.insertParagraph('jump to chapter one', 'End');
  const range = paragraph.getRange();

  range.hyperlink = '#chapter1';
  assert.equal(range.hyperlink, '#chapter1', 'an anchor, with no relationship');
  const xml = new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync('word/document.xml'));
  assert.ok(/<w:hyperlink[^>]*w:anchor="chapter1"/.test(xml));
  assert.ok(!/<w:hyperlink[^>]*r:id=/.test(xml), 'and no r:id');
  assert.match(xml, /<w:hyperlink[^>]*w:history="(1|true)"/, 'an internal link carries it too');

  // address and location together
  range.hyperlink = 'https://example.com/doc#section2';
  assert.equal(range.hyperlink, 'https://example.com/doc#section2');

  range.hyperlink = '';
  assert.equal(range.hyperlink, '');
  assert.equal(paragraph.text, 'jump to chapter one', 'the runs survive the unwrapping');
  const after = new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync('word/document.xml'));
  assert.ok(!after.includes('<w:hyperlink'), 'and the holder is gone');
});

test('Range.hyperlink: reads what Word wrote', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('hyperlink_dupe.docx'));
  await pkg.getMainDocumentPart().getContents();
  const links = pkg.body.paragraphs
    .map((p) => p.getRange().hyperlink)
    .filter((h) => h !== '');
  assert.ok(links.length > 0, 'the fixture has hyperlinks');
  for (const link of links) assert.match(link, /^https?:\/\/|^#/, link);
});
