// CR-010 phase 3: the content API over a kept mc:AlternateContent (loaded with mcePreprocess: false,
// the default since phase 4; the resolved load is mcePreprocess: true).
//
// Readers take the branch load would have resolved the element to (`selectedBranchOf`, docx4j's
// McMode.READ); mutators take every branch (McMode.ALL). So a kept tree answers as a resolved one
// does: the eight .docx fixtures holding the element are loaded both ways and their answers compared
// (CR-010 section 11.2 measured them apart). Body.paragraphs leaves a text box's paragraphs out in
// either load (D7, section 12).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  WordprocessingMLPackage, ZipPartStore, TableContextTracker, NAMESPACE_PREFIXES, UNDERSTOOD_NAMESPACES,
  UNDERSTOOD_PREFIXES, READ_ONLY_UNDERSTOOD_NAMESPACES, READ_UNDERSTOOD_NAMESPACES, READ_UNDERSTOOD_PREFIXES, selectedBranchOf, deepCopy,
} from '../dist/index.mjs';
import { mcBranchOf } from '@docx4j/generated-objects-ts/builders/wml';
import { fixture } from './helpers.mjs';

const MCE = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const KEPT = { mcePreprocess: false };
const RESOLVED = { mcePreprocess: true };
const AC = 'org_docx4j_mce.AlternateContent';

/** The .docx fixtures whose main document holds mc:AlternateContent (CR-010 section 11.2). */
const FIXTURES = [
  'loadAndSave.docx', 'parity/numbering-stories.docx', 'parity/numbering-stories-coretests.docx',
  'parity/tables-textbox-in-cell.docx', 'charts/waterfall.docx', 'revisions/cr043/cr043-tracked.docx',
  'revisions/cr043/cr043-accept-all.docx', 'revisions/cr043/cr043-reject-all.docx',
];

/** The five section 11.2 ran the mutators on. */
const MUTATED = ['loadAndSave.docx', 'parity/numbering-stories.docx', 'parity/tables-textbox-in-cell.docx', 'charts/waterfall.docx', 'revisions/cr043/cr043-tracked.docx'];

/** Every typed object under `root` with the given TYPE_NAME, through every branch (a generic walk). */
function allOf(root, typeName) {
  const out = [];
  const seen = new Set();
  const visit = (o) => {
    if (!o || typeof o !== 'object' || seen.has(o) || typeof o.nodeType === 'number') return;
    seen.add(o);
    if (Array.isArray(o)) { o.forEach(visit); return; }
    if (o.TYPE_NAME === typeName) out.push(o);
    for (const key in o) if (key !== 'PARENT') visit(o[key]);
  };
  visit(root);
  return out;
}

const holdsObject = (p) => allOf(p.p, AC).length > 0 || allOf(p.p, 'org_docx4j_wml.Drawing').length > 0;

async function load(name, kept) {
  const pkg = await WordprocessingMLPackage.load(await fixture(name), kept ? KEPT : RESOLVED);
  return { pkg, body: await pkg.getBody() };
}

// CR-010 section 15: the readers also count four namespaces load leaves alone, which Word 365
// understands (measured 2026-10-20) and docx4j's published understood list has.
const READ_ONLY = { w16sdtdh: 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash', w16sdtfl: 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock', w16du: 'http://schemas.microsoft.com/office/word/2023/wordml/word16du', x14ac: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac' };

test('UNDERSTOOD_PREFIXES: what load understands, by conventional prefix; selectedBranchOf reads that branch, or a Choice requiring one of the four read-only namespaces', () => {
  const byUri = new Map([...UNDERSTOOD_NAMESPACES].filter((uri) => NAMESPACE_PREFIXES[uri]).map((uri) => [uri, NAMESPACE_PREFIXES[uri]]));
  assert.deepEqual([...UNDERSTOOD_PREFIXES].sort(), [...new Set(byUri.values())].sort());
  for (const prefix of ['w', 'wps', 'wp14', 'w14', 'cx1', 'cx2', 'a14']) assert.ok(UNDERSTOOD_PREFIXES.has(prefix), prefix);
  // the read-only four: conventional, in no module, understood by the readers and, since CR-013 (D3), by a resolving load
  assert.deepEqual([...READ_ONLY_UNDERSTOOD_NAMESPACES].sort(), Object.values(READ_ONLY).sort());
  for (const [prefix, uri] of Object.entries(READ_ONLY)) {
    assert.equal(NAMESPACE_PREFIXES[uri], prefix, `${prefix} is conventional`);
    assert.ok(!UNDERSTOOD_NAMESPACES.has(uri) && !UNDERSTOOD_PREFIXES.has(prefix), `${prefix} is in no module`);
  }
  assert.deepEqual([...READ_UNDERSTOOD_PREFIXES].sort(), [...UNDERSTOOD_PREFIXES, ...Object.keys(READ_ONLY)].sort());
  assert.deepEqual([...READ_UNDERSTOOD_NAMESPACES].sort(), [...UNDERSTOOD_NAMESPACES, ...Object.values(READ_ONLY)].sort());
  // a conventional prefix neither understands (CR-010 section 11.5)
  assert.ok(Object.values(NAMESPACE_PREFIXES).includes('cx3') && !READ_UNDERSTOOD_PREFIXES.has('cx3'));
  const choiceItems = [];
  const fallbackItems = [];
  const element = (requires) => ({
    TYPE_NAME: AC,
    choice: [{ TYPE_NAME: `${AC}.Choice`, requires, any: choiceItems }],
    fallback: { TYPE_NAME: `${AC}.Fallback`, any: fallbackItems },
  });
  assert.equal(selectedBranchOf(element('wps')), choiceItems);
  for (const prefix of Object.keys(READ_ONLY)) {
    assert.equal(selectedBranchOf(element(prefix)), choiceItems, `${prefix}: the Choice, as Word takes it`);
    assert.equal(selectedBranchOf(element(`wps ${prefix}`)), choiceItems);
  }
  assert.equal(selectedBranchOf(element('cx3')), fallbackItems, 'not understood: the Fallback');
  assert.equal(mcBranchOf(element('cx3')), choiceItems, "mcBranchOf's own default would take the Choice");
  assert.equal(selectedBranchOf(element('wps cx3')), fallbackItems);
  // the ten Choices of Jason's probe (w14 to wp14), with each of the four put first: the editor's
  // mc-branches.docx cases 7, 9 and 10, and x14ac
  const tenChoices = (first) => {
    const order = ['w14', 'w15', 'w16se', 'w16cid', 'w16', 'w16cex', 'w16sdtdh', 'w16sdtfl', 'w16du', 'wp14'];
    const prefixes = [first, ...order.filter((p) => p !== first)];
    return { TYPE_NAME: AC, choice: prefixes.map((requires) => ({ TYPE_NAME: `${AC}.Choice`, requires, any: [requires] })), fallback: { TYPE_NAME: `${AC}.Fallback`, any: ['fallback'] } };
  };
  assert.deepEqual(selectedBranchOf(tenChoices('w14')), ['w14']);
  for (const first of ['w16du', 'w16sdtdh', 'w16sdtfl', 'x14ac']) assert.deepEqual(selectedBranchOf(tenChoices(first)), [first], `${first} first`);
  assert.equal(selectedBranchOf({ TYPE_NAME: AC }), undefined);
});

/** What the readers of section 11.2 answer for one load. */
async function readers(name, kept) {
  const { pkg, body } = await load(name, kept);
  const resolver = await pkg.getPropertyResolver();
  const labels = [...body.listLabels()].map(([p, label]) => [p.paraId ?? null, label.listString, label.level]);
  const walked = [];
  new TableContextTracker(resolver).walk(pkg.getMainDocumentPart().contents, (p, ctx) => walked.push([p.paraId ?? null, ctx === undefined]));
  const pictures = body.inlinePictures.map((picture) => [picture.relId, picture.width, picture.height, picture.inline.docPr.id]);
  const again = await load(name, kept);
  return {
    paragraphs: body.paragraphs.map((p) => p.text),
    text: body.text,
    labels, walked, pictures,
    accepted: await again.pkg.acceptAllRevisions(),
    rejected: await (await load(name, kept)).pkg.rejectAllRevisions(),
  };
}

test('kept and resolved loads agree: Body.paragraphs, list labels, the table-context walk, inline pictures, the resolve-all counts', async () => {
  const answers = {};
  for (const name of FIXTURES) {
    const resolved = await readers(name, false);
    const kept = await readers(name, true);
    assert.deepEqual(kept, resolved, name);
    answers[name] = resolved;
  }
  // a text box's list paragraphs are numbered once, in the box's own state: 9 labels, 3 of them the box's
  assert.equal(answers['parity/numbering-stories.docx'].labels.length, 9);
  assert.equal(answers['parity/numbering-stories-coretests.docx'].labels.length, 9);
  // the walk reaches the box's 3 paragraphs once, past Body.paragraphs' 11 (D7)
  assert.equal(answers['parity/numbering-stories.docx'].paragraphs.length, 11);
  assert.equal(answers['parity/numbering-stories.docx'].walked.length, 14);
  assert.equal(answers['parity/tables-textbox-in-cell.docx'].walked.length, 8);
  // Accept All counts a text box's changes once: 37, where every branch gave 39
  assert.equal(answers['revisions/cr043/cr043-tracked.docx'].accepted, 37);
});

test('inline pictures are pictures only: no inline text box (check 42) and no inline chart', async () => {
  for (const kept of [false, true]) {
    const at = kept ? 'kept' : 'resolved';
    assert.equal((await load('parity/numbering-stories.docx', kept)).body.inlinePictures.length, 0, `${at}: Office JS lists no picture for the inline text box`);
    assert.equal((await load('parity/tables-textbox-in-cell.docx', kept)).body.inlinePictures.length, 0, `${at}: three inline text boxes`);
    assert.equal((await load('charts/waterfall.docx', kept)).body.inlinePictures.length, 0, `${at}: a chartex frame`);
    const { body } = await load('loadAndSave.docx', kept);
    const pictures = body.inlinePictures;
    assert.equal(pictures.length, 1, `${at}: the picture, not the inline c:chart beside it`);
    assert.equal(pictures[0].imageFormat, 'Svg', 'an SVG icon (CR-002 section 52)');
  }
});

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAAEElEQVR4nGN4YGAARww4OQAs9g8Bfw7KYwAAAABJRU5ErkJggg==';

/**
 * A document whose second run holds a picture in an mc:Choice requiring `requires`, with a copy of
 * the drawing in the Fallback, and the prefix declared (through mc:Ignorable) so that load can
 * resolve it.
 */
async function pictureInBranch(requires) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const p = pkg.body.insertParagraph('a picture: ', 'End');
  const picture = p.insertInlinePictureFromBase64(PNG, 'End');
  const content = picture.run.value.content;
  const choice = picture.element;
  const fallback = deepCopy(choice);
  fallback.value.anchorOrInline[0].docPr.descr = 'the Fallback';
  content[content.indexOf(choice)] = {
    name: { namespaceURI: MCE, localPart: 'AlternateContent' },
    value: {
      TYPE_NAME: AC,
      choice: [{ TYPE_NAME: `${AC}.Choice`, requires, any: [choice] }],
      fallback: { TYPE_NAME: `${AC}.Fallback`, any: [fallback] },
    },
  };
  pkg.getMainDocumentPart().contents.ignorable = requires;
  return pkg.save();
}

const documentXml = (bytes) => new TextDecoder().decode(new ZipPartStore(bytes).loadSync('word/document.xml'));

test('a picture in the selected branch is listed once, from that branch; deleting it removes the element, every branch', async () => {
  const understood = await pictureInBranch('wp14');
  for (const kept of [false, true]) {
    const pkg = await WordprocessingMLPackage.load(understood, kept ? KEPT : RESOLVED);
    const pictures = (await pkg.getBody()).inlinePictures;
    assert.equal(pictures.length, 1, `${kept ? 'kept' : 'resolved'}: the Choice's picture, once`);
    assert.equal(pictures[0].altTextDescription, '', "the Choice's drawing, not the Fallback's");
    assert.equal(await pictures[0].getBase64(), PNG);
    assert.equal(pictures[0].holder?.value.TYPE_NAME, kept ? AC : undefined);
  }
  // with the Choice's prefix not understood, load takes the Fallback, and a kept load reads it too
  const fallback = await pictureInBranch('cx3');
  for (const kept of [false, true]) {
    const pictures = (await (await WordprocessingMLPackage.load(fallback, kept ? KEPT : RESOLVED)).getBody()).inlinePictures;
    assert.deepEqual(pictures.map((p) => p.altTextDescription), ['the Fallback'], kept ? 'kept' : 'resolved');
  }
  // a Choice requiring one of the four read-only namespaces (CR-010 section 15): a kept load reads the
  // Choice, as Word does, and since CR-013 (D3) so does a resolved load, which until then took the Fallback
  for (const prefix of Object.keys(READ_ONLY)) {
    const readOnly = await pictureInBranch(prefix);
    for (const kept of [false, true]) {
      const pictures = (await (await WordprocessingMLPackage.load(readOnly, kept ? KEPT : RESOLVED)).getBody()).inlinePictures;
      assert.deepEqual(pictures.map((p) => p.altTextDescription), [''], `${prefix}, ${kept ? 'kept' : 'resolved'}`);
    }
  }

  // deleted outright: the element goes whole, and on save the image with it
  const pkg = await WordprocessingMLPackage.load(understood, KEPT);
  const body = await pkg.getBody();
  body.inlinePictures[0].delete();
  assert.equal(allOf(body.paragraphs[0].p, AC).length, 0);
  assert.equal(body.paragraphs[0].text, 'a picture: ');
  const saved = await pkg.save();
  assert.ok(!documentXml(saved).includes('mc:AlternateContent'));
  assert.ok(![...new ZipPartStore(saved).partNames()].some((n) => n.startsWith('word/media/')), 'the image part goes too');

  // deleted under tracking: the element in a w:del, then gone on Accept All, back whole on Reject All
  for (const verb of ['acceptAllRevisions', 'rejectAllRevisions']) {
    const tracked = await WordprocessingMLPackage.load(understood, KEPT);
    const trackedBody = await tracked.getBody();
    tracked.changeTrackingMode = 'TrackAll';
    trackedBody.inlinePictures[0].delete();
    const dels = allOf(trackedBody.paragraphs[0].p, 'org_docx4j_wml.RunDel');
    assert.equal(dels.length, 1);
    assert.equal(allOf(dels[0], AC).length, 1, 'the run in the w:del holds the element');
    assert.equal(trackedBody.inlinePictures.length, 0, 'the listing is the accepted view');
    tracked.changeTrackingMode = 'Off';
    await tracked[verb]();
    const left = allOf(trackedBody.paragraphs[0].p, AC);
    if (verb === 'acceptAllRevisions') assert.equal(left.length, 0, 'accepted, the element goes');
    else {
      assert.equal(left.length, 1, 'rejected, the element stays');
      assert.equal(left[0].choice.length, 1);
      assert.ok(left[0].fallback.any.length === 1, 'with its Fallback');
      assert.equal(trackedBody.inlinePictures.length, 1);
    }
  }
});

/** The element, and the drawings, left in a story after an edit, through every branch. */
const left = (pkg) => {
  const contents = pkg.getMainDocumentPart().contents;
  return { elements: allOf(contents, AC).length, drawings: allOf(contents, 'org_docx4j_wml.Drawing').length };
};

test('a mutator removes a kept element whole: a tracked Paragraph.delete then Accept All, a whole-range delete, insertText Replace', async () => {
  for (const name of MUTATED) {
    const results = {};
    for (const kept of [false, true]) {
      const r = {};
      // CR-002 section 48's gap: a tracked delete left the element's run outside the w:del
      let { pkg, body } = await load(name, kept);
      pkg.changeTrackingMode = 'TrackAll';
      const targets = body.paragraphs.filter(holdsObject);
      assert.ok(targets.length > 0, `${name}: a paragraph holds the element`);
      for (const p of targets) p.delete();
      pkg.changeTrackingMode = 'Off';
      r.listed = body.getTrackedChanges().map((c) => [c.type, c.text]);
      await pkg.acceptAllRevisions();
      r.tracked = left(pkg);
      r.trackedSaved = documentXml(await pkg.save()).includes('mc:AlternateContent');

      ({ pkg, body } = await load(name, kept));
      for (const p of body.paragraphs.filter(holdsObject)) p.getRange('Whole').delete();
      r.range = left(pkg);

      ({ pkg, body } = await load(name, kept));
      for (const p of body.paragraphs.filter(holdsObject)) p.insertText('', 'Replace');
      r.replaced = left(pkg);
      r.texts = body.paragraphs.map((p) => p.text);
      results[kept ? 'kept' : 'resolved'] = r;
    }
    assert.deepEqual(results.kept, results.resolved, name);
    for (const step of ['tracked', 'range', 'replaced']) assert.equal(results.kept[step].elements, 0, `${name} ${step}: no element left`);
    assert.equal(results.kept.trackedSaved, false, `${name}: the save writes none`);
  }
});

test("Accept All and Reject All resolve every branch's text box, and count the selected one", async () => {
  for (const verb of ['acceptAllRevisions', 'rejectAllRevisions']) {
    const resolved = await (await load('revisions/cr043/cr043-tracked.docx', false)).pkg[verb]();
    const { pkg } = await load('revisions/cr043/cr043-tracked.docx', true);
    const [element] = allOf(pkg.getMainDocumentPart().contents, AC);
    const boxes = [...element.choice.map((c) => c.any), element.fallback.any].map((branch) => allOf(branch, 'org_docx4j_wml.CTTxbxContent'));
    assert.ok(boxes.every((list) => list.length === 1), 'a text box in each branch');
    const revisions = (box) => ['RunIns', 'RunDel', 'RunTrackChange'].flatMap((t) => allOf(box, `org_docx4j_wml.${t}`)).length;
    assert.ok(boxes.every(([box]) => revisions(box) > 0), 'each with tracked changes');
    assert.equal(await pkg[verb](), resolved, `${verb}: the count a resolved load gives`);
    for (const [box] of boxes) assert.equal(revisions(box), 0, `${verb}: every branch's text box is resolved`);
  }
});

// CR-012 section 12.2 (the editor's ED-009 rule, phase 4): a kept mc:AlternateContent whose selected
// branch holds text reads as that branch's items, among the runs (the editor's case 8, the shape of
// docx4j's text-extraction probes) or in a run (cases 1, 2 and 6), and an edit lands in that branch,
// the other branches kept as they were read. Section 12.4's NotSupportedError (0.5.0's silent misplaced
// edits) is gone with it. Each edit gives the text a resolved load gives.
test('ED-009: an edit over a kept element\'s text lands in its selected branch, the others kept; resolved, the same text', async () => {
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const make = await WordprocessingMLPackage.createPackage();
  make.getMainDocumentPart().setXml(`<w:document xmlns:w="${W_NS}" xmlns:mc="${MCE}" xmlns:wps="${WPS}" mc:Ignorable="wps"><w:body>`
    + '<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:r><w:rPr><w:b/></w:rPr><w:t>bold ying</w:t></w:r>'
    + '<w:r><w:t xml:space="preserve"> plain ying</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>yang</w:t></w:r></mc:Fallback></mc:AlternateContent>'
    + '<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>');
  const source = await make.save();
  const edits = {
    'replace': [(p) => p.search('plain')[0].insertText('PLAIN', 'Replace'), 'Before bold ying PLAIN ying after', 'bold ying PLAIN ying'],
    'insert at a range start': [(p) => p.search('plain')[0].insertText('very ', 'Start'), 'Before bold ying very plain ying after', 'bold ying very plain ying'],
    'delete': [(p) => p.search('plain')[0].delete(), 'Before bold ying  ying after', 'bold ying  ying'],
    'format': [(p) => { p.search('plain')[0].font.italic = true; }, 'Before bold ying plain ying after', 'bold ying plain ying'],
    'insert at the paragraph end': [(p) => p.insertText('!', 'End'), 'Before bold ying plain ying after!', 'bold ying plain ying'],
    'replace across the element': [(p) => p.search('Before bold')[0].insertText('x', 'Replace'), 'x ying plain ying after', ' ying plain ying'],
    'insert where the element starts': [(p) => p.search('Before ')[0].getRange('End').insertText('x', 'Start'), 'Before xbold ying plain ying after', 'xbold ying plain ying'],
  };
  const documentXml = async (pkg) => new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync('word/document.xml'));
  const choiceText = (xml) => [...xml.match(/<mc:Choice Requires="wps">(.*?)<\/mc:Choice>/s)[1].matchAll(/<w:t(?: [^>]*)?>([^<]*)</g)].map((m) => m[1]).join('');
  for (const [name, [edit, expected, choice]] of Object.entries(edits)) {
    const pkg = await WordprocessingMLPackage.load(source, KEPT);
    const p = (await pkg.getBody()).paragraphs[0];
    assert.equal(p.text, 'Before bold ying plain ying after');
    assert.equal(p.segments().map((s) => s.text).join(''), p.text, 'the segments are the text');
    edit(p);
    assert.equal(p.text, expected, name);
    const xml = await documentXml(pkg);
    assert.match(xml, /<mc:Fallback><w:r><w:t>yang<\/w:t><\/w:r><\/mc:Fallback>/, `${name}: the Fallback as it was read`);
    assert.equal(choiceText(xml), choice, `${name}: the Choice`);
    if (name === 'format') assert.equal(p.search('plain')[0].font.italic, true);
    const resolved = await WordprocessingMLPackage.load(source, RESOLVED);
    const q = (await resolved.getBody()).paragraphs[0];
    edit(q);
    assert.equal(q.text, expected, `${name}, resolved`);
  }
  // a tracked replacement: the w:del and the w:ins in the Choice
  const tracked = await WordprocessingMLPackage.load(source, KEPT);
  tracked.changeTrackingMode = 'TrackAll';
  const t = (await tracked.getBody()).paragraphs[0];
  t.search('plain')[0].insertText('PLAIN', 'Replace');
  assert.equal(t.text, 'Before bold ying PLAIN ying after');
  assert.match(await documentXml(tracked), /<mc:Choice Requires="wps">.*<w:del [^>]*><w:r><w:delText>plain<\/w:delText><\/w:r><\/w:del><w:ins [^>]*><w:r><w:t>PLAIN<\/w:t><\/w:r><\/w:ins>.*<\/mc:Choice><mc:Fallback><w:r><w:t>yang/s);
  // the branch's text deleted whole: the element goes, every branch with it (section 12.2 item 4)
  const emptied = await WordprocessingMLPackage.load(source, KEPT);
  const e = (await emptied.getBody()).paragraphs[0];
  e.search('bold ying plain ying')[0].delete();
  assert.equal(e.text, 'Before  after');
  assert.doesNotMatch(await documentXml(emptied), /AlternateContent|yang/);

  // A run-level element (cases 1 and 2): its branch's items are the run's own, at the run's offsets.
  const runLevel = await WordprocessingMLPackage.createPackage();
  runLevel.getMainDocumentPart().setXml(`<w:document xmlns:w="${W_NS}" xmlns:mc="${MCE}" xmlns:wps="${WPS}" mc:Ignorable="wps"><w:body><w:p><w:r><w:t xml:space="preserve">The Earth is </w:t></w:r>`
    + '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>flat</w:t><w:tab/><w:t>really</w:t></mc:Choice><mc:Fallback><w:t>round</w:t></mc:Fallback></mc:AlternateContent><w:t>!</w:t></w:r></w:p><w:sectPr/></w:body></w:document>');
  const runSource = await runLevel.save();
  const runParagraph = async () => (await (await WordprocessingMLPackage.load(runSource, KEPT)).getBody()).paragraphs[0];
  const r = await runParagraph();
  assert.equal(r.text, 'The Earth is flat\treally!');
  assert.equal(r.search('flat').length, 1, 'search finds the branch\'s word (check 49)');
  assert.equal(r.search('round').length, 0, 'and never the Fallback\'s');
  r.search('flat')[0].insertText('oblate', 'Replace');
  assert.equal(r.text, 'The Earth is oblate\treally!');
  assert.match(await r.getXml(), /<mc:Choice Requires="wps"><w:t>oblate<\/w:t><w:tab\/><w:t>really<\/w:t><\/mc:Choice><mc:Fallback><w:t>round<\/w:t><\/mc:Fallback>/);
  // formatting part of it splits the run, and the element with it: the first piece keeps the Fallback, a
  // later piece's other branches are empty (ED-009 section 3.4, P2); only the span is formatted
  const f = await runParagraph();
  f.search('lat')[0].font.bold = true;
  assert.equal(f.text, 'The Earth is flat\treally!');
  assert.match(await f.getXml(), new RegExp('<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>f</w:t></mc:Choice><mc:Fallback><w:t>round</w:t></mc:Fallback></mc:AlternateContent></w:r>'
    + '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:t>lat</w:t></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>'
    + '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:tab/><w:t>really</w:t></mc:Choice><mc:Fallback/></mc:AlternateContent><w:t>!</w:t></w:r>'));
  // a tracked deletion: the Choice's w:delText inside w:del, the Fallback's w:t kept (ED-009 section 3.6, P3)
  const d = await runParagraph();
  d.parentBody.package_.changeTrackingMode = 'TrackAll';
  d.search('flat')[0].delete();
  assert.equal(d.text, 'The Earth is \treally!');
  assert.match(await d.getXml(), /<w:del [^>]*><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:delText>flat<\/w:delText><\/mc:Choice><mc:Fallback><w:t>round<\/w:t><\/mc:Fallback><\/mc:AlternateContent><\/w:r><\/w:del>/);
  // all of its text deleted: the element goes, the run keeps what else it held
  const all = await runParagraph();
  all.search('flat\treally')[0].delete();
  assert.equal(all.text, 'The Earth is !');
  assert.doesNotMatch(await all.getXml(), /AlternateContent|round/);
  // a branch holding a drawing stays one object of no text (section 12.2 item 5): an edit beside it works
  const drawn = await WordprocessingMLPackage.createPackage();
  drawn.getMainDocumentPart().setXml(`<w:document xmlns:w="${W_NS}" xmlns:mc="${MCE}" xmlns:wps="${WPS}"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing/></mc:Choice><mc:Fallback><w:t>round</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const o = (await (await WordprocessingMLPackage.load(await drawn.save(), KEPT)).getBody()).paragraphs[0];
  assert.equal(o.text, '');
  o.insertText('The Earth is ', 'Start');
  assert.equal(o.text, 'The Earth is ');
});
