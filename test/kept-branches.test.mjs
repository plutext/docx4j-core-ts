// CR-010 phase 3: the content API over a kept mc:AlternateContent (loaded with mcePreprocess: false).
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
  UNDERSTOOD_PREFIXES, selectedBranchOf, deepCopy,
} from '../dist/index.mjs';
import { mcBranchOf } from '@docx4j/generated-objects-ts/builders/wml';
import { fixture } from './helpers.mjs';

const MCE = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const KEPT = { mcePreprocess: false };
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
  const pkg = await WordprocessingMLPackage.load(await fixture(name), kept ? KEPT : undefined);
  return { pkg, body: await pkg.getBody() };
}

test('UNDERSTOOD_PREFIXES: what load understands, by conventional prefix; selectedBranchOf reads the branch load resolves to', () => {
  const byUri = new Map([...UNDERSTOOD_NAMESPACES].filter((uri) => NAMESPACE_PREFIXES[uri]).map((uri) => [uri, NAMESPACE_PREFIXES[uri]]));
  assert.deepEqual([...UNDERSTOOD_PREFIXES].sort(), [...new Set(byUri.values())].sort());
  for (const prefix of ['w', 'wps', 'wp14', 'w14', 'cx1', 'cx2', 'a14']) assert.ok(UNDERSTOOD_PREFIXES.has(prefix), prefix);
  // conventional prefixes the objects package knows and load does not understand (CR-010 section 11.5)
  for (const prefix of ['w16du', 'x14ac', 'cx3']) {
    assert.ok(Object.values(NAMESPACE_PREFIXES).includes(prefix), `${prefix} is conventional`);
    assert.ok(!UNDERSTOOD_PREFIXES.has(prefix), `${prefix} is not understood`);
  }
  const choiceItems = [];
  const fallbackItems = [];
  const element = (requires) => ({
    TYPE_NAME: AC,
    choice: [{ TYPE_NAME: `${AC}.Choice`, requires, any: choiceItems }],
    fallback: { TYPE_NAME: `${AC}.Fallback`, any: fallbackItems },
  });
  assert.equal(selectedBranchOf(element('wps')), choiceItems);
  assert.equal(selectedBranchOf(element('w16du')), fallbackItems, 'load takes the Fallback here');
  assert.equal(mcBranchOf(element('w16du')), choiceItems, "mcBranchOf's own default would take the Choice");
  assert.equal(selectedBranchOf(element('wps w16du')), fallbackItems);
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
    assert.equal(pictures[0].imageFormat, 'Png');
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
    const pkg = await WordprocessingMLPackage.load(understood, kept ? KEPT : undefined);
    const pictures = (await pkg.getBody()).inlinePictures;
    assert.equal(pictures.length, 1, `${kept ? 'kept' : 'resolved'}: the Choice's picture, once`);
    assert.equal(pictures[0].altTextDescription, '', "the Choice's drawing, not the Fallback's");
    assert.equal(await pictures[0].getBase64(), PNG);
    assert.equal(pictures[0].holder?.value.TYPE_NAME, kept ? AC : undefined);
  }
  // with the Choice's prefix not understood, load takes the Fallback, and a kept load reads it too
  const fallback = await pictureInBranch('w16du');
  for (const kept of [false, true]) {
    const pictures = (await (await WordprocessingMLPackage.load(fallback, kept ? KEPT : undefined)).getBody()).inlinePictures;
    assert.deepEqual(pictures.map((p) => p.altTextDescription), ['the Fallback'], kept ? 'kept' : 'resolved');
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
