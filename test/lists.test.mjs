// CR-002 phase H: the list views (`List`, `ListItem`, and what they add to `Paragraph` and
// `Body`), specified in CR-002 section 3.9 and built on CR-001 Phase B's `Emulator`.
//
// The strongest test here is the last one: for every paragraph of every story of all 45
// parity goldens, `paragraph.listItem?.listString` must be the label docx4j recorded and
// `listItem.level` its `ilvl`. That holds the whole `listString` path - the story walk, the
// text-box rule, the header/footer shared state - to docx4j's own answers.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdir, readFile, stat } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import {
  WordprocessingMLPackage, ZipPartStore, Body, Paragraph, List,
  HeaderPart, FooterPart,
} from '../dist/index.mjs';
import { Word } from '../dist/office-js/index.mjs';
import { fixture, fixturesDir, bytesEqual } from './helpers.mjs';

const goldenDir = join(dirname(fileURLToPath(import.meta.url)), 'golden');

/** A created package with `count` paragraphs of the given texts, and its body. */
async function packageOf(...texts) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const paragraphs = texts.map((text) => body.insertParagraph(text, 'End'));
  return { pkg, body, paragraphs };
}

/** Every paragraph's label, or null where Word paints none. */
function labels(body) {
  return body.paragraphs.map((p) => (p.isListItem ? p.listItem.listString : null));
}

// ---------------------------------------------------------------- creating a list

test('startNewList: the numbering part is created, with docx4j\'s two default definitions', async () => {
  const { pkg, body, paragraphs } = await packageOf('One', 'Two', 'Three');
  assert.equal(pkg.numberingDefinitionsPart, undefined, 'a new package has no numbering part');

  const list = await paragraphs[0].startNewList();
  const part = pkg.numberingDefinitionsPart;
  assert.ok(part, 'the numbering part was added');
  assert.equal(part.partName.name, '/word/numbering.xml');
  assert.ok(pkg.getMainDocumentPart().relationshipsPart.getRelationshipsByType(
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering').length === 1,
  'and its relationship');

  // docx4j's default numbering is there: w:num 1 is the decimal set and w:num 2 the bullet one
  assert.equal(new List('1', body).levelTypes[0], 'Number');
  assert.equal(new List('2', body).levelTypes[0], 'Bullet');
  // and the new list is a third one, over its own w:abstractNum
  assert.equal(list.id, 3);
  assert.equal(list.levelTypes[0], 'Number');
  assert.equal(list.getLevelString(0), '%1.');
  assert.notEqual(list.abstractElement.abstractNumId, new List('1', body).abstractElement.abstractNumId);

  paragraphs[1].attachToList(list.id);
  paragraphs[2].attachToList(list.id);
  assert.deepEqual(labels(body), ['1.', '2.', '3.']);
  assert.deepEqual(body.lists.map((l) => l.id), [3]);

  // it survives a round trip
  const again = await WordprocessingMLPackage.load(await pkg.save());
  const reloaded = await again.getBody();
  assert.deepEqual(labels(reloaded), ['1.', '2.', '3.']);
});

test('startNewList({ bullet: true }) takes docx4j\'s bullet set', async () => {
  const { body, paragraphs } = await packageOf('One', 'Two');
  const list = await paragraphs[0].startNewList({ bullet: true });
  paragraphs[1].attachToList(list.id);
  assert.equal(list.levelTypes[0], 'Bullet');
  assert.deepEqual(labels(body), ['', '']);
  assert.equal(body.listLabels().get(paragraphs[0].p).isBullet, true);
});

test('attachToList and listItem.level: the nested labels Word paints', async () => {
  const { body, paragraphs } = await packageOf('One', 'One a', 'One b', 'Two');
  const list = await paragraphs[0].startNewList();
  for (const p of paragraphs.slice(1)) p.attachToList(list.id);
  paragraphs[1].listItem.level = 1;
  paragraphs[2].listItem.level = 1;
  assert.deepEqual(labels(body), ['1.', 'a.', 'b.', '2.']);
  assert.deepEqual(body.paragraphs.map((p) => p.listItem.level), [0, 1, 1, 0]);
  // the level is written as the paragraph's own w:ilvl
  assert.equal(paragraphs[1].p.pPr.numPr.ilvl.val, 1);
  assert.deepEqual(list.getLevelParagraphs(1).map((p) => p.text), ['One a', 'One b']);
});

test('detachFromList: the paragraph leaves the list, and the rest renumber', async () => {
  const { body, paragraphs } = await packageOf('One', 'Two', 'Three');
  const list = await paragraphs[0].startNewList();
  paragraphs[1].attachToList(list.id);
  paragraphs[2].attachToList(list.id);
  assert.deepEqual(labels(body), ['1.', '2.', '3.']);

  paragraphs[1].detachFromList();
  assert.equal(paragraphs[1].isListItem, false);
  assert.equal(paragraphs[1].p.pPr.numPr, undefined);
  assert.deepEqual(labels(body), ['1.', null, '2.']);
  assert.equal(paragraphs[1].listItemOrNullObject.isNullObject, true);
  assert.equal(paragraphs[1].listOrNullObject.isNullObject, true);
  assert.throws(() => paragraphs[1].listItem, /not a list item/);
  assert.throws(() => paragraphs[1].list, /not a list item/);
  assert.equal(paragraphs[0].listItemOrNullObject.listString, '1.', 'a real item is not a null object');
});

test('restart: a new w:num over the same w:abstractNum with a w:startOverride', async () => {
  const { pkg, body, paragraphs } = await packageOf('One', 'Two', 'Three', 'Four');
  const list = await paragraphs[0].startNewList();
  for (const p of paragraphs.slice(1)) p.attachToList(list.id);
  assert.deepEqual(labels(body), ['1.', '2.', '3.', '4.']);

  const restarted = paragraphs[2].restartList();
  paragraphs[3].attachToList(restarted.id);
  assert.deepEqual(labels(body), ['1.', '2.', '1.', '2.']);
  assert.equal(restarted.element.abstractNumId.val, list.element.abstractNumId.val,
    'the same abstract definition, so Word treats them as one list');
  assert.equal(restarted.element.lvlOverride[0].startOverride.val, 1);
  assert.deepEqual(body.lists.map((l) => l.id), [list.id, restarted.id]);

  const again = await WordprocessingMLPackage.load(await pkg.save());
  assert.deepEqual(labels(await again.getBody()), ['1.', '2.', '1.', '2.']);
});

// ---------------------------------------------------------------- the level writers

test('setLevelNumbering, setLevelIndents, setLevelAlignment and setLevelStartingNumber', async () => {
  const { body, paragraphs } = await packageOf('One', 'Two');
  const list = await paragraphs[0].startNewList();
  paragraphs[1].attachToList(list.id);

  list.setLevelNumbering(0, 'UpperRoman', ['(', 0, ')']);
  assert.equal(list.getLevelString(0), '(%1)');
  assert.deepEqual(labels(body), ['(I)', '(II)']);

  list.setLevelStartingNumber(0, 5);
  assert.deepEqual(labels(body), ['(V)', '(VI)']);

  list.setLevelIndents(0, 36, 18);
  const ind = list.abstractElement.lvl[0].pPr.ind;
  assert.equal(ind.left, 720);
  assert.equal(ind.hanging, 360);

  list.setLevelAlignment(0, 'Right');
  assert.equal(list.abstractElement.lvl[0].lvlJc.val, 'right');

  // the default format string numbers the level itself with a full stop
  list.setLevelNumbering(1, 'LowerLetter');
  assert.equal(list.getLevelString(1), '%2.');
  assert.equal(list.levelTypes[1], 'Number');
});

test('setLevelBullet: the character and the face Word uses, and Custom', async () => {
  const { body, paragraphs } = await packageOf('One', 'Two');
  const list = await paragraphs[0].startNewList();
  paragraphs[1].attachToList(list.id);

  list.setLevelBullet(0, 'Hollow');
  assert.equal(list.levelTypes[0], 'Bullet');
  assert.deepEqual(labels(body), ['o', 'o']);
  assert.equal(list.getLevelFont(0).name, 'Courier New');

  list.setLevelBullet(0, 'Custom', 0x2605, 'Arial');
  assert.deepEqual(labels(body), ['★', '★']);
  assert.equal(list.abstractElement.lvl[0].rPr.rFonts.ascii, 'Arial');
  assert.throws(() => list.setLevelBullet(0, 'Custom'), /needs a charCode/);
});

test('a level writer copies a shared w:abstractNum first, so the change stays local', async () => {
  const { body, paragraphs } = await packageOf('A one', 'A two', 'B one', 'B two');
  const first = await paragraphs[0].startNewList();
  paragraphs[1].attachToList(first.id);
  // restart() is the way two w:num share one w:abstractNum
  paragraphs[2].attachToList(first.id);
  const second = paragraphs[2].restartList();
  paragraphs[3].attachToList(second.id);
  const sharedAbstract = first.element.abstractNumId.val;
  assert.equal(second.element.abstractNumId.val, sharedAbstract);
  assert.deepEqual(labels(body), ['1.', '2.', '1.', '2.']);

  first.setLevelNumbering(0, 'UpperLetter');

  assert.notEqual(first.element.abstractNumId.val, sharedAbstract, 'the edited list took a copy');
  assert.equal(second.element.abstractNumId.val, sharedAbstract, 'the other one did not move');
  assert.equal(first.abstractElement.lvl[0].numFmt.val, 'upperLetter');
  assert.equal(second.abstractElement.lvl[0].numFmt.val, 'decimal');
  assert.equal(second.getLevelString(0), '%1.');
  assert.deepEqual(labels(body), ['A.', 'B.', '1.', '2.']);
  // the copy is a w:abstractNum of its own, with its own w:nsid
  assert.notEqual(first.abstractElement.nsid.val, second.abstractElement.nsid.val);
});

// ---------------------------------------------------------------- reading an item

test('siblingIndex counts within the level, from where it last restarted', async () => {
  const { body, paragraphs } = await packageOf('One', 'One a', 'One b', 'Two', 'Two a');
  const list = await paragraphs[0].startNewList();
  for (const p of paragraphs.slice(1)) p.attachToList(list.id);
  for (const i of [1, 2, 4]) paragraphs[i].listItem.level = 1;

  assert.deepEqual(labels(body), ['1.', 'a.', 'b.', '2.', 'a.']);
  assert.deepEqual(body.paragraphs.map((p) => p.listItem.siblingIndex), [0, 0, 1, 1, 0]);
  assert.deepEqual([...body.listLabels().values()].map((l) => l.siblingIndex), [0, 0, 1, 1, 0]);
});

test('getAncestor and getDescendants walk the list by level', async () => {
  const { paragraphs } = await packageOf('One', 'One a', 'One a i', 'One b', 'Two');
  const list = await paragraphs[0].startNewList();
  for (const p of paragraphs.slice(1)) p.attachToList(list.id);
  paragraphs[1].listItem.level = 1;
  paragraphs[2].listItem.level = 2;
  paragraphs[3].listItem.level = 1;

  assert.equal(paragraphs[1].listItem.getAncestor().text, 'One');
  assert.equal(paragraphs[2].listItem.getAncestor().text, 'One a');
  assert.equal(paragraphs[2].listItem.getAncestor(true).text, 'One a');
  assert.throws(() => paragraphs[0].listItem.getAncestor(), /has no ancestor/);

  assert.deepEqual(paragraphs[0].listItem.getDescendants().map((p) => p.text), ['One a', 'One a i', 'One b']);
  assert.deepEqual(paragraphs[0].listItem.getDescendants(true).map((p) => p.text), ['One a', 'One b']);
  assert.deepEqual(paragraphs[4].listItem.getDescendants().map((p) => p.text), []);
});

test('insertParagraph on a list adds an item at the level of the one it follows', async () => {
  const { body, paragraphs } = await packageOf('One', 'Two');
  const list = await paragraphs[0].startNewList();
  paragraphs[1].attachToList(list.id);
  const added = list.insertParagraph('Three', 'End');
  assert.equal(added.isListItem, true);
  assert.deepEqual(labels(body), ['1.', '2.', '3.']);
  assert.equal(added.text, 'Three');
});

test('isListItem is true through a numbered paragraph style, and false for w:numId 0', async () => {
  const { pkg, body, paragraphs } = await packageOf('Direct', 'Styled', 'Off');
  const list = await paragraphs[0].startNewList();

  // a "List Number" style carrying the w:numPr, as Word's built-in one does
  const styles = await pkg.getMainDocumentPart().styleDefinitionsPart.getContents();
  styles.style.push({
    TYPE_NAME: 'org_docx4j_wml.Style',
    styleId: 'ListNumber',
    type: 'paragraph',
    name: { TYPE_NAME: 'org_docx4j_wml.Style.Name', val: 'List Number' },
    pPr: {
      TYPE_NAME: 'org_docx4j_wml.PPr',
      numPr: {
        TYPE_NAME: 'org_docx4j_wml.PPrBase.NumPr',
        numId: { TYPE_NAME: 'org_docx4j_wml.PPrBase.NumPr.NumId', val: list.id },
        ilvl: { TYPE_NAME: 'org_docx4j_wml.PPrBase.NumPr.Ilvl', val: 0 },
      },
    },
  });
  pkg.refreshPropertyResolver();

  paragraphs[1].styleId = 'ListNumber';
  assert.equal(paragraphs[1].isListItem, true, 'numbered through the style');
  assert.equal(paragraphs[1].list.id, list.id);
  assert.equal(paragraphs[1].listItem.level, 0);
  assert.deepEqual(labels(body), ['1.', '2.', null]);

  // a w:numId of 0 on the paragraph turns the style's numbering off (ECMA-376 17.9.18)
  paragraphs[2].styleId = 'ListNumber';
  assert.equal(paragraphs[2].isListItem, true);
  paragraphs[2].detachFromList();
  assert.equal(paragraphs[2].p.pPr.numPr.numId.val, 0);
  assert.equal(paragraphs[2].isListItem, false);

  // a w:numId naming no w:num is not a list item either (a dangling id; CR-001 section 15.2)
  paragraphs[2].attachToList(99);
  assert.equal(paragraphs[2].isListItem, false);
});

// ---------------------------------------------------------------- change tracking (phase F)

test('a tracked attachToList records w:pPrChange, as the other property setters do', async () => {
  const { pkg, paragraphs } = await packageOf('One', 'Two');
  const list = await paragraphs[0].startNewList();
  pkg.author = { name: 'Jane Doe', initials: 'JD' };
  pkg.trackedChangeDate = new Date('2026-09-19T00:00:00Z');
  await pkg.setChangeTrackingMode('TrackAll');

  paragraphs[1].attachToList(list.id, 1);
  const change = paragraphs[1].p.pPr.pPrChange;
  assert.ok(change, 'w:pPrChange was recorded');
  assert.equal(change.author, 'Jane Doe');
  assert.equal(change.pPr.numPr, undefined, 'it holds the properties as they stood before');
  assert.equal(paragraphs[1].p.pPr.numPr.numId.val, list.id);

  // level and detachFromList go the same way
  const before = paragraphs[1].p.pPr.pPrChange;
  paragraphs[1].listItem.level = 2;
  assert.equal(paragraphs[1].p.pPr.pPrChange, before, 'one w:pPrChange per paragraph, as Word writes');
  assert.equal(paragraphs[1].p.pPr.numPr.ilvl.val, 2);
});

// ---------------------------------------------------------------- the Word shim (phase I)

test('the Word shim: lists through Word.run, its enums and its null objects', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  await Word.run(pkg, async (context) => {
    const body = context.document.body;
    const one = body.insertParagraph('One', 'End');
    const two = body.insertParagraph('Two', 'End');
    const list = await one.startNewList();
    two.attachToList(list.id, 0);

    assert.equal(list.levelTypes[0], Word.ListLevelType.number);
    assert.equal(two.listItem.listString, '2.');
    assert.equal(two.listItem.siblingIndex, 1);
    assert.deepEqual(body.lists.items.map((l) => l.id), [list.id]);
    assert.equal(body.lists.getFirst().id, list.id);
    assert.deepEqual(one.listItem.getDescendants().items, []);

    list.setLevelNumbering(0, Word.ListNumbering.upperRoman);
    assert.equal(two.listItem.listString, 'II.');
    list.setLevelBullet(1, Word.ListBullet.square);
    assert.equal(list.levelTypes[1], Word.ListLevelType.bullet);

    assert.equal(body.insertParagraph('x', 'End').listOrNullObject.isNullObject, true);
    assert.equal(list.isNullObject, false);
    assert.throws(() => list.getLevelPicture(0), /Word\.List\.getLevelPicture is not supported/);
    for (const member of ['List.setLevelBullet', 'ListItem.listString', 'Paragraph.startNewList', 'Body.lists']) {
      assert.ok(Word.supported.has(member), member);
    }
  });
});

// ---------------------------------------------------------------- the round trip

test('a document whose lists are only read is still saved byte for byte', async () => {
  const bytes = await fixture(join('parity', 'numbering-lvlrestart.docx'));
  const pkg = await WordprocessingMLPackage.load(bytes);
  const body = await pkg.getBody();

  assert.ok(body.lists.length > 0, 'the document has lists');
  assert.ok(body.paragraphs.some((p) => p.isListItem), 'and list items');
  const read = labels(body);
  assert.deepEqual([...body.listLabels().values()].map((l) => l.listString), read.filter((l) => l !== null));
  assert.equal(pkg.numberingDefinitionsPart.isUnmarshalled, false, 'the numbering part stays unread');

  const saved = await pkg.save();
  const src = new ZipPartStore(bytes);
  const out = new ZipPartStore(saved);
  for (const name of ['word/numbering.xml', 'word/styles.xml']) {
    assert.ok(bytesEqual(await src.load(name), await out.load(name)), `${name} byte-identical`);
  }
});

// ---------------------------------------------------------------- against the parity goldens

const goldenNames = (await readdir(goldenDir)).filter((n) => n.endsWith('.json')).sort();

async function fixturePath(name) {
  for (const path of [join(fixturesDir, name), join(fixturesDir, 'parity', name)]) {
    try {
      await stat(path);
      return path;
    } catch {
      // try the other location
    }
  }
  return null;
}

/** True for a jsonix `{ name, value }` element pair (test/parity.test.mjs's). */
function isElement(value) {
  return typeof value === 'object' && value !== null && 'value' in value
    && typeof value.name === 'object' && value.name !== null && 'localPart' in value.name;
}

/**
 * A `Paragraph` view over every `w:p` under `root` in document order - a text box's included,
 * which `Body.paragraphs` does not reach - so that this compares the same paragraphs, in the
 * same order, as `test/parity.test.mjs` does.
 */
function paragraphViews(root, body) {
  const out = [];
  const walk = (value, container) => {
    if (Array.isArray(value)) {
      for (const item of value) walk(item, value);
      return;
    }
    if (typeof value !== 'object' || value === null) return;
    if (isElement(value)) {
      if (value.value && value.value.TYPE_NAME === 'org_docx4j_wml.P') out.push(new Paragraph(value, container, body));
      walk(value.value, container);
      return;
    }
    for (const key of Object.keys(value)) {
      if (key === 'PARENT' || key === 'TYPE_NAME') continue;
      walk(value[key], container);
    }
  };
  walk(root, []);
  return out;
}

/** The body view and the tree of each story the golden records (test/parity.test.mjs's map). */
async function storiesOf(pkg) {
  const main = pkg.getMainDocumentPart();
  const out = new Map();
  out.set('main', [await main.getBody(), await main.getContents()]);
  for (const part of main.headerParts) {
    if (part instanceof HeaderPart) out.set(`header:${part.sourceRelationship?.id}`, [await part.getBody(), await part.getContents()]);
  }
  for (const part of main.footerParts) {
    if (part instanceof FooterPart) out.set(`footer:${part.sourceRelationship?.id}`, [await part.getBody(), await part.getContents()]);
  }
  for (const [story, member] of [['footnotes', 'footnotesPart'], ['endnotes', 'endnotesPart'], ['comments', 'commentsPart']]) {
    const part = main[member];
    // the notes and comments parts have no Body of their own; the view only needs the part,
    // since a list label is counted over the part's whole story
    if (part) out.set(story, [new Body(part, {}, story, pkg), await part.getContents()]);
  }
  return out;
}

test('listString and level equal docx4j\'s labels on every paragraph of all 45 goldens', async () => {
  assert.ok(goldenNames.length >= 45, `${goldenNames.length} goldens`);
  const differences = [];
  let checked = 0;
  let numbered = 0;

  for (const name of goldenNames) {
    const golden = JSON.parse(await readFile(join(goldenDir, name), 'utf8'));
    const path = await fixturePath(golden.header.fixture);
    const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile(path)));
    await pkg.getPropertyResolver();
    const stories = await storiesOf(pkg);

    for (const [story, recorded] of Object.entries(golden.stories)) {
      const entry = stories.get(story);
      if (!entry) { differences.push(`${name} ${story}: not in the package`); continue; }
      const paragraphs = paragraphViews(entry[1], entry[0]);
      if (paragraphs.length !== recorded.paragraphs.length) {
        differences.push(`${name} ${story}: ${paragraphs.length} paragraphs, the golden has ${recorded.paragraphs.length}`);
        continue;
      }
      for (let i = 0; i < paragraphs.length; i++) {
        checked++;
        const where = `${name} ${story}/${recorded.paragraphs[i].index}`;
        const expected = recorded.paragraphs[i].numbering;
        if (expected) numbered++;
        const item = paragraphs[i].isListItem ? paragraphs[i].listItem : undefined;
        const ours = item ? item.listString : null;
        const theirs = expected?.numString ?? null;
        if (ours !== theirs) differences.push(`${where} listString: ${JSON.stringify(ours)} != ${JSON.stringify(theirs)}`);
        const ourLevel = item ? String(item.level) : null;
        const theirLevel = expected?.ilvl ?? null;
        if (ourLevel !== theirLevel) differences.push(`${where} level: ${ourLevel} != ${theirLevel}`);
      }
    }
  }

  assert.deepEqual(differences, [], differences.slice(0, 20).join('\n'));
  assert.equal(checked, 618, 'every paragraph of every story');
  assert.equal(numbered, 183, 'and the numbered ones docx4j recorded');
});

// CR-002 section 19: the list definition verbs off the package, which need no paragraph. The
// editor's case is a document model of its own, holding a w:numId and writing w:numPr itself,
// where going through Paragraph.startNewList() would mean projecting a tree and projecting back.
test('pkg.numbering.newList: a definition without a paragraph, and the part created when absent', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  assert.equal(pkg.getMainDocumentPart().numberingDefinitionsPart, undefined);

  const numId = await pkg.numbering.newList();
  assert.equal(typeof numId, 'string');
  const part = pkg.getMainDocumentPart().numberingDefinitionsPart;
  assert.ok(part, 'the numbering part was created, as startNewList does it');

  // the same definitions Paragraph.startNewList() builds: a copy of docx4j's decimal set
  const definition = part.definitions.list(numId);
  assert.ok(definition, 'the new w:num is in the definitions');
  assert.equal(definition.level('0')?.isBullet, false);
  const bulletId = await pkg.numbering.newList({ bullet: true });
  assert.notEqual(bulletId, numId, 'a second call gives a second definition');
  assert.equal(part.definitions.list(bulletId)?.level('0')?.isBullet, true);

  // nothing was attached: no paragraph is a list item
  assert.deepEqual(pkg.body.paragraphs.filter((p) => p.isListItem), []);

  // and the caller writing w:numPr itself gets a numbered paragraph
  const paragraph = pkg.body.insertParagraph('one', 'End');
  paragraph.attachToList(numId, 0);
  assert.equal(paragraph.isListItem, true);
  assert.equal(paragraph.listItem.listString, '1.');
});

test('pkg.numbering.newList: an existing numbering part is used, and the result survives a save', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('invoice.docx'));
  const numId = await pkg.numbering.newList();
  const part = pkg.getMainDocumentPart().numberingDefinitionsPart;
  assert.ok(part.definitions.list(numId), 'added to the document\'s own numbering part');

  const back = await WordprocessingMLPackage.load(await pkg.save());
  await back.getPropertyResolver();
  const reloaded = back.getMainDocumentPart().numberingDefinitionsPart.definitions.list(numId);
  assert.ok(reloaded, 'the definition is in the saved document');
  assert.equal(reloaded.level('0')?.isBullet, false);
});

test('pkg.numbering.restart: List.restart() without a Body', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const numId = await pkg.numbering.newList();
  const restarted = pkg.numbering.restart(numId);
  assert.notEqual(restarted, numId);

  const part = pkg.getMainDocumentPart().numberingDefinitionsPart;
  const num = part.contents.num.find((n) => String(n.numId) === restarted);
  assert.equal(num.abstractNumId.val, part.contents.num.find((n) => String(n.numId) === numId).abstractNumId.val,
    'the same abstract definition: counters are shared by it');
  assert.equal(num.lvlOverride[0].ilvl, 0);
  assert.equal(num.lvlOverride[0].startOverride.val, 1, 'and a startOverride is what restarts it');

  // the two lists number independently, which is the point of the override
  const first = pkg.body.insertParagraph('a', 'End');
  const second = pkg.body.insertParagraph('b', 'End');
  const third = pkg.body.insertParagraph('c', 'End');
  first.attachToList(numId, 0);
  second.attachToList(numId, 0);
  third.attachToList(restarted, 0);
  assert.deepEqual([first, second, third].map((p) => p.listItem.listString), ['1.', '2.', '1.']);
});

test('pkg.numbering.restart: what it says when there is nothing to restart', async () => {
  const empty = await WordprocessingMLPackage.createPackage();
  assert.throws(() => empty.numbering.restart('1'), /no numbering part/);
  const pkg = await WordprocessingMLPackage.createPackage();
  await pkg.numbering.newList();
  assert.throws(() => pkg.numbering.restart('99'), /No w:num for numId 99/);
});
