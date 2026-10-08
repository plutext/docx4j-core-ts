// CR-006 section 5: selectObjects held to docx4j's getJAXBNodesViaXPath.
//
// test/golden/xpath/ holds, per fixture, what docx4j answers for each of the XPath harness's
// expressions over the main document part (test/java/README.md, XPathHarness): every hit's path,
// the kind of node, and the class of the object docx4j's binder associates with it. The same
// expressions go through pkg.selectObjects here, which must return the same hits in the same order:
//
//  - the same paths. They are spelled with docx4j's prefixes, which are the objects package's too;
//    that is asserted first, so that a prefix which ever came to differ says so, rather than
//    showing as every path of a fixture differing;
//  - the same kind of node;
//  - for an element, an object of the class docx4j gives, which is the object's TYPE_NAME (the
//    harness spells a Java class that way: package dots as underscores, nested classes dotted),
//    and which is an object of the part's own tree, found by walking it - not a copy;
//  - for an attribute or a text node, where docx4j's binder associates nothing, the object of the
//    element that owns it. That is CR-006 section 2's deliberate departure, measured rather than
//    assumed: the owner's class must be the one docx4j gives that element wherever an expression
//    of the same golden selected it.
//
// A fixture whose main part marshals an mc:AlternateContent is recorded without hits: docx4j keeps
// the element and both its branches in the tree, this package resolves the choice before it
// unmarshals, and the two would be answering about different documents (CR-006 section 8 item 3).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdir, readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { WordprocessingMLPackage, NAMESPACE_PREFIXES, walk } from '../dist/index.mjs';
import { fixturesDir } from './helpers.mjs';

const goldenDir = join(dirname(fileURLToPath(import.meta.url)), 'golden', 'xpath');
const names = (await readdir(goldenDir)).filter((n) => n.endsWith('.json')).sort();
const goldens = await Promise.all(names.map(async (n) => JSON.parse(await readFile(join(goldenDir, n), 'utf8'))));

/**
 * The fixtures whose main part has an mc:AlternateContent - a text box, each of them - and so no
 * hits. Listed, so that a new fixture joining them is a decision rather than a silent loss of
 * coverage. tables-textbox-in-cell joined with CR-007's table fixtures (2026-10-08): its text box
 * is the point of the fixture.
 */
const EXCLUDED = ['loadAndSave.docx', 'numbering-stories-coretests.docx', 'numbering-stories.docx', 'tables-textbox-in-cell.docx'];

async function fixtureBytes(name) {
  for (const path of [join(fixturesDir, name), join(fixturesDir, 'parity', name)]) {
    try { return new Uint8Array(await readFile(path)); } catch { /* try the next */ }
  }
  throw new Error(`fixture ${name} is in no fixtures directory`);
}

function kindOf(node) {
  if (node.nodeType === 1) return 'element';
  if (node.nodeType === 2) return 'attribute';
  if (node.nodeType === 3 || node.nodeType === 4) return 'text';
  return `type ${node.nodeType}`;
}

/** The class of a selected object as the golden spells it: its TYPE_NAME, or `dom` for DOM content. */
function classOf(object) {
  if (object === undefined || object === null) return null;
  if (typeof object.nodeType === 'number') return 'dom';
  return object.TYPE_NAME ?? `(${typeof object}, no TYPE_NAME)`;
}

test('xpath goldens: present, one harness version, and the exclusions are the known ones', () => {
  assert.ok(goldens.length >= 40, `${goldens.length} goldens`);
  for (const golden of goldens) assert.equal(golden.header.harnessVersion, '1', golden.header.fixture);
  const excluded = goldens.filter((g) => !g.hits);
  assert.deepEqual(excluded.map((g) => g.header.fixture).sort(), EXCLUDED);
  for (const golden of excluded) assert.ok(golden.header.alternateContent > 0, golden.header.fixture);
});

test("xpath goldens: docx4j's prefixes are the objects package's", () => {
  for (const golden of goldens) {
    for (const [prefix, uri] of Object.entries(golden.header.prefixes)) {
      assert.equal(NAMESPACE_PREFIXES[uri], prefix, `${golden.header.fixture}: ${uri}`);
    }
  }
});

for (const golden of goldens.filter((g) => g.hits)) {
  const fixture = golden.header.fixture;
  test(`selectObjects agrees with docx4j: ${fixture}`, async () => {
    const pkg = await WordprocessingMLPackage.load(await fixtureBytes(fixture));
    // every typed object of the part's tree, which is what a selected object must be one of
    const inTree = new Set();
    walk(await pkg.getMainDocumentPart().getContents(), (value) => { inTree.add(value); });
    // the class docx4j gives each element any expression selected, by path: what the owner of an
    // attribute or a text node is checked against
    const classAt = new Map();
    for (const hits of Object.values(golden.hits)) {
      for (const hit of hits) if (hit.node === 'element') classAt.set(hit.path, hit.object);
    }
    for (const [xpath, expected] of Object.entries(golden.hits)) {
      const bounds = await pkg.selectObjects(xpath);
      const at = `${fixture} ${xpath}`;
      assert.deepEqual(bounds.map((b) => b.path), expected.map((h) => h.path), `${at}: the hits, in docx4j's order`);
      bounds.forEach((bound, i) => {
        const hit = expected[i];
        assert.equal(kindOf(bound.node), hit.node, `${at}: ${hit.path}`);
        if (hit.object !== null) {
          assert.equal(classOf(bound.object), hit.object, `${at}: the object at ${hit.path}`);
          assert.ok(inTree.has(bound.object), `${at}: the object at ${hit.path} is the tree's own`);
        } else {
          // docx4j associates nothing; this package gives the owning element's object
          assert.notEqual(hit.node, 'element', `${at}: docx4j associates no object with the element ${hit.path}`);
          const owner = hit.path.replace(/\/(?:@[^/]+|text\(\))$/, '');
          const ownerClass = classAt.get(owner);
          if (ownerClass !== undefined) assert.equal(classOf(bound.object), ownerClass, `${at}: the owner of ${hit.path}`);
          else assert.ok(classOf(bound.object), `${at}: an object for ${hit.path}`);
        }
      });
    }
  });
}
