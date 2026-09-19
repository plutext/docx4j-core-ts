// CR-001 Phase B, step 1: the parity test skeleton.
//
// The goldens under test/golden/ are docx4j's answers for every fixture (see
// test/golden/README.md and test/java/README.md).  Steps 2, 3 and 4 of the phase fill in
// the comparisons; what this file does today is prove that a golden is *consumable*: its
// header names a fixture that is really there and really that size, and every XML string
// in it parses with this package's `parseXml` and unmarshals through the objects facade
// into a typed object.  If a golden ever carried markup the object model cannot type, the
// port would meet it in step 2 with a resolver half written; here it is one failure with
// the file and the fragment named.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdir, readFile, stat } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { parseXml, unmarshalNode } from '../dist/index.mjs';
import { fixturesDir } from './helpers.mjs';

const goldenDir = join(dirname(fileURLToPath(import.meta.url)), 'golden');
const names = (await readdir(goldenDir)).filter((n) => n.endsWith('.json')).sort();

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * The container each recorded element lives in.  The runtime resolves global elements
 * only (the objects package's `builders/wml` wraps fragments for the same reason), and
 * none of `w:pPr`, `w:rPr`, `w:ind`, `w:lvl`, `w:style` or `w:docDefaults` is global, so
 * each is unmarshalled inside the element the schema puts it in and picked back out.
 */
const CONTAINERS = {
  pPr: ['<w:p>', '</w:p>', (v) => v.pPr],
  rPr: ['<w:r>', '</w:r>', (v) => v.rPr],
  ind: ['<w:p><w:pPr>', '</w:pPr></w:p>', (v) => v.pPr.ind],
  lvl: ['<w:numbering><w:abstractNum w:abstractNumId="0">', '</w:abstractNum></w:numbering>', (v) => v.abstractNum[0].lvl[0]],
  style: ['<w:styles>', '</w:styles>', (v) => v.style[0]],
  docDefaults: ['<w:styles>', '</w:styles>', (v) => v.docDefaults],
};

/** The local name of a fragment's root: the goldens are docx4j's marshalling, always `w:` prefixed. */
function rootLocalName(xml) {
  const match = /^<w:([A-Za-z0-9_.-]+)/.exec(xml);
  return match ? match[1] : null;
}

async function unmarshalFragment(xml) {
  const local = rootLocalName(xml);
  const container = CONTAINERS[local];
  assert.ok(container, `no container known for w:${local}`);
  const [open, close, pick] = container;
  const root = /^<w:([A-Za-z0-9_.-]+)/.exec(open)[1];
  const wrapped = `${open.replace(`<w:${root}`, `<w:${root} xmlns:w="${W}"`)}${xml}${close}`;
  const element = await unmarshalNode(parseXml(wrapped));
  return pick(element.value);
}

/** The keys of a golden whose value is a marshalled element (or null). Named rather than
 *  sniffed: a paragraph's `text` can itself begin with a '<' (NumberingIndents.docx sets
 *  its probe text to the markup being probed). */
const XML_KEYS = new Set([
  'pPr', 'rPr', 'docDefaults',
  'effectivePPr', 'effectiveRPr', 'paragraphMarkRPr',
  'ind', 'indResolved', 'lvl', 'labelRPr',
  'effectiveTableStyle',
]);

/** The golden's XML strings, deduplicated: a document's styles repeat the same `w:rPr`
 *  hundreds of times, and unmarshalling each once keeps this test to half a second. */
function xmlStrings(value, into) {
  if (Array.isArray(value)) {
    for (const item of value) xmlStrings(item, into);
  } else if (value && typeof value === 'object') {
    for (const [key, item] of Object.entries(value)) {
      if (typeof item === 'string') {
        if (XML_KEYS.has(key)) into.add(item);
      } else {
        xmlStrings(item, into);
      }
    }
  }
  return into;
}

test('parity goldens: present', () => {
  assert.ok(names.length >= 45, `${names.length} goldens`);
});

for (const name of names) {
  test(`parity golden ${name}: header, and every XML fragment unmarshals`, async () => {
    const golden = JSON.parse(await readFile(join(goldenDir, name), 'utf8'));

    // the header describes a fixture that is here, unchanged
    const header = golden.header;
    assert.equal(typeof header.docx4jCommit, 'string', 'docx4jCommit');
    assert.equal(typeof header.harnessVersion, 'string', 'harnessVersion');
    assert.match(header.date, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/, 'date');
    assert.equal(header.fixture, `${name.slice(0, -'.json'.length)}.docx`, 'fixture names this golden');
    const candidates = [join(fixturesDir, header.fixture), join(fixturesDir, 'parity', header.fixture)];
    let size = null;
    for (const path of candidates) {
      try {
        size = (await stat(path)).size;
        break;
      } catch {
        // try the other location
      }
    }
    assert.notEqual(size, null, `${header.fixture} is not under test/fixtures/`);
    assert.equal(size, header.fixtureBytes, `${header.fixture} is not the size the golden was made from`);

    // every property the golden records is markup this package can type
    const seen = xmlStrings(golden, new Set());
    assert.ok(seen.size > 0, 'the golden records some XML');
    for (const xml of seen) {
      let value;
      try {
        value = await unmarshalFragment(xml);
      } catch (e) {
        assert.fail(`${xml.slice(0, 200)}: ${e.message}`);
      }
      assert.ok(value && typeof value === 'object' && value.TYPE_NAME, `${xml.slice(0, 120)}: not a typed object`);
    }
  });
}

// The comparisons themselves, each waiting on the step of CR-001 Phase B that writes the
// code it would drive.  Named, not vague, so that the step knows what to turn on.
test.todo('styles: PropertyResolver.getEffectivePPr(styleId) and getEffectiveRPr(styleId) equal the golden (step 2)');
test.todo('paragraphs: getEffectivePPr(pPr) and getEffectiveParagraphMarkRPr(pPr) equal the golden per paragraph (step 2)');
test.todo('runs: getEffectiveRPr(rPr, pPr) equals the golden per run (step 2)');
test.todo('tables: getEffectiveTableStyle(tblPr) and reachesDefaultTableStyle equal the golden (step 2)');
test.todo('numbering: Emulator numString, isBullet, ind, lvl, labelRPr, numRef and stateAfter equal the golden per story (step 3)');
test.todo('fontSpans: RunFontSelector.documentFontFor folds each run into the golden spans (step 4)');
test.todo('fonts: fontsInUse(), stylesInUse() and the IdentityPlusMapper decisions equal the golden (step 4)');
