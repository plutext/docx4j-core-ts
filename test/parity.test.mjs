// CR-001 Phase B: the parity test.
//
// The goldens under test/golden/ are docx4j's answers for every fixture (see
// test/golden/README.md and test/java/README.md).  Two things happen here.
//
//  1. Every golden is *consumable*: its header names a fixture that is really there and really
//     that size, and every XML string in it parses and unmarshals through the objects facade
//     into a typed object.  If a golden ever carried markup the object model cannot type, the
//     port would meet it here rather than half way through a comparison.
//
//  2. **The comparisons.**  Step 2 turned on styles, paragraphs, runs and tables: for every
//     style, `getEffectivePPr(styleId)` and `getEffectiveRPr(styleId)`; the default paragraph
//     style id and the document defaults; for every paragraph of every story,
//     `getEffectivePPr(pPr)` and `getEffectiveParagraphMarkRPr(pPr)`, and for every run
//     `getEffectiveRPr(rPr, pPr)`; for every table `getEffectiveTableStyle(tblPr)` and
//     `reachesDefaultTableStyle(tblPr)`.  Steps 3 and 4 add numbering and fonts.
//
// **How a value is compared.**  Our object is marshalled to XML through the objects facade,
// and then *both* strings are unmarshalled, so both sides took the same path: attribute order,
// namespace prefixes and TYPE_NAME cannot differ, and what is compared is content.  One batch
// of each kind per fixture (all the paragraphs' pPr in one synthetic `w:document`, all the
// rPr in one `w:p`, all the table styles in one `w:styles`), so the round trip costs four
// marshals and four unmarshals per fixture rather than four per property.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdir, readFile, stat } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import {
  parseXml, unmarshalNode, marshalString, WordprocessingMLPackage,
  HeaderPart, FooterPart,
} from '../dist/index.mjs';
import { fixturesDir, plain } from './helpers.mjs';

const goldenDir = join(dirname(fileURLToPath(import.meta.url)), 'golden');
const names = (await readdir(goldenDir)).filter((n) => n.endsWith('.json')).sort();

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * The container each recorded element lives in.  The runtime resolves global elements only
 * (the objects package's `builders/wml` wraps fragments for the same reason), and none of
 * `w:pPr`, `w:rPr`, `w:ind`, `w:lvl`, `w:style` or `w:docDefaults` is global, so each is
 * unmarshalled inside the element the schema puts it in and picked back out.
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
    assert.notEqual(await fixturePath(header.fixture), null, `${header.fixture} is not under test/fixtures/`);
    const size = (await stat(await fixturePath(header.fixture))).size;
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

// ---------------------------------------------------------------- the comparisons

async function fixturePath(fixture) {
  for (const path of [join(fixturesDir, fixture), join(fixturesDir, 'parity', fixture)]) {
    try {
      await stat(path);
      return path;
    } catch {
      // try the other location
    }
  }
  return null;
}

/** True for a jsonix `{ name, value }` element pair. */
function isElement(value) {
  return typeof value === 'object' && value !== null && 'value' in value
    && typeof value.name === 'object' && value.name !== null && 'localPart' in value.name;
}

/**
 * Every object under `root` in document order: own enumerable properties in order, arrays in
 * order, element pairs unwrapped.  The unmarshalled tree keeps document order in its `content`
 * arrays, which is what the harness's TraversalUtil walk gives.
 */
function visit(root, fn) {
  const walk = (value) => {
    if (Array.isArray(value)) {
      for (const item of value) walk(item);
      return;
    }
    if (typeof value !== 'object' || value === null) return;
    if (isElement(value)) { walk(value.value); return; }
    if (fn(value) === false) return; // do not descend
    for (const key of Object.keys(value)) {
      if (key === 'PARENT' || key === 'TYPE_NAME') continue;
      walk(value[key]);
    }
  };
  walk(root);
}

/** Every `w:p` under the content, in document order, text boxes and tables included. */
function paragraphsOf(content) {
  const out = [];
  visit(content, (value) => {
    if (value.TYPE_NAME === 'org_docx4j_wml.P') out.push(value);
  });
  return out;
}

/**
 * Every `w:r` of a paragraph in order - those inside `w:ins`, `w:del`, `w:hyperlink`, `w:sdt`
 * and `w:smartTag` included - but not those of a paragraph nested inside it (a text box's),
 * which belong to that paragraph.  The harness's `collectRuns`.
 */
function runsOfParagraph(p) {
  const out = [];
  let first = true;
  visit(p, (value) => {
    if (value.TYPE_NAME === 'org_docx4j_wml.P' && !first) return false;
    first = false;
    if (value.TYPE_NAME === 'org_docx4j_wml.R') { out.push(value); return false; }
    return true;
  });
  return out;
}

/** Every `w:tbl` under the main document's content, in document order. */
function tablesOf(content) {
  const out = [];
  visit(content, (value) => {
    if (value.TYPE_NAME === 'org_docx4j_wml.Tbl') out.push(value);
  });
  return out;
}

/** The first property path at which two unmarshalled trees differ, or null. */
function firstDifference(ours, golden, path = '') {
  const a = ours === undefined ? null : ours;
  const b = golden === undefined ? null : golden;
  if (a === null || b === null || typeof a !== 'object' || typeof b !== 'object') {
    return a === b ? null : `${path || '.'}: ${JSON.stringify(a)} != ${JSON.stringify(b)}`;
  }
  if (Array.isArray(a) !== Array.isArray(b)) return `${path || '.'}: array/object`;
  if (Array.isArray(a)) {
    if (a.length !== b.length) return `${path || '.'}: ${a.length} items != ${b.length}`;
    for (let i = 0; i < a.length; i++) {
      const d = firstDifference(a[i], b[i], `${path}[${i}]`);
      if (d) return d;
    }
    return null;
  }
  for (const key of new Set([...Object.keys(a), ...Object.keys(b)])) {
    const d = firstDifference(a[key], b[key], path ? `${path}.${key}` : key);
    if (d) return d;
  }
  return null;
}

const P_ELEMENT = (pPr) => ({ name: { namespaceURI: W, localPart: 'p' }, value: { TYPE_NAME: 'org_docx4j_wml.P', ...(pPr === undefined ? {} : { pPr }) } });
const R_ELEMENT = (rPr) => ({ name: { namespaceURI: W, localPart: 'r' }, value: { TYPE_NAME: 'org_docx4j_wml.R', ...(rPr === undefined ? {} : { rPr }) } });

/**
 * Marshal our objects of one kind in one container, unmarshal that, unmarshal the goldens'
 * strings in the same container, and compare each pair with `plain()`.
 *
 * `items` are `{ label, ours, golden }`; `golden` is the marshalled string docx4j recorded, or
 * null/an `{ exception }` object.
 */
async function compareBatch(kind, items, report) {
  if (items.length === 0) return;

  // the goldens whose value is not a string: docx4j answered null, or threw
  const comparable = [];
  for (const item of items) {
    if (typeof item.golden === 'string') { comparable.push(item); continue; }
    if (item.golden === null || item.golden === undefined) {
      if (item.ours !== undefined && item.ours !== null) report(`${item.label}: we answer a ${kind}, docx4j answers null`);
      continue;
    }
    report(`${item.label}: docx4j threw ${item.golden.exception}: ${item.golden.message}`);
  }
  if (comparable.length === 0) return;

  let ourValues;
  let goldenValues;
  if (kind === 'pPr') {
    ourValues = await roundTrip(
      { name: { namespaceURI: W, localPart: 'document' }, value: { TYPE_NAME: 'org_docx4j_wml.Document', body: { TYPE_NAME: 'org_docx4j_wml.Body', content: comparable.map((i) => P_ELEMENT(i.ours)) } } },
      (v) => v.body.content.map((e) => e.value.pPr),
    );
    goldenValues = await unmarshalWrapped(
      `<w:document xmlns:w="${W}"><w:body>${comparable.map((i) => `<w:p>${i.golden}</w:p>`).join('')}</w:body></w:document>`,
      (v) => v.body.content.map((e) => e.value.pPr),
    );
  } else if (kind === 'rPr') {
    ourValues = await roundTrip(
      { name: { namespaceURI: W, localPart: 'p' }, value: { TYPE_NAME: 'org_docx4j_wml.P', content: comparable.map((i) => R_ELEMENT(i.ours)) } },
      (v) => v.content.map((e) => e.value.rPr),
    );
    goldenValues = await unmarshalWrapped(
      `<w:p xmlns:w="${W}">${comparable.map((i) => `<w:r>${i.golden}</w:r>`).join('')}</w:p>`,
      (v) => v.content.map((e) => e.value.rPr),
    );
  } else {
    ourValues = await roundTrip(
      { name: { namespaceURI: W, localPart: 'styles' }, value: { TYPE_NAME: 'org_docx4j_wml.Styles', style: comparable.map((i) => i.ours) } },
      (v) => v.style,
    );
    goldenValues = await unmarshalWrapped(
      `<w:styles xmlns:w="${W}">${comparable.map((i) => i.golden).join('')}</w:styles>`,
      (v) => v.style,
    );
  }

  for (let i = 0; i < comparable.length; i++) {
    const difference = firstDifference(plain(ourValues[i]), plain(goldenValues[i]));
    if (difference) report(`${comparable[i].label}: ${difference}`);
  }
}

async function roundTrip(element, pick) {
  const xml = await marshalString(element);
  const back = await unmarshalNode(parseXml(xml));
  return pick(back.value);
}

async function unmarshalWrapped(xml, pick) {
  const back = await unmarshalNode(parseXml(xml));
  return pick(back.value);
}

/** The block-level content of each story the golden records, in the harness's order. */
async function storiesOf(pkg, golden) {
  const main = pkg.getMainDocumentPart();
  const out = new Map();
  out.set('main', (await main.getContents()).body?.content ?? []);
  for (const part of main.headerParts) {
    if (!(part instanceof HeaderPart)) continue;
    out.set(`header:${part.sourceRelationship?.id}`, (await part.getContents()).content ?? []);
  }
  for (const part of main.footerParts) {
    if (!(part instanceof FooterPart)) continue;
    out.set(`footer:${part.sourceRelationship?.id}`, (await part.getContents()).content ?? []);
  }
  if (main.footnotesPart) {
    const content = [];
    for (const note of (await main.footnotesPart.getContents()).footnote ?? []) content.push(...(note.content ?? []));
    out.set('footnotes', content);
  }
  if (main.endnotesPart) {
    const content = [];
    for (const note of (await main.endnotesPart.getContents()).endnote ?? []) content.push(...(note.content ?? []));
    out.set('endnotes', content);
  }
  if (main.commentsPart) {
    const content = [];
    for (const comment of (await main.commentsPart.getContents()).comment ?? []) content.push(...(comment.content ?? []));
    out.set('comments', content);
  }
  // every story the golden has, we must have
  for (const story of Object.keys(golden.stories)) {
    assert.ok(out.has(story), `story ${story} is in the golden but not in the package`);
  }
  return out;
}

for (const name of names) {
  test(`parity ${name}: styles, paragraphs, runs and tables equal docx4j`, async () => {
    const golden = JSON.parse(await readFile(join(goldenDir, name), 'utf8'));
    const path = await fixturePath(golden.header.fixture);
    const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile(path)));
    const resolver = await pkg.getPropertyResolver();

    const differences = [];
    const report = (message) => differences.push(`${golden.header.fixture} ${message}`);
    const pPrItems = [];
    const rPrItems = [];
    const styleItems = [];

    // --- styles
    assert.equal(resolver.getDefaultParagraphStyleId() ?? null, golden.styles.defaultParagraphStyleId,
      `${golden.header.fixture} defaultParagraphStyleId`);
    pPrItems.push({ label: 'documentDefaults pPr', ours: resolver.getDocumentDefaultPPr(), golden: golden.styles.documentDefaults.pPr });
    rPrItems.push({ label: 'documentDefaults rPr', ours: resolver.getDocumentDefaultRPr(), golden: golden.styles.documentDefaults.rPr });
    for (const [styleId, entry] of Object.entries(golden.styles.byId)) {
      pPrItems.push({ label: `style ${styleId} effectivePPr`, ours: resolver.getEffectivePPr(styleId), golden: entry.effectivePPr });
      rPrItems.push({ label: `style ${styleId} effectiveRPr`, ours: resolver.getEffectiveRPr(styleId), golden: entry.effectiveRPr });
    }

    // --- paragraphs and runs, per story
    const stories = await storiesOf(pkg, golden);
    for (const [story, recorded] of Object.entries(golden.stories)) {
      const paragraphs = paragraphsOf(stories.get(story));
      if (paragraphs.length !== recorded.paragraphs.length) {
        const at = recorded.paragraphs.find((p, i) => (paragraphs[i]?.paraId ?? null) !== p.paraId);
        report(`${story}: ${paragraphs.length} paragraphs, the golden has ${recorded.paragraphs.length}`
          + (at ? `; first mismatch at index ${at.index} (${JSON.stringify(at.text)})` : ''));
        continue;
      }
      for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        const expected = recorded.paragraphs[i];
        const where = `${story}/${expected.index}`;
        pPrItems.push({ label: `${where} effectivePPr`, ours: resolver.getEffectivePPr(p.pPr), golden: expected.effectivePPr });
        rPrItems.push({ label: `${where} paragraphMarkRPr`, ours: resolver.getEffectiveParagraphMarkRPr(p.pPr), golden: expected.paragraphMarkRPr });
        const runs = runsOfParagraph(p);
        if (runs.length !== expected.runs.length) {
          report(`${where}: ${runs.length} runs, the golden has ${expected.runs.length}`);
          continue;
        }
        for (let r = 0; r < runs.length; r++) {
          rPrItems.push({
            label: `${where} run ${r} effectiveRPr`,
            ours: resolver.getEffectiveRPr(runs[r].rPr, p.pPr),
            golden: expected.runs[r].effectiveRPr,
          });
        }
      }
    }

    // --- tables
    const tables = tablesOf((await pkg.getMainDocumentPart().getContents()).body?.content ?? []);
    if (tables.length !== golden.tables.length) {
      report(`${tables.length} tables, the golden has ${golden.tables.length}`);
    } else {
      for (let i = 0; i < tables.length; i++) {
        const expected = golden.tables[i];
        styleItems.push({
          label: `table ${expected.index} effectiveTableStyle`,
          ours: resolver.getEffectiveTableStyle(tables[i].tblPr),
          golden: expected.effectiveTableStyle,
        });
        const reaches = resolver.reachesDefaultTableStyle(tables[i].tblPr);
        if (reaches !== expected.reachesDefaultTableStyle) {
          report(`table ${expected.index} reachesDefaultTableStyle: ${reaches} != ${expected.reachesDefaultTableStyle}`);
        }
      }
    }

    await compareBatch('pPr', pPrItems, report);
    await compareBatch('rPr', rPrItems, report);
    await compareBatch('style', styleItems, report);

    assert.deepEqual(differences, [], differences.join('\n'));
  });
}

// The comparisons still waiting on the step of CR-001 Phase B that writes the code they drive.
test.todo('numbering: Emulator numString, isBullet, ind, lvl, labelRPr, numRef and stateAfter equal the golden per story (step 3)');
test.todo('fontSpans: RunFontSelector.documentFontFor folds each run into the golden spans (step 4)');
test.todo('fonts: fontsInUse(), stylesInUse() and the IdentityPlusMapper decisions equal the golden (step 4)');
