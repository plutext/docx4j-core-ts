// CR-012 phase 6 (D8 c): the text model against docx4j's answers. test/golden/text/ is written by test/java's
// TextHarness (docx4j 17.3.4's org.docx4j.model.text: Texts.of and PlainText.of) for every WordprocessingML
// document under test/fixtures/ and test/text-extraction/, in four readings: the three views, and one with every
// story and every optional item included. Each reading holds the block model (each block's address, story,
// kind, fields, runs and rows; the comments; the warnings) and the plain text. Here the same documents are read
// through Texts.of and PlainText.of with the same options, written by the harness's rule (a value only where it
// is not its default), and compared. Parity is zero differences (CLAUDE.md): a difference is investigated, and
// fixed in the port unless docx4j is plainly wrong.
process.env.TZ = 'UTC';

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { isDeepStrictEqual } from 'node:util';
import { WordprocessingMLPackage, Texts, PlainText, TextOptions } from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const goldenDir = join(here, 'golden', 'text');

async function goldensIn(dir) {
  const out = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    if (entry.isDirectory()) out.push(...(await goldensIn(join(dir, entry.name))));
    else if (entry.name.endsWith('.json')) out.push(join(dir, entry.name));
  }
  return out;
}
const goldens = (await goldensIn(goldenDir)).map((file) => relative(goldenDir, file).split('\\').join('/')).sort();

// ------------------------------------------------------------------ the harness's rule, on this side

/** Puts a value unless it is a default: undefined, null, false, or an empty array (TextHarness.put). */
function put(m, key, value) {
  if (value === undefined || value === null || value === false) return;
  if (Array.isArray(value) && value.length === 0) return;
  m[key] = value;
}

function block(b) {
  const m = { address: b.address, kind: b.kind, story: b.story };
  put(m, 'paraId', b.paraId);
  put(m, 'style', b.style);
  put(m, 'hidden', b.hidden);
  put(m, 'placeholder', b.placeholder);
  put(m, 'decorative', b.decorative);
  if (b.revisions !== 0) m.revisions = b.revisions;
  put(m, 'indentLeft', b.indentLeft);
  put(m, 'indentHanging', b.indentHanging);
  put(m, 'indentFirstLine', b.indentFirstLine);
  put(m, 'alignment', b.alignment);
  put(m, 'monospace', b.monospace);
  put(m, 'listLabel', b.listLabel);
  put(m, 'outlineLevel', b.outlineLevel);
  if (b.numbering) m.numbering = numbering(b.numbering);
  put(m, 'contextualSpacing', b.contextualSpacing);
  put(m, 'bottomBorder', b.bottomBorder);
  put(m, 'noteType', b.noteType);
  if (b.noteNumber !== 0) m.noteNumber = b.noteNumber;
  put(m, 'noteId', b.noteId);
  put(m, 'simpleForm', b.simpleForm);
  if (b.kind === 'CHART') {
    put(m, 'chartTitle', b.chartTitle);
    put(m, 'chartHeader', [...b.chartHeader]);
    put(m, 'chartRows', b.chartRows.map((row) => [...row]));
  }
  put(m, 'runs', b.runs.map(run));
  put(m, 'fields', b.fields.map(field));
  put(m, 'rows', b.rows.map(row));
  return m;
}

function run(r) {
  const m = { kind: r.kind };
  put(m, 'text', r.text);
  if (r.revision) {
    const v = { type: r.revision.type };
    put(v, 'author', r.revision.author);
    put(v, 'date', r.revision.date);
    m.revision = v;
  }
  put(m, 'hidden', r.hidden);
  put(m, 'placeholder', r.placeholder);
  put(m, 'noteType', r.noteType);
  if (r.noteNumber !== 0) m.noteNumber = r.noteNumber;
  put(m, 'noteId', r.noteId);
  put(m, 'commentId', r.commentId);
  if (r.image) {
    const i = {};
    put(i, 'descr', r.image.descr);
    put(i, 'name', r.image.name);
    put(i, 'title', r.image.title);
    put(i, 'decorative', r.image.decorative);
    put(i, 'vml', r.image.vml);
    if (r.image.part) i.part = r.image.part.partName.name;
    m.image = i;
  }
  if (r.description) m.description = { text: r.description.text, provenance: r.description.provenance };
  put(m, 'progId', r.progId);
  if (r.pageNumber !== 0) m.pageNumber = r.pageNumber;
  put(m, 'bold', r.bold);
  put(m, 'italic', r.italic);
  put(m, 'strike', r.strike);
  put(m, 'monospace', r.monospace);
  put(m, 'link', r.link);
  put(m, 'math', r.math);
  put(m, 'displayMath', r.displayMath);
  return m;
}

function field(f) {
  const m = { instr: f.instr };
  put(m, 'result', f.result);
  put(m, 'updated', f.updated);
  return m;
}

function numbering(n) {
  const m = { numId: n.numId, level: n.level };
  put(m, 'format', n.format);
  put(m, 'levelText', n.levelText);
  m.start = n.start;
  put(m, 'value', n.value);
  put(m, 'abstractNumId', n.abstractNumId);
  put(m, 'bullet', n.bullet);
  put(m, 'label', n.label);
  return m;
}

function row(r) {
  const m = { address: r.address };
  put(m, 'header', r.header);
  put(m, 'cells', r.cells.map((c) => {
    const cell = { address: c.address };
    if (c.gridSpan !== 1) cell.gridSpan = c.gridSpan;
    put(cell, 'mergeContinue', c.mergeContinue);
    put(cell, 'mergeRestart', c.mergeRestart);
    put(cell, 'blocks', c.blocks.map(block));
    return cell;
  }));
  return m;
}

function comment(c) {
  const m = { id: c.id };
  put(m, 'author', c.author);
  put(m, 'initials', c.initials);
  put(m, 'date', c.date);
  put(m, 'parentId', c.parentId);
  put(m, 'resolved', c.resolved);
  put(m, 'blocks', c.blocks.map(block));
  put(m, 'replies', c.replies.map((reply) => reply.id));
  return m;
}

function warning(w) {
  const m = { kind: w.kind };
  put(m, 'address', w.address);
  put(m, 'branch', w.branch);
  put(m, 'contentType', w.contentType);
  put(m, 'partName', w.partName);
  return m;
}

/** A reading's body as the harness writes it. */
function body(result) {
  const m = { plainText: PlainText.of(result), blocks: result.blocks.map(block) };
  put(m, 'comments', [...result.comments.values()].map(comment));
  put(m, 'warnings', result.warnings.map(warning));
  return m;
}

/** A reading's body in the golden, following `sameAs`. */
function goldenBody(golden, name) {
  const r = golden.readings[name];
  if (r.sameAs !== undefined) return goldenBody(golden, r.sameAs);
  const { options, ...rest } = r;
  return rest;
}

// ------------------------------------------------------------------ run boundaries (CR-012 section 15.5)

const FORMAT = ['hidden', 'placeholder', 'bold', 'italic', 'strike', 'monospace', 'link', 'math', 'displayMath'];

/**
 * A block's runs joined by docx4j's rule (Extractor.addText): two neighbouring text or field-code runs of one kind
 * whose revision, hidden and placeholder flags and formatting (TextRun.sameFormat) agree are one run, an
 * equation's never. This package keeps such runs apart where Paragraph.text holds text between them that the
 * model leaves out, or where the second's text has no officeOffset (section 15.5: the text is the same, only the
 * run boundaries differ); joined on both sides, the two compare everything but the boundaries.
 */
function joinRuns(runs) {
  const out = [];
  for (const r of runs) {
    const last = out[out.length - 1];
    if (last && (r.kind === 'TEXT' || r.kind === 'FIELD_CODE') && last.kind === r.kind && !r.math && !last.math
      && isDeepStrictEqual(last.revision, r.revision) && FORMAT.every((k) => last[k] === r[k])) {
      out[out.length - 1] = { ...last, text: (last.text ?? '') + (r.text ?? '') };
    } else {
      out.push(r);
    }
  }
  return out;
}

/** A block, its cells' blocks and its runs, with the runs joined. */
function joined(b) {
  const m = { ...b };
  if (b.runs) m.runs = joinRuns(b.runs);
  if (b.rows) m.rows = b.rows.map((row) => ({ ...row, ...(row.cells ? { cells: row.cells.map((c) => ({ ...c, ...(c.blocks ? { blocks: c.blocks.map(joined) } : {}) })) } : {}) }));
  return m;
}

/** A body with every block's runs joined. */
function joinedBody(b) {
  return { ...b, blocks: b.blocks.map(joined), ...(b.comments ? { comments: b.comments.map((c) => ({ ...c, ...(c.blocks ? { blocks: c.blocks.map(joined) } : {}) })) } : {}) };
}

// ------------------------------------------------------------------ what phase 5 and REQ-303 leave (CR-012 section 20)

/**
 * The documents holding content that a statement this package doesn't implement yet governs, as the suite leaves
 * those statements pending (CR-012 sections 15.6 and 16.5): the same 20 documents section 16.5's Markdown sweep
 * found. Each flag sets aside only what its statement governs, on both sides, and the rest of the document is
 * compared:
 * - `math` (phase 5b, REQ-604): an equation's run text, LaTeX in docx4j and its `m:t` text here, masked;
 * - `latexOnly` (phase 5b): the blocks at these addresses, whole, for an equation whose original and accepted forms
 *   differ only in what the LaTeX shows, so that markup writes it as a deletion and an insertion in docx4j and as
 *   one run here, by the same rule over another text;
 * - `charts` and `smartart` (phase 5c, REQ-511), `altchunks` (phase 5e, REQ-513): the blocks docx4j reads at a
 *   `chart:`, `smartart:` or `altchunk:` address, which this package doesn't make, and every
 *   `altchunk.not_expanded` warning;
 * - `req303` (REQ-303 by URI, section 15.3: `Requires` matched by namespace, not prefix): the blocks at and under
 *   these addresses, check 50's rebound prefixes and the probe's case 6, and the warnings there.
 * The plain text is compared with the equations masked, in each reading where nothing else set aside differs. Each
 * flag must still set aside something that differs, in some reading, so that a phase arriving fails here and its
 * entries are removed.
 */
const PROBE = { math: true, charts: true, smartart: true };
const PROBE_CASE_6 = ['body/13'];
const PENDING = {
  'fixtures/charts/waterfall.json': { charts: true },
  'fixtures/check50/50-declarations.json': { req303: ['body/4', 'body/6', 'body/8'] },
  'fixtures/loadAndSave.json': { math: true, charts: true },
  'fixtures/revisions/cr043/equations-accept-all.json': { math: true },
  'fixtures/revisions/cr043/equations-reject-all.json': { math: true },
  // body/12's accent: U+0305 deleted, U+203E inserted, over the same U
  'fixtures/revisions/cr043/tracked-changes-equations.json': { math: true, latexOnly: ['body/12'] },
  'fixtures/revisions/python/artefact8/8-tracked-changes.json': { math: true },
  'fixtures/revisions/python/artefact8/8-word-accepted.json': { math: true },
  'fixtures/revisions/python/artefact8/8-word-rejected.json': { math: true },
  'text-extraction/revisions/python/artefact8/8-tracked-changes.json': { math: true },
  'text-extraction/revisions/python/artefact8/8-word-accepted.json': { math: true },
  'text-extraction/revisions/python/artefact8/8-word-rejected.json': { math: true },
  // the probe in its eight copies: phase 5's cases, and case 6 where the copy keeps the rebound prefix
  'text-extraction/suite/extraction-probe.json': { ...PROBE, req303: PROBE_CASE_6 },
  'text-extraction/suite/extraction-probe.word-resave.json': PROBE,
  'text-extraction/suite/phase0/extraction-probe.json': { ...PROBE, altchunks: true, req303: PROBE_CASE_6 },
  'text-extraction/suite/phase0/evidence/checked-52/extraction-probe.json': { ...PROBE, altchunks: true, req303: PROBE_CASE_6 },
  'text-extraction/suite/phase0/evidence/extraction-probe-r2-word-save.json': PROBE,
  'text-extraction/suite/phase0/evidence/extraction-probe-r3-word-save.json': PROBE,
  'text-extraction/suite/phase0/evidence/extraction-probe-v1-rejected.json': { ...PROBE, altchunks: true, req303: PROBE_CASE_6 },
  'text-extraction/suite/phase0/evidence/extraction-probe-v1-word-repaired.json': PROBE,
};

/** The flags that set blocks aside by their address's last segment, and the flags that set them aside by address. */
const MARKERS = { charts: '/chart:', smartart: '/smartart:', altchunks: '/altchunk:' };
const AT = ['latexOnly', 'req303'];

/** Every block, its cells' blocks and its comments' blocks, in order. */
function allBlocks(body) {
  const out = [];
  const walk = (blocks) => {
    for (const b of blocks) {
      out.push(b);
      for (const r of b.rows ?? []) for (const c of r.cells ?? []) walk(c.blocks ?? []);
    }
  };
  walk(body.blocks);
  for (const c of body.comments ?? []) walk(c.blocks ?? []);
  return out;
}

/** A block with its equations' run text masked, deeply. */
function maskMath(b) {
  const m = { ...b };
  if (b.runs) m.runs = b.runs.map((r) => (r.math ? { ...r, text: '<math>' } : r));
  if (b.rows) m.rows = b.rows.map((row) => ({ ...row, ...(row.cells ? { cells: row.cells.map((c) => ({ ...c, ...(c.blocks ? { blocks: c.blocks.map(maskMath) } : {}) })) } : {}) }));
  return m;
}

/** The plain text with each equation's run text, in order, masked. */
function maskMathText(plainText, body) {
  let out = plainText;
  let from = 0;
  for (const b of allBlocks(body)) {
    for (const r of b.runs ?? []) {
      if (!r.math || !r.text) continue;
      const at = out.indexOf(r.text, from);
      if (at < 0) continue;
      out = out.slice(0, at) + '<math>' + out.slice(at + r.text.length);
      from = at + '<math>'.length;
    }
  }
  return out;
}

/** One side's body with what `pending` sets aside taken out (`kept`), and what each flag took (`aside`). */
function setAside(body, pending) {
  const under = (address, at) => address === at || address.startsWith(`${at}/`);
  let blocks = body.blocks;
  let warnings = body.warnings ?? [];
  const aside = {};
  for (const [flag, marker] of Object.entries(MARKERS)) {
    if (!pending[flag]) continue;
    aside[flag] = { blocks: blocks.filter((b) => b.address.includes(marker)) };
    blocks = blocks.filter((b) => !b.address.includes(marker));
  }
  if (pending.altchunks) {
    aside.altchunks.warnings = warnings.filter((w) => w.kind === 'altchunk.not_expanded');
    warnings = warnings.filter((w) => w.kind !== 'altchunk.not_expanded');
  }
  for (const flag of AT) {
    const addresses = pending[flag];
    if (!addresses) continue;
    const at = (address) => address !== undefined && addresses.some((a) => under(address, a));
    aside[flag] = { blocks: blocks.filter((b) => at(b.address)), warnings: warnings.filter((w) => at(w.address)) };
    blocks = blocks.filter((b) => !at(b.address));
    warnings = warnings.filter((w) => !at(w.address));
  }
  const kept = { ...body, blocks };
  delete kept.warnings;
  if (warnings.length > 0) kept.warnings = warnings;
  return { kept, aside };
}

// ------------------------------------------------------------------ the tests

test('text goldens: one per document under test/fixtures/ and test/text-extraction/, from TextHarness version 1', () => {
  assert.equal(goldens.length, 282, 'the WordprocessingML documents committed under the two directories (CR-012 section 20)');
  for (const name of Object.keys(PENDING)) assert.ok(goldens.includes(name), `PENDING names a golden: ${name}`);
  return Promise.all(goldens.map(async (name) => {
    const golden = JSON.parse(await readFile(join(goldenDir, name), 'utf8'));
    assert.equal(golden.header.harnessVersion, '1', name);
    assert.equal(golden.header.timeZone, 'UTC', name);
    assert.equal(golden.header.fixture.replace(/\.(docx|docm|dotx)$/, '.json'), name, name);
    assert.deepEqual(Object.keys(golden.readings), ['accepted', 'original', 'markup', 'everything'], name);
  }));
});

for (const name of goldens) {
  test(`text parity: ${name}`, async () => {
    const golden = JSON.parse(await readFile(join(goldenDir, name), 'utf8'));
    const pkg = await WordprocessingMLPackage.load(await readFile(join(here, golden.header.fixture)));
    const pending = PENDING[name] ?? {};
    const needed = new Set();
    let metadata;
    for (const [reading, { options }] of Object.entries(golden.readings)) {
      const result = await Texts.of(pkg, TextOptions.of(options));
      metadata ??= result.metadata;
      const oursWhole = joinedBody(body(result));
      const theirsWhole = joinedBody(goldenBody(golden, reading));
      const { kept: ours, aside: oursAside } = setAside(oursWhole, pending);
      const { kept: theirs, aside: theirsAside } = setAside(theirsWhole, pending);
      const at = `${name} ${reading}`;

      // what each flag set aside, and whether it differs in this reading
      const differs = Object.keys(theirsAside).filter((flag) => !isDeepStrictEqual(oursAside[flag], theirsAside[flag]));
      for (const flag of differs) needed.add(flag);
      if (pending.math && !isDeepStrictEqual(ours.blocks, theirs.blocks)) needed.add('math');
      const mask = (blocks) => (pending.math ? blocks.map(maskMath) : blocks);

      if (differs.length === 0) {
        const plain = (side, whole) => (pending.math ? maskMathText(side.plainText, whole) : side.plainText);
        assert.equal(plain(ours, oursWhole), plain(theirs, theirsWhole), `${at}: the plain text`);
      }
      assert.deepStrictEqual(ours.blocks.map((b) => b.address), theirs.blocks.map((b) => b.address), `${at}: the blocks' addresses`);
      const oursBlocks = mask(ours.blocks);
      const theirsBlocks = mask(theirs.blocks);
      oursBlocks.forEach((b, i) => assert.deepStrictEqual(b, theirsBlocks[i], `${at}: block ${b.address}`));
      assert.deepStrictEqual(ours.comments, theirs.comments, `${at}: the comments`);
      assert.deepStrictEqual(ours.warnings, theirs.warnings, `${at}: the warnings`);
    }
    for (const flag of Object.keys(pending)) {
      assert.ok(needed.has(flag), `${name}: PENDING's ${flag} sets aside nothing that differs now; if its phase has landed, remove the entry`);
    }
    assert.deepStrictEqual(JSON.parse(JSON.stringify(metadata)), golden.metadata, `${name}: the metadata`);
  });
}
