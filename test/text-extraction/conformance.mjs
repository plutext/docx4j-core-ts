// CR-012 phase 0 (D8 a): the text extraction conformance suite's protocol, a mirror of docx4j's
// org.docx4j.textextraction.ExtractionConformance as it stands at the commit suite/ was copied from
// (manifest.json). The rows, their requirements and their comparison are docx4j's, as written,
// including the named exceptions in textRequirements and the supplemental scenarios in run: a
// difference in what the suite asks goes back to docx4j, not into this file (CR-012 section 8).
//
// An engine is the text model's adapter (docx4j TextsAdapter):
//   implementedStatements()                      the statements (and capabilities) it implements
//   extract(fixture, view, scenario, format)     { text, findings }: findings maps
//                                                channel<TAB>kind<TAB>selector to a count or value
//   inspect(fixture), counts(fixture, output)    an inspector's (WD section 9), optional
// Each may return a promise. A row whose requirements the engine doesn't all implement is PENDING,
// with its phase and what is missing; with no engine (null) every row is.
import { AssertionError } from 'node:assert';
import { existsSync, readFileSync } from 'node:fs';
import { join, basename } from 'node:path';

/** A UTF-8 TSV with a header: one object per non-empty line, keyed by the header. */
export function table(file) {
  const lines = readFileSync(file, 'utf8').split(/\r\n|\n|\r/);
  if (lines.length > 0 && lines[lines.length - 1] === '') lines.pop();
  const header = lines[0].split('\t');
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    if (lines[i] === '') continue;
    const values = lines[i].split('\t');
    if (values.length !== header.length) throw new Error(`${file}:${i + 1}: column count`);
    const row = {};
    header.forEach((name, j) => { row[name] = values[j]; });
    rows.push(row);
  }
  return rows;
}

/**
 * REQ-1201 only: CriticMarkup's delimiters, then white space (Java's \s, the ASCII six, and \p{Z}),
 * `*`, `_`, `\`, `{`, `}` and `$`. Zero-width spaces, tag characters, hyphens and Unicode
 * composition are kept.
 */
export function normalize(value) {
  return value.replace(/\{\+\+|\+\+\}|\{--|--\}|\{>>|<<\}/g, '')
    .replace(/[\t\n\x0B\f\r\p{Z}*_\\{}$]/gu, '');
}

/**
 * Occurrences of a marker, not counting those inside a longer marker token of the suite: after
 * normalisation removes underscores and white space, FLD19_RESULT is also found inside
 * FLD19_RESULT_TWO, and MCE8_OUTER inside MCE8_OUTER_FALLBACK.
 */
export function count(text, marker, suite = []) {
  const haystack = normalize(text);
  const needle = normalize(marker);
  if (needle === '') throw new Error('Empty marker');
  const covered = new Uint8Array(haystack.length);
  for (const other of suite) {
    const longer = normalize(other);
    if (longer.length <= needle.length || !longer.includes(needle)) continue;
    for (let i = 0; (i = haystack.indexOf(longer, i)) >= 0; i += longer.length) covered.fill(1, i, i + longer.length);
  }
  let result = 0;
  for (let i = 0; (i = haystack.indexOf(needle, i)) >= 0; i += needle.length) if (!covered[i]) result++;
  return result;
}

const show = (v) => v instanceof Map || (v !== null && typeof v === 'object') ? JSON.stringify(v instanceof Map ? Object.fromEntries(v) : v) : String(v);

/** docx4j's `equal`: Objects.equals, an AssertionError naming the context otherwise. */
export function equal(expected, actual, context) {
  const same = expected !== null && typeof expected === 'object'
    ? JSON.stringify(sorted(expected)) === JSON.stringify(sorted(actual))
    : expected === actual;
  if (!same) throw new AssertionError({ message: `${context}: expected ${show(expected)}, got ${show(actual)}` });
}
const sorted = (o) => o === null || o === undefined ? o : Object.fromEntries(Object.entries(o instanceof Map ? Object.fromEntries(o) : o).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)));

const MARKING = { ins: ['{++', '++}'], del: ['{--', '--}'], cmt: ['{>>', '<<}'] };
const quote = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** A token row's expectation: a count, or CriticMarkup kinds joined by `+` (`del+ins`). */
export function verify(text, token, expected, context, suite = []) {
  if (/^\d+$/.test(expected)) {
    equal(Number.parseInt(expected, 10), count(text, token, suite), context);
    return;
  }
  const kinds = expected.split('+');
  equal(kinds.length, count(text, token, suite), `${context} total`);
  for (const kind of kinds) {
    if (!(kind in MARKING)) throw new Error(`Unknown marking: ${kind}`);
    const [opening, closing] = MARKING[kind];
    const matches = new RegExp(`${quote(opening)}([\\s\\S]*?)${quote(closing)}`, 'g');
    let marked = 0;
    for (const m of text.matchAll(matches)) marked += count(m[1], token, suite);
    equal(1, marked, `${context} marking ${kind}`);
  }
}

/** Java's String.split(",\\s*"), trailing empty strings dropped, into an ordered set. */
export function requirements(value) {
  const parts = value.split(/,\s*/);
  while (parts.length > 1 && parts[parts.length - 1] === '') parts.pop();
  return new Set(parts);
}

/** A findings, order or golden row's requirements: its statements and its scenario's. */
export function rowRequirements(row) {
  const req = requirements(row.statements);
  for (const r of scenarioRequirements(row.scenario)) req.add(r);
  return req;
}

/** A token row's statements, with docx4j's named exceptions. */
export function textRequirements(row, format) {
  const req = requirements(row.statements);
  req.delete('REQ-902'); // an extractor does not claim inspector conformance
  if (row.token === 'CHT34_AXIS_TITLE') req.add('CR-047');
  if (row.case === '21' && format === 'txt') req.delete('REQ-703');
  // docx4j CR-046 section 4: case 25's text boxes are phase 1; only its legacy WordArt needs REQ-511
  if (row.case === '25' && !row.token.startsWith('WA25')) req.delete('REQ-511');
  return req;
}

/** Capabilities, not options: whether an HTML reader is present (REQ-513). */
export const HTML_READER_PRESENT = 'html reader present';
export const HTML_READER_ABSENT = 'html reader absent';

const OPTION_STATEMENTS = {
  text_boxes: 'REQ-518', notes: 'REQ-503', note_references: 'REQ-503', comments: 'REQ-504',
  hidden: 'REQ-507', field_codes: 'REQ-505', fields: 'REQ-506', placeholders: 'REQ-508',
  toc: 'REQ-520', bindings: 'REQ-509', list_labels: 'REQ-510', charts: 'REQ-511', smartart: 'REQ-511',
  embedded: 'REQ-512', altchunks: 'REQ-513', alt_text: 'REQ-514', metadata: 'REQ-515',
  symbols: 'REQ-601', math: 'REQ-604', tables: 'REQ-605', markup_detail: 'REQ-702',
  page_markers: 'REQ-708', layout: 'REQ-709',
};

/** The statement a scenario's option belongs to (WD section 10), which the engine must implement too. */
export function scenarioRequirements(scenario) {
  if (scenario === undefined || scenario === null || scenario === 'default') return new Set();
  if (scenario === HTML_READER_PRESENT) return new Set(['REQ-513', HTML_READER_PRESENT]);
  if (scenario === HTML_READER_ABSENT) return new Set(['REQ-513', HTML_READER_ABSENT]);
  if (scenario === 'markdown') return new Set(['REQ-703']);
  if (scenario.includes(';')) {
    // several options, "; " apart: each one's statement
    const all = new Set();
    for (const option of scenario.split(';')) for (const r of scenarioRequirements(option.trim())) all.add(r);
    return all;
  }
  const colon = scenario.indexOf(':');
  const name = (colon < 0 ? scenario : scenario.slice(0, colon)).trim();
  const value = colon < 0 ? '' : scenario.slice(colon + 1).trim();
  if (name === 'include') {
    return new Set([value.includes('glossary') ? 'REQ-516' : value.includes('comments') ? 'REQ-504' : 'REQ-502']);
  }
  const statement = OPTION_STATEMENTS[name];
  if (statement === undefined) throw new Error(`Unknown scenario: ${scenario}`);
  return new Set([statement]);
}

/** The phase a row waits for, as docx4j's CR-046 numbers them; 9 is the inspector (CR-050). */
export function phase(reqs) {
  let result = 1;
  for (const req of reqs) {
    if (!req.startsWith('REQ-')) { result = Math.max(result, 2); continue; }
    const n = Number.parseInt(req.slice(4), 10);
    if (n >= 901 && n <= 907) return 9;
    if ([506, 509, 511, 512, 513, 514, 515, 516, 604, 605, 708, 709].includes(n)) result = Math.max(result, 2);
    if (n === 703 || n === 704 || n === 705) result = Math.max(result, 3);
  }
  return result;
}

// ------------------------------------------------------------------ the inspection (CR-050)

export const INSPECTION_TABLE = 'expected-inspection.tsv';
export const INSPECTION_STATUS = 'expected-inspection-status.tsv';
/** A finding's identity: its kind and its stored location. */
export const INSPECTION_KEY = ['kind', 'chain', 'part', 'path', 'start', 'end_path', 'end'];
/** The fields compared for a finding whose identity matched. */
export const INSPECTION_FIELDS = ['address', 'visible', 'reveal', 'plain', 'markdown', 'blocks', 'hidden',
  'embedded', 'comments', 'excerpt_length', 'excerpt_cut'];
/** The statements an inspection row needs (WD section 9), REQ-903's counts apart. */
export const INSPECTION_STATEMENTS = 'REQ-901, REQ-902, REQ-904, REQ-905, REQ-906, REQ-907';
/** The six outputs the table gives each finding's group under. */
export const INSPECTION_OUTPUTS = ['plain', 'markdown', 'blocks', 'hidden', 'embedded', 'comments'];
/** What every inspector says it does not detect (REQ-904), by id. */
export const INSPECTION_LIMITATIONS = ['layout.covered', 'layout.position', 'layout.clipped',
  'background.image', 'font.careful_swap', 'payload.opaque', 'chart.workbook', 'intent'];

const byCodeUnit = (a, b) => (a < b ? -1 : a > b ? 1 : 0);

/** location key -> the sorted field tuples of the findings there (a multiset) */
function occurrences(findings) {
  const out = new Map();
  for (const f of findings) {
    const key = INSPECTION_KEY.map((k) => f[k] ?? '').join(' | ');
    const fields = INSPECTION_FIELDS.map((k) => f[k] ?? '').join(' | ');
    if (!out.has(key)) out.set(key, []);
    out.get(key).push(fields);
  }
  for (const l of out.values()) l.sort(byCodeUnit);
  return out;
}

/** A fixture's inspection against the table: by occurrence, then field by field, then the status. */
export function compareInspection(fixture, expected, status, actual) {
  const want = occurrences(expected);
  const got = occurrences(actual.findings);
  const problems = [];
  for (const key of [...new Set([...want.keys(), ...got.keys()])].sort(byCodeUnit)) {
    const w = want.get(key) ?? [];
    const g = got.get(key) ?? [];
    if (w.length !== g.length) {
      problems.push(`${w.length > g.length ? 'missing ' : 'unexpected '}${Math.abs(w.length - g.length)}: ${key}`);
    } else if (w.join('\n') !== g.join('\n')) {
      problems.push(`fields of ${key}: expected [${w.join(', ')}], got [${g.join(', ')}] (${INSPECTION_FIELDS.join(', ')})`);
    }
  }
  if (problems.length > 0) {
    throw new AssertionError({ message: `${fixture}: inspection differs (${problems.length}):\n  ${problems.slice(0, 40).join('\n  ')}` });
  }
  equal(status.status, actual.status, `${fixture}: status`);
  equal(status.skipped, [...new Set(actual.skipped)].sort(byCodeUnit).join('; '), `${fixture}: skipped`);
  equal(status.undecided, [...new Set(actual.undecided)].sort(byCodeUnit).join('; '), `${fixture}: undecided`);
  for (const id of INSPECTION_LIMITATIONS) {
    if (!actual.limitations.some((l) => l.startsWith(`${id}:`) || l === id)) {
      throw new AssertionError({ message: `${fixture}: limitation not stated: ${id}` });
    }
  }
}

// ------------------------------------------------------------------ order and goldens

/** Each marker once, and in this order (` > ` apart), after normalisation. */
export function verifyOrder(text, markers, context) {
  const normalized = normalize(text);
  let after = 0;
  for (const token of markers.split(/\s*>\s*/)) {
    const marker = normalize(token);
    equal(1, count(text, token), `${context}: order marker ${token}`);
    const position = normalized.indexOf(marker);
    if (position < after) throw new AssertionError({ message: `${context}: wrong order at ${token}` });
    after = position + marker.length;
  }
}

const IMAGE = /!\[((?:\\.|[^\]\\])*)\]\([^)]*\)/g;

/** Markdown with each image's destination written `…`: REQ-703 leaves it to the implementation. */
export function imageDestinations(markdown) {
  return markdown.replace(IMAGE, (_, alt) => `![${alt}](…)`);
}

/** A whole output against a golden, after REQ-1201's normalisation (and, for Markdown, image destinations). */
export function verifyGolden(file, actual) {
  let golden = readFileSync(file, 'utf8');
  if (basename(file).endsWith('.md')) {
    golden = imageDestinations(golden);
    actual = imageDestinations(actual);
  }
  equal(normalize(golden), normalize(actual), file);
}

// ------------------------------------------------------------------ the run

class Run {
  constructor(engine, implemented) {
    this.engine = engine;
    this.implemented = implemented;
    this.tokens = new Set();
    this.outputs = new Map();
    this.report = ['fixture\tcase\titem\tview\tscenario\tstatus\treason'];
    this.passed = 0;
    this.pending = 0;
  }

  ready(row, item, view, scenario, reqs) {
    const missing = [...reqs].filter((r) => !this.implemented.has(r));
    const ready = this.engine !== null && missing.length === 0;
    if (!ready) {
      this.pending++;
      const p = phase(reqs);
      this.report.push(`${row.fixture}\t${row.case}\t${item}\t${view}\t${scenario}\tPENDING\t`
        + `${p === 9 ? 'the inspector (CR-050)' : `phase ${p}`}; ${missing.join(', ')}`
        + `${this.engine === null ? '; Texts adapter not installed' : ''}`);
    }
    return ready;
  }

  async output(dir, row, view, scenario, format) {
    const key = `${row.fixture}\t${view}\t${scenario}\t${format}`;
    if (!this.outputs.has(key)) this.outputs.set(key, await this.engine.extract(join(dir, row.fixture), view, scenario, format));
    return this.outputs.get(key);
  }

  pass(row, item, view, scenario) {
    this.passed++;
    this.report.push(`${row.fixture}\t${row.case}\t${item}\t${view}\t${scenario}\tPASS\t`);
  }
}

const VIEWS = ['accepted', 'original', 'markup'];
const findingOf = (findings, key) => (findings instanceof Map ? findings.get(key) : findings?.[key]) ?? null;

async function checkToken(run, dir, row, view, scenario, format, expected) {
  const reqs = textRequirements(row, format);
  for (const r of scenarioRequirements(scenario)) reqs.add(r);
  if (!run.ready(row, row.token, view, scenario, reqs)) return;
  const output = await run.output(dir, row, view, scenario, format);
  verify(output.text, row.token, expected, `${row.case}:${row.token}:${view}:${scenario}`, run.tokens);
  if (row.token === 'TAG48_MARK') {
    equal(true, output.text.includes('TAG48_MARK\u{E0041}\u{E0042}'), 'REQ-602 tag characters retained');
  }
  run.pass(row, row.token, view, scenario);
}

async function inspections(run, dir) {
  if (!existsSync(join(dir, INSPECTION_TABLE))) return;
  const expected = new Map();
  for (const row of table(join(dir, INSPECTION_TABLE))) {
    if (!expected.has(row.fixture)) expected.set(row.fixture, []);
    expected.get(row.fixture).push(row);
  }
  const statuses = table(join(dir, INSPECTION_STATUS));
  for (const status of statuses) {
    // the whole fixture's comparison is REQ-1204's, its checks REQ-901, 902 and 904 to 907's
    const statements = `${INSPECTION_STATEMENTS}, REQ-1204`;
    const row = { fixture: status.fixture, case: '0', statements };
    if (!run.ready(row, 'inspection', '-', 'default', requirements(statements))) continue;
    compareInspection(status.fixture, expected.get(status.fixture) ?? [], status, await run.engine.inspect(join(dir, status.fixture)));
    run.pass(row, 'inspection', '-', 'default');
  }
  // REQ-903: an extraction's counts under each output, against the table's groups, with no HTML reader
  const countsRequire = requirements('REQ-903');
  countsRequire.add(HTML_READER_ABSENT);
  for (const status of statuses) {
    const row = { fixture: status.fixture, case: '0', statements: 'REQ-903' };
    for (const output of INSPECTION_OUTPUTS) {
      if (!run.ready(row, 'counts', '-', output, countsRequire)) continue;
      const want = {};
      for (const group of ['unseen', 'omitted', 'active', 'info']) {
        want[group] = String((expected.get(status.fixture) ?? []).filter((f) => f[output] === group).length);
      }
      want.status = status.status;
      equal(want, await run.engine.counts(join(dir, status.fixture), output), `${status.fixture}: counts under ${output}`);
      run.pass(row, 'counts', '-', output);
    }
  }
}

/**
 * Runs the suite in `dir` (the fixture directory, `phase0/`) against an engine, or null for none.
 * Throws an AssertionError at the first enabled assertion that fails; returns the report otherwise.
 */
export async function run(dir, engine) {
  const implemented = new Set(engine === null ? [] : await engine.implementedStatements());
  const r = new Run(engine, implemented);
  const tokens = table(join(dir, 'expected-tokens.tsv'));
  for (const row of tokens) r.tokens.add(row.token);
  for (const row of tokens) {
    for (const view of VIEWS) {
      // a row gated on an HTML reader (an HTML or MHT chunk) runs its default scenario without one
      const scenario = row.option === HTML_READER_PRESENT ? HTML_READER_ABSENT : 'default';
      await checkToken(r, dir, row, view, scenario, 'txt', row[view]);
      const option = row.option;
      if (option !== '' && (option !== 'comments: true' || view === 'markup')) {
        await checkToken(r, dir, row, view, option, option === 'markdown' ? 'md' : 'txt', row.with_option);
      }
    }
  }
  // supplemental scenarios pin combined capabilities, no-fetch fields, and character retention
  for (const scenario of ['embedded: true', 'altchunks: false', 'fields: update', HTML_READER_PRESENT, 'metadata: true', 'field_codes: true']) {
    for (const row of tokens) {
      const n = Number.parseInt(row.case, 10);
      const applies = scenario === 'embedded: true' && ((n >= 39 && n <= 42) || n === 53 || n === 57)
        || scenario === 'altchunks: false' && ((n >= 43 && n <= 47) || n === 36 || row.token === 'CHK63_SCRIPTED')
        || scenario === 'fields: update' && n >= 49 && n <= 51
        || scenario === HTML_READER_PRESENT && n === 47
        || (scenario === 'metadata: true' || scenario === 'field_codes: true' || scenario === 'embedded: true') && n >= 50 && n <= 51;
      if (!applies || row.option === scenario) continue;
      for (const view of VIEWS) await checkToken(r, dir, row, view, scenario, 'txt', scenario === 'altchunks: false' ? '0' : row[view]);
    }
  }
  for (const row of table(join(dir, 'expected-findings.tsv'))) {
    const item = `${row.channel}:${row.kind}:${row.selector}`;
    // an inspection is the package's, whatever the view: its rows are checked once
    for (const view of row.channel === 'inspect' ? ['accepted'] : VIEWS) {
      if (!r.ready(row, item, view, row.scenario, rowRequirements(row))) continue;
      const output = await r.output(dir, row, view, row.scenario, 'txt');
      equal(row.expected, findingOf(output.findings, `${row.channel}\t${row.kind}\t${row.selector}`), item);
      r.pass(row, item, view, row.scenario);
    }
  }
  if (existsSync(join(dir, 'expected-order.tsv'))) {
    for (const row of table(join(dir, 'expected-order.tsv'))) {
      for (const view of VIEWS) {
        if (!r.ready(row, 'order', view, row.scenario, rowRequirements(row))) continue;
        const output = await r.output(dir, row, view, row.scenario, 'txt');
        verifyOrder(output.text, row.tokens, `${row.case}:${view}`);
        r.pass(row, 'order', view, row.scenario);
      }
    }
  }
  await inspections(r, dir);
  // every published golden is checked; one is cut without an HTML reader
  for (const row of table(join(dir, 'goldens.tsv'))) {
    const item = `golden:${row.golden}`;
    const required = rowRequirements(row);
    if (row.scenario === 'default') required.add(HTML_READER_ABSENT);
    if (!r.ready(row, item, row.view, row.scenario, required)) continue;
    const output = await r.output(dir, row, row.view, row.scenario, row.format);
    verifyGolden(join(dir, row.golden), output.text);
    r.pass(row, item, row.view, row.scenario);
  }
  return { passed: r.passed, pending: r.pending, report: r.report };
}
