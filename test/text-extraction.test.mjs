// CR-012 phases 0 and 2: the text extraction specification's conformance suite (D8 a), its agreement
// test (D8 b) and the understood set (REQ-306), and this package's text model held to them.
//
// test/text-extraction/suite/ is docx4j's docx4j-core-tests text-extraction/ directory, copied
// unedited by scripts/copy-text-extraction-suite.mjs at the commit manifest.json names; the runner
// (text-extraction/conformance.mjs) mirrors docx4j's ExtractionConformance, and the tests below
// mirror its ExtractionConformanceTest's checks of the protocol, which use built tables and engines.
// Phase 2 installs the engine (text-extraction/adapter.mjs, docx4j's TextsAdapter): the rows whose
// statements it claims run, and the others stay pending with their phase.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdtempSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  WordprocessingMLPackage, UNDERSTOOD_NAMESPACES, READ_ONLY_UNDERSTOOD_NAMESPACES,
} from '../dist/index.mjs';
import {
  table, normalize, count, verify, verifyGolden, run, INSPECTION_TABLE, INSPECTION_STATUS,
  INSPECTION_HTML_TABLE, INSPECTION_HTML_STATUS, INSPECTION_LIMITATIONS, HTML_READER_ABSENT, HTML_READER_PRESENT,
} from './text-extraction/conformance.mjs';
import { agreement, agreementFixtures } from './text-extraction/agreement.mjs';
import { textsAdapter, agreementEngine, IMPLEMENTED } from './text-extraction/adapter.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const root = dirname(here);
const base = join(here, 'text-extraction');
const phase0 = join(base, 'suite', 'phase0');
const manifest = JSON.parse(readFileSync(join(base, 'manifest.json'), 'utf8'));

/** The Texts adapter (docx4j TextsAdapter), CR-012 phase 2. */
const ENGINE = textsAdapter();

const scratch = (name) => mkdtempSync(join(tmpdir(), `${name}-`));
const sha1 = (file) => createHash('sha1').update(readFileSync(file)).digest('hex');
function filesUnder(dir) {
  const out = [];
  const walk = (d) => {
    for (const e of readdirSync(d, { withFileTypes: true })) {
      if (e.isDirectory()) walk(join(d, e.name));
      else out.push(relative(root, join(d, e.name)).split(sep).join('/'));
    }
  };
  walk(dir);
  return out;
}

// ---------------------------------------------------------------- the copy

test('the suite is docx4j\'s, unedited: every file as manifest.json records it, and nothing more', () => {
  assert.match(manifest.docx4j.commit, /^[0-9a-f]{40}$/);
  for (const f of manifest.files) assert.equal(sha1(join(root, f.path)), f.sha1, `${f.path} (docx4j ${f.docx4j})`);
  const listed = new Set(manifest.files.map((f) => f.path));
  for (const dir of ['suite', 'revisions']) {
    for (const file of filesUnder(join(base, dir))) assert.ok(listed.has(file), `${file} is not docx4j's: recopy, never edit`);
  }
});

test('the suite\'s case statements and README rows match WD Appendix A and the approved additions', () => {
  // docx4j ExtractionConformanceTest.requirementsAndReadmeMatchAppendixA
  const support = join(base, 'suite', 'support');
  const appendix = table(join(support, 'appendix-a-statements.tsv'));
  const generated = table(join(phase0, 'case-statements.tsv'));
  assert.equal(appendix.length, 52);
  const additions = table(join(support, 'additional-case-statements.tsv'));
  // 53-54 (WD 0.5), 55-57 (extraction-probe-structure.docx), 58-67 (extraction-probe-inspect.docx, CR-050)
  assert.equal(additions.length, 15);
  appendix.push(...additions);
  assert.equal(generated.length, 67);
  const readme = readFileSync(join(phase0, 'README.md'), 'utf8');
  appendix.forEach((row, i) => {
    assert.equal(generated[i].case, row.case);
    assert.equal(generated[i].statements, row.statements);
    const section = readme.split(`### Case ${i + 1}:`)[1].split('### Case ')[0];
    for (const line of section.split('\n')) {
      if (line.startsWith('| `')) assert.ok(line.endsWith(`| ${row.statements} |`), line);
    }
  });
});

// ---------------------------------------------------------------- the protocol

test('REQ-1201: normalisation keeps the characters that matter', () => {
  assert.equal(normalize('{++ $MTH32\\_A$ ++}'), 'MTH32A');
  assert.equal(normalize('ZW48_SP​LIT'), 'ZW48SP​LIT');
  assert.equal(normalize('TAG48_MARK\u{E0041}\u{E0042}'), 'TAG48MARK\u{E0041}\u{E0042}');
  assert.notEqual(normalize('a­b'), normalize('ab'));
  assert.notEqual(normalize('a‑b'), normalize('ab'));
  assert.equal(count('MTH32\\_A', 'MTH32_A'), 1);
  // Java's \s is the ASCII six and \p{Z} the separators: U+FEFF (JavaScript's \s) is kept
  assert.equal(normalize('A B C\u000bD﻿E'), 'ABCD﻿E');
  // a marker inside a longer marker of the suite is not counted
  assert.equal(count('FLD19_RESULT_TWO FLD19_RESULT', 'FLD19_RESULT', ['FLD19_RESULT_TWO']), 1);
});

test('a marking must be right before normalisation', () => {
  verify('{-- TC10_MOVED --}{++ TC10_MOVED ++}', 'TC10_MOVED', 'del+ins', 'move');
  assert.throws(() => verify('{-- TC9_NEW --}', 'TC9_NEW', 'ins', 'wrong marking'), /marking ins/);
});

test('the golden reader normalises as REQ-1201 says and finds a change', () => {
  const dir = scratch('golden');
  const golden = join(dir, 'example.accepted.txt');
  writeFileSync(golden, 'A_B\n');
  verifyGolden(golden, 'A\\_B');
  assert.throws(() => verifyGolden(golden, 'A_C'), /expected AB/);
});

test('a Markdown golden ignores image destinations and compares alt text (REQ-703)', () => {
  const dir = scratch('golden-md');
  const golden = join(dir, 'example.md');
  writeFileSync(golden, 'See ![IMG\\_ALT \\[1\\]](…) here.\n');
  verifyGolden(golden, 'See ![IMG\\_ALT \\[1\\]](data:image/png;base64,iVBORw0KGgo=) here.\n');
  verifyGolden(golden, 'See ![IMG\\_ALT \\[1\\]](images/image1.png) here.\n');
  assert.throws(() => verifyGolden(golden, 'See ![IMG\\_OTHER](images/image1.png) here.\n'), /IMGALT/);
  const text = join(dir, 'example.accepted.txt');
  writeFileSync(text, '![a](…)\n');
  assert.throws(() => verifyGolden(text, '![a](x.png)'));
});

const EMPTY_TOKENS = 'case\ttoken\taccepted\toriginal\tmarkup\toption\twith_option\tfixture\tstatements\n';
const EMPTY_FINDINGS = 'case\tfixture\tchannel\tkind\tselector\tscenario\texpected\tstatements\n';
const EMPTY_GOLDENS = 'fixture\tcase\tview\tscenario\tformat\tgolden\tstatements\n';

test('an engine\'s rows and a published golden are checked, and a golden that differs fails', async () => {
  const dir = scratch('suite');
  writeFileSync(join(dir, 'expected-tokens.tsv'), `${EMPTY_TOKENS}9\tTC9_NEW\t1\t0\tins\t\t\tfixture.docx\tREQ-402, REQ-403\n`);
  writeFileSync(join(dir, 'expected-findings.tsv'), EMPTY_FINDINGS);
  writeFileSync(join(dir, 'goldens.tsv'), `${EMPTY_GOLDENS}fixture.docx\t0\taccepted\tdefault\ttxt\tfixture.accepted.txt\tREQ-402\n`);
  writeFileSync(join(dir, 'fixture.accepted.txt'), 'TC9_NEW\n');
  const engine = {
    // an engine says which of the two HTML-reader capabilities it has; the goldens need "absent"
    implementedStatements: () => ['REQ-402', 'REQ-403', HTML_READER_ABSENT],
    extract: async (p, view) => ({ text: view === 'original' ? '' : view === 'markup' ? '{++ TC9_NEW ++}' : 'TC9_NEW', findings: {} }),
  };
  const result = await run(dir, engine);
  assert.equal(result.report.length, 5); // header, three views and the golden
  assert.ok(!result.report.join('\n').includes('PENDING'));
  writeFileSync(join(dir, 'fixture.accepted.txt'), 'BROKEN\n');
  await assert.rejects(run(dir, engine), /fixture\.accepted\.txt/);
});

test('the presentation and glossary order assertions run in every view, and a wrong order fails', async () => {
  const dir = scratch('ordered-suite');
  writeFileSync(join(dir, 'expected-tokens.tsv'), EMPTY_TOKENS);
  writeFileSync(join(dir, 'expected-findings.tsv'), EMPTY_FINDINGS);
  writeFileSync(join(dir, 'goldens.tsv'), EMPTY_GOLDENS);
  writeFileSync(join(dir, 'expected-order.tsv'), 'case\tfixture\tscenario\ttokens\tstatements\n'
    + '53\tfixture.docx\tembedded: true\tFIRST53_SHAPE > LAST53_SHAPE\tREQ-512\n');
  let reversed = false;
  const engine = {
    implementedStatements: () => ['REQ-512'],
    extract: async (p, view, scenario) => {
      assert.equal(scenario, 'embedded: true');
      return { text: reversed ? 'LAST53_SHAPE FIRST53_SHAPE' : 'FIRST53_SHAPE LAST53_SHAPE', findings: {} };
    },
  };
  const result = await run(dir, engine);
  assert.equal(result.report.length, 4);
  assert.ok(!result.report.join('\n').includes('PENDING'));
  reversed = true;
  await assert.rejects(run(dir, engine), /wrong order/);
});

test('an inspection is compared by occurrence: a finding missing, doubled or of another group fails', async () => {
  const dir = scratch('inspected');
  writeFileSync(join(dir, 'expected-tokens.tsv'), EMPTY_TOKENS);
  writeFileSync(join(dir, 'expected-findings.tsv'), EMPTY_FINDINGS);
  writeFileSync(join(dir, 'goldens.tsv'), EMPTY_GOLDENS);
  const header = 'fixture\tcase\tkind\tchain\tpart\tpath\tstart\tend_path\tend\taddress\tvisible\treveal\tplain\tmarkdown'
    + '\tblocks\thidden\tembedded\tcomments\texcerpt_length\texcerpt_cut\tselector\tsummary\n';
  const one = 'f.docx\t1\ttext.hidden\t\t/word/document.xml\t/w:document[1]/w:body[1]/w:p[1]/w:r[1]\t0\t\t4\tbody/0\tno'
    + '\toptions: hidden: true\tomitted\tomitted\tomitted\tunseen\tomitted\tomitted\t4\tfalse\tA_B\t\n';
  const two = one.replace('w:p[1]', 'w:p[2]').replace('body/0', 'body/1');
  writeFileSync(join(dir, INSPECTION_TABLE), header + one + two);
  writeFileSync(join(dir, INSPECTION_STATUS), 'fixture\tstatus\tskipped\tundecided\nf.docx\tCOMPLETE\t\t\n');
  const expected = table(join(dir, INSPECTION_TABLE));
  let reported = expected;
  const engine = {
    implementedStatements: () => ['REQ-901', 'REQ-902', 'REQ-904', 'REQ-905', 'REQ-906', 'REQ-907', 'REQ-1204', HTML_READER_ABSENT],
    extract: () => { throw new Error('not an extraction'); },
    inspect: async () => ({ findings: reported, status: 'COMPLETE', skipped: [], undecided: [], limitations: INSPECTION_LIMITATIONS }),
  };
  const result = await run(dir, engine);
  assert.ok(result.report.includes('f.docx\t0\tinspection\t-\tdefault\tPASS\t'));
  const moved = { ...expected[1], plain: 'unseen' };
  for (const wrong of [[expected[0]], [expected[0], expected[1], expected[1]], [expected[0], moved]]) {
    reported = wrong;
    await assert.rejects(run(dir, engine), /inspection differs/, `a wrong inspection passed: ${wrong.length}`);
  }
});

test('an HTML reader\'s inspection: parsed is compared, and a row pending a ruling is left out and reported', async () => {
  // docx4j CR-053 phase 3's htmlInspections, which docx4j's own tests reach only through its engine
  const dir = scratch('inspected-html');
  writeFileSync(join(dir, 'expected-tokens.tsv'), EMPTY_TOKENS);
  writeFileSync(join(dir, 'expected-findings.tsv'), EMPTY_FINDINGS);
  writeFileSync(join(dir, 'goldens.tsv'), EMPTY_GOLDENS);
  const header = 'fixture\tcase\tkind\tchain\tpart\tpath\tstart\tend_path\tend\taddress\tvisible\treveal\tplain\tmarkdown'
    + '\tblocks\thidden\tembedded\tcomments\texcerpt_length\texcerpt_cut\tselector\tsummary\tparsed\tstatements\tevidence\tpending\n';
  const checked = 'h.docx\t0\ttext.hidden\t\t/word/chunk.htm\t/html[1]/body[1]/p[1]\t\t\t\t\tno\toptions: hidden: true; html reader present'
    + '\tomitted\tomitted\tomitted\tunseen\tomitted\tomitted\t6\tfalse\tH_SEEN\t\thtml\tREQ-902\tr10-1\t\n';
  const waiting = checked.replace('p[1]', 'p[2]').replace('H_SEEN', 'H_WAITS').replace('\tr10-1\t\n', '\tr11-2\tWD 0.29: a rule not yet ruled on\n');
  writeFileSync(join(dir, INSPECTION_HTML_TABLE), header + checked + waiting);
  // the HTML pass runs inside the inspection pass, which needs its own tables, here empty
  writeFileSync(join(dir, INSPECTION_TABLE), header.split('\tparsed')[0] + '\n');
  writeFileSync(join(dir, INSPECTION_STATUS), 'fixture\tstatus\tskipped\tundecided\n');
  writeFileSync(join(dir, INSPECTION_HTML_STATUS), 'fixture\tstatus\tskipped\tundecided\nh.docx\tCOMPLETE\t\t\n');
  const [seen, waits] = table(join(dir, INSPECTION_HTML_TABLE));
  let reported = [seen, { ...waits, plain: 'unseen' }];
  const engine = {
    implementedStatements: () => ['REQ-901', 'REQ-902', 'REQ-904', 'REQ-905', 'REQ-906', 'REQ-907', 'REQ-1204', 'REQ-513', HTML_READER_PRESENT],
    extract: () => { throw new Error('not an extraction'); },
    inspect: async () => ({ findings: reported, status: 'COMPLETE', skipped: [], undecided: [], limitations: INSPECTION_LIMITATIONS }),
  };
  // the pending row is left out on both sides, whatever the inspector reports there
  const result = await run(dir, engine);
  assert.ok(result.report.includes(`h.docx\t0\tinspection\t-\t${HTML_READER_PRESENT}\tPASS\t`));
  assert.ok(result.report.includes(`h.docx\t0\tinspect:text.hidden:H_WAITS\t-\t${HTML_READER_PRESENT}\tPENDING\tWD 0.29: a rule not yet ruled on`));
  reported = [{ ...seen, parsed: 'xml' }];
  await assert.rejects(run(dir, engine), /inspection differs[\s\S]*parsed/);
});

// ---------------------------------------------------------------- the suite

test('the suite with no engine: every row pending, with its phase', async () => {
  const result = await run(phase0, null);
  const text = result.report.join('\n');
  assert.ok(text.includes('VBA51_MODULE_MARK'));
  assert.ok(text.includes('ALT47_NESTED_CHUNK'));
  assert.ok(text.includes('the inspector (CR-050)'));
  // docx4j 198d015a3's suite: 1,584 rows, the same rows in the same order as docx4j's own run (CR-012
  // section 13.3)
  assert.equal(result.passed, 0);
  assert.equal(result.pending, 1584, 'the suite changed: recount, and compare with docx4j\'s run');
  for (const line of result.report.slice(1)) assert.match(line, /\tPENDING\t(phase [123]|the inspector \(CR-050\)); .*; Texts adapter not installed$/);
});

test('the suite with the text model (CR-012 phase 2): every row it claims passes, the others pend for a statement it does not claim', async () => {
  const result = await run(phase0, ENGINE);
  // run() throws at the first row that fails; these are the counts at docx4j 198d015a3's suite
  assert.equal(result.passed + result.pending, 1584);
  assert.equal(result.passed, 735, 'a row moved: recount (CR-012 section 15)');
  assert.equal(result.pending, 849);
  const claimed = new Set(IMPLEMENTED);
  for (const line of result.report.slice(1)) {
    const [, , , , , status, reason] = line.split('\t');
    if (status !== 'PENDING') continue;
    const missing = reason.split('; ')[1].split(', ');
    assert.ok(missing.length > 0 && missing.every((r) => !claimed.has(r)), line);
  }
  // what pends, by what it waits for: phase 5's statements, Markdown, the inspector, and REQ-303 and REQ-1203
  const waits = (statement) => result.report.filter((l) => l.includes('\tPENDING\t') && l.split('; ')[1].split(', ').includes(statement)).length;
  assert.equal(waits('REQ-303'), 9, 'case 6: the URI decides; this package matches by prefix (REQ-303\'s permitted route)');
  assert.equal(waits('REQ-1203'), 12, 'the whole-document goldens hold phase 5\'s content');
  for (const statement of ['REQ-506', 'REQ-511', 'REQ-512', 'REQ-513', 'REQ-604', 'REQ-703', 'REQ-901']) assert.ok(waits(statement) > 0, statement);
});

test('the whole-document goldens (REQ-1203, not claimed until phase 5): exact but for phase 5\'s content and case 6', async () => {
  // docx4j's ExtractionGoldens writes the text with its trailing line ends trimmed to one
  const asWritten = (text) => `${text.replace(/\n+$/, '')}\n`;
  // the lines phase 5 adds (5b LaTeX, 5c the chart and SmartArt, 5e the chunks) and case 6's (REQ-303 by prefix)
  const PENDING = /MCE6_|MTH32|CHT34_|SMA35_|ALT4[3-7]_|^\[chart: /;
  const without = (text) => text.split('\n').filter((line) => !PENDING.test(line)).join('\n');
  let exact = 0;
  for (const row of table(join(phase0, 'goldens.tsv'))) {
    if (row.format !== 'txt') continue;
    const ours = asWritten((await ENGINE.extract(join(phase0, row.fixture), row.view, row.scenario, 'txt')).text);
    const golden = readFileSync(join(phase0, row.golden), 'utf8');
    if (ours === golden) { exact++; continue; }
    assert.equal(row.fixture, 'extraction-probe.docx', `${row.golden} differs`);
    assert.equal(without(ours), without(golden), row.golden);
  }
  assert.equal(exact, 6, 'the external and structure probes\' goldens, in every view');
});

// ---------------------------------------------------------------- the agreement test

const fixtures = agreementFixtures();

test('the agreement test\'s fixtures: docx4j\'s 132 revisions fixtures and 15 text-extraction documents', () => {
  assert.equal(fixtures.filter((f) => f.name.startsWith('revisions/')).length, 132);
  assert.equal(fixtures.filter((f) => f.name.startsWith('text-extraction/')).length, 15);
});

test('the agreement test (REQ-406, REQ-1202): the views\' walk is the text of the copy acceptAllRevisions or rejectAllRevisions leaves, on every fixture', async () => {
  const result = await agreement(fixtures, agreementEngine(), (bytes) => WordprocessingMLPackage.load(bytes));
  assert.equal(result.pending, 0);
  assert.equal(result.compared, fixtures.length * 4, 'two views, default scope and all stories');
  assert.deepEqual(result.differences, []);
});

test('the agreement runner reports where the walk and the resolved copy part', async () => {
  const sample = fixtures.filter((f) => f.name.endsWith('cr043-tracked.docx'));
  assert.equal(sample.length, 1);
  const engine = {
    walk: async (pkg, view, all) => `same\n${view}${all ? ' all' : ''}\nwalk`,
    resolved: async (copy, view, all) => `same\n${view}${all ? ' all' : ''}\ncopy`,
  };
  const result = await agreement(sample, engine, (bytes) => WordprocessingMLPackage.load(bytes));
  assert.equal(result.compared, 4);
  assert.deepEqual(result.differences, [
    'revisions/cr043/cr043-tracked.docx accepted: line 3\n  walk: walk\n  copy: copy',
    'revisions/cr043/cr043-tracked.docx accepted (all stories): line 3\n  walk: walk\n  copy: copy',
    'revisions/cr043/cr043-tracked.docx original: line 3\n  walk: walk\n  copy: copy',
    'revisions/cr043/cr043-tracked.docx original (all stories): line 3\n  walk: walk\n  copy: copy',
  ]);
});

// ---------------------------------------------------------------- REQ-306, the understood set

test('the readers\' understood set is 111 of docx4j\'s published 118, and the other 7 are named', () => {
  const published = readFileSync(join(base, 'understood-namespaces.txt'), 'utf8').split('\n').filter((l) => l !== '');
  assert.equal(published.length, 118);
  // load's set is the list's first 107, in its order
  assert.deepEqual([...UNDERSTOOD_NAMESPACES], published.slice(0, 107));
  // the readers add lines 109 to 112, which Word 365 understands (CR-010 section 15)
  assert.deepEqual([...READ_ONLY_UNDERSTOOD_NAMESPACES].sort(), published.slice(108, 112).sort());
  // the difference: no mc:Choice in the corpus requires any of these (CR-012 section 13)
  assert.deepEqual([published[107], ...published.slice(112)], [
    'http://schemas.microsoft.com/office/drawing/2016/11/diagram',
    'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2',
    'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3',
    'http://schemas.microsoft.com/office/spreadsheetml/2016/revision6',
    'http://schemas.microsoft.com/office/spreadsheetml/2016/revision9',
    'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10',
    'http://schemas.microsoft.com/office/spreadsheetml/2017/revision16',
  ]);
});
