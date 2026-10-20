// CR-012 phase 0: the text extraction specification's conformance suite (D8 a), its agreement test
// (D8 b) and the understood set (REQ-306), before there is a text model to hold to them.
//
// test/text-extraction/suite/ is docx4j's docx4j-core-tests text-extraction/ directory, copied
// unedited by scripts/copy-text-extraction-suite.mjs at the commit manifest.json names; the runner
// (text-extraction/conformance.mjs) mirrors docx4j's ExtractionConformance, and the tests below
// mirror its ExtractionConformanceTest's checks of the protocol, which use built tables and engines.
// No engine is installed until phase 2, so every row of the suite is pending, as docx4j's run with
// `-Dtextextraction.adapter=none` reports it.
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
  INSPECTION_LIMITATIONS, HTML_READER_ABSENT,
} from './text-extraction/conformance.mjs';
import { agreement, agreementFixtures } from './text-extraction/agreement.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const root = dirname(here);
const base = join(here, 'text-extraction');
const phase0 = join(base, 'suite', 'phase0');
const manifest = JSON.parse(readFileSync(join(base, 'manifest.json'), 'utf8'));

/** Phase 2 installs the Texts adapter here (docx4j TextsAdapter). */
const ENGINE = null;

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
    implementedStatements: () => ['REQ-901', 'REQ-902', 'REQ-904', 'REQ-905', 'REQ-906', 'REQ-907', 'REQ-1204'],
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

// ---------------------------------------------------------------- the suite, with no engine

test('the suite: every row pending, with its phase, until phase 2 installs an engine', async () => {
  const result = await run(phase0, ENGINE);
  const text = result.report.join('\n');
  assert.ok(text.includes('VBA51_MODULE_MARK'));
  assert.ok(text.includes('ALT47_NESTED_CHUNK'));
  assert.ok(text.includes('the inspector (CR-050)'));
  // docx4j 1d108c0af's suite: 1,583 rows, the same rows in the same order as docx4j's own run, less
  // the 7 of the fifth fixture its CR-053 phase 3 has still to commit (CR-012 section 13)
  assert.equal(result.passed, 0);
  assert.equal(result.pending, 1583, 'the suite changed: recount, and compare with docx4j\'s run');
  for (const line of result.report.slice(1)) assert.match(line, /\tPENDING\t(phase [123]|the inspector \(CR-050\)); .*; Texts adapter not installed$/);
});

// ---------------------------------------------------------------- the agreement test

const fixtures = agreementFixtures();

test('the agreement test\'s fixtures: docx4j\'s 128 revisions fixtures and 14 text-extraction documents', () => {
  assert.equal(fixtures.filter((f) => f.name.startsWith('revisions/')).length, 128);
  assert.equal(fixtures.filter((f) => f.name.startsWith('text-extraction/')).length, 14);
});

test('the agreement test\'s copy route, the reference, runs on every fixture in both views', async () => {
  for (const f of fixtures) {
    const bytes = readFileSync(f.path);
    await (await WordprocessingMLPackage.load(bytes)).acceptAllRevisions();
    await (await WordprocessingMLPackage.load(bytes)).rejectAllRevisions();
  }
  const result = await agreement(fixtures, ENGINE, (bytes) => WordprocessingMLPackage.load(bytes));
  assert.equal(result.compared, 0);
  assert.equal(result.pending, fixtures.length * 4, 'two views, default scope and all stories');
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
