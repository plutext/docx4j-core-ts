// CR-012 phases 0, 2 and 3: the text extraction specification's conformance suite (D8 a), its agreement
// test (D8 b) and the understood set (REQ-306), and this package's text model held to them.
//
// test/text-extraction/suite/ is docx4j's docx4j-core-tests text-extraction/ directory, copied
// unedited by scripts/copy-text-extraction-suite.mjs at the commit manifest.json names; the runner
// (text-extraction/conformance.mjs) mirrors docx4j's ExtractionConformance, and the tests below
// mirror its ExtractionConformanceTest's checks of the protocol, which use built tables and engines.
// Phase 2 installs the engine (text-extraction/adapter.mjs, docx4j's TextsAdapter): the rows whose
// statements it claims run, and the others stay pending with their phase. Phase 3 adds Markdown (format md).
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
  table, normalize, count, verify, verifyGolden, verifyWordSave, imageDestinations, run, INSPECTION_TABLE, INSPECTION_STATUS,
  INSPECTION_LIMITATIONS, HTML_READER_ABSENT, HTML_READER_PRESENT,
} from './text-extraction/conformance.mjs';
import { agreement, agreementFixtures } from './text-extraction/agreement.mjs';
import { textsAdapter, agreementEngine, IMPLEMENTED } from './text-extraction/adapter.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const root = dirname(here);
const base = join(here, 'text-extraction');
const phase0 = join(base, 'suite', 'phase0');
/** docx4j CR-069's suite (`dpp/README.md`), beside phase0. */
const dpp = join(base, 'suite', 'dpp');
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
  // 53-54 (WD 0.5), 55-57 (extraction-probe-structure.docx), 58-67 (extraction-probe-inspect.docx, CR-050), 68-74 (the
  // same fixture, the text a picture draws, CR-050 phase 3), 75-82 (the same fixture, an integer attribute beyond Word's
  // storage width, CR-066 phase 4)
  assert.equal(additions.length, 30);
  appendix.push(...additions);
  assert.equal(generated.length, 82);
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
  assert.ok(result.report.includes(`f.docx\t0\tinspection\t-\t${HTML_READER_ABSENT}\tPASS\t`));
  const moved = { ...expected[1], plain: 'unseen' };
  for (const wrong of [[expected[0]], [expected[0], expected[1], expected[1]], [expected[0], moved]]) {
    reported = wrong;
    await assert.rejects(run(dir, engine), /inspection differs/, `a wrong inspection passed: ${wrong.length}`);
  }
});

test('the inspection table\'s capability column: each environment\'s rows checked in it, the other\'s pending by count, a row pending a ruling left out', async () => {
  // docx4j CR-050 phase 4: the table folds the reader-present rows in under capability, with parsed and pending
  const dir = scratch('inspected-capability');
  writeFileSync(join(dir, 'expected-tokens.tsv'), EMPTY_TOKENS);
  writeFileSync(join(dir, 'expected-findings.tsv'), EMPTY_FINDINGS);
  writeFileSync(join(dir, 'goldens.tsv'), EMPTY_GOLDENS);
  const header = 'fixture\tcase\tkind\tchain\tpart\tpath\tstart\tend_path\tend\taddress\tvisible\treveal\tplain\tmarkdown'
    + '\tblocks\thidden\tembedded\tcomments\texcerpt_length\texcerpt_cut\tselector\tsummary\tcapability\tparsed\tevidence\tpending\n';
  const any = 'h.docx\t0\ttext.hidden\t\t/word/document.xml\t/w:document[1]/w:body[1]/w:p[1]/w:r[1]\t0\t\t4\tbody/0\tno'
    + '\toptions: hidden: true\tomitted\tomitted\tomitted\tunseen\tomitted\tomitted\t4\tfalse\tA_B\t\tany\t\t\t\n';
  const present = 'h.docx\t0\ttext.hidden\t\t/word/chunk.htm\t/html[1]/body[1]/p[1]\t\t\t\t\tno\toptions: hidden: true; html reader present'
    + `\tomitted\tomitted\tomitted\tunseen\tomitted\tomitted\t6\tfalse\tH_SEEN\t\t${HTML_READER_PRESENT}\thtml\tr10-1\t\n`;
  const waiting = present.replace('p[1]', 'p[2]').replace('H_SEEN', 'H_WAITS').replace('\tr10-1\t\n', '\tr11-2\tWD 0.29: a rule not yet ruled on\n');
  writeFileSync(join(dir, INSPECTION_TABLE), header + any + present + waiting);
  writeFileSync(join(dir, INSPECTION_STATUS), `fixture\tcapability\tstatus\tskipped\tundecided\nh.docx\t${HTML_READER_ABSENT}\tPARTIAL\tnot_decoded@|/word/chunk.htm|\t\n`
    + `h.docx\t${HTML_READER_PRESENT}\tCOMPLETE\t\t\n`);
  const [anyRow, seen, waits] = table(join(dir, INSPECTION_TABLE));
  const STATEMENTS = ['REQ-901', 'REQ-902', 'REQ-904', 'REQ-905', 'REQ-906', 'REQ-907', 'REQ-1204', 'REQ-513'];
  let reported;
  const engine = (environment, status, skipped) => ({
    implementedStatements: () => [...STATEMENTS, environment],
    extract: () => { throw new Error('not an extraction'); },
    inspect: async () => ({ findings: reported, status, skipped, undecided: [], limitations: INSPECTION_LIMITATIONS }),
  });
  // with a reader: the any row and the reader's, the pending row left out on both sides whatever the inspector reports there
  reported = [anyRow, seen, { ...waits, plain: 'unseen' }];
  const withReader = await run(dir, engine(HTML_READER_PRESENT, 'COMPLETE', []));
  assert.ok(withReader.report.includes(`h.docx\t0\tinspection\t-\t${HTML_READER_PRESENT}\tPASS\t`));
  assert.ok(withReader.report.includes(`h.docx\t0\tinspect:text.hidden:H_WAITS\t-\t${HTML_READER_PRESENT}\tPENDING\tWD 0.29: a rule not yet ruled on`));
  reported = [anyRow, { ...seen, parsed: 'xml' }];
  await assert.rejects(run(dir, engine(HTML_READER_PRESENT, 'COMPLETE', [])), /inspection differs[\s\S]*parsed/);
  // without one: the any row alone, with that environment's status, the reader's two rows pending by count
  reported = [anyRow];
  const without = await run(dir, engine(HTML_READER_ABSENT, 'PARTIAL', ['not_decoded@|/word/chunk.htm|']));
  assert.ok(without.report.includes(`h.docx\t0\tinspection\t-\t${HTML_READER_ABSENT}\tPASS\t`));
  assert.ok(without.report.includes(`h.docx\t0\tinspection\t-\t${HTML_READER_PRESENT}\tPENDING\t2 rows of ${HTML_READER_PRESENT}, not this run's environment`));
  reported = [anyRow, seen];
  await assert.rejects(run(dir, engine(HTML_READER_ABSENT, 'PARTIAL', ['not_decoded@|/word/chunk.htm|'])), /inspection differs/);
});

test('Word\'s text save is read by its encoding and compared by line (WD REQ-1002\'s NOTE)', () => {
  // docx4j ExtractionConformanceTest.wordSaveIsReadByItsEncodingAndComparedByLine
  const dir = scratch('word-save');
  const cp1252 = join(dir, 'cp1252.txt');
  // Windows-1252 has é at 0xE9; Word wrote ? for U+2082, which the code page lacks
  writeFileSync(cp1252, Buffer.from('L1_ H?O caf\xe9\r\nC04_\r1.3161.335\rL3_ end.\r\n', 'latin1'));
  const text = 'L1_ H\u2082O caf\u00e9\nC04_\t1.3161.335\nL3_ end.\n';
  verifyWordSave(cp1252, 'windows-1252', text);
  // a character Word did not write; a break between cells lost
  for (const wrong of [text.replace('H\u2082O', 'H2O'), text.replace('C04_\t', 'C04_ ')]) {
    assert.throws(() => verifyWordSave(cp1252, 'windows-1252', wrong), /: line /, wrong);
  }
  assert.throws(() => verifyWordSave(cp1252, 'UTF-8', text), /no byte order mark/);
  const utf8 = join(dir, 'utf8.txt');
  writeFileSync(utf8, '\ufeff\u0130STANBUL F\u0130LM\r\n');
  verifyWordSave(utf8, 'UTF-8', '\u0130STANBUL F\u0130LM\n');
  assert.throws(() => verifyWordSave(utf8, 'UTF-8', 'ISTANBUL FILM\n'), /: line 1/);
  assert.throws(() => verifyWordSave(utf8, 'windows-1252', '\u0130STANBUL F\u0130LM\n'), /a byte order mark/);
});

// ---------------------------------------------------------------- the suite

test('the suite with no engine: every row pending, with its phase', async () => {
  const result = await run(phase0, null);
  const text = result.report.join('\n');
  assert.ok(text.includes('VBA51_MODULE_MARK'));
  assert.ok(text.includes('ALT47_NESTED_CHUNK'));
  assert.ok(text.includes('the inspector (CR-050)'));
  // docx4j 17.3.5's suite: 1,772 rows, the same rows in the same order as docx4j's own run (CR-012 sections 13.3 and 24.2);
  // an inspection's rows of the other environment pend by count, whatever the engine
  assert.equal(result.passed, 0);
  assert.equal(result.pending, 1772, 'the suite changed: recount, and compare with docx4j\'s run');
  for (const line of result.report.slice(1)) {
    assert.match(line, /\tPENDING\t((phase [123]|the inspector \(CR-050\)); .*; Texts adapter not installed|\d+ rows? of html reader present, not this run's environment)$/);
  }
  const caps = await run(dpp, null);
  assert.equal(caps.pending, 192);
  for (const line of caps.report.slice(1)) assert.match(line, /\tPENDING\t(phase [123]|CR-069 phase [1-4]|the inspector \(CR-050\)); .*; Texts adapter not installed$/);
});

test('the suite with the text model (CR-012 phases 2 and 3): every row it claims passes, the others pend for a statement it does not claim', async () => {
  const result = await run(phase0, ENGINE);
  // run() throws at the first row that fails; these are the counts at docx4j 17.3.5's suite (CR-012 section 24.2: the
  // 12 altchunk.not_expanded rows re-keyed by part name, section 22.2 (a), among the 786)
  assert.equal(result.passed + result.pending, 1772);
  assert.equal(result.passed, 786, 'a row moved: recount (CR-012 sections 15, 16 and 24)');
  assert.equal(result.pending, 986);
  const claimed = new Set(IMPLEMENTED);
  // an inspection's rows of the other environment pend by count, with no statement missing
  const byStatement = result.report.slice(1).filter((l) => l.includes('\tPENDING\t') && !/\tPENDING\t\d+ rows? of /.test(l));
  for (const line of byStatement) {
    const missing = line.split('\t')[6].split('; ')[1].split(', ');
    assert.ok(missing.length > 0 && missing.every((r) => !claimed.has(r)), line);
  }
  assert.equal(result.report.length - 1 - result.passed - byStatement.length, 3, 'three fixtures\' reader-present rows, by count');
  // what pends, by what it waits for: phase 5's statements, the inspector, and REQ-303 and REQ-1203
  const waits = (statement) => byStatement.filter((l) => l.split('; ')[1].split(', ').includes(statement)).length;
  assert.equal(waits('REQ-303'), 9, 'case 6: the URI decides; this package matches by prefix (REQ-303\'s permitted route)');
  assert.equal(waits('REQ-1203'), 12, 'the whole-document goldens hold phase 5\'s content');
  for (const statement of ['REQ-506', 'REQ-511', 'REQ-512', 'REQ-513', 'REQ-604', 'REQ-901']) assert.ok(waits(statement) > 0, statement);
  assert.equal(waits('REQ-703'), 0, 'Markdown (phase 3)');
  for (const statement of ['REQ-1002', 'REQ-607', 'REQ-521', 'REQ-710']) assert.equal(waits(statement), 0, statement);
});

test('the dpp suite (docx4j CR-069, CR-012 section 24): every row checked but the inspector\'s and the whole-document goldens', async () => {
  // docx4j ExtractionConformanceTest.dppSuiteIsCheckedOrPendingOnItsPhase, its rows; Word's 22 saves under profile: word-text
  const result = await run(dpp, ENGINE);
  const text = result.report.join('\n');
  assert.equal(result.passed + result.pending, 192, 'the same rows as docx4j\'s own run (target/text-extraction/dpp.tsv)');
  const word = '\taccepted\tprofile: word-text\t';
  for (const save of ['bisect/control.docx\t0\tword save:bisect/control.txt', 'bisect/C03.docx\t3\tword save:bisect/C03.txt',
    'bisect/C08.docx\t8\tword save:bisect/C08.txt', 'bisect/C05.docx\t5\tword save:bisect/C05.txt',
    'caps/caps-lang7d.docx\t5\tword save:caps/caps-lang7d-utf8.txt', 'caps/caps-lang6a.docx\t5\tword save:caps/caps-lang6a.txt',
    'bisect/C02.docx\t2\tword save:bisect/C02.txt', 'bisect/C06.docx\t6\tword save:bisect/C06.txt', 'bisect/C07.docx\t7\tword save:bisect/C07.txt']) {
    assert.ok(text.includes(`${save}${word}PASS`), save);
  }
  assert.equal(result.report.filter((l) => l.includes('\tword save:') && l.endsWith('\tPASS\t')).length, 22, 'Word\'s 22 saves');
  // phase 1: the case mapping, against the tables
  assert.ok(text.includes('caps/caps-lang5.docx\t5\tH\u2082O\toriginal\tdefault\tPASS'));
  assert.ok(text.includes('caps/caps-lang6a.docx\t5\t\u0130STANBUL F\u0130LM\taccepted\tsmall_caps: upper\tPASS'));
  assert.ok(text.includes('dpp-forks-word365.docx\t5\tROBERTS\taccepted\tmarkdown\tPASS'));
  assert.ok(text.includes('dpp-forks-word365.docx\t5\tJackson\tmarkup\tdefault\tPASS'));
  // phase 2: the drop-cap join and its warning, in every view
  for (const view of ['accepted', 'original', 'markup']) {
    assert.ok(text.includes(`dpp-forks-word365.docx\t2\twarning:paragraph.drop_cap_joined:/word/document.xml\t${view}\tdefault\tPASS`));
    assert.ok(text.includes(`bisect/C02.docx\t2\twarning:paragraph.drop_cap_joined:/word/document.xml\t${view}\tdefault\tPASS`));
  }
  // phase 3: the tags in Markdown, the baseline in plain text
  for (const token of ['H<sub>2</sub>O', 'x<sup>2</sup>', '1<sup>st</sup>', '48,210<sup>3</sup>']) {
    assert.ok(text.includes(`dpp-forks-word365.docx\t${token.startsWith('48') ? 7 : 6}\t${token}\taccepted\tmarkdown\tPASS`), token);
  }
  assert.ok(text.includes('dpp-forks-word365.docx\t7\t48,2103\taccepted\tdefault\tPASS'));
  // what pends: the inspector's six (phase 4's text.formatting_join among them) and the two whole-document goldens
  const pending = result.report.filter((l) => l.includes('\tPENDING\t'));
  assert.equal(pending.length, 8);
  assert.equal(pending.filter((l) => l.includes('\tinspect:')).length, 6);
  assert.equal(pending.filter((l) => l.includes('\tgolden:') && l.endsWith('; REQ-1203')).length, 2);
});

test('the dpp suite\'s two goldens (REQ-1203, not claimed until phase 5): exact', async () => {
  // docx4j's ExtractionGoldens writes the text with its trailing line ends trimmed to one
  const asWritten = (text) => `${text.replace(/\n+$/, '')}\n`;
  for (const row of table(join(dpp, 'goldens.tsv'))) {
    const ours = asWritten(imageDestinations((await ENGINE.extract(join(dpp, row.fixture), row.view, row.scenario, row.format)).text));
    assert.equal(ours, readFileSync(join(dpp, row.golden), 'utf8'), row.golden);
  }
});

test('the whole-document goldens (REQ-1203, not claimed until phase 5): exact but for phase 5\'s content and case 6', async () => {
  // docx4j's ExtractionGoldens writes the text with its trailing line ends trimmed to one
  const asWritten = (text) => `${text.replace(/\n+$/, '')}\n`;
  // the lines phase 5 adds (5b LaTeX, 5c the chart and SmartArt, 5e the chunks) and case 6's (REQ-303 by prefix)
  const PENDING = /MCE6_|MTH32|CHT34_|SMA35_|ALT4[3-7]_|^\[chart: /;
  const without = (text) => text.split('\n').filter((line) => !PENDING.test(line)).join('\n');
  // in Markdown, each block a paragraph apart, its underscores escaped
  const PENDING_MD = /MCE6\\_|MTH32|CHT34\\_|SMA35\\_|ALT4[3-7]\\_/;
  const withoutMd = (text) => text.split('\n\n').filter((block) => !PENDING_MD.test(block)).join('\n\n');
  let exact = 0;
  for (const row of table(join(phase0, 'goldens.tsv'))) {
    const ours = asWritten(imageDestinations((await ENGINE.extract(join(phase0, row.fixture), row.view, row.scenario, row.format)).text));
    const golden = readFileSync(join(phase0, row.golden), 'utf8');
    if (ours === golden) { exact++; continue; }
    assert.equal(row.fixture, 'extraction-probe.docx', `${row.golden} differs`);
    if (row.format === 'md') assert.equal(withoutMd(ours), withoutMd(golden), row.golden);
    else assert.equal(without(ours), without(golden), row.golden);
  }
  assert.equal(exact, 8, 'the external and structure probes\' goldens, in every view and in Markdown');
});

// ---------------------------------------------------------------- the agreement test

const fixtures = agreementFixtures();

test('the agreement test\'s fixtures: docx4j\'s 174 revisions fixtures and 38 text-extraction documents', () => {
  // at docx4j 17.3.5: CR-068's revisions/cr068/, check 54's and 57's Word saves under revisions/core-ts/, CR-043 section
  // 17's probes; the dpp suite's fixtures and bind-props.docx
  assert.equal(fixtures.filter((f) => f.name.startsWith('revisions/')).length, 174);
  assert.equal(fixtures.filter((f) => f.name.startsWith('text-extraction/')).length, 38);
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
