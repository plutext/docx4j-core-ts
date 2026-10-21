// CR-014, ranges across paragraphs. Phase 0 is check 52 (test/README.md), which also measures CR-002 section
// 55's D2 (Word's table rule, cases 86 to 95): its file and its Script Lab snippet must agree, since Word runs
// the snippet on the file and nothing can be fixed after the run. Every case's BEGIN paragraph parses with the
// snippet's own parser to the case the file was built from, and every token the snippet searches for is in the
// file once, as a whole word in upper case. Word 365's two saves of the run (test/check52/, outside fixtures/
// so that the weekly RevisionsHarness cuts no golden for them) hold this package's Accept All to Word's.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../dist/index.mjs';
import { check52Bytes, check52Cases } from './fixtures/check52/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));

/** The snippet's CHECK52, loaded as Script Lab runs it, with no page around it. */
async function snippet() {
  const source = await readFile(join(here, 'fixtures', 'check52', 'check52-script-lab.js'), 'utf8');
  return new Function(`${source}\nreturn CHECK52;`)();
}

test("check 52's file and snippet agree: every case parses, every token the snippet looks for is there once", async () => {
  const { parseRegion } = await snippet();
  const pkg = await WordprocessingMLPackage.load(await check52Bytes());
  await pkg.getBody();
  const texts = pkg.body.paragraphs.map((p) => p.text);
  const regions = texts.map(parseRegion).filter((r) => r !== null);
  const cases = check52Cases();
  assert.equal(cases.length, 95);
  assert.deepEqual(regions.map((r) => [r.n, r.kind, r.op, r.tracked]), cases.map((c) => [c.n, c.kind, c.op, c.tracking === 'tracked']));

  const count = (token) => texts.reduce((sum, text) => sum + (text.match(new RegExp(`(?<![A-Za-z0-9])${token}(?![A-Za-z0-9])`, 'g')) ?? []).length, 0);
  const tags = pkg.body.contentControls.map((c) => c.tag);
  for (const r of regions) {
    const tokens = [`BEGIN${r.n}`, `END${r.n}`];
    if (r.range) {
      const { from, to, toMark, whole, control, cell, paragraph } = r.range;
      tokens.push(...[from, to, toMark, whole, cell, paragraph].filter(Boolean));
      if (control) assert.equal(tags.filter((t) => t === control).length, 1, control);
    } else if (r.queries) {
      for (const q of r.queries) tokens.push(...q.split(/\^./).filter((s) => s !== ''));
    } else {
      assert.equal(r.kind, 'VIEWS');
    }
    for (const token of tokens) assert.equal(count(token), 1, `case ${r.n}: ${token}`);
  }
  // every kind of range, every operation and both tracking modes are there
  assert.deepEqual([...new Set(regions.filter((r) => r.range).map((r) => Object.keys(r.range).join('+')))].sort(),
    ['cell', 'control', 'from+to', 'from+toMark', 'paragraph', 'whole']);
  assert.deepEqual([...new Set(regions.map((r) => r.op))].sort(),
    ['bold', 'clear', 'comment', 'control-delete', 'delete', 'ooxml-replace', 'paragraph-after', 'paragraph-before', 'paragraph-delete',
      'read', 'replace', 'search', 'table-end', 'text-end', 'text-start']);
  // CR-002 section 55's cases: the token each names is in a table cell, the control's in one too
  for (const r of regions.filter((x) => x.kind.startsWith('TC'))) {
    const token = r.range.cell ?? r.range.paragraph;
    const paragraph = token ? pkg.body.paragraphs.find((p) => p.text.includes(token)) : pkg.body.contentControls.find((c) => c.tag === r.range.control).paragraphs[0];
    assert.ok(paragraph.parentTableCell, `case ${r.n}`);
  }
});

test("check 52's table for ooxml-replace is a flat OPC package this package reads: a table of one cell, nothing after it", async () => {
  const { TABLE_OOXML } = await snippet();
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  body.insertParagraph('before', 'End');
  const inserted = await body.insertOoxml(TABLE_OOXML, 'End');
  assert.equal(inserted.length, 1);
  assert.deepEqual(body.tables.map((t) => t.values), [[['NEWOOXML']]]);
});

test("check 52's run and its report after Accept All (Word 365, 2026-10-22) are of this file: the same 95 cases, each its kind, operation and tracking", async () => {
  for (const [name, mode] of [['52-ranges-word365-run.json', 'run'], ['52-ranges-word365-accepted.json', 'report']]) {
    const word = JSON.parse(await readFile(join(here, 'fixtures', 'check52', name), 'utf8'));
    assert.equal(word.mode, mode);
    assert.deepEqual(word.cases.map((c) => [c.n, c.kind, c.op, c.tracked]),
      check52Cases().map((c) => [c.n, c.kind, c.op, c.tracking === 'tracked']));
  }
});

const saves = join(here, 'check52');
const saveBytes = (name) => readFile(join(saves, name));

test("check 52's Word saves are Word's, as copied: both files listed, each SHA-1 matching", async () => {
  const manifest = JSON.parse(await readFile(join(saves, 'manifest.json'), 'utf8'));
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await saveBytes(name)).digest('hex'), sha1, name);
  assert.deepEqual(manifest.files.map((f) => f.name), ['52-ranges-word365.docx', '52-ranges-word365-accepted.docx']);
});

/** Each case's paragraphs between its BEGIN and END paragraphs: text, style, alignment, list, table nesting. */
async function regionsOf(pkg) {
  await pkg.getBody();
  const nesting = (p) => {
    let depth = 0;
    for (let cell = p.parentTableCell; cell; cell = cell.parentTable.parentTableCell) depth++;
    return depth;
  };
  const regions = new Map();
  let n = null;
  for (const p of pkg.body.paragraphs) {
    const begin = /^BEGIN(\d+) /.exec(p.text);
    if (begin) { n = Number(begin[1]); regions.set(n, []); continue; }
    if (n !== null && p.text === `END${n}`) { n = null; continue; }
    if (n !== null) regions.get(n).push({ text: p.text, style: p.styleBuiltIn, alignment: p.alignment, listItem: p.isListItem, nesting: nesting(p) });
  }
  return regions;
}

test("check 52: acceptAllRevisions on Word 365's save after the run gives Word's Accept All, region for region (CR-014 section 6.2)", async () => {
  const ours = await WordprocessingMLPackage.load(await saveBytes('52-ranges-word365.docx'));
  assert.ok((await ours.getAllTrackedChanges()).length > 0);
  await ours.acceptAllRevisions();
  assert.equal((await ours.getAllTrackedChanges()).length, 0);
  const mine = await regionsOf(ours);
  const word = await regionsOf(await WordprocessingMLPackage.load(await saveBytes('52-ranges-word365-accepted.docx')));
  assert.equal(word.size, 95);
  assert.deepEqual([...mine.keys()], [...word.keys()]);
  const differing = [...word.keys()].filter((n) => JSON.stringify(mine.get(n)) !== JSON.stringify(word.get(n)));
  // case 26 (CC, its content deleted under tracking): Word fills the emptied control with five spaces, this package
  // leaves it empty; the paragraph is the same otherwise
  assert.deepEqual(differing, [26]);
  assert.deepEqual(word.get(26).map((p) => p.text), ['     ']);
  assert.deepEqual(mine.get(26), word.get(26).map((p) => ({ ...p, text: '' })));
  // the joins: the joined paragraph in the properties of the paragraph the range ended in (HB, BH, LB, BL, TBL)
  assert.deepEqual([16, 17, 18, 19, 21].map((n) => word.get(n).map((p) => [p.text, p.style, p.alignment, p.listItem])), [
    [['HEAD16A  BODY16B', 'Normal', 'Centered', false]],
    [['BODY17A  HEAD17B', 'Heading1', 'Left', false]],
    [['LIST18A  BODY18B', 'Normal', 'Centered', false]],
    [['BODY19A  LIST19B', 'ListParagraph', 'Left', true]],
    [['HEAD21A  BODY21B', 'Normal', 'Centered', false]],
  ]);
});
