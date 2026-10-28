// Check 72 (CR-016 D5, 2026-10-27; Jason: "D5 do check 72"): one file per misplaced child docx4j CR-064's rules refuse, for
// Word 365 to open, so that "Word repairs it" becomes a measurement. The files are built here and copied to the share; this
// test holds the build and records what this package does with each file: before CR-016 phase 2 (2026-10-27) the loads
// never failed, the readers saw almost none of the misplaced content, and a save wrote most of it back as it was; since phase
// 2 (2026-10-28) the load warns of each (target.misplaced, D7), the save is refused (D2) unless pkg.checkContentOnSave is
// off, when it writes as before. misplaced-content.test.mjs holds the fields and the rules; this file the files' fates.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { unzipSync, strFromU8 } from 'fflate';
import { WordprocessingMLPackage, MisplacedContentException } from '../dist/index.mjs';
import { FORMS, check72Files, documentOf } from './fixtures/check72/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));

test("check 72's files: one misplaced child each, between BEFORE and AFTER, one per rule of CR-064's five sets", async () => {
  const files = await check72Files();
  assert.equal(files.length, 18);
  assert.deepEqual(files.map((f) => f.name), FORMS.map((f) => f.name));
  const sets = new Set(FORMS.map((f) => f.set.replace(/ \(.*\)$/, '')));
  assert.deepEqual([...sets].sort(), ['block-level container', 'row', 'run', 'run-level container', 'table']);
  for (const f of files) {
    const xml = strFromU8(unzipSync(f.bytes)['word/document.xml']);
    assert.equal(xml, documentOf(f), f.name);
    assert.ok(/<w:t[ >][^<]*BEFORE<\/w:t>[\s\S]*<w:t[ >][^<]*AFTER<\/w:t>/.test(xml), `${f.name}: BEFORE then AFTER`);
    assert.ok(f.name === '72e-break-in-body' || f.name === '72j-break-in-paragraph' || f.name === '72g-table-in-paragraph' || xml.includes('MISPLACED'), `${f.name}: the child marked`);
  }
});

/**
 * What this package does with each file (measured 2026-10-27, before CR-016 phase 2; the same since under the opt-out): whether
 * the readers see the misplaced content (`read`), and whether a save under `checkContentOnSave = false` writes it back
 * (`saved`). Not what Word does: that is the check.
 *
 * A bare break (72e, 72j) carries no MISPLACED marker, so its `saved` is the `w:br` itself, kept where it was: both are (the facade
 * reads each typed, `Br`, and writes it as held; objects CR-017's finding of 2026-10-28, which corrected an earlier reading here).
 */
const AS_HELD = {
  '72a-run-in-body': { read: false, saved: true }, '72b-text-in-body': { read: false, saved: true },
  '72c-cell-in-body': { read: true, saved: true }, '72d-row-in-body': { read: true, saved: true },
  '72e-break-in-body': { read: false, saved: true }, '72f-run-in-cell': { read: false, saved: true },
  '72g-table-in-paragraph': { read: false, saved: true }, '72h-paragraph-in-paragraph': { read: false, saved: true },
  '72i-text-in-paragraph': { read: false, saved: true }, '72j-break-in-paragraph': { read: false, saved: true },
  '72k-run-in-run': { read: false, saved: true }, '72l-paragraph-in-run': { read: false, saved: true },
  '72m-string-in-run': { read: false, saved: false }, '72n-cell-in-table': { read: false, saved: true },
  '72o-paragraph-in-table': { read: false, saved: true }, '72p-table-in-table': { read: false, saved: true },
  '72q-paragraph-in-row': { read: false, saved: true }, '72r-run-in-row': { read: false, saved: true },
};

test("every check 72 file loads with one target.misplaced warning, the readers see a misplaced child only in a cell or a row in the body, the save is refused but for the string in a run (gone at the read), and under the opt-out writes every one back but that string (CR-016 sections 7 and 8)", async () => {
  for (const f of await check72Files()) {
    const pkg = await WordprocessingMLPackage.load(f.bytes);
    const body = await pkg.getBody();
    const isString = f.name === '72m-string-in-run';
    assert.deepEqual(pkg.getMainDocumentPart().warnings.map((w) => w.code), ['target.misplaced'], `${f.name}: the load says what is wrong (D7)`);
    const isTable = f.name === '72g-table-in-paragraph';
    const isBreak = f.name.includes('break');
    const read = isTable ? body.tables.length > 0 : isBreak ? body.text.includes('\u000b') : body.text.includes('MISPLACED');
    if (isString) await pkg.save(); // the runtime dropped the string at the read: nothing misplaced is in the tree to refuse
    else await assert.rejects(pkg.save(), (e) => e instanceof MisplacedContentException && e.findings.length === 1, `${f.name}: refused (D2)`);
    pkg.checkContentOnSave = false;
    const saved = strFromU8(unzipSync(await pkg.save())['word/document.xml']);
    assert.equal(pkg.lastSave.misplaced.length, isString ? 0 : 1, `${f.name}: reported under the opt-out`);
    const kept = isTable ? saved.includes('<w:tbl>')
      : isBreak ? saved.includes(f.name === '72e-break-in-body' ? '</w:p><w:br/><w:p>' : '</w:r><w:br/><w:r>')
      : saved.includes('MISPLACED');
    assert.deepEqual({ read, saved: kept }, AS_HELD[f.name], f.name);
    assert.ok(body.paragraphs.some((p) => p.text === 'BEFORE') && body.paragraphs.some((p) => p.text === 'AFTER'), `${f.name}: BEFORE and AFTER read`);
  }
});

// --- Word 365's run (Jason, 2026-10-27; CR-016 section 7.1): 11 files don't open ("Word experienced an error trying to open the
// file", no repair offered); 7 opened with no message, repaired silently

/** The files Word 365 opened with no error, and the body of its save (rsids, paraIds and the table borders left out). */
const WORD_OPENED = {
  // the break goes into the next paragraph, as its first run
  '72e-break-in-body': '<w:p><w:r><w:t>BEFORE</w:t></w:r></w:p><w:p><w:r><w:br/></w:r><w:r><w:t>AFTER</w:t></w:r></w:p>',
  // the break goes into the run that follows it
  '72j-break-in-paragraph': '<w:p><w:r><w:t>a run, then a bare break</w:t></w:r><w:r><w:br/><w:t>then a run</w:t></w:r></w:p>',
  // the inner paragraph's run joins the outer paragraph, and its mark becomes an empty paragraph after
  '72l-paragraph-in-run': '<w:p><w:r><w:t xml:space="preserve">outer text </w:t></w:r><w:r><w:t>MISPLACED paragraph in a run</w:t></w:r></w:p><w:p/><w:p><w:r><w:t>AFTER</w:t></w:r></w:p>',
  // the string is dropped, with no message: its paragraph is left empty
  '72m-string-in-run': '<w:p><w:r><w:t>BEFORE</w:t></w:r></w:p><w:p/><w:p><w:r><w:t>AFTER</w:t></w:r></w:p>',
  // the paragraph goes before the table, in the body
  '72o-paragraph-in-table': '<w:p><w:r><w:t>MISPLACED paragraph in a table</w:t></w:r></w:p><w:tbl>',
  // the nested table becomes the outer table's first row, its one cell filling the first column (gridAfter 1)
  '72p-table-in-table': '<w:tr><w:trPr><w:gridAfter w:val="1"/><w:wAfter w:w="3000" w:type="dxa"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>MISPLACED nested table</w:t></w:r></w:p></w:tc></w:tr><w:tr>',
  // the paragraph goes into the row's first cell, before its paragraph
  '72q-paragraph-in-row': '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>MISPLACED paragraph in a row</w:t></w:r></w:p><w:p><w:r><w:t>CELL A1</w:t></w:r></w:p></w:tc>',
};
const WORD_ERROR = ['72a-run-in-body', '72b-text-in-body', '72c-cell-in-body', '72d-row-in-body', '72f-run-in-cell', '72g-table-in-paragraph',
  '72h-paragraph-in-paragraph', '72i-text-in-paragraph', '72k-run-in-run', '72n-cell-in-table', '72r-run-in-row'];

const bodyOf = (xml) => xml.replace(/ w:rsid\w*="[^"]*"| w14:(?:paraId|textId)="[^"]*"/g, '').replace(/<w:tblBorders>.*?<\/w:tblBorders>/g, '')
  .replace(/^[\s\S]*<w:body>/, '').replace(/<w:sectPr[\s\S]*$/, '');

test("check 72, Word 365's run: 11 files don't open, no repair offered; 7 open with no message and are repaired silently, a break, a paragraph or a nested table moved into a neighbour, a string in a run dropped (CR-016 section 7.1)", async () => {
  assert.deepEqual([...Object.keys(WORD_OPENED), ...WORD_ERROR].sort(), FORMS.map((f) => f.name).sort(), 'every file accounted for');
  const manifest = JSON.parse(await readFile(join(here, 'check72', 'manifest.json'), 'utf8'));
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await readFile(join(here, 'check72', name))).digest('hex'), sha1, name);
  assert.deepEqual(manifest.files.filter((f) => f.name.endsWith('.docx')).map((f) => f.name), Object.keys(WORD_OPENED).map((n) => `${n}-word365.docx`));
  for (const [name, expected] of Object.entries(WORD_OPENED)) {
    const body = bodyOf(strFromU8(unzipSync(new Uint8Array(await readFile(join(here, 'check72', `${name}-word365.docx`))))['word/document.xml']));
    assert.ok(body.includes(expected), `${name}: Word's save holds the repair\n${body}`);
    assert.ok(body.includes('<w:t>BEFORE</w:t>') && body.includes('<w:t>AFTER</w:t>'), `${name}: BEFORE and AFTER kept`);
    const hasText = !name.includes('break');
    assert.equal(body.includes('MISPLACED'), hasText && name !== '72m-string-in-run', `${name}: the child's text kept, but the string in a run`);
    // and this package reads Word's save as any file: the text where Word put it
    const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile(join(here, 'check72', `${name}-word365.docx`))));
    const text = (await pkg.getBody()).text;
    assert.equal(text.includes('MISPLACED'), hasText && name !== '72m-string-in-run', `${name}: read back`);
    if (!hasText) assert.ok(text.includes('\u000b'), `${name}: the break read where Word put it`);
    assert.deepEqual(pkg.getMainDocumentPart().warnings, [], name);
  }
});
