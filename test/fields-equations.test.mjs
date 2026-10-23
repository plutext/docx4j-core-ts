// CR-002 section 62: a paragraph break inside a field's result, and a tracked delete of a paragraph holding an
// equation (docx4j-python's review, items B2 and B11). Check 57 (test/README.md) measured what Word does: a file for a
// Script Lab snippet, a region per case, two of them only opened. The snippet's file and the snippet must agree, since
// Word runs one on the other and nothing can be fixed after the run. Section 62.2 has Word's run, and D1, D2, F1 and F2,
// which this package follows since: a w:fldSimple made a complex field before a break splits it, an equation deleted
// under tracking in Word's form, Word's form accepted and read as Word has it.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../dist/index.mjs';
import { NotSupportedError } from '../dist/office-js/index.mjs';
import { Word } from '../dist/office-js/index.mjs';
import { check57Bytes, check57Cases, documentOf, documentXmlOf, packageOf, sourceParagraph } from './fixtures/check57/build.mjs';
import { strFromU8, unzipSync } from 'fflate';

const here = dirname(fileURLToPath(import.meta.url));

const snippet = async () => new Function(`${await readFile(join(here, 'fixtures', 'check57', 'check57-script-lab.js'), 'utf8')}\nreturn CHECK57;`)();

/** Word 365's run of check 57 (CR-002 section 62.2), as its snippet reported it. */
const wordRun = async () => JSON.parse(await readFile(join(here, 'fixtures', 'check57', '57-fields-equations-word365-run.json'), 'utf8'));
const saved = (name) => readFile(join(here, 'check57', name));
/** A string the snippet made visible (`vis`), as it was. */
const unvis = (s) => s.replace(/\\u([0-9a-f]{4})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
/** A region's OOXML as two runs can compare it: no revision ids, authors or dates, no xml:space, no getOoxml's last empty paragraph. */
const comparable = (xml) => xml.replace(/ (?:w:id|w:author|w:date|w16du:dateUtc|xml:space)="[^"]*"/g, '').replace(/<w:p\/>$/, '');
/** A region's OOXML split in paragraphs: its BEGIN paragraph, the case's, ..., its END paragraph. */
const paragraphsOf = (xml) => xml.match(/<w:p\b(?:(?!<w:p\b)[\s\S])*?<\/w:p>|<w:p\/>/g);

/** A saved document's body in paragraphs, compared as Word saves: no rsids, no ids or dates, a run of the same properties as its neighbour's joined to it. */
function bodyParagraphs(bytes) {
  let body = /<w:body>([\s\S]*)<\/w:body>/.exec(strFromU8(unzipSync(bytes)['word/document.xml']))[1]
    .replace(/<w:sectPr\b[\s\S]*<\/w:sectPr>/, '').replace(/<w:lastRenderedPageBreak\/>/g, '')
    .replace(/ (?:w:rsid\w*|w14:paraId|w14:textId|w:id|w:author|w:date|w16du:dateUtc|xml:space)="[^"]*"| xmlns:\w+="[^"]*"/g, '');
  for (let was = ''; was !== body;) {
    was = body;
    body = body.replace(/<w:r>((?:<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>)?)<w:t>([^<]*)<\/w:t><\/w:r><w:r>\1<w:t>([^<]*)<\/w:t><\/w:r>/g, '<w:r>$1<w:t>$2$3</w:t></w:r>')
      .replace(/<m:r>((?:<w:del>)?(?:<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>)?)<m:t>([^<]*)<\/m:t>((?:<\/w:del>)?)<\/m:r><m:r>\1<m:t>([^<]*)<\/m:t>\3<\/m:r>/g, '<m:r>$1<m:t>$2$4</m:t>$3</m:r>');
  }
  return paragraphsOf(body);
}

/** The paragraph after a case's BEGIN paragraph, of a list of paragraphs as `bodyParagraphs` gives them. */
const caseParagraphOf = (paragraphs, n) => paragraphs[paragraphs.findIndex((p) => p.includes(`>BEGIN${n} `)) + 1];

test("check 57's snippet file and snippet agree: every case parses to the case it was built as, of its shape", async () => {
  const { parseRegion } = await snippet();
  const pkg = await WordprocessingMLPackage.load(await check57Bytes());
  const paragraphs = (await pkg.getBody()).paragraphs;
  const cases = check57Cases();
  assert.equal(cases.length, 10);
  assert.deepEqual(paragraphs.map((p) => parseRegion(p.text)).filter((r) => r !== null), cases);
  for (const c of cases) {
    const at = paragraphs.findIndex((p) => p.text === `BEGIN${c.n} ${c.op} ${c.shape}`);
    const p = paragraphs[at + 1];
    assert.equal(paragraphs[at + 2].text, `END${c.n}`);
    const xml = JSON.stringify(p.p, (k, v) => (k === 'PARENT' ? undefined : v));
    const has = (typeName) => xml.includes(`"TYPE_NAME":"${typeName}"`);
    switch (c.shape) {
      case 'complex-field': assert.ok(has('org_docx4j_wml.FldChar') && !has('org_docx4j_wml.CTSimpleField'), `case ${c.n}`); break;
      case 'simple-field': assert.ok(has('org_docx4j_wml.CTSimpleField') && !has('org_docx4j_wml.FldChar'), `case ${c.n}`); break;
      case 'display-equation': assert.ok(has('org_docx4j_math.CTOMathPara'), `case ${c.n}`); break;
      default: assert.ok(has('org_docx4j_math.CTOMath') && !has('org_docx4j_math.CTOMathPara'), `case ${c.n}`);
    }
    // the words the snippet searches for, each once in the document
    const words = c.op.startsWith('insert') ? [`FA${c.n}`] : c.op === 'range-delete' && c.shape === 'inline-equation' ? [`L${c.n}`, `R${c.n}`] : [];
    for (const w of words) assert.equal(paragraphs.filter((q) => q.text.split(/\W+/).includes(w)).length, 1, `case ${c.n}: ${w}`);
    assert.equal(c.op === 'open', xml.includes('"TYPE_NAME":"org_docx4j_wml.RunDel"'), `case ${c.n}`);
  }
});

test("check 57's run: Word's JSON is of this file's cases, and its saves are Word's, as copied", async () => {
  const manifest = JSON.parse(await readFile(join(here, 'check57', 'manifest.json'), 'utf8'));
  assert.deepEqual(manifest.files.map((f) => f.name), ['57-fields-equations-word365.docx', '57-fields-equations-word365-accepted.docx', '57-fields-equations-word365-rejected.docx']);
  for (const { name, sha1 } of manifest.files) assert.equal(createHash('sha1').update(await saved(name)).digest('hex'), sha1, name);
  const word = await wordRun();
  assert.equal(word.host.version, '16.0.20326.20232');
  assert.deepEqual(word.cases.map(({ n, op, shape }) => ({ n, op, shape })), check57Cases());
});


test("CR-002 section 62.2: check 57's snippet through the Word shim: every case goes ahead, and only Range.fields is missing", async () => {
  const { runCheck } = await snippet();
  const pkg = await WordprocessingMLPackage.load(await check57Bytes());
  const ours = await runCheck((f) => Word.run(pkg, f), 'run');
  assert.equal(ours.cases.length, 10);
  assert.deepEqual(ours.cases.filter((c) => c.error !== undefined), []);
  const errors = new Set(ours.cases.flatMap((c) => Object.values(c.region)).filter((v) => typeof v === 'string' && v.startsWith('error:')));
  assert.deepEqual([...errors], ['error: NotImplemented: Word.Range.fields is not supported by @docx4j/core-ts']);
  await pkg.save();
});

test("CR-002 section 62.2: check 57's snippet through the Word shim gives Word 365's run, case for case", async () => {
  const { runCheck } = await snippet();
  const pkg = await WordprocessingMLPackage.load(await check57Bytes());
  const ours = Object.fromEntries((await runCheck((f) => Word.run(pkg, f), 'run')).cases.map((c) => [c.n, c]));
  const word = Object.fromEntries((await wordRun()).cases.map((c) => [c.n, c]));
  /** A region's words of case `from` (keep1a, FA1, END1) as case `to`'s. */
  const renumbered = (xml, from, to) => xml.replace(new RegExp(`(?<=[A-Za-z])${from}(?!\\d)`, 'g'), String(to));
  const caseParagraph = (c) => paragraphsOf(comparable(c.region.ooxml))[1];
  for (let n = 1; n <= 10; n++) assert.deepEqual(ours[n].result, word[n].result, `case ${n}`);
  // B2: a break inside a field's result splits it, the field running across the paragraphs; a w:fldSimple made a
  // complex field first (D1)
  for (const n of [1, 2, 3, 4]) {
    for (const k of ['text', 'paragraphs', 'trackedChanges']) assert.deepEqual(ours[n].region[k], word[n].region[k], `case ${n}: ${k}`);
    assert.equal(comparable(ours[n].region.ooxml), comparable(word[n].region.ooxml), `case ${n}`);
  }
  assert.deepEqual(word[2].region.fields, [{ code: ' QUOTE "FA2 FB2" ', result: 'FA2x\\u000dy FB2' }]);
  // B11: a tracked Paragraph.delete and a tracked range delete taking an equation write Word's form (D2): a w:del in
  // each m:r and in the structure's m:ctrlPr, the runs either side in w:dels of their own
  for (const n of [7, 8, 9, 10]) assert.equal(comparable(ours[n].region.ooxml), comparable(word[n].region.ooxml), `case ${n}`);
  // cases 5 and 6 hold the one w:del this package wrote before; Word read them as what it writes now
  for (const [n, as] of [[5, 7], [6, 8]]) {
    assert.equal(renumbered(caseParagraph(word[n]), n, as), caseParagraph(ours[as]), `case ${n}`);
    assert.match(caseParagraph(ours[n]), /<w:del><(?:m:oMathPara|w:r)>/, `case ${n}`);
  }
  // one change a case, as Word lists them. Word's text under tracking is All Markup's, the deleted text in it (CR-014
  // section 6.2), where this package's is Simple Markup's; a change's text is docx4j's reader's, the equation's m:t
  // (none of it in the one w:del form), where Word's is the equation as it draws it
  for (let n = 5; n <= 10; n++) {
    assert.equal(ours[n].region.paragraphs.length, word[n].region.paragraphs.length, `case ${n}`);
    assert.deepEqual(ours[n].region.trackedChanges.map((c) => c.type), word[n].region.trackedChanges.map((c) => c.type), `case ${n}`);
  }
  assert.deepEqual([5, 6, 7, 8, 9, 10].map((n) => ours[n].region.paragraphs[1]), ['', '', '', '', '', 'keep10a  keep10b']);
  assert.deepEqual([5, 6, 7, 8, 9, 10].map((n) => ours[n].region.trackedChanges[0].text),
    ['\\u000d', 'keep6a L6  R6 keep6b\\u000d', 'E=mc2\\u000d', 'keep8a L8 E=mc2 R8 keep8b\\u000d', 'E=mc2', 'L10 E=mc2 R10']);
  assert.equal(word[10].region.trackedChanges[0].text, 'L10 \\ud835\\udc38=\\ud835\\udc5a,\\ud835\\udc50-2. R10');
});

test("CR-002 section 62.2: check 57's run here, saved, accepted and rejected, gives Word's saves region for region", async () => {
  const { runCheck } = await snippet();
  const pkg = await WordprocessingMLPackage.load(await check57Bytes());
  await runCheck((f) => Word.run(pkg, f), 'run');
  const run = await pkg.save();
  const regions = (paragraphs) => paragraphs.slice(1, -1);
  // the save after the run: every case's paragraphs Word's, but cases 5 and 6, which keep the file's form
  const ours = bodyParagraphs(run);
  const word = bodyParagraphs(await saved('57-fields-equations-word365.docx'));
  assert.equal(ours.length, word.length);
  const differ = regions(ours).map((p, i) => (p === regions(word)[i] ? null : p)).filter((p) => p !== null);
  assert.deepEqual(differ, [caseParagraphOf(ours, 5), caseParagraphOf(ours, 6)]);
  for (const [verb, name] of [['acceptAllRevisions', 'accepted'], ['rejectAllRevisions', 'rejected']]) {
    const resolved = await WordprocessingMLPackage.load(run);
    await resolved[verb]();
    assert.deepEqual(bodyParagraphs(await resolved.save()), bodyParagraphs(await saved(`57-fields-equations-word365-${name}.docx`)), verb);
  }
});

test("CR-002 section 62.2: Word's save after check 57's run, accepted and rejected here, gives Word's Accept All and Reject All (F1)", async () => {
  const run = await saved('57-fields-equations-word365.docx');
  for (const [verb, name] of [['acceptAllRevisions', 'accepted'], ['rejectAllRevisions', 'rejected']]) {
    const pkg = await WordprocessingMLPackage.load(run);
    await pkg[verb]();
    assert.deepEqual(bodyParagraphs(await pkg.save()), bodyParagraphs(await saved(`57-fields-equations-word365-${name}.docx`)), verb);
  }
  // accepted, an equation deleted whole goes with its m:oMath and m:oMathPara, and its paragraph where the mark went
  const accepted = bodyParagraphs(await saved('57-fields-equations-word365-accepted.docx'));
  assert.equal(caseParagraphOf(accepted, 9), '<w:p/>');
  assert.ok(accepted.every((p) => !p.includes('oMath')));
});

test("CR-002 section 62.2: this package's readers on Word's saves of check 57 (F2)", async () => {
  const word = Object.fromEntries((await wordRun()).cases.map((c) => [c.n, c]));
  const caseParagraphs = async (bytes) => {
    const paragraphs = (await (await WordprocessingMLPackage.load(bytes)).getBody()).paragraphs;
    return Object.fromEntries(check57Cases().map((c) => [c.n, paragraphs[paragraphs.findIndex((p) => p.text.startsWith(`BEGIN${c.n} `)) + 1].text]));
  };
  // rejected, each equation reads as Word read it in the run, where All Markup showed the deleted text
  const rejected = await caseParagraphs(await saved('57-fields-equations-word365-rejected.docx'));
  for (const n of [5, 6, 7, 8, 9, 10]) assert.equal(rejected[n], unvis(word[n].region.paragraphs[1]), `case ${n}`);
  // Word's form of an equation deleted whole reads as nothing in Simple Markup's rule, its structure with its runs, as
  // the one w:del form does
  const run = await caseParagraphs(await saved('57-fields-equations-word365.docx'));
  assert.deepEqual([5, 6, 7, 8, 9, 10].map((n) => run[n]), ['', '', '', '', '', 'keep10a  keep10b']);
  const ours = await caseParagraphs(await check57Bytes());
  assert.deepEqual([ours[5], ours[6]], ['', '']);
  // the listing: one change a case, as Word's run lists them, its text docx4j's reader's
  const changes = await (await WordprocessingMLPackage.load(await saved('57-fields-equations-word365.docx'))).getAllTrackedChanges();
  assert.deepEqual(changes.map((c) => [c.type, c.text, c.author]), [
    ['Deleted', 'E=mc2\r', 'docx4j'], ['Deleted', 'keep6a L6 E=mc2 R6 keep6b\r', 'docx4j'],
    ['Deleted', 'E=mc2\r', 'Author A'], ['Deleted', 'keep8a L8 E=mc2 R8 keep8b\r', 'Author A'],
    ['Deleted', 'E=mc2', 'Author A'], ['Deleted', 'L10 E=mc2 R10', 'Author A'],
  ]);
  assert.deepEqual([5, 6, 7, 8, 9, 10].map((n) => word[n].region.trackedChanges.length), [1, 1, 1, 1, 1, 1]);
});

/** A package of this document part's paragraphs (`documentOf`), its body. */
async function bodyOf(paragraphs) {
  const pkg = await WordprocessingMLPackage.load(await packageOf(documentOf(paragraphs)));
  return { pkg, body: await pkg.getBody() };
}

test('CR-002 section 62.2 D1: a w:fldSimple a break splits is made a complex field, its lock, dirty flag and data on the begin, its runs\' look on the field characters', async () => {
  const bold = '<w:rPr><w:b/></w:rPr>';
  const field = `<w:p><w:r><w:t xml:space="preserve">a </w:t></w:r><w:fldSimple w:instr="DATE" w:dirty="true" w:fldLock="true"><w:r>${bold}<w:t>one two</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> b</w:t></w:r></w:p>`;
  for (const write of [(p) => p.getRange().search('one')[0].insertParagraph('NEW', 'After'), (p) => p.getRange().search('one')[0].insertText('\n', 'After')]) {
    const { pkg, body } = await bodyOf([field]);
    const p = body.paragraphs.find((q) => q.text.includes('one'));
    write(p);
    const paragraphs = equationOrFieldParagraphs(documentXmlOf(await pkg.save()));
    // the first half's new mark aside (insertParagraph's in the run's formatting, a line feed's in none, as before)
    assert.equal(paragraphs[0].replace(/<w:pPr>.*?<\/w:pPr>/, ''), `<w:p><w:r><w:t xml:space="preserve">a </w:t></w:r><w:r>${bold}<w:fldChar w:dirty="true" w:fldCharType="begin" w:fldLock="true"/></w:r>`
      + `<w:r>${bold}<w:instrText>DATE</w:instrText></w:r><w:r>${bold}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${bold}<w:t>one</w:t></w:r></w:p>`);
    assert.equal(paragraphs[paragraphs.length - 1], `<w:p><w:r>${bold}<w:t xml:space="preserve"> two</w:t></w:r><w:r>${bold}<w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> b</w:t></w:r></w:p>`);
    assert.ok(!documentXmlOf(await pkg.save()).includes('fldSimple'));
  }
});

/** The paragraphs of a document part between the check's title and its last line, without namespace declarations. */
const equationOrFieldParagraphs = (xml) => (xml.match(/<w:p\b(?:(?!<w:p\b)[\s\S])*?<\/w:p>/g) ?? []).slice(1, -1)
  .map((p) => p.replace(/ xmlns:\w+="[^"]*"| w14:\w+="[^"]*"/g, ''));

test('CR-002 section 62.2 D2: a tracked delete of an equation holding a revision of its own puts it whole in one w:del, as before; a tracked Range replacement goes ahead, a Paragraph one is refused', async () => {
  const revised = `<w:p><w:r><w:t xml:space="preserve">x </w:t></w:r><m:oMath><m:r><w:ins w:id="90" w:author="Other" w:date="2026-10-01T00:00:00Z"><m:t>q</m:t></w:ins></m:r></m:oMath><w:r><w:t xml:space="preserve"> y</w:t></w:r></w:p>`;
  {
    const { pkg, body } = await bodyOf([revised]);
    pkg.changeTrackingMode = 'TrackAll';
    body.paragraphs[1].delete();
    assert.match(equationOrFieldParagraphs(documentXmlOf(await pkg.save()))[0], /<w:del [^>]*><w:r><w:delText xml:space="preserve">x <\/w:delText><\/w:r><m:oMath>/);
  }
  {
    const { pkg, body } = await bodyOf([sourceParagraph(10, 'inline-equation')]);
    pkg.changeTrackingMode = 'TrackAll';
    const p = body.paragraphs[1];
    const before = documentXmlOf(await pkg.save());
    // Paragraph.insertText's Replace puts its text after the deletion, where Word's beside a deleted equation isn't measured
    assert.throws(() => p.insertText('new', 'Replace'), (e) => e instanceof NotSupportedError && /tracked replacement of an equation/.test(e.message));
    assert.equal(documentXmlOf(await pkg.save()), before);
    // a Range's puts its text before the deletion, as Word's does (check 52 cases 39 to 48), the equation deleted in Word's form
    p.getRange().search('L10')[0].expandTo(p.getRange().search('R10')[0]).insertText('new', 'Replace');
    assert.equal(p.text, 'keep10a new keep10b');
    assert.match(equationOrFieldParagraphs(documentXmlOf(await pkg.save()))[0], /<w:ins [^>]*><w:r><w:t>new<\/w:t><\/w:r><\/w:ins><w:del [^>]*><w:r><w:delText xml:space="preserve">L10 <\/w:delText><\/w:r><\/w:del><m:oMath><m:r><w:del /);
  }
});
