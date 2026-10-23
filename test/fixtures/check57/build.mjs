// The file for Word check 57 (CR-002 section 62: a paragraph break inside a field's result, and a tracked delete of
// a paragraph holding an equation; docx4j-python's review, items B2 and B11). Run from the repository root after
// `npm run build`:
//
//   node test/fixtures/check57/build.mjs
//
// writes 57-fields-equations.docx into this directory, which git ignores (Word reads the share's copy);
// `check57Bytes` gives a test the same parts. A region per case: a paragraph `BEGIN<n> <op> <shape>`, the case's
// paragraph, and a paragraph `END<n>`. The shapes:
// - `complex-field`: `keep<n>a `, a complex QUOTE field (begin, instruction, separate, its result `FA<n> FB<n>` in one
//   run, end), ` keep<n>b`;
// - `simple-field`: the same with a `w:fldSimple`;
// - `display-equation`: a paragraph holding only an `m:oMathPara`, E=mc² with a superscript structure;
// - `inline-equation`: `keep<n>a L<n> `, the same equation as an `m:oMath`, ` R<n> keep<n>b`.
// An `open` case's paragraph is the one this package's tracked `Paragraph.delete` wrote for its shape before CR-002
// section 62.3 (one `w:del` around the runs and the equation, the mark deleted), so Word opens it as written; it still
// writes it for an equation holding a revision of its own. The other cases' are as above, for the snippet to write.
// The parts are written as text, not marshalled.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The author and date of the open cases' revisions: this package's default author. */
export const AUTHOR = 'docx4j';
export const DATE = '2026-10-23T00:00:00Z';

/** Check 57's cases, in document order: { n, op, shape }. */
export function check57Cases() {
  const rows = [
    ['insertText-after', 'complex-field'], ['insertText-after', 'simple-field'],
    ['insertParagraph-after', 'complex-field'], ['insertParagraph-after', 'simple-field'],
    ['open', 'display-equation'], ['open', 'inline-equation'],
    ['paragraph-delete', 'display-equation'], ['paragraph-delete', 'inline-equation'],
    ['range-delete', 'display-equation'], ['range-delete', 'inline-equation'],
  ];
  return rows.map(([op, shape], i) => ({ n: i + 1, op, shape }));
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, op, shape }) => `BEGIN${n} ${op} ${shape}`;

const t = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const mathRun = (text) => `<m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>${text}</m:t></m:r>`;
/** E=mc², as Word writes an equation: math runs in Cambria Math, a structure's m:ctrlPr. */
const EQUATION = `<m:oMath>${mathRun('E=m')}<m:sSup><m:sSupPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr></m:ctrlPr></m:sSupPr>`
  + `<m:e>${mathRun('c')}</m:e><m:sup>${mathRun('2')}</m:sup></m:sSup></m:oMath>`;

/** A case's paragraph content, before anything is done to it. */
export function sourceContent(n, shape) {
  switch (shape) {
    case 'complex-field':
      return t(`keep${n}a `) + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        + `<w:r><w:instrText xml:space="preserve"> QUOTE "FA${n} FB${n}" </w:instrText></w:r>`
        + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + t(`FA${n} FB${n}`) + '<w:r><w:fldChar w:fldCharType="end"/></w:r>' + t(` keep${n}b`);
    case 'simple-field':
      return t(`keep${n}a `) + `<w:fldSimple w:instr=' QUOTE "FA${n} FB${n}" '>${t(`FA${n} FB${n}`)}</w:fldSimple>` + t(` keep${n}b`);
    case 'display-equation': return `<m:oMathPara>${EQUATION}</m:oMathPara>`;
    case 'inline-equation': return t(`keep${n}a L${n} `) + EQUATION + t(` R${n} keep${n}b`);
    default: throw new Error(`no shape ${shape}`);
  }
}

/** The paragraph this package's tracked Paragraph.delete wrote for a case's content before CR-002 section 62.3 (B11). */
function deletedParagraph(n, shape) {
  const attrs = (id) => `w:id="${id}" w:author="${AUTHOR}" w:date="${DATE}"`;
  const content = sourceContent(n, shape).replace(/<w:t xml:space="preserve">([^<]*)<\/w:t>/g, '<w:delText xml:space="preserve">$1</w:delText>');
  return `<w:p><w:pPr><w:rPr><w:del ${attrs(5700 + 2 * n)}/></w:rPr></w:pPr><w:del ${attrs(5701 + 2 * n)}>${content}</w:del></w:p>`;
}

/** A case's paragraph before anything is done to it. */
export const sourceParagraph = (n, shape) => para(sourceContent(n, shape));

/** A case's paragraph as the file holds it. */
const caseParagraph = ({ n, op, shape }) => (op === 'open' ? deletedParagraph(n, shape) : sourceParagraph(n, shape));

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** A document part of these paragraphs, between the check's title and its last line. */
export function documentOf(paragraphs) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:m="${M}" xmlns:r="${REL}"><w:body>`
    + para(t('Check 57. Each case is a region: its BEGIN paragraph, the case\'s paragraph, its END paragraph.'))
    + paragraphs.join('') + para(t('End of check 57.')) + `${sectPr}</w:body></w:document>`;
}

/** Check 57's file, in memory. */
export async function check57Bytes() {
  const regions = check57Cases().flatMap((c) => [para(t(beginText(c))), caseParagraph(c), para(t(`END${c.n}`))]);
  return packageOf(documentOf(regions));
}

/** A created package with this document part. */
export async function packageOf(documentXml) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  if (!entries['word/document.xml']) throw new Error('no word/document.xml in the created package');
  entries['word/document.xml'] = strToU8(documentXml);
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

/** The document part of a file, as text. */
export const documentXmlOf = (bytes) => strFromU8(unzipSync(bytes)['word/document.xml']);

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check57Bytes();
  await writeFile(join(here, '57-fields-equations.docx'), bytes);
  console.log(`wrote 57-fields-equations.docx (${bytes.length} bytes, ${check57Cases().length} cases)`);
}
