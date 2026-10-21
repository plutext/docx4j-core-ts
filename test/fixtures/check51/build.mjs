// The file for Word check 51 (CR-012 section 17.7): what CR-012 phase 4 left unmeasured. (1) The joiner of a
// cell's, a content control's and a comment's text when they hold several paragraphs; (2) a footnote's and an
// endnote's own mark in the note's text; (3) an equation's letters beyond check 48's `x`: the math styles
// (`m:sty`), the scripts (`m:scr`), Greek, digits and operators, normal text (`m:nor`) and the structures.
// Run from the repository root after `npm run build`:
//
//   node test/fixtures/check51/build.mjs
//
// writes 51-text.docx into this directory, which git ignores (Word reads the share's copy, as for check 48);
// `check51Bytes` gives the same bytes to a test. Each case is a paragraph "Case n: <what>." and then the item,
// between upper-case markers where it is inline. The parts are written as text, not marshalled.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
};
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const ROOT_NS = Object.entries(NS).map(([p, uri]) => ` xmlns:${p}="${uri}"`).join('');
const AUTHOR = 'w:author="Check 51" w:date="2026-10-21T00:00:00Z"';

const escape = (text) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const t = (text) => `<w:r><w:t xml:space="preserve">${escape(text)}</w:t></w:r>`;
const run = (inner, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}${inner}</w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const label = (n, what) => para(t(`Case ${n}: ${what}.`));
const testCase = (n, what, inner) => para(t(`Case ${n}: ${what}. `) + inner);

// --- math -------------------------------------------------------------------------------------------
/** An m:r: its text, and its m:rPr's style (`p`, `b`, `i`, `bi`), script and normal-text flag. */
const mr = (text, { sty, scr, nor } = {}) => {
  const props = (nor ? '<m:nor/>' : '') + (scr ? `<m:scr m:val="${scr}"/>` : '') + (sty ? `<m:sty m:val="${sty}"/>` : '');
  return `<m:r>${props === '' ? '' : `<m:rPr>${props}</m:rPr>`}<m:t xml:space="preserve">${escape(text)}</m:t></m:r>`;
};
const math = (inner) => `<m:oMath>${inner}</m:oMath>`;
/** A paragraph of equations, each between markers MnA and MnB: what Word's text gives for each. */
const mathCase = (n, what, equations) => testCase(n, what, equations.map(([marker, inner]) => t(` ${marker}A `) + math(inner) + t(` ${marker}B`)).join(''));

const cases = [];

// (1) the joiners
const tc = (paragraphs, width = 3000) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${paragraphs.map((p) => para(t(p))).join('')}</w:tc>`;
const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((b) => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
cases.push(
  label(1, 'a table: its first cell holds three paragraphs, CELLP1 to CELLP3, its second one, OTHER'),
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}<w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>`
    + `<w:tr>${tc(['CELLP1', 'CELLP2', 'CELLP3'])}${tc(['OTHER'])}</w:tr></w:tbl>`,
  para(t('AFTERTABLE')),
  label(2, 'a block content control (rich text, titled Check 51 block) of three paragraphs, CCP1 to CCP3'),
  '<w:sdt><w:sdtPr><w:alias w:val="Check 51 block"/><w:tag w:val="block"/><w:id w:val="51002"/></w:sdtPr><w:sdtContent>'
    + para(t('CCP1')) + para(t('CCP2')) + para(t('CCP3')) + '</w:sdtContent></w:sdt>',
  para(t('AFTERCONTROL')),
  testCase(3, 'an inline content control (titled Check 51 inline)',
    t('INLA ') + '<w:sdt><w:sdtPr><w:alias w:val="Check 51 inline"/><w:tag w:val="inline"/><w:id w:val="51003"/></w:sdtPr><w:sdtContent>'
    + t('INLINECC') + '</w:sdtContent></w:sdt>' + t(' INLB')),
  testCase(4, 'a comment of three paragraphs, COMP1 to COMP3, on the middle word',
    t('COMA ') + '<w:commentRangeStart w:id="0"/>' + t('COMMENTED') + '<w:commentRangeEnd w:id="0"/>' + run('<w:commentReference w:id="0"/>') + t(' COMB')),
);

// (2) the notes' own marks
const superscript = '<w:vertAlign w:val="superscript"/>';
cases.push(testCase(5, 'a footnote of two paragraphs (NOTEP1, NOTEP2), then an endnote of one (ENDP1)',
  t('FOOTA') + run('<w:footnoteReference w:id="1"/>', superscript) + t(' FOOTB ENDA') + run('<w:endnoteReference w:id="1"/>', superscript) + t(' ENDB')));

// (3) the letters
cases.push(
  mathCase(6, 'equations in the default style: Latin letters; digits and operators', [
    ['M6', mr('abhxyzABHXYZ')], ['N6', mr('0123+-=<>()')],
  ]),
  mathCase(7, 'equations in the default style: small Greek, capital Greek, the Greek variants and nabla', [
    ['M7', mr('αβγδεθλμπσςφω')], ['N7', mr('ΓΔΘΛΠΣΦΨΩ')], ['P7', mr('ϑϕϖϵϱϰ∂∇')],
  ]),
  mathCase(8, 'equations with m:sty: plain, bold, italic, bold italic', [
    ['M8', mr('sinx2Γ', { sty: 'p' })], ['N8', mr('bx2Γ', { sty: 'b' })], ['P8', mr('ix2Γ', { sty: 'i' })], ['Q8', mr('bix2Γ', { sty: 'bi' })],
  ]),
  mathCase(9, 'equations with m:scr: script, fraktur, double-struck, sans-serif, monospace', [
    ['M9', mr('BELeagz1', { scr: 'script' })], ['N9', mr('CHIRZgz1', { scr: 'fraktur' })], ['P9', mr('CNRZz1', { scr: 'double-struck' })],
    ['Q9', mr('ab1α', { scr: 'sans-serif' })], ['R9', mr('az1', { scr: 'monospace' })],
  ]),
  mathCase(10, 'equations with a bold script: bold script, bold fraktur, bold sans-serif, bold italic sans-serif', [
    ['M10', mr('Ba', { scr: 'script', sty: 'b' })], ['N10', mr('Ca', { scr: 'fraktur', sty: 'b' })],
    ['P10', mr('a1α', { scr: 'sans-serif', sty: 'b' })], ['Q10', mr('aα', { scr: 'sans-serif', sty: 'bi' })],
  ]),
  mathCase(11, 'an equation with m:nor (normal text), then one mixing a normal run and a math run', [
    ['M11', mr('nx', { nor: true })], ['N11', mr('nx', { nor: true }) + mr('y')],
  ]),
  mathCase(12, 'equation structures: a fraction a over b, x squared, the square root of x, a sum, x in parentheses, sin x', [
    ['F12', '<m:f><m:num>' + mr('a') + '</m:num><m:den>' + mr('b') + '</m:den></m:f>'],
    ['S12', '<m:sSup><m:e>' + mr('x') + '</m:e><m:sup>' + mr('2') + '</m:sup></m:sSup>'],
    ['R12', '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>' + mr('x') + '</m:e></m:rad>'],
    ['N12', '<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>' + mr('i=1') + '</m:sub><m:sup>' + mr('n') + '</m:sup><m:e>' + mr('x') + '</m:e></m:nary>'],
    ['D12', '<m:d><m:e>' + mr('x') + '</m:e></m:d>'],
    ['U12', '<m:func><m:fName>' + mr('sin', { sty: 'p' }) + '</m:fName><m:e>' + mr('x') + '</m:e></m:func>'],
  ]),
  label(13, 'the next paragraph is a display equation (m:oMathPara) of two equations, x=1 and y=2'),
  para('<m:oMathPara>' + math(mr('x=1')) + math(mr('y=2')) + '</m:oMathPara>'),
);

/** Word writes a separator and a continuation separator first in each notes part. */
const separators = (kind) => `<w:${kind} w:type="separator" w:id="-1">${para(run('<w:separator/>'))}</w:${kind}>`
  + `<w:${kind} w:type="continuationSeparator" w:id="0">${para(run('<w:continuationSeparator/>'))}</w:${kind}>`;
const notes = (kind, ref, paragraphs) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind}s${ROOT_NS}>`
  + separators(kind) + `<w:${kind} w:id="1">${paragraphs.map((text, i) => para((i === 0 ? run(`<w:${ref}/>`, superscript) : '') + t(i === 0 ? ` ${text}` : text))).join('')}</w:${kind}></w:${kind}s>`;
const comments = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments${ROOT_NS}>`
  + `<w:comment w:id="0" ${AUTHOR} w:initials="C51">${para(run('<w:annotationRef/>') + t('COMP1')) + para(t('COMP2')) + para(t('COMP3'))}</w:comment></w:comments>`;

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document${ROOT_NS}><w:body>`
  + [para(t('Check 51. Each case is a paragraph, then its item.')), ...cases, para(t('End of check 51.'))].join('') + `${sectPr}</w:body></w:document>`;

/** Check 51's file, in memory: what a test reads beside Word's answers. */
export async function check51Bytes() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };
  put('word/document.xml', documentXml);
  put('word/footnotes.xml', notes('footnote', 'footnoteRef', ['NOTEP1', 'NOTEP2']));
  put('word/endnotes.xml', notes('endnote', 'endnoteRef', ['ENDP1']));
  put('word/comments.xml', comments);
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck51Footnotes" Type="${REL}/footnotes" Target="footnotes.xml"/>`
    + `<Relationship Id="rIdCheck51Endnotes" Type="${REL}/endnotes" Target="endnotes.xml"/>`
    + `<Relationship Id="rIdCheck51Comments" Type="${REL}/comments" Target="comments.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    `<Override PartName="/word/footnotes.xml" ContentType="${CT}.footnotes+xml"/>`
    + `<Override PartName="/word/endnotes.xml" ContentType="${CT}.endnotes+xml"/>`
    + `<Override PartName="/word/comments.xml" ContentType="${CT}.comments+xml"/></Types>`));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check51Bytes();
  await writeFile(join(here, '51-text.docx'), bytes);
  console.log(`wrote 51-text.docx (${bytes.length} bytes)`);
}
