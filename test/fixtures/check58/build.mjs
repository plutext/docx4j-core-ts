// The file for Word check 58 (2026-10-23): what CR-002 section 58, CR-012 sections 18 and 19 and CR-002 section 32's
// item 1 left to measure. Run from the repository root after `npm run build`:
//
//   node test/fixtures/check58/build.mjs
//
// writes 58-measures.docx into this directory, which git ignores (it holds tracked changes, for which the weekly
// RevisionsHarness would cut a golden; Word reads the share's copy); `check58Bytes` gives the same bytes to a test.
// The parts are written as text, not marshalled.
//
// Each case is a region: a paragraph `BEGIN<n> <op> <args>`, its content, a paragraph `END<n>`. The arguments are
// lower case in the marker and upper case in the content, so a whole-word, case-sensitive search finds them in the
// content only. The ops (check58-script-lab.js reads, or in its edit mode writes, each):
//   changes            the region's tracked changes, and each paragraph's, its whole range's and its content range's
//   range x y          the range from word x's start to word y's end (expandTo): its text and paragraphs
//   cell               the region's first table's first cell: its paragraphs' text and ranges, its body's, its value,
//                      and the table's values and whole range
//   placeholder tag    a control showing its placeholder: its text, its paragraph's, its ranges'
//   search id          the region searched with the snippet's queries `id` (QUERIES), each hit's text
//   accept i, reject i edit mode: the first tracked change paragraph i of the region lists (BEGIN is 0), accepted or
//                      rejected; then the region's paragraphs and changes. Read mode: the region's changes, as `changes`
//   fill tag how       edit mode: the placeholder control tagged tag takes `insertText('FILL<n>', how)`; then its text
//                      and its paragraph's. Read mode: as `placeholder`
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const ROOT_NS = ` xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const author = (name = 'Check 58') => `w:author="${name}" w:date="2026-10-23T00:00:00Z"`;

const run = (inner, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}${inner}</w:r>`;
const t = (text, rPr = '') => run(`<w:t xml:space="preserve">${text}</w:t>`, rPr);
const para = (pPr, inner) => `<w:p>${pPr === '' ? '' : `<w:pPr>${pPr}</w:pPr>`}${inner}</w:p>`;
const words = (rPr, ...ws) => ws.map((w, i) => t(i < ws.length - 1 ? `${w} ` : w, rPr)).join('');
let revisionId = 9000;
const del = (inner, who) => `<w:del w:id="${revisionId++}" ${author(who)}>${inner}</w:del>`;
const ins = (inner, who) => `<w:ins w:id="${revisionId++}" ${author(who)}>${inner}</w:ins>`;
const delRun = (text, rPr = '') => run(`<w:delText xml:space="preserve">${text}</w:delText>`, rPr);
const markDel = (who) => `<w:rPr>${del('', who)}</w:rPr>`;
const markIns = (who) => `<w:rPr>${ins('', who)}</w:rPr>`;
/** A paragraph wholly deleted: its text and its mark. */
const deletedParagraph = (text, who) => para(markDel(who), del(delRun(text), who));

const HEADING = '<w:pStyle w:val="Heading1"/>';
const CENTRED = '<w:jc w:val="center"/>';
const ITALIC = '<w:i/>';
const list = (numId) => `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`;

const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((b) => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
const tblPr = `<w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}<w:tblLook w:val="04A0"/></w:tblPr>`;
const grid = (n, w = 2400) => `<w:tblGrid>${`<w:gridCol w:w="${w}"/>`.repeat(n)}</w:tblGrid>`;
const tc = (inner, w = 2400) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const tr = (...cells) => `<w:tr>${cells.join('')}</w:tr>`;
const tbl = (cols, ...rows) => `<w:tbl>${tblPr}${grid(cols)}${rows.join('')}</w:tbl>`;

const PLACEHOLDER = 'Click or tap here to enter text.';
const PLACEHOLDER_RUN = t(PLACEHOLDER, '<w:rStyle w:val="PlaceholderText"/>');
let controlId = 58000;
const sdtPr = (tag, extra = '') => `<w:sdtPr><w:alias w:val="Check 58 ${tag}"/><w:tag w:val="${tag}"/><w:id w:val="${controlId++}"/>${extra}</w:sdtPr>`;
const sdt = (tag, content, extra = '') => `<w:sdt>${sdtPr(tag, extra)}<w:sdtContent>${content}</w:sdtContent></w:sdt>`;
/** A block control showing Word's own placeholder text. */
const blockPlaceholder = (tag) => sdt(tag, para('', PLACEHOLDER_RUN), '<w:showingPlcHdr/>');
/** A paragraph holding a run-level control showing the placeholder text between two words. */
const runPlaceholder = (tag, n) => para('', t(`PLAIN${n}A `) + sdt(tag, PLACEHOLDER_RUN, '<w:showingPlcHdr/>') + t(` PLAIN${n}B`));

/** check 53's case 25: a deletion from a heading's second word through its mark into the next paragraph's first word. */
const throughMark = (n) => [para(HEADING + markDel(), t(`HEAD${n}A `) + del(delRun(`HEAD${n}B`))), para(CENTRED, del(delRun(`BODY${n}A `, ITALIC)) + t(`BODY${n}B`, ITALIC))];
/** check 53's case 26: an insertion from a list item's second word through its mark into the next paragraph. */
const insertedThroughMark = (n) => [para(list(n) + markIns(), t(`LIST${n}A`) + ins(t(` LIST${n}B`))), para('', ins(t(`PLAIN${n}A `)) + t(`PLAIN${n}B`))];

/** The cases: [op and arguments, content]. Tokens carry the case's number, given as `n`. */
const CASES = [
  // --- CR-002 section 58: a wholly deleted paragraph's own listing, again and in variants (check 53 case 30 listed none)
  [() => 'changes', (n) => [para('', t(`PLAIN${n}A`)), deletedParagraph(`OLD${n}A OLD${n}B`), para('', t(`PLAIN${n}C`))]],
  [() => 'changes', (n) => [para('', t(`PLAIN${n}A`)), deletedParagraph(`OLD${n}A OLD${n}B`), deletedParagraph(`OLD${n}C OLD${n}D`), para('', t(`PLAIN${n}C`))]],
  [() => 'changes', (n) => [para('', t(`PLAIN${n}A`)), deletedParagraph(`OLD${n}A OLD${n}B`), deletedParagraph(`OLD${n}C OLD${n}D`, 'Check 58 other'), para('', t(`PLAIN${n}C`))]],
  [() => 'changes', (n) => [para('', t(`PLAIN${n}A`)), para('', del(delRun(`OLD${n}A OLD${n}B`))), para('', t(`PLAIN${n}C`))]],
  [() => 'changes', (n) => [para('', t(`PLAIN${n}A`)), para(markIns(), ins(t(`NEW${n}A NEW${n}B`))), para('', t(`PLAIN${n}C`))]],
  // --- CR-012 section 19: a range across a deleted mark, in Word's default view
  [(n) => `range head${n}a body${n}b`, throughMark],
  [(n) => `range two${n}a two${n}b`, (n) => [para(markDel(), t(`TWO${n}A`)), para('', t(`TWO${n}B`))]],
  [(n) => `range plain${n}a plain${n}c`, (n) => [para('', t(`PLAIN${n}A`)), deletedParagraph(`OLD${n}A OLD${n}B`), para('', t(`PLAIN${n}C`))]],
  // a deleted mark at a cell's last paragraph, which Word never writes: alone in its cell, and after another paragraph
  [() => 'cell', (n) => [tbl(2, tr(tc(para(markDel(), t(`CELL${n}A`))), tc(para('', t(`CELL${n}B`))))), para('', t(`AFTER${n}A`))]],
  [() => 'cell', (n) => [tbl(2, tr(tc(para('', t(`CELL${n}A`)) + para(markDel(), t(`CELL${n}B`))), tc(para('', t(`CELL${n}C`))))), para('', t(`AFTER${n}A`))]],
  // --- CR-012 section 18: a control showing its placeholder, run-level, and the placeholder's text in Word's search
  [(n) => `placeholder c58-${n}`, (n) => [runPlaceholder(`c58-${n}`, n)]],
  [() => 'search placeholder', (n) => [blockPlaceholder(`c58-${n}`), runPlaceholder(`c58-${n}r`, n), para('', t(PLACEHOLDER))]],
  // --- CR-002 section 32, item 1: Word's wildcards over a surrogate pair and a combining mark
  [() => 'search wildcards', () => [para('', t('x\u{1F600}y a\u{1F600}b كَتَبَ z'))]],
  // --- CR-002 section 58: a paragraph's part of a change, accepted or rejected (edit mode)
  [() => 'accept 1', throughMark],
  [() => 'accept 2', throughMark],
  [() => 'reject 1', throughMark],
  [() => 'reject 2', throughMark],
  [() => 'accept 2', insertedThroughMark],
  [() => 'reject 2', insertedThroughMark],
  // --- CR-012 section 18: writing into a control showing its placeholder (edit mode)
  [(n) => `fill c58-${n} replace`, (n) => [blockPlaceholder(`c58-${n}`)]],
  [(n) => `fill c58-${n} end`, (n) => [blockPlaceholder(`c58-${n}`)]],
  [(n) => `fill c58-${n} end`, (n) => [runPlaceholder(`c58-${n}`, n)]],
];

/** Check 58's cases, in document order: { n, spec, content }. */
export function check58Cases() {
  revisionId = 9000;
  controlId = 58000;
  return CASES.map(([spec, content], i) => ({ n: i + 1, spec: spec(i + 1), content: content(i + 1) }));
}

const STYLES = '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>'
  + '<w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>'
  + '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>';

const numbering = (cases) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering${ROOT_NS}>`
  + '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/>'
  + '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
  + cases.filter((c) => c.content.join('').includes(`<w:numId w:val="${c.n}"/>`))
    .map((c) => `<w:num w:numId="${c.n}"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`).join('')
  + '</w:numbering>';
const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 58's file, in memory. */
export async function check58Bytes() {
  const cases = check58Cases();
  const regions = cases.flatMap((c) => [para('', t(`BEGIN${c.n} ${c.spec}`)), ...c.content, para('', t(`END${c.n}`))]);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document${ROOT_NS}><w:body>`
    + para('', t('Check 58. Each case is a region: its BEGIN paragraph, its content, its END paragraph.')) + regions.join('')
    + para('', t('End of check 58.')) + `${sectPr}</w:body></w:document>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };
  put('word/document.xml', documentXml);
  put('word/numbering.xml', numbering(cases));
  edit('word/styles.xml', (xml) => xml.replace('</w:styles>', `${STYLES}</w:styles>`));
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck58Numbering" Type="${REL}/numbering" Target="numbering.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>', `<Override PartName="/word/numbering.xml" ContentType="${CT}.numbering+xml"/></Types>`));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check58Bytes();
  await writeFile(join(here, '58-measures.docx'), bytes);
  console.log(`wrote 58-measures.docx (${bytes.length} bytes, ${check58Cases().length} cases)`);
}
