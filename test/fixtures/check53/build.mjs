// The file for Word check 53 (CR-014 phase 1's choices where Word was not measured; CR-014 section 9): read only.
// Run from the repository root after `npm run build`:
//
//   node test/fixtures/check53/build.mjs
//
// writes 53-reads.docx into this directory, which git ignores (it holds tracked changes, for which the weekly
// RevisionsHarness would cut a golden; Word reads the share's copy); `check53Bytes` gives the same bytes to a test.
// The parts are written as text, not marshalled.
//
// Each case is a region: a paragraph `BEGIN<n> <op> <args>`, its content, a paragraph `END<n>`. The arguments are
// lower case in the marker and upper case in the content, so a whole-word, case-sensitive search finds them in the
// content only. The ops (check53-script-lab.js reads each):
//   range x y        the range from word x's start to word y's end (expandTo): every read a Range has
//   para x           the paragraph holding x: its style, its whole range's, and the tracked changes of both
//   control tag      the content control tagged tag: its text and its ranges
//   table x          the table holding x: its ranges
//   cellbody x       the body of the cell holding x: its ranges
//   changes          the region's tracked changes, and each paragraph's
//   stories x        x's range expandTo the first footnote's whole range (two stories)
//   body             the main body's ranges
//   placeholder tag  a control showing its placeholder text: its text and its paragraph's (CR-012)
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
const AUTHOR = 'w:author="Check 53" w:date="2026-10-22T00:00:00Z"';

const run = (inner, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}${inner}</w:r>`;
const t = (text, rPr = '') => run(`<w:t xml:space="preserve">${text}</w:t>`, rPr);
const para = (pPr, inner) => `<w:p>${pPr === '' ? '' : `<w:pPr>${pPr}</w:pPr>`}${inner}</w:p>`;
const words = (rPr, ...ws) => ws.map((w, i) => t(i < ws.length - 1 ? `${w} ` : w, rPr)).join('');
let revisionId = 9000;
const del = (inner) => `<w:del w:id="${revisionId++}" ${AUTHOR}>${inner}</w:del>`;
const ins = (inner) => `<w:ins w:id="${revisionId++}" ${AUTHOR}>${inner}</w:ins>`;
const delRun = (text, rPr = '') => run(`<w:delText xml:space="preserve">${text}</w:delText>`, rPr);
const markDel = () => `<w:rPr>${del('')}</w:rPr>`;
const markIns = () => `<w:rPr>${ins('')}</w:rPr>`;

const HEADING = '<w:pStyle w:val="Heading1"/>';
const HEADING2 = '<w:pStyle w:val="Heading2"/>';
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
/** A table of cells each holding one paragraph of two words, `<prefix>R<r>C<c>A <prefix>R<r>C<c>B`. */
const plainTable = (prefix, rows, cols) => tbl(cols, ...Array.from({ length: rows }, (_, r) => tr(...Array.from({ length: cols },
  (_, c) => tc(para('', words('', `${prefix}R${r + 1}C${c + 1}A`, `${prefix}R${r + 1}C${c + 1}B`)))))));
const nested = (token) => tbl(1, tr(tc(para('', t(token)), 1200)));
const sdtPr = (tag, extra = '') => `<w:sdtPr><w:alias w:val="Check 53 ${tag}"/><w:tag w:val="${tag}"/><w:id w:val="${53000 + Number(tag.replace(/\D/g, ''))}"/>${extra}</w:sdtPr>`;
const sdt = (tag, content, extra = '') => `<w:sdt>${sdtPr(tag, extra)}<w:sdtContent>${content}</w:sdtContent></w:sdt>`;

const footnotes = [];
/** A footnote reference run; the note reads NOTE<n>. */
const note = (n) => { footnotes.push(n); return run(`<w:footnoteReference w:id="${footnotes.length}"/>`, '<w:rStyle w:val="FootnoteReference"/>'); };

/** The cases: [op and arguments, content]. Tokens carry the case's number, given as `n`. */
const CASES = [
  // --- style across paragraphs (CR-014 section 9.1: their shared style, else Normal) and Word's 'Other'
  [(n) => `range head${n}b sub${n}a`, (n) => [para(HEADING, words('', `HEAD${n}A`, `HEAD${n}B`)), para(HEADING2, words('', `SUB${n}A`, `SUB${n}B`))]],
  [(n) => `range body${n}b plain${n}a`, (n) => [para(CENTRED, words(ITALIC, `BODY${n}A`, `BODY${n}B`)), para('', words('', `PLAIN${n}A`, `PLAIN${n}B`))]],
  [(n) => `range head${n}b head${n}c`, (n) => [para(HEADING, words('', `HEAD${n}A`, `HEAD${n}B`)), para(HEADING, words('', `HEAD${n}C`, `HEAD${n}D`))]],
  [(n) => `range head${n}a head${n}b`, (n) => [para(HEADING, words('', `HEAD${n}A`, `HEAD${n}B`) + note(n) + t(` HEAD${n}C`))]],
  [(n) => `range head${n}b head${n}c`, (n) => [para(HEADING, words('', `HEAD${n}A`, `HEAD${n}B`) + note(n) + t(` HEAD${n}C`))]],
  [(n) => `para head${n}a`, (n) => [para(HEADING, words('', `HEAD${n}A`, `HEAD${n}B`) + note(n) + t(` HEAD${n}C`))]],
  [(n) => `range emph${n}b emph${n}b`, (n) => [para('', t(`PLAIN${n}A `) + t(`EMPH${n}B`, '<w:rStyle w:val="Emphasis"/>') + t(` PLAIN${n}C`))]],
  [(n) => `range plain${n}a emph${n}b`, (n) => [para('', t(`PLAIN${n}A `) + t(`EMPH${n}B`, '<w:rStyle w:val="Emphasis"/>') + t(` PLAIN${n}C`))]],
  [(n) => `range plain${n}a plain${n}a`, (n) => [para('', t(`PLAIN${n}A `) + t(`EMPH${n}B`, '<w:rStyle w:val="Emphasis"/>') + t(` PLAIN${n}C`))]],
  // a paragraph after a nested table in a cell (check 52's TC2 read 'Other'): a list item, a plain one, a heading
  [(n) => `para list${n}a`, (n) => [tbl(1, tr(tc(nested(`N${n}A`) + para(list(n), words('', `LIST${n}A`, `LIST${n}B`)))))]],
  [(n) => `para plain${n}a`, (n) => [tbl(1, tr(tc(nested(`N${n}A`) + para('', words('', `PLAIN${n}A`, `PLAIN${n}B`)))))]],
  [(n) => `para head${n}a`, (n) => [tbl(1, tr(tc(nested(`N${n}A`) + para(HEADING, words('', `HEAD${n}A`, `HEAD${n}B`)))))]],
  [(n) => `para list${n}a`, (n) => [tbl(1, tr(tc(para(list(n), words('', `LIST${n}A`, `LIST${n}B`)))))]],
  [(n) => `para plain${n}a`, (n) => [tbl(1, tr(tc(para('', words('', `PLAIN${n}A`, `PLAIN${n}B`)) + nested(`N${n}A`) + para('', t(`LAST${n}A`)))))]],
  // --- content controls' ranges (section 9.2: a row control's and one ending in a table end \r\n)
  [(n) => `control c53-${n}`, (n) => [`<w:tbl>${tblPr}${grid(2)}${sdt(`c53-${n}`, tr(tc(para('', t(`ROW${n}A`))), tc(para('', t(`ROW${n}B`)))))}${tr(tc(para('', t(`ROW${n}C`))), tc(para('', t(`ROW${n}D`))))}</w:tbl>`]],
  [(n) => `control c53-${n}`, (n) => [sdt(`c53-${n}`, para('', words('', `CTRL${n}A`, `CTRL${n}B`)) + plainTable(`T${n}`, 1, 2)), para('', t(`AFTER${n}A`))]],
  [(n) => `control c53-${n}`, (n) => [sdt(`c53-${n}`, plainTable(`T${n}`, 1, 2)), para('', t(`AFTER${n}A`))]],
  [(n) => `control c53-${n}`, (n) => [`<w:tbl>${tblPr}${grid(2)}<w:tr>${sdt(`c53-${n}`, tc(para('', t(`CELL${n}A`)) + para('', t(`CELL${n}B`))))}${tc(para('', t(`CELL${n}C`)))}</w:tr></w:tbl>`]],
  // --- the controls a range lists (section 9.1: those whose content it holds whole)
  [(n) => `range plain${n}b ctrl${n}a`, (n) => [para('', words('', `PLAIN${n}A`, `PLAIN${n}B`)), sdt(`c53-${n}`, para('', words('', `CTRL${n}A`, `CTRL${n}B`)) + para('', t(`CTRL${n}C`)))]],
  [(n) => `range ctrl${n}a ctrl${n}b`, (n) => [sdt(`c53-${n}`, para('', words('', `CTRL${n}A`, `CTRL${n}B`)))]],
  [(n) => `range plain${n}a plain${n}c`, (n) => [para('', t(`PLAIN${n}A `) + sdt(`c53-${n}`, t(`INL${n}B`)) + t(` PLAIN${n}C`))]],
  // --- a table's ranges, and the tables a range lists (section 9.1, 9.2)
  [(n) => `table t${n}r1c1a`, (n) => [plainTable(`T${n}`, 2, 2), para('', t(`AFTER${n}A`))]],
  [(n) => `cellbody t${n}a`, (n) => [tbl(2, tr(tc(para('', t(`T${n}A`)) + nested(`N${n}A`) + para('', t(`T${n}B`))), tc(para('', t(`T${n}C`)))))]],
  [(n) => `range n${n}a n${n}a`, (n) => [tbl(1, tr(tc(para('', t(`T${n}A`)) + nested(`N${n}A`) + para('', t(`T${n}B`)))))]],
  // --- tracked changes' ranges (section 9.2), in each view: through a mark, a lone mark, a whole paragraph
  [() => 'changes', (n) => [para(HEADING + markDel(), t(`HEAD${n}A `) + del(delRun(`HEAD${n}B`))), para(CENTRED, del(delRun(`BODY${n}A `, ITALIC)) + t(`BODY${n}B`, ITALIC))]],
  [() => 'changes', (n) => [para(list(n) + markIns(), t(`LIST${n}A`) + ins(t(` LIST${n}B`))), para('', ins(t(`PLAIN${n}A `)) + t(`PLAIN${n}B`))]],
  [() => 'changes', (n) => [para(markIns(), t(`ONE${n}A`)), para('', t(`ONE${n}B`))]],
  [() => 'changes', (n) => [para(markDel(), t(`TWO${n}A`)), para('', t(`TWO${n}B`))]],
  [() => 'changes', (n) => [para('', t(`PLAIN${n}A`)), para(markIns(), ins(t(`NEW${n}A NEW${n}B`))), para('', t(`PLAIN${n}C`))]],
  [() => 'changes', (n) => [para('', t(`PLAIN${n}A`)), para(markDel(), del(delRun(`OLD${n}A OLD${n}B`))), para('', t(`PLAIN${n}C`))]],
  // --- font where runs differ in colour and highlight (MixedFont)
  [(n) => `range red${n}a blue${n}a`, (n) => [para('', t(`RED${n}A`, '<w:color w:val="FF0000"/>') + t(' ') + t(`BLUE${n}A`, '<w:color w:val="0000FF"/><w:highlight w:val="yellow"/>'))]],
  // --- two stories, the body's ranges
  [(n) => `stories head${n}a`, (n) => [para(HEADING, words('', `HEAD${n}A`, `HEAD${n}B`) + note(n))]],
  [() => 'body', () => [para('', t('BODYCASE'))]],
  // --- a control showing its placeholder (CR-012: check 52's emptied control read as no text)
  [(n) => `placeholder c53-${n}`, (n) => [sdt(`c53-${n}`, para('', t('Click or tap here to enter text.', '<w:rStyle w:val="PlaceholderText"/>')), '<w:showingPlcHdr/>')]],
  [(n) => `placeholder c53-${n}`, (n) => [sdt(`c53-${n}`, para('', t('     ', '<w:b/>')), '<w:showingPlcHdr/>')]],
];

/** Check 53's cases, in document order: { n, spec, content }. */
export function check53Cases() {
  footnotes.length = 0;
  revisionId = 9000;
  return CASES.map(([spec, content], i) => ({ n: i + 1, spec: spec(i + 1), content: content(i + 1) }));
}

const STYLES = '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>'
  + '<w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
  + '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
  + '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>';

const numbering = (cases) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering${ROOT_NS}>`
  + '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/>'
  + '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
  + cases.filter((c) => c.content.join('').includes(`<w:numId w:val="${c.n}"/>`))
    .map((c) => `<w:num w:numId="${c.n}"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`).join('')
  + '</w:numbering>';
const separators = '<w:footnote w:type="separator" w:id="-1">' + para('', run('<w:separator/>')) + '</w:footnote>'
  + '<w:footnote w:type="continuationSeparator" w:id="0">' + para('', run('<w:continuationSeparator/>')) + '</w:footnote>';
const footnotesXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes${ROOT_NS}>${separators}`
  + footnotes.map((n, i) => `<w:footnote w:id="${i + 1}">${para('<w:pStyle w:val="FootnoteText"/>',
    run('<w:footnoteRef/>', '<w:rStyle w:val="FootnoteReference"/>') + t(` NOTE${n}`))}</w:footnote>`).join('') + '</w:footnotes>';
const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 53's file, in memory. */
export async function check53Bytes() {
  const cases = check53Cases();
  const regions = cases.flatMap((c) => [para('', t(`BEGIN${c.n} ${c.spec}`)), ...c.content, para('', t(`END${c.n}`))]);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document${ROOT_NS}><w:body>`
    + para('', t('Check 53. Each case is a region: its BEGIN paragraph, its content, its END paragraph. Read only.')) + regions.join('')
    + para('', t('End of check 53.')) + `${sectPr}</w:body></w:document>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };
  put('word/document.xml', documentXml);
  put('word/numbering.xml', numbering(cases));
  put('word/footnotes.xml', footnotesXml());
  edit('word/styles.xml', (xml) => xml.replace('</w:styles>', `${STYLES}</w:styles>`));
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck53Numbering" Type="${REL}/numbering" Target="numbering.xml"/>`
    + `<Relationship Id="rIdCheck53Footnotes" Type="${REL}/footnotes" Target="footnotes.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    `<Override PartName="/word/numbering.xml" ContentType="${CT}.numbering+xml"/>`
    + `<Override PartName="/word/footnotes.xml" ContentType="${CT}.footnotes+xml"/></Types>`));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check53Bytes();
  await writeFile(join(here, '53-reads.docx'), bytes);
  console.log(`wrote 53-reads.docx (${bytes.length} bytes, ${check53Cases().length} cases)`);
}
