// The file for Word check 52 (CR-014 phase 0): what Office JS does with a range across paragraphs. Run from
// the repository root after `npm run build`:
//
//   node test/fixtures/check52/build.mjs
//
// writes 52-ranges.docx into this directory, which git ignores (Word reads the share's copy, as for checks 48
// and 51); `check52Bytes` gives the same bytes to a test. The parts are written as text, not marshalled.
//
// Each case is a region of its own, so that one run of the snippet (check52-script-lab.js) measures every case
// on content no other case has touched: a paragraph `BEGIN<n> <kind> <op> <tracking> <range>`, the kind's
// paragraphs, and a paragraph `END<n>`. The range is written in lower case, so that the snippet's whole-word,
// case-sensitive search for a token (HEAD7B) finds it in the content and never in the marker:
//   from head7b to body7a      the start of HEAD7B's word to the end of BODY7A's (expandTo)
//   from head7b to-mark body7a  the start of HEAD7B's word to the end of BODY7A's paragraph, its mark included
//   whole body7a                BODY7A's paragraph, its mark included (getRange('Whole'))
//   control c52-7               the content control tagged c52-7, its content (getRange('Content'))
// The paragraphs are told apart by their formatting: a heading (Heading1, its runs plain, HEAD), a body
// paragraph (centred, its runs italic, BODY), a list item (List Paragraph, numbered, its runs underlined,
// LIST), a right-aligned paragraph (NEXT), a plain one (PLAIN), a table of two rows (T<n>R1C1A ...), a block
// content control of a heading, a body paragraph and a list item. test/README.md check 52 lists the kinds.
// Cases 86 to 95 are CR-002 section 55's (Word's table rule), on a cell's paragraph, which the range names:
//   cell list86a               the cell holding LIST86A (parentTableCell), its body
//   paragraph list87a          the paragraph holding LIST87A
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const ROOT_NS = Object.entries(NS).map(([p, uri]) => ` xmlns:${p}="${uri}"`).join('');
const AUTHOR = 'w:author="Check 52" w:date="2026-10-22T00:00:00Z"';

const run = (inner, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}${inner}</w:r>`;
const t = (text, rPr = '') => run(`<w:t xml:space="preserve">${text}</w:t>`, rPr);
const para = (pPr, inner) => `<w:p>${pPr === '' ? '' : `<w:pPr>${pPr}</w:pPr>`}${inner}</w:p>`;
/** Words in a run each, a space after all but the last. */
const words = (rPr, ...ws) => ws.map((w, i) => t(i < ws.length - 1 ? `${w} ` : w, rPr)).join('');

// --- the paragraphs ---------------------------------------------------------------------------------
const HEADING = '<w:pStyle w:val="Heading1"/>';
const ITALIC = '<w:i/>';
const UNDERLINE = '<w:u w:val="single"/>';
const listPPr = (n) => `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${n}"/></w:numPr>`;

const H = (n, inner = words('', `HEAD${n}A`, `HEAD${n}B`)) => para(HEADING, inner);
const B = (n) => para('<w:jc w:val="center"/>', words(ITALIC, `BODY${n}A`, `BODY${n}B`));
const L = (n) => para(listPPr(n), words(UNDERLINE, `LIST${n}A`, `LIST${n}B`));
const N = (n) => para('<w:jc w:val="right"/>', words('', `NEXT${n}A`, `NEXT${n}B`));
const P = (n) => para('', words('', `PLAIN${n}A`, `PLAIN${n}B`));

const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((b) => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
const cell = (n, r, c) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${para('', words('', `T${n}R${r}C${c}A`, `T${n}R${r}C${c}B`))}</w:tc>`;
const T = (n) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}<w:tblLook w:val="04A0"/></w:tblPr>`
  + '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
  + `<w:tr>${cell(n, 1, 1)}${cell(n, 1, 2)}</w:tr><w:tr>${cell(n, 2, 1)}${cell(n, 2, 2)}</w:tr></w:tbl>`;
const CC = (n) => `<w:sdt><w:sdtPr><w:alias w:val="Check 52 control ${n}"/><w:tag w:val="c52-${n}"/><w:id w:val="${52000 + n}"/></w:sdtPr>`
  + `<w:sdtContent>${H(n)}${B(n)}${L(n)}</w:sdtContent></w:sdt>`;

/**
 * A table of one row whose first cell holds `first` (blocks) and whose second holds T<n>R1C2A T<n>R1C2B: the cells
 * CR-002 section 55's D2 measures (Word's table rule, a cell's last block a paragraph).
 */
const cellTable = (n, first) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}<w:tblLook w:val="04A0"/></w:tblPr>`
  + '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="3000"/></w:tblGrid>'
  + `<w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${first}</w:tc>${cell(n, 1, 2)}</w:tr></w:tbl>`;
/** A table nested in a cell: one cell, N<n>A N<n>B. */
const nested = (n) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}<w:tblLook w:val="04A0"/></w:tblPr>`
  + '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
  + `<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${para('', words('', `N${n}A`, `N${n}B`))}</w:tc></w:tr></w:tbl>`;

/** A heading with a hidden run after its second word, then a third word. */
const hiddenHeading = (n) => H(n, words('', `HEAD${n}A`, `HEAD${n}B`) + t(` HIDDEN${n}`, '<w:vanish/>') + t(` HEAD${n}C`));
/** A heading with a footnote reference after its second word, then a third word; the note reads NOTE<n>. */
const footnotes = [];
const noteHeading = (n) => {
  footnotes.push(n);
  return H(n, words('', `HEAD${n}A`, `HEAD${n}B`) + run(`<w:footnoteReference w:id="${footnotes.length}"/>`, '<w:rStyle w:val="FootnoteReference"/>') + t(` HEAD${n}C`));
};

// --- the kinds: their content and their range --------------------------------------------------------
const KINDS = {
  HB: { content: (n) => [H(n), B(n)], range: (n) => `from head${n}b to body${n}a` },
  BH: { content: (n) => [B(n), H(n)], range: (n) => `from body${n}b to head${n}a` },
  LB: { content: (n) => [L(n), B(n)], range: (n) => `from list${n}b to body${n}a` },
  BL: { content: (n) => [B(n), L(n)], range: (n) => `from body${n}b to list${n}a` },
  HMB: { content: (n) => [H(n), P(n), B(n)], range: (n) => `from head${n}b to body${n}a` },
  TBL: { content: (n) => [H(n), T(n), B(n)], range: (n) => `from head${n}b to body${n}a` },
  INTO: { content: (n) => [H(n), T(n), B(n)], range: (n) => `from head${n}b to t${n}r2c1a` },
  CELL: { content: (n) => [H(n), T(n), B(n)], range: (n) => `from t${n}r1c1b to body${n}a` },
  HID: { content: (n) => [hiddenHeading(n), B(n)], range: (n) => `from head${n}b to body${n}a` },
  FN: { content: (n) => [noteHeading(n), B(n)], range: (n) => `from head${n}b to body${n}a` },
  CC: { content: (n) => [CC(n)], range: (n) => `control c52-${n}` },
  W1: { content: (n) => [B(n), N(n)], range: (n) => `whole body${n}a` },
  W2: { content: (n) => [H(n), B(n), N(n)], range: (n) => `from head${n}b to-mark body${n}a` },
  // CR-002 section 55's D2: a cell's last paragraph is a list item, so whether Word keeps its properties shows
  TC1: { content: (n) => [cellTable(n, L(n))], range: (n, op) => (op === 'paragraph-delete' ? `paragraph list${n}a` : `cell list${n}a`) },
  TC2: { content: (n) => [cellTable(n, nested(n) + L(n))], range: (n) => `paragraph list${n}a` },
  TC3: { content: (n) => [cellTable(n, `<w:sdt><w:sdtPr><w:alias w:val="Check 52 control ${n}"/><w:tag w:val="c52-${n}"/><w:id w:val="${52000 + n}"/></w:sdtPr>`
    + `<w:sdtContent>${L(n)}</w:sdtContent></w:sdt>`)], range: (n) => `control c52-${n}` },
};

// --- the two read-only regions ------------------------------------------------------------------------
const del = (id, inner) => `<w:del w:id="${id}" ${AUTHOR}>${inner}</w:del>`;
const ins = (id, inner) => `<w:ins w:id="${id}" ${AUTHOR}>${inner}</w:ins>`;
const delRun = (text, rPr = '') => run(`<w:delText xml:space="preserve">${text}</w:delText>`, rPr);
/**
 * VIEWS: what a view across paragraphs reads before anything is written. A tracked deletion from the heading's
 * second word through its mark to the body paragraph's first word; a comment from the body paragraph's second
 * word to the list item's first; a tracked insertion from the list item's second word through its mark to the
 * plain paragraph's first word; a table.
 */
const views = (n) => [
  para(HEADING + `<w:rPr>${del(9002, '')}</w:rPr>`, t(`HEAD${n}A `) + del(9001, delRun(`HEAD${n}B`))),
  para('<w:jc w:val="center"/>', del(9003, delRun(`BODY${n}A `, ITALIC)) + '<w:commentRangeStart w:id="0"/>' + t(`BODY${n}B`, ITALIC)),
  para(listPPr(n) + `<w:rPr>${ins(9005, '')}</w:rPr>`, t(`LIST${n}A`, UNDERLINE) + '<w:commentRangeEnd w:id="0"/>'
    + run('<w:commentReference w:id="0"/>', '<w:rStyle w:val="CommentReference"/>') + ins(9004, t(` LIST${n}B`, UNDERLINE))),
  para('', ins(9006, t(`PLAIN${n}A `)) + t(`PLAIN${n}B`)),
  T(n),
];
/** SEARCH: Word's ^p (a paragraph mark) against paragraphs, a list item, a table's cells and rows. */
const searchQueries = (n) => ['^p', `head${n}b^pbody${n}a`, `body${n}b^plist${n}a`, `list${n}b^pt${n}r1c1a`, `t${n}r1c1b^p`,
  `t${n}r1c1b^pt${n}r1c2a`, `t${n}r1c2b^p`, `t${n}r1c2b^pt${n}r2c1a`, `t${n}r2c2b^p`, `t${n}r2c2b^pplain${n}a`, `^p^p`];

// --- the cases ----------------------------------------------------------------------------------------
const ALL = ['HB', 'BH', 'LB', 'BL', 'HMB', 'TBL', 'INTO', 'CELL', 'HID', 'FN', 'CC', 'W1', 'W2'];
const MATRIX = [
  ['delete', ALL, ['untracked', 'tracked']],
  ['replace', ['HB', 'BH', 'LB', 'BL', 'TBL', 'INTO', 'CELL', 'CC', 'W1', 'W2'], ['untracked', 'tracked']],
  ['bold', ['HB', 'TBL', 'CELL', 'CC'], ['untracked', 'tracked']],
  ['paragraph-before', ['HB', 'TBL', 'CELL', 'CC', 'W1', 'W2'], ['untracked']],
  ['paragraph-after', ['HB', 'TBL', 'CELL', 'CC', 'W1', 'W2'], ['untracked']],
  ['paragraph-before', ['HB'], ['tracked']],
  ['paragraph-after', ['HB'], ['tracked']],
  ['text-start', ['HB', 'CC', 'W1', 'W2'], ['untracked']],
  ['text-end', ['HB', 'CC', 'W1', 'W2'], ['untracked']],
  ['comment', ['HB', 'TBL', 'CELL', 'CC', 'W1', 'W2', 'FN'], ['untracked']],
  // CR-002 section 55's D2 (Word's table rule): what Office JS's edits leave in a cell
  ['table-end', ['TC1'], ['untracked', 'tracked']],
  ['paragraph-delete', ['TC1', 'TC2'], ['untracked', 'tracked']],
  ['clear', ['TC1'], ['untracked', 'tracked']],
  ['ooxml-replace', ['TC1'], ['untracked']],
  ['control-delete', ['TC3'], ['untracked']],
];

/** Check 52's cases, in document order: { n, kind, op, tracking, spec, content }. */
export function check52Cases() {
  footnotes.length = 0;
  const cases = [];
  const add = (kind, op, tracking, spec, content) => {
    const n = cases.length + 1;
    cases.push({ n, kind, op, tracking, spec: spec(n, op), content: content(n) });
  };
  add('VIEWS', 'read', 'untracked', () => '', views);
  add('SEARCH', 'search', 'untracked', (n) => `queries ${searchQueries(n).join(';')}`, (n) => [H(n), B(n), L(n), T(n), P(n)]);
  for (const [op, kinds, trackings] of MATRIX) {
    for (const tracking of trackings) {
      for (const kind of kinds) add(kind, op, tracking, KINDS[kind].range, KINDS[kind].content);
    }
  }
  return cases;
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, kind, op, tracking, spec }) => `BEGIN${n} ${kind} ${op} ${tracking}${spec === '' ? '' : ` ${spec}`}`;

const LIST_PARAGRAPH = '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>'
  + '<w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>';
// the created package's styles part has the footnote styles only in an XML comment
const FOOTNOTE_STYLES = '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
  + '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
  + '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>';
const COMMENT_STYLES = '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/>'
  + '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/>'
  + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>';

/** One decimal list, and a w:num restarting it at 1 for each case holding a list item. */
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
    run('<w:footnoteRef/>', '<w:rStyle w:val="FootnoteReference"/>') + t(` NOTE${n}`))}</w:footnote>`).join('')
  + '</w:footnotes>';
const commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments${ROOT_NS}><w:comment w:id="0" ${AUTHOR} w:initials="C52">`
  + para('<w:pStyle w:val="CommentText"/>', run('<w:annotationRef/>', '<w:rStyle w:val="CommentReference"/>') + t('SPANNING')) + '</w:comment></w:comments>';

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 52's file, in memory: what a test reads beside Word's answers. */
export async function check52Bytes() {
  const cases = check52Cases();
  const regions = cases.flatMap((c) => [para('', t(beginText(c))), ...c.content, para('', t(`END${c.n}`))]);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document${ROOT_NS}><w:body>`
    + para('', t('Check 52. Each case is a region: its BEGIN paragraph, its content, its END paragraph.')) + regions.join('')
    + para('', t('End of check 52.')) + `${sectPr}</w:body></w:document>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };
  put('word/document.xml', documentXml);
  put('word/numbering.xml', numbering(cases));
  put('word/footnotes.xml', footnotesXml());
  put('word/comments.xml', commentsXml);
  edit('word/styles.xml', (xml) => xml.replace('</w:styles>', `${LIST_PARAGRAPH}${FOOTNOTE_STYLES}${COMMENT_STYLES}</w:styles>`));
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck52Numbering" Type="${REL}/numbering" Target="numbering.xml"/>`
    + `<Relationship Id="rIdCheck52Footnotes" Type="${REL}/footnotes" Target="footnotes.xml"/>`
    + `<Relationship Id="rIdCheck52Comments" Type="${REL}/comments" Target="comments.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    `<Override PartName="/word/numbering.xml" ContentType="${CT}.numbering+xml"/>`
    + `<Override PartName="/word/footnotes.xml" ContentType="${CT}.footnotes+xml"/>`
    + `<Override PartName="/word/comments.xml" ContentType="${CT}.comments+xml"/></Types>`));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check52Bytes();
  await writeFile(join(here, '52-ranges.docx'), bytes);
  console.log(`wrote 52-ranges.docx (${bytes.length} bytes, ${check52Cases().length} cases)`);
}
