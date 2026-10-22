// The file for Word check 54 (CR-014 phase 3, and phase 2's choices check 52 didn't measure: CR-014 section 10.6).
// Run from the repository root after `npm run build`:
//
//   node test/fixtures/check54/build.mjs
//
// writes 54-writes.docx into this directory, which git ignores (Word reads the share's copy, as for check 52);
// `check54Bytes` gives the same bytes to a test. The parts are written as text, not marshalled.
//
// As check 52's: each case is a region, a paragraph `BEGIN<n> <kind> <op> <tracking> <range>`, the kind's
// paragraphs, a paragraph `END<n>`; the range in lower case in the marker and the content's words in upper case,
// so that the snippet's whole-word, case-sensitive search finds a word in the content only. The ranges:
//   from head7b to body7a        the start of HEAD7B's word to the end of BODY7A's (expandTo)
//   from head7b to-mark body7a   the start of HEAD7B's word to the end of BODY7A's paragraph, its mark included
//   from head7b to-table t7r1c1a the start of HEAD7B's word to the end of the table holding T7R1C1A (its 'Whole' range)
//   start head7a to-mark body7a  the start of HEAD7A's paragraph to the end of BODY7A's, its mark included
//   whole body7a                 BODY7A's paragraph, its mark included (getRange('Whole'))
//   word head7b                  HEAD7B's search hit: a range within one paragraph
// The paragraphs are check 52's (a heading, HEAD; a centred body paragraph, its runs italic, BODY; a numbered list
// item, its runs underlined, LIST; a plain paragraph, PLAIN; a table of two rows, T<n>R1C1A ...) and four more: a
// paragraph in a style whose runs are italic (QUOTE), one in a style that underlines, capitalises, sizes and
// colours its runs (STYL), a block content control (c54-<n>), and words a comment or a bookmark runs over.
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
const AUTHOR = 'w:author="Check 54" w:date="2026-10-22T00:00:00Z"';

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

const H = (n) => para(HEADING, words('', `HEAD${n}A`, `HEAD${n}B`, `HEAD${n}C`));
const B = (n) => para('<w:jc w:val="center"/>', words(ITALIC, `BODY${n}A`, `BODY${n}B`));
const L = (n) => para(listPPr(n), words(UNDERLINE, `LIST${n}A`, `LIST${n}B`));
const P = (n) => para('', words('', `PLAIN${n}A`, `PLAIN${n}B`));
/** A paragraph in Check54Italic, whose runs are italic through the style. */
const Q = (n) => para('<w:pStyle w:val="Check54Italic"/>', words('', `QUOTE${n}A`, `QUOTE${n}B`));
/** A paragraph in Check54Styled, whose runs the style underlines, capitalises, sets at 16 points and colours red. */
const S = (n) => para('<w:pStyle w:val="Check54Styled"/>', words('', `STYL${n}A`, `STYL${n}B`));

const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((b) => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
const tblPr = `<w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}<w:tblLook w:val="04A0"/></w:tblPr>`;
const cell = (n, r, c) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${para('', words('', `T${n}R${r}C${c}A`, `T${n}R${r}C${c}B`))}</w:tc>`;
const T = (n) => `<w:tbl>${tblPr}<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>`
  + `<w:tr>${cell(n, 1, 1)}${cell(n, 1, 2)}</w:tr><w:tr>${cell(n, 2, 1)}${cell(n, 2, 2)}</w:tr></w:tbl>`;
/** A table of one row: its first cell holds `first` (blocks), its second T<n>R1C2A T<n>R1C2B. */
const cellTable = (n, first) => `<w:tbl>${tblPr}<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="3000"/></w:tblGrid>`
  + `<w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${first}</w:tc>${cell(n, 1, 2)}</w:tr></w:tbl>`;
/** A block content control tagged c54-<n> around `inner` (blocks). */
const control = (n, inner) => `<w:sdt><w:sdtPr><w:alias w:val="Check 54 control ${n}"/><w:tag w:val="c54-${n}"/><w:id w:val="${54000 + n}"/></w:sdtPr>`
  + `<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** The comments the file holds, as [id, text]; a kind adds its own. */
const comments = [];
const commented = (n, text) => {
  comments.push([comments.length, text]);
  return comments.length - 1;
};
const reference = (id) => run(`<w:commentReference w:id="${id}"/>`, '<w:rStyle w:val="CommentReference"/>');
/** A heading whose comment runs over its second word only, wholly inside a range from its first word to the body paragraph. */
const commentInside = (n) => {
  const id = commented(n, `INSIDE${n}`);
  return [para(HEADING, t(`HEAD${n}A `) + `<w:commentRangeStart w:id="${id}"/>` + t(`HEAD${n}B`) + `<w:commentRangeEnd w:id="${id}"/>` + reference(id) + t(` HEAD${n}C`)), B(n)];
};
/** A comment from the heading's second word to the body paragraph's second, which a range ending at BODY<n>A cuts. */
const commentAcross = (n) => {
  const id = commented(n, `ACROSS${n}`);
  return [para(HEADING, t(`HEAD${n}A `) + `<w:commentRangeStart w:id="${id}"/>` + t(`HEAD${n}B HEAD${n}C`)),
    para('<w:jc w:val="center"/>', t(`BODY${n}A `, ITALIC) + t(`BODY${n}B`, ITALIC) + `<w:commentRangeEnd w:id="${id}"/>` + reference(id))];
};
/** A bookmark from the heading's second word to the body paragraph's second, which a range ending at BODY<n>A cuts. */
const bookmarkAcross = (n) => [para(HEADING, t(`HEAD${n}A `) + `<w:bookmarkStart w:id="${5400 + n}" w:name="Check54Mark${n}"/>` + t(`HEAD${n}B HEAD${n}C`)),
  para('<w:jc w:val="center"/>', t(`BODY${n}A `, ITALIC) + t(`BODY${n}B`, ITALIC) + `<w:bookmarkEnd w:id="${5400 + n}"/>`)];

// --- the kinds: their content and their range --------------------------------------------------------
const KINDS = {
  // a range within one paragraph: a word's search hit
  IN1: { content: (n) => [H(n), B(n)], range: (n) => `word head${n}b` },
  // from a paragraph's second word to the end of the table after it, the table's whole range
  TE: { content: (n) => [H(n), T(n), P(n)], range: (n) => `from head${n}b to-table t${n}r1c1a` },
  // a paragraph's whole range, a table after it
  WT: { content: (n) => [B(n), T(n), P(n)], range: (n) => `whole body${n}a` },
  // the whole range of a cell's last paragraph, a body paragraph before it in the cell
  WL: { content: (n) => [cellTable(n, B(n) + L(n))], range: (n) => `whole list${n}a` },
  // the whole range of a content control's last paragraph
  WC: { content: (n) => [control(n, H(n) + B(n) + L(n))], range: (n) => `whole list${n}a` },
  // a whole block content control (of a plain paragraph) between the two paragraphs
  CCW: { content: (n) => [H(n), control(n, P(n)), B(n)], range: (n) => `from head${n}b to body${n}a` },
  // into a block content control, from the paragraph before it
  CCI: { content: (n) => [H(n), control(n, B(n) + L(n))], range: (n) => `from head${n}b to body${n}a` },
  // out of a block content control, to the paragraph after it
  CCO: { content: (n) => [control(n, H(n) + B(n)), P(n)], range: (n) => `from body${n}b to plain${n}a` },
  // italic through a paragraph style, then a plain paragraph
  QP: { content: (n) => [Q(n), P(n)], range: (n) => `from quote${n}b to plain${n}a` },
  // a heading (its runs 14 points through its style), then a plain paragraph
  HP: { content: (n) => [H(n), P(n)], range: (n) => `from head${n}b to plain${n}a` },
  // a list item (underlined directly), then a plain paragraph
  LP: { content: (n) => [L(n), P(n)], range: (n) => `from list${n}b to plain${n}a` },
  // a plain paragraph, then one in Check54Styled
  PS: { content: (n) => [P(n), S(n)], range: (n) => `from plain${n}b to styl${n}a` },
  // a comment over the heading's second word, wholly inside the range
  CMT: { content: commentInside, range: (n) => `from head${n}a to body${n}a` },
  // a comment from the heading into the body paragraph, its start inside the range and its end after it
  CMS: { content: commentAcross, range: (n) => `from head${n}a to body${n}a` },
  // a bookmark from the heading into the body paragraph, its start inside the range and its end after it
  BMK: { content: bookmarkAcross, range: (n) => `from head${n}a to body${n}a` },
  // a paragraph this run inserts under tracking after the heading (OWN<n>A OWN<n>B), the range into it
  OWN: { content: (n) => [H(n), B(n)], range: (n) => `from head${n}b to own${n}a` },
  // a body paragraph, then a list item
  BL: { content: (n) => [B(n), L(n)], range: (n) => `from body${n}b to list${n}a` },
  // check 52's kinds, for the content controls of phase 3
  HB: { content: (n) => [H(n), B(n)], range: (n) => `from head${n}b to body${n}a` },
  W2: { content: (n) => [H(n), B(n), P(n)], range: (n) => `from head${n}b to-mark body${n}a` },
  WP: { content: (n) => [H(n), B(n), P(n)], range: (n) => `start head${n}a to-mark body${n}a` },
  TBL: { content: (n) => [H(n), T(n), B(n)], range: (n) => `from head${n}b to body${n}a` },
  INTO: { content: (n) => [H(n), T(n), B(n)], range: (n) => `from head${n}b to t${n}r2c1a` },
  CELL: { content: (n) => [H(n), T(n), B(n)], range: (n) => `from t${n}r1c1b to body${n}a` },
};

// --- the cases, as [kind, op, tracking]; CR-014 section 10.6's item each answers ------------------------
const CASES = [
  // 1: insertParagraph on a range within one paragraph; 7: a tracked Replace within one paragraph
  ['IN1', 'paragraph-before', 'untracked'], ['IN1', 'paragraph-after', 'untracked'], ['IN1', 'paragraph-before', 'tracked'],
  ['IN1', 'replace', 'tracked'],
  // 2: a range ending with a whole table after a paragraph's tail
  ['TE', 'delete', 'untracked'], ['TE', 'delete', 'tracked'], ['TE', 'replace', 'untracked'],
  // 3: a 'Whole' range with no paragraph after it in its list
  ['WT', 'delete', 'untracked'], ['WT', 'delete', 'tracked'], ['WL', 'delete', 'untracked'], ['WL', 'delete', 'tracked'],
  ['WC', 'delete', 'untracked'], ['WC', 'delete', 'tracked'],
  // 4: a whole block control in a range, a range into one and out of one
  ['CCW', 'delete', 'untracked'], ['CCW', 'delete', 'tracked'], ['CCW', 'replace', 'untracked'], ['CCW', 'replace', 'tracked'],
  ['CCI', 'delete', 'untracked'], ['CCI', 'delete', 'tracked'], ['CCO', 'delete', 'untracked'],
  // 5: toggles other than bold, and properties that aren't toggles
  ['QP', 'italic', 'untracked'], ['QP', 'italic', 'tracked'], ['HP', 'size14', 'untracked'], ['HP', 'size14', 'tracked'],
  ['LP', 'underline', 'untracked'],
  // 6: the look a Replace gives into a paragraph whose style sets what no style set before
  ['PS', 'replace', 'untracked'],
  // 8: comments and bookmarks a deletion takes or cuts
  ['CMT', 'delete', 'untracked'], ['CMT', 'delete', 'tracked'], ['CMS', 'delete', 'untracked'], ['BMK', 'delete', 'untracked'],
  // 9: a tracked deletion over a mark this author inserted
  ['OWN', 'own-delete', 'tracked'],
  // 10: a style, a hyperlink and insertOoxml across paragraphs
  ['BL', 'style', 'untracked'], ['BL', 'style', 'tracked'], ['HB', 'hyperlink', 'untracked'], ['HB', 'ooxml', 'untracked'],
  // phase 3: insertContentControl across paragraphs
  ['HB', 'control', 'untracked'], ['W2', 'control', 'untracked'], ['WP', 'control', 'untracked'], ['TBL', 'control', 'untracked'],
  ['INTO', 'control', 'untracked'], ['CELL', 'control', 'untracked'], ['HB', 'control', 'tracked'], ['WP', 'control', 'tracked'],
];

/** Check 54's cases, in document order: { n, kind, op, tracking, spec, content }. */
export function check54Cases() {
  comments.length = 0;
  return CASES.map(([kind, op, tracking], i) => {
    const n = i + 1;
    return { n, kind, op, tracking, spec: KINDS[kind].range(n), content: KINDS[kind].content(n) };
  });
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, kind, op, tracking, spec }) => `BEGIN${n} ${kind} ${op} ${tracking} ${spec}`;

const STYLES = '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>'
  + '<w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>'
  + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Check54Italic"><w:name w:val="Check 54 Italic"/><w:basedOn w:val="Normal"/>'
  + '<w:qFormat/><w:rPr><w:i/><w:iCs/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Check54Styled"><w:name w:val="Check 54 Styled"/><w:basedOn w:val="Normal"/>'
  + '<w:qFormat/><w:rPr><w:caps/><w:color w:val="C00000"/><w:sz w:val="32"/><w:szCs w:val="32"/><w:u w:val="single"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/>'
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

const commentsXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments${ROOT_NS}>`
  + comments.map(([id, text]) => `<w:comment w:id="${id}" ${AUTHOR} w:initials="C54">`
    + para('<w:pStyle w:val="CommentText"/>', run('<w:annotationRef/>', '<w:rStyle w:val="CommentReference"/>') + t(text)) + '</w:comment>').join('')
  + '</w:comments>';

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 54's file, in memory: what a test reads beside Word's answers. */
export async function check54Bytes() {
  const cases = check54Cases();
  const regions = cases.flatMap((c) => [para('', t(beginText(c))), ...c.content, para('', t(`END${c.n}`))]);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document${ROOT_NS}><w:body>`
    + para('', t('Check 54. Each case is a region: its BEGIN paragraph, its content, its END paragraph.')) + regions.join('')
    + para('', t('End of check 54.')) + `${sectPr}</w:body></w:document>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };
  put('word/document.xml', documentXml);
  put('word/numbering.xml', numbering(cases));
  put('word/comments.xml', commentsXml());
  edit('word/styles.xml', (xml) => xml.replace('</w:styles>', `${STYLES}</w:styles>`));
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck54Numbering" Type="${REL}/numbering" Target="numbering.xml"/>`
    + `<Relationship Id="rIdCheck54Comments" Type="${REL}/comments" Target="comments.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    `<Override PartName="/word/numbering.xml" ContentType="${CT}.numbering+xml"/>`
    + `<Override PartName="/word/comments.xml" ContentType="${CT}.comments+xml"/></Types>`));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check54Bytes();
  await writeFile(join(here, '54-writes.docx'), bytes);
  console.log(`wrote 54-writes.docx (${bytes.length} bytes, ${check54Cases().length} cases)`);
}
