// The file for Word check 70 (CR-002 section 63.25: where a paragraph's End is, when the paragraph ends in a marker, a
// revision, a field or a text box). Run from the repository root after `npm run build`:
//
//   node test/fixtures/check70/build.mjs
//
// writes 70-paragraph-endings.docx into this directory, which git ignores (Word reads the share's copy); `check70Bytes`
// gives a test the same parts. The parts are written as text, not marshalled.
//
// Check 69 found that Word writes at a paragraph's End after everything the paragraph holds, outside a trailing run-level
// control, and that the text takes a trailing hidden run's properties. Check 70 asks the same of the endings check 69 left
// unmeasured, and settles section 63.24's D13 (which run's properties the text takes after a control). Each case is a
// region, `BEGIN<n> <op> <shape>`, its paragraph, `END<n>`; the paragraph is `TEXT<n> ` and then:
//   bookmark-span      a bookmark BM<n> around TEXT<n>, its w:bookmarkEnd last
//   bookmark-point     an empty bookmark PT<n> after TEXT<n>
//   comment            a comment around TEXT<n>: its w:commentRangeEnd, then its reference run, last
//   tracked-del        a tracked deletion, DELETED<n>
//   tracked-ins        a tracked insertion, INSERTED<n>
//   field-complex      a complex QUOTE field whose result is FIELD<n>
//   field-simple       a w:fldSimple QUOTE field whose result is FIELD<n>
//   textbox            an inline text box (wps, in mc:AlternateContent) holding BOX<n>
//   bold-ph            TEXT<n> bold, then a run-level control showing its placeholder in Word's form (D13)
//   bold-control-text  TEXT<n> bold, then a run-level control holding VALUE<n> (D13)
//   control-bold-text  TEXT<n> plain, then a run-level control holding VALUE<n> bold (D13)
// Every op writes NEW<n> at the paragraph's End, by check 69's routes but the document body's: Paragraph.insertText 'End',
// its getRange('End')'s insertText at each location, its getRange('Content')'s 'End', and a one-cell table's cell body's
// insertText 'End' (cell-body-end). Developer > Design Mode stays off, Review > Display for Review is All Markup, and
// the snippet turns change tracking off.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';
import { GLOSSARY_WITH_STYLES, DEFAULT_ENTRY, PLACEHOLDER } from '../check66/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
};
const REL = NS.r;
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const ROOT_NS = Object.entries(NS).map(([p, uri]) => ` xmlns:${p}="${uri}"`).join('');
export { DEFAULT_ENTRY, PLACEHOLDER };
export const AUTHOR = 'Check 70';
const REVISION = `w:author="${AUTHOR}" w:date="2026-10-25T00:00:00Z"`;

const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const run = (inner, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}${inner}</w:r>`;
const t = (text, rPr = '') => run(`<w:t xml:space="preserve">${text}</w:t>`, rPr);
const para = (inner) => `<w:p>${inner}</w:p>`;
const BOLD = '<w:b/>';

export const SHAPES = ['bookmark-span', 'bookmark-point', 'comment', 'tracked-del', 'tracked-ins', 'field-complex', 'field-simple',
  'textbox', 'bold-ph', 'bold-control-text', 'control-bold-text'];
/** The writes at a paragraph's End: check 69's, the document body's left out. */
export const OPS = ['paragraph-end', 'end-before', 'end-after', 'end-start', 'end-end', 'end-replace', 'content-end', 'cell-body-end'];

/** Check 70's cases, in document order: { n, op, shape }. */
export function check70Cases() {
  const out = [];
  for (const shape of SHAPES) for (const op of OPS) out.push({ n: out.length + 1, op, shape });
  return out;
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, op, shape }) => `BEGIN${n} ${op} ${shape}`;

/** A run-level control: showing its placeholder in Word's form, or holding a value (bold or not). */
function control(n, kind) {
  const entry = kind === 'ph' ? `<w:placeholder><w:docPart w:val="${DEFAULT_ENTRY}"/></w:placeholder><w:showingPlcHdr/>` : '';
  const inner = kind === 'ph' ? t(PLACEHOLDER, '<w:rStyle w:val="PlaceholderText"/>') : t(`VALUE${n}`, kind === 'bold-value' ? BOLD : '');
  return `<w:sdt><w:sdtPr><w:alias w:val="Check 70 control ${n}"/><w:tag w:val="c70-${n}"/><w:id w:val="${70000 + n}"/>${entry}</w:sdtPr>`
    + `<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
}

/** An inline text box holding `text` (check 50's), in mc:AlternateContent with a plain-text Fallback. */
const textBox = (id, text) => run(`<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">`
  + `<wp:extent cx="1828800" cy="365760"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Text Box ${id}"/>`
  + `<wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${NS.wps}"><wps:wsp><wps:cNvSpPr txBox="1"/>`
  + '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="365760"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
  + '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="0050DC"/></a:solidFill></a:ln></wps:spPr>'
  + `<wps:txbx><w:txbxContent>${para(t(text))}</w:txbxContent></wps:txbx>`
  + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
  + `</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:t xml:space="preserve">${text}</w:t></mc:Fallback></mc:AlternateContent>`);

/** A case's paragraph. */
export function paragraphOf({ n, shape }) {
  const lead = t(`TEXT${n} `);
  switch (shape) {
    case 'bookmark-span': return para(`<w:bookmarkStart w:id="${n}" w:name="BM${n}"/>${lead}<w:bookmarkEnd w:id="${n}"/>`);
    case 'bookmark-point': return para(`${lead}<w:bookmarkStart w:id="${n}" w:name="PT${n}"/><w:bookmarkEnd w:id="${n}"/>`);
    case 'comment': return para(`<w:commentRangeStart w:id="${n}"/>${lead}<w:commentRangeEnd w:id="${n}"/>`
      + run(`<w:commentReference w:id="${n}"/>`, '<w:rStyle w:val="CommentReference"/>'));
    case 'tracked-del': return para(`${lead}<w:del w:id="${7000 + n}" ${REVISION}>${run(`<w:delText xml:space="preserve">DELETED${n}</w:delText>`)}</w:del>`);
    case 'tracked-ins': return para(`${lead}<w:ins w:id="${7000 + n}" ${REVISION}>${t(`INSERTED${n}`)}</w:ins>`);
    case 'field-complex': return para(lead + run('<w:fldChar w:fldCharType="begin"/>') + run(`<w:instrText xml:space="preserve"> QUOTE "FIELD${n}" </w:instrText>`)
      + run('<w:fldChar w:fldCharType="separate"/>') + t(`FIELD${n}`) + run('<w:fldChar w:fldCharType="end"/>'));
    case 'field-simple': return para(`${lead}<w:fldSimple w:instr=" QUOTE &quot;FIELD${n}&quot; ">${t(`FIELD${n}`)}</w:fldSimple>`);
    case 'textbox': return para(lead + textBox(70000 + n, `BOX${n}`));
    case 'bold-ph': return para(t(`TEXT${n} `, BOLD) + control(n, 'ph'));
    case 'bold-control-text': return para(t(`TEXT${n} `, BOLD) + control(n, 'value'));
    case 'control-bold-text': return para(lead + control(n, 'bold-value'));
    default: throw new Error(`no shape ${shape}`);
  }
}

const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((b) => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
const cell = (inner) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}<w:tblLook w:val="04A0"/></w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="7200"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="7200" w:type="dxa"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`;

/** A case's region: its BEGIN paragraph, its paragraph (in a one-cell table for cell-body-end), its END paragraph. */
const regionOf = (c) => para(t(beginText(c))) + (c.op === 'cell-body-end' ? cell(paragraphOf(c)) : paragraphOf(c)) + para(t(`END${c.n}`));

/** The comments part: one comment per comment case, `COMMENT<n>`. */
const commentsXml = (cases) => `${declaration}<w:comments${ROOT_NS}>`
  + cases.filter((c) => c.shape === 'comment').map((c) => `<w:comment w:id="${c.n}" ${REVISION} w:initials="C70">`
    + `<w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>${run('<w:annotationRef/>', '<w:rStyle w:val="CommentReference"/>')}${t(`COMMENT${c.n}`)}</w:p></w:comment>`).join('')
  + '</w:comments>';

const STYLES = '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/>'
  + '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/>'
  + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>';
const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 70's file, in memory. */
export async function check70Bytes() {
  const cases = check70Cases();
  const documentXml = `${declaration}<w:document${ROOT_NS}><w:body>`
    + para(t("Check 70: where a paragraph's End is, when the paragraph ends in a marker, a revision, a field or a text box."))
    + cases.map(regionOf).join('') + para(t('End of check 70.')) + `${sectPr}</w:body></w:document>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };
  put('word/document.xml', documentXml);
  put('word/comments.xml', commentsXml(cases));
  edit('word/styles.xml', (xml) => xml.replace('</w:styles>', `${STYLES}</w:styles>`));
  for (const [part, text] of Object.entries(GLOSSARY_WITH_STYLES)) put(part, text);
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck70Glossary" Type="${REL}/glossaryDocument" Target="glossary/document.xml"/>`
    + `<Relationship Id="rIdCheck70Comments" Type="${REL}/comments" Target="comments.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    `<Override PartName="/word/glossary/document.xml" ContentType="${CT}.document.glossary+xml"/>`
    + `<Override PartName="/word/glossary/styles.xml" ContentType="${CT}.styles+xml"/>`
    + `<Override PartName="/word/comments.xml" ContentType="${CT}.comments+xml"/></Types>`));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check70Bytes();
  await writeFile(join(here, '70-paragraph-endings.docx'), bytes);
  console.log(`wrote 70-paragraph-endings.docx (${bytes.length} bytes, ${check70Cases().length} cases)`);
}
