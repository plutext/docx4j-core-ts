// The file for Word check 69 (CR-002 section 63.23: where a paragraph's End is, when the paragraph ends in something with
// no width). Run from the repository root after `npm run build`:
//
//   node test/fixtures/check69/build.mjs
//
// writes 69-paragraph-end.docx into this directory, which git ignores (Word reads the share's copy); `check69Bytes` gives
// a test the same parts. The parts are written as text, not marshalled.
//
// A write at a paragraph's End, where the paragraph ends in a run-level control showing its placeholder: this package
// writes the text on the end of the last run with width, before the control; docx4j-python wrote it inside the control,
// after the prompt, flag on, where no reader sees it. Neither is measured, and the same choice decides a trailing hidden
// run (this package before it, docx4j-python after it) and a trailing control holding text (this package inside it).
// Each case is a region, `BEGIN<n> <op> <shape>`, its paragraph, `END<n>`; the paragraph is `TEXT<n> ` and then:
//   plain         nothing more (the baseline)
//   ph-word       a run-level control showing its placeholder in Word's form (w:placeholder naming the default entry,
//                 66a's glossary), tagged c69-<n>
//   ph-bare       the same in the bare form (w:showingPlcHdr, no w:placeholder)
//   hidden        a hidden run (w:vanish), HIDDEN<n>
//   control-text  a run-level control holding VALUE<n>, not showing a placeholder, tagged c69-<n>
//   ph-only       (no TEXT<n>) the paragraph is the ph-word control alone
// Every op writes NEW<n>, at the paragraph's End by a different route: Paragraph.insertText 'End', its getRange('End')'s
// insertText at each location, its getRange('Content')'s 'End', and Body.insertText 'End' on a body whose last paragraph it
// is: a one-cell table's cell (cell-body-end), and the document's body for the last case (body-end), its region running to
// the end of the document with no END paragraph. Developer > Design Mode stays off throughout (check 63: in Design Mode a
// placeholder is editable text).
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';
import { GLOSSARY_WITH_STYLES, DEFAULT_ENTRY, PLACEHOLDER } from '../check66/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export { DEFAULT_ENTRY, PLACEHOLDER };

const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const t = (text, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const STYLE = '<w:rStyle w:val="PlaceholderText"/>';

export const SHAPES = ['plain', 'ph-word', 'ph-bare', 'hidden', 'control-text', 'ph-only'];
/** The writes at a paragraph's End, through the paragraph and its ranges. */
export const PARAGRAPH_OPS = ['paragraph-end', 'end-before', 'end-after', 'end-start', 'end-end', 'end-replace', 'content-end'];

/** Check 69's cases, in document order: { n, op, shape }. The last, body-end, ends the document. */
export function check69Cases() {
  const out = [];
  const add = (op, shape) => out.push({ n: out.length + 1, op, shape });
  for (const shape of SHAPES) {
    for (const op of PARAGRAPH_OPS) add(op, shape);
    add('cell-body-end', shape);
  }
  add('body-end', 'ph-word');
  return out;
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, op, shape }) => `BEGIN${n} ${op} ${shape}`;

/** A run-level control: showing its placeholder (Word's form or the bare form), or holding a value. */
function control(n, kind) {
  const entry = kind === 'word' ? `<w:placeholder><w:docPart w:val="${DEFAULT_ENTRY}"/></w:placeholder>` : '';
  const flag = kind === 'value' ? '' : '<w:showingPlcHdr/>';
  const inner = kind === 'value' ? t(`VALUE${n}`) : t(PLACEHOLDER, STYLE);
  return `<w:sdt><w:sdtPr><w:alias w:val="Check 69 control ${n}"/><w:tag w:val="c69-${n}"/><w:id w:val="${69000 + n}"/>${entry}${flag}</w:sdtPr>`
    + `<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
}

/** A case's paragraph. */
export function paragraphOf({ n, shape }) {
  const lead = t(`TEXT${n} `);
  switch (shape) {
    case 'plain': return para(lead);
    case 'ph-word': return para(lead + control(n, 'word'));
    case 'ph-bare': return para(lead + control(n, 'bare'));
    case 'hidden': return para(lead + t(`HIDDEN${n}`, '<w:vanish/>'));
    case 'control-text': return para(lead + control(n, 'value'));
    case 'ph-only': return para(control(n, 'word'));
    default: throw new Error(`no shape ${shape}`);
  }
}

const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((b) => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
const cell = (inner) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}<w:tblLook w:val="04A0"/></w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="7200"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="7200" w:type="dxa"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`;

/** A case's region: its BEGIN paragraph, its paragraph (in a one-cell table for cell-body-end), its END paragraph. */
function regionOf(c) {
  const content = c.op === 'cell-body-end' ? cell(paragraphOf(c)) : paragraphOf(c);
  return para(t(beginText(c))) + content + (c.op === 'body-end' ? '' : para(t(`END${c.n}`)));
}

const STYLES = '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>';
const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 69's file, in memory. */
export async function check69Bytes() {
  const cases = check69Cases();
  const documentXml = `${declaration}<w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body>`
    + para(t("Check 69: where a paragraph's End is, when the paragraph ends in something with no width. The last case ends the document."))
    + cases.map(regionOf).join('') + `${sectPr}</w:body></w:document>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };
  put('word/document.xml', documentXml);
  edit('word/styles.xml', (xml) => xml.replace('</w:styles>', `${STYLES}</w:styles>`));
  for (const [part, text] of Object.entries(GLOSSARY_WITH_STYLES)) put(part, text);
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck69Glossary" Type="${REL}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>'
    + '<Override PartName="/word/glossary/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check69Bytes();
  await writeFile(join(here, '69-paragraph-end.docx'), bytes);
  console.log(`wrote 69-paragraph-end.docx (${bytes.length} bytes, ${check69Cases().length} cases)`);
}
