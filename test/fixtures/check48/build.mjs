// The files for Word checks 48 and 48b (CR-012 D6, sections 6.2 and 12.3): what Office JS's text holds for
// each kind of inline content, before phase 4 moves Paragraph.text onto this package's own run reader. Run
// from the repository root after `npm run build`:
//
//   node test/fixtures/check48/build.mjs
//
// (`readers.test.mjs` imports `check48Bytes` instead, which builds them in memory: CR-012 phase 4)
//
// writes 48-text.docx, and 48b-text.docx with two cases more (symbols in other fonts, a run hidden through
// its character style), into this directory, which git ignores (test/README.md check 48 says why, and
// what to do with it; the share holds the copy Word reads). Each case is one paragraph, "Case n:
// <what>.", then two upper-case marker words with the item between them (FOOTA, a footnote reference,
// FOOTB), so the snippet can read each paragraph's text, search across the item, and read a range from
// one marker to the other. The parts are written as text, not
// marshalled, so the markup is exactly what is listed here; the picture is a 48 by 24 pixel green PNG.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';
import { pngOf } from '../../helpers.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
};
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const ROOT_NS = Object.entries(NS).map(([p, uri]) => ` xmlns:${p}="${uri}"`).join('');
const REVISION = 'w:author="Check 48" w:date="2026-10-21T00:00:00Z"';

const t = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const run = (inner, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}${inner}</w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
/** A case: its label, then the markers and the item, in one paragraph. */
const testCase = (n, what, inner) => para(t(`Case ${n}: ${what}. `) + inner);
const field = (code, result) => run('<w:fldChar w:fldCharType="begin"/>')
  + run(`<w:instrText xml:space="preserve"> ${code} </w:instrText>`)
  + run('<w:fldChar w:fldCharType="separate"/>') + t(result) + run('<w:fldChar w:fldCharType="end"/>');

const picture = '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="457200" cy="228600"/>'
  + '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="16" name="Picture 16" descr="a green rectangle"/>'
  + '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
  + `<a:graphic><a:graphicData uri="${NS.pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="16" name="green.png"/><pic:cNvPicPr/></pic:nvPicPr>`
  + '<pic:blipFill><a:blip r:embed="rIdCheck48Image"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
  + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
  + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';

const cell = (text) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${para(t(text))}</w:tc>`;
const table = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
  + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((b) => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('')
  + '</w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>'
  + `<w:tr>${cell('CELLA')}${cell('CELLB')}</w:tr></w:tbl>`;

const superscript = '<w:vertAlign w:val="superscript"/>';
const cases = [
  testCase(1, 'a footnote reference, then an endnote reference',
    t('FOOTA') + run('<w:footnoteReference w:id="1"/>', superscript) + t('FOOTB and ENDA')
    + run('<w:endnoteReference w:id="1"/>', superscript) + t('ENDB')),
  testCase(2, 'a comment on the middle word, its reference after it',
    t('COMA ') + '<w:commentRangeStart w:id="0"/>' + t('COMMENTED') + '<w:commentRangeEnd w:id="0"/>'
    + run('<w:commentReference w:id="0"/>') + t('COMB')),
  testCase(3, 'a line break (w:br)', t('LINEA') + run('<w:br/>') + t('LINEB')),
  testCase(4, 'a page break (w:br type page)', t('PAGEA') + run('<w:br w:type="page"/>') + t('PAGEB')),
  testCase(5, 'a column break (w:br type column)', t('COLA') + run('<w:br w:type="column"/>') + t('COLB')),
  testCase(6, 'a carriage return (w:cr)', t('CRA') + run('<w:cr/>') + t('CRB')),
  testCase(7, 'a tab, then a right-aligned positional tab (w:ptab)',
    t('TABA') + run('<w:tab/>') + t('TABB') + run('<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/>') + t('RIGHTB')),
  testCase(8, 'a non-breaking hyphen, then an optional hyphen',
    t('HYPHA') + run('<w:noBreakHyphen/>') + t('HYPHB') + run('<w:softHyphen/>') + t('HYPHC')),
  testCase(9, 'a Wingdings symbol (w:sym, F04A, a smiling face)',
    t('SYMA') + run('<w:sym w:font="Wingdings" w:char="F04A"/>') + t('SYMB')),
  testCase(10, 'hidden text (w:vanish) between the markers',
    t('HIDA') + run('<w:t>HIDDENTEXT</w:t>', '<w:vanish/>') + t('HIDB')),
  testCase(11, 'a tracked deletion, then a tracked insertion',
    t('REVA') + `<w:del w:id="101" ${REVISION}>${run('<w:delText>DELETEDTEXT</w:delText>')}</w:del>`
    + `<w:ins w:id="102" ${REVISION}>${t('INSERTEDTEXT')}</w:ins>` + t('REVB')),
  testCase(12, 'a field (begin, code, separate, result, end) whose result is not its code\'s',
    t('FLDA ') + field('QUOTE "CODETWELVE"', 'RESULTTWELVE') + t(' FLDB')),
  testCase(13, 'a simple field (w:fldSimple) whose result is not its code\'s',
    t('SIMA ') + '<w:fldSimple w:instr=" QUOTE &quot;CODETHIRTEEN&quot; ">' + t('RESULTTHIRTEEN') + '</w:fldSimple>' + t(' SIMB')),
  testCase(14, 'a hyperlink to the bookmark at the end',
    t('LINKA ') + `<w:hyperlink w:anchor="check48end">${run('<w:t>LINKTEXT</w:t>', '<w:color w:val="0563C1"/><w:u w:val="single"/>')}</w:hyperlink>` + t(' LINKB')),
  testCase(15, 'an inline content control',
    t('SDTA ') + '<w:sdt><w:sdtPr><w:alias w:val="Check 48"/><w:id w:val="48015"/></w:sdtPr><w:sdtContent>'
    + t('SDTTEXT') + '</w:sdtContent></w:sdt>' + t(' SDTB')),
  testCase(16, 'an inline picture', t('PICA') + run(picture) + t('PICB')),
  testCase(17, 'an equation (m:oMath), one variable equal to one',
    t('MATHA ') + '<m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>' + t(' MATHB')),
  para(t('Case 18: a table of one row and two cells, then a paragraph.')),
  table,
  para(t('AFTERTABLE')),
];

/**
 * Check 48b's two cases more (CR-012 section 6.2, the editor's refinements): symbols in other fonts,
 * to tell a table of Word's from one constant character, and a run hidden through a character style.
 */
const cases48b = [
  testCase(19, 'w:sym: a Wingdings check mark (F0FC), Symbol alpha (F061) and plus-minus (F0B1), Segoe UI Symbol 263A',
    t('SYM2A') + run('<w:sym w:font="Wingdings" w:char="F0FC"/>') + t('SYM2B SYM3A') + run('<w:sym w:font="Symbol" w:char="F061"/>')
    + t('SYM3B SYM4A') + run('<w:sym w:font="Symbol" w:char="F0B1"/>') + t('SYM4B SYM5A')
    + run('<w:sym w:font="Segoe UI Symbol" w:char="263A"/>') + t('SYM5B')),
  testCase(20, 'a run hidden through its character style (Check48Hidden, w:vanish)',
    t('HIDSA') + run('<w:t>STYLEHIDDEN</w:t>', '<w:rStyle w:val="Check48Hidden"/>') + t('HIDSB')),
];
const hiddenStyle = '<w:style w:type="character" w:customStyle="1" w:styleId="Check48Hidden"><w:name w:val="Check48Hidden"/>'
  + '<w:rPr><w:vanish/></w:rPr></w:style>';

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
const documentOf = (name, extra) => {
  const body = [para(t(`${name}. Each case is one paragraph: two upper-case markers with the item between them.`)), ...cases, ...extra,
    para('<w:bookmarkStart w:id="480" w:name="check48end"/>' + t(`End of ${name.toLowerCase()}.`) + '<w:bookmarkEnd w:id="480"/>')];
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document${ROOT_NS}><w:body>${body.join('')}${sectPr}</w:body></w:document>`;
};

/** Word writes a separator and a continuation separator first in each notes part. */
const separators = (kind) => `<w:${kind} w:type="separator" w:id="-1">${para(run('<w:separator/>'))}</w:${kind}>`
  + `<w:${kind} w:type="continuationSeparator" w:id="0">${para(run('<w:continuationSeparator/>'))}</w:${kind}>`;
const notes = (kind, ref, text) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind}s${ROOT_NS}>`
  + separators(kind) + `<w:${kind} w:id="1">${para(run(`<w:${ref}/>`, superscript) + t(` ${text}`))}</w:${kind}></w:${kind}s>`;
const comments = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments${ROOT_NS}>`
  + `<w:comment w:id="0" ${REVISION} w:initials="C48">${para(run('<w:annotationRef/>') + t('The comment of case 2.'))}</w:comment></w:comments>`;

/** One file's bytes: check 48's cases, and 48b's after them where `extra` holds them. */
async function bytesOf(name, extra) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };

  put('word/document.xml', documentOf(name, extra));
  if (extra.length > 0) edit('word/styles.xml', (xml) => xml.replace('</w:styles>', `${hiddenStyle}</w:styles>`));
  put('word/footnotes.xml', notes('footnote', 'footnoteRef', 'The footnote of case 1.'));
  put('word/endnotes.xml', notes('endnote', 'endnoteRef', 'The endnote of case 1.'));
  put('word/comments.xml', comments);
  entries['word/media/check48.png'] = new Uint8Array(Buffer.from(pngOf(48, 24, [0, 160, 80]), 'base64'));
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck48Footnotes" Type="${REL}/footnotes" Target="footnotes.xml"/>`
    + `<Relationship Id="rIdCheck48Endnotes" Type="${REL}/endnotes" Target="endnotes.xml"/>`
    + `<Relationship Id="rIdCheck48Comments" Type="${REL}/comments" Target="comments.xml"/>`
    + `<Relationship Id="rIdCheck48Image" Type="${REL}/image" Target="media/check48.png"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    (xml.includes('Extension="png"') ? '' : '<Default Extension="png" ContentType="image/png"/>')
    + `<Override PartName="/word/footnotes.xml" ContentType="${CT}.footnotes+xml"/>`
    + `<Override PartName="/word/endnotes.xml" ContentType="${CT}.endnotes+xml"/>`
    + `<Override PartName="/word/comments.xml" ContentType="${CT}.comments+xml"/></Types>`));

  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

/** Check 48's file (`'48'`) or 48b's, in memory: what a test reads beside Word's answers (CR-012 phase 4). */
export async function check48Bytes(which) {
  return which === '48b' ? bytesOf('Check 48b', cases48b) : bytesOf('Check 48', []);
}

async function build(file, which) {
  const bytes = await check48Bytes(which);
  await writeFile(join(here, file), bytes);
  console.log(`wrote ${file} (${bytes.length} bytes)`);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  await build('48-text.docx', '48');
  await build('48b-text.docx', '48b');
}
