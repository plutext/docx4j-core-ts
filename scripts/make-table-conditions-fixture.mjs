// The parity fixtures test/fixtures/parity/tables-conditions.docx and tables-conditions-mode12.docx (CR-007 phase 1): tables whose
// table style has every conditional format, read through every way a table states which of them
// apply. docx4j CR-030's own probes (the other tables-*.docx fixtures) settle the name rule, the
// story boundary and the compatibility exception in Word; none of them has banding, columns,
// corners, w:cnfStyle caches, spans or a nested table, which is the arithmetic of
// TableStyleConditions. These two documents are not Word-measured: its golden is what docx4j answers.
//
//   npm run build && node scripts/make-table-conditions-fixture.mjs
//
// then regenerate the goldens (test/java/README.md).
import { writeFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from '../dist/index.mjs';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

// One colour per condition, so the last applicable one shows in w:color; toggles and paragraph
// properties on some, so the composition is exercised and not only the precedence.
const cond = (type, rPr, pPr = '') => `<w:tblStylePr w:type="${type}">${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:rPr>${rPr}</w:rPr></w:tblStylePr>`;
const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="left"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="BoldPara"><w:name w:val="Bold Para"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="right"/><w:ind w:left="240"/></w:pPr><w:rPr><w:b/><w:caps/><w:sz w:val="30"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="UnboldPara"><w:name w:val="Unbold Para"/><w:basedOn w:val="Normal"/><w:rPr><w:b w:val="0"/><w:i/></w:rPr></w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>
  <w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/><w:color w:val="112233"/></w:rPr></w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
  <w:style w:type="table" w:styleId="Base"><w:name w:val="Probe Base"/><w:basedOn w:val="TableNormal"/>
    <w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr>
    <w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/></w:tblPr>
    ${cond('wholeTable', '<w:color w:val="000001"/>', '<w:keepNext/>')}
    ${cond('firstRow', '<w:b/><w:color w:val="0000F1"/>', '<w:jc w:val="center"/>')}
    ${cond('lastRow', '<w:i/><w:color w:val="0000F2"/>')}
    ${cond('firstCol', '<w:b/><w:color w:val="0000C1"/>')}
    ${cond('lastCol', '<w:caps/><w:color w:val="0000C2"/>')}
    ${cond('band1Vert', '<w:color w:val="0000A1"/>')}
    ${cond('band2Vert', '<w:color w:val="0000A2"/>', '<w:ind w:left="120"/>')}
    ${cond('band1Horz', '<w:color w:val="0000B1"/>')}
    ${cond('band2Horz', '<w:strike/><w:color w:val="0000B2"/>')}
  </w:style>
  <w:style w:type="table" w:styleId="Full"><w:name w:val="Probe Full"/><w:basedOn w:val="Base"/>
    <w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr>
    ${cond('firstRow', '<w:smallCaps/>', '<w:spacing w:before="80"/>')}
    ${cond('nwCell', '<w:color w:val="00E001"/>')}
    ${cond('neCell', '<w:color w:val="00E002"/>')}
    ${cond('swCell', '<w:color w:val="00E003"/>')}
    ${cond('seCell', '<w:b w:val="0"/><w:color w:val="00E004"/>')}
  </w:style>
  <w:style w:type="table" w:styleId="Wide"><w:name w:val="Probe Wide Bands"/><w:basedOn w:val="Full"/>
    <w:tblPr><w:tblStyleRowBandSize w:val="2"/><w:tblStyleColBandSize w:val="2"/><w:tblLook w:val="01E0"/></w:tblPr>
  </w:style>
  <w:style w:type="table" w:styleId="NoText"><w:name w:val="Probe No Text"/>
    <w:tblPr><w:tblCellMar><w:left w:w="60" w:type="dxa"/></w:tblCellMar></w:tblPr>
    <w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/></w:tcPr></w:tblStylePr>
  </w:style>
</w:styles>`;

let n = 0;
const para = (text, { style, pPr = '', rPr = '', cnf } = {}) => {
  const props = `${style ? `<w:pStyle w:val="${style}"/>` : ''}${cnf ? `<w:cnfStyle w:val="${cnf}"/>` : ''}${pPr}`;
  return `<w:p>${props ? `<w:pPr>${props}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t>${text}</w:t></w:r></w:p>`;
};
const tc = (content, { span, cnf } = {}) => {
  const props = `${cnf ? `<w:cnfStyle w:val="${cnf}"/>` : ''}<w:tcW w:w="1200" w:type="dxa"/>${span ? `<w:gridSpan w:val="${span}"/>` : ''}`;
  return `<w:tc><w:tcPr>${props}</w:tcPr>${content}</w:tc>`;
};
const tr = (cells, { cnf, before, after } = {}) => {
  const props = `${cnf ? `<w:cnfStyle w:val="${cnf}"/>` : ''}${before ? `<w:gridBefore w:val="${before}"/>` : ''}${after ? `<w:gridAfter w:val="${after}"/>` : ''}`;
  return `<w:tr>${props ? `<w:trPr>${props}</w:trPr>` : ''}${cells.join('')}</w:tr>`;
};
const tbl = (tblPr, rows, cols = 4) => `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid>${'<w:gridCol w:w="1200"/>'.repeat(cols)}</w:tblGrid>${rows.join('')}</w:tbl>`;
/** A rows x cols table of plain cells, each holding one paragraph named by its position. */
const grid = (label, rows, cols, options = () => ({})) => Array.from({ length: rows }, (_, r) =>
  tr(Array.from({ length: cols }, (_, c) => tc(para(`${label} r${r}c${c}`, options(r, c))))));
const heading = (text) => para(`T${String(++n).padStart(2, '0')} ${text}`);

const body = [
  heading('Full, every look attribute on, 5 x 5'),
  tbl('<w:tblStyle w:val="Full"/><w:tblLook w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/>', grid('a', 5, 5), 5),
  heading('Full, the look as the bitmask 05E0 alone (no column banding)'),
  tbl('<w:tblStyle w:val="Full"/><w:tblLook w:val="05E0"/>', grid('b', 4, 4)),
  heading('Full, no w:tblLook at all: Word\'s 04A0'),
  tbl('<w:tblStyle w:val="Full"/>', grid('c', 4, 4)),
  heading('Full, look 0000 with banding off: the whole table only'),
  tbl('<w:tblStyle w:val="Full"/><w:tblLook w:val="0600"/>', grid('d', 3, 3), 3),
  heading('Wide: band sizes 2 and the look from the style'),
  tbl('<w:tblStyle w:val="Wide"/>', grid('e', 7, 7), 7),
  heading('Wide, the table\'s own band sizes 3 and 1 and its own look'),
  tbl('<w:tblStyle w:val="Wide"/><w:tblStyleRowBandSize w:val="3"/><w:tblStyleColBandSize w:val="1"/><w:tblLook w:val="0000"/>', grid('f', 7, 4)),
  heading('Full, paragraph styles and direct formatting over the table level'),
  tbl('<w:tblStyle w:val="Full"/><w:tblLook w:val="04A0"/>', grid('g', 3, 4, (r, c) => [
    { style: 'BoldPara' }, { style: 'UnboldPara' }, { rPr: '<w:rStyle w:val="Strong"/>' },
    { pPr: '<w:jc w:val="both"/><w:rPr><w:b w:val="0"/></w:rPr>', rPr: '<w:b w:val="0"/><w:color w:val="FF0000"/>' },
  ][c])),
  heading('Full, w:cnfStyle caches which contradict the positions, and one the look gates'),
  tbl('<w:tblStyle w:val="Full"/><w:tblLook w:val="04A0"/>', [
    tr([tc(para('h r0c0')), tc(para('h r0c1'), { cnf: '000100000000' }), tc(para('h r0c2', { cnf: '010000000000' }))], { cnf: '000000100000' }),
    tr([tc(para('h r1c0'), { cnf: '001000000000' }), tc(para('h r1c1', { cnf: '000000010000' })), tc(para('h r1c2', { cnf: '000010000000' }))]),
    tr([tc(para('h r2c0')), tc(para('h r2c1'), { cnf: '000000000000' }), tc(para('h r2c2', { cnf: '1' }))], { cnf: '100000000000' }),
  ], 3),
  heading('Full, grid spans, w:gridBefore and w:gridAfter'),
  tbl('<w:tblStyle w:val="Full"/><w:tblLook w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/>', [
    tr([tc(para('i r0c0-1'), { span: 2 }), tc(para('i r0c2')), tc(para('i r0c3'))]),
    tr([tc(para('i r1c1')), tc(para('i r1c2-3'), { span: 2 })], { before: 1 }),
    tr([tc(para('i r2c0')), tc(para('i r2c1'))], { after: 2 }),
    tr([tc(para('i r3c0-3'), { span: 4 })]),
  ]),
  heading('Full, a nested table (NoText) in a cell, and content controls around a row and a cell'),
  tbl('<w:tblStyle w:val="Full"/><w:tblLook w:val="04A0"/>', [
    tr([tc(para('j r0c0')), tc(para('j r0c1 before the nested table')
      + tbl('<w:tblStyle w:val="NoText"/><w:tblLook w:val="04A0"/>', grid('j nested', 2, 2), 2) + para('j r0c1 after it'))]),
    `<w:sdt><w:sdtPr><w:id w:val="101"/></w:sdtPr><w:sdtContent>${tr([tc(para('j r1c0 in a row control')), tc(`<w:sdt><w:sdtPr><w:id w:val="102"/></w:sdtPr><w:sdtContent>${para('j r1c1 in a block control')}</w:sdtContent></w:sdt>`)])}</w:sdtContent></w:sdt>`,
    tr([`<w:sdt><w:sdtPr><w:id w:val="103"/></w:sdtPr><w:sdtContent>${tc(para('j r2c0 in a cell control'))}</w:sdtContent></w:sdt>`, tc(para('j r2c1'))]),
  ], 2),
  heading('No style named: the default table style, Normal Table'),
  tbl('<w:tblLook w:val="04A0"/>', grid('k', 2, 2), 2),
  heading('A style which is missing'),
  tbl('<w:tblStyle w:val="Missing"/><w:tblLook w:val="04A0"/>', grid('l', 2, 2), 2),
  heading('NoText: conditions with no text formatting'),
  tbl('<w:tblStyle w:val="NoText"/><w:tblLook w:val="04A0"/>', grid('m', 2, 2), 2),
  para('End.'),
].join('');

const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;

// Two documents, because the [MS-DOCX] size and justification exception is a document setting:
// createPackage()'s settings (compatibility mode 15, the override on), and none at all (mode 12),
// where Normal's 12pt and left give way to the table style's size and justification.
for (const [name, settings] of [['tables-conditions', undefined], ['tables-conditions-mode12', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}/>`]]) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const main = pkg.getMainDocumentPart();
  main.setXml(document);
  pkg.ensureStyleDefinitionsPart().setXml(styles.replace(/>\s+</g, '><'));
  if (settings !== undefined) main.documentSettingsPart.setXml(settings);
  await writeFile(new URL(`../test/fixtures/parity/${name}.docx`, import.meta.url), await pkg.save());
  console.log(`wrote test/fixtures/parity/${name}.docx`);
}
