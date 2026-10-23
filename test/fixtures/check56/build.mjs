// The files for Word check 56 (CR-002 section 61: a bound block control of plain text holds one paragraph).
// Run from the repository root after `npm run build`:
//
//   node test/fixtures/check56/build.mjs
//
// writes into this directory, which git ignores (Word reads the share's copies):
// - 56-writes.docx, for the Script Lab snippet: a region per case, a paragraph `BEGIN<n> <op> <kind>`, a block-level
//   content control tagged c56-<n> holding one paragraph `start<n>`, bound to `value<n>` in the custom XML part whose
//   namespace is `urn:check56`, and a paragraph `END<n>`; `check56Bytes` gives a test the same parts;
// - four files only to open, each made by this package's own writers today (section 61 not implemented): a bound
//   block control given a second block (`check56OpenFiles`).
// The kinds: `untyped` (no type element), `text-multi` (`w:text w:multiLine="1"`), `text-single` (`w:text`),
// `richtext` (`w:richText`). The parts are written as text, not marshalled, but for what the four files' writes add.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATA_NS = 'urn:check56';
export const STORE_ITEM_ID = '{56565656-5656-4565-8565-565656565656}';

const t = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;

/** Each kind's type element in w:sdtPr. */
export const KINDS = { untyped: '', 'text-multi': '<w:text w:multiLine="1"/>', 'text-single': '<w:text/>', richtext: '<w:richText/>' };
/** What the snippet does to each control, through Office JS's ContentControl. */
export const OPS = ['paragraph-end', 'table-end', 'text-newline', 'ooxml-end'];

/** Check 56's snippet cases, in document order: { n, op, kind }. */
export function check56Cases() {
  const out = [];
  for (const kind of Object.keys(KINDS)) for (const op of OPS) out.push({ n: out.length + 1, op, kind });
  return out;
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, op, kind }) => `BEGIN${n} ${op} ${kind}`;

/** A block-level control tagged c56-<n>, of a kind, bound to value<n>, holding one paragraph start<n>. */
const boundControl = (n, kind) => `<w:sdt><w:sdtPr><w:alias w:val="Check 56 control ${n}"/><w:tag w:val="c56-${n}"/><w:id w:val="${56000 + n}"/>`
  + `<w:dataBinding w:prefixMappings="xmlns:ns0='${DATA_NS}'" w:xpath="/ns0:check56[1]/ns0:value${n}[1]" w:storeItemID="${STORE_ITEM_ID}"/>`
  + `${KINDS[kind]}</w:sdtPr><w:sdtContent>${para(t(`start${n}`))}</w:sdtContent></w:sdt>`;

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
const documentOf = (title, blocks) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body>`
  + para(t(title)) + blocks.join('') + para(t('End of check 56.')) + `${sectPr}</w:body></w:document>`;
const dataOf = (ns) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><check56 xmlns="${DATA_NS}">`
  + ns.map((n) => `<value${n}>start${n}</value${n}>`).join('') + '</check56>';

/** Check 56's snippet file, in memory. */
export async function check56Bytes() {
  const cases = check56Cases();
  const regions = cases.flatMap((c) => [para(t(beginText(c))), boundControl(c.n, c.kind), para(t(`END${c.n}`))]);
  return packageOf(documentOf('Check 56. Each case is a region: its BEGIN paragraph, a bound block content control, its END paragraph.', regions),
    dataOf(cases.map((c) => c.n)));
}

/** A one-row, two-cell table, as a fragment. */
const TABLE = `<w:tbl xmlns:w="${W}"><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>`
  + `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>CELLA</w:t></w:r></w:p></w:tc>`
  + `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>CELLB</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

/**
 * The four files only to open, by name: one bound block control each, of a kind, given a second block by this
 * package's own writer, its node left as it was (docx4j-python's 3a: the node doesn't matter).
 */
const OPEN_FILES = [
  ['56a-untyped-two-paragraphs.docx', 'untyped', 'paragraph'],
  ['56b-text-two-paragraphs.docx', 'text-multi', 'paragraph'],
  ['56c-richtext-two-paragraphs.docx', 'richtext', 'paragraph'],
  ['56d-untyped-table.docx', 'untyped', 'table'],
];

/** The open-only files, { name, kind, add, bytes }. */
export async function check56OpenFiles() {
  const out = [];
  for (const [name, kind, add] of OPEN_FILES) {
    const title = `Check 56, ${name}: a bound block control (${kind}) given a ${add === 'table' ? 'table before its paragraph' : 'second paragraph'} by @docx4j/core-ts.`;
    const pkg = await WordprocessingMLPackage.load(await packageOf(documentOf(title, [boundControl(1, kind)]), dataOf([1])));
    const control = (await pkg.getBody()).contentControls.find((c) => c.tag === 'c56-1');
    // the table before the paragraph, so that the control still ends with one: the table alone is measured
    if (add === 'table') await control.paragraphs[0].insertOoxml(TABLE, 'Before');
    else control.insertParagraph('second paragraph', 'End');
    out.push({ name, kind, add, bytes: await pkg.save() });
  }
  return out;
}

/** A created package with this document part and this custom XML part, bound by STORE_ITEM_ID. */
async function packageOf(documentXml, data) {
  const props = `<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="${STORE_ITEM_ID}" `
    + 'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>';
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };
  put('word/document.xml', documentXml);
  put('customXml/item1.xml', data);
  put('customXml/itemProps1.xml', props);
  put('customXml/_rels/item1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + `<Relationship Id="rId1" Type="${REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>`);
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck56Data" Type="${REL}/customXml" Target="../customXml/item1.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>'));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check56Bytes();
  await writeFile(join(here, '56-writes.docx'), bytes);
  console.log(`wrote 56-writes.docx (${bytes.length} bytes, ${check56Cases().length} cases)`);
  for (const { name, bytes: b } of await check56OpenFiles()) {
    await writeFile(join(here, name), b);
    console.log(`wrote ${name} (${b.length} bytes)`);
  }
}
