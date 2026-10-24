// The files for Word check 66 (CR-002 section 63.10: D9, how much glossary this package writes where a file has none).
// Run from the repository root after `npm run build`:
//
//   node test/fixtures/check66/build.mjs
//
// writes 66a-glossary-with-styles.docx and 66b-word-glossary.docx into this directory, which git ignores (Word reads the
// share's copies); `check66Bytes(file)` gives a test the same parts. The parts are written as text, not marshalled.
//
// Check 65's 65b had a glossary of the document part alone: Word kept it, but its entries lost the Placeholder Text style,
// which the glossary's own styles lacked, and a reopened save typed into moved controls off the default entry. Here:
//   66a  D9 (a): the glossary document part, holding the default entry in the style, and a styles part holding Placeholder
//        Text and the style it is based on (Word's definitions, as check 34's glossary has them), related from it;
//   66b  D9 (b): check 34's glossary, Word's six parts, unedited (check 64's).
// The same cases in both, each a region `BEGIN<n> <op> <shape> <mode>` ... `END<n>`, its control tagged c66-<n>:
//   1 open block                a control showing its placeholder, naming the default entry; nothing written
//   2 type block-bound          the same as a w:text control bound to an empty node; Jason types TYPED2 into it by hand
//                               once Word's save after the run is reopened
//   3 content-delete block-value  a control naming the default entry and holding VALUE3, not showing it; the snippet
//                               deletes its content (does Word bring the entry back, in the style?)
//   4 make made-block           no control: the snippet makes one with insertContentControl, filled from the default entry
//   5 set-placeholder-text block  control.placeholderText = 'Enter the name 5.', a new entry in this glossary
//   6 type bare-bound           a bound w:text control in the bare form (no w:placeholder), typed into as case 2 (TYPED6)
import { readFileSync } from 'node:fs';
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATA_NS = 'urn:check66';
export const STORE_ITEM_ID = '{66666666-6666-4666-8666-666666666666}';
export const PLACEHOLDER = 'Click or tap here to enter text.';
export const DEFAULT_ENTRY = 'DefaultPlaceholder_-1854013440';
/** Check 34's glossary: Word's six parts (66b's), and its default entry's w:guid (66a's too). */
export const GLOSSARY_SOURCE = join(here, '..', 'revisions', 'check34', '34a-mapped.docx');
const DEFAULT_GUID = '{7D21F9BC-12A3-463E-B08F-7DD4FFC77A03}';

const t = (text, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const placeholderRun = (text) => t(text, '<w:rStyle w:val="PlaceholderText"/>');

/** The two files: name, glossary ('with-styles', D9 (a); 'word', D9 (b)), letter. */
export const CHECK66_FILES = [
  { name: '66a-glossary-with-styles.docx', glossary: 'with-styles', letter: 'a' },
  { name: '66b-word-glossary.docx', glossary: 'word', letter: 'b' },
];

/** The cases, the same in both files, in document order. */
export function check66Cases() {
  return [
    ['open', 'block'], ['type', 'block-bound'], ['content-delete', 'block-value'], ['make', 'made-block'],
    ['set-placeholder-text', 'block'], ['type', 'bare-bound'],
  ].map(([op, shape], i) => ({ n: i + 1, op, shape, mode: 'untracked' }));
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, op, shape, mode }) => `BEGIN${n} ${op} ${shape} ${mode}`;

const bound = (shape) => shape.endsWith('-bound');
const showing = (shape) => shape !== 'block-value';

/** A control tagged c66-<n>: w:placeholder naming the default entry (but in the bare form), the flag, a binding and w:text. */
function control(n, shape, content) {
  return `<w:sdt><w:sdtPr><w:alias w:val="Check 66 control ${n}"/><w:tag w:val="c66-${n}"/><w:id w:val="${66000 + n}"/>`
    + (shape === 'bare-bound' ? '' : `<w:placeholder><w:docPart w:val="${DEFAULT_ENTRY}"/></w:placeholder>`)
    + (showing(shape) ? '<w:showingPlcHdr/>' : '')
    + (bound(shape) ? `<w:dataBinding w:prefixMappings="xmlns:ns0='${DATA_NS}'" w:xpath="/ns0:check66[1]/ns0:value${n}[1]" w:storeItemID="${STORE_ITEM_ID}"/><w:text/>` : '')
    + `</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
}

function contentOf({ n, shape }) {
  if (shape === 'made-block') return '<w:p/>';
  return control(n, shape, para(showing(shape) ? placeholderRun(PLACEHOLDER) : t(`VALUE${n}`)));
}

/** Word's definitions of Placeholder Text and the style it is based on, as check 34's glossary has them (w:rsid left out). */
export const PLACEHOLDER_STYLES = '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>'
  + '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
  + '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>';

const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
/** 66a's glossary: the document part with the default entry (Word's docPartPr, check 34's guid), its styles part, the relationship. */
export const GLOSSARY_WITH_STYLES = {
  'word/glossary/document.xml': `${declaration}<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr>`
    + `<w:name w:val="${DEFAULT_ENTRY}"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>`
    + '<w:types><w:type w:val="bbPlcHdr"/></w:types><w:behaviors><w:behavior w:val="content"/></w:behaviors>'
    + `<w:guid w:val="${DEFAULT_GUID}"/></w:docPartPr><w:docPartBody>${para(placeholderRun(PLACEHOLDER))}</w:docPartBody></w:docPart>`
    + '</w:docParts></w:glossaryDocument>',
  'word/glossary/styles.xml': `${declaration}<w:styles xmlns:w="${W}">${PLACEHOLDER_STYLES}</w:styles>`,
  'word/glossary/_rels/document.xml.rels': `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + `<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
};
const WITH_STYLES_TYPES = '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>'
  + '<Override PartName="/word/glossary/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>';

/** Check 34's glossary parts, unedited, and their content types (as check 64's build takes them). */
function wordGlossary() {
  const source = unzipSync(readFileSync(GLOSSARY_SOURCE));
  const parts = Object.fromEntries(Object.entries(source).filter(([k]) => k.startsWith('word/glossary/')));
  const types = strFromU8(source['[Content_Types].xml']).match(/<Override PartName="\/word\/glossary\/[^"]*" ContentType="[^"]*"\/>/g).join('');
  return { parts, types };
}

const STYLES = '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>';
const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** A check 66 file, in memory. */
export async function check66Bytes(file) {
  const cases = check66Cases();
  const regions = cases.flatMap((c) => [para(t(beginText(c))), contentOf(c), para(t(`END${c.n}`))]);
  const title = `Check 66${file.letter}: controls naming the default placeholder entry, ${file.glossary === 'word' ? "with Word's glossary of six parts" : 'with a glossary of the document and styles parts written by @docx4j/core-ts'}.`;
  const documentXml = `${declaration}<w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body>`
    + para(t(title)) + regions.join('') + para(t('End of check 66.')) + `${sectPr}</w:body></w:document>`;
  const data = `${declaration}<check66 xmlns="${DATA_NS}">` + cases.filter((c) => bound(c.shape)).map((c) => `<value${c.n}/>`).join('') + '</check66>';
  const props = `<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="${STORE_ITEM_ID}" `
    + 'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>';
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };
  put('word/document.xml', documentXml);
  edit('word/styles.xml', (xml) => xml.replace('</w:styles>', `${STYLES}</w:styles>`));
  put('customXml/item1.xml', data);
  put('customXml/itemProps1.xml', props);
  put('customXml/_rels/item1.xml.rels', `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + `<Relationship Id="rId1" Type="${REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>`);
  let types = '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>';
  if (file.glossary === 'word') {
    const glossary = wordGlossary();
    Object.assign(entries, glossary.parts);
    types += glossary.types;
  } else {
    for (const [part, text] of Object.entries(GLOSSARY_WITH_STYLES)) put(part, text);
    types += WITH_STYLES_TYPES;
  }
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck66Data" Type="${REL}/customXml" Target="../customXml/item1.xml"/>`
    + `<Relationship Id="rIdCheck66Glossary" Type="${REL}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>', `${types}</Types>`));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  for (const file of CHECK66_FILES) {
    const bytes = await check66Bytes(file);
    await writeFile(join(here, file.name), bytes);
    console.log(`wrote ${file.name} (${bytes.length} bytes, ${check66Cases().length} cases)`);
  }
}
