// The file for Word check 64 (CR-002 section 63.4: check 63's writes on placeholder controls in Word's own form, and a
// bound control's node). Run from the repository root after `npm run build`:
//
//   node test/fixtures/check64/build.mjs
//
// writes 64-placeholders-word.docx into this directory, which git ignores (Word reads the share's copy); `check64Bytes`
// gives the same parts to a test. The parts are written as text, not marshalled.
//
// Check 63's controls had no glossary part and no w:placeholder, so Word had only the content's run to format a write
// from. Here each control is in Word's form (as Word saved one in check 34, fixtures/revisions/check34/34a-mapped.docx):
// w:placeholder naming DefaultPlaceholder_-1854013440, and that file's glossary parts, copied unedited. Each case is a
// region: a paragraph `BEGIN<n> <op> <shape> <mode>`, its content, a paragraph `END<n>`. The shapes:
//   block, run            a control showing its placeholder, block-level or between `PLAIN<n>A ` and ` PLAIN<n>B`
//   block-bound, run-bound    the same as w:text controls bound to an empty node value<n> (check 63's, in Word's form)
//   block-value, run-value    w:text controls bound to value<n> holding `VALUE<n>`, which they show: no placeholder
//   made-block, made-run  no control: the snippet makes one with Office JS's insertContentControl, around the region's
//                         empty paragraph or between the two words, before the write
// The ops are check 63's (check63/build.mjs lists them), with control-text-start: control.insertText(NEW<n>, 'Start').
import { readFileSync } from 'node:fs';
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATA_NS = 'urn:check64';
export const STORE_ITEM_ID = '{64646464-6464-4646-8646-646464646464}';
export const PLACEHOLDER = 'Click or tap here to enter text.';
/** Word's default placeholder building block, which the glossary from check 34 holds. */
export const DOC_PART = 'DefaultPlaceholder_-1854013440';
/** The Word-made file the glossary parts come from. */
export const GLOSSARY_SOURCE = join(here, '..', 'revisions', 'check34', '34a-mapped.docx');

const t = (text, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const PLACEHOLDER_RUN = t(PLACEHOLDER, '<w:rStyle w:val="PlaceholderText"/>');

const TEXT_OPS = ['control-text-replace', 'control-text-start', 'control-text-end', 'content-text-replace', 'content-text-end'];
const BLOCK_OPS = [...TEXT_OPS, 'control-paragraph-end', 'paragraph-text-replace', 'paragraph-text-start', 'paragraph-text-end'];
const MADE_BLOCK_OPS = ['control-text-replace', 'control-text-end', 'content-text-replace', 'paragraph-text-end'];
const MADE_RUN_OPS = ['control-text-replace', 'control-text-end', 'content-text-replace'];
const NODE_BLOCK_OPS = ['control-text-replace', 'content-text-replace', 'paragraph-text-replace'];
const NODE_RUN_OPS = ['control-text-replace', 'content-text-replace'];
export const OPS = [...new Set([...BLOCK_OPS, ...NODE_BLOCK_OPS])];
export const MODES = ['untracked', 'tracked'];

/** Check 64's cases, in document order: { n, op, shape, mode }. */
export function check64Cases() {
  const out = [];
  const add = (op, shape, mode) => out.push({ n: out.length + 1, op, shape, mode });
  for (const mode of MODES) {
    for (const op of BLOCK_OPS) add(op, 'block', mode);
    for (const op of TEXT_OPS) add(op, 'run', mode);
    for (const op of MADE_BLOCK_OPS) add(op, 'made-block', mode);
    for (const op of MADE_RUN_OPS) add(op, 'made-run', mode);
  }
  for (const shape of ['block-bound', 'block-value']) for (const op of NODE_BLOCK_OPS) add(op, shape, 'untracked');
  for (const shape of ['run-bound', 'run-value']) for (const op of NODE_RUN_OPS) add(op, shape, 'untracked');
  return out;
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, op, shape, mode }) => `BEGIN${n} ${op} ${shape} ${mode}`;

/** Whether a shape's control is bound to value<n>, and whether it shows its placeholder. */
const bound = (shape) => /-(bound|value)$/.test(shape);
const showing = (shape) => !shape.endsWith('-value');

/** A control tagged c64-<n> in Word's form (34a-mapped's): w:placeholder, then the flag, then a binding and w:text. */
function control(n, shape, content) {
  return `<w:sdt><w:sdtPr><w:alias w:val="Check 64 control ${n}"/><w:tag w:val="c64-${n}"/><w:id w:val="${64000 + n}"/>`
    + `<w:placeholder><w:docPart w:val="${DOC_PART}"/></w:placeholder>${showing(shape) ? '<w:showingPlcHdr/>' : ''}`
    + (bound(shape) ? `<w:dataBinding w:prefixMappings="xmlns:ns0='${DATA_NS}'" w:xpath="/ns0:check64[1]/ns0:value${n}[1]" w:storeItemID="${STORE_ITEM_ID}"/><w:text/>` : '')
    + `</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
}

/** A case's content. */
function contentOf({ n, shape }) {
  const runLevel = (inner) => para(t(`PLAIN${n}A `) + inner + t(` PLAIN${n}B`));
  const inner = showing(shape) ? PLACEHOLDER_RUN : t(`VALUE${n}`);
  if (shape === 'made-block') return '<w:p/>';
  if (shape === 'made-run') return runLevel('');
  return shape.startsWith('run') ? runLevel(control(n, shape, inner)) : control(n, shape, para(inner));
}

/** Word's Placeholder Text style, as 34a-mapped's styles part has it (its w:rsid left out). */
const STYLES = '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>';
const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** The glossary parts of 34a-mapped.docx, unedited: { part: bytes }, and their content types. */
export function glossaryParts() {
  const source = unzipSync(readFileSync(GLOSSARY_SOURCE));
  const parts = Object.fromEntries(Object.entries(source).filter(([k]) => k.startsWith('word/glossary/')));
  const types = strFromU8(source['[Content_Types].xml']).match(/<Override PartName="\/word\/glossary\/[^"]*" ContentType="[^"]*"\/>/g);
  return { parts, types };
}

/** Check 64's file, in memory. */
export async function check64Bytes() {
  const cases = check64Cases();
  const regions = cases.flatMap((c) => [para(t(beginText(c))), contentOf(c), para(t(`END${c.n}`))]);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body>`
    + para(t('Check 64. Each case is a region: its BEGIN paragraph, a content control in Word\'s form (or a place for one), its END paragraph.'))
    + regions.join('') + para(t('End of check 64.')) + `${sectPr}</w:body></w:document>`;
  const nodes = cases.filter((c) => bound(c.shape));
  const data = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><check64 xmlns="${DATA_NS}">`
    + nodes.map((c) => (showing(c.shape) ? `<value${c.n}/>` : `<value${c.n}>VALUE${c.n}</value${c.n}>`)).join('') + '</check64>';
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
  put('customXml/_rels/item1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + `<Relationship Id="rId1" Type="${REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>`);
  const glossary = glossaryParts();
  Object.assign(entries, glossary.parts);
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck64Data" Type="${REL}/customXml" Target="../customXml/item1.xml"/>`
    + `<Relationship Id="rIdCheck64Glossary" Type="${REL}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    `<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>${glossary.types.join('')}</Types>`));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check64Bytes();
  await writeFile(join(here, '64-placeholders-word.docx'), bytes);
  console.log(`wrote 64-placeholders-word.docx (${bytes.length} bytes, ${check64Cases().length} cases)`);
}
