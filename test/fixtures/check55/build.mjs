// The files for Word checks 55 and 55b (CR-002 section 60: what Office JS's writers do with the characters its readers
// give; 55b, the bound controls 55 couldn't find).
// Run from the repository root after `npm run build`:
//
//   node test/fixtures/check55/build.mjs
//
// writes 55-characters.docx and 55b-bound.docx into this directory, which git ignores (Word reads the share's copy,
// as for check 54); `check55Bytes` and `check55bBytes` give a test the same parts (only the zip's timestamps differ).
// The parts are written as text, not marshalled.
//
// As check 54's: each case is a region, a paragraph `BEGIN<n> <op> <tracking> <characters>`, the case's paragraph, a
// paragraph `END<n>`. The characters are code points, `U+000B` or `U+000D,U+000A`; the snippet writes `x`, them, `y`.
// The case's paragraph is `TARGET<n>`, or for a bound control `BOUND<n> ` and a plain-text control tagged `c55-<n>`,
// bound to `value<n>` in the custom XML part `urn:check55`, holding `start<n>`.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATA_NS = 'urn:check55';
export const STORE_ITEM_ID = '{55555555-5555-4555-8555-555555555555}';

const t = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;

/** The characters each op is measured with: every Office JS reader's character, the XML 1.0 exclusions, docx4j's hyphens. */
const ALL = [[0x09], [0x0a], [0x0d], [0x0d, 0x0a], [0x0b], [0x0c], [0x0e], [0x1e], [0x1f], [0x02], [0x01], [0x07], [0x2011], [0xad]];
const BREAKS = [[0x0d], [0x0a], [0x0d, 0x0a], [0x0b]];
const CASES = [
  // insertText('x' + c + 'y', 'End') on a paragraph, and insertParagraph('x' + c + 'y', 'After') after it
  ...ALL.map((chars) => ['text-end', 'untracked', chars]),
  ...ALL.map((chars) => ['paragraph-after', 'untracked', chars]),
  // insertComment('x' + c + 'y') on the paragraph's word; then comment.content = 'x' + c + 'y' on a comment made with COMMENT<n>
  ...ALL.map((chars) => ['comment', 'untracked', chars]),
  ...BREAKS.map((chars) => ['comment-content', 'untracked', chars]),
  // insertText under tracking: the w:ins around what the character becomes
  ...[[0x09], [0x0a], [0x0d], [0x0b], [0x1e]].map((chars) => ['text-end', 'tracked', chars]),
  // a bound plain-text control, multi-line and not, given insertText('x' + c + 'y', 'Replace'): its custom XML node after
  ...[[0x0b], [0x0a], [0x0d], [0x1e], [0x02]].map((chars) => ['bound-multi', 'untracked', chars]),
  ...[[0x0b], [0x0a]].map((chars) => ['bound-single', 'untracked', chars]),
];

const code = (cp) => 'U+' + cp.toString(16).toUpperCase().padStart(4, '0');

/** Check 55's cases, in document order: { n, op, tracking, chars, spec }. */
export function check55Cases() {
  return CASES.map(([op, tracking, chars], i) => ({ n: i + 1, op, tracking, chars, spec: chars.map(code).join(',') }));
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, op, tracking, spec }) => `BEGIN${n} ${op} ${tracking} ${spec}`;

/** A plain-text control bound to value<n>, its text start<n>; multi-line or not. */
const boundControl = (n, multiLine) => `<w:sdt><w:sdtPr><w:alias w:val="Check 55 bound ${n}"/><w:tag w:val="c55-${n}"/><w:id w:val="${55000 + n}"/>`
  + `<w:dataBinding w:prefixMappings="xmlns:ns0='${DATA_NS}'" w:xpath="/ns0:check55[1]/ns0:value${n}[1]" w:storeItemID="${STORE_ITEM_ID}"/>`
  + `<w:text${multiLine ? ' w:multiLine="1"' : ''}/></w:sdtPr><w:sdtContent>${t(`start${n}`)}</w:sdtContent></w:sdt>`;

const content = (c) => c.op.startsWith('bound-') ? para(t(`BOUND${c.n} `) + boundControl(c.n, c.op === 'bound-multi')) : para(t(`TARGET${c.n}`));

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 55's file, in memory: what a test reads beside Word's answers. */
export async function check55Bytes() {
  const cases = check55Cases();
  const regions = cases.flatMap((c) => [para(t(beginText(c))), content(c), para(t(`END${c.n}`))]);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body>`
    + para(t('Check 55. Each case is a region: its BEGIN paragraph, its paragraph, its END paragraph.')) + regions.join('')
    + para(t('End of check 55.')) + `${sectPr}</w:body></w:document>`;
  const data = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><check55 xmlns="${DATA_NS}">`
    + cases.filter((c) => c.op.startsWith('bound-')).map((c) => `<value${c.n}>start${c.n}</value${c.n}>`).join('') + '</check55>';
  return packageOf(documentXml, data);
}

// --- check 55b: the bound controls check 55 couldn't find (CR-002 section 60.5) ---------------------------------
//
// Check 55's body.contentControls.getByTag found none of its seven plain-text controls. Here each case's control is
// looked for five ways, each recorded, and written through the first that finds it. U+0041 is a baseline: a plain
// character, to show a write reaches the node at all. The last case writes nothing: its node holds `line1`, a line
// feed and `line2`, and its control `stale<n>`, for what Word shows of a node's line feed on open.
const DATA_NS_B = 'urn:check55b';
const CASES_B = [
  ...[[0x41], [0x0b], [0x0a], [0x0d], [0x1e], [0x02]].map((chars) => ['bound-multi', 'untracked', chars]),
  ...[[0x41], [0x0b], [0x0a]].map((chars) => ['bound-single', 'untracked', chars]),
  ['bound-read', 'untracked', [0x0a]],
];

/** Check 55b's cases, in document order: { n, op, tracking, chars, spec }. */
export function check55bCases() {
  return CASES_B.map(([op, tracking, chars], i) => ({ n: i + 1, op, tracking, chars, spec: chars.map(code).join(',') }));
}

/** A plain-text control tagged c55b-<n>, bound to value<n> in urn:check55b, holding `text`. */
const boundControlB = (n, multiLine, text) => `<w:sdt><w:sdtPr><w:alias w:val="Check 55b bound ${n}"/><w:tag w:val="c55b-${n}"/><w:id w:val="${55100 + n}"/>`
  + `<w:dataBinding w:prefixMappings="xmlns:ns0='${DATA_NS_B}'" w:xpath="/ns0:check55b[1]/ns0:value${n}[1]" w:storeItemID="${STORE_ITEM_ID}"/>`
  + `<w:text${multiLine ? ' w:multiLine="1"' : ''}/></w:sdtPr><w:sdtContent>${t(text)}</w:sdtContent></w:sdt>`;

/** Check 55b's file, in memory, as check55Bytes makes check 55's. */
export async function check55bBytes() {
  const cases = check55bCases();
  const content = (c) => para(t(`BOUND${c.n} `) + boundControlB(c.n, c.op !== 'bound-single', c.op === 'bound-read' ? `stale${c.n}` : `start${c.n}`));
  const regions = cases.flatMap((c) => [para(t(beginText(c))), content(c), para(t(`END${c.n}`))]);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body>`
    + para(t('Check 55b. Each case is a region: its BEGIN paragraph, its paragraph, its END paragraph.')) + regions.join('')
    + para(t('End of check 55b.')) + `${sectPr}</w:body></w:document>`;
  const data = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><check55b xmlns="${DATA_NS_B}">`
    + cases.map((c) => `<value${c.n}>${c.op === 'bound-read' ? 'line1\nline2' : `start${c.n}`}</value${c.n}>`).join('') + '</check55b>';
  return packageOf(documentXml, data);
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
    `<Relationship Id="rIdCheck55Data" Type="${REL}/customXml" Target="../customXml/item1.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>'));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check55Bytes();
  await writeFile(join(here, '55-characters.docx'), bytes);
  console.log(`wrote 55-characters.docx (${bytes.length} bytes, ${check55Cases().length} cases)`);
  const b = await check55bBytes();
  await writeFile(join(here, '55b-bound.docx'), b);
  console.log(`wrote 55b-bound.docx (${b.length} bytes, ${check55bCases().length} cases)`);
}
