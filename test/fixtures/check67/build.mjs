// The files for Word check 67 (CR-002 section 63.13: why checks 58 and 63 disagree on the bare form's look). Run from the
// repository root after `npm run build`:
//
//   node test/fixtures/check67/build.mjs
//
// writes 67a-plain.docx, 67b-custom-xml.docx and 67c-revisions.docx into this directory, which git ignores (Word reads the
// share's copies); `check67Bytes(file)` gives a test the same parts. The parts are written as text, not marshalled.
//
// Checks 58 (cases 20 to 22) and 63 (cases 1, 2 and 14) gave a control in the bare form (w:showingPlcHdr, no
// w:placeholder, no glossary) the same untracked insertText, on the same markup and the same Placeholder Text definition,
// in the same Word build. Word's save of check 58 holds a plain run; check 63's keeps the Placeholder Text style, with
// w:color auto for 'End'. What differed between the two:
//   - tracking: check 63 set changeTrackingMode 'Off' before each write; check 58 never touched it;
//   - the view: check 63 ran in Simple Markup, check 58's fills in All Markup;
//   - the lookup: check 63 found the control in body.getContentControls() by tag, check 58 by
//     body.contentControls.getByTag(tag).getFirst();
//   - the file: check 63's held a custom XML part bound to some controls, check 58's tracked changes elsewhere.
// Each file here gives check 58's three writes - a block control's insertText Replace and End, a run-level control's End -
// in four phases, each through both lookups. The phases run in this order, the two leaving tracking alone first, since
// setting it may last for the session:
//   simple-untouched   Simple Markup, changeTrackingMode never touched (check 58's tracking)
//   all-untouched      All Markup, never touched (check 58's fills)
//   all-off            All Markup, changeTrackingMode set 'Off' before each write and after it (check 63's tracking)
//   simple-off         Simple Markup, set 'Off' (check 63's run)
// The files differ only in what else they hold: 67a nothing; 67b a custom XML part and a control bound to it, left alone
// (check 63's file); 67c a tracked insertion and deletion by another author in its first paragraph (check 58's file held
// tracked changes). Each case is a region, `BEGIN<n> <op> <shape> <phase> <lookup>`, its control tagged c67-<n>, `END<n>`.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATA_NS = 'urn:check67';
export const STORE_ITEM_ID = '{67676767-6767-4676-8676-676767676767}';
export const PLACEHOLDER = 'Click or tap here to enter text.';

const t = (text, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const PLACEHOLDER_RUN = t(PLACEHOLDER, '<w:rStyle w:val="PlaceholderText"/>');

/** The three files: name, letter, what else each holds. */
export const CHECK67_FILES = [
  { name: '67a-plain.docx', letter: 'a', holds: 'nothing else' },
  { name: '67b-custom-xml.docx', letter: 'b', holds: 'a custom XML part and a control bound to it' },
  { name: '67c-revisions.docx', letter: 'c', holds: 'tracked changes by another author' },
];

/** The phases, in the order they run: the view Jason sets, and whether the snippet sets changeTrackingMode. */
export const PHASES = ['simple-untouched', 'all-untouched', 'all-off', 'simple-off'];
/** Check 58's three writes: the op and the shape. */
export const WRITES = [['control-text-replace', 'block'], ['control-text-end', 'block'], ['control-text-end', 'run']];
/** The two ways to find the control: check 58's (`tag`, getByTag) and check 63's (`list`, getContentControls). */
export const LOOKUPS = ['tag', 'list'];

/** Check 67's cases, the same in every file, in document order: { n, op, shape, phase, lookup }. */
export function check67Cases() {
  const out = [];
  for (const phase of PHASES) {
    for (const lookup of LOOKUPS) {
      for (const [op, shape] of WRITES) out.push({ n: out.length + 1, op, shape, phase, lookup });
    }
  }
  return out;
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, op, shape, phase, lookup }) => `BEGIN${n} ${op} ${shape} ${phase} ${lookup}`;

/** A control tagged c67-<n> in the bare form, showing Word's placeholder text: check 58's and 63's markup. */
const control = (n, content, bound = '') => `<w:sdt><w:sdtPr><w:alias w:val="Check 67 control ${n}"/><w:tag w:val="c67-${n}"/><w:id w:val="${67000 + n}"/>`
  + `<w:showingPlcHdr/>${bound}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

function contentOf({ n, shape }) {
  if (shape === 'block') return control(n, para(PLACEHOLDER_RUN));
  return para(t(`PLAIN${n}A `) + control(n, PLACEHOLDER_RUN) + t(` PLAIN${n}B`));
}

const STYLES = '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>';
const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** The title paragraph: in 67c, with a word inserted and one deleted by Author A, tracked. */
function title(file) {
  const text = `Check 67${file.letter}: check 58's writes into a control in the bare form, in four phases; this file holds ${file.holds}.`;
  if (file.letter !== 'c') return para(t(text));
  const mark = 'w:author="Author A" w:date="2026-10-20T00:00:00Z"';
  return para(t(text + ' ') + `<w:ins w:id="901" ${mark}>${t('Inserted.')}</w:ins>`
    + `<w:del w:id="902" ${mark}><w:r><w:delText xml:space="preserve"> Deleted.</w:delText></w:r></w:del>`);
}

/** A check 67 file, in memory. */
export async function check67Bytes(file) {
  const cases = check67Cases();
  const regions = cases.flatMap((c) => [para(t(beginText(c))), contentOf(c), para(t(`END${c.n}`))]);
  // 67b: a control bound to a node, left alone, as check 63's file had controls bound to its part
  const boundControl = file.letter === 'b'
    ? para(t('A bound control, left alone: ')) + control(0, para(PLACEHOLDER_RUN),
      `<w:dataBinding w:prefixMappings="xmlns:ns0='${DATA_NS}'" w:xpath="/ns0:check67[1]/ns0:value0[1]" w:storeItemID="${STORE_ITEM_ID}"/><w:text/>`)
    : '';
  const documentXml = `${declaration}<w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body>`
    + title(file) + boundControl + regions.join('') + para(t('End of check 67.')) + `${sectPr}</w:body></w:document>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };
  put('word/document.xml', documentXml);
  edit('word/styles.xml', (xml) => xml.replace('</w:styles>', `${STYLES}</w:styles>`));
  if (file.letter === 'b') {
    put('customXml/item1.xml', `${declaration}<check67 xmlns="${DATA_NS}"><value0/></check67>`);
    put('customXml/itemProps1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="${STORE_ITEM_ID}" `
      + 'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>');
    put('customXml/_rels/item1.xml.rels', `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
      + `<Relationship Id="rId1" Type="${REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>`);
    edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
      `<Relationship Id="rIdCheck67Data" Type="${REL}/customXml" Target="../customXml/item1.xml"/></Relationships>`));
    edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
      '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>'));
  }
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  for (const file of CHECK67_FILES) {
    const bytes = await check67Bytes(file);
    await writeFile(join(here, file.name), bytes);
    console.log(`wrote ${file.name} (${bytes.length} bytes, ${check67Cases().length} cases)`);
  }
}
