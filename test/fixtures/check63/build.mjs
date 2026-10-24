// The file for Word check 63 (CR-002 section 63: writing into a content control showing its placeholder). Run from the
// repository root after `npm run build`:
//
//   node test/fixtures/check63/build.mjs
//
// writes 63-placeholders.docx into this directory, which git ignores (Word reads the share's copy); `check63Bytes` gives
// the same bytes to a test. The parts are written as text, not marshalled.
//
// Each case is a region: a paragraph `BEGIN<n> <op> <shape> <mode>`, its content, a paragraph `END<n>`. The content is
// one control tagged c63-<n>, showing Word's own placeholder text in the Placeholder Text style, with `w:showingPlcHdr`
// (no glossary part, as check 58's). The shapes:
//   block       a block-level control holding one paragraph
//   run         a run-level control between two words, `PLAIN<n>A ` and ` PLAIN<n>B`
//   block-bound, run-bound   the same, `w:text`, bound to an empty node value<n> in the custom XML part `urn:check63`
// The ops (check63-script-lab.js writes `NEW<n>`, with change tracking off or on as the mode says):
//   control-text-replace, control-text-end         control.insertText(NEW<n>, 'Replace' | 'End')
//   control-paragraph-end                          control.insertParagraph(NEW<n>, 'End')
//   content-text-replace, content-text-end         control.getRange('Content').insertText(NEW<n>, 'Replace' | 'End')
//   content-delete                                 control.getRange('Content').delete()
//   whole-text-replace                             control.getRange('Whole').insertText(NEW<n>, 'Replace')
//   paragraph-text-replace, -start, -end           the control's paragraph (block) or the paragraph holding it (run):
//                                                  paragraph.insertText(NEW<n>, 'Replace' | 'Start' | 'End')
//   paragraph-after, paragraph-before              paragraph.insertParagraph(NEW<n>, 'After' | 'Before')
//   paragraph-content-delete                       paragraph.getRange('Content').delete()
// and `open`, written nothing: what this package leaves after a delete, a control showing its placeholder with no
// placeholder text (`run-empty`: an empty w:sdtContent; `block-empty`: an empty paragraph).
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATA_NS = 'urn:check63';
export const STORE_ITEM_ID = '{63636363-6363-4636-8636-636363636363}';
export const PLACEHOLDER = 'Click or tap here to enter text.';

const t = (text, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const PLACEHOLDER_RUN = t(PLACEHOLDER, '<w:rStyle w:val="PlaceholderText"/>');

/** What each op writes: through the control, its ranges, or its paragraph; which shapes it is run on. */
const BLOCK_OPS = ['control-text-replace', 'control-text-end', 'control-paragraph-end', 'content-text-replace', 'content-text-end',
  'content-delete', 'paragraph-text-replace', 'paragraph-text-start', 'paragraph-text-end', 'paragraph-after', 'paragraph-before',
  'paragraph-content-delete'];
const RUN_OPS = ['control-text-replace', 'control-text-end', 'content-text-replace', 'content-text-end', 'content-delete',
  'whole-text-replace', 'paragraph-text-replace', 'paragraph-content-delete'];
const BOUND_OPS = ['control-text-replace', 'content-text-replace', 'paragraph-text-replace'];
export const OPS = [...new Set([...BLOCK_OPS, ...RUN_OPS])];
export const MODES = ['untracked', 'tracked'];

/** Check 63's cases, in document order: { n, op, shape, mode } (`mode` absent for `open`). */
export function check63Cases() {
  const out = [];
  const add = (op, shape, mode) => out.push({ n: out.length + 1, op, shape, ...(mode ? { mode } : {}) });
  for (const mode of MODES) {
    for (const op of BLOCK_OPS) add(op, 'block', mode);
    for (const op of RUN_OPS) add(op, 'run', mode);
    for (const op of BOUND_OPS) add(op, 'block-bound', mode);
    for (const op of BOUND_OPS) add(op, 'run-bound', mode);
  }
  add('open', 'run-empty');
  add('open', 'block-empty');
  return out;
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, op, shape, mode }) => `BEGIN${n} ${op} ${shape}${mode ? ` ${mode}` : ''}`;

/** A control tagged c63-<n> showing its placeholder, its content as given; bound to value<n> if `bound`. */
const control = (n, content, bound) => `<w:sdt><w:sdtPr><w:alias w:val="Check 63 control ${n}"/><w:tag w:val="c63-${n}"/><w:id w:val="${63000 + n}"/>`
  + '<w:showingPlcHdr/>'
  + (bound ? `<w:dataBinding w:prefixMappings="xmlns:ns0='${DATA_NS}'" w:xpath="/ns0:check63[1]/ns0:value${n}[1]" w:storeItemID="${STORE_ITEM_ID}"/><w:text/>` : '')
  + `</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/** A case's content: its control, block-level or between two words. */
function contentOf({ n, shape }) {
  const runLevel = (inner) => para(t(`PLAIN${n}A `) + inner + t(` PLAIN${n}B`));
  switch (shape) {
    case 'block': return control(n, para(PLACEHOLDER_RUN), false);
    case 'run': return runLevel(control(n, PLACEHOLDER_RUN, false));
    case 'block-bound': return control(n, para(PLACEHOLDER_RUN), true);
    case 'run-bound': return runLevel(control(n, PLACEHOLDER_RUN, true));
    case 'run-empty': return runLevel(control(n, '', false));
    case 'block-empty': return control(n, '<w:p/>', false);
    default: throw new Error(`no shape ${shape}`);
  }
}

const STYLES = '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>';
const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 63's file, in memory. */
export async function check63Bytes() {
  const cases = check63Cases();
  const regions = cases.flatMap((c) => [para(t(beginText(c))), contentOf(c), para(t(`END${c.n}`))]);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body>`
    + para(t('Check 63. Each case is a region: its BEGIN paragraph, a content control showing its placeholder, its END paragraph.'))
    + regions.join('') + para(t('End of check 63.')) + `${sectPr}</w:body></w:document>`;
  const bound = cases.filter((c) => c.shape.endsWith('-bound')).map((c) => c.n);
  const data = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><check63 xmlns="${DATA_NS}">${bound.map((n) => `<value${n}/>`).join('')}</check63>`;
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
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck63Data" Type="${REL}/customXml" Target="../customXml/item1.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>'));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check63Bytes();
  await writeFile(join(here, '63-placeholders.docx'), bytes);
  console.log(`wrote 63-placeholders.docx (${bytes.length} bytes, ${check63Cases().length} cases)`);
}
