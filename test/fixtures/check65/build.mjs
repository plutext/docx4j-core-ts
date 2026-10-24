// The files for Word check 65 (CR-002 section 63.7: what Word needs of a placeholder in its own form, for D8). Run from
// the repository root after `npm run build`:
//
//   node test/fixtures/check65/build.mjs
//
// writes 65a-reference-only.docx and 65b-minimal-glossary.docx into this directory, which git ignores (Word reads the
// share's copies); `check65Bytes(file)` gives a test the same parts. The parts are written as text, not marshalled.
//
// Check 64 measured controls naming DefaultPlaceholder_-1854013440 with Word's whole glossary (six parts). Here:
//   65a  the controls name it, and the file has no glossary part;
//   65b  a glossary written here: the glossary document part alone, holding that entry (Word's docPartPr, as check
//        34's glossary has it) and one of its own text, `Enter the name.`, under a name of Word's kind (32 hex digits).
// Each case is a region, `BEGIN<n> <op> <shape> <mode>` ... `END<n>`, its control tagged c65-<n>. The shapes:
//   block        a block-level control showing its placeholder, naming the default entry
//   block-bound  the same as a w:text control bound to an empty node value<n> (urn:check65)
//   run-bound    the same, run-level, between `PLAIN<n>A ` and ` PLAIN<n>B`
//   custom       (65b) a block control showing `Enter the name.`, naming the glossary's own entry
//   missing      (65b) a block control naming an entry the glossary hasn't
//   made-block   no control: the snippet makes one with insertContentControl around the region's empty paragraph
//   bare-bound   a bound w:text control in the bare form, check 63's: no w:placeholder (docx4j CR-063's question)
// The ops are check 63's (check63/build.mjs), and:
//   set-placeholder-text   control.placeholderText = 'Enter the name <n>.', which Word keeps in the glossary
//   open                   nothing written: the control as Word reads it, and as it shows when the file is reopened
//   make                   (made-block) the control made, nothing written
//   type                   (bare-bound) nothing written by the snippet: Jason types TYPED<n> into it by hand after the run
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATA_NS = 'urn:check65';
export const STORE_ITEM_ID = '{65656565-6565-4656-8656-656565656565}';
export const PLACEHOLDER = 'Click or tap here to enter text.';
export const DEFAULT_ENTRY = 'DefaultPlaceholder_-1854013440';
/** The glossary's own entry in 65b, named as Word names one (32 hex digits), and its text. */
export const CUSTOM_ENTRY = '65656565656545658565656565650001';
export const CUSTOM_TEXT = 'Enter the name.';
/** An entry no glossary holds. */
export const MISSING_ENTRY = '65656565656545658565656565650002';

const t = (text, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const placeholderRun = (text) => t(text, '<w:rStyle w:val="PlaceholderText"/>');

/** The two files: name, glossary, cases ({ n, op, shape, mode }). */
export const CHECK65_FILES = [
  { name: '65a-reference-only.docx', glossary: 'none', letter: 'a' },
  { name: '65b-minimal-glossary.docx', glossary: 'minimal', letter: 'b' },
];

/** A file's cases, in document order. */
export function check65Cases(file) {
  const out = [];
  const add = (op, shape, mode = 'untracked') => out.push({ n: out.length + 1, op, shape, mode });
  add('control-text-end', 'block');
  add('control-paragraph-end', 'block');
  add('control-text-replace', 'block', 'tracked');
  add('control-text-replace', 'block-bound');
  add('content-text-replace', 'run-bound');
  add('set-placeholder-text', 'block');
  add('open', 'block');
  add('open', 'block-bound');
  add('make', 'made-block');
  if (file.glossary === 'minimal') {
    add('control-text-end', 'custom');
    add('open', 'custom');
    add('control-text-end', 'missing');
    add('open', 'missing');
  }
  add('type', 'bare-bound');
  return out;
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, op, shape, mode }) => `BEGIN${n} ${op} ${shape} ${mode}`;

const bound = (shape) => shape.endsWith('-bound');
const entryOf = (shape) => (shape === 'custom' ? CUSTOM_ENTRY : shape === 'missing' ? MISSING_ENTRY : DEFAULT_ENTRY);

/** A control tagged c65-<n> in Word's form: w:placeholder naming its entry, the flag, then a binding and w:text. */
function control(n, shape, content) {
  return `<w:sdt><w:sdtPr><w:alias w:val="Check 65 control ${n}"/><w:tag w:val="c65-${n}"/><w:id w:val="${65000 + n}"/>`
    + (shape === 'bare-bound' ? '' : `<w:placeholder><w:docPart w:val="${entryOf(shape)}"/></w:placeholder>`) + '<w:showingPlcHdr/>'
    + (bound(shape) ? `<w:dataBinding w:prefixMappings="xmlns:ns0='${DATA_NS}'" w:xpath="/ns0:check65[1]/ns0:value${n}[1]" w:storeItemID="${STORE_ITEM_ID}"/><w:text/>` : '')
    + `</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
}

function contentOf({ n, shape }) {
  const text = shape === 'custom' ? CUSTOM_TEXT : PLACEHOLDER;
  if (shape === 'made-block') return '<w:p/>';
  if (shape === 'run-bound') return para(t(`PLAIN${n}A `) + control(n, shape, placeholderRun(text)) + t(` PLAIN${n}B`));
  return control(n, shape, para(placeholderRun(text)));
}

/** A glossary entry as check 34's glossary (Word's) has one: its docPartPr, and its text in the Placeholder Text style. */
const docPart = (name, guid, text) => `<w:docPart><w:docPartPr><w:name w:val="${name}"/><w:category><w:name w:val="General"/>`
  + '<w:gallery w:val="placeholder"/></w:category><w:types><w:type w:val="bbPlcHdr"/></w:types><w:behaviors><w:behavior w:val="content"/></w:behaviors>'
  + `<w:guid w:val="${guid}"/></w:docPartPr><w:docPartBody>${para(placeholderRun(text))}</w:docPartBody></w:docPart>`;

/** 65b's glossary document part: the default entry (check 34's guid) and the custom one. */
export const MINIMAL_GLOSSARY = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument xmlns:w="${W}"><w:docParts>`
  + docPart(DEFAULT_ENTRY, '{7D21F9BC-12A3-463E-B08F-7DD4FFC77A03}', PLACEHOLDER)
  + docPart(CUSTOM_ENTRY, '{65656565-6565-4565-8565-656565650001}', CUSTOM_TEXT)
  + '</w:docParts></w:glossaryDocument>';

const STYLES = '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>';
const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** A check 65 file, in memory. */
export async function check65Bytes(file) {
  const cases = check65Cases(file);
  const regions = cases.flatMap((c) => [para(t(beginText(c))), contentOf(c), para(t(`END${c.n}`))]);
  const title = `Check 65${file.letter}: controls naming a placeholder entry, ${file.glossary === 'none' ? 'with no glossary part' : 'with a glossary of one part written by @docx4j/core-ts'}.`;
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body>`
    + para(t(title)) + regions.join('') + para(t('End of check 65.')) + `${sectPr}</w:body></w:document>`;
  const data = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><check65 xmlns="${DATA_NS}">`
    + cases.filter((c) => bound(c.shape)).map((c) => `<value${c.n}/>`).join('') + '</check65>';
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
  let rels = `<Relationship Id="rIdCheck65Data" Type="${REL}/customXml" Target="../customXml/item1.xml"/>`;
  let types = '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>';
  if (file.glossary === 'minimal') {
    put('word/glossary/document.xml', MINIMAL_GLOSSARY);
    rels += `<Relationship Id="rIdCheck65Glossary" Type="${REL}/glossaryDocument" Target="glossary/document.xml"/>`;
    types += '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>';
  }
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>', `${rels}</Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>', `${types}</Types>`));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  for (const file of CHECK65_FILES) {
    const bytes = await check65Bytes(file);
    await writeFile(join(here, file.name), bytes);
    console.log(`wrote ${file.name} (${bytes.length} bytes, ${check65Cases(file).length} cases)`);
  }
}
