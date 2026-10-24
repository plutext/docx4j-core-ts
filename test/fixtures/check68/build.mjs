// The file for Word check 68 (CR-002 section 63.19: the run properties a write into a placeholder takes, and how Word
// combines a glossary entry with a control's own run properties). Run from the repository root after `npm run build`:
//
//   node test/fixtures/check68/build.mjs
//
// writes 68-placeholder-run-properties.docx into this directory, which git ignores (Word reads the share's copy);
// `check68Bytes` gives a test the same parts. The parts are written as text, not marshalled.
//
// Every placeholder Word was measured on so far carried the Placeholder Text style alone, so two rules agreed: this
// package writes the text plain, and docx4j-python gives it the placeholder run's properties but the style. And section
// 63.18 applies a control's own run properties (w:sdtPr/w:rPr) over a glossary entry's runs, as docx4j does, unmeasured.
// The two sources carry different properties here, so a case shows where each property came from: the placeholder run
// (or the entry) italic, the control bold. Each case is a region, `BEGIN<n> <op> <shape> <form> run-<r> control-<c>`, its
// control tagged c68-<n>, `END<n>`:
//   form bare   w:showingPlcHdr, no w:placeholder; its placeholder run in Placeholder Text, with italic (run-i), bold
//               (run-b) or nothing more (run-none); the control's w:sdtPr/w:rPr bold (control-b) or none (control-none).
//               run-b control-b is what docx4j's binder writes: placeholder.xml's run under the control's properties
//   form word   w:placeholder naming the default entry (66a's glossary: the document and styles parts), the control bold;
//               its content Word's (run-none: the entry's run as Word copies it) or this package's since section 63.18
//               (run-b: the entry's run with the control's bold over it)
//   form value  the same Word's form control holding a bold value, VALUE<n>, naming the default entry (run-none) or a
//               custom entry whose run is italic (run-i); the snippet deletes its content, and Word brings the entry back
// The ops: control-text-replace and control-text-end write NEW<n>; content-delete deletes a value control's content.
// Developer > Design Mode stays off throughout (check 63: in Design Mode a placeholder is editable text).
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
/** A custom entry, its run italic, named as Word names one. */
export const CUSTOM_ENTRY = '68686868686846868686868686868686';
export const CUSTOM_TEXT = 'Enter the name.';

const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const t = (text, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const STYLE = '<w:rStyle w:val="PlaceholderText"/>';
/** A run's properties for a placeholder: the style, then the extra one. */
const runRPr = (r) => STYLE + (r === 'i' ? '<w:i/>' : r === 'b' ? '<w:b/>' : '');

/** Check 68's cases, in document order: { n, op, shape, form, run, control }. */
export function check68Cases() {
  const out = [];
  const add = (op, shape, form, run, control) => out.push({ n: out.length + 1, op, shape, form, run, control });
  for (const [run, control] of [['i', 'none'], ['none', 'b'], ['b', 'b']]) {
    for (const shape of ['block', 'run']) for (const op of ['control-text-replace', 'control-text-end']) add(op, shape, 'bare', run, control);
  }
  for (const run of ['none', 'b']) {
    for (const shape of ['block', 'run']) for (const op of ['control-text-replace', 'control-text-end']) add(op, shape, 'word', run, 'b');
  }
  for (const run of ['none', 'i']) for (const shape of ['block', 'run']) add('content-delete', shape, 'value', run, 'b');
  return out;
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, op, shape, form, run, control }) => `BEGIN${n} ${op} ${shape} ${form} run-${run} control-${control}`;

/** A case's control: its properties, and its content (a placeholder's run, or a bold value). */
function control({ n, form, run, control: c }, inner) {
  const entry = form === 'bare' ? '' : `<w:placeholder><w:docPart w:val="${form === 'value' && run === 'i' ? CUSTOM_ENTRY : DEFAULT_ENTRY}"/></w:placeholder>`;
  return `<w:sdt><w:sdtPr>${c === 'b' ? '<w:rPr><w:b/></w:rPr>' : ''}<w:alias w:val="Check 68 control ${n}"/><w:tag w:val="c68-${n}"/>`
    + `<w:id w:val="${68000 + n}"/>${entry}${form === 'value' ? '' : '<w:showingPlcHdr/>'}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
}

function contentOf(c) {
  const run = c.form === 'value' ? t(`VALUE${c.n}`, '<w:b/>') : t(PLACEHOLDER, runRPr(c.run));
  if (c.shape === 'block') return control(c, para(run));
  return para(t(`PLAIN${c.n}A `) + control(c, run) + t(` PLAIN${c.n}B`));
}

/** 66a's glossary (the document and styles parts), given the custom entry, its run italic. */
function glossary() {
  const custom = `<w:docPart><w:docPartPr><w:name w:val="${CUSTOM_ENTRY}"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>`
    + '<w:types><w:type w:val="bbPlcHdr"/></w:types><w:behaviors><w:behavior w:val="content"/></w:behaviors>'
    + '<w:guid w:val="{68686868-6868-4686-8686-686868686868}"/></w:docPartPr>'
    + `<w:docPartBody>${para(t(CUSTOM_TEXT, `${STYLE}<w:i/>`))}</w:docPartBody></w:docPart>`;
  return { ...GLOSSARY_WITH_STYLES, 'word/glossary/document.xml': GLOSSARY_WITH_STYLES['word/glossary/document.xml'].replace('</w:docParts>', `${custom}</w:docParts>`) };
}

const STYLES = '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>';
const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 68's file, in memory. */
export async function check68Bytes() {
  const cases = check68Cases();
  const regions = cases.flatMap((c) => [para(t(beginText(c))), contentOf(c), para(t(`END${c.n}`))]);
  const documentXml = `${declaration}<w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body>`
    + para(t("Check 68: the run properties a write into a placeholder takes, and a glossary entry under a control's own run properties."))
    + regions.join('') + para(t('End of check 68.')) + `${sectPr}</w:body></w:document>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  const put = (part, text) => { entries[part] = strToU8(text); };
  const edit = (part, f) => { if (!entries[part]) throw new Error(`no ${part} in the created package`); put(part, f(strFromU8(entries[part]))); };
  put('word/document.xml', documentXml);
  edit('word/styles.xml', (xml) => xml.replace('</w:styles>', `${STYLES}</w:styles>`));
  for (const [part, text] of Object.entries(glossary())) put(part, text);
  edit('word/_rels/document.xml.rels', (xml) => xml.replace('</Relationships>',
    `<Relationship Id="rIdCheck68Glossary" Type="${REL}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`));
  edit('[Content_Types].xml', (xml) => xml.replace('</Types>',
    '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>'
    + '<Override PartName="/word/glossary/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'));
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check68Bytes();
  await writeFile(join(here, '68-placeholder-run-properties.docx'), bytes);
  console.log(`wrote 68-placeholder-run-properties.docx (${bytes.length} bytes, ${check68Cases().length} cases)`);
}
