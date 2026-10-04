// The inputs of test/README.md check 36 (2026-10-04, at Jason's word): does Word 2010 open a part
// whose Office extension content is NOT declared ignorable? Check 35 opened only a file whose w15
// content was declared (mc:Ignorable="w15"); objects CR-008's open question 2 and core-ts CR-002
// section 40 item 2 both rest on the undeclared case, which nobody has opened.
//
// Four documents, a pair per kind of extension content, alike but for the root's mc:Ignorable:
//   36a  a content control with <w15:appearance w15:val="tags"/>   xmlns:w15 declared, NOT ignorable
//   36b  the same                                                   mc:Ignorable="w15"
//   36c  a tracked insertion with w16du:dateUtc                     xmlns:w16du declared, NOT ignorable
//   36d  the same                                                   mc:Ignorable="w16du"
// The document part is written as a string (setXml), so the engine's own declarations play no part.
//
//   npm run build && node scripts/make-check36-inputs.mjs
//
// writes test/fixtures/revisions/check36/36[a-d]-*.docx.
import { mkdir, writeFile } from 'node:fs/promises';
import { WordprocessingMLPackage, ZipPartStore } from '../dist/index.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';

const p = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const control = `<w:sdt><w:sdtPr><w:alias w:val="36"/><w:id w:val="360001"/><w15:appearance w15:val="tags"/></w:sdtPr><w:sdtContent>${p('Text in the content control.')}</w:sdtContent></w:sdt>`;
const insertion = `<w:p><w:r><w:t xml:space="preserve">Kept text, then </w:t></w:r><w:ins w:id="1" w:author="Check 36" w:date="2026-10-04T10:00:00Z" w16du:dateUtc="2026-10-04T00:00:00Z"><w:r><w:t>inserted text</w:t></w:r></w:ins><w:r><w:t>.</w:t></w:r></w:p>`;
const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

const cases = [
  ['36a-w15-undeclared', `xmlns:w15="${W15}"`, 'a content control with w15:appearance; w15 is NOT declared ignorable', control],
  ['36b-w15-declared', `xmlns:mc="${MC}" xmlns:w15="${W15}" mc:Ignorable="w15"`, 'a content control with w15:appearance; w15 is declared ignorable', control],
  ['36c-w16du-undeclared', `xmlns:w16du="${W16DU}"`, 'a tracked insertion with w16du:dateUtc; w16du is NOT declared ignorable', insertion],
  ['36d-w16du-declared', `xmlns:mc="${MC}" xmlns:w16du="${W16DU}" mc:Ignorable="w16du"`, 'a tracked insertion with w16du:dateUtc; w16du is declared ignorable', insertion],
];

const dir = new URL('../test/fixtures/revisions/check36/', import.meta.url);
await mkdir(dir, { recursive: true });
for (const [name, declarations, what, content] of cases) {
  const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" ${declarations}><w:body>`
    + `${p(`Check 36, ${name.slice(0, 3)}: ${what}.`)}${content}${p('After.')}${sectPr}</w:body></w:document>`;
  pkg.getMainDocumentPart().setXml(xml);
  const bytes = await pkg.save();
  // the part must be exactly what was written: no declaration added on the way out
  const saved = new TextDecoder().decode(await new ZipPartStore(bytes).load('word/document.xml'));
  if (saved !== xml) throw new Error(`${name}: word/document.xml is not what was written`);
  await writeFile(new URL(`${name}.docx`, dir), bytes);
  console.log(`wrote ${name}.docx: ${/<w:document [^>]*>/.exec(saved)[0].replace(/xmlns:w="[^"]*" /, '')}`);
}
