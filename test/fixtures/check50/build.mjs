// The file for Word check 50 (CR-010 section 14.9): whether Word 365 reads a Requires prefix alike
// wherever it is declared. Run from the repository root after `npm run build`:
//
//   node test/fixtures/check50/build.mjs
//
// writes 50-declarations.docx into this directory (test/README.md check 50 says what to do with it).
// Five run-level mc:AlternateContent elements, each a Choice holding a text box (wps, inline) and a
// Fallback of plain text, so which branch Word takes shows on the page. Case 1 is Word's own form
// (Requires="wps", declared on the root). Cases 2 to 4 bind the wps namespace to another prefix, declared
// on the root (2), on the mc:AlternateContent (3) or on the mc:Choice (4), and write the Choice's shape
// elements with that prefix. Case 5 requires a namespace Word does not know (zz9), so it should show its
// Fallback. The part is written as text, not marshalled, so every declaration stays where it is put.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
};
const ZZ9 = 'urn:example:zz9';

/** A text box whose shape elements carry `prefix` (bound to the wps namespace), its text `text`. */
const textBox = (prefix, id, text) => `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">`
  + `<wp:extent cx="3657600" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Text Box ${id}"/>`
  + `<wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${NS.wps}"><${prefix}:wsp><${prefix}:cNvSpPr txBox="1"/>`
  + `<${prefix}:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3657600" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
  + `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="0050DC"/></a:solidFill></a:ln></${prefix}:spPr>`
  + `<${prefix}:txbx><w:txbxContent><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:txbxContent></${prefix}:txbx>`
  + `<${prefix}:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></${prefix}:bodyPr>`
  + `</${prefix}:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`;

const label = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
/** One case: its label, then a paragraph holding the element; `on` says where `xmlns:<prefix>` goes. */
const element = (n, prefix, on, labelText) => {
  const decl = ` xmlns:${prefix}="${prefix === 'zz9' ? ZZ9 : NS.wps}"`;
  return label(`${n}. ${labelText}`)
    + `<w:p><w:r><w:t xml:space="preserve">Case ${n}: </w:t></w:r><w:r>`
    + `<mc:AlternateContent${on === 'ac' ? decl : ''}><mc:Choice Requires="${prefix}"${on === 'choice' ? decl : ''}>`
    + textBox(prefix === 'zz9' ? 'wps' : prefix, n, `${n}: the Choice (a text box)`)
    + `</mc:Choice><mc:Fallback><w:t xml:space="preserve">${n}: the Fallback (plain text, no box)</w:t></mc:Fallback></mc:AlternateContent>`
    + `</w:r></w:p>`;
};

const root = `<w:document xmlns:w="${NS.w}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:mc="${NS.mc}" xmlns:wps="${NS.wps}"`
  + ` xmlns:sroot="${NS.wps}" xmlns:zz9="${ZZ9}">`;
const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + root + '<w:body>'
  + label('Check 50. Each case is a text box if Word takes its Choice, plain text if it takes its Fallback.')
  + element(1, 'wps', 'root', 'Requires="wps", declared on the root (Word\'s own form). Expected: a text box.')
  + element(2, 'sroot', 'root', 'Requires="sroot", the wps namespace under another prefix, declared on the root.')
  + element(3, 'sac', 'ac', 'Requires="sac", the wps namespace under another prefix, declared on the mc:AlternateContent.')
  + element(4, 'schoice', 'choice', 'Requires="schoice", the wps namespace under another prefix, declared on the mc:Choice.')
  + element(5, 'zz9', 'root', 'Requires="zz9", a namespace Word does not know, declared on the root. Expected: plain text.')
  + label('End of check 50.')
  + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
  + '</w:body></w:document>';

const pkg = await WordprocessingMLPackage.createPackage();
const entries = unzipSync(await pkg.save());
const name = Object.keys(entries).find((n) => /^word\/document\.xml$/.test(n));
if (!name) throw new Error('no word/document.xml in the created package');
entries[name] = strToU8(xml);
await writeFile(join(here, '50-declarations.docx'), zipSync(entries));
console.log(`wrote 50-declarations.docx (${strFromU8(entries[name]).length} characters of document.xml)`);
