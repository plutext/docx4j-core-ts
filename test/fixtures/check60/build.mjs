// The file for Word check 60 (CR-015 D4: what Office JS's Font reports for run properties Word reads leniently).
// Read only. Run from the repository root:
//
//   node test/fixtures/check60/build.mjs
//
// writes 60-runprops.docx into this directory, which git ignores (Word reads the share's copy); `check60Bytes`
// gives the same bytes to a test. The parts are written as text, not marshalled, so every value reaches Word as the
// probe wrote it.
//
// It is docx4j's probes (test/runprops-word365/, CR-015 section 8) in one document: for each probe in
// CHECK60_PROBES a paragraph `PROBE <name>`, then the probe's paragraphs but its title (and its w:sectPr), as it wrote them, each
// `label: MARKER (end)` with only the marker run carrying the property tested. The styles are the probes' together
// (no two probes define a style id differently; the build refuses if one does). The three props-enum probes are left
// out: Word repairs them, and its repair, not its reading, decides what they show (CR-015 section 2, rule 2).
import { readFileSync } from 'node:fs';
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';

const here = dirname(fileURLToPath(import.meta.url));
const probes = join(here, '..', '..', 'runprops-word365');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** The probes in the file, in order: docx4j's RunPropertiesWord365Test's, but the three Word repairs. */
export const CHECK60_PROBES = [
  'onoff-valid', 'onoff-empty', 'onoff-space', 'onoff-TRUE', 'onoff-yes',
  'props-b-case', 'props-b-styles', 'props-toggles-1', 'props-toggles-2',
  'props-u-valid', 'props-u-case', 'props-u-other',
  'colour-names-highlight', 'colour-names-css', 'colour-case', 'colour-hex',
  'colour-table-grey', 'colour-table-other',
  'props-u-styles', 'colour-styles', 'colour-elsewhere',
  'props-u-onoff-styles', 'colour-auto-styles', 'highlight-char-styles', 'colour-elsewhere-styles',
];

const escape = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** A probe's body paragraphs (they hold no table or nested paragraph) and its style elements, as written. */
function probeParts(name) {
  const entries = unzipSync(readFileSync(join(probes, `${name}.docx`)));
  const doc = strFromU8(entries['word/document.xml']);
  const body = doc.slice(doc.indexOf('<w:body>') + '<w:body>'.length, doc.lastIndexOf('</w:body>')).replace(/<w:sectPr>[\s\S]*<\/w:sectPr>$/, '');
  const paragraphs = body.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? [];
  if (paragraphs.join('') !== body) throw new Error(`${name}: the body is not paragraphs and a w:sectPr alone`);
  const styles = entries['word/styles.xml'] === undefined ? [] : strFromU8(entries['word/styles.xml']).match(/<w:style\b[\s\S]*?<\/w:style>/g) ?? [];
  return { paragraphs, styles };
}

export function check60Bytes() {
  const paragraphs = [];
  const styles = new Map();
  for (const name of CHECK60_PROBES) {
    const parts = probeParts(name);
    paragraphs.push(`<w:p><w:r><w:t xml:space="preserve">PROBE ${escape(name)}</w:t></w:r></w:p>`, ...parts.paragraphs.slice(1));
    for (const style of parts.styles) {
      const id = /w:styleId="([^"]+)"/.exec(style)[1];
      if (styles.has(id) && styles.get(id) !== style) throw new Error(`${name}: style ${id} differs from another probe's`);
      styles.set(id, style);
    }
  }
  const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const files = {
    '[Content_Types].xml': `${declaration}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    'word/document.xml': `${declaration}<w:document xmlns:w="${W}"><w:body>${paragraphs.join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `${declaration}<w:styles xmlns:w="${W}">${[...styles.values()].join('')}</w:styles>`,
  };
  // a fixed date, so that the file is the same bytes on every run
  const mtime = new Date('2026-10-24T00:00:00Z');
  return zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, [strToU8(v), { mtime }]])));
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  await writeFile(join(here, '60-runprops.docx'), check60Bytes());
  console.log('wrote 60-runprops.docx');
}
