// The file for Word check 61 (CR-015 section 12: Word's highlight palette through Office JS, read and written).
// Run from the repository root:
//
//   node test/fixtures/check61/build.mjs
//
// writes 61-highlights.docx into this directory, which git ignores (Word reads the share's copy); `check61Bytes` gives
// the same bytes to a test. The parts are written as text. The cases are the snippet's (check61-script-lab.js, loaded
// here as a test loads it): a `READ <name>: R_<name> (end)` paragraph for each of w:highlight's sixteen names, `none`
// and a character shading, then a `WRITE <n> <value>: W<nn> (end)` paragraph for each value the snippet writes (the
// marker's number two digits, so that no marker holds another).
import { readFileSync } from 'node:fs';
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, strToU8 } from 'fflate';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export const CHECK61 = new Function(`${readFileSync(join(here, 'check61-script-lab.js'), 'utf8')}\nreturn CHECK61;`)();

const t = (text, rPr = '') => `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;
/** A probe paragraph: `label: MARKER (end)`, only the marker run carrying `rPr`. */
const probe = (label, marker, rPr) => `<w:p>${t(`${label}: `)}${t(marker, rPr)}${t(' (end)')}</w:p>`;

/** The read cases: [label, marker, the marker's w:rPr]. */
export const CHECK61_READS = [
  ...CHECK61.NAMES.map((name) => [`READ ${name}`, `R_${name}`, `<w:highlight w:val="${name}"/>`]),
  ['READ none', 'R_none', '<w:highlight w:val="none"/>'],
  ['READ shd', 'R_SHD', '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>'],
];

export function check61Bytes() {
  const paragraphs = [`<w:p>${t('Check 61: highlight colours, read and written (CR-015 section 12)')}</w:p>`];
  for (const [label, marker, rPr] of CHECK61_READS) paragraphs.push(probe(label, marker, rPr));
  CHECK61.WRITES.forEach((value, n) => {
    paragraphs.push(probe(`WRITE ${n} ${value === null ? 'null' : value}`, `W${String(n).padStart(2, '0')}`, value === null ? '<w:highlight w:val="yellow"/>' : ''));
  });
  const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const files = {
    '[Content_Types].xml': `${declaration}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    '_rels/.rels': `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/document.xml': `${declaration}<w:document xmlns:w="${W}"><w:body>${paragraphs.join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  const mtime = new Date('2026-10-24T00:00:00Z');
  return zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, [strToU8(v), { mtime }]])));
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  await writeFile(join(here, '61-highlights.docx'), check61Bytes());
  console.log('wrote 61-highlights.docx');
}
