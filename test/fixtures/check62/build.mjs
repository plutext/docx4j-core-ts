// The file for Word check 62 (CR-015 section 12.2: Word's rule for the highlight a written colour becomes). Run from the
// repository root:
//
//   node test/fixtures/check62/build.mjs
//
// writes 62-highlight-snap.docx into this directory, which git ignores (Word reads the share's copy); `check62Bytes`
// gives the same bytes to a test. A title paragraph, then one paragraph per colour, its text the colour's six hex digits
// (CHECK62_POINTS, each once, in the order of their families):
// - the grid: 16 levels per channel (00, 11, ..., FF), 4,096 colours, for the overall shape;
// - the grey ramp, 000000 to FFFFFF in steps of 1, for where black, darkGray, lightGray and white meet;
// - each channel's ramp from 00 to FF with the other two at 00, and again with them at FF;
// - check 61's two counterexamples to nearest-in-RGB, #A9A9A9 (Word: darkGray) and #123456 (Word: darkCyan), each
//   channel within 8 of them in steps of 2.
// The objects session (its CR-014's H4) tests candidate rules against Word's answers.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, strToU8 } from 'fflate';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const hex = (r, g, b) => [r, g, b].map((v) => v.toString(16).toUpperCase().padStart(2, '0')).join('');

/** The colours, by family; a colour in more than one is kept in the first. */
export const CHECK62_FAMILIES = (() => {
  const grid = [];
  for (let r = 0; r < 256; r += 17) for (let g = 0; g < 256; g += 17) for (let b = 0; b < 256; b += 17) grid.push(hex(r, g, b));
  const grey = Array.from({ length: 256 }, (_, v) => hex(v, v, v));
  const ramps = (other) => [0, 1, 2].flatMap((channel) => Array.from({ length: 256 }, (_, v) => {
    const c = [other, other, other];
    c[channel] = v;
    return hex(...c);
  }));
  const around = ([r0, g0, b0]) => {
    const out = [];
    for (let dr = -8; dr <= 8; dr += 2) for (let dg = -8; dg <= 8; dg += 2) for (let db = -8; db <= 8; db += 2) out.push(hex(r0 + dr, g0 + dg, b0 + db));
    return out;
  };
  return { grid, grey, rampsOn00: ramps(0), rampsOnFF: ramps(255), aroundA9A9A9: around([0xA9, 0xA9, 0xA9]), around123456: around([0x12, 0x34, 0x56]) };
})();

export const CHECK62_POINTS = [...new Set(Object.values(CHECK62_FAMILIES).flat())];

export function check62Bytes(points = CHECK62_POINTS) {
  const p = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const paragraphs = [p('Check 62: the highlight Word stores for a colour written to it (CR-015 section 12.2)'), ...points.map(p)];
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
  await writeFile(join(here, '62-highlight-snap.docx'), check62Bytes());
  console.log(`wrote 62-highlight-snap.docx: ${CHECK62_POINTS.length} colours`);
}
