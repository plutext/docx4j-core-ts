// The file for Word check 71 (CR-002 section 65: what Office JS's measurement setters do with a value past the range Word
// stores its attribute in). Run from the repository root after `npm run build`:
//
//   node test/fixtures/check71/build.mjs
//
// writes 71-measurement-setters.docx into this directory, which git ignores (Word reads the share's copy); `check71Bytes`
// gives a test the same parts. The parts are written as text, not marshalled.
//
// docx4j CR-066's round 5 measured how Word reads such a value from a file: in 16 bits, wrapped, with no error. Check 71
// asks the other side, the one an add-in sees: what each setter does when code sets it past that range (refuses, clamps,
// wraps or keeps), and what Word's save then holds. Each case is a region, `BEGIN<n> <setter> <value>`, its target
// paragraph `TARGET<n>` (in a one-cell table for columnWidth), `END<n>`. The setters:
//   leftIndent, rightIndent, firstLineIndent, spaceBefore, spaceAfter, lineSpacing   the paragraph's, in points
//   columnWidth                                                                       the table's only cell's
//   fontSize                                                                          the paragraph's font.size
//   listIndent                                                                        startNewList(), then setLevelIndents(0, value, 0)
// each at the values either side of Word's 16-bit ranges (signed 32767 twips is 1638.35 pt, unsigned 65535 is 3276.75 pt;
// for a font, 65535 half-points is 32767.5 pt), 400000 pt and negatives. Developer > Design Mode stays off.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** The setters, in order. */
export const SETTERS = ['leftIndent', 'rightIndent', 'firstLineIndent', 'spaceBefore', 'spaceAfter', 'lineSpacing', 'columnWidth', 'listIndent', 'fontSize'];
/** The values in points, either side of 32767 twips (1638.35 pt) and 65535 twips (3276.75 pt), and beyond. */
export const TWIPS_VALUES = [1638, 1638.35, 1638.4, 1639, 3276.7, 3276.75, 3276.8, 400000, -1, -1638.4, -1638.45];
/** A font's, either side of Word's UI limit (1638 pt) and of 65535 half-points (32767.5 pt). */
export const SIZE_VALUES = [1638, 1638.5, 1639, 32767.5, 32768, 400000, 0, -1];

/** Check 71's cases, in document order: { n, setter, value }. */
export function check71Cases() {
  const out = [];
  for (const setter of SETTERS) {
    for (const value of setter === 'fontSize' ? SIZE_VALUES : TWIPS_VALUES) out.push({ n: out.length + 1, setter, value });
  }
  return out;
}

/** The marker paragraph that opens a case. */
export const beginText = ({ n, setter, value }) => `BEGIN${n} ${setter} ${value}`;

const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((b) => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
const cell = (inner) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`;

/** A case's target paragraph, as built: plain, in a one-cell table for columnWidth. */
export const targetOf = (c) => (c.setter === 'columnWidth' ? cell(para(`TARGET${c.n}`)) : para(`TARGET${c.n}`));

/** A case's region: its BEGIN paragraph, its target, its END paragraph. */
const regionOf = (c) => para(beginText(c)) + targetOf(c) + para(`END${c.n}`);

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 71's file, in memory. */
export async function check71Bytes() {
  const cases = check71Cases();
  const documentXml = `${declaration}<w:document xmlns:w="${W}"><w:body>`
    + para("Check 71: what Office JS's measurement setters do with a value past the range Word stores its attribute in.")
    + cases.map(regionOf).join('') + para('End of check 71.') + `${sectPr}</w:body></w:document>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  if (!entries['word/document.xml']) throw new Error('no word/document.xml in the created package');
  entries['word/document.xml'] = strToU8(documentXml);
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check71Bytes();
  await writeFile(join(here, '71-measurement-setters.docx'), bytes);
  console.log(`wrote 71-measurement-setters.docx (${bytes.length} bytes, ${check71Cases().length} cases)`);
}
