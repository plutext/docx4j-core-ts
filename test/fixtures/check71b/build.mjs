// The file for Word check 71b (CR-002 section 65.5 D2: where exactly Word 365's own setters refuse, below storage). Run
// from the repository root after `npm run build`:
//
//   node test/fixtures/check71b/build.mjs
//
// writes 71b-setter-limits.docx into this directory, which git ignores (Word reads the share's copy); `check71bBytes`
// gives a test the same parts. The parts are written as text, not marshalled. The snippet is check 71's, unchanged
// (check71b-script-lab.js is a copy of check71-script-lab.js, held equal by measurements.test.mjs): it parses each
// region's `BEGIN<n> <setter> <value>` and sets the setter, so only the file differs.
//
// Check 71 found Word's spacing, line, list and font setters refusing with InvalidArgument below what Word stores:
// spaceBefore, spaceAfter and lineSpacing refused 1638 pt and every value above, and negatives; setLevelIndents refused a
// text indent of 1638 pt and up and -1638.4 and down, but took -1; font.size took 1638 and refused 1638.5, 1639 and 0.
// Between those values the limit was not measured: 1584 pt (22 inches) is Word's UI limit for spacing and indents, 1 to
// 1638 pt for a font, and an exact line has a smallest value too. This check sets each setter at and either side of
// those boundaries, so this package can refuse where Word does (D2), and no earlier.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';
import { beginText, targetOf } from '../check71/build.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** The values per setter, in points: at and either side of the boundaries check 71 left between its values. */
export const VALUES = {
  // 22 inches is 1584 pt; check 71 refused 1638 and -1 and took nothing in between
  spaceBefore: [0, 0.05, 1583.95, 1584, 1584.05, 1584.1, 1585, 1600, 1637.95, -0.05],
  spaceAfter: [0, 0.05, 1583.95, 1584, 1584.05, 1584.1, 1585, 1600, 1637.95, -0.05],
  // an exact line: Word's dialog takes 0.7 pt and up; and the upper limit as the spacing's
  lineSpacing: [0, 0.05, 0.5, 0.65, 0.7, 0.75, 1, 1583.95, 1584, 1584.05, 1585, 1637.95],
  // setLevelIndents(0, value, 0): check 71 refused 1638 and -1638.4, took -1
  listIndent: [1583.95, 1584, 1584.05, 1585, 1637.95, -1583.95, -1584, -1584.05, -1585, -1638.35],
  // font.size: check 71 took 1638 and refused 1638.5 and 0; Word's dialog takes 1 to 1638
  fontSize: [0.5, 0.9, 1, 1.5, 1637.5, 1638, 1638.05, 1638.4, 1638.45],
};

/** Check 71b's cases, in document order: { n, setter, value }, check 71's shape. */
export function check71bCases() {
  const out = [];
  for (const [setter, values] of Object.entries(VALUES)) for (const value of values) out.push({ n: out.length + 1, setter, value });
  return out;
}

/** A case's region: its BEGIN paragraph, its target, its END paragraph (check 71's). */
const regionOf = (c) => para(beginText(c)) + targetOf(c) + para(`END${c.n}`);

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 71b's file, in memory. */
export async function check71bBytes() {
  const cases = check71bCases();
  const documentXml = `${declaration}<w:document xmlns:w="${W}"><w:body>`
    + para("Check 71b: where Word 365's own spacing, line, list and font setters refuse, below what Word stores.")
    + cases.map(regionOf).join('') + para('End of check 71b.') + `${sectPr}</w:body></w:document>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  if (!entries['word/document.xml']) throw new Error('no word/document.xml in the created package');
  entries['word/document.xml'] = strToU8(documentXml);
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check71bBytes();
  await writeFile(join(here, '71b-setter-limits.docx'), bytes);
  console.log(`wrote 71b-setter-limits.docx (${bytes.length} bytes, ${check71bCases().length} cases)`);
}
