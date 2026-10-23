// The file for Word check 59 (2026-10-23, CR-002 section 32's K5): where Word's wildcards match a paragraph mark. Check 58
// found `?` matching a paragraph's mark at its end, in one body paragraph; this measures the places `^p` needed care
// (CR-014 phase 3): across two paragraphs, a cell's end, a row's end, and a mark inside a block content control. Read
// only. Run from the repository root after `npm run build`:
//
//   node test/fixtures/check59/build.mjs
//
// writes 59-wildcard-marks.docx into this directory, which git ignores (Word reads the share's copy); `check59Bytes`
// gives the same bytes to a test. The parts are written as text, not marshalled.
//
// Each case is a region: a paragraph `BEGIN<n> <shape>`, its content, a paragraph `END<n>`. The content's paragraphs
// hold one or two letters each (Q, Z, W), so that a search for `?` lists every character and mark Word matches.
// check59-script-lab.js searches each region with its queries in four scopes: the region's range, each paragraph,
// the first table's first cell's body, the first content control.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const ROOT_NS = ` xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;

const t = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (text) => `<w:p>${t(text)}</w:p>`;
const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((b) => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
const tblPr = `<w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}<w:tblLook w:val="04A0"/></w:tblPr>`;
const tc = (...ps) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${ps.map(para).join('')}</w:tc>`;
const tbl = (cols, ...rows) => `<w:tbl>${tblPr}<w:tblGrid>${'<w:gridCol w:w="2400"/>'.repeat(cols)}</w:tblGrid>${rows.map((r) => `<w:tr>${r.join('')}</w:tr>`).join('')}</w:tbl>`;
let controlId = 59000;
const sdt = (...ps) => `<w:sdt><w:sdtPr><w:alias w:val="Check 59"/><w:tag w:val="c59-${controlId}"/><w:id w:val="${controlId++}"/></w:sdtPr><w:sdtContent>${ps.map(para).join('')}</w:sdtContent></w:sdt>`;

/** The cases: [shape, content]. */
const CASES = [
  ['body', () => [para('Q'), para('Z')]],                                          // two body paragraphs: the mark between
  ['cell-only', () => [tbl(1, [tc('Q')]), para('Z')]],                            // a cell's only paragraph, then the cell's end
  ['cell-two', () => [tbl(1, [tc('Q', 'Z')]), para('W')]],                        // a mark inside a cell, then its end
  ['cells-rows', () => [tbl(2, [tc('Q'), tc('Z')], [tc('W'), tc('X')]), para('Y')]], // across a cell's end, a row's end
  ['control', () => [sdt('Q', 'Z'), para('W')]],                                  // a mark inside a block control, its last
];

/** Check 59's cases, in document order: { n, shape, content }. */
export function check59Cases() {
  controlId = 59000;
  return CASES.map(([shape, content], i) => ({ n: i + 1, shape, content: content() }));
}

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** Check 59's file, in memory. */
export async function check59Bytes() {
  const cases = check59Cases();
  const regions = cases.flatMap((c) => [para(`BEGIN${c.n} ${c.shape}`), ...c.content, para(`END${c.n}`)]);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document${ROOT_NS}><w:body>`
    + para('Check 59. Each case is a region: its BEGIN paragraph, its content, its END paragraph. Read only.') + regions.join('')
    + para('End of check 59.') + `${sectPr}</w:body></w:document>`;
  const pkg = await WordprocessingMLPackage.createPackage();
  const entries = unzipSync(await pkg.save());
  if (!entries['word/document.xml']) throw new Error('no word/document.xml in the created package');
  entries['word/document.xml'] = strToU8(documentXml);
  void strFromU8;
  const bytes = zipSync(entries);
  await WordprocessingMLPackage.load(bytes);   // it loads here, at least
  return bytes;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = await check59Bytes();
  await writeFile(join(here, '59-wildcard-marks.docx'), bytes);
  console.log(`wrote 59-wildcard-marks.docx (${bytes.length} bytes, ${check59Cases().length} cases)`);
}
