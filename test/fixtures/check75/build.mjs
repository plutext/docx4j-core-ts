// The files for Word check 75 (CR-024, 2026-11-04; Jason: "CR-024 decisions as recommended, subject to: Word's reading of a
// non-ASCII case pair is unmeasured - let's measure it"). Run from the repository root after `npm run build`:
//
//   node test/fixtures/check75/build.mjs [directory [round]]
//
// writes the files into this directory, or the one given, which git ignores (Word reads the share's copies); `check75Files`
// gives a test the same bytes. Each is a package whose main document has one header, and whose zip holds that header's
// entry beside another whose name differs from it only in case: is the pair one name to Word (a duplicate, OPC's case-
// insensitive equivalence) or two (a part and an orphan)? The zips are written entry by entry here, since fflate refuses a
// repeated name, and each name in UTF-8 with the language encoding flag set.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';
import { storedZip } from '../../helpers.mjs';

export { storedZip };

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const A_UPPER = 'headerÉ.xml';   // É, U+00C9
const A_LOWER = 'headeré.xml';   // é, U+00E9
const E_UPPER = 'header%C3%89.xml';     // É, percent-encoded, as OPC maps a part name to a zip item name
const E_LOWER = 'header%C3%A9.xml';     // é, the same
const E_LOWER_HEX = 'header%c3%a9.xml'; // é, its hex digits in lower case

/**
 * The files: `name`, what it asks (`asks`), the header entries in zip order (`headers`: entry name and the header's text),
 * and the relationship's target (`target`), percent-encoded as a URI unless `rawTarget`.
 */
export const FILES = [
  { name: '75a-ascii-pair', asks: 'an ASCII case pair, the relationship naming the second entry',
    headers: [['headerA.xml', 'HEADER 1: the first entry, word/headerA.xml'], ['headera.xml', 'HEADER 2: the second entry, word/headera.xml, the one the relationship names']],
    target: 'headera.xml' },
  { name: '75b-nonascii-single', asks: 'one non-ASCII part name, no pair (the control: does Word find word/headeré.xml by a percent-encoded target?)',
    headers: [[A_LOWER, 'HEADER 2: the only entry, word/header-e-acute.xml, the one the relationship names']],
    target: 'header%C3%A9.xml' },
  { name: '75c-nonascii-pair', asks: 'a non-ASCII case pair (E acute and e acute), the relationship naming the second entry',
    headers: [[A_UPPER, 'HEADER 1: the first entry, word/header-E-ACUTE.xml (capital)'], [A_LOWER, 'HEADER 2: the second entry, word/header-e-acute.xml (small), the one the relationship names']],
    target: 'header%C3%A9.xml' },
  { name: '75d-nonascii-pair-reversed', asks: 'the same pair the other way round, the relationship naming the first entry',
    headers: [[A_LOWER, 'HEADER 2: the first entry, word/header-e-acute.xml (small), the one the relationship names'], [A_UPPER, 'HEADER 1: the second entry, word/header-E-ACUTE.xml (capital)']],
    target: 'header%C3%A9.xml' },
  { name: '75e-nonascii-single-raw', asks: 'one non-ASCII part name, the target written raw (a second control, in case 75b\'s percent-encoding is what Word misses)',
    headers: [[A_LOWER, 'HEADER 2: the only entry, word/header-e-acute.xml, the one the relationship names']],
    target: A_LOWER },
  // Round 2 (2026-11-04, after round 1): Word reported unreadable content for every zip entry named in raw UTF-8 (75b to
  // 75e) and dropped the header, so these name the entries as OPC maps a part name to a zip item, percent-encoded ASCII.
  { name: '75f-encoded-single', round: 2, asks: 'one non-ASCII part name, its zip entry percent-encoded (the control: does Word read word/header%C3%A9.xml?)',
    headers: [[E_LOWER, 'HEADER 2: the only entry, word/header%C3%A9.xml (e acute), the one the relationship names']],
    target: 'header%C3%A9.xml' },
  { name: '75g-encoded-pair', round: 2, asks: 'a non-ASCII case pair, percent-encoded: %C3%89 (E acute) then %C3%A9 (e acute), the relationship naming the second',
    headers: [[E_UPPER, 'HEADER 1: the first entry, word/header%C3%89.xml (E acute)'], [E_LOWER, 'HEADER 2: the second entry, word/header%C3%A9.xml (e acute), the one the relationship names']],
    target: 'header%C3%A9.xml' },
  { name: '75h-hex-case-pair', round: 2, asks: 'one character, its percent-encoding in two cases: %c3%a9 then %C3%A9, the relationship naming the second',
    headers: [[E_LOWER_HEX, 'HEADER 1: the first entry, word/header%c3%a9.xml (lower-case hex)'], [E_LOWER, 'HEADER 2: the second entry, word/header%C3%A9.xml (upper-case hex), the one the relationship names']],
    target: 'header%C3%A9.xml' },
  // Round 3 (2026-11-04, after round 2): every non-ASCII file of rounds 1 and 2 lost its header, the controls too, and no
  // file had an ASCII header alone through this writer; these separate the writer from the name (the portfolio's
  // suggestion: an ASCII header and a percent-encoded one in one zip, from one writer).
  { name: '75i-writer-and-name', round: 3, asks: 'two headers from one writer: page 1 has the percent-encoded entry word/header%C3%A9.xml, page 2 the ASCII entry word/header1.xml',
    question: 'Which header does page 1 show (HEADER E, or none), and which page 2 (HEADER A, or none)?',
    headers: [['header1.xml', 'HEADER A: page 2\'s header, the ASCII entry word/header1.xml'], [E_LOWER, 'HEADER E: page 1\'s header, the percent-encoded entry word/header%C3%A9.xml']],
    references: [{ type: 'first', target: 'header%C3%A9.xml' }, { type: 'default', target: 'header1.xml' }] },
  { name: '75j-writer-ascii-only', round: 3, asks: 'one ASCII header, word/header1.xml, through the same writer as every file before (the writer\'s control)',
    question: 'Does it open with no message, and does the page show HEADER A?',
    headers: [['header1.xml', 'HEADER A: the only entry, word/header1.xml']],
    target: 'header1.xml' },
  { name: '75k-fflate-encoded-single', round: 3, writer: 'fflate', asks: '75f again, written by fflate, the writer every saved package here uses (the name\'s control)',
    question: 'Which header does the page show: HEADER 2, or none?',
    headers: [[E_LOWER, 'HEADER 2: the only entry, word/header%C3%A9.xml (e acute), the one the relationship names']],
    target: 'header%C3%A9.xml' },
];

const p = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const partNameOf = (entry) => `/word/${entry.includes('%') ? entry : encodeURIComponent(entry.normalize('NFC'))}`;

/** The files, { name, asks, round, bytes }, built over a created package's parts; those of `round` (1 to 3), else every one. */
export async function check75Files(round) {
  const base = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  return FILES.filter((f) => round === undefined || (f.round ?? 1) === round).map((f) => {
    // one header reference per header the page shows: the file's own list, else one default header at `target`
    const references = (f.references ?? [{ type: 'default', target: f.target }]).map((r, k) => ({ ...r, id: `rIdHeader75${k === 0 ? '' : k + 1}` }));
    const types = strFromU8(base['[Content_Types].xml']).replace('</Types>',
      `${f.headers.map(([entry]) => `<Override PartName="${partNameOf(entry)}" ContentType="${HEADER_TYPE}"/>`).join('')}</Types>`);
    const rels = strFromU8(base['word/_rels/document.xml.rels']).replace('</Relationships>',
      `${references.map((r) => `<Relationship Id="${r.id}" Type="${R}/header" Target="${r.target}"/>`).join('')}</Relationships>`);
    const firstPage = references.some((r) => r.type === 'first');
    const document = `${declaration}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>`
      + p(`Check 75, ${f.name}: ${f.asks}.`) + p(f.question ?? 'Which header does the page show: HEADER 1, HEADER 2, or none?')
      + (firstPage ? '<w:p><w:r><w:br w:type="page"/><w:t>Page 2.</w:t></w:r></w:p>' : '')
      + `<w:sectPr>${references.map((r) => `<w:headerReference w:type="${r.type}" r:id="${r.id}"/>`).join('')}<w:pgSz w:w="11906" w:h="16838"/>`
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
      + `${firstPage ? '<w:titlePg/>' : ''}</w:sectPr></w:body></w:document>`;
    const entries = [];
    for (const [name, bytes] of Object.entries(base)) {
      if (name === '[Content_Types].xml') entries.push([name, strToU8(types)]);
      else if (name === 'word/_rels/document.xml.rels') entries.push([name, strToU8(rels)]);
      else if (name === 'word/document.xml') entries.push([name, strToU8(document)]);
      else entries.push([name, bytes]);
    }
    for (const [entry, text] of f.headers) entries.push([`word/${entry}`, strToU8(`${declaration}<w:hdr xmlns:w="${W}">${p(text)}</w:hdr>`)]);
    // fflate stamps each entry with the time; a fixed one makes the file the same on every build
    const bytes = f.writer === 'fflate' ? zipSync(Object.fromEntries(entries), { mtime: new Date(Date.UTC(2026, 10, 4)) }) : storedZip(entries);
    return { name: f.name, asks: f.asks, round: f.round ?? 1, bytes };
  });
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const dir = process.argv[2] ?? here;
  const files = await check75Files(process.argv[3] === undefined ? undefined : Number(process.argv[3]));
  for (const f of files) await writeFile(join(dir, `${f.name}.docx`), f.bytes);
  console.log(`wrote ${files.length} files to ${dir}`);
}
