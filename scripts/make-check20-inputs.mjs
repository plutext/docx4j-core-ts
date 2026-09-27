// The hand-edited inputs of test/README.md check 20, case 4 (CR-002 section 29): a move with a half
// missing its partner or its range markers, made from the Word-made
// test/fixtures/revisions/revisions-word15.docx by editing word/document.xml as text, so that every
// other byte of every part is Word's. Section 3 of the fixture is the move.
//
//   node scripts/make-check20-inputs.mjs
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { unzipSync, zipSync } from 'fflate';

const dir = new URL('../test/fixtures/revisions/', import.meta.url);
const out = new URL('check20/input/', dir);
const source = unzipSync(new Uint8Array(await readFile(new URL('revisions-word15.docx', dir))));
const document = new TextDecoder().decode(source['word/document.xml']);

/** The text with the first match of each pattern removed; throws when one is missing. */
function without(text, ...patterns) {
  for (const pattern of patterns) {
    const match = pattern.exec(text);
    if (!match) throw new Error(`not found: ${pattern}`);
    text = text.slice(0, match.index) + text.slice(match.index + match[0].length);
  }
  return text;
}

/** The w:p holding the given marker, whole. */
const paragraphHolding = (text, marker) => {
  const at = text.indexOf(marker);
  if (at < 0) throw new Error(`not found: ${marker}`);
  const start = text.lastIndexOf('<w:p ', at);
  const end = text.indexOf('</w:p>', at) + '</w:p>'.length;
  return new RegExp(text.slice(start, end).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
};

const inputs = {
  // (a) the destination half alone: the source paragraph (its w:moveFrom, its range start and its
  // mark) and the source range end are gone, so the move's w:name names no other side
  'partner-missing.docx': without(document, paragraphHolding(document, '<w:moveFromRangeStart'), /<w:moveFromRangeEnd w:id="\d+"\/>/),
  // (b) a w:moveTo with no range markers: nothing names its move
  'moveTo-no-ranges.docx': without(document, /<w:moveToRangeStart [^>]*\/>/, /<w:moveToRangeEnd w:id="\d+"\/>/),
  // (c) a w:moveFrom with no range markers
  'moveFrom-no-ranges.docx': without(document, /<w:moveFromRangeStart [^>]*\/>/, /<w:moveFromRangeEnd w:id="\d+"\/>/),
};

await mkdir(out, { recursive: true });
for (const [name, xml] of Object.entries(inputs)) {
  const entries = { ...source, 'word/document.xml': new TextEncoder().encode(xml) };
  await writeFile(new URL(name, out), zipSync(entries, { level: 6 }));
  console.log(name, xml.length - document.length, 'characters');
}
