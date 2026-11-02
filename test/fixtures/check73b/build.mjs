// Check 73b's one built file (test/README.md check 73b): docx4j's bind-plain-two-parts.docx with the main document
// part's two customXml relationships listed in the other order, nothing else changed, so that Word's order of the
// custom XML parts, and the part a stale B7 binds to, can be seen to follow that order or not.
//   node test/fixtures/check73b/build.mjs <docx4j's bind-plain-two-parts.docx> <out.docx>
import { readFileSync, writeFileSync } from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

/** The probe's bytes with its document.xml.rels relationships in reverse order. */
export function swappedRelationships(bytes) {
  const entries = unzipSync(bytes);
  const rels = strFromU8(entries['word/_rels/document.xml.rels']);
  const items = rels.match(/<Relationship\b[^>]*\/>/g) ?? [];
  if (items.length !== 2) throw new Error(`expected two relationships, found ${items.length}`);
  entries['word/_rels/document.xml.rels'] = strToU8(rels.replace(items[0], '\u0000').replace(items[1], items[0]).replace('\u0000', items[1]));
  return zipSync(entries);
}

const [source, out] = process.argv.slice(2);
if (process.argv[1]?.endsWith('build.mjs') && source && out) writeFileSync(out, swappedRelationships(new Uint8Array(readFileSync(source))));
