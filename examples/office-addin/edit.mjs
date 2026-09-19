// The edit itself, written once and run in two places: inside Word, over the flat OPC package
// that `getOoxml()` hands out, and in Node, over the same package through the `./office-js` shim.
// Nothing here knows which: `body` is a `Word.Body`, Office JS's or this package's.
//
// In a real add-in the import is `from '@docx4j/core-ts'`; here it is the built dist of this
// checkout, so the example runs from the repository without an install.
import { WordprocessingMLPackage } from '../../dist/index.mjs';

/**
 * One edit, on any `Word.Body`: the first paragraph becomes a Heading 1, and every occurrence of
 * `term` is highlighted and struck through. Returns how many it highlighted.
 *
 * @param {import('../../dist/model/content/Body.mjs').Body} body
 * @param {string} [term]
 * @returns {number}
 */
export function edit(body, term = 'draft') {
  const first = body.paragraphs[0];
  if (first) first.styleBuiltIn = 'Heading1';
  let count = 0;
  for (const range of body.search(term, { matchCase: false })) {
    range.font.highlightColor = '#FFFF00';
    range.font.strikeThrough = true;
    count++;
  }
  return count;
}

/**
 * The add-in's half: a flat OPC package in (what `Range.getOoxml()` returns), the edited flat OPC
 * package out (what `Range.insertOoxml()` takes).
 *
 * @param {string} ooxml
 * @param {string} [term]
 * @returns {Promise<string>}
 */
export async function applyEdit(ooxml, term) {
  const pkg = await WordprocessingMLPackage.load(ooxml);
  edit(await pkg.getBody(), term);
  return pkg.saveFlatOpc();
}
