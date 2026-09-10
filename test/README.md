# Tests

`npm test` builds and runs every `test/*.test.mjs` with Node's test runner. One file:
`node --test test/roundtrip.test.mjs` (after `npm run build`).

Fixtures under `fixtures/` come from docx4j's `docx4j-core-tests` resources (Apache-2.0):
`loadAndSave.docx` (a current Word: comments, headers and footers, a chart with an embedded
workbook, PNG and SVG images, custom XML, `mc:AlternateContent`), `HelloWordOnline.docx` (main
part named `document22.xml`), `header-no-rels.docx`, `hyperlink_dupe.docx`,
`loadAndSave.pptx`, `loadAndSave.xlsx`, and `mc-alternate-content-header.xml`, a Word 2010 flat
OPC package (`pkg:package`) with `mc:AlternateContent` in a header.

## Word acceptance (manual)

Saved output must open in Word without a repair prompt. After a change to marshalling,
namespaces, content types or the zip writer, check by hand:

1. Round trip: load `fixtures/loadAndSave.docx`, save, open in Word. No repair prompt; comments,
   headers, footers, the chart and both images still there.
2. Re-marshalled main part: as 1, but `await pkg.getMainDocumentPart().getContents()` before
   saving. Word must accept the `mc:Ignorable` prefixes and the `xml:space` attributes.
3. New document: `WordprocessingMLPackage.createPackage()` with a paragraph added, saved, opened.
   Styles pane shows Normal and Heading 1 to 4.
4. Flat OPC: `saveFlatOpc()` output pasted through `insertOoxml()` in a Word add-in (or saved as
   `.xml` and opened in Word, which reads `pkg:package` files directly).
5. pptx and xlsx: round trips of the two fixtures open in PowerPoint and Excel.

A small Node script for 1 to 3 is:

```js
import { writeFile, readFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
const pkg = await WordprocessingMLPackage.load(await readFile('test/fixtures/loadAndSave.docx'));
await pkg.getMainDocumentPart().getContents();
await writeFile('out.docx', await pkg.save());
```
