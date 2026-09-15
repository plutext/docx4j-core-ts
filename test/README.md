# Tests

`npm test` builds and runs every `test/*.test.mjs` with Node's test runner. One file:
`node --test test/roundtrip.test.mjs` (after `npm run build`).

`office-js-subset.ts` is the compile-time Office JS promise (CR-002 section 3.4), compiled by
`npm run typecheck`; `npm run generate` turns it into `src/office-js/supported.generated.mts`,
which is `Word.supported`. Add a member to the subset and regenerate in the same commit.
`office-js.test.mjs` covers the `Word` shim and `toApiScript` (CR-002 phase I).

Fixtures under `fixtures/` come from docx4j's `docx4j-core-tests` resources (Apache-2.0):
`loadAndSave.docx` (a current Word: comments, headers and footers, a chart with an embedded
workbook, PNG and SVG images, custom XML, `mc:AlternateContent`), `HelloWordOnline.docx` (main
part named `document22.xml`), `header-no-rels.docx`, `hyperlink_dupe.docx`,
`loadAndSave.pptx`, `loadAndSave.xlsx`, `invoice.docx` (from `OpenDoPE/`: content controls at
block, row and cell level, with data bindings and a table, the fixture of the content-control
tests), `comments-two.docx` (from `AlteredParts/`: two comments with no w15, w16cid or
`w:people` parts, so the comment API must cope without them), and
`mc-alternate-content-header.xml`, a Word 2010 flat OPC package (`pkg:package`) with
`mc:AlternateContent` in a header.

The images in `content-c.test.mjs` are base64 constants rather than files: a 4 x 3 PNG at 96 dpi
(a real one, deflated with `node:zlib`), a 2 x 2 GIF87a, a 2 x 2 24-bit BMP at 3780 px/m, and a
JPEG header (SOI, a JFIF APP0 at 300 dpi, SOF0, EOI) for the header reader.

## Word acceptance (manual)

Last run: 2026-09-16, Word 365 version 2608 (build 20326.20144, Click-to-Run), after CR-002
phases C, G and I: checks 6 to 8, with the script below, all passed. Check 7 as expected shows the fixture's own text, not the
replacement, on the document surface: the control is bound to a custom XML part, and Word
re-reads a bound control from the XML part on open, so the replaced `w:sdtContent` is
overwritten by the stale binding. The control, its title and its binding survived, which is
what the check is for; writing through the binding is CR-002 phase E (`XmlMapping`).

Previous run: 2026-09-10, Word 2016, after CR-001 Phase A and CR-002 phases B and D. All four
files below opened without a repair prompt: the untouched round trip, the round trip with the
main part and a header re-marshalled and a paragraph added through the content API, a
document created from nothing with a heading and a formatted span, and the flat OPC of the
second opened as a `pkg:package` file.

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
6. Tables and pictures (CR-002 phase C): a package created from nothing with
   `body.insertTable(3, 2, 'End', values)`, `table.styleBuiltIn = 'TableGrid'` and
   `table.headerRowCount = 1`, then `body.insertInlinePictureFromBase64(png, 'End')`. Word must
   show a bordered table whose first row repeats on a page break, and the picture at its natural
   size (4 x 3 px at 96 dpi is 3 x 2.25 pt; use a larger image to judge it), with the alt text in
   Format Picture. Check `insertOoxml` of that package's `saveFlatOpc()` into another: the image
   appears once, from `word/media/image1.png` of the target.
7. Content controls: `fixtures/invoice.docx` loaded, `body.contentControls[0].insertText(...)`,
   saved; Word must still show the control and its data binding. Expect the fixture's text, not
   the replacement, on the surface: the control is bound, and Word refreshes a bound control from
   the custom XML part on open (the replaced text is in `word/document.xml`, and is what a
   non-bound control would show). Until phase E writes through the binding, check the XML.
8. Comments (CR-002 phase G): in a package from `createPackage()`, comment a range
   (`range.insertComment('...')`), reply to it and resolve it; in `fixtures/loadAndSave.docx`,
   comment a paragraph and delete Word's own comment. Saved and opened in Word: the review pane
   shows the comment, the reply under it in the same thread, the resolved thread greyed, the
   author and initials from `pkg.author`, and the deleted comment gone. Note that a comments part
   this package re-marshals loses the `mc:Ignorable` attribute of `w:comments` (the object model
   does not carry it, CR-002 section 9), so this check is the one that would catch a repair prompt
   from that.

A small Node script for 1 to 3 is:

```js
import { writeFile, readFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
const pkg = await WordprocessingMLPackage.load(await readFile('test/fixtures/loadAndSave.docx'));
await pkg.getMainDocumentPart().getContents();
await writeFile('out.docx', await pkg.save());
```

And one for 6 to 8 (run from the repository root after `npm run build`; `logo.png` is any PNG):

```js
import { writeFile, readFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from './dist/index.mjs';

// 6. tables and pictures
const a = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
const rows = [['Item', 'Price'], ...Array.from({ length: 60 }, (_, i) => [`item ${i + 1}`, `$${i + 1}`])];  // spans a page break
const table = a.body.insertTable(rows.length, 2, 'End', rows);
table.styleBuiltIn = 'TableGrid';
table.headerRowCount = 1;
a.body.insertInlinePictureFromBase64((await readFile('logo.png')).toString('base64'), 'End', { altTextDescription: 'The logo' });
await writeFile('check6-a.docx', await a.save());
const b = await WordprocessingMLPackage.createPackage();
await b.body.insertOoxml(await a.saveFlatOpc(), 'End');
await writeFile('check6-b.docx', await b.save());           // the image once, as word/media/image1.png

// 7. content controls
const c = await WordprocessingMLPackage.load(await readFile('test/fixtures/invoice.docx'));
(await c.getBody()).contentControls[0].insertText('Jane Doe', 'Replace');
await writeFile('check7.docx', await c.save());

// 8. comments
const d = await WordprocessingMLPackage.createPackage();
d.author = { name: 'Acceptance Tester', initials: 'AT' };
const p = d.body.insertParagraph('The quick brown fox jumps over the lazy dog.', 'End');
const [range] = p.search('brown fox');
const comment = await range.insertComment('Is it brown?');
await comment.reply('Yes, brown.');
comment.resolved = true;
await writeFile('check8-a.docx', await d.save());
const e = await WordprocessingMLPackage.load(await readFile('test/fixtures/loadAndSave.docx'));
const body = await e.getBody();
await body.paragraphs[0].insertComment('A new comment on the first paragraph');
for (const existing of await body.getComments()) if (existing.authorName !== 'docx4j') await existing.delete();
await writeFile('check8-b.docx', await e.save());
```
