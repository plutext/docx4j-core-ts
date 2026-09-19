# @docx4j/core-ts

docx4j-core for TypeScript: Open Packaging (docx, pptx and xlsx as zip, and the flat OPC packages
that Office JS `getOoxml()` returns), typed parts and relationships, and docx4j's style, numbering
and font resolution, over the Office Open XML object model of
[`@docx4j/generated-objects-ts`](https://github.com/plutext/docx4j-generated-objects-ts). Node,
browsers and Word add-ins.

Status: the Open Packaging layer, the typed parts and the packages (Phase A of
[CR-001](docs/change-requests/CR-001-engine.md)) are implemented, and so is `PropertyResolver`
(Phase B step 2), so paragraph and run reads report the formatting that actually applies. List
numbering and font selection (Phase B steps 3 and 4) are next.

```
npm install @docx4j/core-ts
npm install xpath        # optional, Node only: XPath over custom XML parts (browsers use document.evaluate)
```

```ts
import { WordprocessingMLPackage, ImagePart, ContentTypes } from '@docx4j/core-ts';

// a zip (Uint8Array), a flat OPC string from Office JS getOoxml(), or a PartStore
const pkg = await WordprocessingMLPackage.load(bytes);
const body = await pkg.getBody();                   // Office JS Word.Body, over the typed tree
body.insertParagraph('Hello World', 'End');
for (const range of body.search('draft', { matchWholeWord: true })) range.insertText('final', 'Replace');
body.paragraphs[0].font.bold = true;

const main = pkg.getMainDocumentPart();
const document = main.contents;                     // org_docx4j_wml.Document: the tree behind the views
const styles = await main.styleDefinitionsPart.getContents();

const image = new ImagePart('/word/media/image1.png', ContentTypes.IMAGE_PNG);
image.setBytes(pngBytes);
const rel = main.addTargetPart(image);              // rel.id is the r:embed to use in the document

const docx = await pkg.save();                      // zip: only the parts you unmarshalled are re-marshalled
const ooxml = await pkg.saveFlatOpc();              // pkg:package, for insertOoxml()
```

A part that is never touched is written back byte for byte; `getContents()` (or a view over
it) marks a part for re-marshalling. `mc:AlternateContent` is resolved when a part is
unmarshalled, as Word does on open (`{ mcePreprocess: false }` to keep it). Subpaths
`@docx4j/core-ts/opc`, `/parts`, `/packages` and `/model` give the layers separately, and
`/office-js` the `Word` shim (below).

### The content API

`Body`, `Paragraph`, `Range`, `Font`, `Table`, `InlinePicture` and `ContentControl` follow
Office JS's `Word.*` shapes (a compile-time
check keeps them assignable to a subset of those types), so add-in code runs against a
package with its `load` and `sync` lines removed. Hello World, from nothing to a `.docx`:

```ts
import { writeFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
pkg.body.insertParagraph('Hello World', 'End');
await writeFile('hello.docx', await pkg.save());
```

Editing what is there:

```ts
const body = await pkg.getBody();
const title = body.insertParagraph('Report', 'Start');
title.styleBuiltIn = 'Heading1';                       // Word.Style value; title.style reads 'Heading 1'
title.alignment = 'Centered';
const [hit] = body.search('quick brown fox');          // across runs
hit.font.italic = true;                                // runs split at the boundaries
hit.insertText('slow red fox', 'Replace');             // keeps the first replaced character's formatting
body.paragraphs.at(-1)?.insertParagraph('The end.', 'After');
body.text;                                             // a paragraph per line
```

For agents, every paragraph has an address (`body/3`, or the `w14:paraId` Word wrote), and
`outline()` lists them all:

```ts
const outline = await pkg.outline();   // { paragraphs: [{ address: 'body/0', paraId, style, text }, ...], tables, headers, footers }
const p = await pkg.paragraphAt('body/3');
p?.insertParagraph('Inserted after the fourth block', 'After');
(await pkg.paragraphAt({ contains: 'Chapter 2' }))?.delete();
```

Docx4j's names are there as aliases (`addParagraphOfText`, `addStyledParagraphOfText`,
`addObject`, `getContent`), and the tree stays reachable: `paragraph.p` is the `P`,
`body.content` the live array.

### Change tracking

`pkg.changeTrackingMode` is Office JS's `document.changeTrackingMode`, backed by
`w:trackRevisions` in the settings part. While it is on, every edit made through the content
API writes Word's revision markup instead of changing the tree in place: an insertion becomes
`w:ins`, a deletion `w:del` with `w:delText`, a replacement a `w:del` followed by a `w:ins`, a
new paragraph's mark `w:pPr/w:rPr/w:ins`, a deleted one `w:pPr/w:rPr/w:del`, and a formatting
change keeps what was there in `w:rPrChange` or `w:pPrChange`. The author is `pkg.author` and
the date `pkg.trackedChangeDate` (now, unless you fix one).

```ts
pkg.author = { name: 'Ada Lovelace' };      // the same author phase G's comments use
pkg.changeTrackingMode = 'TrackAll';           // await pkg.setChangeTrackingMode(...) on a loaded package
body.search('quick brown fox')[0].insertText('slow red fox', 'Replace');
body.replaceText('colour', 'color');           // returns the number replaced, tracked like any other edit

body.text;                                     // the accepted view: w:ins in, w:del out
body.getText({ view: 'original' });            // the other way round

for (const change of body.getTrackedChanges()) {
  change.type;        // 'Added' | 'Deleted' | 'Formatted'
  change.author; change.date; change.text;
  change.getRange();  // where it is
}
body.acceptAll();     // or rejectAll(), or accept()/reject() one at a time
```

`getTrackedChanges()` is also on `Paragraph` and `Range`. Text and searches read the accepted
view throughout, so an agent sees the document as it will read once the changes are taken.

### Comments

`getComments()` on the body, a paragraph or a range, and `insertComment` on a range or a
paragraph, over the four parts Word writes (`w:comments`, `w15:commentsEx`, `w16cid` and
`w:people`, plus `w16cex` when the document has one); threads, `resolved` (`w15:done`) and
`delete()` keep them all in step, and any of the parts the document lacks is created with its
relationship, content type and the `CommentText` and `CommentReference` styles.

```ts
pkg.author = { name: 'Ada Lovelace', initials: 'AL', email: 'ada@example.com' };   // whose comments these are

const [hit] = body.search('quick brown fox');
const comment = await hit.insertComment('Is this the right idiom?');
await comment.reply('Yes, Word writes it this way');
comment.resolved = true;                       // w15:done

for (const c of await body.getComments()) {    // document order, replies nested
  console.log(c.authorName, c.creationDate, c.content, c.replies.length, c.getRange().map((r) => r.text));
  if (c.resolved) await c.delete();            // the comment, its replies, the markers and every side entry
}
```

### XML in

A whole `word/document.xml` becomes the main document part's contents with `setXml`, and a
fragment (one or more `w:p`, `w:tbl`, ... as written inside `document.xml`, standard prefixes
declared for you) goes in through `insertXml`:

```ts
main.setXml(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p><w:sectPr/></w:body>
</w:document>`);
const body = await main.getBody();
await body.insertXml('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>From XML</w:t></w:r></w:p>', 'Start');
```

Typed content is built with the objects package's generated factories and appended with
`addObject` (docx4j) or `insertElement`, which links parent pointers and rejects what a body
cannot hold:

```ts
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
body.addObject(el.p({ content: [el.r({ content: [el.t({ value: 'Hello World' })] })] }));
```

### Tables, pictures and content controls

`Table`, `TableRow` and `TableCell` are Office JS shapes too; a cell is a `Body` of its own, so
everything above works inside one:

```ts
const table = body.insertTable(3, 2, 'End', [['Item', 'Price'], ['apples', '$20']]);
table.styleBuiltIn = 'TableGrid';           // Word.Style value; table.style is then 'Table Grid', table.styleId 'TableGrid'
table.headerRowCount = 1;                   // w:tblHeader: the row repeats on every page
table.getCell(1, 1).value = '$25';
table.addRows('End', 1, [['pears', '$30']]);
table.values;                               // [['Item', 'Price'], ['apples', '$25'], ['pears', '$30']]
body.paragraphs[4].parentTableCell?.cellIndex;
```

A picture adds its own `ImagePart`, the relationship and a `w:drawing` sized from the image's
pixels and resolution (96 dpi when the file declares none, as docx4j does), scaled down if it
is wider than the text area:

```ts
const picture = body.insertInlinePictureFromBase64(pngBase64, 'End', { altTextDescription: 'The logo' });
picture.width = 144;                        // points; the height follows the aspect ratio
picture.imageFormat;                        // 'Png'
await body.inlinePictures[0].getBase64();
```

`insertOoxml` takes what Word's does, a flat OPC `pkg:package` string: the body content comes
over with the parts it references (images, embedded objects) copied in under fresh names and
relationship ids. A bare fragment works too, as `insertXml`.

```ts
await body.insertOoxml(await other.saveFlatOpc(), 'End');       // or an add-in's getOoxml() string
```

Content controls (`w:sdt`) are read at block, row, cell and run level:

```ts
for (const control of body.contentControls) {
  console.log(control.form, control.type, control.title, control.tag, control.text);
}
body.contentControls[0].insertText('Jane Doe', 'Replace');
body.contentControls[1].delete(true);       // unwrap, keeping what it held
```

`Font` and `Paragraph` reads report **effective** formatting - the document defaults, the style
chain and the numbering level resolved, as Office JS reports it - through docx4j's
`PropertyResolver` (CR-001 Phase B step 2). Writes are always direct formatting, and
`paragraph.getFont({ direct: true })`, `range.getFont({ direct: true })` and
`paragraph.formatting({ direct: true })` read the direct values.

```ts
const p = (await pkg.getBody()).paragraphs[0];
p.font.size;                                 // 14: the style's w:sz, not the run's
p.getFont({ direct: true }).size;             // 0: the run states none
p.alignment;                                 // 'Left' when nothing states a w:jc, as Word lays it out
p.formatting({ direct: true }).alignment;     // 'Unknown'
p.effectivePPr;                               // docx4j's resolved w:pPr, for anything the view does not cover
```

### Custom XML and content controls

`pkg.customXmlParts` is Office JS's `CustomXmlPartCollection` over the custom XML data storage
parts, with XPath 1.0 over their DOM: `document.evaluate` in a browser or a Word add-in, the
optional peer dependency [`xpath`](https://www.npmjs.com/package/xpath) in Node (`npm install
xpath`; `pkg.xpathEngine` takes any other engine). One `await` warms it, and the node model is
synchronous from then on:

```ts
const parts = await pkg.customXmlParts.load();        // parses the DOMs, warms the XPath engine
const data = pkg.customXmlParts.getItem('{8B049945-9DFE-4726-9DE9-CF5691E53858}');
data.selectSingleNode('/invoice/customer/name').text;         // 'Joe Bloggs'
data.selectNodes('/invoice/items/item').map((n) => n.text);
data.documentElement.appendChildNode('<note>paid</note>');
```

A content control's `xmlMapping` is its `w:dataBinding`, and the two directions are docx4j's,
under docx4j's names:

```ts
const control = (await pkg.getBody()).contentControls[0];
control.xmlMapping.isMapped;                          // true
control.xmlMapping.xpath;                             // '/invoice[1]/customer[1]/name[1]'
control.xmlMapping.customXmlNode.text;                // 'Joe Bloggs'

await pkg.customXmlParts.applyBindings();             // docx4j BindingHandler: XML -> controls
control.insertText('Jane Doe', 'Replace');            // writes through to the bound node as well
await pkg.customXmlParts.updateFromContentControls(); // docx4j: controls -> XML
```

That last pair matters: Word refreshes a bound control from the custom XML part when it opens a
document, so an edit that only touched `w:sdtContent` would not show.

A new part comes with its properties part, a fresh `ds:itemID` and the relationship from the main
document part (Word drops a custom XML part the main part does not relate to):

```ts
const part = pkg.customXmlParts.add('<greeting xmlns="http://example.com/g"><to>World</to></greeting>');
control.xmlMapping.setMapping('/ns0:greeting[1]/ns0:to[1]', "xmlns:ns0='http://example.com/g'", part);
control.xmlMapping.setMappingByNode(part.selectSingleNode('/ns0:greeting/ns0:to', "xmlns:ns0='http://example.com/g'"));
```

Controls are created by wrapping what is there, and the kind-specific views are Office JS's:

```ts
const control = paragraph.insertContentControl('CheckBox');   // also on body and range
control.checkboxContentControl.isChecked = true;              // shows ☒, as Word does
control.title = 'Applies';  control.appearance = 'Tags';  control.cannotDelete = true;

const list = body.paragraphs[1].insertContentControl('DropDownList').dropDownListContentControl;
list.addListItem('Apples', 'apples');
```

### Use in Node: running add-in code against a package

`@docx4j/core-ts/office-js` is a `Word` shim: `Word.run(pkg, fn)` gives the callback a
`context` whose `document.body` is the package's body, `load()` is a no-op and `context.sync()`
resolves what was asynchronous. An add-in's batch therefore runs unchanged in a test, in CI,
or on a server, with the saved `.docx` as the assertion — and a member this package does not
implement throws `NotSupportedError` naming it, rather than doing nothing.

```ts
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { Word } from '@docx4j/core-ts/office-js';

const pkg = await WordprocessingMLPackage.load(bytes);
await Word.run(pkg, async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load('text');
  await context.sync();

  paragraphs.items[0].styleBuiltIn = Word.Style.heading1;
  for (const range of context.document.body.search('draft')) range.font.highlightColor = '#FFFF00';
  context.document.properties.lastAuthor = 'an add-in';
  context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

  const ooxml = context.document.body.getOoxml();   // a ClientResult, as in Office JS
  await context.sync();
  ooxml.value;                                      // the flat OPC package
});
await writeFile('edited.docx', await pkg.save());
```

`Word.supported` is the set of `Class.member` strings this package implements
(`Word.supported.has('Body.insertParagraph')`), generated from the compile-time subset, so a
tool can check a script before running it. The selection an add-in edits is passed in:
`Word.run(pkg, fn, { selection })`.

`toApiScript` goes the other way: given a paragraph, a range or an element, it emits the
content-API calls that reproduce it, falling back to `insertXml` for what the verbs cannot say.

```ts
import { toApiScript } from '@docx4j/core-ts/model';

await toApiScript(body.paragraphs[1]);
// const p1 = body.insertParagraph('Chapter 1', 'End');
// p1.styleBuiltIn = 'Heading1';
// p1.alignment = 'Centered';
// const r1 = p1.insertText(' (draft)', 'End');
// r1.font.italic = true;
```

Custom XML parts, XML mapping, typed content controls and lists are the phases still to come of
[CR-002](docs/change-requests/CR-002-content-api.md). A generated script reproduces the
document's *markup*, so it emits direct formatting only, not the effective values the reads
report (CR-001 Phase B step 2).

## Development

```
npm ci                              # the locked dependencies from npm
npm run typecheck && npm test       # test = build, then node --test test/*.test.mjs
npm run generate                    # src/office-js/supported.generated.mts from test/office-js-subset.ts
```

`npm run generate` is committed output: the build does not depend on the script, but `npm test`
regenerates it so that `Word.supported` stays in step with the Office JS subset.

Releases are published to npm from GitHub Actions; see `RELEASING.md`.

## Licence

Apache-2.0, as docx4j. See NOTICE.
