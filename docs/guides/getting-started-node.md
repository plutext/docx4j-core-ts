# Getting started in Node

Load a `.docx`, read and edit it through the content API, save it. Everything on this page runs
on Node 18 or later; the same code runs in a browser and in a Word add-in (see
[office-addin.md](office-addin.md)).

```
npm install @docx4j/core-ts
npm install xpath        # optional, Node only: XPath over custom XML parts
```

## Load and save

```js
import { readFile, writeFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile('in.docx')));
// ... edits ...
await writeFile('out.docx', await pkg.save());
```

`load` takes zip bytes, a flat OPC `pkg:package` string (what Office JS `getOoxml()` returns), or
a `PartStore`. `save()` gives a zip, `saveFlatOpc()` a `pkg:package` string, `saveTo(sink)`
anything else — including a directory:

```js
import { DirectoryPartStore, DirectoryPartSink } from '@docx4j/core-ts/node';

await pkg.saveTo(new DirectoryPartSink('unzipped'));               // one file per part
const same = await WordprocessingMLPackage.load(await DirectoryPartStore.open('unzipped'));
```

**A part you never touch is written back byte for byte.** Only the parts whose content you read
as objects are re-marshalled, which is what makes a round trip safe with markup the object model
does not know, and what keeps saving fast.

## The content API

`pkg.body` is a `Word.Body` in Office JS's shape — `paragraphs`, `insertParagraph`, `search`,
`font`, `tables`, `contentControls` — over the typed tree. The part has to be unmarshalled first,
which `getBody()` does; `body` is the synchronous accessor afterwards.

```js
const body = await pkg.getBody();

body.text;                                        // the whole document as text
body.paragraphs.length;

const p = body.insertParagraph('Chapter 1', 'End');
p.styleBuiltIn = 'Heading1';
p.alignment = 'Centered';

const run = p.insertText(' (draft)', 'End');
run.font.italic = true;
run.font.highlightColor = '#FFFF00';

for (const range of body.search('draft', { matchWholeWord: true })) {
  range.insertText('final', 'Replace');
}
```

Reads report the formatting that **actually applies**, resolved through the style hierarchy, the
document defaults and the theme, as Word paints it — not just the direct formatting in the
markup. `font.name`, `font.size`, `paragraph.alignment` and the indents all work this way; pass
`{ direct: true }` where you want the markup's own value.

```js
await pkg.getMainDocumentPart().getRunFontSelector();  // once: reads the theme and the font table
body.paragraphs[0].font.name;                          // 'Aptos Display', resolved through the theme
body.paragraphs[0].font.size;                          // 14, from the Heading 1 style
body.paragraphs[0].getFont({ direct: true }).size;     // 0: the paragraph itself states no size
```

## Where things are

```js
const outline = await pkg.outline();              // every paragraph and table, with an address
outline.paragraphs[3];                            // { address: 'body/3', style, text, paraId }
const p = await pkg.paragraphAt('body/3');        // and back again
await pkg.paragraphAt('w14:1F2A3B4C');            // or by Word's own paragraph id
```

## Lists, tracked changes and comments

```js
const items = body.paragraphs.filter((p) => p.isListItem);
items[0].listItem.listString;                     // '1.', 'a.', a bullet — the label Word paints

pkg.author = { name: 'Jane Doe', initials: 'JD' };
pkg.changeTrackingMode = 'TrackAll';              // every edit below is now a revision
body.paragraphs[0].insertText('tracked', 'End');
pkg.getTrackedChanges().map((c) => `${c.type} by ${c.author}`);

const comments = await body.getComments();
comments[0].content;
body.paragraphs[0].getRange().insertComment('Check this');
```

## The parts underneath

The content API is a view; the package, its parts and their relationships are the model, named as
in docx4j so Java code and documentation transfer.

```js
import { ImagePart, ContentTypes } from '@docx4j/core-ts';

const main = pkg.getMainDocumentPart();
main.contents;                                    // org_docx4j_wml.Document, the tree behind the views
const styles = await main.styleDefinitionsPart.getContents();

const image = new ImagePart('/word/media/image1.png', ContentTypes.IMAGE_PNG);
image.setBytes(pngBytes);
const rel = main.addTargetPart(image);            // rel.id is the r:embed to use in the document
```

`pkg.parts` is every part, `pkg.getPart('/word/document.xml')` one of them, and
`part.relationshipsPart` its `.rels`. A `pkg.clone()` is an independent copy of the whole thing.

## What holds it to docx4j

Every effective property, every list label and every resolved font on 45 documents is compared,
in the test suite, against what docx4j itself answers for the same document — goldens produced by
a Java harness against docx4j `VERSION_17_1_1`. That is the promise the name makes; see
[parity.md](parity.md) for how to run it yourself against a newer docx4j.

## Next

- [office-addin.md](office-addin.md) — the same code inside Word, over flat OPC
- [presentationml-spreadsheetml.md](presentationml-spreadsheetml.md) — pptx and xlsx
- [`examples/node/`](../../examples/node) — runnable versions of everything above
