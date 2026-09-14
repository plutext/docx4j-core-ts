# @docx4j/core-ts

docx4j-core for TypeScript: Open Packaging (docx, pptx and xlsx as zip, and the flat OPC packages
that Office JS `getOoxml()` returns), typed parts and relationships, and docx4j's style, numbering
and font resolution, over the Office Open XML object model of
[`@docx4j/generated-objects-ts`](https://github.com/plutext/docx4j-generated-objects-ts). Node,
browsers and Word add-ins.

Status: the Open Packaging layer, the typed parts and the packages (Phase A of
[CR-001](docs/change-requests/CR-001-engine.md)) are implemented; the resolution utilities
(`PropertyResolver`, list numbering, fonts: Phase B) are next.

```
npm install @docx4j/core-ts
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
`@docx4j/core-ts/opc`, `/parts`, `/packages` and `/model` give the layers separately.

### The content API

`Body`, `Paragraph`, `Range` and `Font` follow Office JS's `Word.*` shapes (a compile-time
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
title.styleBuiltIn = 'Heading 1';
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

Tables, pictures, `insertOoxml` from a `pkg:package`, custom XML parts, XML mapping and typed
content controls are the next phases of
[CR-002](docs/change-requests/CR-002-content-api.md); effective formatting (styles resolved,
as Office JS reports it) comes with [CR-001](docs/change-requests/CR-001-engine.md) Phase B.

## Development

```
npm ci                              # the locked dependencies from npm
npm run typecheck && npm test       # test = build, then node --test test/
```

Releases are published to npm from GitHub Actions; see `RELEASING.md`.

## Licence

Apache-2.0, as docx4j. See NOTICE.
