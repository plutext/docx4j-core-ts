# pptx and xlsx

The engine is one Open Packaging layer for all three formats. A `.pptx` and an `.xlsx` load, round
-trip and save exactly as a `.docx` does, with their parts typed by the same object model; what
they do **not** have is the content API — `Body`, `Paragraph`, `Range` and the rest are
WordprocessingML (CR-002), and PresentationML and SpreadsheetML are worked through their typed
parts instead, as in docx4j.

## What you get

| | WordprocessingML | PresentationML | SpreadsheetML |
|---|---|---|---|
| Load, save, flat OPC, byte-identical untouched parts | yes | yes | yes |
| Typed parts and relationships | yes | yes | yes |
| `createPackage()` | yes | yes | yes |
| Content API (`body.insertParagraph`, search, `font`) | yes | no | no |
| Style, numbering and font resolution | yes | no | no |

## A presentation

```js
import { writeFile } from 'node:fs/promises';
import { PresentationMLPackage } from '@docx4j/core-ts';

const pkg = await PresentationMLPackage.createPackage({ slideSize: 'SCREEN16x9' });
await pkg.addSlide();                       // a second slide on the same layout

pkg.slideParts;                             // in p:sldIdLst order
pkg.slideMasterParts;                       // in p:sldMasterIdLst order
pkg.slideLayoutParts;                       // every master's layouts, in p:sldLayoutIdLst order
pkg.themePart;
pkg.slideParts[0].slideLayoutPart.slideMasterPart;

await writeFile('out.pptx', await pkg.save());
```

A created presentation has docx4j's part set — `ppt/presentation.xml`, one slide master, one
slide layout, `ppt/theme/theme1.xml` (the Office theme `pkg.fonts.defaultTheme` names, the same
one a new `.docx` gets) — plus one empty slide, which docx4j leaves to the caller. The slide
sizes are docx4j's `SlideSizesWellKnown`: `SCREEN4x3`, `SCREEN16x9`, `SCREEN16x10`, `LETTER`,
`A3`, `A4`, `LEDGER`, `B4ISO`, `B5ISO`, `MM35`, `OVERHEAD`, `BANNER`, landscape unless
`{ landscape: false }`.

Shapes are the object model, built with the generated factories:

```js
import { createSld } from '@docx4j/generated-objects-ts/factory/org_pptx4j_pml';

const slide = await pkg.slideParts[0].getContents();
slide.cSld.spTree.spOrGrpSpOrGraphicFrame;   // the shapes: p:sp, p:pic, p:graphicFrame, ...
```

Loading an existing deck gives the same shortcuts, plus `notesMasterPart`,
`presentationPropertiesPart`, `viewPropertiesPart`, `tableStylesPart`, `commentAuthorsPart`, and
on a slide `notesSlidePart` and `commentsPart`.

## A workbook

```js
import { writeFile } from 'node:fs/promises';
import { SpreadsheetMLPackage } from '@docx4j/core-ts';
import { createRow, createCell } from '@docx4j/generated-objects-ts/factory/org_xlsx4j_sml';

const pkg = await SpreadsheetMLPackage.createPackage();
const sheet = pkg.createWorksheetPart('Sales');       // the tab, its r:id and the worksheet part
pkg.createWorksheetPart('Cover', 0);                  // inserted as the first tab

sheet.contents.sheetData.row = [
  createRow({ r: 1, c: [createCell({ r: 'A1', v: '42' })] }),
];

pkg.worksheetParts;                                   // in `sheets` order, the tab order
pkg.sharedStringsPart;
pkg.stylesPart;
pkg.calcChainPart;
sheet.tableParts;
sheet.drawingPart;
sheet.commentsPart;

await writeFile('out.xlsx', await pkg.save());
```

A created workbook is `xl/workbook.xml` with one `bookViews/workbookView` (docx4j adds it because
without it Excel 2010 could crash on print) and one worksheet part per `createWorksheetPart`,
each with an empty `sheetData` — docx4j's part set exactly.

## Markup compatibility

`mc:AlternateContent` is resolved when a part is unmarshalled, as PowerPoint and Excel do on open:
the first `mc:Choice` whose `Requires` namespaces the object model knows, else the `mc:Fallback`.
Where neither applies the element is dropped — which is what happens to the `x15ac:absPath` Excel
writes at the top of `xl/workbook.xml` (an `x15` Choice with no Fallback, holding the author's
local folder). A workbook nobody unmarshals still round-trips byte for byte; one that is
re-marshalled loses that element, as Excel itself would. See CR-001 sections 5.6 and 16.

## What is not here

Typed convenience over slide layouts and placeholders (docx4j's `ResolvedLayout`,
`ShapeWrapper`), the pptx-to-html line, and anything that reads or writes cell values as values
rather than as markup, are later CRs — as is a content API for either format. The parts, the
relationships and the object model are all there in the meantime.

Runnable versions: [`examples/node/pptx.mjs`](../../examples/node/pptx.mjs) and
[`examples/node/xlsx.mjs`](../../examples/node/xlsx.mjs).
