# CR-009: The chart read model (docx4j CR-035 phase 2), held to its JSON goldens

**Status:** Proposed 2026-10-12, on docx4j CR-035 phase 2's hand-off (the docx4j session,
2026-10-12). Not scheduled; Jason decides. New scope, not a parity-bound correction. Section 8 (a
chartex chart lost on save, and kept: `cx1` and `cx2` understood) fixed 2026-10-17, **released in 0.5.0** (2026-10-19).

## 1. Why

docx4j CR-035 (`docs/developer/change-requests/CR-035-charts-box-read-model-describe-edit.md`,
sections 4.2 and 11) added `org.docx4j.model.chart` to docx4j-core, on `VERSION_17_3_3` at
`276325167` (not pushed when handed off): a read model over every chart a docx, pptx or xlsx
places, with a JSON form and 32 goldens, one per chart part of every fixture. This package has the
parts (`ChartPart` over `c:chartSpace`, `ChartShapePart`; `src/parts/dml/`) and the typed model
(`org_docx4j_dml_chart`), and no reading of them. docx4j's goldens are the oracle a port would be
held to, as the parity goldens hold the resolver and the XPath goldens hold `selectObjects`.

## 2. What docx4j built

- **`Charts.in(pkg)`**: every chart the package places, in document order, as a `ChartRef`: the
  host part, the anchor, the extent, the name, the chart part (`c:` or `cx:` chartex), and whether
  it sat in an `mc:AlternateContent` Choice. All three chartex waterfalls of the fixtures are in
  Choices. docx4j leaves an unselected Choice's content as DOM in xlsx and pptx; here the MCE
  preprocessor resolves a Choice before unmarshalling where its prefixes are understood
  (CR-001 section 21, CR-004 section 6), so whether `cx` content comes typed or stays DOM is the
  first thing a port checks (`UNDERSTOOD_NAMESPACES`).
- **`ChartModel.read(chart, hostTheme)`**: groups, series, axes, title, legend and chart-wide
  settings. Series go through one adapter over the eight `ser` classes; series hidden by Excel's
  chart filter (`c15:filtered*Series`) are included, marked hidden. Data is keyed by `c:pt/@idx`
  with `ptCount`: a gap is null, `#N/A` is kept as text. Fills are automatic where unstated, and
  solid colours are resolved in the host's theme through CR-034 (here CR-008 phase 2). "Shown"
  values follow MS-OI29500: a title with no text shows "Chart Title", or the name of a single named
  series (2.1.1561); a series with no name is "SeriesN", N = idx + 1 (2.1.1565); `varyColors` is
  ignored with more than one series, and on area, stock and filled radar (2.1.1571); a pie shows
  only its first series (2.1.1519); data-label positions have per-kind defaults (2.1.1456); each
  kind has its own supported `dispBlanksAs` values and default (2.1.1452).
- **`ChartExModel`**: chartex from its caches (Excel's xlsx chartex has none, only `_xlchart.v1.N`
  names).
- **`ChartCells`**: the data from the cells (the workbook, or the one embedded with the chart) and
  where it differs from the caches.
- **A measured fact:** Excel hides a filtered category by narrowing every series' `c:f` to a union
  of the visible cells, for example `(Sheet1!$A$2:$A$3,Sheet1!$A$5)`, the cache holding only those
  points and the full range in `c15:fullRef`; it does not use `categoryFilterExceptions`.

## 3. The oracle

`docx4j-core-tests/src/test/resources/charts/json/`: 32 goldens (29 charts, 3 chartex), one per
chart part of every fixture under `charts/` (its `README.md`), in JSON version 1: fixed member
order, two-space indent, cache numbers as written, gaps as null. The fixtures are the earlier
`bar-chart`, `clustered-column`, `line-chart`, `pie-chart`, `stacked-column` and `waterfall` in
docx, xlsx and (waterfall) pptx, and 16 Excel-made probes `p2-*`: every chart kind, a combination
with a secondary axis, a chartsheet, gaps, multi-level categories, a filtered chart, and a chart
pasted into PowerPoint.

A port compares object trees, never text, as `test/parity.test.mjs` does: the fixtures copied
under `test/fixtures/charts/`, the goldens under `test/golden/charts/`, their provenance (the
docx4j commit) in a README, and the weekly workflow extended to regenerate them if the model moves
(`Charts` would need a harness main class in `test/java`, as the XPath oracle has).

## 4. Phases, if scheduled

1. **Enumeration**: `Charts.in(pkg)` as `pkg.charts` over the three package kinds, the anchor and
   extent read from the drawing, the Choice flag from the loaded tree. Half a day.
2. **The read model**: `ChartModel.read` with the MS-OI29500 "shown" rules, the series adapter,
   the data keyed by index, the filtered series; fills resolved once CR-008 phase 2 exists, and
   until then the unresolved reference and the automatic flag. Two to three days against the 29
   chart goldens.
3. **Chartex and cells**: `ChartExModel` from the caches, `ChartCells` from the workbook or the
   embedded one (this package reads `.xlsx` through `SpreadsheetMLPackage`; an embedded workbook is a
   binary part). One to two days against the 3 chartex goldens and the cell comparisons.

Consumers: none has asked. The editor's ED-005 has no chart item; a describe-or-edit surface for an
agent (docx4j CR-035's own motive) would be the one to want it. Depends on CR-008 phase 2 for the
resolved fills only.

## 5. Decisions

None yet.

## 6. Noted, not in scope: editing a chart's data (docx4j CR-035 phase 4)

Handed off by the docx4j session on 2026-10-12, for information only; nothing is proposed. docx4j
`b7a811bbf` (`VERSION_17_3_3`, unpushed when handed off; CR-035 sections 4.4, 8 (D10 to D12) and 11)
completes CR-035 with `org.docx4j.model.chart.ChartDataEditor`. It can set a value, a blank, an
error, a category or a series name, add and remove categories and series, and refresh the caches.
Each edit keeps the cells, the table, the chart's references and its caches in step, with the
cell writers `org.xlsx4j.model.CellRef`, `AreaRef`, `SheetCells` and `Tables` underneath. Every
rule was measured from Excel's own edits. The oracle is 21 before/after pairs saved from Excel,
under `docx4j-core-tests/src/test/resources/charts/edit/` (`e1` to `e7`, `r1` to `r14`, each with a
`.txt` of the edit), compared part by part by `ChartDataEditorTest` (27 tests). Some things are
refused rather than guessed: resizing a table with a totals row or without a header row, adding a
bubble chart's series, and data on more than one sheet.

Relevant here only if this package ports chart editing, which would come after this CR's read
model. The pairs would then be its oracle, as the JSON goldens are this CR's. Also in docx4j, from
phase 3 (`326b88bd2`): `ChartTable`, a chart described as a title and a table with categories down
and series across.


## 7. Noted, not in scope: chart authoring (docx4j CR-036)

Handed off by the docx4j session on 2026-10-12 for the backlog. Nothing is proposed. docx4j CR-036
(`docs/developer/change-requests/CR-036-chart-authoring-edit-docx-pptx-create.md`, its section 9
names this package) creates charts and edits their data in every host. Phase 1 (`88ce2aacb`, pushed)
edits a chart's data in docx and pptx and keeps the embedded workbook live in its
`EmbeddedPackagePart`.

Phase 2 (`334d80f07`, `VERSION_17_3_3`, unpushed when handed off) creates charts in xlsx, in
`org.docx4j.model.chart`:

- **`ChartType`**: twelve types, each with its template and Excel's default style id. They are
  column and bar (clustered, stacked, stacked 100%), line, line with markers, pie, doughnut, area
  and scatter.
- **`ChartDefinition`**: shaped like POI's XDDF and built of groups. Its categories are all text,
  all numbers or all dates. It sets the title, legend, axis titles and number formats, and data
  labels, and `fromData` gives it `setData`'s shape.
- **`ChartCreator`**:
  - `inSheet` places a chart from a definition.
  - `fromCells` places one over a range, read as Excel reads it.
  - A seeded `Random` reproduces a workbook's ids.
- **`ChartSettings`**: the same settings, written on any existing chart.
- **Excel's templates**: one per type, plus a date axis, eight style parts and a colours part,
  kept byte for byte under `docx4j-core/src/main/resources/org/docx4j/model/chart/templates/`.
- **Cell writers**: `Tables.create` and `SheetCells.setNumberFormat`.

The rules a port would have to reproduce are CR-036 section 11's phase 2 findings 1 to 6. Among
them:

- Excel's text properties for each style.
- The date axis, with its base time unit taken from the dates.
- When the legend is shown or hidden, by series count.
- Each pie series repeats the first series' `c:dPt`.
- How `fromCells` decides the header row and column, and which way the series run.
- Excel's spelling of the built-in date formats in the cache.

The oracle is the Excel-made workbooks under `docx4j-core-tests/src/test/resources/charts/authoring/`
(`p1-*`, `p2-*`, `p6-*`, with `.txt` files of their series formulas). `ChartCreatorTest` compares
106 charts against them part by part, leaving the ids aside, and Excel opened 108 without repair.

Phase 3 (`9ca75ced2`, unpushed when handed off, and handed off the same day) creates charts in docx
and pptx, and deletes them. CR-036 is done with it. The API is on `ChartCreator`:

- `inDocument` returns a `w:r` holding the `wp:inline`.
- `inNewParagraph` adds a paragraph holding the chart.
- `inSlide` adds a `p:graphicFrame` to the slide.
- `delete(ChartRef)` works in all three hosts.

The rules, measured on what Word and PowerPoint write:

- **The embedded workbook** is Word's own for each type, and PowerPoint's is the same, shipped byte
  for byte. Its table has no `autoFilter` and no `tableStyleInfo` name. Creation empties it, writes
  the data at A1 with the xlsx writer, and gives it fresh uids.
- **A docx chart** is Excel's template plus Word's host step:
  - `c:printSettings` is removed.
  - `c:externalData` points at the embedded workbook, with `c:autoUpdate` 0.
  - A scatter series' line width is halved.
  - Each of a pie's points gets a `c16:uniqueId`.
- **A pptx chart** is PowerPoint's own chart part and style part for each type.
- **The legend** is the template's, whatever the data. Excel's rules for showing it apply only when
  the host is an xlsx.
- **Ids and names**: Word's `wp:docPr` id is random and unused, named "Chart n". PowerPoint's frame
  id is the slide's highest plus one. The chart parts and `Microsoft_Excel_Worksheet*.xlsx` are named
  as Office names them, and the workbook is typed by an `xlsx` Default.
- **Deleting a chart** removes the run or the frame, the chart part, its style, colours and workbook,
  and their overrides. The `xlsx` Default goes with the last workbook.

The oracles are `charts/authoring/p3-*.docx` and `p4-*.pptx`, plus `p3-no-chart.docx` for deletion.
`ChartCreatorHostTest` finds all 24 type and host cases identical to Office's, leaving the ids aside.

Phase 3 also touched two classes this package ports, with no change to what they do by default:

- `io3/Save` and `SaveToZipFile` now call `DocPropsCorePart.writeTimes()`. It writes
  `dcterms:modified`, and `dcterms:created` on a new package, only where the new
  `docx4j.dcterms.write` property is true. The default is false.
- `DrawingPropsIdTracker` gained `generateRandomId(Random)`, and `generateId` is unchanged.

Nothing here follows either.

This package would come to it after this CR's read model, if at all. It reads the same chart parts,
and `SpreadsheetMLPackage` is the xlsx host. No consumer has asked.

## 8. A chartex chart lost on save, and kept (found 2026-10-14, fixed 2026-10-17)

Found while checking objects `5b43f6c` for release, and fixed at Jason's word. It is the question
section 2 left open: whether a Choice's `cx` content comes typed or stays DOM. Until this fix it came
as neither, because the Choice wasn't taken.

- **The case.** Word, Excel and PowerPoint write a chartex chart's frame in `mc:AlternateContent`.
  The Choice requires a later chartex namespace than the chart's own: `Requires="cx1"` (2015/9/8)
  for a waterfall or a histogram, `"cx2"` (2015/10/21) for a funnel. The Fallback is a picture of
  the chart in Word and PowerPoint, and in Excel a shape whose text says the chart isn't available.
  Neither namespace was in `UNDERSTOOD_NAMESPACES`, so loading took the Fallback. Any save that
  re-marshalled the part then wrote the picture or the shape in the chart's place, with no
  `cx:chart`. That happens after `unmarshalAll`, and in a docx after `getBody`. The chartEx part
  stayed in the package with nothing pointing to it. Measured on docx4j's five chartex files: the
  three waterfalls under `charts/` and the two histogram workbooks under `charts/chartex/`.
- **The Choice.** Nothing in it is in the namespace it requires. The frame's `a:graphicData` holds
  a `cx:chart` of the 2014 namespace, which the model binds: it loads typed, as `CTRelId`. Nothing
  binds the 2015 namespaces (docx4j has no schema for them), but the Choice round-trips whole. The
  check compared each element's namespace, name, attributes and children, leaving prefix
  declarations aside. It covered docx4j's five files and all 251 chartex files in the test
  resources of Plutext-Enterprise-Java-11's `docx4j-Enterprise-ChartEx` module (239 `cx1`, 12
  `cx2`), read in place and not copied here.
- **Fixed** in `src/opc/mce/understood.mts`: the two namespaces are a hand-kept group of their own.
  Any later chartex namespace is left out, since no file seen gates a Choice with one.
- **The cost.** As with every Choice taken, the save writes the Choice alone, without the
  `mc:AlternateContent` or its Fallback (CR-001 section 21). A consumer that doesn't understand
  `cx1` then has nothing to draw there, where it had the picture before.
- **Tests,** in `mce.test.mjs`:
  - Each of docx4j's waterfalls is loaded, unmarshalled and saved. They are copied into
    `test/fixtures/charts/` from docx4j `f9708f5b3` (docx, pptx) and `e560d03f3` (xlsx), both
    pushed. The saved frame equals the Choice's, `cx:chart` is there, and its relationship still
    names the chartEx part.
  - A built funnel Choice tests `cx2`.
  - All four fail without the two namespaces. The suite passes 826 of 826.
- **What moved.** Load, `unmarshalAll` and save of 456 fixtures (8,153 parts) were hashed before
  and after. Five parts differ: the hosts of the five chartex frames, each now keeping its Choice.
- **Found, not fixed.** The content API lists every inline drawing as an `InlinePicture`, a chart
  included. On an ordinary `c:chart` (`loadAndSave.docx`), and now on a chartex one, `imageFormat`
  throws a TypeError (`Cannot read properties of undefined (reading 'blip')`) rather than saying
  the drawing isn't a picture. Before the fix the waterfall's `InlinePicture` was its Fallback
  picture, and its reads worked.
