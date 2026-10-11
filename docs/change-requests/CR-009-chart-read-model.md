# CR-009: The chart read model (docx4j CR-035 phase 2), held to its JSON goldens

**Status:** Proposed 2026-10-12, on docx4j CR-035 phase 2's hand-off (the docx4j session,
2026-10-12). Not scheduled; Jason decides. New scope, not a parity-bound correction.

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
