# CR-007: Table conditions in the `PropertyResolver`

**Status:** Proposed 2026-10-04, on docx4j CR-030's hand-off (its section 9). Not scheduled; Jason
decides. CR-001 Phase B left table conditional formatting out (its sections 15.1 and 19) because
docx4j did not have it; docx4j CR-030 (VERSION_17_3_1, phases 1 to 5 done 2026-10-03 and 2026-10-04,
commits 3577e5181, ecaa95c1d, 8c387f484, 446d65c5d, 843ac12df) now does, and this is its port.

## 1. What docx4j CR-030 did, as far as this package reads it

The reference is `docs/developer/change-requests/CR-030-table-conditions-in-the-resolver.md` in
`../docx4j`, section 4 (design) and section 6 (what each phase did, with the Word probes T1 to T8
behind it). The parts a port of the resolver takes:

- **`TableContext`** (`resolver.tableContext(tbl)`, `org.docx4j.model.table`): the table style the
  table resolves to, its look (`w:tblLook`, else the style's, else Word's 04A0), the band sizes, each
  row's index and each cell's first grid column and span (`w:gridBefore`, then the spans, nested
  tables excluded). `forParagraph(tr, tc, pPr)` and `forCell(tr, tc)` give a `CellContext`;
  `rowConditions(tr)` and `applicable(conditions)` serve the table writers, which this package has
  none of.
- **`CellContext`**, immutable: the table style id, the conditions (`TableStyleConditions.resolve`:
  the row's, cell's and paragraph's `w:cnfStyle` when written, else the position; gated by the look),
  the applicable `w:tblStylePr` entries in ECMA-376 17.7.6 order, and a structured cache key
  (paragraph style id, table style id, condition set), never a concatenated string.
- **The overloads:** `getEffectivePPr(pPr, ctx)`, `getEffectiveRPr(rPr, pPr, ctx)`,
  `getEffectiveParagraphMarkRPr(pPr, ctx)`. Composition: document defaults, then the table level
  (the table style's own `pPr`/`rPr`, then the applicable conditions in order), then the paragraph
  style chain, then direct; toggles combine across the table and paragraph-style levels. A null
  context gives exactly the old answer.
- **`cellContextOf(p)`**: the nearest `w:tc`, `w:tr`, `w:tbl` up the parent pointers, through
  `w:sdt`, `w:customXml`, `w:smartTag` and `mc:AlternateContent`; null at a story root reached
  first (a text box, a note, a comment, a header or footer, the body). One `TableContext` per call;
  code resolving every paragraph keeps a **`TableContextTracker`** (a traversal callback holding the
  table stack, emptied inside `w:txbxContent`).
- **Decided by name, not id** (phase 2; T1, T2, T7): walking a table's style chain, a style *named*
  "Normal Table" contributes Word's built-in (108 twips left and right, nothing to text) and the walk
  ends there; every other style applies as written, the document's `w:default` table style included,
  with nothing beneath it; a style-less table takes the default style whatever its name.
  `reachesDefaultTableStyle` follows the same rule. This package decides by the default style's id
  (CR-001 section 15.1, from CR-015's probe), which is the same answer only where the default style
  is the one named "Normal Table".
- **The [MS-DOCX] `overrideTableStyleFontSizeAndJustification` exception** (12pt and left-aligned
  table-style text yielding to the paragraph style) applies only below compatibility mode 15 and
  there only where the setting is off or absent, which includes a document with no
  `compatibilityMode` (mode 12); never in mode 15 (T5, T6; `appliesTableStyleSizeJcException()`).
  This package has no such exception today, so there is nothing to correct, only to add.
- Not for this package: the synthetic style ids (`styleIdFor`, `syntheticStyle`, `sourceStyleOf`),
  which serve docx4j's HTML and FO XSLT pathways; the writers' conditions; `FontsAnalysis`'s
  in-context walk (this package's `fontsInUse` and the `RunFontSelector` are the counterpart, and
  take a `CellContext` the same way if they are to agree).

## 2. What this package would do

- `PropertyResolver`: `tableContext(tbl)`, `TableContext`, `CellContext`, `TableStyleConditions`
  (`src/model/properties/table.mts`), the three overloads, `cellContextOf(p)`, the name rule in
  `getEffectiveTableStyle` and `reachesDefaultTableStyle`, and the compatibility-mode exception read
  from the settings part. Names follow docx4j's (CLAUDE.md), `TableContextTracker` as a callback the
  content API's walkers use.
- The content API: `Paragraph.formatting()`, `Font` and the paragraph mark's effective reads in a
  table cell pass `cellContextOf(p)` (the paragraph's `PARENT` chain is kept by `linkParents`), so
  "bold because the header row is bold" reads as bold, which the editor's ED-003 noted it does not
  (its table CSS comes from `getEffectiveTableStyle` and merges no conditions per cell). The
  `RunFontSelector` reads the same effective `rPr`, so the document font follows.
- Parity (CR-001 Phase B's contract): the harness and goldens gain, per table paragraph, the table
  style id, the condition key and the in-context `pPr`, `rPr` and mark `rPr`, as CR-030 section 9
  asks; the table entries' `effectiveTableStyle` and `reachesDefaultTableStyle` move to the name
  rule when docx4j's do. Zero differences stays the bar.

## 3. What CR-030 changes under the goldens now, before any port

Checked on 2026-10-04 by building docx4j at 843ac12df (a `git archive` export, a separate Maven
repository, `test/java/README.md`'s recipe; the jar `docx4j-core-17.3.1-SNAPSHOT.jar`, SHA-256
`b58d718ef31f152a`, against the goldens' `ed92b2e826456394`) and running the harness over the 46
fixtures: **zero differences** in the properties, styles, tables, numbering and fonts of every golden;
only the header's date moved. So none of the fixtures has a default table style other than the one
named "Normal Table", none a table-styled run whose font decision the in-context resolution would
change, and the goldens stay at `a8d20c1cc`'s answers. The weekly workflow will say the same once
docx4j pushes `VERSION_17_3_1`. The name rule therefore has no golden to hold it yet; phase 1 adds
one (a fixture whose `w:default` table style is not named "Normal Table", and one whose chain reaches
a style so named).

## 4. Order and dependencies

Depends on docx4j CR-030 phase 1 (done). Phase 1 here: the resolver's types, overloads and the
name rule, held by the goldens. Phase 2: the content API reads in context, and the editor's
ED-003 limitation goes. Effort: phase 1 two to three days with the goldens extended; phase 2 one.
