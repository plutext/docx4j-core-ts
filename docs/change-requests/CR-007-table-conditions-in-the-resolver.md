# CR-007: Table conditions in the `PropertyResolver`

**Status:** Implemented 2026-10-04, unreleased: phase 1 (the resolver, held by the goldens; section
5) and phase 2 (the content API reads in context; section 6). Proposed 2026-10-04, on docx4j CR-030's
hand-off (its section 9). CR-001 Phase B left table conditional formatting out (its sections 15.1 and 19) because
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

## 5. Phase 1 as implemented (2026-10-04)

Started at Jason's word the day it was proposed. Ported against docx4j `VERSION_17_3_1` at
`843ac12df` (the commits of section 1, local in `../docx4j` and not pushed when this was written).

**What landed.**

- `src/model/properties/tableStyleConditions.mts`, exported as the namespace
  `TableStyleConditions`: the look, the band sizes, `fromCnf`, `atPosition`, `resolve`,
  `rowConditions`, `gate`, `lookup`, `applicable`, `formatsText`, `key`. docx4j has had this class
  since 17.2.0; this package had no port of it.
- `src/model/properties/table.mts`: `TableContext`, `CellContext` with its `CellContextKey`,
  `TableContextTracker` and `enclosingCellOf`.
- `PropertyResolver`: `tableContext(tbl)`, `cellContextOf(p)`, `getTableStyleIdOf`,
  `getTableStyleChain`, `appliesTableStyleSizeJcException()`, and a `CellContext` argument on
  `getEffectivePPr`, `getEffectiveRPr` and `getEffectiveParagraphMarkRPr`. `getEffectiveTableStyle`
  and `reachesDefaultTableStyle` decide by name. `ResolverSource` gains `settings`, which `create`
  reads privately, so the settings part still saves from its source.
- Parity: harness version 4 (`test/java`), 60 goldens (the 46 regenerated, every old answer
  unchanged, and 14 new fixtures), `test/parity.test.mjs` comparing the table contexts, each cell
  paragraph's context and its in-context `pPr`, mark `rPr` and run `rPr`. **Zero differences.** The
  test was checked to bite: reordering the precedence in `dist/` failed 17 tests, dropping the
  toggle combination 5. `test/table-context.test.mjs` holds what a golden cannot say.

**Departures from the Java, and why.**

- `TableContextTracker` is not a `TraversalUtil` callback to extend, there being no `TraversalUtil`
  here: it is the stack on its own (`enter(o)` returning the function to call on the way out,
  `cellContext(p)`, `tableContext()`), with `walk(root, visit)` as a walker over a whole tree.
- The cache key is `CellContextKey` (`tableStyleId`, `conditions`, and `mask`, the conditions as
  bits over `PRECEDENCE`); the composition caches are nested maps on (paragraph style, table
  style, mask), so nothing is a joined string, as CR-030 D8 asks.
- `null` is `undefined` throughout, as elsewhere in this package.
- Not ported, as section 1 said: `styleIdFor`, `syntheticStyle`, `sourceStyleOf`, and the table
  writers' `conditionalTrPr` / `conditionalTcPr` / `hBandRows` / `vBandCols` / `hasTblHeader`.

**Fixtures (section 3's promise).** docx4j CR-030's twelve Word probes are now parity fixtures
(`tables-*.docx`); T2 and T7 are the ones whose default table style is not named "Normal Table" and
whose chain reaches a style so named, so the name rule has goldens. They have one table shape each
(first row only), so `scripts/make-table-conditions-fixture.mjs` makes two more for the condition
arithmetic: every look, band sizes, `w:cnfStyle` caches, spans, `w:gridBefore`, a nested table,
content controls around rows and cells. Those two are not Word-measured; their goldens are docx4j's
answers.

**Consequences to know about.**

- **The harness now needs docx4j 17.3.1.** `test/java/pom.xml`'s default is `17.3.1-SNAPSHOT`; it
  does not compile against 17.3.0. The weekly workflow builds docx4j's pushed `VERSION_17_3_1`, so
  it fails until docx4j pushes CR-030. Push docx4j first, or expect one red run.
- **A behaviour change without a context:** `getEffectiveTableStyle` and `reachesDefaultTableStyle`
  answer differently for a document whose default table style is not named "Normal Table", or
  which has another style so named. None of the 46 earlier fixtures is one.
- The `RunFontSelector` took a context in phase 2 (section 6). `fontsInUse` needs none: see there.

## 6. Phase 2 as implemented (2026-10-04)

- `Paragraph.cellContext` is `resolver.cellContextOf(p)`. `Paragraph.effectivePPr`,
  `effectiveParagraphMarkRPr`, `formatting()` and the getters over them, and `Paragraph.getFont()`
  and `Range.getFont()`, pass it. So a run in a header row the table style makes bold reads
  `font.bold === true`, and `font.name` is the table style's font where it names one. Reads with
  `{ direct: true }` are unchanged, and writes are direct as before.
- The alignment setter's "do not write a value the paragraph already inherits" (CR-002 section 38)
  and `settleFormatting` compare in context too, so in a cell what is inherited includes the table
  style's.
- `RunFontSelector.documentFontFor` and `spans` take the context (a fifth argument, and
  `SpansOptions.cellContext`), as docx4j's `FontsAnalysis` walk does since CR-030 phase 5.
- `fontsInUse` is unchanged, and needs no context: it is docx4j's `FontAndStyleFinder`, which
  collects the fonts every style in use *names*, the `w:tblStylePr` run properties of a table style
  included, and resolves nothing. docx4j's in-context walk is `FontsAnalysis`'s usage walk, which
  this package has no port of. Section 1 was wrong to call `fontsInUse` its counterpart.
- Cost: each effective read of a cell's paragraph builds a `TableContext`, a walk of the table's
  rows, as `cellContextOf` does in docx4j. A caller reading every paragraph of a large table should
  walk with a `TableContextTracker` and call the resolver itself.
- Not Word-checked here: the content API's answers are the resolver's, which the goldens hold to
  docx4j, and docx4j's are Word-measured for the twelve probes only.
- `test/table-context.test.mjs` holds both phases. No existing test changed its expectation.
- The editor's ED-003 limitation (table CSS merging no conditions per cell) is the editor's to
  lift: `Paragraph.cellContext`, or `resolver.tableContext(tbl).forCell(tr, tc)` for a cell's own
  conditions, is what it would use.
