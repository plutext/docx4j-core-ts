# CR-013: `mc:AlternateContent` at block, row and cell level: every reader takes the selected branch (docx4j CR-051's hand-off)

**Status:** Proposed 2026-10-21, at Jason's request ("open a CR"), on docx4j CR-051 phase 3's hand-off
(the docx4j session, `VERSION_17_3_4` at `69af8cf42`). D1 decided the same day (section 5): the default
load stays as it is, every branch kept. Jason: "the default here is as it is since we want the branches
in the editor. if you were to make the default match docx4j, then the editor would need to ask for the
non-default. that consideration aside, i'm relaxed." He then added: "in fact i think docx4j ought to change
its default to match what we do here." That view was passed to docx4j as his. D2 to D4 were decided the same day,
as recommended (Jason: "D2, D3, D4 as recommended."). Addresses keep the element's form. `mcePreprocess:
true` takes the readers' set, after the round trip measured in section 3, item 6. A view in a branch is
read-only. The editor, asked about D4 the same day, also wanted (a), with its ED-009 rule if editing is
ever needed, and not resolving on edit. Phases 0 and 1 implemented 2026-10-21 (Jason: "do CR-013 phases 0
and 1"), not released: docx4j's probes in `test/mc-fun/` with every reader's answer asserted, the walker
audit (section 7), and the text model reading row- and cell-level branches as Word 365 does, kept or
resolved, with `mcePreprocess: true` over the readers' set (section 8). Phase 2 implemented 2026-10-21 (Jason: "do CR-012
section 17.13 then CR-013 phase 2"), not released: the content API's readers and the tracked-change listing take the
selected branch at block, row and cell level (and the listing at paragraph level), so every probe with a Word save
reads as the save does through every Office JS reader, kept or resolved; every view in a branch refuses its writes
(D4); a table goes only when no branch keeps a row, which closes section 7.3's two content-losing shapes (section 9).
Phase 3 implemented 2026-10-22 (Jason: "start CR-013 phase 3"), not released: the mutators reach every branch. A
listed change's twins in the other branches resolve with it and resolve-all resolves every branch, held to docx4j's
`TrackedChangesAlternateContentTest`, all 20. A table's records, cells and columns span every branch. The bulk edits,
the binding passes, replies, the ids and `toApiScript` reach every branch too. Views in a branch stay read-only
(section 10). Every phase is done.

## 1. Why

docx4j's CR-051 (phase 1, `1cff2cd5c`) admits `mc:AlternateContent` where its schema didn't: as a child of
a block container (the body, a cell, a block content control, a header, a note, a text box), of a table
(in place of a row), and of a row (in place of a cell).
- **Word reads such an element by ISO's rule** (the portfolio's paper, section 7.8). Word 365 takes the
  first Choice it understands, `w16sdtfl` on Jason's newest-first probes, in a cell, a table, a row and
  every kind of content control. Word never writes one: its save resolves the element to the branch it
  read.
- **docx4j resolves at load by default**, as Word does on open: its `word` mode, the property
  `docx4j.jaxb.mc.onLoad` (enum `McLoadMode`; named `docx4j.jaxb.mc.blockLevel` and `McBlockLevel` until docx4j
  renamed them on 2026-10-22, before 17.3.4 shipped, the setting covering row and cell level too since its phase 2).
  Its `preserve` mode keeps the element whole. Run and paragraph level are
  kept whole in both modes, as since docx4j 3.3.8 and 17.2.0.
- **The hand-off** asked how this package's readers do on the 16 `mc-fun` probes, whether its load
  takes both modes, and whether its understood set has the three Word 365 namespaces. Section 2 is the
  answer, as sent on 2026-10-21.

## 2. What this package does (measured 2026-10-21, at `2782d2b`)

The probes are docx4j-core-tests' `mc-fun/` at `69af8cf42`: 16 probes and 14 Word 365 saves. "Word" below is
this package's reading of the save, in which Word has resolved every element.
- **Load keeps the element at every level, typed.** At `body`, in a block `w:sdtContent`, in a `w:tc`'s
  content, in a `w:tbl`'s (where its branches hold `w:tr`) and in a `w:tr`'s (where they hold `w:tc`), the
  tree holds `org_docx4j_mce.AlternateContent` with its branches' content typed. The generated
  declarations don't list it in those content unions (`Body`, `Tbl`, `Tr`, `Tc`, `SdtContentBlock`):
  section 11.1 of CR-010 says block containers can't hold it, which is the declarations' view, not the
  load's.
- **Save keeps it.** On all 16 probes, an unmarshal and a save keep every element, every Choice and
  every `w:t`.
- **The text model** (`Texts.of`, CR-012), on the default load:
  - At block level, in a block content control, in a cell's block content, and around the four text
    boxes, it reads Word's text: `w16sdtfl`, and each box after its anchor. Each comes with
    `mce.branches_differ` at the element's address.
  - **At row and cell level it reads nothing of the element.** Its `rows` and `cells` walks have no case
    for it (`Extractor.mts`). The table's "row / w16sdtfl" row and the row's "w16sdtfl" cell are missing,
    with no warning. This is a reader's gap: the save keeps the content.
  - A paragraph in a block-level branch is addressed from the element: `body/2/0`, `body/2/0/0` in a
    control, `body/1/0/1/0/0` in a cell.
- **The Office JS readers** (`Paragraph.text`, `Body.text`, `Body.paragraphs`, `Table.rowCount`, ...) don't
  see a block-, row- or cell-level element at all:
  - `Body.paragraphs` doesn't list the paragraphs in a block-level branch. `Body.text` and the paragraph
    count skip them.
  - A table's rows and a row's cells leave out those in a row- or cell-level element.
  - Yet `TableCell.values` reads a block-level element in its cell (`["I prefer","w16sdtfl"]`), which its
    `paragraphs` don't list. The two readers disagree.
  - Run-level content is CR-012 section 12's rule (the editor's ED-009), in CR-012 phase 4, and isn't
    this CR's.
- **`mcePreprocess: true`** resolves on the DOM before unmarshalling, at every level, run and paragraph
  included. Every reader then agrees, but it resolves by load's understood set (`UNDERSTOOD_NAMESPACES`),
  which lacks `w16sdtfl`, `w16sdtdh` and `w16du`. So on all 13 newest-first probes it takes `w16cex`,
  where Word 365 takes `w16sdtfl`.
- **The readers' set** has counted those three and `x14ac` since CR-010 section 15
  (`READ_ONLY_UNDERSTOOD_NAMESPACES`). That is why the text model reads `w16sdtfl`.
- **A `w:sdt` directly in a branch** (docx4j CR-051 section 11.3, a defect docx4j fixed): no probe has
  one, so a block control in a block-level Choice and a run control in a run-level Choice were built.
  The text model reads both, and the save keeps both. Nothing is lost.
- `mc-run-word365.docx` is Word's save of the original-order probe, which isn't in docx4j's resources
  (docx4j's session confirmed it). It reads `w14`, rightly for that probe.

The two load modes don't map one to one onto docx4j's:

| | block, row, cell | run, paragraph |
|---|---|---|
| docx4j `word` (its default through 17.3.x) | resolved, by the readers' set | kept |
| docx4j `preserve` (its default from 17.4.0, D1) | kept (rows and cells resolved until its phase 2) | kept |
| this package's default | kept | kept |
| this package's `mcePreprocess: true` | resolved, by load's set | resolved, by load's set |

## 3. Proposal

1. **The default stays: every branch kept** (D1, decided). The editor wants the branches; a load that
   resolved them would have the editor ask for the other mode.
2. **Every reader takes the selected branch at block, row and cell level**, as CR-010 has them do at run
   level: `selectedBranchOf`, over the readers' set. A paragraph in the selected branch of a block-level
   element is a paragraph of the story, a row in a row-level element's is a row of its table, and a cell
   in a cell-level element's is a cell of its row. The other branches are invisible to readers.
3. **Mutators** reach every branch where CR-010's rule has them do so (the resolve-all verbs, `replaceText`,
   the binding passes). A view of a paragraph, row or cell in a branch is **read-only** in this CR (D4):
   the editor needs only to read there. If editing is ever wanted, it would be ED-009's rule at these
   levels: the selected branch edited in place, the others kept verbatim.
4. **The walker audit** is phase 0, as docx4j CR-051 section 10 did its own (done, section 7.3, which
   corrects and extends the list below). It goes through each walk of
   block, row or cell content in `src/` and classifies it: reader, mutator, or indifferent (CR-010
   section 11.1's three kinds). Known already:
   - **the text model:** `Extractor.rows` and `cells` (no case for the element); its block walk is right;
   - **the content API:** `Body.paragraphs`, `tables` and `contentControls`; `Table`'s and `TableRow`'s rows
     and cells (`rowsOf`, `cellsOf`), `rowCount`, `values` and `getCell`; `TableCell`'s paragraphs against
     its `values`; `listLabels` and the `Emulator`'s story walk (a numbered paragraph in a branch counts in
     Word); `outline`; `search`; addresses (`addressOf`, `elementAt`, `paragraphAt`); `paraIdTaken`;
   - **the tracked changes** (CR-011): the listing, `acceptAllRevisions` and `rejectAllRevisions` over a
     revision inside a block-, row- or cell-level branch (docx4j CR-051 section 10.3, item 3, found the
     same gap in its revisions model);
   - **the comments** (CR-002 section 54's rule): markers in a branch;
   - **the `PropertyResolver`'s `TableContextTracker`** (CR-007): which rows and cells count for a table
     style's conditional formats when one is in a row- or cell-level element;
   - **the binding passes** (`applyBindings`, the XHTML and picture passes, CR-005): a control in a branch.
   `insertXml` refusing a block-level element stays.
5. **Addresses** (D2). The text model already addresses a branch's paragraph from the element
   (`body/2/0`). docx4j's `word` mode splices the branch in, so its first paragraph is `body/2`. Its
   `preserve` mode follows the element, as this package does. Decided: the element's form, for the
   content API too: it is what the tree holds, and an address resolves without knowing which
   branch was read.
6. **`mcePreprocess: true`** (D3). Decided: the readers' set, so that a resolved load takes the
   branch Word 365 takes, as a kept load's readers already do. The two loads then agree where CR-010
   section 15.4 has them differ. CR-010 section 15.3 kept load's set because
   taking a Choice turns DOM into typed content, and nothing of the four namespaces inside it is bound, so
   a re-marshal would drop it. Measured on 2026-10-21, that reason doesn't hold for the markup Word writes:
   - **Only attributes.** In this repository's tests, docx4j's test resources and the editor's, the four
     appear only as attributes: `w16du:dateUtc` (752), `w16sdtdh:storeItemChecksum` (2, check 35's
     file), `x14ac:dyDescent` and `x14ac:knownFonts`. `w16sdtfl` is only ever declared.
   - **The attributes survive.** The object model binds none of them, yet each survives a re-marshal,
     whether or not it was in a taken Choice. Checked: a block-level Choice requiring `w16du` with
     `w16du:dateUtc` on a `w:ins`; check 35's `w15:dataBinding` with its checksum, in a Choice requiring
     `w16sdtdh`; and `cr022-checkbox.xlsx`, every part re-marshalled.
   - **An element fails the part, in every mode.** A `w16sdtfl` element in a `w:sdtPr` (made up, since
     no file has one) fails the part's unmarshal under the default load as well, inside a Choice or
     outside one. So it is no reason for one set over the other. Only load's set dodges it, and only
     inside a Choice that has another branch.

   A block-level-only mode like docx4j's `word` isn't proposed: no caller has asked for one.
7. **The declarations** (objects package). The content unions could list `AlternateContent` at block, row
   and cell level, so that code here is typed where the tree already is. docx4j asked the objects session
   to regenerate from its schema at `1cff2cd5c`. Nothing here waits on it.

## 4. Phases

| Phase | What | Size |
|---|---|---|
| 0 | The probes copied with provenance, to `test/mc-fun/` rather than `test/fixtures/` so that the parity, XPath and revisions harnesses don't take them up; tests asserting today's behaviour per reader and level (CLAUDE.md's rule: a gap fixed announces itself); the walker audit (section 3, item 4). **Done 2026-10-21, section 7** | 1 day |
| 1 | The text model at row and cell level (`rows`, `cells`, `compareBranches`), held to the 13 Word saves; `mcePreprocess: true` over the readers' set (D3), with section 3 item 6's round trips as tests. **Done 2026-10-21, section 8**, with `surfaceText` (section 7.3) | ½ day |
| 2 | The content API's readers: paragraphs, tables, rows, cells, controls, lists, outline, search, addresses; the views in a branch read-only (D4); since the audit (section 7.4) also the tracked-change listing, which the text model's views are planned from, the `TableContext`, and `textOf`'s readers. **Done 2026-10-21, section 9**, with phase 3's two content-losing shapes | 2½ to 3½ days |
| 3 | The mutators the audit lists (section 7.3), the two that lose content first (`removeRow`, `Table.deleteColumns`): the resolve-all verbs, table edits and revisions, comments, the binding passes, the ids. **Done 2026-10-22, section 10**, the two in phase 2 | 1½ to 2½ days |

## 5. Decisions

- **D1, the default** (decided 2026-10-21): kept as it is, every branch kept, for the editor. Jason also
  thinks docx4j's default should be the same ("in fact i think docx4j ought to change its default to match
  what we do here"). That is docx4j's to take up, and it was passed to its session.
  - *docx4j decided it the same day* (its CR-051 D5, revised; Jason): `preserve` becomes docx4j's default
    at 17.4.0, after its phase 2 (rows and cells in `preserve` mode); 17.3.x keeps `word`, with a CHANGELOG
    notice. From 17.4.0 docx4j's default matches this package's and docx4j-python's: every branch kept,
    selected on read. Nothing is asked of this package.
- **D2, addresses** (decided 2026-10-21, (a)): (a) the element's form, `body/2/0`, as the text model has it today; (b)
  spliced, as docx4j's `word` mode numbers the content. (b) would shift the ordinal of everything after
  the element, and an address would depend on the branch read.
- **D3, `mcePreprocess: true`'s set** (decided 2026-10-21, (b)): (a) load's set, as now, documented; (b)
  the readers' set, after a round-trip check for each namespace (CR-010's test for adding one). The check
  ran on 2026-10-21 (section 3, item 6): the three namespaces with markup in the corpora survive, and
  `w16sdtfl` has none to check. The recommendation was (a) until that measurement.
- **D4, edits through a view in a branch** (decided 2026-10-21, (a)):
  - (a) read-only, `NotSupportedError`, as CR-012 D5 has text boxes;
  - (b) the editor's ED-009 rule at block, row and cell level: the selected branch is edited in place,
    and every other branch is kept verbatim, never edited. No twins are matched;
  - (c) resolve on the first edit: the element is replaced by its selected branch's content, as Word's own
    save replaces it, and the edit applies to that. The other branches go, as they do when Word saves.

  Asked of the editor on 2026-10-21 (Jason: "ask the editor about D4"). Its answers the same day:
  - **Reading is enough for now.** Its projection shows a block-level element as an opaque block today,
    read-only, its content unseen. So its agent can't reach such a paragraph either. The gain is seeing
    the selected branch's paragraphs at the element's addresses (D2). Its projection will take the same
    branch at the same three levels when this CR lands.
  - **(a), with `NotSupportedError`**, as for text-box paragraphs. Its companion maps that to its CR-003's
    `target.invalid`, with the address, as already planned for text boxes.
  - **Not (c).** It changes addresses mid-session (`body/2/0` becomes `body/2`), which the agent's address
    and `if_text` contract doesn't cover. And it drops the other branches silently, against ED-009's rule.
  - **(b), if editing is ever needed**, in the form above. Word's save discards the other branches anyway
    (ED-009 section 6), so no consumer loses what Word would have kept.
  - **No such document in its fixtures.** A scan of every `.docx` in its repository finds no element
    under `w:body`, `w:tc`, `w:tr`, `w:tbl`, `w:sdtContent`, `w:txbxContent`, a note, a comment, a header or
    a footer: only run-level ones (Word's `wps` drawings) and ED-009's built cases. It asks where the 16
    probes live once this CR lands, to hold its projection and diagnostics to them.

## 6. References

- docx4j CR-051, `docs/developer/change-requests/CR-051-mc-alternate-content-block-level.md` at `69af8cf42`:
  section 3.2 (the mode), 3.3 (the preprocessor), 10 (the walker audit), 11.3 (a `w:sdt` in a branch).
- The probes: docx4j-core-tests' `src/test/resources/mc-fun/`; their originals on Jason's share
  (`Office 2016/mc-fun/`).
- This package: CR-010 (sections 11 and 15), CR-012 (section 12, phase 4; the text model), CR-004
  section 7.
- The portfolio's paper, section 7.8: Word's reading of the probes. docx4j's six cells read as Word 365
  does; the portfolio session would add this package's row.

## 7. Phase 0: the probes, the baseline and the walker audit (implemented 2026-10-21)

Jason: "do CR-013 phases 0 and 1".

### 7.1 The probes

docx4j's 30 `mc-fun` files are copied unedited into `test/mc-fun/`, from docx4j `69af8cf42`
(`VERSION_17_3_4`; unchanged since `1cff2cd5c`, which added them). These are its 16 probes and 14 Word
365 saves. Three original-order probes in docx4j's working tree (`mc-block`, `mc-run`, `mc-text`) aren't
committed there, so they aren't copied. `manifest.json` gives each file's SHA-1, `README.md` says what each
holds, and `mc-fun.test.mjs` fails on a file that isn't listed or doesn't match. The directory is outside
`fixtures/`, so no harness or fixture-walking test takes it up.

### 7.2 The baseline

`mc-fun.test.mjs` asserts every reader's answer on the probes. Phase 1 moved the text model's (section
8). The Office JS readers' are phase 2's, asserted as they are:
- `Body.paragraphs`, and so `Body.text`, leave out a block-level branch's paragraphs (`mc-block-`,
  `mc-sdt-block-newest-first`, `mc-mixed`'s body element, `mc-txbx-block`).
- A table's rows and cells leave out a row- or cell-level branch's (`mc-row-`, `mc-sdt-row-`, `mc-cell-`,
  `mc-sdt-cell-newest-first`).
- In `mc-tc-block-newest-first`, `TableCell.values` reads `w16sdtfl` and the cell's `paragraphs` are empty.

Every file round-trips under the default load with the main part re-marshalled: every element, branch and
`w:t` is saved.

### 7.3 The walker audit

Each walk of block, row or cell content in `src/` was read and classed as a reader, a mutator or
indifferent (CR-010 section 11.1), on `2782d2b`. Findings marked ✓ were reproduced: on the probes, or on
two variants, a `w:del` inside the block-level `w16sdtfl` Choice and a deleted plain row beside the
row-level element.

**Why they all skip it.** An `mc:AlternateContent` holds `choice[].any` and `fallback.any`, not `content`.
So `childrenOf` (`tree.mts`) answers nothing for it, and `rowsOf` and `cellsOf` match it as neither a row
nor a cell nor a holder. Every walk over them steps past the element without a word. A branch item's
`PARENT` is its Choice or Fallback, so a locator that climbs through `PARENT` and asks `childrenOf` fails
too: `paragraphFor`, `cellOf`, `ancestorTable`, `currentOwner` and `ResolvedView`'s `listOf`. And
`addressOf`/`pathOf` give `undefined` for a branch's paragraph, where D2's `body/2/0` needs the Choice
level stepped over, as `pathOf` already steps over `w:sdtContent`.

**Already right.** These read the selected branch at every level:
- the text model's block walk;
- `List`'s `walkStory`, and so `listLabels` (the `Emulator` has no story walk of its own);
- `TableContextTracker.walk`;
- `textBoxesIn`;
- `preprocessAlternateContent`, a DOM walk.

**The readers that miss it** (phase 2):
- **The content API:**
  - `Body.paragraphs`, with about 20 callers: `text`, `getText`, `search` (no hit in a branch ✓),
    `inlinePictures`, `List.paragraphs`, `listsOf`, `ContentControl.paragraphs`, `TableCell.paragraphs`,
    `Comment.paragraphs` and `getRange`;
  - `Body.tables`, and `Body.select`'s views;
  - `elementAt` (`body/2` gives the element, `body/2/0` gives `undefined` ✓), `addressOf`, `pathOf`;
  - `outline` (a row in a branch isn't counted ✓, and a table whose first row is one counts no columns).
- **Tables:**
  - every `rowsOf`/`cellsOf` reader of `Table`, `TableRow` and `TableCell` (✓ on the probes);
  - `gridWidthsOf`;
  - `TableContext`'s constructor. A branch's rows and cells get no conditional formats, and the row
    count leaves them out, so last-row and banding conditions shift. Item 4 put this under phase 3's
    mutators; it is a reader.
- **Content controls:** `collectControls`, `collectInRows` and `collectInCells`, behind
  `Body.contentControls` and `ContentControl.contentControls`.
- **Comments:** `collect`, behind `markersOf` (no marker in a branch ✓) and `commentsOf`.
- **The tracked-change listing**, `Body.getTrackedChanges`: a revision in a branch isn't listed ✓. The
  text model's views are planned from that listing (`ResolvedView`), so they plan nothing in a branch: the
  accepted view of the `w:del` variant reads the deleted text, `w16sdtflGONEc` ✓. That makes this listing
  the text model's gap too, though the text model's own walk is right.
- **`textOf`** (the objects package's, behind `TableCell.text` and `value`, `Table.text` and
  `ContentControl.text`) reads every branch differently:
  - it joins a block-level branch's paragraphs onto one line;
  - it gives `''` at row and cell level ✓;
  - it matches `Requires` with `mcBranchOf`'s own default prefixes, not the readers' set.

  That is why `TableCell.values` and `paragraphs` disagree. The objects package's `textOf` stays as it
  is (Jason, 2026-10-19), so phase 2 has to choose between an objects CR and these readers no longer
  calling it.

**The mutators that miss it** (phase 3):
- **Two that lose content:**
  - Accepting a deleted row whose table's other rows are all in a row-level element deletes the table,
    element and all ✓. `removeRow` counts the table's rows with `rowsOf`, finds none left, and removes
    the table.
  - `Table.deleteColumns` on `mc-cell-newest-first` deletes the table ✓, by the same count over
    `cellsOf`.

  These are the first to fix in phase 3.
- **Tracked changes:**
  - `resolveAll` over the listing: Accept All leaves the `w:del` in a branch ✓;
  - `resolveUnlisted`, `movePiecesOf` and `dissolveBrokenMoves`;
  - the table revisions (`acceptTable`, `rejectTable`, `collapseGrid`);
  - tracked `Body.clear` and `rowElements`;
  - the `trackInserted*` helpers, which only an insertion holding such an element reaches.
  - `joinWithNext`: a paragraph mark taken away before such an element is dropped, not joined. What
    Word does there is unmeasured.
- **Table edits:** `addRows`, `deleteRows`, `markRowsDeleted`, `markRowsInserted`, and the
  `tableColumns` passes. After `addColumns` the grid has three columns and the branch's row keeps two
  cells ✓.
- **Comments:** `removeMarkers` (`Comment.delete`, a resolution's comment removal) and `placeAfter`.
- **Bindings:** `controlsOf`, behind `applyBindings`, `updateFromContentControls` and the custom XML
  part's `removePart`.
- **Ids,** which should see every branch: `assignParaId`/`usesParaIds`, `paraIdTaken`, `nextCommentId`,
  `commentId`, and `commentParaId` (a comment whose first block is such an element has no paraId).

**Elsewhere:**
- **The text model's `surfaceText`** (a bound control's surface, REQ-509) read every branch's text, so
  a control whose content held an element warned `sdt.binding_differs` against a value its selected
  branch matched ✓. It is fixed in phase 1 (section 8.1).
- **`toApiScript`** writes a block-level element as an `insertXml` that `insertXml` refuses ✓, so the
  script can't be replayed. Its `tableGrid` would drop a row- or cell-level element from a plain-grid
  table (read, not run: the probes' tables fall back to XML).
- **Indifferent:** the XHTML and picture passes, `fontsInUse`, `nextDrawingId`, the relationship-id
  scans, `insertOoxml`'s `walkAll`, `binder` and `select`.

**The fixes with the widest reach:**
1. `rowsOf`/`cellsOf`/`located`, in a readers' form (the selected branch) and a mutators' form (every
   branch). Between them they cover the tables, the table context, the table revisions, the column passes,
   tracking, `removeRow`, `resolveUnlisted`, `Body`'s table walks and `outline`.
2. `childrenOf`, split the same way. Also answering for a Choice or Fallback would fix the `PARENT`
   locators, and `pathOf` would step over the Choice for D2.
3. `Body.paragraphs`, for its 20 callers.
4. `Body.getTrackedChanges`: the listing, the resolve-all verbs and the text model's views.
5. `collect`, and the three control collectors.

### 7.4 What the audit changes in the plan

- **The tracked-change listing** moves from phase 3 to phase 2. It is a reader, and the text model's
  accepted and original views depend on it. Its resolve-all stays in phase 3.
- **`TableContext`** moves to phase 2 as well, being a reader.
- **Phase 3 starts with the two content-losing mutators.** No file in this repository's tests or the
  editor's has a row- or cell-level element (section 8.3), so neither has met a real document.
- **`textOf`** is a choice for phase 2 (section 7.3), not this phase's.

## 8. Phase 1: the text model's rows and cells, and D3 (implemented 2026-10-21)

Jason: "do CR-013 phases 0 and 1".

### 8.1 What was built

- **Rows and cells** (`Extractor.rows`, `Extractor.cells`, `src/model/text/Extractor.mts`). An
  `mc:AlternateContent` among a table's rows, or a row's cells, now reads as its selected branch's rows or
  cells (`selectedBranchOf`, the readers' set), as the block walk already read one among blocks. Each is
  addressed from the element (D2): the second row of `mc-row-newest-first` is the element at `body/1/1`,
  and the branch's row is `body/1/1/0`. The branches are compared (`compareBranches`, REQ-307), so the
  element raises `mce.branches_differ` at its own address, as at block level.
- **`mcePreprocess: true` over the readers' set** (D3, `src/opc/mce/`). `READ_UNDERSTOOD_NAMESPACES` is
  `UNDERSTOOD_NAMESPACES` and `READ_ONLY_UNDERSTOOD_NAMESPACES`, exported from `./opc`. It is now the
  default set of `preprocessAlternateContent`, so of `mcePreprocess`, `mceKeep` (whose warning names the
  branch the readers take), `createMcePreprocessor` and `resolveAlternateContent`. `UNDERSTOOD_NAMESPACES`
  is unchanged: it is still the object model's modules, which `mce.test.mjs` holds it to.
  `READ_ONLY_UNDERSTOOD_NAMESPACES` keeps its name, CR-010's, for the API; its comment says resolving
  counts it too now. CR-010 sections 15.3 and 15.4 carry a dated note.
- **A bound control's surface** (`Extractor.surfaceText`, REQ-509) reads the selected branch of an element
  in the control's content, as the reading does, where it read every branch's text and warned
  `sdt.binding_differs` falsely (section 7.3). A branch's own items go through the visitor, so a branch in
  a `w:r`, which holds `w:t` directly, is read too.

The Markdown export follows the block model, so a row in a row-level branch is a row of the GFM table:

```
|I prefer|below|
|---|---|
|row|w16sdtfl|
```

### 8.2 What the probes say

| | before | after |
|---|---|---|
| The text model, kept, against Word 365's 13 saves | 9 agree; the 4 row and cell probes miss the element's row or cell, with no warning | 13 agree, each element warning at its address |
| The text model, resolved | `w16cex` on every probe whose first Choice requires `w16sdtfl` | `w16sdtfl` on every probe, as Word 365 reads it |
| The Office JS readers, resolved | `w16cex`'s row, cell or paragraph | `w16sdtfl`'s |
| The Office JS readers, kept | the element unseen at block, row and cell level | unchanged: phase 2 |

All 16 probes now read as Word 365 does in the text model, kept or resolved. The three with no Word save
(`mc-text-`, `mc-run-` and `mc-block-newest-first`) read `w16sdtfl`, Word's reading in the paper's
section 7.8.

### 8.3 What moved elsewhere

- **One test announced D3**, as it was meant to: `kept-branches.test.mjs` had a resolved load reading a
  picture's Fallback behind a Choice requiring one of the four. It now reads the Choice's, as the kept load
  does.
- **Nothing in the corpus.** In this repository's tests and the editor's (476 packages) no element stands
  at block, row or cell level, and a `Requires` names one of the four only in the editor's two run-level
  check files (`mc-branches.docx`, `mc-for-word.docx`), which it loads kept. The suite's 786 rows, the
  agreement test's 588 comparisons and the parity, XPath and revisions goldens are unchanged.

### 8.4 Tests

`test/mc-fun.test.mjs`, seven tests: the copy's manifest; every probe's reading, kept and resolved, held to
Word's save where there is one; the addresses and warnings at block, row and cell level, in a content
control too; the Office JS readers' gaps, asserted as they are (phase 2's); a resolved load giving those
readers Word's branch; every element saved whole after a re-marshal, on all 30 files; and D3's round trips,
`w16du:dateUtc` in a taken `w16du` Choice and check 35's `w16sdtdh:storeItemChecksum` in a taken
`w16sdtdh` one. `text-model.test.mjs` adds a bound control's surface in a branch: run-level, in a `w:r`,
and a selected branch that does differ, which still warns. The first two warned before. 1,105 tests pass.

### 8.5 What CR-012 phase 4 changes here (2026-10-21)

CR-012 phase 4 (its section 17) came between this CR's phases 1 and 2, as agreed. It changes what section
7.3's audit found in three places:
- **`textOf` is gone from the cell, table and control readers.** `TableCell.text` and `value`, `Table.text`
  and `ContentControl.text` read their paragraphs' `Paragraph.text`. So `TableCell.values` and `paragraphs`
  no longer disagree. Both miss a block-level element in the cell until phase 2, and `mc-fun.test.mjs`
  asserts it so. Section 7.3's choice between an objects CR and these readers leaving `textOf` is made: they
  left.
- **The run reader reads a branch at run and paragraph level** (`segmentsOf`, the editor's ED-009). Block,
  row and cell level are still phase 2's.
- **A PARENT climb steps out of a text box.** `addressOf` goes from a box's paragraph to its anchor
  (`textbox:N`). A text box inside a block-level element is not addressed: its paragraph's address starts
  in the element, which `elementAt` doesn't enter. The three `mc-fun` `txbx` probes are asserted as that
  gap (CR-012 section 17.4), for phase 2's addresses (D2) to close.

### 8.6 docx4j CR-051 phase 2, for phases 2 and 3 (2026-10-21)

docx4j-25 handed over CR-051 phase 2 (`9db1a3cbd`, unpushed when told). In `preserve` mode docx4j now keeps the
element whole at row and cell level too, so its `preserve` load matches this package's default at every
level. It becomes docx4j's default at 17.4.0.
- **Addresses.** docx4j addresses the element as a level, a row inside it `body/3/1/0`. This package's text
  model has done the same since phase 1: on the probes the row-level element is `body/1/1`, its row
  `body/1/1/0`; the cell-level one is `body/1/0/1`. `mce.branches_differ` is raised at the element's address.
  Nothing differs, and docx4j was told so.
- **Its fixes are the shapes section 7.3 found here:**
  - `Lister.hasRows` decides over every branch whether a table keeps a row: phase 3's first item (`removeRow`,
    `Table.deleteColumns`);
  - `Lister.alternateRows` and `alternateCells` list a row's or a cell's revisions in a branch: phase 2's
    listing.
- **Its tests are the oracle for those phases:** `McLoadModeProbesTest.preserveModeReadsAsWordsSave`,
  `McLoadModeReadersTest` (the element's address, both modes) and `TrackedChangesAlternateContentTest` (row and
  cell bodies), with `acceptingADeletedRowKeepsATableWhoseOtherBranchHasARow`.

## 9. Phase 2: the content API's readers and the tracked-change listing (implemented 2026-10-21)

Jason: "do CR-012 section 17.13 then CR-013 phase 2". Phase 2 began from `9307bdb`, CR-012 phase 4.

### 9.1 The walks' modes

The audit's three fixes with the widest reach (section 7.3), in `tree.mts`:
- **`childrenOf(value, 'read')`** answers for a kept `mc:AlternateContent` with its selected branch
  (`selectedBranchOf`, the readers' set). Without the mode it still answers nothing for the element, as the
  mutators' walks need until phase 3.
- **A Choice or the Fallback answers with its items** (`any`), mode or not. So every locator that climbs
  `PARENT` from a branch's paragraph, row or cell finds the list it is in: `paragraphFor`, `cellOf`,
  `ResolvedView`'s `listOf`, the tracked changes' `currentOwner`.
- **`rowsOf(tbl, mode)` and `cellsOf(tr, mode)`** descend into a row- or cell-level element: `'read'` its
  selected branch, `'every'` every branch, the Choices first. `BranchMode` names the two; with no mode the
  element is stepped over, as before.
- **`inBranch(value)`** is true for a value under a Choice or the Fallback, from `PARENT` links. A
  paragraph's ancestors never include a run- or paragraph-level element, which is inside the paragraph, so
  ED-009's edits there are untouched. **`assertOutsideBranch(value, className)`** throws D4's
  `NotSupportedError`.

### 9.2 The readers

Each takes the selected branch at block, row and cell level:
- **`Body`:** `paragraphs` (so `search`, `inlinePictures`, `lists`, `getText` and every reader built on
  them), `text`, `tables` (a table in a block-level branch is the body's, as Word reads it), `elementAt` and
  `pathOf`.
- **Addresses (D2).** `pathOf` steps over the Choice level, so a branch's paragraph comes back in the
  element's form: `body/2/0`, a row's `body/1/1/0/0/0`, a cell's `body/1/0/1/0/0`. A paragraph in another
  branch has no address, since no reader sees it. So the three text boxes in a block-level element agree
  with the text model now, and all 132 do (CR-012 section 17.4).
- **Tables:** `Table.rowCount`, `rows`, `values`, `getCell`, `headerRowCount`, `columnWidths`;
  `TableRow.rowIndex`, `cellCount`, `cells`; `TableCell.cellIndex`; the `cellOf` locator; `gridWidthsOf`.
- **Content controls:** `collectControls`, `collectInRows`, `collectInCells` and `collectRunControls` (a
  paragraph-level element's controls too).
- **Comments:** the marker walk (`markersOf`) behind `getComments` and `Comment.getRange`.
- **The `TableContext`:** a branch's rows and cells count for a table style's first row, last row and bands.
- **`outline()`** is the text model's, right since phase 1. `Body.select` builds its views from the readers
  above, so a hit in the selected branch is a view, and one in another branch stays a `Bound`.

### 9.3 The tracked-change listing (section 7.4)

`Body.getTrackedChanges` lists a block-level element's selected branch, and a table's row- and cell-level
branches, with nothing kept apart. Word's save has no element there. `trackedChangeTokensOfParagraph` does
the same for a paragraph-level element: the same gap, and docx4j's CR-051 section 10.5 item 3. The rows' and
the table property revisions' texts read the branch too.
- **docx4j's shapes.** On `TrackedChangesAlternateContentTest`'s four (block, paragraph, row and cell level;
  docx4j CR-051 phase 2, `9db1a3cbd`), each lists docx4j's two changes, an insertion and a deletion, from the
  selected branch. They had listed none.
- **The text model's views are planned from this listing** (`ResolvedView`). So its accepted view no longer
  reads a deletion in a branch (`C INSERTED`, where it read `C INSERTEDDELETED`), and its original view no
  longer reads an insertion there.
- **Resolution resolves what was listed:** the selected branch. Accept All counts and resolves docx4j's
  two. **The other branches keep their revisions**, which docx4j resolves with their twins
  (`acceptAllResolvesEveryBranch`). That is phase 3's resolve-all, asserted as it is.
- **The table property revisions** are listed and resolved in the selected branch. `collapseGrid` counts
  and renumbers every branch's cells, since the grid is every branch's.

### 9.4 Two of phase 3's mutators, pulled in

The listing made section 7.3's content-losing shape reachable more often: a deleted row in a branch is now
listed, so accepting it reached `removeRow`, which counted the table's rows outside the branches. So:
- **`removeRow`** removes the table only when no branch keeps a row (`rowsOf(tbl, 'every')`, docx4j's
  `Lister.hasRows`). docx4j's `acceptingADeletedRowKeepsATableWhoseOtherBranchHasARow` passes: the Choice's
  row goes, the Fallback's stays, and the table with it. Section 7.3's shape, a plain deleted row beside the
  element, keeps its table now.
- **`Table.deleteColumns`** deletes the table only when no branch keeps a cell. On `mc-cell-newest-first`
  the table stays, with the branch's cell.
- docx4j's `acceptingADeletedRowInEveryBranchRemovesTheTable` differs, until phase 3. The Choice's row is
  accepted, and the Fallback's waits for its twin, so the table stays where docx4j removes it. Asserted as
  it is.
- **Comment markers**: `removeMarkers` and `nextCommentId` walk every branch (`markersOf(..., 'every')`),
  as `paraIdTaken` does. So a comment deleted, or removed with a resolved revision (CR-002 section 54),
  leaves no marker in any branch.

### 9.5 Views in a branch: read-only (D4)

Every view a reader gives in a branch refuses its writes with `NotSupportedError` ("... is in a branch of an
mc:AlternateContent, which is read-only (CR-013 D4)"):
- **`Paragraph`:** `assertWritable` refuses it, as it refuses a text box's paragraph, so every paragraph
  mutator, `Range` and `Font` write does. `Paragraph.inBranch` says which.
- **`Body`**, a cell's or a control's in a branch: every insert, `clear`, `insertText`, `insertXml`,
  `insertOoxml`, `insertContentControl`. Inserting beside a block in a branch is refused too.
- **`Table`, `TableRow`, `TableCell`:** each setter and verb, and `inBranch` on each. `Table.values` and
  `headerRowCount`, and `TableRow.values`, check every row and cell first, so nothing is written before a
  refusal. `Table.deleteRows` counts rows as the readers do, and refuses a row in a branch.
- **`ContentControl`:** its property setters, `putProperty`, `removeProperty`, `insertText`,
  `insertParagraph`, `delete`, `setBoundContent`, `insertCopyAfter`. Its kinds' setters too: the checkbox's,
  the date picker's, the lists' and their items', and the repeating section's. A control holding an element
  in its content is not in a branch: a write to its content is refused by its paragraph.
- **`InlinePicture`** in a branch's paragraph: its size, alt text and `delete`. A picture in a run-level
  element's branch is unchanged: deleting it removes the element (CR-010 phase 3).

The bulk edits step over a branch, as before this phase, until phase 3 has them reach every branch:
`replaceText`, a tracked `Body.clear`, a tracked table or row deletion (`markRowsDeleted`), and the binding
passes (`applyBindings`, `updateFromContentControls`, the picture pass). Each filters out what is in a branch.

### 9.6 What was checked

- **Every probe with a Word 365 save** (13; `mc-run-word365` is the original-order probe's) reads as its
  save does through `Body.paragraphs`, `Body.text` and every table's values and cells. A resolved load
  (D3) reads the same through the content controls as well, on all 16 probes.
- **Each branch paragraph's address** gives the paragraph back (`paragraphAt`), in the element's form.
- **The sweep.** Over 790 `.docx` files (this repository's tests, docx4j's test resources, the editor's
  packages), this tree and `9307bdb` were compared on `Body.paragraphs`, `Body.text`, the tables' values, the
  content controls' text, the tracked-change listing, a search and every comment's range. The 76 differences
  are all in the `mc-fun` probes (both copies), each now Word's reading. Nothing else moves, and nothing
  throws.
- No suite row, agreement comparison or golden moves: parity, XPath and the 96 revisions goldens.

### 9.7 Tests

- **`mc-fun.test.mjs`** (11 tests):
  - phase 0's gap test became "the Office JS readers take the selected branch": the 13 saves, the resolved
    loads, the addresses, the probes' readings and a search;
  - docx4j's four listing shapes: listed once from the selected branch, the text model's two views, Accept
    All and what it leaves for phase 3;
  - the row-removal cases, docx4j's two and section 7.3's two;
  - D4's refusals, view by view, with nothing written before one, and the bulk edits stepping over a branch;
  - the `TableContext`'s last row in a branch, and a comment anchored in a block-level branch: listed, its
    range read, and deleted from both branches.
- **`readers.test.mjs`:** the `textbox:` test holds all 132 boxes, where it asserted three missed.
- 1,122 tests pass.

### 9.8 What phase 3 has left

Done in phase 3 (section 10), every item.

- **Resolve-all in every branch:** the other branches' revisions, docx4j's twins. This includes a deleted
  row in every branch, which then removes the table.
- **The bulk edits that step over a branch** (section 9.5's list), and `resolveUnlisted`, `movePiecesOf`
  and `dissolveBrokenMoves` in a branch.
- **The table edits:** `addRows` beside an element, and the column passes (`addColumns`, `deleteColumns`,
  `tableColumns`). These still leave a branch's row with its old cells.
- **The ids that should see every branch:** `assignParaId` and `usesParaIds` read the selected branch;
  `commentParaId`.
- **`toApiScript`'s** block-level `insertXml`, and `joinWithNext` before an element (section 7.3).

## 10. Phase 3: the mutators reach every branch (implemented 2026-10-22)

Jason: "start CR-013 phase 3". Phase 3 began from `df5b4af`, phase 2. The oracle is docx4j's
`TrackedChangesAlternateContentTest` at `VERSION_17_3_4`: its CR-051 phase 2 (`9db1a3cbd`), and two fixes made
since, both from docx4j-python's port. `60bd4b9ad` removes a deleted cell's twin in its own branch. `901a922a3`
lists and resolves a table's property records over every branch, and collapses the grid by each reading of a
row. A third, in docx4j-25's working tree when read and committed since in `84b202c5a`, has a table change's
text read Word's branch (`aTableChangesTextReadsTheBranchWordReads`).

### 10.1 Twins

- **Two listings**, docx4j's two `Lister`s (`Listing` in `TrackedChange.mts`):
  - **The public listing** (`getTrackedChanges`, `getAllTrackedChanges`) lists the selected branch's changes
    where the element stands, as phase 2 did. Each other branch's changes become their twins, paired by kind
    and position, never by id (`pairTwins`, docx4j's). `accept()` and `reject()` resolve a change's twins
    with it (`resolveTwins`), and each twin its own, for an element in an element's branch.
  - **Resolve-all's listing** lists every other branch's changes on their own, after the story's. So each is
    resolved, and a filter matches each on its own. The count stays the public listing's.
- **At every level:** block, paragraph (`alternateTokens` among runs), row and cell. `Body.getTrackedChanges`
  now walks a table's rows and a row's cells itself (`visitRows`, `visitCells`), to meet the element there.
- **Departure: pieces are paired, not listed changes.** docx4j pairs listed changes, groups included. This
  listing groups the selected branch's changes with what touches the element (phase 2 keeps nothing apart
  there), so a group can run across the element's edge. Pairing pieces gives the same pairs on every shape
  docx4j tests. It also means a piece resolved on its own resolves its twin, where docx4j leaves the twin to
  `acceptAll`.
- **Text boxes** (a gap CR-011 section 7 left open): a text box in a kept element's selected branch has as twin
  the box at its place in each other branch (docx4j's `Lister.twins`), paired the same way. So one change
  accepted in a text box resolves its `mc:Fallback` twin, as docx4j's does. Resolve-all already resolved
  every branch's box (CR-010 section 11.2), so its results don't move.
- A change in another branch that pairs with nothing is left to resolve-all, as in docx4j.

### 10.2 Tables

- **A deleted row in every branch:** accepted, its twin goes with it, and the table with both. docx4j's
  `acceptingADeletedRowInEveryBranchRemovesTheTable` passes, which phase 2 asserted as differing.
- **`removeCell`'s neighbour** is the row's as a reader sees it. A twin's is its own branch's, and a twin alone
  in its branch widens nothing, the row's neighbour having been widened once (docx4j `60bd4b9ad`).
- **A table's property records** are one change over every branch (docx4j `901a922a3`): `inChange`,
  `hasTableRecords`, `tableRecordsOf`, `acceptTable` and `rejectTable` read every branch's rows and cells.
  A record only another branch holds is listed and resolved. **The change's text** reads the selected branch,
  what Word shows, so such a change has none, as in docx4j's `84b202c5a`.
- **`collapseGrid`** counts the grid along each reading of a row (`cellLayoutsOf`, docx4j's `Lister.cellLayoutsOf`):
  the selected branches' reading, then each other branch of each cell-level element in its place. Phase 2
  counted every branch's cells along one row, which put the second branch's cells past the grid and dropped
  their spans.
  - **docx4j writes each reading's spans as it counts**, the last reading first. A cell outside the element is
    in every reading, so the first reading counts it with the span already written. Where a collapse changes
    such a cell's span, the cells after it are counted wrong. One example: a cell spanning two columns before
    the element, its branches' cells spanning four, and six columns collapsed to three. Reading docx4j's code,
    the Choice's cell would lose its span there. Here every span is counted before any is written, and the
    result is the first cell at one column and each branch's at two. Reported to docx4j (section 10.7).
- **The column passes** (`addColumns`, `insertColumns`, `deleteColumns`, `deleteColumn`, and their records) walk
  every branch's rows and each row's readings:
  - a new cell goes where the boundary falls in each reading, once where readings share the place, so each
    branch of a cell-level element gains it;
  - a deleted column's cell goes in every branch;
  - `gridRangeOf` reads along the readings, so a cell after a cell-level element is placed past its columns.
    It was placed as if the element weren't there.
- **`addRows` beside an element** was right already: new rows go into the table's own content, outside it. A
  test holds it.

### 10.3 The bulk edits

- **`writingEveryBranch`** (`tree.mts`): inside a bulk edit, the views it makes of what is in a branch write.
  A caller's views still refuse (D4).
- **`replaceText`** searches every branch's paragraphs (`paragraphsIn('every')`), each on its own, and counts
  the matches a reader sees, as resolve-all counts the listed changes. A body in a branch refuses it (D4).
- **A tracked `Body.clear`** marks every branch's paragraphs and rows deleted.
- **A tracked row or table deletion** (`deleteRows`, `TableRow.delete`, `Table.delete`) marks every branch's rows
  and their cells' content deleted (`Body.markContentDeleted`). **Inserted content** (`trackInsertedTable`,
  `trackInsertedBlock`, `markRowsInserted`) marks every branch's rows, cells and blocks.
- **The binding passes:**
  - `applyBindings` binds every branch's controls (`Body.contentControlsIn('every')`, an extension) and counts
    the ones a reader sees;
  - `updateFromContentControls` reads the controls a reader sees, since the selected branch is what Word shows
    and writes back;
  - `removePart` unbinds every branch's controls.
- **A reply** (`Comment.reply`) is anchored after its parent's markers in every branch that holds them.
- **`resolveUnlisted`** sweeps every branch, and its markers sweep reaches run-level branches.
- **Moves:** the selected branch is part of the story it stands in, and another branch is a story of its own
  (`storyOf`; docx4j's twin stories). So a move is paired within its branch. `movePiecesOf` reads the selected
  branch at block, row and cell level. A run-level element stays one object, as before.

### 10.4 Ids, `toApiScript`, joins

- **`assignParaId`** decides whether the document uses `w14:paraId`, and finds a free one, over every
  branch. `paraIdTaken` and `nextCommentId` have since phase 2, and `nextSdtId` walks the whole tree.
- **`commentParaId`**: a comment whose content starts with an element reads the selected branch's first
  paragraph.
- **`toApiScript`** writes a block-level element as its selected branch's content, through the verbs. That is
  Word's reading, and what its save writes. A table with a row- or cell-level element goes whole through
  `insertXml`, which takes it: only the element at block level is refused. Its plain-grid reading had dropped
  the element's rows.
- **`joinWithNext` before an element** stays as it is, decided as docx4j has it: a deleted mark before an
  element, or at a branch's end, is dropped, not joined. "A join never crosses a container's boundary" (docx4j's
  U1). What Word does there is unmeasured.

### 10.5 What was checked

- **docx4j's `TrackedChangesAlternateContentTest`**, all 20 cases, built in memory as docx4j builds them, give
  docx4j's answers here. That includes the three phase 2 asserted as differing, and the four from docx4j's
  fixes since (`60bd4b9ad`, `901a922a3` and `84b202c5a`'s text case).
- **The sweep.** Over 797 `.docx` files (313 of this repository's, 296 of docx4j's test resources, 188 of the
  editor's), this tree and `df5b4af` were compared on: the listing, Accept All, Reject All, every listed change
  resolved one at a time, `replaceText`, a tracked `Body.clear`, `addColumns` then `deleteColumns`, a tracked
  `Table.delete`, `applyBindings`, and `toApiScript`.
  - **35 files differ, each as intended:**
    - the `mc-fun` probes, both copies (25 files): the tracked clear, the tracked table deletion, the column
      edits and `toApiScript` now reach the element;
    - the 10 with revisions in a text box's `mc:Fallback` (`cr043-tracked` and the extraction probe's Word
      saves, both copies): resolved one at a time, phase 2 left two revisions in each Fallback box and phase 3
      leaves none.
  - **The editor's `nested-table.docx` first differed in revision ids only**, from marking a nested table's rows
    before its cell's paragraphs. `markContentDeleted` keeps the old order now, and the file is identical.
  - **Everywhere else, nothing moves.** That covers Accept All, Reject All, the listing, `replaceText` and the
    binding passes on all 797 files, since no file outside the probes has a block-, row- or cell-level element
    with revisions or a bound control. Nothing throws.
- **No suite row, agreement comparison or golden moves**: parity, XPath and the 96 revisions goldens.

### 10.6 Tests

- **`mc-fun.test.mjs`** (19 tests, 8 more). Phase 2's three gap tests now assert phase 3's answers: Accept All in
  every branch, a deleted row in every branch taking its table, and the bulk edits reaching a branch. The new
  tests:
  - docx4j's twins at all four levels, Reject All, and a piece resolved alone;
  - a filtered resolve-all, matching each branch's changes on their own;
  - the cell and record cases, and a grid collapsed along each reading of a row;
  - a text box's twin (`cr043-tracked`);
  - a move within a branch;
  - the tracked table deletion, the binding passes in both directions, `removePart`, a reply, a table inserted
    under tracking;
  - the column passes on `mc-cell-newest-first` and a row-level element, with rows added beside it;
  - a `w14:paraId` only in a Fallback, and `toApiScript`.
- 1,130 tests pass.

### 10.7 What is left

Nothing of this CR. Three notes:
- **Views in a branch stay read-only** (D4). Editing there would be option (b), ED-009's rule at these levels,
  which nobody has asked for.
- **Comment markers in a run-level branch other than the selected one** aren't reached by the marker walk.
  That walk reads a paragraph by its segments, so it sees the selected branch only. This is CR-010's run-level
  rule (the editor's ED-009), not this CR's: no deleted comment has been found leaving one.
- **docx4j's `collapseGrid`** reads a span it has already rewritten (section 10.2). Reported to docx4j-25 on 2026-10-22, from
  reading its code. docx4j-25 measured the shape the same day: it gave the Choice's cell span 1 where 2, as
  traced. It fixed it in `84b202c5a` (`VERSION_17_3_4`, unpushed when told): the counts are taken along every
  reading from the spans as they were, then written. Its test is
  `rejectingAGridChangeCountsEveryBranchsSpansFromTheOldOnes`, and its CR-051 section 12.5 records it. The same
  commit has the table change's text read Word's branch. It also confirmed the table change's text: a record held only by another branch is
  listed with empty text, as intended.
