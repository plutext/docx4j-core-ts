# CR-029: An `mc:AlternateContent` that selects nothing, read as Word reads it (WD REQ-305 and REQ-306)

**Status:** Proposed 2026-11-07, at Jason's "draft CR-029", on his rulings in the WD: "REQ-305 itself changes for all three engines", "selecting nothing should raise a warning" and "two sets". Measured here the same day (section 2), Word's `te-*` saves read here. D1 to D8 decided the same day, all as recommended (Jason: "CR-029's decisions: D1 to D8 as recommended. Please implement"). Implemented 2026-11-07 against the objects release candidate (section 11), and committed with the raise to objects 0.8.0 on 2026-11-08 (Jason: "raise"), released in 0.7.0 (2026-11-09). docx4j built its side at `06628f550` and `af26a5f78`, and docx4j-python as its CR-040, complete. The objects package's CR-021 (`a7c4ac8`, released in 0.8.0) does the selecting. D9, raised 2026-11-08 by the `cell-paragraph` round (section 12), decided the same day (Jason: "D9: always write Word's spacing on the cell paragraph"), built and committed at his "commit".

## 0. Summary

Where an `mc:AlternateContent` has `mc:Choice` children, none understood, and no `mc:Fallback`, REQ-305 selects nothing: the element and all it holds are absent, as ISO and Word have it, and `mce.nothing_selected` is raised. This package takes the first Choice instead, so it shows words Word doesn't, in a namespace of the author's choosing, and says nothing.

Of 142 files Word was measured on, 118 read here as Word does. The other 24 all show a Choice's words where Word shows neither branch. With objects' CR-021 in place, in a scratch build, all 142 match, on a kept load and a resolved one.

So half of this CR is forced by the objects raise: `selectedBranchOf` has to change in that commit, since CR-021 removed the option it passes. The other half is new here: the warning, `mcePreprocess: true`, which still keeps such an element, the sets REQ-306 asks an extractor to publish, and a cell a removal leaves with no block.

Root handling stays in CR-028. No package in the corpus has this shape (section 2.3), so this is about a rare file, and about an author being unable to show a reader words Word never shows.

## 1. What Word does

Word 16.0.20326.20232, through the bridge, each file twice, the runs agreeing.

| Round | Files | Shows |
|---|---|---|
| `ul-*` | The 118 URIs of the old list, and a made-up control | The Choice for 105; neither branch for 13 and the control |
| `uc-*` | 11 namespaces the list lacked, and two controls | The Choice for `cx4`, `cx6`, `cx7`, `c16r2`, `ax`, `xsi`; neither for `cx5`, `cx8`, `mo`, `arto`, `dcmitype` |
| `c305-*` | Block and run level: `wpi`, `oel`, `cx3`, `w16sdtfl`, unknown namespaces | The Choice where understood; neither otherwise |
| `rw-25` | An unknown namespace in a body, no Fallback | Neither |
| `rw-24` | The same at the root | An empty document |

**What a removal leaves.** The `tc-empty` round, read here from Word's saves. `te-11` to `te-15` are docx4j-python's `word`-mode saves of `te-01` to `te-05`, with the element taken out, and Word treats each as it treats its source.

| Left empty | Files | Word |
|---|---|---|
| A `w:tc`, beside another cell or alone, or holding an element with no branch at all | `te-01` to `te-03`, `te-11` to `te-13` | Offers a repair, and saves an empty paragraph in the cell |
| A `w:tr` | `te-04`, `te-14` | Refuses the file |
| A `w:tbl` | `te-05`, `te-15` | Opens silently, the table gone |

**The limit.** Every file in the rounds is Choice-only, in a main document part, on one Word build. REQ-306 says so: the set is silent on spreadsheets and presentations, and a namespace not measured counts as not understood. Taking a Fallback where a Choice isn't understood is REQ-302 and ISO, and these rounds didn't put one in.

## 2. Here today (measured 2026-11-07)

### 2.1 Word's files

Each source loaded, read through `Texts.of`, and held to Word's text save, whitespace collapsed. The first column is this package at `4003051`, the second a scratch copy of `1cae4ef` with objects' CR-021 and `tree.mts:77` changed to `mcBranchOf(value)`.

| Round | Files | As Word, today | As Word, with CR-021 |
|---|---|---|---|
| `ul-*` | 119 | 105 | 119 |
| `uc-*` | 13 | 7 | 13 |
| `c305-*` | 9 | 6 | 9 |
| `rw-25` | 1 | 0 | 1 |
| **All** | **142** | **118** | **142** |

A kept load and `mcePreprocess: true` gave the same count on each row. All 24 misses show `CHOICE WORDS` where Word shows neither. None of these files has a Fallback, so the third step hides the 16 namespaces section 3.1 adds: the Choice is read today for the wrong reason.

### 2.2 The shapes

| File | Route | Today | With CR-021 |
|---|---|---|---|
| `rw-25` | Either | `CHOICE WORDS` read, nothing said | Not read, nothing said |
| `rw-24` | Either | Empty, nothing said | The same |
| `te-01` | Either | `HIDDEN` read | Not read |
| `te-04` | Either | `HIDDEN` read | Not read, and the text model writes `[table: listed format]` for the row with no cell |
| `te-05` | Either | `HIDDEN` read | Not read |

No warning is raised on any of them, and every save still holds the element, `mcePreprocess: true` included, since `preprocessAlternateContent` keeps it (`preprocessor.mts:17`, `:229`).

### 2.3 How often

Of 1,479 `mc:AlternateContent` in 1,894 packages, this repository's `test/` tree and docx4j's test resources, none selects nothing under the 114. Python found one in 719 packages, `rw-24`. An earlier count here found 985 Choices in 1,045 packages: none requires one of the 13 namespaces the set loses, and one requires one of the 16 it gains (`wpi`).

### 2.4 The readers

Every caller of `selectedBranchOf` already copes with no branch, since an element with no branches gave `undefined` before CR-021: the content API's walks, `Paragraph`, `tracking`, `TrackedChange`, `Body`, `tableRules`, `table.mts`, `List`, `toApiScript` and the Extractor's sites were each read for this. A view whose branch isn't selected already counts as in an unread branch. So nothing changes at those sites, and a mutator still reaches every branch, as CR-013 has it.

Two places still hold the old rule:
- `preprocessor.mts:296`, `textOfBranch`: a nested element with nothing selected gives `choices[0]`, in what `mce.branches_differ` compares.
- `preprocessor.mts:17`, the file comment that argues for keeping such an element.

### 2.5 The objects release candidate

A scratch copy of this package with section 3.2's change, objects packed from its checkout at `b4373c8`, whose `src/` and `modules/` are `a7c4ac8`'s, and jsonix 3.13.0, as that release's lockfile will have it. Node 22.

- **The suite:** typecheck passes, and 1,895 of 1,907 tests pass. The 12 failures are the earlier trial's: section 7's eight, and four pins announcing a fix (check 76's files waiting on the runtime, CR-025 section 1's two on `w:stylePaneFormatFilter`, and the facade's walk).
- **`onNothingSelected`,** through a prototype in `compareBranches`: on the 142 files it fires exactly on the 24 where Word shows neither branch, and all 14 `detail` values the round's index can check name the right URI. An unbound prefix and an empty `Requires` give an empty `detail`, two Choices give both namespaces, one namespace under two prefixes gives it once, and it reaches a namespace bound on the Choice alone and an element nested in a taken Fallback. An element with no branches stays silent. `{ namespaces, value }` holds all section 3.5 needs.
- **Every save against today's:** of 1,059 packages one differs, `NumberingImplicitNumId.docx`. Its `word/settings.xml` is marshalled where it was copied, since CR-019 keeps the 15 attributes that made it skipped content. All 16 of `w:stylePaneFormatFilter`'s attributes and every element survive.
- **A resolved round trip of the sixteen new namespaces:** only `wpi` occurs, in 2 of 1,894 packages, two copies of one file, in a `w:r`. It saves with nothing lost or skipped. The other fifteen gate no content in the corpus, so nothing here tries them.

## 3. Proposal

### 3.1 The sets (REQ-306, D5)

REQ-306 has an extractor publish two sets. Its *understood set* decides every markup compatibility question in a word-processing package, and its *traversal set* decides nothing about selection.

| Today, on `./opc` | Holds | Becomes |
|---|---|---|
| `UNDERSTOOD_NAMESPACES` | 107: the model's namespaces | The understood set, 114, re-exported from objects' `./builders/wml` |
| `READ_UNDERSTOOD_NAMESPACES` | 111: those and four more | Gone, its uses moving to the 114 |
| `READ_ONLY_UNDERSTOOD_NAMESPACES` | The four | Gone |
| `UNDERSTOOD_PREFIXES`, `READ_UNDERSTOOD_PREFIXES` | The same by prefix | Gone: nothing selects by prefix any more |
| (none) | | `TRAVERSAL_NAMESPACES`, docx4j's 118, whose first 107 are today's list |

Against the 114 today's readers' set is 16 short (`wpi`, `oel`, `cx3`, `cx4`, `cx6`, `cx7`, `c16r2`, `ax`, `xsi`, `dgm1611`, `xr2`, `xr3`, `xr6`, `xr9`, `xr10`, `xr16`) and 13 too wide (extended-properties, docPropsVTypes, core-properties, `dc`, `dcterms`, coverPageProps, `office:excel`, `office:powerpoint`, excel/2008/2/main, encryption, the two key encryptors, custom-properties).

Three more uses follow the understood set, as REQ-306's "every markup compatibility question" has it, and as python's did:
- `preprocessAlternateContent`, which already matches by URI;
- `breaks.mts:279`, CR-021's namespace Word doesn't know;
- `unknownElements.mts:186`, CR-026's.

`mce.test.mjs`'s check that every module's namespace is listed moves to the traversal set, and `understood.mts`'s header, which defines understanding as surviving a load and a save, is rewritten.

### 3.2 Selection

`selectedBranchOf` becomes `mcBranchOf(value)`, with no option: the 114, by URI, through the bindings a read keeps, and nothing where nothing is understood. A read through the facade resolves a prefix like `q`. In a tree built by hand a prefix the table doesn't hold counts as not understood, which is objects' D4. Objects' CR-021 section 5 lists what moves for a reader.

**The risk in `mcePreprocess: true`.** Taking a Choice converts DOM into typed content, so a binding gap the Choice hid can lose content on a re-marshal (CLAUDE.md; CR-004 section 7). The set now follows what Word understands, not what round-trips here, so sixteen namespaces go from the Fallback to the Choice. `wpi` loses nothing, and the other fifteen have no real content here to try (section 2.5), so this can't be closed by testing. D6 says what follows if one loses content.

### 3.3 `mcePreprocess: true` (D2, D3)

The pass stops keeping such an element and removes it, as it removes an element with no branch at all, in a word-processing part:
- **At every level,** run included, as Word does (`c305-*`). docx4j's splice stops at block, row and cell.
- **Not at a root.** Removing `word/document.xml`'s root would leave a part with no root at all, so `rw-24` stays kept, as today, and reads empty. CR-028 owns it.
- **`KEPT_HOSTS` come first,** as now.
- **`textOfBranch`** loses its `?? choices[0]`.
- **In any other part** the element stays, as today (D3). The word-processing test is the part's content type, which `mceKeep` and `mcePreprocess` already receive.

The default load is unchanged: the element stays in the tree with every branch, a save writes it, and only a reader sees nothing.

### 3.4 A cell left with no block (D4)

Where the removal leaves a `w:tc` with no block, the pass writes the paragraph Word's repair writes there: `w:spacing w:after="0" w:line="240" w:lineRule="auto"` on it, whatever the document's defaults say (D9, section 12). D4 first wrote a plain `w:p`, as python's `_keep_a_paragraph` does. Word repairs such a cell whether or not a removal made it, and python's save of that form opens silently where the source prompts (its section 8.2a, CR-039 section 9.7). A `w:tr` with no cell is left: Word refused the source too. An empty `w:tbl` is left: Word drops it silently.

A kept load writes what it read, so it saves the shape Word repairs, as the source is.

### 3.5 The warning (D7)

**In the text model.** `TextWarning.MCE_NOTHING_SELECTED`, `mce.nothing_selected`:
- **Raised** in `Extractor.compareBranches`, which already stands at all six sites (block, row, cell, paragraph and run), where `selected === undefined` now returns: through `mcBranchOf`'s `onNothingSelected`, for an element with Choices. An element with no Choice and no Fallback isn't REQ-305's case and stays silent.
- **Selector** the part's name, so `compareBranches` takes it.
- **`detail`** the namespaces, one string of URIs separated by spaces, empty if every prefix binds to nothing, as REQ-801 has it. `TextWarning.detail` widens from `number` to `number | string`.
- **Address** where one exists. One per element, as `mce.branches_differ` is. `TextWarning.key` gains `detail`, so two at one address requiring different namespaces are two.

**On the part.** `preprocessAlternateContent` already stands at the element, so it records `McNothingSelectedWarning` (`code`, `partName`, `line`, `column`, `namespaces`, `removed`, `root`, `message`) on `XmlPart.warnings`, and so on `pkg.warnings`, on a kept load and a resolved one alike, beside `McBranchesDifferWarning`. A nested element in a branch not taken isn't visited, as now.

**Where the walk can't meet it.** On a resolved load the element has gone, and at a root the walk never starts. The Extractor carries such a part's record into `TextResult.warnings` with no address, as docx4j's sweep does for a container left empty and REQ-801 has for a root. So `rw-24` says what it read, ahead of CR-028.

### 3.6 Changed and added, for the release notes

- **Changed:** an `mc:AlternateContent` with Choices, none understood, and no Fallback reads as absent, as Word reads it. Previously its first Choice was read. `Requires` is matched by URI over the set Word was measured to understand, so a few Choices and Fallbacks swap. `mcePreprocess: true` removes such an element in a word-processing part.
- **Added:** `mce.nothing_selected` in `TextResult.warnings` and `McNothingSelectedWarning` on `pkg.warnings`; `TRAVERSAL_NAMESPACES`.
- **Removed:** `READ_UNDERSTOOD_NAMESPACES`, `READ_ONLY_UNDERSTOOD_NAMESPACES`, `READ_UNDERSTOOD_PREFIXES`, `UNDERSTOOD_PREFIXES`. `UNDERSTOOD_NAMESPACES` now holds the understood set.

## 4. Not here

- **`mc:AlternateContent` at a root (REQ-317):** CR-028 reads it. Only the warning comes here, and WD 0.101 has since widened REQ-317 to other parts.
- **A table's grid with a row left with no cell.** `te-04` is a file Word refuses, so nothing measures it. The text model's `[table: listed format]` there is pinned, not chosen.
- **`mc:Ignorable`.** REQ-306 has the set decide it too. Here it only fills `part.skipped`'s `ignorable`, which reads the attribute and asks no set.
- **Spreadsheets and presentations.** Neither the set nor REQ-305's measurements reach them. Selection through `mcePreprocess` uses the one set there, as objects' D5, and no fixture's selection moves.

## 5. Decisions (taken 2026-11-07, all as recommended)

| # | Question | Recommendation |
|---|---|---|
| D1 | Which release? | With 0.7.0, both phases. Phase 1 has no choice, since it rides the raise, and phase 2 should go with it so a reader never selects nothing with nothing said. The alternative is phase 2 after 0.7.0, with an element that selects nothing silent until then |
| D2 | Does `mcePreprocess: true` remove such an element? | Yes, as ISO, Word, python and docx4j's `word` mode do. The file comment's reason for keeping was the danger of promoting a Choice, and removing promotes nothing. The alternative is keeping it, so a resolved save differs from Word's |
| D3 | Removal and the part warning in which parts? | Word-processing parts only. The set is measured there and REQ-306 is silent elsewhere, so a Choice-only element in a slide or a sheet could be lost for a namespace nobody asked Word about. docx4j didn't check those hosts either. The alternative is every part |
| D4 | A `w:tc` the removal leaves with no block? | An empty `w:p`, as python's, which is what Word's repair gives. The alternative leaves a save that prompts |
| D5 | The public names? | As section 3.1: `UNDERSTOOD_NAMESPACES` means what it says and the four others go. Nothing in the editor uses them. It is an API break, recorded in the release notes. The alternative keeps the old names as deprecated aliases for a release, though their content would change anyway |
| D6 | If a resolved round trip of a new namespace loses content (section 3.2)? | Report it, and leave the set as the WD has it: REQ-306 requires the namespace, and `mcePreprocess` is an opt-in. A binding gap goes to the objects package. The alternative is holding that namespace back from the pass alone, which breaks REQ-306 |
| D7 | Does the warning reach `pkg.warnings` and a resolved load? | Yes to both. docx4j raises it in both modes and REQ-305's MUST has no exception. Python's `mce_preprocess` stays silent (its Q3), but only because it has no record, and this pass has one. The alternative is the text model's kept reading alone |
| D8 | A Word check of resolved saves, as python's section 8.4? | Yes, in phase 3, through the bridge, on Jason's word: this package's `mcePreprocess: true` saves of the round's files, each held to Word's reading of its source. It's the only check that such a save opens silently |
| D9 | Raised by D8's check and the `cell-paragraph` round (section 12): should the paragraph D4 puts in an emptied cell carry Word's spacing? | **Decided** (Jason, 2026-11-08: "D9: always write Word's spacing on the cell paragraph"), and built. `w:spacing w:after="0" w:line="240" w:lineRule="auto"` on it, always. Word's repair writes exactly that wherever the paragraph would otherwise read differently, and the cell then renders as Word's. Without it the cell grows by the document's spacing, 32.6pt in `cp-11`. Where the defaults already read so (`cp-02`) Word writes none, and the extra element there changes nothing drawn. docx4j-python should move with this package. |

## 6. Phases and size

| Phase | Work | Size |
|---|---|---|
| 1 | With the raise, in its commit: section 3.1's sets, section 3.2's selection, the tests pinned to the old rule, `wpi`'s resolved round trip as a test | ½ day |
| 2 | Section 3.3's removal, section 3.4's cell, section 3.5's warning, `test/nothing-selected/` | ¾ day |
| 3 | The records, CLAUDE.md's `mc:AlternateContent` paragraph, D8's check | ¼ day |

Phase 1 waits on the objects release, and the editor is told before it (section 8). Phase 2 needs `onNothingSelected` and so follows it.

## 7. Tests and oracles

- **Word's files:** `test/nothing-selected/`, outside `fixtures/`, copied unedited with a SHA-1 manifest: the `ul-*`, `uc-*`, `c305-*`, `rw-24`, `rw-25` and `te-*` sources, Word's text saves, and Word's saves of `te-*`. Each file reads as Word's, by both routes, and the warning is raised exactly where Word shows neither branch.
- **The warning:** selector, `detail`, one per element, at each of the six sites, an empty `detail`, a resolved load and a root with no address, none where a branch is selected, none for an element with no branches, none under `warnings: false`.
- **The removal:** every level, not a root, not under `KEPT_HOSTS`, not in another part, a nested element in `mce.branches_differ`, the cell's paragraph, a `w:tr` and a `w:tbl` left as they are.
- **The default load:** an untouched part byte for byte, a save writing every branch, and a mutator reaching a branch no reader sees.

**Tests that move with phase 1,** eight in the trial, each stopping at its first failed assertion, so editing them may show more:
- `text-parity.test.mjs`: five `PENDING` `req303` entries go, since the URI now decides.
- The `extraction-probe-inspect` golden: 17.3.5 reads its `wpi` ink's Fallback and this package now reads the Choice, as Word does. docx4j re-baselines it, so a `PENDING` entry holds it until the goldens are recut.
- `kept-branches.test.mjs:59` and `:179` use `cx3` as not understood, and take a namespace like `cx5`.

Also: `mce.test.mjs:241` (kept whole) changes in phase 2, `text-extraction.test.mjs:405` holds the two sets, and the adapter claims REQ-303, its nine `waits` rows going. `create.test.mjs:263`, Excel's `x15ac:absPath`, stays kept.

**The suite stays at 17.3.5** until docx4j releases 17.3.6, when the recopy brings the two published files and `phase0/expected-inspection.tsv`'s two `wpi` rows, and `scripts/copy-text-extraction-suite.mjs` takes `traversal-namespaces.txt`. Until then the 118 in the suite's copy are the traversal set under the old file name, and the understood set is held to objects' `UNDERSTOOD_NAMESPACES`, itself held to docx4j's file.

## 8. Consumers and risks

- **The editor** calls `selectedBranchOf` in `labels.mts`, `drawings.mts`, `import.mts`, `export.mts` and `AskBox.tsx`, so its projection changes. `packages/editor-model/test/alternate-content.test.mts:134` pins the old rule by name: "with no Fallback the engine takes the first Choice (mcBranchOf), never nothing". It needs telling before the raise, and whether it shows such an element as opaque is its own decision.
- **A namespace Word knows and the set lacks** loses its words. The warning names the URI, which is why it goes with the selection (D1).
- **Released behaviour changes** for a caller who got a Choice's words. Section 3.6's line says so.
- **Order with CR-028.** Its root pass will meet `rw-24`, and its D4 already expects this warning. This CR raises it first, so that CR takes over reading the root and adds nothing to the warning.

## 9. Registry

`core-ts/CR-029`, proposed, with phases 1 to 3. Phase 1 depends on the objects package's CR-021 being released. It joins 0.7.0's list under D1.

## 10. Sources

- `docx4j-portfolio/docs/text_extraction_specification_wd.md`: REQ-302 to REQ-306 and REQ-801 at WD 0.99 (`0309843`), and WD 0.101 (`7150868`).
- The objects package's `docs/change-requests/CR-021-selection-by-namespace.md` (`a7c4ac8`) and `src/builders/wml.mts`.
- docx4j `824565363` (the two published sets), `06628f550` and `af26a5f78` (`McSplice.NOTHING`, `Extractor.compareBranches`, `McSelection.requiredNamespaces`).
- docx4j-python's CR-040, sections 3, 7, 8.2, 8.2a and 8.4.
- The share's `bridge/portfolio/`: `ul-*`, `uc-*`, `c305-*`, `rw-24`, `rw-25`, `te-*`, Word's saves and text read here.
- CR-010 section 15, CR-013 D3, CR-028 section 5's D4.

## 11. Implemented (2026-11-07)

Jason, 2026-11-07: "CR-029's decisions: D1 to D8 as recommended. Please implement". Built as section 3 proposes, phases 1 and 2 and these records, in the working tree, against objects packed at `b4373c8` and jsonix 3.13.0 through `npm install --no-save`. `package.json` and the lockfile are untouched. It can't be committed ahead of objects 0.8.0, since `main` would fail typecheck, so it lands in the raise's one commit.

- **The sets** (`understood.mts`, D5): `UNDERSTOOD_NAMESPACES` is the objects package's, `TRAVERSAL_NAMESPACES` docx4j's 118, and the other four names are gone.
- **Selection** (`tree.mts`): `selectedBranchOf` is `mcBranchOf`, and takes its options, so the text model passes `onNothingSelected`.
- **The pass** (`preprocessor.mts`): `wordProcessing`, from the part's content type, or true for a bare node. It records `McNothingSelectedWarning` and removes such an element, but for a root, a `KEPT_HOSTS` host and another part. `keepACellsParagraph` runs after any removal from a `w:tc`, so `te-03`'s element with no branches gets its paragraph too. `textOfBranch` lost `choices[0]`.
- **The text model** (`Extractor.mts`): `compareBranches` at its six sites, and `nothingSelectedUnmet` over the parts the reading took, for records marked `removed` or `root`. A part that wasn't read says nothing. `TextWarning.detail` is `number | string`, and `key` holds it.
- **`breaks.mts` and `unknownElements.mts`** read the understood set. No test moved for it.

**The raise's other pins,** which this work flipped so the suite could run:
- CR-026's seven files read through (`WAITS_ON_THE_RUNTIME` gone), `76t71` pinned as a known difference, as CR-026 D10 expected.
- CR-025 section 1's two tests: objects' CR-019 keeps `NumberingImplicitNumId.docx`'s 15 attributes, so no fixture reports any, and all 16 of `w:stylePaneFormatFilter`'s survive a change.
- CR-016's gap: the facade's walk reaches a container with no `TYPE_NAME`, so the test asserts the refusal.

**Tests.** `test/nothing-selected.test.mjs`, ten tests, over `test/nothing-selected/`: section 7's Word files, 318 of them, 1.5 MB. The suite's rows went from 786 to 795, the nine REQ-303 rows of case 6, which the adapter now claims. `PENDING` lost its five `req303` entries and gained `wpi` on `extraction-probe-inspect`'s `body/41`. Seventeen deliberate breaks of the build were each caught, on a green baseline: the removal off and at a root, every part, a kept host not first, the cell's paragraph off and fooled by a `w:tcPr`, `textOfBranch`'s old rule, the record off, `removed` and `root` false, the macro-enabled type, the walk's warning off, the sweep off, taking everything and ignoring `warnings: false`, the main part left out of the sweep at the early exit, and `key` without `detail`.

**D8's Word check** (`cr029-word-check`, Word 16.0.20326.20232 through the bridge, 55 files twice, 2026-11-07). This package's `mcePreprocess: true` saves of the 24 files that select nothing, those of the 16 new namespaces and the rounds' controls, `rw-24`, `te-00` to `te-05`, and three with a paragraph added through `Body`.
- **54 open silently, with no dialog,** and the two runs agree on all 55. `te-04` is refused, as Word refuses its source.
- **Words:** each save's text is Word's reading of its source, and the three edited files hold the new paragraph once.
- **Markup,** read where Word saved the source too (`rw-*`, `te-*`): `rw-24`, `rw-25` and `te-05` are the same, leaving rsids and the section aside. Four differ, for two reasons:
  - `te-01` to `te-03`: Word's repair of the source writes `w:spacing` on the cell's paragraph, and Word's save of this package's plain `w:p` has none. The word "repair" was too loose in this CR's first draft: the cell gets a paragraph, not Word's exact paragraph.
  - `te-00`, the control: Word writes `w:tblLook` as `0000` from the source and `04A0` from this package's save. CR-029 isn't the cause, since the file holds no such element and this build marshals it byte for byte as the build before it does. The two differ in the root's unused declarations and in `_rels/.rels`. What Word reacts to isn't established.

**Left as it stands.**
- `Body.text` on a kept load of `te-01` is `BEFORE\tRIGHT\r\nAFTER`, the cell holding no paragraph to break before, where a resolved load gives `BEFORE\r\tRIGHT\r\nAFTER`. Word repairs such a file, and its reading of the cell is the resolved one.
- `te-04`'s row with no cell, as section 4 says. A save of it is refused, as its source is.
- A `w:sdt` in a cell whose content a removal empties gets no paragraph, and nothing measures it.

**The raise (2026-11-08, Jason: "raise").** Objects 0.8.0 is out (tag `0.8.0`, `d62959e`, which is `9a2a945` and its version line), with `@docx4j/jsonix` locked at 3.13.0. The published tarball differs from the pack tested before its release only in that line. Typecheck passes, and 1,917 of 1,917 on Node 18.20.8, 20.20.2 and 22.23.4, the first run of this CR on 18 and 20. All 1,224 files read for the comparison save the same on the published tarball as on that pack. `CLAUDE.md`'s floor sentence and CR-026 phase 3's records go with this commit. The editor was told first, and has nothing to hold: its callers handle `undefined`, and `alternate-content.test.mts:134` to `:136` move when it installs 0.7.0.

The cell paragraph went in as the plain `w:p` D4 decided. Word's repair writes `w:spacing w:after="0" w:line="240" w:lineRule="auto"` on it, and the `cell-paragraph` round (section 12) showed that spacing fixed, so D9 writes it.

## 12. The `cell-paragraph` round (2026-11-08)

Jason, 2026-11-08: "Cell paragraph: let's do the small Word round". D8's check found that Word's repair of a `w:tc` with no block (`te-01` to `te-03`) puts `w:spacing w:after="0" w:line="240" w:lineRule="auto"` on the paragraph it adds, where D4's is a plain `w:p`. Those sources had no styles part, so the question was whether that spacing is fixed or taken from the document. Eight files, each a one-cell table between BEFORE and AFTER, each opened twice through the bridge (Word 16.0.20326.20232, bridge `4dcb436`), the two runs alike in every file. The portfolio's two controls ran silently too. Word's saves and PDFs are in the share's `bridge/portfolio/cp-*`. AFTER's top is read from each PDF in points (`pdftotext -bbox`).

| File | Holds | Word | The paragraph Word's save holds in the cell | AFTER |
|---|---|---|---|---|
| `cp-00` | The cell holding words; no styles part | Silent | As written | 121.98 |
| `cp-01` | An empty cell; no styles part | Repaired | `w:spacing w:after="0" w:line="240" w:lineRule="auto"` | 111.66 |
| `cp-02` | The same; document defaults after 0, line 240 | Repaired | A plain `w:p` | 97.70 |
| `cp-03` | The same; defaults after 400, line 480 | Repaired | `w:spacing w:after="0" w:line="240" w:lineRule="auto"` | 130.37 |
| `cp-04` | The same; no default spacing, Normal after 300, line 360 | Repaired | The same spacing | 119.09 |
| `cp-05` | `cp-03`'s defaults, and a table style of after 120, line 300 | Repaired | The same spacing | 131.57 |
| `cp-11` | `cp-03`'s defaults; the cell holding a plain `w:p`, as D4 writes | Silent | As written | 163.01 |
| `cp-12` | `cp-03`'s defaults; the cell holding `cp-03`'s spaced `w:p` | Silent | As written | 130.37 |

- **The spacing is fixed, not taken from the document.** Word writes after 0, line 240 (single) whatever the defaults, the Normal style or the table style say. It writes nothing where the paragraph already reads so (`cp-02`), so the rule is that the paragraph reads single-spaced with no space after.
- **It shows.** With wide defaults, D4's plain `w:p` makes the row 32.6pt taller than Word's repair (`cp-11` against `cp-03`). Word's spaced paragraph, opened silently, stands exactly where Word's repair does (`cp-12`). `cp-05`'s extra 1.2pt is its table style's borders.
- **Not measured:** `w:before` and `w:contextualSpacing` in the defaults, and a plain `w:p` in a package with no styles part on its own (`te-01`'s row takes its height from the other cell, and Word's `te-*` jobs saved no PDF). By the same rule a plain `w:p` there would take Word's defaults, after 160 and line 278, and stand as `cp-00`'s cell of words does, about 10pt below `cp-01`. That is inferred.

So D9 asked whether to write that spacing. Always writing it renders as Word in every case measured, and differs from Word's markup only where the defaults already read so (`cp-02`). Matching that too would need the effective spacing in a DOM pass that runs before the styles part is read.

**Built (2026-11-08, D9).** `keepACellsParagraph` writes the spacing. `test/nothing-selected/` holds the round's sources and Word's saves (`cp-*`, its `-a` run, with the manifest's SHA-1s). `nothing-selected.test.mjs` holds Word's repaired paragraph to that spacing in every repaired file but `cp-02`, and the paragraph the pass writes in `te-01` to `te-03` to Word's save of each.
