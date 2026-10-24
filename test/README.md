# Tests

`npm test` builds and runs every `test/*.test.mjs` with Node's test runner. One file:
`node --test test/roundtrip.test.mjs` (after `npm run build`).

`office-js-subset.ts` is the compile-time Office JS promise (CR-002 section 3.4), compiled by
`npm run typecheck`; `npm run generate` turns it into `src/office-js/supported.generated.mts`,
which is `Word.supported`. Add a member to the subset and regenerate in the same commit.
`office-js.test.mjs` covers the `Word` shim and `toApiScript` (CR-002 phase I).

Fixtures under `fixtures/` come from docx4j's `docx4j-core-tests` resources (Apache-2.0):
`loadAndSave.docx` (a current Word: comments, headers and footers, a chart with an embedded
workbook, PNG and SVG images, custom XML, `mc:AlternateContent`), `HelloWordOnline.docx` (main
part named `document22.xml`), `header-no-rels.docx`, `hyperlink_dupe.docx`,
`loadAndSave.pptx`, `loadAndSave.xlsx`, `invoice.docx` (from `OpenDoPE/`: content controls at
block, row and cell level, with data bindings and a table, the fixture of the content-control
tests), `comments-two.docx` (from `AlteredParts/`: two comments with no w15, w16cid or
`w:people` parts, so the comment API must cope without them), `invoice2013.docx` (from
`docx4j-samples-docx4j/sample-docs/databinding/`: the typed content controls, a `w14:checkbox`, a
`w:date`, a `w:picture` and two `w15:repeatingSection`s, all data bound — the fixture of the
phase E typed-kind tests; no document in the docx4j checkout has a `w:dropDownList`,
`w:comboBox` or `w:group` control, so those are built by the API in `customxml.test.mjs`), and
`mc-alternate-content-header.xml`, a Word 2010 flat OPC package (`pkg:package`) with
`mc:AlternateContent` in a header. `tracked-changes.docx` is docx4j's
`docx4j-samples-docx4j/sample-docs/sample-docx.docx` (also Apache-2.0), a Word file with one
`w:ins` and one `w:del` by "Jason Harrop", the fixture of `tracking.test.mjs`.
`invoice_Saxon_XPath2.docx` is docx4j's `sample-docs/databinding` template of that name, whose
conditions are the cases where the OpenDoPE boolean conversion modes disagree (`wantspam` holds
`false`) and where XPath 2.0 syntax is required (`xs:date(...) > xs:date(...)`), the fixture of
`xpath-fonto.test.mjs` (CR-005 phase A).
`cr022-slicers-timelines.xlsx` and `cr022-checkbox.xlsx` are docx4j's fixtures of those names,
added by its CR-022 at `16844ff03`, the fixtures of `sml-extensions.test.mjs` (CR-004). The first
is Excel 365 with two pivot tables, two slicers with their caches, a timeline with its cache,
sparklines and a table; its `xl/drawings/drawing1.xml` and `drawing2.xml` frame the slicers and
the timeline through `mc:AlternateContent`, and since CR-004 phase B those Choices are taken and
the slicer content kept as DOM (it needs the lax `a:graphicData` wildcard of docx4j `8e8f6ea83`,
in objects 0.2.0). The second is a check box: its worksheet's only `mc:AlternateContent` requires
`x14` and has no Fallback, so it is what the `x14` half of phase B's `UNDERSTOOD_NAMESPACES`
change is measured on - but its check box is inert in Excel before any save (docx4j-generated, no
`fmlaLink`), so it cannot witness whether a re-marshalled control still works.
`cr022-checkbox-linked.xlsx`, added 2026-09-26, is the same workbook with the box linked to D4, and
is the fixture acceptance check 14 was settled on.
`revisions/revisions-word15.docx` is the docx4j-ts-editor's Word-made revision fixture of that name
(its `packages/editor-model/test/fixtures/revisions/`, ED-005 section 13, at its `a0e1c4b`):
the engine built the text, and Word 15 made every revision in it, as Author A and Author B, by
following the instructions in the document. Section 3 is a move - a paragraph cut and pasted above
the one before it - so it carries Word's `w:moveFrom`, `w:moveTo`, their range markers and the
move's `w:name`, which is what CR-002 section 29 is measured on (check 16). It sits in a directory
of its own so that neither Java harness makes a golden of it. `revisions/check18/` holds what Word
made of it in check 18 - one action each through the Review tab, on a fresh copy, saved by Jason
2026-09-28 - the oracle of `tracking.test.mjs`'s move tests.
`revisions/review/` holds three of the docx4j-ts-editor's Word saves from its E4.g session (its
`packages/editor-model/test/fixtures/review/`, at its `1c90869`, 2026-09-30):
`review-for-word-word15.docx`, the editor's review file saved untouched by Word 15, and Word 15's
Accept All and Reject All saves of the same file (`-accepted`, `-rejected`). They are the oracle of
`tracking.test.mjs`'s footnote case (CR-002 section 44). Beside them, from docx4j's
`docx4j-core-tests/src/test/resources/revisions/editor/review/` at its `a8bc6cda9`:
`review-for-word.docx`, the editor's review file itself, and `review-for-word-word365-rejected.docx`,
Jason's Word 365 Reject All of it (2026-10-21), which removed the comment "Changed from arrears, as
agreed." whose markers lie outside its `w:ins`, where Word 2010 and Word 15 kept it. They, and the two
directories below, are the oracle of `tracking.test.mjs`'s comment cases (CR-002 section 54).
`revisions/python/artefact8/` holds docx4j-python's acceptance artefact 8 (`8-tracked-changes.docx`,
a comment on inserted "second" with every marker outside the `w:ins`) and Word 365's Accept All and
Reject All saves of it (`8-word-accepted`, `8-word-rejected`; Jason, 2026-10-21), from docx4j's
`revisions/python/artefact8/` at `a8bc6cda9`. `revisions/cr043/comments/` holds docx4j-25's probes U21
to U23 for docx4j CR-043 section 15 and Word 365's saves of them (Jason, 2026-10-21), from docx4j's
`revisions/cr043/comments/` at `a8bc6cda9`: U21, Accept All of a deletion whose comment's markers lie
outside the `w:del`; U22, one insertion rejected, which removes the only comment while a deletion
stays tracked; U23, Reject All of two inserted paragraphs commented whole.
`revisions/cr043/` holds three files from docx4j's CR-043 phase 0 probe (its
`docx4j-core-tests/src/test/resources/revisions/cr043/`, at its `5e895232f`; Word 16.0.20326,
2026-10-16, driven through Word's object model by its `cr043-probe.ps1`): `cr043-tracked.docx`, the
probe's tracked edits on a base document (its `cr043-log.txt` lists each case), and Word's Accept
All and Reject All saves of it (`cr043-accept-all`, `cr043-reject-all`). They are the oracle of
`tracking.test.mjs`'s bookmark case (CR-002 section 45), its parts case (section 46) and its
every-story case (section 49). Beside them, `tracked-changes-equations.docx` is docx4j's file of that
name (its `docx4j-core-tests/src/test/resources/`, at its `992c8d572`, issue 348), and
`equations-accept-all.docx` and `equations-reject-all.docx` are Word's Accept All and Reject All saves of
it from the same probe run (at `5e895232f`): the oracle of the equations case (section 51).

`onoff/pandoc-code.docx` is what pandoc 3.6.1 makes of `onoff/pandoc-code.md` beside it
(`pandoc pandoc-code.md -o pandoc-code.docx`, 2026-10-17): its table's header row has
`w:tblHeader w:val="on"` and its Source Code style `w:wordWrap w:val="off"`, ST_OnOff spellings
that @docx4j/jsonix 3.4.0 refused. It is the fixture of `load.test.mjs`'s on/off case.

`charts/waterfall.docx`, `.xlsx` and `.pptx` are docx4j's `docx4j-core-tests/src/test/resources/charts/`
files of those names (`f9708f5b3`, the xlsx `e560d03f3`; its `charts/README.md` describes them): a
waterfall chart made by Office, whose frame is in an `mc:AlternateContent` with `Requires="cx1"`. They
are the fixtures of `mce.test.mjs`'s chartex case (CR-009 section 8).

The images in `content-c.test.mjs` are base64 constants rather than files: a 4 x 3 PNG at 96 dpi
(a real one, deflated with `node:zlib`), a 2 x 2 GIF87a, a 2 x 2 24-bit BMP at 3780 px/m, and a
JPEG header (SOI, a JFIF APP0 at 300 dpi, SOF0, EOI) for the header reader.

## Parity fixtures (`fixtures/parity/`, CR-001 Phase B step 0, 2026-09-19)

Thirty-seven documents (fifty-three since CR-007, below) the Java harness (`test/java/`, step 1) turns into goldens under
`test/golden/`; `parity-fixtures.test.mjs` checks each loads, unmarshals fully and round-trips
untouched byte for byte (relationships parts and `[Content_Types].xml` excepted, as always).
All are docx4j's, Apache-2.0, from `VERSION_17_1_1` at `0077d749c`:

- `styles-*.docx` (5), `fonts-*.docx` (12), `numbering-*.docx` (10): the verification probes
  of docx4j CR-015, CR-016 and CR-014, generated by `docx4j-layout-fidelity`'s `Corpus.java`
  (`java -cp "target/classes:target/lib/*" org.docx4j.fidelity.Fidelity generate <dir>`; the
  Word goldens for them are in that module's `goldens/word/`). Regenerated at `0077d749c`, so
  not byte-identical to the six docx4j-python copied earlier under the same names (root namespace
  declarations differ, and `numbering-stories.docx` in attribute order inside a VML fallback);
  everything the resolver, the emulator and the selector read is identical, so one golden serves
  both.
- From `docx4j-core-tests/src/test/resources/`: `numbering_indentation.docx`,
  `numbering_indentation_firstline.docx`, `NumberingImplicitNumId.docx`, `Mac_OSX_Fonts.docx`.
- From `docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/`:
  `article-section-isLgl.docx`, `article-section-NotIsLgl.docx`, `NumberingIndents.docx`,
  `startOverride.docx`, and `numbering-stories.docx` renamed `numbering-stories-coretests.docx`
  (the probe corpus has a `numbering-stories.docx` of its own).
- From `docx4j-samples-docx4j/sample-docs/`: `fonts-modesOfApplication.docx`.

Added for CR-007 (the table context, 2026-10-04), sixteen more:

- `tables-banding-merged-row.docx` and `tables-banding-band-size.docx`: docx4j CR-030's probes T9
  and T10, from the same share (CR-007 section 7).
- `tables-*.docx` (12): the Word probes of docx4j CR-030 (its section 5, T1 to T8), generated by
  `docx4j-layout-fidelity`'s `Corpus.java` at `1954a7b89` and copied from the Word share's
  `fidelity/corpus/`, where Jason cut the goldens of 2026-10-03. They settle, in Word, the table
  style decided by name, the story boundary (a text box, a footnote) and the [MS-DOCX] size and
  justification exception by compatibility mode.
- `tables-conditions.docx` and `tables-conditions-mode12.docx`: made by
  `scripts/make-table-conditions-fixture.mjs` for the arithmetic of `TableStyleConditions` the
  probes do not reach (every look, band sizes, `w:cnfStyle` caches, spans, `w:gridBefore`, a nested
  table, content controls around rows and cells, paragraph and character styles over the table
  level). **Not Word-measured:** their goldens are what docx4j answers.

## Numbering indent fixtures (`fixtures/ind/`, CR-001 Phase B step 3)

The seven flat OPC documents of docx4j's `ListNumberIndTest`, copied unchanged from
`docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/ind/` at `VERSION_17_1_1`
(Apache-2.0): `abstract_style_with`, `abstract_style_without`, `abstract_nostyle_ppr`,
`abstract_nostyle_noppr`, `override_nostyle_ppr`, `abstract_style_ind_only` and
`abstract_style_basedon`. Each is one `w:num` 1 whose level 0 states its indent, or its linked
`w:pStyle` does, or neither; `numbering.test.mjs` asserts the `w:ind/@w:left`
`NumberingDefinitionsPart.getIndOf('1', '0')` answers (2880, 12, 13, none, 23, 11, 31).

## Level owner fixtures (`fixtures/listnumbering/`, CR-012 phase 1)

The three documents of docx4j's `LevelOwnerTest`, copied unchanged from
`docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/` at `1d108c0af` (added in
`d9acdf5e2`, Apache-2.0): `numbering-linked-level-other-style`, `numbering-unlinked-level-claim` and
`numbering-shared-abstract-styles`. They are docx4j's layout-fidelity probes as Word 365 converted
them, with several paragraph styles carrying one numbering level. `numbering.test.mjs` asserts Word's
labels, read off Word's PDFs, through the `Emulator` and `Body.listLabels` (REQ-510: a level
belongs to one paragraph style). They are outside `fixtures/parity/`, so the parity harness writes
no golden for them, and they hold no revision markup.

## The text extraction suite (`text-extraction/`, CR-012 phase 0)

docx4j's text extraction conformance suite and the agreement test's fixtures, copied unedited, with
the Node runners that mirror docx4j's `ExtractionConformance` and `ViewAgreementTest`. Its own
`README.md` gives the provenance and how to copy it again. It is kept outside `fixtures/` because
the weekly `RevisionsHarness` writes a golden for every `.docx` with revision markup under
`fixtures/`.

## The `mc-fun` probes (`mc-fun/`, CR-013 phase 0)

docx4j's 30 probes of `mc:AlternateContent` at every level (run, paragraph, block, row, cell, in content
controls and around text boxes) and Word 365's saves of them, copied unedited from docx4j at `69af8cf42`
(added in `1cff2cd5c`, docx4j CR-051, Apache-2.0). Their `README.md` says what each holds and
`manifest.json` gives each file's SHA-1, which `mc-fun.test.mjs` checks. A Word save shows the branch
Word took, so the test holds every reader to it: the text model and, since CR-013 phase 2, the Office JS
readers, kept and resolved. Since phase 3 it also holds the mutators to docx4j's
`TrackedChangesAlternateContentTest`, built in memory as docx4j builds them, and to every branch. They are
kept outside `fixtures/` for the reason the suite is.

## The theme colour vectors (`theme-colors/`, CR-008 phase 2)

docx4j's vectors for its colour resolver (`org.docx4j.model.color`, its CR-034), copied unedited from docx4j at
`1fffc20a3` (added in `3dbe4eee5`, Apache-2.0): seven probe files of 1,068 labelled cases (WordprocessingML colours whose
written value disagrees with their theme reference, a swapped `w:clrSchemeMapping`, 403 DrawingML transform cases in Word and
in PowerPoint, SpreadsheetML theme, tint and indexed fills, a slide's colour map override), what Word, PowerPoint and Excel 365
drew for each (`readings.tsv`), and the 117 distinct tint and shade triples of four real-document corpora (`word-tint-shade.tsv`).
Their `README.md` describes them; `manifest.json` gives each file's SHA-1, which `theme-colors.test.mjs` checks. Kept outside
`fixtures/` so that no fixture-wide test or harness counts the probes.

## The run property probes (`runprops-word365/`, CR-015)

docx4j's probes for run properties Word 365 reads leniently (its CR-060 and CR-061), copied unedited from docx4j at
`5163c2986` (`docx4j-core-tests/src/test/resources/runprops/word365/`, added in `87e6ace56`, Apache-2.0): Jason's seven
rounds of 2026-10-24, 33 probes and Word 365's save of each (`<name>-word365.docx`), with docx4j's `README.md` of what each
round tests. Each probe paragraph after the first is `label: MARKER (end)`, and only the marker run carries the property
tested. `manifest.json` gives each file's SHA-1, which `runprops-word365.test.mjs` checks. Kept outside `fixtures/` so that no
fixture-wide test or harness counts the probes; check 60's file (`fixtures/check60/build.mjs`) is built from 25 of them.

## Check 52's Word saves (`check52/`, CR-014 phase 0)

Word 365's two saves of check 52's file after its Script Lab run (2026-10-22): `52-ranges-word365.docx` as the run left
it, and `52-ranges-word365-accepted.docx` after Accept All. `manifest.json` gives each one's SHA-1, which
`ranges.test.mjs` checks, with the Word build and the view. They hold Word's own markup for ranges across paragraphs, tracked
and not, and for CR-002 section 55's cells. `ranges.test.mjs` accepts the first with `acceptAllRevisions()` and compares
it with the second, region for region. They are kept outside `fixtures/` because they hold tracked changes, for which the
weekly `RevisionsHarness` would cut CR-011 goldens. Check 52 in the manual checklist below has the run.

## Parity goldens (`golden/`, CR-001 Phase B step 1)

What docx4j answers for every fixture above and for the eight `.docx` in `fixtures/`, written
by the Java harness in `java/` and read by `parity.test.mjs`. 62 goldens at harness version 4
from docx4j `VERSION_17_3_1` at `0349796f9` since CR-007 (2026-10-04; 45 at version 3 from
`VERSION_17_1_1` at `7fba7a150` when Phase B landed); the resolver, the emulator and the selector
reproduce every recorded answer. `golden/README.md` is the review guide (counts per golden, and
what was found producing them); `java/README.md` says how to build docx4j, run the harness and
read a golden. `.github/workflows/parity.yml` reruns it weekly against docx4j's head and opens
a pull request when an answer changes.

`golden/xpath/` holds CR-006's oracle, made by the same project's `XPathHarness` and read by
`xpath-parity.test.mjs`: what docx4j's `getJAXBNodesViaXPath` selects for 47 expressions over each
fixture's main document part, from the docx4j 17.2.1 release (the 46 of CR-006) and the 17.3.1
release (the 16 table fixtures of CR-007, added 2026-10-08). `selectObjects` returns the same hits
in the same order, with objects of the same classes, on all 58 fixtures that have hits.

## Word acceptance (manual)

Last run: 2026-09-19, Word 365, PowerPoint and Excel, after CR-001 Phase C: checks 3 and 9 to 13.
What passed: check 3's theme fonts (Design > Fonts showed the Office theme each `defaultTheme`
names, and the body text in the theme's body face), check 9 (Word showed **Jane Doe** from the
custom XML part, and `invoice2013.docx`'s checkbox, date, repeating sections and picture control
all still worked), check 11's list labels and the shared-abstract edit, check 12's slide sizes,
master, layout and theme fonts, check 13's tab order and values, and every file but one opened
without a repair prompt. Five things were found, and are fixed in the same change as this record:

1. **Every created docx opened in compatibility mode** (check 3). `createPackage()` wrote an
   empty `w:settings`, so Word 365 treated the document as pre-2013. It now writes
   `w:compat/w:compatSetting compatibilityMode` 15 (CR-001 section 17).
2. **A tracked row deletion showed no struck-out content** (check 10). Word marks a deleted row
   with `w:trPr/w:del` *and* every cell's content deleted; only the `w:trPr` was written
   (CR-002 section 13).
3. **The created pptx's slides were blank** (check 12): the slide's `p:spTree` was empty, so
   PowerPoint had no "Click to add title" to offer. The layout and every created slide now carry
   a title and a body placeholder (CR-001 section 17).
4. **`check13-remarshalled.xlsx` was the one file Excel offered to repair** (check 13): the
   re-marshalled `xl/workbook.xml` kept `mc:Ignorable="x15 xr xr6 xr10 xr2"` while declaring only
   `x15` and `xr`. `XmlPart` now re-declares an ignorable prefix the model dropped
   (CR-001 section 17).
5. **Check 13's expectation was wrong**, not the output: the script inserts `Cover` at index 0
   after writing the values into `Sales`, so values on the second tab is right. The check text
   below now says so.

Checks 3, 10, 12 and 13 are to be re-run against the fixed code. Check 14 (added by CR-004 phase B,
2026-09-25) ran on 2026-09-26 and **passed on both halves**, the form-control half after being
re-run on a live control - see the check itself and CR-004 section 7.

Previous run: 2026-09-16, Word 365 version 2608 (build 20326.20144, Click-to-Run), after CR-002
phases C, G and I: checks 6 to 8, with the script below, all passed. Check 7 as expected shows the fixture's own text, not the
replacement, on the document surface: the control is bound to a custom XML part, and Word
re-reads a bound control from the XML part on open, so the replaced `w:sdtContent` is
overwritten by the stale binding. The control, its title and its binding survived, which is
what the check is for; writing through the binding is CR-002 phase E (`XmlMapping`).

Previous run: 2026-09-10, Word 2016, after CR-001 Phase A and CR-002 phases B and D. All four
files below opened without a repair prompt: the untouched round trip, the round trip with the
main part and a header re-marshalled and a paragraph added through the content API, a
document created from nothing with a heading and a formatted span, and the flat OPC of the
second opened as a `pkg:package` file.

Saved output must open in Word without a repair prompt. After a change to marshalling,
namespaces, content types or the zip writer, check by hand:

**Saving each step of a check under its own name: Save As first, then make the change, then
Ctrl+S.** Word's AutoSave writes a change back to the file that is open as soon as it is made, so
"make the change, then Save As the new name" overwrites the file the step started from. Where a
check below says "carrying on from", "from a fresh copy of" or "save as", do it in that order: open
the starting file, File > Save As the step's file name, make the change, Ctrl+S. Where the cases
all start from the same input (check 29 on), the copies are made beforehand, one per case and
already named, so each case starts by opening its own file.

1. Round trip: load `fixtures/loadAndSave.docx`, save, open in Word. No repair prompt; comments,
   headers, footers, the chart and both images still there.
2. Re-marshalled main part: as 1, but `await pkg.getMainDocumentPart().getContents()` before
   saving. Word must accept the `mc:Ignorable` prefixes and the `xml:space` attributes.
3. New document: `WordprocessingMLPackage.createPackage()` with a paragraph added, saved, opened.
   The title bar must **not** say Compatibility Mode: since CR-001 section 17 the settings part
   carries `w:compat/w:compatSetting compatibilityMode` 15, as Word 2013 and later write (File >
   Info > Convert must be absent, and Word's own re-save must not change the layout).
   Styles pane shows Normal and Heading 1 to 4. Since CR-001 Phase B step 4 the package also
   carries a theme part (`word/theme/theme1.xml`, nine parts in the zip instead of eight), so
   Design > Fonts must show the Office theme - Aptos Display / Aptos by default, and Calibri
   Light / Calibri or Cambria / Calibri for `createPackage({ defaultTheme: '2013' | '2007' })` -
   and the body text must be set in the theme's body face rather than Word's fallback.
4. Flat OPC: `saveFlatOpc()` output pasted through `insertOoxml()` in a Word add-in (or saved as
   `.xml` and opened in Word, which reads `pkg:package` files directly).
5. pptx and xlsx: round trips of the two fixtures open in PowerPoint and Excel.
6. Tables and pictures (CR-002 phase C): a package created from nothing with
   `body.insertTable(3, 2, 'End', values)`, `table.styleBuiltIn = 'TableGrid'` and
   `table.headerRowCount = 1`, then `body.insertInlinePictureFromBase64(png, 'End')`. Word must
   show a bordered table whose first row repeats on a page break, and the picture at its natural
   size (4 x 3 px at 96 dpi is 3 x 2.25 pt; use a larger image to judge it), with the alt text in
   Format Picture. Check `insertOoxml` of that package's `saveFlatOpc()` into another: the image
   appears once, from `word/media/image1.png` of the target.
7. Content controls: `fixtures/invoice.docx` loaded, `body.contentControls[0].insertText(...)`,
   saved; Word must still show the control and its data binding. Expect the fixture's text, not
   the replacement, on the surface: the control is bound, and Word refreshes a bound control from
   the custom XML part on open (the replaced text is in `word/document.xml`, and is what a
   non-bound control would show). Until phase E writes through the binding, check the XML.
8. Comments (CR-002 phase G): in a package from `createPackage()`, comment a range
   (`range.insertComment('...')`), reply to it and resolve it; in `fixtures/loadAndSave.docx`,
   comment a paragraph and delete Word's own comment. Saved and opened in Word: the review pane
   shows the comment, the reply under it in the same thread, the resolved thread greyed, the
   author and initials from `pkg.author`, and the deleted comment gone. (Before objects 0.1.3 a
   re-marshalled comments part lost `mc:Ignorable` on `w:comments`; it is kept now, CR-002
   section 14.)
9. Custom XML and the bindings (CR-002 phase E): `fixtures/invoice.docx` loaded,
   `await pkg.customXmlParts.load()`, then `insertText('Jane Doe', 'Replace')` on the customer-name
   control (`contentControls[0]`) and `await pkg.customXmlParts.updateFromContentControls()`, saved.
   Word must show **Jane Doe** on the document surface when it opens the file — this is the check
   that item 7 could not pass: Word refreshes a bound control from the custom XML part on open, so
   the value has to be in `customXml/item3.xml` as well as in `word/document.xml`. The three bound
   controls, the repeat and the two conditions must survive. In the same file, check
   `invoice2013.docx` with `applyBindings()` after a change to the custom XML: the checkbox, the
   date and the repeating sections keep working in Word, and the picture control (which phase E
   leaves alone) still shows its image.

10. Tracked changes (CR-002 phase F): with `pkg.author` set and `pkg.changeTrackingMode =
   'TrackAll'`, make one of each kind — insert text, delete text, replace a word, change a
   run's formatting and a paragraph's style, add a paragraph, delete a paragraph, add a table
   row and delete another — then save and open in Word. The deleted row must show **struck out**,
   cell text and all (CR-002 section 13: `w:trPr/w:del` plus a `w:del` around every run and
   `w:pPr/w:rPr/w:del` on every paragraph mark in it), and the inserted row underlined.
   No repair prompt; Word must show every
   change in the review pane under the author from `pkg.author`, with the date from
   `pkg.trackedChangeDate`; Accept All must leave the document reading as `body.text` did, and
   Reject All as `body.getText({ view: 'original' })` did. Check too that Word's own Accept All
   and this package's `acceptAll()` agree on the same file, and that a comment inserted while
   tracking is on is a comment only: no `w:ins` around its reference run or range markers.

11. Lists (CR-002 phase H): in a package from `createPackage()`, `await p.startNewList()` on the
   first paragraph, `attachToList(list.id)` on the next few with `listItem.level = 1` on two of
   them, `p.restartList()` on a later one, and a second list from
   `startNewList({ bullet: true })`. Saved and opened in Word, the labels must be the ones
   `listItem.listString` reported: 1. a. b. 2. for the nested run, the restarted list beginning
   at 1. again while Word still treats it as the same list (Home > Multilevel List shows one
   definition, and Continue Numbering on the restarted item joins them), and the bulleted list
   showing Word's solid bullet in Symbol. Check `list.setLevelNumbering(0, 'UpperRoman')` and
   `list.setLevelBullet(0, 'Square')` on a list that shares its `w:abstractNum` with a restart:
   only the edited list may change in Word.

12. A created presentation (CR-001 Phase C): `PresentationMLPackage.createPackage()`, saved and
   opened in PowerPoint. No repair prompt; one slide in the Slides pane, at the slide size
   asked for (View > Slide Master shows the one master with its one layout, and Design > Variants >
   Fonts the theme `defaultTheme` names). The slide must show **Click to add title** and
   **Click to add text**, both clickable and typable, at the layout's positions (CR-001 section
   17); `addSlide({ title, body })` must show that text instead, one paragraph per body line.
   Check `createPackage({ slideSize: 'SCREEN16x9' })` and `{ slideSize: 'A4', landscape: false }`,
   and a second slide from `addSlide()`.

13. A created workbook (CR-001 Phase C): `SpreadsheetMLPackage.createPackage()` with
   `createWorksheetPart('Sales')` and a few rows written into `sheetData`, saved and opened in
   Excel. No repair prompt. The script below writes the values into `Sales` and then inserts
   `Cover` at index 0, so Excel must show three tabs in `sheets` order — **Cover first and empty,
   Sales second with the Q1 and Q2 rows in A1 down, Notes third and empty** — which is what
   `createWorksheetPart(name, index?)` inserting rather than appending means. Also round-trip the
   two fixtures (check 5) after unmarshalling their main parts, to see a re-marshalled
   `ppt/presentation.xml` and `xl/workbook.xml` open (the workbook loses Excel's
   `x15ac:absPath` and its `xr:revisionPtr`, which is expected: CR-001 sections 16 and 17).
14. Excel's extension parts (CR-004 phase B): load `fixtures/cr022-slicers-timelines.xlsx`,
   `await pkg.unmarshalAll()`, save and open in Excel. No repair prompt; both slicers and the
   timeline must still be there and still filter their pivot tables (the drawings now keep the
   slicer graphic instead of falling back to the "supported in Excel 2010 or later" placeholder
   box). Then the same with `fixtures/cr022-checkbox.xlsx`: the check box must still be on the
   sheet and still toggle. That one is the check worth making: resolving the `mc:AlternateContent`
   on load cannot rebuild the wrapper, so the saved worksheet holds a bare `<controls>` where
   Excel writes `mc:AlternateContent > mc:Choice Requires="x14" > controls`. It is schema-valid
   (`CT_Worksheet` has `controls`). **Accepted 2026-09-26 (Excel 365), measured on
   `fixtures/cr022-checkbox-linked.xlsx`:** the box toggles and drives D4 after a whole-package
   re-marshal, so Excel does not need the wrapper and resolve-on-load stands (CR-004 section 7).
   Since CR-010 phase 4 (2026-10-18) the default load keeps the wrapper and writes it back, which is
   Excel's own markup and docx4j's (its CR-021 section 8.13 found Excel accepts both forms); the bare
   form this check measured is `mcePreprocess: true`'s. Use
   the linked fixture for this check, not `cr022-checkbox.xlsx`, whose check box does not toggle
   *before* a save either - it is docx4j-generated with no `fmlaLink` and no `x:FmlaLink`, and three
   rounds of bisecting over it measured nothing. **The lesson for any check added here: name a
   fixture whose feature is known to work before the save.** Files at
   `fidelity/cr004b-core-ts-controls/` on the share.
15. A new hyperlink's style (CR-002 section 25), **Word's own answer, before the engine changes**.
   What Office JS's `range.hyperlink = ...` does in Word decides what this package's setter does,
   since the editor's `api` lines are held to it. Run the Script Lab snippet below (Insert > Get
   Add-ins > Script Lab, then Script Lab > Code > New snippet: paste the Script and HTML tabs,
   Run, press **Run the check**) **in a new blank document**, and copy the text box's JSON back.
   It makes five paragraphs, each its own case with its own `sync`, so one failing case does not
   hide the others: A an external link over plain text, B an internal `#` link, C a link over
   text that already has a character style (`Strong`), D a link over direct formatting (bold, red),
   E a link set and then removed with `""`. For each it records what Office JS reports back
   (`style`, `styleBuiltIn`, the font's colour and underline) and the paragraph's XML from
   `body.getOoxml()`, and from the same package the `Hyperlink` and `FollowedHyperlink` style
   definitions Word wrote. The questions it answers: does Word give the wrapped runs
   `w:rStyle w:val="Hyperlink"`, what happens to a run that already names a character style and to
   direct formatting, does an internal link get the style too, does removing the link remove the
   style, and what the style definition is.
16. A tracked move, accepted and rejected (CR-002 section 29), and a hyperlink over part of another
   (section 25), **Word's own answers, before the engine changes**. Open a **copy** of
   `fixtures/revisions/revisions-word15.docx` in Word and run the second Script Lab snippet below
   (Script tab, HTML tab, Run, then **Run the check** with "All" chosen), and copy the JSON back.
   It needs WordApi 1.6 (`getTrackedChanges`, `accept`, `reject`, `acceptAll`, `rejectAll`). It
   turns tracking off, keeps the body's OOXML, and for each of six scenarios puts that OOXML back
   (`insertOoxml`, "Replace"), records section 3 (the move) before, acts, and records it after,
   with the tracked changes Office JS reports: accept the move's first change in document order
   (its destination), accept its second (its source), reject each, accept all, reject all. The
   first thing it records is how Office JS lists the move - two changes or one, and of which types.
   If a scenario's "before" shows that `insertOoxml` did not bring the move back as a move, run the
   scenarios one at a time instead, each on a fresh copy, choosing it in the list. Then three
   hyperlink cases, appended as paragraphs: F links "alpha beta" and then "beta gamma"; G links
   "alpha beta gamma" and then "beta"; H links "alpha beta gamma" and then sets "beta" to `""`.
   Each records the paragraph's XML and the link each word reads afterwards. The questions: does
   accepting or rejecting one half resolve the other, what happens to the four range markers, how a
   move is listed, and what a link over part of an existing link does to the rest of it.
   **Run 2026-09-28 (Word 16.0.20326.20158, Windows): the links answered, the moves did not.** F, G
   and H are in CR-002 section 25 and `content.test.mjs`. Every move scenario failed with
   `GeneralException`: over this fixture Office JS listed **one** tracked change, a `Formatted` one
   whose text is the whole document - the fixture's `w:sectPrChange` (section 12, the page margins),
   a section-properties revision spanning the only section - and none of the other 23; the snippet,
   looking for the change holding the moved text, found that one and accepted it, after which every
   `getTrackedChanges` threw. The move is therefore asked again in a document holding nothing else,
   check 17. (`insertOoxml` with "Replace" did bring the move back intact each time.)
17. A tracked move on its own (CR-002 section 29): check 16's move scenarios in a **new blank
   document**. The third Script Lab snippet below inserts section 3 of the fixture exactly as check
   16 recorded it - the destination paragraph with `w:moveToRangeStart`, the paragraph it moved
   before, the source paragraph with `w:moveFromRangeStart`, the paragraph after, the two range ends
   between them, the move's `w:name` and Author A's revisions - through `insertOoxml` into an
   emptied body, once per scenario, and records how Office JS lists the move (every change, whole
   text), acts, and records the blocks and the changes after. Press **Run the check** and copy the
   JSON back. **Run 2026-09-28 (Word 16.0.20326.20158): Office JS cannot list a move.** The move went
   in intact every time (Word renumbering its ids 0 to 5), and the first `getTrackedChanges()` after
   it threw `GeneralException`, before anything was accepted or rejected. With check 16, where it
   listed only the revision before the move and threw once past it, that is this build's answer:
   there is no Office JS behaviour for a move to hold the engine to, so Word's own Review commands
   are asked instead, check 18.
18. A tracked move accepted and rejected through Word's Review tab (CR-002 section 29), by hand.
   Each scenario on a **fresh copy** of `fixtures/revisions/revisions-word15.docx`, saved under the
   name given, into `fixtures/revisions/check18/`; section 3 is the move, where "The second
   paragraph of the move" appears twice - its destination above "The first paragraph of the move",
   and its source, struck through, below it:
   - `accept-destination.docx`: click in the destination copy, Review > Accept > **Accept This
     Change**.
   - `accept-source.docx`: click in the struck-through source copy, Review > Accept > **Accept This
     Change**.
   - `reject-destination.docx`: click in the destination copy, Review > Reject > **Reject Change**.
   - `reject-source.docx`: click in the source copy, Review > Reject > **Reject Change**.
   - `accept-all.docx`: Review > Accept > **Accept All Changes**.
   - `reject-all.docx`: Review > Reject > **Reject All Changes**.
   Nothing else is changed in any copy. Note anything Word says or selects on the way (a message,
   both copies highlighted at once, the change it moves to). The saved files are compared with the
   fixture's section 3: whether acting on one half resolved the other, and what became of the four
   range markers and the paragraph marks. **Run 2026-09-28 (Word 16.0.20326.20158): a move is one
   change.** Accepting either half left the paragraph at its destination only, the source paragraph
   gone with its mark; rejecting either half put it back at its source, the destination paragraph
   gone with its mark; all eight move elements went every time, and nothing outside section 3
   changed (the same text in both views, 24 changes down to 22; Word's re-save merged runs and
   dropped `w:proofErr`, no more). Accept all and reject all did the same to the move. The files are
   `tracking.test.mjs`'s oracle.
19. A hyperlink set at a caret (CR-002 section 25), **Office JS's own answer, before the engine
   changes**. The editor's link dialog edits the link under a caret, as Word's Edit Hyperlink does;
   the setter removes that link and puts an empty `w:hyperlink` at the caret; what Office JS's
   `range.hyperlink = ...` does on an empty range is not known. Run the fourth Script Lab snippet
   below in a **new blank document** and copy the JSON back. Four paragraphs, each its own case: I, a
   caret inside a link ("alpha beta gamma" linked, the caret in "beta"), given a new address; J, a
   caret in plain text, given an address; K, a caret just after a link on "alpha", given another
   address; L, a caret inside a link, given `""`. For each: what the caret reads before and after,
   what each word reads afterwards, and the paragraph's XML. **Run 2026-09-28 (Word
   16.0.20326.20158): Office JS refuses a hyperlink on an empty range.** All four threw
   `GeneralException` and left the paragraph as it was, `""` included. Reading at the caret gave
   the link it lies in (I, L), `""` in plain text (J), and **the link just before it** for a caret
   right after one (K).
20. What Word does to a move that is edited, moved again, or missing a part (CR-002 section 29,
   asked by the editor for its tracker, E4.c step C3), by hand, as check 18 was. Section 3 of
   `fixtures/revisions/revisions-word15.docx` is the move: "The second paragraph of the move: it goes
   before the first." stands at its **destination** (the first of section 3's paragraphs) and,
   struck through, at its **source** (below "The first paragraph of the move"). Before starting, in
   Word 15: Review > Track Changes on, "Track moves" kept on in the tracking options, and File >
   Options > General with "Always use these values regardless of sign in to Office" ticked, so that
   the user name set there is the author: **Author A** (initials AA) made the move, **Author B**
   (BB) is another person. Each numbered case starts from a **fresh copy** of the file named, and is
   saved into `fixtures/revisions/check20/` as `NN-name.docx`. A case marked **+ accept/reject**
   is then taken on twice from that saved file: open it, click in the destination copy's original
   words (on "second"), Review > Accept > **Accept This Change**, save as `NN-name-accept.docx`;
   open it again, click in the same place, Review > Reject > **Reject Change**, save as
   `NN-name-reject.docx`. Note anything Word says, refuses or selects on the way. Change nothing
   else.

   | Case | Start from | As | Do | Then |
   |---|---|---|---|---|
   | 01-type-inside-a | the fixture | A | in the destination copy, click after "goes" and type " quickly" | + accept/reject |
   | 02-type-inside-b | the fixture | B | the same | |
   | 03-type-at-start | the fixture | B | click at the very start of the destination copy (before "The") and type "Now " | + accept/reject |
   | 04-type-at-end | the fixture | B | click at the end of the destination copy (after "first.", before the paragraph mark) and type " Really." | + accept/reject |
   | 05-type-in-source | the fixture | B | click inside the struck-through source copy, after "goes", and type "X" (note where the text goes, or whether Word refuses) | |
   | 06-delete-part-a | the fixture | A | in the destination copy select "it goes " and press Delete | + accept/reject |
   | 07-delete-part-b | the fixture | B | the same | |
   | 08-delete-all | the fixture | B | select all of the destination copy's text (not its paragraph mark) and press Delete | + accept/reject |
   | 09-delete-in-source | the fixture | B | select "it goes " in the struck-through source copy and press Delete (note what Word does) | |
   | 10-move-again | the fixture | A | triple-click the destination copy (its mark with it), cut (Ctrl+X), click at the start of "The third paragraph of the move", paste (Ctrl+V) | + accept/reject (click in the pasted copy) |
   | 11-phrase-move | the fixture | A | in "The first paragraph of the move: it stays where it is." select "it stays where it is", cut, click just after "The ", paste (note whether Word marks a move, green, or a deletion and an insertion) | + accept/reject (click in the pasted phrase), if Word marked a move |
   | 12-partner-missing | `check20/input/partner-missing.docx` | - | open it; note how Word shows the destination copy (a move? an insertion?); save it unchanged | + accept/reject |
   | 13-moveTo-no-ranges | `check20/input/moveTo-no-ranges.docx` | - | the same | + accept/reject |
   | 14-moveFrom-no-ranges | `check20/input/moveFrom-no-ranges.docx` | - | the same; for accept/reject click in the struck-through source copy instead | + accept/reject |

   Case 4's inputs (12 to 14) are the fixture with `word/document.xml` edited as text by
   `scripts/make-check20-inputs.mjs`, every other byte Word's: the source paragraph and its range
   removed (12), the destination's two range markers removed (13), the source's two removed (14).
   The saved files are compared with the fixture's section 3, as check 18's were.
   **When a dialog appears** (rejecting case 01 showed Word's "Tracked Moves Conflict Dialog",
   asking whether to keep the original location's text or the edited new location's), take the
   default and save under the planned name, then redo that step choosing the other option and save
   with `-keep-new` (or a suffix naming the choice); note which cases and actions showed one.
   **Run 2026-09-28 (Word 16.0.20326.20158), 39 files; the findings are CR-002 section 29's table.**
   The dialog came up on rejecting 01, 04, 06, 07 and 08 (their `-keep-new` files), not on 03.
   Jason's notes: 11 showed as an insertion and a deletion, 12 as moved (insertion) (so the re-run's reading; the first run's note said an insertion), 13 as added then
   moved, 14 as moved (insertion) then deletion, for the two locations respectively. Word does not
   save a file it has not changed: 12 and 13 as "saved unchanged" are byte-identical to their inputs,
   and in 14 Jason typed a space at the end of section 2 to make Word save it, which is where its
   conversion of the broken pair comes from.
   **Re-run of 12 to 14 (asked 2026-09-28)**, for two things the first run could not show. First, how
   Word *writes* each broken pair: open each input, change nothing, and File > **Save As** into
   `check20/` as `NN-name-resaved.docx` (Save As, unlike Ctrl+S on an unchanged file, should write it;
   if the file comes out byte-identical to its input, redo it by typing "x" at the end of section 2
   and pressing Backspace - with tracking on, deleting your own insertion leaves no markup - then
   save). Second, accept and reject on 14's broken markup itself: the first run took them from the
   saved file, where Word had already written the pair as a plain `w:ins` and `w:del`, so they measured
   a deleted paragraph, not the broken move. From `input/moveFrom-no-ranges.docx` each time: click in
   the struck-through source copy, Accept This Change, Save As `14-moveFrom-no-ranges-accept-input.docx`;
   the same with Reject Change, `-reject-input`; then both again clicking in the destination copy,
   `-accept-dest-input` and `-reject-dest-input`. Seven files; note anything Word says or selects.
   **Re-run 2026-09-28: Save As wrote all three**, each broken pair dissolved: every half Word could
   not pair written as a plain `w:ins` or `w:del`, run and mark, its range markers kept. 14's accept
   and reject from the input match the first run's from the saved file (ids apart), so Word dissolves
   the pair when it reads it; clicking the destination instead kept it, plain, either way (as 12),
   the source left a pending `w:del`. Jason also re-did 12's and 13's accept and reject; they match
   the first run, so only the seven files above were added. Word's labels, read again, for the
   destination and source respectively: 12 "moved (insertion)"; 13 "added", "moved"; 14 "moved
   (insertion)", "deletion". CR-002 section 29, fix E.
21. Typing with tracking **off** next to tracked changes (CR-002 phase F's `insertText`, asked by
   the editor for its tracker, E4.c step C3): does Word keep typed text inside the insertion it lands
   in, or make it plain? Section 1 of `fixtures/revisions/revisions-word15.docx` reads "The quick
   brown ~~red~~ fox jumps over the ~~very~~ lazy dog near the ~~old~~ stone bridge ~~at dawn~~.",
   where "red " is Author B's insertion, "very " Author A's, and "old " Author A's deletion. Each
   case on a **fresh copy**, with Review > Display for Review set to **All Markup** (so deletions
   show), Review > **Track Changes off**, and the author (the "As" column) set as the user name in File > Options > General, with "Always use these values regardless of sign in to Office" ticked; saved into
   `fixtures/revisions/check21/`:

   | Case | As | Do |
   |---|---|---|
   | `21a-in-insertion.docx` | B | click inside Author A's inserted "very", between "ve" and "ry", and type "X" |
   | `21b-at-insertion-end.docx` | B | click just after Author A's inserted "very " (before "lazy"), and type "Y" |
   | `21c-in-deletion.docx` | B | click inside Author A's deleted "old", between "o" and "ld", and type "Z" |
   | `21d-in-own-insertion.docx` | A | as 21a, but as Author A, whose insertion it is |

   The engine keeps text typed inside a `w:ins` in that `w:ins` when tracking is off (the editor's
   `api.test` holds it to that); if Word makes it plain, both change. **Run 2026-09-28: tracking was
   still on** - all four files have `w:trackRevisions` in `word/settings.xml` - so they show typing
   with tracking on (CR-002 section 29 has what they show), and the tracking-off question needs a
   re-run: turn Track Changes off (the Review tab button not highlighted) before typing, and check
   that the saved `settings.xml` has no `w:trackRevisions`. **Re-run 2026-09-28, tracking off
   (checked): Word makes the typed text plain** in all four, splitting the `w:ins` (21a, 21d) or
   `w:del` (21c) it lands in, or following the `w:ins` it ends (21b). The files are the re-run's;
   CR-002 section 29 has the first run's tracking-on results.
22. A deleted paragraph, listed, accepted and rejected (CR-002 section 29, what check 20's case 14
   left open): Word accepted and rejected a deleted paragraph's run deletion and its deleted mark
   together, where the engine lists and resolves them as two changes. No move is involved, so Office
   JS can list the changes (it throws on a move, checks 16 and 17). In a **new blank document**, with
   **Author A** as the user name in File > Options > General, with "Always use these values regardless of sign in to Office" ticked, Home > Show/Hide (the pilcrows showing) and Review >
   Display for Review > **All Markup**, type seven paragraphs with tracking **off**: "One.", "Two.",
   "Three.", "Four.", "Five six.", "Seven eight.", "Nine.". Then turn Track Changes **on** and:
   triple-click "Two." (its mark selected with it) and press Delete; click at the end of "Four."
   (before its pilcrow) and press **Delete** once, deleting that mark alone (note whether Word marks
   it or refuses); select from "six." through "Seven " (across the mark between them) and press
   Delete. Save into `fixtures/revisions/check22/` as `22-deleted.docx`. Then, from that file each
   time, as check 20 did: click in "Two.", Review > Accept > Accept This Change, save
   `22-accept-whole.docx`; the same with Reject Change, `22-reject-whole.docx`; then clicking in the
   deleted "six.", `22-accept-across.docx` and `22-reject-across.docx`. Note what Word selects on
   each. Last, open `22-deleted.docx`, run the Script Lab snippet for 22 below, and copy the JSON back
   (it lists the changes as Office JS does, whole and per paragraph, then accepts and rejects each
   of the first three from a fresh copy of the body, and puts the document back). The questions:
   is a deleted paragraph one change or two, is a lone deleted mark a change of its own, and what
   does accepting or rejecting one of them take with it. **Run 2026-09-28 (Word 16.0.20326.20158):
   one change for deleted text and marks that touch.** Office JS listed three: `Deleted "Two.\r"`
   (the paragraph's text and its mark), `Deleted "\r"` (the lone mark, which Word did mark), and
   `Deleted "six.\rSeven "` (across the mark), a mark's text being `\r`. Accepting or rejecting one
   resolved all of it, through Office JS and on the Review tab alike (the four saved files agree with
   the JSON, run boundaries apart). `Paragraph.getTrackedChanges()` listed the changes starting in the
   paragraph, cut at its end: "Five six." gave `six.\r`, "Seven eight." nothing. CR-002 section 29.
23. What makes tracked changes one change (CR-002 section 29, after check 22): do an inserted
   paragraph's text and mark list as one change, as a deleted one's do, and does a different date or
   author split deletions that touch? In a **new blank document**, run the Script Lab snippet for 23
   below (**Run the check**) and copy the JSON back; nothing is typed by hand. For each of nine cases
   it empties the body, inserts the case's paragraphs through `insertOoxml` with tracking off (the
   revisions written out, as check 17's were), lists the changes, then accepts the first one listed
   and records the paragraphs and the list after; then does the same again rejecting it. The cases:
   an inserted paragraph (`ins-paragraph`); an insertion across a mark (`ins-across`); a lone
   inserted mark (`ins-mark-alone`); a deleted paragraph whose mark is an hour younger than its text
   (`del-paragraph-dates`) or by another author (`del-paragraph-authors`); two touching `w:del` of
   the same author and date (`del-same`), an hour apart (`del-dates`), by two authors
   (`del-authors`); and a deletion touching an insertion (`del-then-ins`). **Run 2026-09-28 (Word
   16.0.20326.20158): one change is touching pieces of one kind by one author, whatever their
   dates.** Insertions group as deletions do (`Two.\r`, `five.\rSix `, `\r`); an hour's difference
   does not split, another author does, and a deletion and an insertion are two. `del-same`'s two
   elements Word merged into one when inserting them. Office JS reported every `w:date` as local
   time (01:00Z came back as 15:00Z the day before). CR-002 section 29, fix F, and its open item on
   dates.
24. Formatting changed and changed back, tracked (CR-002 section 29, asked by the editor for its C5
   agreement suite): does Word keep a `w:rPrChange` whose recorded properties equal the run's
   current ones, or drop it? The engine keeps it (a `w:rPrChange` with `<w:rPr/>` after bold then
   unbold); the editor drops it. It matters to `Range.hyperlink` with tracking on, whose restyle is a
   formatting change: a link made and then removed leaves the run as it was. In a **new blank
   document**, with **Author A** as the user name in File > Options > General, with "Always use these values regardless of sign in to Office" ticked, type "alpha beta gamma delta" with tracking off,
   then turn Track Changes **on** and save each step into `fixtures/revisions/check24/`:

   | File | As | Do |
   |---|---|---|
   | `24a-bold.docx` | A | select "beta", Ctrl+B |
   | `24b-unbold.docx` | A | carrying on from 24a, select "beta", Ctrl+B again |
   | `24c-unbold-by-b.docx` | B | open `24a-bold.docx`, change the user name to **Author B**, select "beta", Ctrl+B |
   | `24d-style.docx` | A | open `24b-unbold.docx`, select "gamma", apply the **Strong** character style (Home > Styles) |
   | `24e-style-cleared.docx` | A | carrying on from 24d, select "gamma", Ctrl+Space |

   Note what Word's Review pane lists after each step. The question in each of 24b, 24c and 24e is
   whether the run keeps a `w:rPrChange`. **Run 2026-09-28 (Word 16.0.20326.20158): Word drops it,
   all three times.** Bold then unbold by Author A, unbold by Author B of A's bold, and the Strong
   style then Ctrl+Space each left the run with no `w:rPrChange` and no `w:rPr` - no revision by
   anyone. CR-002 section 29.
25. A rejected list change's level (CR-002 section 29, asked by the editor after its C5): check 18's
   section 9 recorded `w:numPr` with no `w:ilvl` in `w:pPrChange`, and Word's reject wrote
   `w:ilvl="0"` - but the current level was 0 as well, so Word may keep the current level or write
   the recorded one (a missing `w:ilvl` meaning 0), and the engine and the editor each write one of
   the two. In a **new blank document**, with **Author A** as the user name in File > Options > General, with "Always use these values regardless of sign in to Office" ticked, and tracking **off**:
   type "One.", "Two." and "Three." as a numbered list (Home > Numbering), put the caret at the very
   start of "Two." and press Tab twice (it goes to the third level), then save
   `fixtures/revisions/check25/25-input.docx`. Then, each from a **fresh copy** of that file, with
   tracking **on**, as Author A:

   | File | Do |
   |---|---|
   | `25a-bullets.docx` | select all three items, click Bullets; save |
   | `25b-bullets-reject.docx` | carrying on from 25a, Review > Reject > **Reject All Changes**; save |
   | `25c-demote.docx` | put the caret at the very start of "One.", click Bullets, then press Tab twice (it goes to the third level, bulleted); save |
   | `25d-demote-reject.docx` | carrying on from 25c, Review > Reject > **Reject All Changes**; save |

   The questions: what `w:pPrChange` records for "Two." in 25a (its level, 2, or none) and what 25b
   restores; and in 25c, whether "One."'s recorded numbering still has no `w:ilvl` while its current
   one has `w:ilvl="2"` - the case that separates the two readings - and whether 25d writes
   `w:ilvl` 0, 2, or none. **Run 2026-09-28 (Word 16.0.20326.20158; `25-input.docx` was not kept).**
   25a recorded the level-2 item's numbering with its `w:ilvl="2"` and its level's own indent, and
   the level-0 items' with no `w:ilvl` and `w:ind w:hanging="360"` alone; 25b restored all three with
   no indent and explicit `w:ilvl` (2, and 0 where the recorded had none - the current level being 0
   again, so the question stays open). **25c did not do what was meant: no change was tracked.**
   Bullets clicked with only a caret in "One." made level 0 of the list's own definition a bullet
   (so "Three." turned too), and Tab at the start of a list's first item indented every level of the
   definition by 1080; the paragraphs kept their numbering. Word does not track a list definition,
   so the Review pane showed nothing and Reject All left 25d as 25c (recorded as a known issue in
   Word, `docs/word-known-issues.md` entry 1). The case is therefore asked
   again, with the paragraph selected and its level set by name, both of which change the paragraph.
   As **Author A** throughout:
   1. If `25-input.docx` is gone, make it again: new blank document, tracking **off**, "One.", "Two."
      and "Three." as a numbered list, the caret at the very start of "Two.", Tab twice; File > Save
      As `25-input.docx`.
   2. Open `25-input.docx`, File > **Save As** `25e-bullet-demote.docx` (before changing anything).
   3. Turn Track Changes **on**. Triple-click "Three." to select it, click Bullets, then, with it
      still selected, the arrow beside Bullets > **Change List Level** > the second level. Check that
      the Review pane shows a formatting change on "Three.". Ctrl+S.
   4. File > **Save As** `25f-bullet-demote-reject.docx` (before changing anything), then Review >
      Reject > **Reject All Changes**, Ctrl+S.

   25e should record "Three."'s numbering with no `w:ilvl` and carry `w:ilvl="1"`; 25f then writes 1
   (Word keeps the current level), or 0 or none (it writes the recorded one). **Run 2026-09-28: Word
   writes the recorded level.** 25e recorded `numId 1` alone (with `w:ind w:hanging="360"`) against a
   current `ilvl 1, numId 2`; 25f restored `ilvl 0, numId 1` and no indent - the 0 written out, as in
   every reject so far. CR-002 section 29.
26. Paragraph formatting changed and changed back, tracked (CR-002 section 29, check 24's question
   for a paragraph, asked by the editor): does Word keep a `w:pPrChange` once the paragraph's
   properties are back to what it recorded? Check 24 found Word drops a run's `w:rPrChange` then;
   the editor's tracker drops a paragraph's too, and the engine keeps it. Save into
   `fixtures/revisions/check26/`:
   1. As **Author A** (the user name in File > Options > General, with "Always use these values
      regardless of sign in to Office" ticked): new blank document, Track Changes **off**, type "The
      first paragraph." and "The second paragraph."; File > Save As `26-input.docx`.
   2. Open `26-input.docx`, File > **Save As** `26a-centred.docx`. Turn Track Changes **on**. Click in
      "The first paragraph.", press Ctrl+E (centre). Ctrl+S.
   3. Open `26a-centred.docx`, File > **Save As** `26b-left-again.docx`. Click in "The first
      paragraph.", press Ctrl+L (align left). Ctrl+S.
   4. Change the user name to **Author B**. Open `26a-centred.docx`, File > **Save As**
      `26c-left-by-b.docx`. Click in "The first paragraph.", press Ctrl+L. Ctrl+S. Change the user
      name back to **Author A**.
   5. Open `26-input.docx`, File > **Save As** `26d-indented.docx`. Turn Track Changes **on**. Click
      in "The second paragraph.", Home > **Increase Indent**. Ctrl+S.
   6. Open `26d-indented.docx`, File > **Save As** `26e-indent-back.docx`. Click in "The second
      paragraph.", Home > **Decrease Indent**. Ctrl+S.

   Note what Word's Review pane lists after each step. The questions: 26a and 26d should each carry a
   `w:pPrChange`; do 26b, 26c and 26e keep it? **Run 2026-09-28 (Word 16.0.20326.20158; 26e saved as
   `26d-indented-back.docx`): Word drops it, all three times,** and writes no `w:pPr`: Ctrl+L removed
   the `w:jc`, Decrease Indent the `w:ind`, and B's undoing of A's change left no revision by either.
   CR-002 section 29.
27. Tracked changes the engine does not list (CR-002 section 35), through Office JS: which of them
   current Word writes, how Office JS lists them, and what accepting and rejecting each leaves. In a
   **new blank document**, run the Script Lab snippet for 27 below (**Run the check**) and copy the
   JSON back; nothing is typed by hand and nothing needs saving. The snippet switches tracking on and
   off itself, so the user name does not matter. Two parts:
   - **Word writes** (11 cases): a paragraph, a 3x3 table and a paragraph are put in untracked, then,
     with tracking on, Office JS changes one thing - the table's alignment, width or style, a cell's
     shading, width or vertical alignment, the first row's height, a column added or deleted, two
     cells merged vertically, the first paragraph made bold (its mark too?), a paragraph made a list
     item and moved a level down. What revision markup Word writes, if any, is the answer.
   - **Written by hand** (10 cases): each kind of section 35's markup, by Author A, put in through
     `insertOoxml` with tracking off - `w:tblPrChange`, `w:tblGridChange`, `w:trPrChange`,
     `w:tcPrChange`, `w:cellIns`, `w:cellDel`, `w:cellMerge`, `w:numberingChange`, a mark's
     `w:rPrChange`, and a section break's `w:sectPrChange`.

   For every case, from a fresh start each time: the markup, what `getTrackedChanges()` lists (type,
   author, text), then the first change accepted, the first rejected, all accepted and all rejected,
   each with the markup after (a block unchanged from before shows as `=`) and what is still listed.
   A case that throws records the error and the next goes on. The document is left holding the last
   case; close it without saving. It needs WordApi 1.6; merging cells needs `Table.mergeCells`, and
   the case records an error where the build lacks it. **Run 2026-09-28 (Word 16.0.20326.20158);
   CR-002 section 35 has the findings, `docs/word-known-issues.md` entries 2 to 4 the Word issues.**
   Two faults of the snippet's own: `body.clear()` keeps the last paragraph's properties, so the
   list made in `list-level` (whose accept and reject threw) stayed on the body's last paragraph
   through part B and was listed there (in `tblPrChange` it was the only change listed); and its
   `revisions` summary came out empty although the markup holds the revisions.
28. Check 27's three open questions (CR-002 section 35), by hand in one document:
   `fixtures/revisions/check28/28-input.docx` (made by the engine; a copy is in the shared `__tmp/28/`).
   Question 1: when a table changes, does Word write a `w:trPrChange` for rows that already have
   properties of their own, and does rejecting keep those properties? Question 2: is a row's height
   removed when the change is rejected from a **saved file**, not in the session? Question 3: are two
   touching runs' formatting changes listed as one change or two? As **Author A** throughout (the
   user name in File > Options > General, with "Always use these values regardless of sign in to
   Office" ticked):
   1. Open `28-input.docx`. File > **Save As** `28a-changes.docx`. Turn Track Changes **on**.
   2. Table 1 (rows 1a to 3a): click in it, then the table's **Layout** tab > **Properties** > **Table**
      tab > Alignment **Center** > OK.
   3. Table 2 (x1 to y2): click in its first row, then the table's **Layout** tab > **Height** (Cell
      Size group): type 1.5 cm, Enter.
   4. In the last paragraph, select the words "alpha beta" (the italic "alpha" and the plain "beta")
      and press Ctrl+B.
   5. Ctrl+S. Then, with `28a-changes.docx` still open, run the Script Lab snippet for 28 below
      (**Run the check**) and copy the JSON back: it only lists the changes.
   6. Close the document. Open `28a-changes.docx` again, File > **Save As** `28b-reject-all.docx`,
      then Review > Reject > **Reject All Changes**, Ctrl+S.
   7. Close it. Open `28a-changes.docx` again, File > **Save As** `28c-accept-all.docx`, then Review >
      Accept > **Accept All Changes**, Ctrl+S.

   Save the four files into `fixtures/revisions/check28/` (or leave them in `__tmp/28/`). **Run
   2026-09-28 (Word 16.0.20326.20158; the files saved as `28a-changes`, `28a-reject-all`,
   `28a-accept-all`).** Question 1: Word wrote a `w:trPrChange` for each row that had properties of its
   own (row 1's exact height and header row, row 2's no-break), none for row 3, which had none, and
   rejecting from the saved file gave rows 1 and 2 back their recorded properties and row 3 none.
   Question 2: yes, the height went (table 2). Question 3: **one** change, `Formatted "alpha beta"`,
   the two runs having recorded different formatting (italic, none). CR-002 sections 29 and 35.
29. Enter at the **end** of a paragraph: which paragraph mark does Word treat as the new one (CR-002
   section 29, asked by the editor for its ED-005 section 12.16)? In the middle of a paragraph Word
   puts the new mark at the caret, so the first half's mark is the inserted one (the revisions
   fixture's section 6). At the end, if it does the same, the paragraph that keeps the old mark is
   the new empty one - and after a Heading 1, whose next style is Normal, that paragraph changes
   style: does Word then record a `w:pPrChange`, or mark the other pilcrow? Input:
   `fixtures/revisions/check29/29-input.docx` (made by the engine; a copy is in the shared
   `__tmp/29/`): a Heading 1 ("Case a heading") and Normal paragraphs, two of whose marks Author A
   inserted ("Cases c and d: ..." and "Case e: ..."). Before starting: Home > **Show/Hide** (the
   pilcrows showing) and Review > Display for Review > **All Markup**. The author is the user name in
   File > Options > General, with "Always use these values regardless of sign in to Office" ticked.
   Each case has its own copy of `29-input.docx`, already named, in `__tmp/29/`: open the case's
   file, set the author and tracking as the table says, act, Ctrl+S, close.

   | File | Author | Track Changes | Do |
   |---|---|---|---|
   | `29a-heading-end-tracked.docx` | **Author B** | **on** | click at the very end of "Case a heading" (press End), press Enter, type New |
   | `29b-normal-end-tracked.docx` | **Author B** | **on** | click at the very end of "Case b: the end of this Normal paragraph." (End), press Enter, type New |
   | `29c-inserted-mark-end-untracked.docx` | **Author B** | **off** | click at the very end of "Cases c and d: Author A inserted this paragraph's mark." (End: the caret just before Author A's inserted pilcrow), press Enter, type New |
   | `29d-inserted-mark-end-tracked.docx` | **Author B** | **on** | the same as 29c |
   | `29e-inserted-mark-middle-untracked.docx` | **Author B** | **off** | in "Case e: before the split, after the split. ...", click just after the first comma (after "split,"), press Enter |

   For 29a and 29d, before closing, note what Review > **Reviewing Pane** lists (for example whether
   a "Formatted" entry appears, and which paragraph each inserted pilcrow is on). If Word 2010 is to
   hand, the same five again, each file name ending `-word2010`. The questions, per file, for the
   paragraphs around the caret: their text and `w:pStyle`; which carries `w:pPr/w:rPr/w:ins` (or
   `w:del`), and whose; and whether a `w:pPrChange` was written, and where. **Run 2026-09-29 (Word
   16.0.20326.20158; Word 2010 not run; the Reviewing Pane not noted): the caret rule at the end
   too.** The new mark is the one at the caret, so the paragraph before it ends with the new mark and
   the paragraph after it - "New" - keeps the old one. 29a: "Case a heading" stays Heading 1, its mark
   inserted by B; "New" keeps the old mark and turns Normal, with a `w:pPrChange` recording Heading 1,
   its text B's `w:ins`. 29b: the first mark inserted by B, "New" with no `w:pPrChange`. 29c (tracking
   off): the first paragraph a new plain mark, "New" keeping Author A's inserted mark, its text plain.
   29d: the first mark inserted by B, "New" keeping A's inserted mark, its text B's `w:ins`. 29e: the
   first half plain, the second keeping A's inserted mark. CR-002 section 29.

30. What Word writes for content inserted under tracking that holds a hyperlink, a field, a content
   control, a smart tag, custom XML, a block control or a nested table, and what rejecting it leaves
   (CR-002 section 37, items 4 and 6). In a **new blank document**, run the Script Lab snippet for 30
   below (**Run the check**) and copy the JSON back; nothing is typed by hand and nothing needs saving.
   The snippet switches tracking on and off itself, so the user name does not matter. The Script and
   HTML tabs are also in the shared `__tmp/30/` (`check30-script.js`, `check30.html`). Every case
   starts from "Kept." and the body's last, empty paragraph, put in untracked. Three parts:
   - **Word inserts** (8 cases): with tracking on, `body.insertOoxml` at the end of one fragment - a
     paragraph with a hyperlink, with a `w:fldSimple` PAGE field, with an inline content control, with
     a smart tag, with custom XML; a paragraph then a block content control; a paragraph then a table;
     a table whose cell holds a nested table. Each ends with a plain "(end of insertion)" paragraph.
   - **Gestures** (5 cases), tracking on: text typed into a new empty content control; inserted text
     then wrapped in a content control; a content control put in **untracked**, then typed into with
     tracking on (its markup before the typing is `setup`); a PAGE field inserted (`Range.insertField`,
     WordApi 1.5); a hyperlink set on inserted text.
   - **Engine markup** (8 cases): the same fragments as `@docx4j/core-ts` f96fa6c writes them when they
     are inserted after "Kept." under tracking by Author A - the `w:ins` inside each run-level holder -
     put in through `insertOoxml` with tracking **off** (known issue 4: check that `markup` still
     holds the revisions).

   For every case, from a fresh start each time: the markup (`markup`, the body's blocks with ids and
   rsids taken out), the revision elements in it, and what `getTrackedChanges()` lists; then reject all
   and accept all, each with the markup, the body's text and what is still listed after. A case that
   throws records the error and the next goes on. Close the document without saving.

   The questions: where Word puts the `w:ins` for an inline content control and a smart tag - inside
   the holder (as the engine does, and as the schema forces for a hyperlink and a field) or around
   it; whether Word keeps a smart tag and custom XML at all; and, after reject all, whether an emptied
   hyperlink, field, inline or block content control, smart tag or custom XML element is left, whether
   a paragraph whose inserted mark had a table or a block control after it is left, and whether an
   **existing** control that only an insertion was typed into stays (it should: the engine cannot tell
   it from an inserted one by its markup). The engine today (f96fa6c), rejecting each fragment
   inserted after "Kept.": each run-level holder left empty at the end of "Kept."'s paragraph (the
   inserted paragraph's mark rejected, its remains join the paragraph before); a block control left
   holding an empty paragraph, and an empty paragraph after it; a table, nested or not, removed with
   its last row, and one empty paragraph left after "Kept.".
   **Run 2026-09-30 (Word 16.0.20326.20158), from a new snippet; the JSON is `fixtures/revisions/check30/result.json`.**
   (Two earlier runs the same day from an older snippet came back with every `markup` empty and a
   paragraph mark per XML tag of the package; a diagnostic, `diag30-script.js` in the shared folder,
   showed `body.insertOoxml` at `End` working, tracked or not, so the snippet now opens with a `probe`
   and records `markupRead` when a case's markup is empty. Two Word facts from the diagnostic: a
   `w:hyperlink` with `w:anchor` comes back as a HYPERLINK **field**, and `Paragraph.insertOoxml`
   rejects `After`.)

   **Word inserts** (tracking on): the paragraph's mark inserted in `w:pPr/w:rPr`, its runs in **one**
   `w:ins`, and the holders:
   - *hyperlink* and *field*: the link (as a HYPERLINK field: `w:fldChar` runs) and the `w:fldSimple`
     (as a complex field) go into that one `w:ins` with the runs; accepted, the link is a `w:hyperlink`
     again, the field stays complex. Neither element has a form of its own in a `w:ins`.
   - *control-inline* and *control-block*: the `w:ins` **inside** `w:sdtContent` (the engine's placement),
     and the control itself marked inserted with `w:customXmlInsRangeStart` / `w:customXmlInsRangeEnd`:
     one start **before** the `w:sdt` whose end is the **first** item of `w:sdtContent`, and a second start
     the **last** item of `w:sdtContent` whose end follows the `w:sdt`; the two pairs' ids run in
     document order with the `w:ins` ids (2, 3, 4).
   - *smart-tag*: the `w:smartTag` goes **into** the one `w:ins` beside the runs, its own runs plain.
   - *custom-xml*: `insertOoxml` dropped the `w:customXml` element; its runs came in as plain text.
   - *table* and *nested-table*: `w:trPr/w:ins` on every row, nested ones too, every cell paragraph's
     mark and runs inserted. Word also added `w:tblInd`, `w:tblCellMar`, `w:tblLook` and a `w:tblPrEx`
     per row.
   - The "(end of insertion)" paragraph took the body's last (untracked) mark, and Word added one
     empty paragraph after it, so every reject-all left `"Kept."` and two empty paragraphs, nothing
     else: **the emptied holders are all removed** - the link, field, both controls, smart tag, and a
     table with the paragraph before it (whose mark was inserted). Accept all left everything, no
     marker, no revision. `getTrackedChanges` listed each case as **one** change, the block control and
     the table included (`"before the control\rin a block control\r(end of insertion)"`).

   **Gestures**: `new-control-typed` and `inserted-text-wrapped` wrote the same markers around the new
   control, the `w:ins` inside it; reject all removed control and text, accept all kept the control
   with no marker. `existing-control-typed` (the control put in untracked, then typed into): **no
   markers**, the `w:ins` inside; reject all **kept the control**, which shows its placeholder - and
   Word left the placeholder run inside a `w:ins` of Author B's, so one `Added` change with empty text
   was still listed (`docs/word-known-issues.md` entry 6). `field-inserted`: `w:ins` around the
   field's begin and instruction runs and around its end run, the separator and result runs plain.
   `link-on-inserted-text`: one `w:ins` around the HYPERLINK field runs; accepted, a `w:hyperlink` with
   an `r:id`. Each gesture listed as one change.

   **Engine markup** (f96fa6c, inserted untracked): every revision survived `insertOoxml`, but the last
   paragraph's inserted mark, which took the target paragraph's plain one (the same in every case).
   Word listed the hyperlink, field and inline control cases as **three** changes each (`"before "`,
   `"link"`, `" after\r(end of insertion)"`: a holder that is not itself an insertion keeps the `w:ins`
   inside it apart from those outside; its own forms, above, group as one - this is the editor's E4.g
   question, and the engine's listing now splits the same way), the block control as three, the smart
   tag and custom XML as one. Rejecting all left, in an empty paragraph after
   "Kept.": an **empty `<w:hyperlink w:anchor="target"/>`**; an **empty PAGE field** (begin, instruction,
   separate, end); the inline control, **kept**, showing its placeholder (`w:showingPlcHdr` and a run of
   five spaces); the block control kept the same way, and an empty paragraph; nothing for the table
   cases. The smart tag: Word had **rewritten** the engine's three `w:ins` as one around everything, the
   `w:smartTag` inside it, and reject all removed it; the custom XML element was dropped on the way in.
   So Word's answer to the engine's markup was the engine's own (an emptied holder stays), and the
   difference is in what Word *writes*: the markers, and the smart tag inside the `w:ins`. **The engine
   now writes both** (CR-002 section 37 item 7): a control or custom XML element it inserts under
   tracking gets the marker pairs, is listed as a piece of the change and goes when the change is
   rejected, leaving what was in it; a smart tag goes into the `w:ins`. A hyperlink or a field is left
   as it was, empty, since Word leaves the engine's form the same way and has no form of its own for
   them. What Word does with an *inserted* control whose markup carries no markers - stays, as an
   existing one - is what the engine did before, and what it still does for a control that was there.

31. How Office JS lists an inserted table of several rows, with and without insertions touching it,
   and rows added to a table that was there (the open question of CR-002 section 37 item 7: check 30's
   one-row tables listed as one change with the paragraphs before and after them, where the engine and
   the editor list a table and each row apart). In a **new blank document**, run the Script Lab snippet
   for 31 below (**Run the check**) and copy the JSON back; nothing is typed by hand and nothing needs
   saving. The tabs are in the shared `__tmp/31/` (`check31-script.js`, `check31.html`). Three parts,
   every case from "Kept." and the body's last, empty paragraph, put in untracked:
   - **Word inserts** (3 cases), tracking on, `body.insertOoxml` at the end: a paragraph, a three-row
     table and a paragraph (`table-between`, check 30's table case with three rows); the three-row table
     alone (`table-alone`); a two-row table, a paragraph and another two-row table (`two-tables`).
   - **Gestures** (5 cases) on a three-row table put in untracked, tracking on: `table.addRows` of two
     rows at the end, and at the start; one row after the first and one after the third (`rows-apart`);
     a paragraph inserted before the table then a row at its start (`text-then-rows`); a row at the end
     then a paragraph after the table (`rows-then-text`).
   - **Engine markup** (4 cases, `@docx4j/core-ts` 426e7b5's forms), tracking off: `table-between` as
     the engine writes it (every row `w:trPr/w:ins`, every mark and run inserted, one author); the same
     with the middle row Author B's; two inserted rows after a plain one, plain paragraphs around; an
     inserted first and last row around two plain ones.

   For every case: the markup, what `getTrackedChanges()` lists (type, author, text), then reject all
   and accept all with the markup, text and what is still listed. The questions: whether touching
   inserted rows are one change or one per row; whether an inserted table groups with inserted text
   before and after it (as check 30's one-row tables did); whether another author's row, or plain
   rows between, keep them apart. The engine today lists `table-between` as five changes (the
   paragraph, each row, the paragraph), `two-authors` as five, `rows-added` as two, `rows-apart` as
   two; the editor lists rows apart too, and both follow the answer.

   **Run 2026-10-01 (Word 16.0.20326.20158); the JSON is `fixtures/revisions/check31/result.json`.**
   Rows group with whatever inserted content touches them, and a table boundary is no break:
   - *table-between*: **one** change, `"before the table\rrow one\r\nrow two\r\nrow three\r\n(end of
     insertion)"`; *table-alone*: one, `"row one\r\nrow two\r\nrow three\r\n"`; *two-tables*: one across
     both tables and the paragraph between. A row's text is its cells and `\r\n`.
   - *addRows* of two rows at the end, and at the start: one change each; *rows-apart*: two;
     *rows-then-text*: **one**, `"new last\r\ntyped after the table\r"`; *text-then-rows*: **two** -
     because `insertParagraph('Before')` on the table wrote the paragraph's runs in a `w:ins` but left
     its **mark plain** (`docs/word-known-issues.md` entry 8; reject all left an empty paragraph), and a
     plain mark keeps the text apart from the rows, as anywhere.
   - Engine markup: *engine-table-between* one change, as Word's own; *engine-rows-added* one;
     *engine-rows-apart* two; *engine-two-authors* **two**, Author A's halves - and Author B's row,
     inserted like the others, was **not listed at all** (known issue 8's neighbour, entry 7), though
     reject all removed it and accept all kept it.
   - Every reject all and accept all left what the markup says, nothing listed after.

   So the rule is check 22's, with rows as pieces: touching pieces of one kind and author are one
   change, and neither a table nor a row is a boundary. **The engine's listing follows (CR-002
   section 37 item 7): a row change groups with the rows and text touching it, a table's property
   records stay apart, and a row's text is Office JS's.** The engine lists Author B's row where Office
   JS did not.

32. Two rules neither side has measured (the editor's request of 2026-10-02, at Jason's word; its
   ED-005 section 12.47 item 10 and 12.32 item 3): what rejecting a pending insertion that holds a
   comment's word does to the comment (CR-002 section 38 item 1's rule: the markers stay, the
   comment survives collapsed where its text was), and what Word writes for typing into a link that
   was there. The tabs are in the shared `__tmp/32/` (`check32-script.js`, `check32.html`, three
   buttons). In a **new blank document**:
   - **Case A** (automated: **Run A and B-api**, copy the JSON back): three paragraphs, each put in
     untracked through `insertOoxml` with its comment part ("On monthly." / "On beta."), from a fresh
     start each time, then that one change rejected (A3: accepted), reject all, and accept all. A1 is
     Word's own form from the editor's `measure-word15.docx` (the range start inside the `w:ins`, the
     end marker and the reference after it); A2 the editor's (the end marker just inside the
     `w:ins`); A3 is A1 with a `w:del` for the `w:ins` and a plain run after. Each snapshot records the
     markup, the text, what is listed, `body.getComments()` with each comment's anchored text, and
     the comments part. The same run does **B through the API**: the link paragraph put in untracked,
     tracking on, then `insertText` of "x" after "le", "y" after "lexase" and "z" before it, a snapshot
     after each.
   - **Case B by hand**: click **Set up typing (B)**, which leaves "Read the lease today." with "lease"
     an external link and tracking on; type **x** between "le" and "ase", **y** right after "lexase"
     (before the space), **z** right before its "l"; click **Record typing (B)** and copy the JSON
     back; then File > Save As `check32-typed.docx` into `__tmp/32/` (the saved file shows what Word
     rewrites at save, as it rewrote `w:hyperlink` holders in check 30). Current Word is enough; Word
     2010 or 15 only if an answer would change a rule.

   The engine's answers for A (426e7b5): A1 lists "The tenant pays rent monthly" and " in advance.\r"
   apart (the plain reference run between them); rejecting the first leaves the paragraph's inserted
   mark, then `commentRangeStart`, `commentRangeEnd` and the reference run, then the second `w:ins`;
   reject all leaves the markers and the reference at the end of "Kept."; accept all is the text with
   the comment on "monthly". A2 the same shape. A3: accepting the deletion leaves the markers and the
   reference before the second `w:del`; accept all leaves them before "Kept after."; reject all is the
   text with the comment on "monthly". For B the engine writes typed text inside a `w:hyperlink` as a
   `w:ins` inside it, at the link's start and end as in its middle.

   **Run 2026-10-02 (Word 16.0.20430.20118); the JSON is `fixtures/revisions/check32/result.json`
   (A and B-api) and `typed.json` (B by hand), the save `check32-typed.docx`.**
   - **A: Word removes the comment with its text.** Rejecting the insertion that holds "monthly" (A1,
     Word's form), rejecting all, and accepting the deletion in the same shape (A3) each left no
     marker, no reference run, no comment in `body.getComments()` and **no comments part** in
     `getOoxml()`; accepting all (A1, A2) and rejecting all (A3) left the comment on its word. So
     CR-002 section 38 item 1's rule - the markers stay, the comment survives collapsed - was wrong
     for this form, and the engine now removes the comment too (its entries from the parts when the
     comments are next read or the package saved). The same day the editor showed Word 2010 and Word
     15 **keeping** a comment whose markers lay **outside** the rejected insertion (0.3.0's form, its
     `review-for-word.docx` Reject All saves): the reference run alone, no range markers, the entry
     kept. So the rule is by the start marker's place, and the engine keeps the outside form's comment
     that way. Two things `insertOoxml` did on the way in: it moved A2's end marker
     out of the `w:ins`, into A1's form, and it dropped A1's inserted paragraph mark (the target mark
     wins, as in checks 30 and 31). And Office JS listed each of A1 and A2 as **one** change across the
     plain reference run between the two `w:ins` - where the editor's Word 15 Reviewing Pane had
     split "lease" apart at a reference run (its ED-005 section 12.29 item 3) - so a run holding only
     a comment reference no longer keeps pieces apart in the engine's listing either.
   - **B: typed text goes into the link in its middle and outside it at either end; the API's goes
     inside at all three.** The link came in as a HYPERLINK field (`w:fldChar` runs), as check 30
     found. Typing "x" between "le" and "ase": `w:ins` inside, between the field's separate and end.
     Typing "y" at the link's end: `w:ins` **after** the field's end run; "z" at its start: `w:ins`
     **before** the field's begin run. `range.insertText('x', 'After')` on "le", `('y', 'After')` on
     "lexase" and `('z', 'Before')` on it: each a `w:ins` **inside** the field. Three changes listed,
     "z", "x", "y". The save kept the field form, `w16du:dateUtc` on each `w:ins`: with pending
     insertions inside, Word does not write the link back as a `w:hyperlink`. The engine writes a
     `w:ins` inside the `w:hyperlink` for all three, as the API does; `Range.insertText('Before')`
     used to go outside at the link's start, and joins the run the range starts in now.

33. Whether deleting a **whole** paragraph under tracking, text and mark, hands its properties to
   the next paragraph as deleting its mark alone does (CR-002 section 38 item 2, measured in Word 15
   for Delete at a heading's end; the editor's finding on 0.3.1, where the carry turned an agent's
   new heading into body text). Open `fixtures/revisions/check33/33-input.docx` (made by the engine;
   a copy is in the shared `__tmp/33/`): "Before the heading.", a Heading 2 "A heading deleted
   whole", "The body paragraph after it.", "After.". In current Word, **File > Save As**
   `33a-deleted.docx` first (AutoSave overwrites), user name Author A. Review > Track Changes **on**.
   Triple-click the heading, which selects the whole paragraph with its mark, and press **Delete**.
   Ctrl+S. Then Review > Accept > **Accept All Changes**, File > Save As `33b-accepted.docx`. Put
   both files in `fixtures/revisions/check33/` (or leave them in `__tmp/33/`). The questions: in
   `33a`, whether "The body paragraph after it." gained `w:pStyle Heading2` with a `w:pPrChange`
   recording its own (as the mark-alone case did), or is untouched; and in `33b`, whether it is a
   heading or body text. The engine (since 2026-10-02) carries nothing for a whole-paragraph delete,
   so `33b`'s paragraph would be body text; if Word carries, the carry comes back for this case.
   **Run 2026-10-02 (current Word), files in `fixtures/revisions/check33/`: Word carries nothing.**
   `33a`: the heading's text in a `w:del` and its mark `w:pPr/w:rPr/w:del`, its `w:pStyle Heading2`
   kept; "The body paragraph after it." untouched, no `w:pStyle`, no `w:pPrChange`. `33b`: the heading
   gone, the body paragraph Normal. So the carry of CR-002 section 38 item 2 belongs to a mark deleted
   alone, with text kept, and not to a paragraph deleted whole; the engine's withdrawal of it stands,
   measured.

34. A plain text content control created already mapped to a custom XML node, and a control mapped to
   a node whose value differs from its text (the editor's request of 2026-10-03, at Jason's word, for
   its Data tab's "bind"; CR-002 phase E). Open `fixtures/revisions/check34/34-input.docx` (made by the
   engine; a copy is in the shared `__tmp/34/`): a custom XML part `<data><name>Ann</name><empty/></data>`
   and a labelled paragraph per case, "1a." to "2b.", "The quick **brown** fox." with "brown" bold;
   1d's is the empty paragraph under its label, 1f's two paragraphs, 1g's holds a small picture, 1h's
   an existing rich text control titled "Existing" (no footnote: a selection holding one is not
   covered). In current Word, **File > Save As** `34a-mapped.docx` first (AutoSave overwrites);
   user name Author A; Track Changes **off**. Developer > **XML Mapping Pane**; in its Custom XML
   Part list choose the part whose name is `(no namespace)` or shows `data`. For each case put the
   caret or selection in that case's paragraph, then in the pane right-click **name** (under data)
   and choose **Insert Content Control > Plain Text**:
   - 1a: the caret inside "quick" ("qu|ick"). 1b: between "quick" and "brown". 1c: at the paragraph's
     end. 1d: in the empty paragraph.
   - 1e: "quick brown" selected (the selection's first run plain, its second bold).
   - 1f: a selection from "quick" in the first of the two paragraphs to "brown" in the second.
   - 1g: a selection holding the picture ("here: " through the picture).
   - 1h: the caret inside the existing control's text.
   - 1i: the caret at the end of its paragraph, but right-click **empty** instead of name.
   Note, as you go, what each control shows (its text, or a placeholder), and anything Word refused
   or said. **Ctrl+S.** Then 1j: Track Changes **on**, "quick brown" selected in the "1j." paragraph,
   name > Insert Content Control > Plain Text; **File > Save As** `34b-tracked.docx`. Put both files
   in `fixtures/revisions/check34/` (or leave them in `__tmp/34/`).

   **Part 2, the API.** Which side wins when a control that already has text is mapped: the node's
   value into the control, or the control's text into the node? First the Script Lab probe (new
   blank document; `__tmp/34/` `check34-script.js` and `check34.html`; **Run the check**, copy the
   JSON back): it adds the part with `customXmlParts.add`, makes the control with
   `insertContentControl('PlainText')` (falling back to rich text if the build refuses the type), and
   looks for `ContentControl.xmlMapping`; if Office JS has no mapping API, which is likely, it says
   so and the answer comes from VBA, which is the API the engine's `XmlMapping` mirrors. For that,
   open `34-input.docx` again, **Save As** `34c-vba.docx`, Alt+F11, Ctrl+G (the Immediate window),
   and run these three lines one at a time (each is one line; "2a." is paragraph 14, "2b." 15):

   ```
   Set r = ActiveDocument.Paragraphs(14).Range: r.Find.Execute FindText:="quick brown": Set cc = ActiveDocument.ContentControls.Add(wdContentControlText, r): cc.XMLMapping.SetMapping "/data/name": Debug.Print "2a control: "; cc.Range.Text; " | node: "; cc.XMLMapping.CustomXMLNode.Text
   Set r = ActiveDocument.Paragraphs(15).Range: r.Find.Execute FindText:="quick": r.Collapse 0: Set cc = ActiveDocument.ContentControls.Add(wdContentControlText, r): cc.XMLMapping.SetMapping "/data/name": Debug.Print "2b caret control: "; cc.Range.Text; " | node: "; cc.XMLMapping.CustomXMLNode.Text
   Set r = ActiveDocument.Paragraphs(15).Range: r.Find.Execute FindText:="brown": Set cc = ActiveDocument.ContentControls.Add(wdContentControlText, r): cc.Range.Text = "Ann": cc.XMLMapping.SetMapping "/data/name": Debug.Print "2b Ann control: "; cc.Range.Text; " | node: "; cc.XMLMapping.CustomXMLNode.Text
   ```

   Copy the Immediate window's three printed lines back, and **Ctrl+S** (the file then holds the
   controls and the part as Word left them).

   The engine today: `Range.insertContentControl('PlainText')` wraps the selected runs as they are,
   formatting kept, and an empty control at a caret; a `Range` is within one paragraph, so 1f cannot
   be asked of it; inside an existing control it nests a new one. `xmlMapping.setMapping('/data/name')`
   writes the `w:dataBinding` and moves **no** value either way: the control keeps "quick brown" and
   the node "Ann" until the caller runs `applyBindings` (node into controls) or
   `updateFromContentControls` (controls into nodes). Whichever side Word's SetMapping takes, the
   engine's `setMapping` will follow.

   **Run 2026-10-03 (Word 16.0.20430.20118); files in `fixtures/revisions/check34/`: `34a-mapped.docx`,
   `34b-tracked.docx`, `34c-vba.docx` (three plain text controls made by VBA, unmapped),
   `34d-mapped-existing.docx`, `probe.json` (the Office JS probe's second run; `probe-first-run.json`
   its first, which read an unloaded property before the mapping).**
   - **Part 1.** Every insert made a plain text control showing the node's **"Ann"** (1i the
     placeholder), with `w:dataBinding w:prefixMappings="" w:xpath="/data[1]/name[1]"` (the canonical
     path), the part's `storeItemID`, `<w:text/>`, and Word's default placeholder `w:docPart`. 1a, 1b,
     1c: a run-level control at the caret, the word split around it ("qu", control, "ick "), the
     control's run plain. 1d: a **block-level** control around the empty paragraph, holding "Ann".
     1e: the selection **replaced** by "Ann" in a plain run (the selection's first run was plain; a
     bold first run is not measured). 1f: **not refused**: the control replaced the first paragraph's
     selected text ("quick") and the second paragraph was left as it was. 1g: the selected text
     replaced and the picture **kept, moved to the end of the paragraph** after " after it." (why it
     could not be seen). 1h: a new mapped control **nested** inside the existing one, its text split
     around it. 1i: `w:showingPlcHdr` and the placeholder run "Click or tap here to enter text." in
     `PlaceholderText`. 1j, tracking on: the control marked inserted with `w:customXmlInsRangeStart` /
     `End` (check 30's form), and inside it the selected runs in a `w:del` and "Ann" in a `w:ins`.
   - **Part 2.** Office JS **has** `ContentControl.xmlMapping` (`setMapping`, `setMappingByNode`,
     `delete`, `customXmlNode`, `customXmlPart`, `isMapped`, `prefixMappings`, `xpath`: the engine's
     shape) and `insertContentControl('PlainText')` works (type `PlainText`, subtype
     `PlainTextInline`). `setMapping('/data/name')` returned **true** and put the node's value into the
     control - "quick brown" became "Ann", an empty control "Ann", "Ann" stayed - with the part
     unchanged: **the node wins**. The binding is written with the XPath as given, `/data/name`, and
     the `storeItemID` of the **first** part in which it selects a node (the probe added a part per
     case, and the later cases bound to the first case's part). By hand, the pane's **Map to Selected
     Content Control** on the VBA-made control holding "quick brown" (`34d`) did the same: "Ann", the
     canonical path, the part unchanged. VBA's `XMLMapping.SetMapping` raised error 4198 in the
     Immediate window and was not pursued; its `ContentControls.Add(wdContentControlText, range)`
     refused a range already inside a plain text control, which is what the later 4198s were.
   - **The engine follows:** `xmlMapping.setMapping` and `setMappingByNode` now put the node's value
     into the control (CR-002 section 39). Its `Range.insertContentControl` still keeps the selected
     runs as they are; an insert-and-map in one step is the editor's to compose from the two.

35. Rich text content controls mapped to custom XML (the editor's request of 2026-10-03, at Jason's
   word, for its Data tab's bind; CR-002 phase E). Does Word let a rich text control carry a binding,
   in which element (`w:dataBinding`, or `w15:dataBinding` as `invoice2013.docx`'s "terms" control
   has it), and what goes into the control and into the node? Open
   `fixtures/revisions/check35/35-input.docx` (made by the engine with
   `scripts/make-check35-input.mjs`; a copy is in the shared `__tmp/35/`). Its custom XML part is
   `<data><name>Ann</name><rich1/><rich2/><bound>Bea</bound></data>`, and it has a labelled paragraph
   per case:
   - "1a.": a run-level rich text control titled 1a around "quick **brown**";
   - "1b.": a block-level rich text control titled 1b holding two paragraphs, "bold" in bold;
   - "1c." and "1d.": an empty paragraph each, under the label;
   - "1f.": a run-level rich text control titled 1f showing "Old text", bound with `w:dataBinding` to
     `/data/bound` ("Bea"), as the engine's `setMapping` and the editor bind a rich text control.

   **Part 1, current Word.** **File > Save As** `35a-word15.docx` first (AutoSave overwrites); Track
   Changes **off**. First note what 1f shows: "Old text" or "Bea". Developer > **XML Mapping Pane**;
   in its Custom XML Part list choose the part that shows `data`. Then, in order:
   - 1a: the caret inside 1a's text; in the pane right-click **name** and choose **Map to Selected
     Content Control**. Note what the control shows, and anything Word refuses or says.
   - 1b: the caret inside 1b's first paragraph; right-click **rich1** > Map to Selected Content
     Control. Note what it shows: its two paragraphs, a placeholder, or something else. Then type
     " Edited." at the end of its second paragraph.
   - 1c: the caret in the empty paragraph under "1c."; right-click **rich2** > **Insert Content
     Control**, note which kinds the submenu offers, and choose **Rich Text**. In the new control type
     `Hello `, then **Ctrl+B**, `world`, **Ctrl+B**, **Enter**, `Second line.`
   - 1d: the caret in the empty paragraph under "1d."; right-click **rich2** > Insert Content Control
     > **Rich Text**. Note what it shows: 1c's two paragraphs with "world" bold, their text alone, or
     a placeholder.
   - 1e: the caret inside 1a's control, mapped since 1a; right-click **name** > Insert Content Control
     > **Plain Text**. Note whether Word nests a control in it, refuses, or says something.
   - 1f: type " more" at the end of 1f's text.

   **Ctrl+S.**

   **Part 2, Word 2010** (revised 2026-10-03 after part 1: Word 2010 maps no rich text, so what
   matters is whether it opens these files, the editor's round-trip oracle being Word 2010; no VBA):
   - Open `35a-word15.docx`, Word 15's forms (`w15:dataBinding`, `w16sdtdh:storeItemChecksum`).
     Note any message, and what 1a to 1f show. Type " 2010" at the end of 1f's text, and **File >
     Save As** `35c-word2010.docx`.
   - Open `35d-editor.docx` (in `__tmp/35/`, made from `35-input.docx` by the editor's
     `test/fixtures/build/check35-for-word.mjs`). In it: 1a mapped to `name` with
     `w15:dataBinding`, showing "Ann", and inside it a plain text control bound to `rich2`
     (`w:dataBinding`, the placeholder); 1b mapped to `rich1` with `w15:dataBinding` (the
     placeholder); 1f untouched; "After." in a `w15:repeatingSection`; `w15` listed in `mc:Ignorable`
     by the editor's save. Note any message, and what 1a, 1b, 1f and "After." show. Type " 2010" at
     the end of 1f's text, and **File > Save As** `35e-word2010.docx`.

   Put the files in `fixtures/revisions/check35/` (or leave them in `__tmp/35/`).

   The questions:
   - whether Word 15 writes a binding on a rich text control, as `w:dataBinding` or
     `w15:dataBinding`, and whether Word 2010's `SetMapping` refuses one (`False`, or an error);
   - which side wins when a rich text control holding content is mapped: 1a to a node holding "Ann",
     1b to an empty node. For a plain text control the node's value wins (check 34);
   - what a mapped rich text control's edit writes to its node (1b, 1c, 1f): its text, or its content
     as escaped WordprocessingML (a Flat OPC `pkg:package`, as `invoice2013.docx`'s "terms" node holds
     it). And whether a second control mapped to the same node (1d) shows the formatting;
   - whether the pane nests a plain text control inside a mapped rich text control (1e);
   - whether a rich text control bound with `w:dataBinding` (1f) is honoured: refreshed from its node
     on open, and written to its node by an edit, in Word 15 and in Word 2010;
   - whether Word 2010 opens `35a` and the editor's `35d` without complaint, which `w15` bindings
     survive its save, and whether it takes 1f's `w:dataBinding` on rich text as Word 15 does
     (showing "Bea", writing " 2010" to `bound`).

   The engine today: `xmlMapping.setMapping` writes `w:dataBinding` on any control, rich text
   included, and `applyBindingTo` puts the node's text into an untyped control and skips one with
   `w:richText` ("explicit rich text is bound from flat OPC or XHTML in docx4j; deferred").
   `XmlMapping.dataBinding` reads `w15:dataBinding` too. The editor binds a rich text control with
   `w:dataBinding` and shows the node's text in it.

   **Part 1 run 2026-10-03 (current Word), `fixtures/revisions/check35/35a-word15.docx`; part 2 run
   the same day (Word 2010), below.** Jason's notes, then the file:
   - **Rich text is mapped with `w15:dataBinding`** (canonical path, `w:prefixMappings=""`), with
     `w16sdtdh:storeItemChecksum` on 1a's. The pane's Insert Content Control offers Rich Text, Plain
     Text, Picture, Check Box, Combo Box, Drop-Down List and Date Picker.
   - **The node's value wins on mapping:** 1a, holding "quick brown", showed "Ann". 1b, mapped to the
     empty `rich1`, showed the placeholder and "became a single para": the saved file has 1b as a
     run-level control at the start of the "1c." label's paragraph, its two paragraphs gone.
   - **An edit writes the content as escaped Flat OPC:** after 1e, `name` holds a whole `pkg:package`
     (document, styles, the glossary; about 51 KB) of 1a's content, the nested control included.
   - **Typing in a mapped control showing its placeholder wrote nothing:** 1b ("     edited") and 1c
     ("Hello **world**" and "Second line.", in the `PlaceholderText` style) keep `w:showingPlcHdr`,
     and `rich1` and `rich2` are still empty, so 1d, mapped to `rich2`, shows the placeholder. The
     typed text looked grey in Word (Jason): it stayed placeholder text.
   - **1e: the pane nests a plain text control inside the mapped 1a,** with the alert "The custom XML
     node is already mapped to a rich text content control, so it can't be mapped to a plain text
     content control." The nested control is saved with a `w:dataBinding` to `name` anyway, showing
     its placeholder.
   - **1f: a rich text control bound with `w:dataBinding` is taken as plain text:** on open it showed
     the node's "Bea", not "Old text"; Word added `<w:text/>` to it; " more" typed wrote "Bea more" to
     `bound`.

   **Part 2 run 2026-10-03 (Word 2010), `35c-word2010.docx` (from `35a`) and `35e-word2010.docx`
   (from the editor's `35d-editor.docx`, saved by Jason as `35e-Word2010.docx`).** Both opened
   without complaint. Jason's notes on `35a`: "1a shows flatopc xml; 1b shows our '    edited' text;
   1c and 1d both show the placeholder text; 1f 'Bea more'". The files:
   - **Word 2010 drops every `w15:dataBinding` on saving** (1a to 1d in `35c`, 1a and 1b in `35e`),
     with `w16sdtdh:storeItemChecksum`, and the `w15:repeatingSection` and its item in `35e`. The
     controls stay, as rich text controls.
   - **It honours the `w:dataBinding` Word 15 saved on 1e's inert control:** the plain text control
     inside 1a shows `name`'s Flat OPC as text, split into runs by proofing. That is the "flatopc
     xml" Jason saw.
   - **It shows 1c's typed text as the placeholder:** the content Word 15 kept under
     `w:showingPlcHdr` is replaced by the placeholder text.
   - **It takes a rich text control with `w:dataBinding` as plain text, as Word 15 does:** 1f in `35e`
     showed the node's "Bea" in place of "Old text", and gained `<w:text/>`.
   - **It drops a `w:rStyle` naming a style the styles part lacks:** the editor's placeholder runs in
     `35e` lost `PlaceholderText`, which `35d`'s styles part, the engine's, does not define. Word
     15's save (`35a`) defines it, so `35c` kept it.
   - The data part is as it was in each: `name` still Flat OPC in `35c`, "Ann" in `35e`.

36. Extension content that is **not** declared ignorable (2026-10-04, at Jason's word; objects CR-008's
   open question 2, and the premise of CR-002 section 40 item 2). Check 35 opened, in Word 2010, a
   file whose `w15` content was declared (`mc:Ignorable="w15"`); nobody has opened one where it was
   not. ECMA-376 Part 3 says a consumer that does not understand a namespace not declared ignorable
   must treat the markup as an error. Does Word?

   Four files, made by `scripts/make-check36-inputs.mjs`, in `fixtures/revisions/check36/` with copies
   in the shared `__tmp/36/`. Each pair is alike but for the root's `mc:Ignorable`:
   - `36a-w15-undeclared.docx`: a block content control titled 36 with
     `<w15:appearance w15:val="tags"/>`; `xmlns:w15` declared, **no** `mc:Ignorable`;
   - `36b-w15-declared.docx`: the same with `mc:Ignorable="w15"` (the control: check 35's form);
   - `36c-w16du-undeclared.docx`: a tracked insertion ("inserted text", author Check 36) carrying
     the attribute `w16du:dateUtc`; `xmlns:w16du` declared, **no** `mc:Ignorable`;
   - `36d-w16du-declared.docx`: the same with `mc:Ignorable="w16du"`.

   **Part 1, Word 2010**, which knows neither `w15` nor `w16du`. For each of the four files, in the
   order a, b, c, d:
   - Open it from `__tmp/36/` (double-click, or **File > Open**). Note **exactly** what Word does:
     opens it with no message; shows a message (write down its words, and which button you pressed);
     offers to recover or repair; or refuses to open it.
   - If it opened: note whether the three paragraphs are there ("Check 36, 36x: ...", the control's
     "Text in the content control." or "Kept text, then inserted text.", and "After."), and for
     36c and 36d whether "inserted text" shows as a tracked insertion (**Review** tab; All Markup
     or Final: Show Markup).
   - If it opened: **File > Save As**, in `__tmp/36/`, the same name with `-word2010` added
     (`36a-w15-undeclared-word2010.docx`, and so on). Close it.

   **Part 2, current Word**, which knows both namespaces: open `36a-w15-undeclared.docx` and
   `36c-w16du-undeclared.docx` only. Note any message. If each opens, **File > Save As** with
   `-word365` added to the name (Save As first if AutoSave is on), and close it.

   Leave the saved files in `__tmp/36/`.

   The questions:
   - whether Word 2010 opens extension content its root does not declare ignorable: an unknown
     **element** in a `w:sdtPr` (36a) and an unknown **attribute** on a `w:ins` (36c), each against
     its declared twin (36b, 36d);
   - if it opens them, what its save keeps of the content, and whether its save differs between
     the declared and undeclared file;
   - whether current Word minds the missing declaration for namespaces it knows.

   Why it matters: if Word 2010 refuses 36a or 36c, the objects package's CR-008 (the marshaller
   declares every extension namespace a part uses) is a correctness fix, and `insertOoxml` and
   `insertXml`, which declare nothing for content a caller brings, write files Word 2010 cannot
   open. If it opens both, the declaration is tidiness and CR-008 stays a safeguard.

   The engine today: it lists `w15` and `w16du` in `mc:Ignorable` for the elements and attributes it
   writes itself (CR-002 sections 29 and 40), and nothing for inserted content.

   **Run 2026-10-04 by Jason, Word 2010 and Word 365; the saves are in `fixtures/revisions/check36/`.**
   - **Word 2010 refuses an undeclared extension element.** `36a` did not open: "problems with the
     contents", an unspecified error at `document.xml` line 1 column 440, and no offer to repair.
     Column 440 is the end of `<w15:appearance w15:val="tags"/>`. Its declared twin `36b` opened with
     all three paragraphs.
   - **Word 2010 tolerates an undeclared extension attribute.** `36c` opened, as `36d` did, and
     "inserted text" showed as a tracked insertion in both.
   - **Word 2010's saves drop what it does not know**, declared or not: `36b`'s control lost
     `w15:appearance`, and the `w:ins` of `36c` and `36d` lost `w16du:dateUtc`; each root lists
     `w14 wp14`. The two `w16du` saves do not differ in the insertion.
   - **Word 365 does not mind the missing declaration** for namespaces it knows: `36a` and `36c`
     opened without a message, and its saves keep the element and the attribute and write its own
     full `mc:Ignorable` list.

   So the declaration is required for an extension **element** and not, as far as this measures, for
   an **attribute**: one element (`w15:appearance` in a `w:sdtPr`) and one attribute
   (`w16du:dateUtc` on a `w:ins`) were tried, in Word 2010 only among the versions that know
   neither. `insertOoxml` and `insertXml` bringing a `w15` element into a part that does not list
   `w15` write a file Word 2010 cannot open. CR-002 section 40 item 2's premise is confirmed, and
   objects CR-008 (its section 8) is the fix for every writer. **Fixed the same day** by
   `@docx4j/generated-objects-ts` 0.3.1, which this package now requires: inserted content has its
   extension namespaces declared.

37. Escaped XHTML bound through an `od:xpath` entry (CR-005 section 8.2, 2026-10-08): the files are
   built by `fixtures/check37/build.mjs`, and its `README.md` says what to look for in Word 2010 and
   Word 15. **Run 2026-10-08 (the saves beside the originals):** every control kept in both Words
   with its content and tag, no `w:dataBinding` added to the XHTML controls, nothing restyled;
   `html-binding.test.mjs` holds the saves. CR-005 section 8.2.

38. Picture bindings (CR-005 section 9, 2026-10-08): a picture content control mapped to a base64
   node (does Word fill it on open; what does a picture changed by hand write to the node), and
   `od:Handler=picture` controls bound by the engine, floating without `width` and inline with
   `width=4500`, `width=auto` and in a table cell. Built by `fixtures/check38/build.mjs`; its
   `README.md` says what to look for. **Run 2026-10-08 (the saves beside the originals):** both
   Words fill a mapped picture control from its node on open (the template opened showing the node's
   blue image, the placeholder part dropped from the save); Word 15 writes a picture changed by hand
   back to the node as base64, exactly the new media part's bytes (Word 2010 not tried); the
   `od:Handler` pictures, floating and inline, keep their new images, Word 2010 minding neither the
   tag nor the anchor's `true`/`false` booleans; an SVG part is kept and pointed at by both, drawn as
   a red cross, neither reading SVG. **Word 365** (`38b-handler-bound-word365.docx`, the same day)
   showed the SVG's orange circle, and its save rewrote the blip into Word's form: a PNG it
   rasterised, the SVG kept under `asvg:svgBlip`. `picture-binding.test.mjs` holds the saves.
   CR-005 section 9.

39. A column inserted in each of Word's table layout modes (CR-002 section 42, 2026-10-09): the
   editor's probe (`fixtures/check39/`, built by docx4j-ts-editor's `table-autofit-probe.mjs`), one
   two-column table per mode, Insert Right in the first cell of each. **Run 2026-10-09 in Word 365
   (Jason; `table-autofit-probe-word365.docx`):** fixed copies the width beside the caret and grows
   `w:tblW` past the margin (9,026 to 13,539); AutoFit to window keeps 100% and writes the cells as
   1666, 1667, 1667 pct over a grid of 3005, 3006, 3006; AutoFit to contents stays `auto` with a
   grid measured from the text (593, 222, 597). The editor's `check39.test.mts` holds its command
   to the save; the engine's `addColumns` / `insertColumns`, implemented 2026-10-09 (CR-002
   section 42.5), is held to the same save by `table-columns.test.mjs`. A Word 2010 save is
   optional.

40. Columns: the three points section 42.5 left open (CR-002, 2026-10-09; asked by Jason: "let's
   measure"). Five files from `fixtures/check40/build.mjs`, copies in the shared `__tmp/40/`, its
   `README.md` saying what to do with each in Word 365: the AutoFit-to-window rescale with unequal
   columns (Insert Right in the narrow cell, Insert Left in the wide one: proportional, as the
   engine, or an equal share); Insert Right beside and inside a cell spanning two columns (does Word
   widen the span, as the engine, or split it; what width the new column takes beside a span); and,
   through Office JS with tracking on (`check40-script.js` in Script Lab), `deleteColumns` and
   `addColumns`: whether the deletion tracks anything, and whether the add writes a `w:tcPrChange`
   on the new cells. **Run 2026-10-09 in Word 365 (Jason; the saves and `40c-result.json` beside
   the originals):** the new column copies the column to the right of the boundary (the wide one
   in 40a1, column 3 in 40b2), the rescale is proportional (40a2's shares are the engine's to the
   unit); a boundary inside a spanning cell puts the new cell before it, the span kept (40b1);
   Office JS's deletion removes the cells untracked but records the table change, and its add
   records a `w:tcPrChange` on the new cells too, their text in a `w:ins` with no mark insertion,
   the `w:tblGridChange` holding the grid as it is after the change in both. The engine moved to
   all of it; CR-002 section 42.6 has the table, `table-columns.test.mjs` the saves.

41. Columns: the two points check 40 left (CR-002 section 42.6, 2026-10-09; Jason: "let's measure
   those 2 things"). Three files from `fixtures/check41/build.mjs`, copies in the shared `__tmp/41/`,
   its `README.md` saying what to do with each in Word 365: whose properties a new column's cells
   take when the two neighbouring columns differ in shading, vertical alignment, borders and
   paragraph formatting (Insert Right from the yellow side, Insert Left from the blue side); and,
   through Office JS with tracking on (`check41-script.js` in Script Lab), `addColumns` on an
   AutoFit-to-contents table and an AutoFit-to-window one: what Word writes and lists. **Run
   2026-10-09 in Word 365 (Jason; the saves and `41c-result.json` beside the originals):** the new
   column takes the right-hand cell's shading and its paragraph's properties (right-aligned, an
   italic mark) and not its borders or vertical alignment, from either side; Office JS's add on the
   contents table recorded nothing and listed two `Added`; on the window table it recorded every
   cell with its old share spanning a seven-column merged grid (check 27's form), the new cells as
   `1 pct`, and listed one `Formatted`. The engine moved to all of it; CR-002 section 42.6,
   `table-columns.test.mjs`.
42. Whether Office JS counts a text box's paragraphs as the body's (CR-010 section 12, D7, 2026-10-18;
   Jason: "D7 and D8 agreed, subject to a check in word"). In Word 365 (desktop), open a **copy** of
   `fixtures/parity/numbering-stories.docx`. Its eighth paragraph, "P23 A text box:", ends in an inline
   text box holding three numbered paragraphs, P20 to P22, in an `mc:AlternateContent` (a `wps` Choice
   and a VML Fallback). Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New
   snippet), paste the HTML and Script tabs of the snippet for 42 below over the defaults, click
   **Run the check** and copy the JSON back. Nothing is edited; close without saving. Here
   `Body.paragraphs` gives 11, `Body.text` leaves the box's text out, and `Body.inlinePictures` gives 1
   (the box's inline `w:drawing`). If Office JS gives 14 paragraphs, or the box's text in `body.text`,
   it counts the box as the body's, and D7 is revisited. **Run 2026-10-18 (Jason, Word
   16.0.20326.20232, PC):** 11 paragraphs, with `Body.paragraphs`' texts except that each note
   reference stands as `\u0002` ("P11 Two footnotes\u0002 and\u0002", "P15 An endnote\u0002"; nothing
   for P19's comment reference); `body.text` without the box's text; and **no** inline picture, where
   `Body.inlinePictures` gives 1. So Office JS leaves the box out of the body, and D7 stands; it does
   not count an inline text box as an inline picture. CR-010 section 12.6 records the two findings.
43. Whether Office JS counts an inline chart as an inline picture (CR-010 section 13.4, 2026-10-18;
   Jason: "we'd better check whether Office JS counts an inline chart as a picture"). In Word 365
   (desktop), open `__tmp/43/loadAndSave.docx`, a copy of `fixtures/loadAndSave.docx` made for this
   check. It has 41 paragraphs. The 30th, an empty one, holds an inline picture ("Graphic 2", 72 by
   72 points). The 41st, the last, holds an inline chart ("Chart 2", 347.55 by 202.75 points, a
   `c:chart` in the drawing's `a:graphicData`). The text box in the 26th paragraph is floating
   (`wp:anchor`), so it is no inline picture either way. Open Script Lab (Insert > Get Add-ins >
   Script Lab, then Script Lab > Code > New snippet), paste the HTML and Script tabs of the snippet
   for 43 below over the defaults, click **Run the check** and copy the JSON back. Nothing is edited;
   close without saving. Since CR-010 phase 3, `Body.inlinePictures` gives 1 (the picture), and 2
   before (the chart too). If Office JS lists the chart, a second entry of about 347.55 by 202.75 in
   the 41st paragraph, `inlinePictures` lists inline charts again. **Run 2026-10-18 (Jason, Word
   16.0.20326.20232, PC):** 41 paragraphs and **one** inline picture, 72 by 72 in the 30th paragraph,
   alt text "Badge Tick1 with solid fill", the same as `Body.inlinePictures` gives. The chart is not
   listed, so Office JS does not count an inline chart as an inline picture, and phase 3's rule
   stands. Also found: `getBase64ImageSrc()` gave 1016 characters. That is the length of the base64
   of `word/media/image2.svg`, the SVG the picture's `asvg:svgBlip` embeds. `getBase64()` here gives
   the PNG that the `a:blip` embeds (`image1.png`, 5652). Only the lengths were compared. CR-010
   section 13.4 records both findings, and CR-002 section 52 the SVG fix.
44. What Office JS reports for an SVG picture (CR-002 section 52, 2026-10-18; Jason: "find out what
   Office JS reports for an SVG picture"). In Word 365 (desktop), open `__tmp/44/loadAndSave.docx`, a
   copy of `fixtures/loadAndSave.docx` made for this check. Its one inline picture, in the 30th of 41
   paragraphs, is an SVG icon ("Badge Tick1 with solid fill", 72 by 72 points). Its `a:blip` embeds
   `media/image1.png`, and the `asvg:svgBlip` in the blip's `a:extLst` embeds `media/image2.svg`. Open
   Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet), paste the HTML
   and Script tabs of the snippet for 44 below over the defaults, click **Run the check** and copy the
   JSON back. Nothing is edited; close without saving. The snippet reports:
   - `imageFormat`, or the error if this Word lacks it. `InlinePicture.imageFormat` gives `Png` here,
     the `a:blip`'s part's format.
   - whether `getBase64ImageSrc()` is the SVG's base64 character for character (`equalsSvg`), where
     check 43 compared only lengths. `getBase64()` gives the SVG since CR-002 section 52.
   - the image's first 16 characters: `PHN2Zy` begins an SVG's `<svg`, and `iVBORw0KGgo` a PNG.

   **Run 2026-10-18 (Jason, Word 16.0.20326.20232, PC):** `equalsSvg` true (1016 characters, beginning
   `PHN2ZyB2aWV3Qm94`), so Office JS's image for an SVG picture is the SVG, as `getBase64()` gives since
   CR-002 section 52. `imageFormat` came back `null`, with no error. One picture cannot say whether that
   is the SVG picture's value or every picture's in this build, which is check 45's question.
45. Whether Office JS's `imageFormat` is `null` for an SVG picture only, or for every picture (CR-002
   section 52, 2026-10-18; check 44 gave `null`). In Word 365 (desktop), open
   `__tmp/45/svg-and-plain.docx`. It was made from `fixtures/loadAndSave.docx` by this package (main
   part re-marshalled), with two paragraphs added at the end, so it has 43. Its three inline pictures:
   - in the 30th paragraph, the SVG icon ("Badge Tick1 with solid fill", 72 by 72);
   - in the 42nd, "plain PNG": the icon's own PNG (`a:blip` only, no `svgBlip`), 72 by 72;
   - in the 43rd, "plain GIF": a 2 by 2 GIF shown at 36 by 36.

   Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet), paste the
   HTML and Script tabs of the snippet for 45 below over the defaults, click **Run the check** and copy
   the JSON back. Nothing is edited; close without saving. `InlinePicture.imageFormat` gives `Png`, `Png`
   and `Gif`. If Office JS gives `Png` and `Gif` for the two plain ones and `null` for the icon, `null` is
   its value for an SVG picture. If it gives `null` for all three, this build reports no format for any
   picture. **Run 2026-10-18 (Jason, Word 16.0.20326.20232, PC):** `null` for the icon, `Png` for "plain
   PNG" and `Gif` for "plain GIF". So `null` is Office JS's `imageFormat` for an SVG picture, though
   `@types/office-js` declares the property non-null and `Word.ImageFormat` has an `Svg`. The images were
   the icon's SVG (1016 characters, `PHN2Zy...`), the PNG (5652, `iVBORw0KGgo...`) and the GIF (48), the
   sizes as stated. CR-002 section 52 records it, and the departure Jason chose: this package gives
   `Svg` there.
46. What Word draws for an SVG picture after this package's picture binding has swapped its image
   (CR-002 section 52, 2026-10-18; Jason: "let's do check 46"). The binding's keep mode (CR-005
   section 9) points the `a:blip`'s `r:embed` at the new image and leaves the `asvg:svgBlip` on the
   old SVG. docx4j's `bind.xslt` writes a bare `a:blip` instead.
   - **The file:** `__tmp/46/46-svg-binding.docx`, a copy of `fixtures/check46/46-svg-binding.docx` (built by
     `fixtures/check46/build.mjs`). It holds
     three copies of `loadAndSave.docx`'s SVG icon, a black circle with a tick cut out of it, each
     72 by 72 points, under a label paragraph.
     - **46a:** the icon, not bound. Its `a:blip` is on the icon's PNG and its `svgBlip` on the SVG.
     - **46b:** an `od:Handler=picture` rich text control (no `w:dataBinding`), bound by this package
       to a node holding a solid red PNG.
     - **46c:** a picture content control with a `w:dataBinding` to the same node, bound by this
       package.
     In both bound copies the `a:blip` is on a red PNG and the `svgBlip` still on the icon's SVG.
   - **Steps.** In Word 365 (desktop), open `__tmp/46/46-svg-binding.docx`, and before anything else
     File > Save As `46-svg-binding-word365.docx` in the same folder. That keeps Word's form of the
     file, a refreshed binding included. In the saved copy:
     1. Say what each of 46a, 46b and 46c shows: the black circle with the tick, a red square, or
        something else.
     2. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet),
        paste the HTML and Script tabs of the snippet for 46 below over the defaults, click **Run the
        check** and copy the JSON back.
     3. Close without saving again.
   - **What it decides.**
     - 46b shows the circle: Word draws the stale `svgBlip`, and the binding needs fixing (CR-002
       section 52's options).
     - 46b shows red: Word draws the `a:blip` there.
     - 46c also says whether Word refills a bound picture control from its node when it opens the
       file. The save shows what Word writes for both.
   - **Run 2026-10-18 (Jason, Word 16.0.20326.20232, PC).** The original was left unchanged. Jason's
     screenshot of the saved copy shows the black circle with the tick for 46a and **46b**, and a red
     square for 46c. Word's save is `fixtures/check46/46-svg-binding-word365.docx`.
     - 46a: the icon's SVG (1016 characters, `PHN2Zy...`), `imageFormat` `null`, in no control.
     - 46b: **the icon's SVG** (1016, `PHN2Zy...`), `imageFormat` `null`, in the `RichText` control with
       the `od:Handler=picture` tag. So for Word the picture is the stale `svgBlip`'s old icon, not the
       red PNG that the binding put in the `a:blip`.
     - 46c: a PNG of 464 characters (`iVBORw0KGgo...`), `imageFormat` `Png`, in the `Picture` control.
       That is exactly the length of the node's red PNG (346 bytes) in base64. So Word refilled the
       bound picture control from its node when it opened the file, and holds it as a plain PNG
       picture, the SVG gone.
     - **Word's save:**
       - 46a and 46b: both drawings have the same `a:blip`, on `media/image1.png`, a PNG of 2414 bytes
         that is neither the icon's own PNG (4238) nor the red one. Their `svgBlip` is on the icon's
         SVG. So for 46b Word dropped the red PNG the binding had put in the `a:blip`, and gave it the
         same PNG as 46a, which appears to be one Word made from the SVG.
       - 46c: a bare `<a:blip r:embed="rId10"/>` on the red PNG (`media/image3.png`, 346 bytes), the
         `a:extLst` and its `svgBlip` gone. Word's own binding writes the form docx4j's `bind.xslt`
         writes.
     - CR-002 section 52 records it.
47. Which form an SVG picture's binding should write (CR-002 section 52, 2026-10-18; Jason: "build check
   47"). Check 46 found that Word's own binding of a picture control writes a bare
   `<a:blip r:embed="..."/>`, as docx4j's `bind.xslt` does. This check asks three things: what Word
   writes when the bound value is an SVG (47a), whether Word reads an `a:blip` with only the
   `asvg:svgBlip` and no `r:embed` (47b, Jason's suggestion), and whether Word keeps a picture's effects
   when it rebinds (47c). Two files in `__tmp/47/`, copies of `fixtures/check47/` (built by
   `fixtures/check47/build.mjs`).
   - **`47-bindings.docx`, in Word 365 only.** It has two picture content controls, each with a
     `w:dataBinding` that this package did not apply: each control's picture is a green square, 72 by
     72 points, and Word is left to bind it when it opens the file, as it did in check 46.
     - **47a:** the node holds an SVG, a blue circle. A blue circle means Word bound it, and its save
       shows the form Word writes for a bound SVG. A green square means Word did not bind it.
     - **47c:** the node holds a red PNG, and the green square's `a:blip` has an `a:alphaModFix amt="30000"`
       (30% opacity). Pink means Word bound it and kept the effect; solid red means it bound it and
       dropped the effect; pale green means it did not bind it.
     - Steps: open `__tmp/47/47-bindings.docx`, and before anything else File > Save As
       `47-bindings-word365.docx` in the same folder. In the saved copy, say what 47a and 47c show (or
       take a screenshot), run the snippet for 46 below (Script Lab, as in check 46) and copy the JSON
       back, then close without saving again.
   - **`47-svgblip-only.docx`, in Word 365, then Word 2010, then Word 15.** It has two copies of
     `loadAndSave.docx`'s icon, a black circle with a tick cut out of it, each 72 by 72 points.
     - **Baseline:** the icon as Word writes it, with its `a:blip` on the PNG and its `svgBlip` on the SVG.
     - **47b:** the same icon with `<a:blip><a:extLst><a:ext uri="{96DAC541-...}"><asvg:svgBlip
       r:embed="..."/></a:ext></a:extLst></a:blip>`, no `r:embed` on the `a:blip`. The schema allows
       that (`AG_Blob`'s `r:embed` is optional); Office never writes it.
     - Steps, in each Word: open `__tmp/47/47-svgblip-only.docx` and say whether Word reports a
       problem or offers a repair. Before anything else, File > Save As `47-svgblip-only-word365.docx`
       (or `-word2010.docx`, `-word15.docx`) in the same folder. Say what the baseline and 47b show:
       the circle with the tick, an empty frame, a red cross, or something else. In Word 365 only, run
       the snippet for 46 below and copy the JSON back. Close without saving again.
     - Word 2010 and 15 do not read SVG (check 38), so they show what a reader without SVG support makes
       of 47b. The baseline gives them the `a:blip`'s PNG.
   - **Run 2026-10-18 (Jason; Word 365 16.0.20326.20232, Word 2010, Word 2013).** Jason wrote his
     observations into each save, in bold underline. The saves are in `fixtures/check47/`, the Word 2013
     one renamed `47-svgblip-only-word2013.docx`. No problem on opening was reported.
     - **47a: a blue circle.** Word bound the SVG. Its save writes Office's own SVG picture: the
       `a:blip` on a PNG Word made (`image1.png`, 2348 bytes), and an `asvg:svgBlip` on the node's SVG,
       byte for byte (113 bytes). Office JS gave that SVG (152 characters) and `imageFormat` `null`.
       The node was left as it was.
     - **47c: a pink square.** Word bound the red PNG and kept the effect. Its save is
       `<a:blip r:embed="rId9"><a:alphaModFix amt="30000"/></a:blip>` on the red PNG (346 bytes).
       Office JS gave the red PNG (464 characters), `Png`.
     - **47b, the `svgBlip` alone:**
       - Word 365 drew the circle with the tick, as for the baseline. Saving, it gave 47b an `r:embed`
         on a PNG it made (2414 bytes), shared with the baseline's, so the save holds Office's own form
         twice.
       - Word 2010 and Word 2013 showed **a red cross** for 47b and the circle with the tick for the
         baseline (from its PNG). Each kept 47b's `a:blip` as it was, without `r:embed`.
     - CR-002 section 52 records it, with the binding rule it gives, which the binding now follows.
       `picture-binding.test.mjs` holds the binding to these saves and to check 46's.

A small Node script for 1 to 3 is:

```js
import { writeFile, readFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

// 1 and 2: the round trip, with the main part re-marshalled
const pkg = await WordprocessingMLPackage.load(await readFile('test/fixtures/loadAndSave.docx'));
await pkg.getMainDocumentPart().getContents();
await writeFile('out.docx', await pkg.save());

// 3: a created document, with the theme part CR-001 Phase B step 4 added
for (const theme of ['2023', '2013', '2007']) {
  const fresh = await WordprocessingMLPackage.createPackage({ pageSize: 'A4', defaultTheme: theme });
  fresh.body.insertParagraph('Hello World', 'End').styleBuiltIn = 'Heading1';
  fresh.body.insertParagraph('Body text in the theme face.', 'End');
  await writeFile(`check3-${theme}.docx`, await fresh.save());   // nine zip entries, not eight
}
```

And one for 6 to 8 (run from the repository root after `npm run build`; `logo.png` is any PNG):

```js
import { writeFile, readFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from './dist/index.mjs';

// 6. tables and pictures
const a = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
const rows = [['Item', 'Price'], ...Array.from({ length: 60 }, (_, i) => [`item ${i + 1}`, `$${i + 1}`])];  // spans a page break
const table = a.body.insertTable(rows.length, 2, 'End', rows);
table.styleBuiltIn = 'TableGrid';
table.headerRowCount = 1;
a.body.insertInlinePictureFromBase64((await readFile('logo.png')).toString('base64'), 'End', { altTextDescription: 'The logo' });
await writeFile('check6-a.docx', await a.save());
const b = await WordprocessingMLPackage.createPackage();
await b.body.insertOoxml(await a.saveFlatOpc(), 'End');
await writeFile('check6-b.docx', await b.save());           // the image once, as word/media/image1.png

// 7. content controls
const c = await WordprocessingMLPackage.load(await readFile('test/fixtures/invoice.docx'));
(await c.getBody()).contentControls[0].insertText('Jane Doe', 'Replace');
await writeFile('check7.docx', await c.save());

// 8. comments
const d = await WordprocessingMLPackage.createPackage();
d.author = { name: 'Acceptance Tester', initials: 'AT' };
const p = d.body.insertParagraph('The quick brown fox jumps over the lazy dog.', 'End');
const [range] = p.search('brown fox');
const comment = await range.insertComment('Is it brown?');
await comment.reply('Yes, brown.');
comment.resolved = true;
await writeFile('check8-a.docx', await d.save());
const e = await WordprocessingMLPackage.load(await readFile('test/fixtures/loadAndSave.docx'));
const body = await e.getBody();
await body.paragraphs[0].insertComment('A new comment on the first paragraph');
for (const existing of await body.getComments()) if (existing.authorName !== 'docx4j') await existing.delete();
await writeFile('check8-b.docx', await e.save());

// 9. custom XML and the bindings (needs the optional `xpath` package in Node)
const f = await WordprocessingMLPackage.load(await readFile('test/fixtures/invoice.docx'));
await f.customXmlParts.load();
(await f.getBody()).contentControls[0].insertText('Jane Doe', 'Replace');
await f.customXmlParts.updateFromContentControls();
await writeFile('check9-a.docx', await f.save());    // Word shows Jane Doe, not Joe Bloggs

const g = await WordprocessingMLPackage.load(await readFile('test/fixtures/invoice2013.docx'));
await g.customXmlParts.load();
g.customXmlParts.items[0].selectSingleNode('/invoice/customer/company').text = 'Acme Pty Ltd';
await g.customXmlParts.applyBindings();
await writeFile('check9-b.docx', await g.save());
```

And one for 12 and 13 (the pptx and xlsx of CR-001 Phase C):

```js
import { writeFile, readFile } from 'node:fs/promises';
import { PresentationMLPackage, SpreadsheetMLPackage } from './dist/index.mjs';
import { createRow, createCell } from '@docx4j/generated-objects-ts/factory/org_xlsx4j_sml';

// 12. a created presentation, at two slide sizes: slide 1 shows the prompts, slide 2 the text
for (const [name, options] of [['16x9', { slideSize: 'SCREEN16x9' }], ['a4-portrait', { slideSize: 'A4', landscape: false }]]) {
  const deck = await PresentationMLPackage.createPackage(options);
  if (name === '16x9') await deck.addSlide({ title: 'The second slide', body: 'One body line\nAnd another' });
  await writeFile(`check12-${name}.pptx`, await deck.save());
}

// 13. a created workbook: values into Sales, then Notes appended and Cover inserted first, so the
// tabs are Cover (empty), Sales (Q1 and Q2), Notes (empty) - `sheets` order, not creation order
const book = await SpreadsheetMLPackage.createPackage();
const sheet = book.createWorksheetPart('Sales');
book.createWorksheetPart('Notes');
book.createWorksheetPart('Cover', 0);
sheet.contents.sheetData.row = [['Q1', '100'], ['Q2', '120']].map(([label, value], i) => createRow({
  r: i + 1,
  c: [createCell({ r: `A${i + 1}`, t: 'inlineStr', is: { t: label } }), createCell({ r: `B${i + 1}`, v: value })],
}));
await writeFile('check13.xlsx', await book.save());

// 13 (continued): the two fixtures re-marshalled
const deck = await PresentationMLPackage.load(await readFile('test/fixtures/loadAndSave.pptx'));
await deck.getMainPresentationPart().getContents();
await writeFile('check13-remarshalled.pptx', await deck.save());
const wb = await SpreadsheetMLPackage.load(await readFile('test/fixtures/loadAndSave.xlsx'));
await wb.getWorkbookPart().getContents();
await writeFile('check13-remarshalled.xlsx', await wb.save());
```

And one for 10 and 11 (run from the repository root after `npm run build`):

```js
import { writeFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from './dist/index.mjs';
const OUT = '.';

// 10. tracked changes: every kind, under one author
const t = await WordprocessingMLPackage.createPackage();
t.author = { name: 'Acceptance Tester', initials: 'AT' };
const p1 = t.body.insertParagraph('The quick brown fox jumps over the lazy dog.', 'End');
const p2 = t.body.insertParagraph('A second paragraph that will be deleted.', 'End');
const p3 = t.body.insertParagraph('A third paragraph whose formatting changes.', 'End');
const table = t.body.insertTable(2, 2, 'End', [['a', 'b'], ['c', 'd']]);
t.changeTrackingMode = 'TrackAll';
p1.search('brown fox')[0].insertText('red hen', 'Replace');      // w:del + w:ins
p1.insertText(' Inserted at the end.', 'End');                    // w:ins
p2.delete();                                                       // paragraph mark and content deleted
p3.getRange().font.bold = true;                                    // w:rPrChange
p3.alignment = 'Centered';                                         // w:pPrChange
p3.insertParagraph('A new tracked paragraph.', 'After');           // inserted paragraph mark
table.addRows('End', 1, [['e', 'f']]);                              // w:trPr/w:ins
table.deleteRows(0, 1);                                            // w:trPr/w:del
await p1.search('lazy dog')[0].insertComment('A comment made while tracking is on: no w:ins around it.');
console.log('tracked changes:', t.body.getTrackedChanges().map((c) => c.type).join(', '));
await writeFile(`${OUT}/check10.docx`, await t.save());

// 11. lists: numbered, a nested level, a restart, a bulleted list
const l = await WordprocessingMLPackage.createPackage();
const a = l.body.insertParagraph('First item', 'End');
const list = await a.startNewList();
const b = l.body.insertParagraph('Second item', 'End'); b.attachToList(list.id, 0);
const c = l.body.insertParagraph('Nested under the second', 'End'); c.attachToList(list.id, 1);
const d = l.body.insertParagraph('Third item', 'End'); d.attachToList(list.id, 0);
l.body.insertParagraph('Some text between the lists.', 'End');
const restarted = list.restart();
const e = l.body.insertParagraph('Restarts at one', 'End'); e.attachToList(restarted.id, 0);
const f = l.body.insertParagraph('And continues at two', 'End'); f.attachToList(restarted.id, 0);
const g = l.body.insertParagraph('A bullet', 'End');
const bullets = await g.startNewList({ bullet: true });
const h = l.body.insertParagraph('Another bullet', 'End'); h.attachToList(bullets.id, 0);
console.log('labels:', [...l.body.listLabels().values()].map((x) => x.listString).join(' | '));
await writeFile(`${OUT}/check11.docx`, await l.save());

// 11b. two lists sharing one w:abstractNum (a restart); editing one must leave the other alone
const m = await WordprocessingMLPackage.createPackage();
const m1 = m.body.insertParagraph('Original list, item one', 'End');
const orig = await m1.startNewList();
const m2 = m.body.insertParagraph('Original list, item two', 'End'); m2.attachToList(orig.id, 0);
const again = orig.restart();
const m3 = m.body.insertParagraph('Restarted list, item one (edited to upper roman)', 'End'); m3.attachToList(again.id, 0);
const m4 = m.body.insertParagraph('Restarted list, item two', 'End'); m4.attachToList(again.id, 0);
again.setLevelNumbering(0, 'UpperRoman');                          // copies the shared abstract first
console.log('shared-abstract edit:', [...m.body.listLabels().values()].map((x) => x.listString).join(' | '));
await writeFile(`${OUT}/check11-b.docx`, await m.save());
```

And the Script Lab snippet for 15 (Word, Office JS; WordApi 1.3 for `Range.hyperlink` and
`styleBuiltIn`). The HTML tab:

```html
<button id="run">Run the check</button>
<p>Copy this back:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

async function check() {
  const report = { host: Office.context.diagnostics, cases: {} };
  const cases = [
    ['A', 'Case A: an external link on this text', 'this text', async (r) => { r.hyperlink = 'https://example.com/a'; }],
    ['B', 'Case B: an internal link on this text', 'this text', async (r) => { r.hyperlink = '#_top'; }],
    ['C', 'Case C: a link on these strong words', 'these strong words', async (r, context) => {
      r.styleBuiltIn = Word.Style.strong; await context.sync(); r.hyperlink = 'https://example.com/c'; }],
    ['D', 'Case D: a link on this bold red text', 'this bold red text', async (r, context) => {
      r.font.bold = true; r.font.color = '#FF0000'; await context.sync(); r.hyperlink = 'https://example.com/d'; }],
    ['E', 'Case E: a link set and then removed here', 'removed here', async (r, context) => {
      r.hyperlink = 'https://example.com/e'; await context.sync(); r.hyperlink = ''; }],
  ];
  for (const [id, text, target, act] of cases) {
    const entry = (report.cases[id] = {});
    try {
      await Word.run(async (context) => {
        const paragraph = context.document.body.insertParagraph(text, 'End');
        const range = paragraph.search(target, { matchCase: true }).getFirst();
        await act(range, context);
        await context.sync();
        range.load('hyperlink,style,styleBuiltIn,font/color,font/underline,font/bold');
        await context.sync();
        Object.assign(entry, {
          hyperlink: range.hyperlink, style: range.style, styleBuiltIn: range.styleBuiltIn,
          color: range.font.color, underline: range.font.underline, bold: range.font.bold,
        });
      });
    } catch (e) {
      entry.error = String(e && e.message || e);
    }
  }
  await Word.run(async (context) => {
    const ooxml = context.document.body.getOoxml();
    await context.sync();
    const doc = new DOMParser().parseFromString(ooxml.value, 'application/xml');
    const xml = (node) => new XMLSerializer().serializeToString(node).replace(/ xmlns:\w+="[^"]*"/g, '');
    const all = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
    const part = (name) => all(doc, PKG, 'part').find((p) => p.getAttributeNS(PKG, 'name') === name);
    const main = part('/word/document.xml');
    const styles = part('/word/styles.xml');
    for (const p of main ? all(main, W, 'p') : []) {
      const m = /^Case ([A-E]):/.exec(p.textContent);
      if (m) report.cases[m[1]].xml = xml(p);
    }
    report.styles = styles
      ? all(styles, W, 'style')
          .filter((s) => ['Hyperlink', 'FollowedHyperlink'].includes(s.getAttributeNS(W, 'styleId')))
          .map(xml)
      : 'no styles part in body.getOoxml()';
  });
  out.value = JSON.stringify(report, null, 2);
}
```

And the Script Lab snippet for 16 (Word, Office JS, WordApi 1.6), on a copy of
`fixtures/revisions/revisions-word15.docx`. The HTML tab:

```html
<select id="scenario">
  <option value="all">All (restores the body between scenarios)</option>
  <option value="accept-destination">Accept the move's destination only</option>
  <option value="accept-source">Accept the move's source only</option>
  <option value="reject-destination">Reject the move's destination only</option>
  <option value="reject-source">Reject the move's source only</option>
  <option value="accept-all">Accept all</option>
  <option value="reject-all">Reject all</option>
  <option value="links">The hyperlink cases only</option>
</select>
<button id="run">Run the check</button>
<p>Copy this back:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const MOVED = 'second paragraph of the move';
const SCENARIOS = {
  'accept-destination': { act: 'accept', index: 0 },
  'accept-source': { act: 'accept', index: 1 },
  'reject-destination': { act: 'reject', index: 0 },
  'reject-source': { act: 'reject', index: 1 },
  'accept-all': { act: 'accept', all: true },
  'reject-all': { act: 'reject', all: true },
};
const LINKS = [
  ['F', 'Case F: alpha beta gamma', [['alpha beta', 'https://example.com/f1'], ['beta gamma', 'https://example.com/f2']]],
  ['G', 'Case G: alpha beta gamma', [['alpha beta gamma', 'https://example.com/g1'], ['beta', 'https://example.com/g2']]],
  ['H', 'Case H: alpha beta gamma', [['alpha beta gamma', 'https://example.com/h1'], ['beta', '']]],
];
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () =>
  check(document.getElementById('scenario').value).catch((e) => { out.value = String(e.stack || e); }));

const all = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
const xml = (node) => new XMLSerializer().serializeToString(node)
  .replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*)="[^"]*"/g, '');

/** The body's block-level children, from the document's own body OOXML. */
async function blocks(context) {
  const ooxml = context.document.body.getOoxml();
  await context.sync();
  const doc = new DOMParser().parseFromString(ooxml.value, 'application/xml');
  const main = all(doc, PKG, 'part').find((p) => p.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const body = main && all(main, W, 'body')[0];
  return body ? Array.from(body.childNodes).filter((n) => n.nodeType === 1) : [];
}

/** Section 3, the move: everything between its heading and section 4's, less the instructions. */
async function section(context) {
  const kids = await blocks(context);
  const start = kids.findIndex((n) => n.localName === 'p' && n.textContent.trim() === '3. A move');
  const end = kids.findIndex((n, i) => i > start && n.localName === 'p' && n.textContent.startsWith('4. '));
  if (start < 0 || end < 0) return `section 3 not found (${start}, ${end})`;
  return kids.slice(start + 1, end).filter((n) => !n.textContent.startsWith('[')).map(xml);
}

async function listed(context, filter) {
  const changes = context.document.body.getTrackedChanges();
  changes.load('items/type,items/text,items/author');
  await context.sync();
  return { changes, list: changes.items.map((c, i) => ({ i, type: c.type, author: c.author, text: c.text.slice(0, 70) }))
    .filter((c) => !filter || c.text.includes(filter)) };
}

async function check(which) {
  const report = { host: Office.context.diagnostics, which, moves: {}, links: {} };
  let original;
  await Word.run(async (context) => {
    context.document.load('changeTrackingMode');
    await context.sync();
    report.trackingModeBefore = context.document.changeTrackingMode;
    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    const ooxml = context.document.body.getOoxml();
    await context.sync();
    original = ooxml.value;
    report.listed = (await listed(context)).list;          // how Office JS lists every change, the move's among them
  });
  const names = which === 'all' ? Object.keys(SCENARIOS) : which === 'links' ? [] : [which];
  for (const name of names) {
    const entry = (report.moves[name] = {});
    try {
      await Word.run(async (context) => {
        if (which === 'all') { context.document.body.insertOoxml(original, 'Replace'); await context.sync(); }
        entry.before = await section(context);
        const { changes, list } = await listed(context);
        entry.moveBefore = list.filter((c) => c.text.includes('paragraph of the move'));
        const scenario = SCENARIOS[name];
        if (scenario.all) {
          if (scenario.act === 'accept') changes.acceptAll(); else changes.rejectAll();
        } else {
          const move = changes.items.filter((c) => c.text.includes(MOVED));
          const target = move[scenario.index];
          if (!target) throw new Error(`no change ${scenario.index} of the move (${move.length} listed)`);
          entry.acted = { type: target.type, text: target.text.slice(0, 70) };
          if (scenario.act === 'accept') target.accept(); else target.reject();
        }
        await context.sync();
        entry.after = await section(context);
        entry.moveAfter = (await listed(context, 'paragraph of the move')).list;
      });
    } catch (e) {
      entry.error = String(e && e.message || e);
    }
  }
  if (which === 'all' || which === 'links') {
    for (const [id, text, steps] of LINKS) {
      const entry = (report.links[id] = {});
      try {
        await Word.run(async (context) => {
          context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
          const paragraph = context.document.body.insertParagraph(text, 'End');
          for (const [target, address] of steps) {
            paragraph.search(target, { matchCase: true }).getFirst().hyperlink = address;
            await context.sync();
          }
          const words = ['alpha', 'beta', 'gamma'].map((w) => paragraph.search(w, { matchCase: true }).getFirst());
          words.forEach((r) => r.load('hyperlink'));
          await context.sync();
          entry.reads = Object.fromEntries(words.map((r, i) => [['alpha', 'beta', 'gamma'][i], r.hyperlink]));
          const p = (await blocks(context)).find((n) => n.localName === 'p' && n.textContent.startsWith(`Case ${id}:`));
          entry.xml = p ? xml(p) : 'not found';
        });
      } catch (e) {
        entry.error = String(e && e.message || e);
      }
    }
  }
  out.value = JSON.stringify(report, null, 2);
}
```

And the Script Lab snippet for 17 (Word, Office JS, WordApi 1.6), in a new blank document. The HTML
tab:

```html
<button id="run">Run the check</button>
<p>Copy this back:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const MOVED = 'second paragraph of the move';
const A = 'w:author="Author A" w:date="2026-09-28T05:47:00Z"';
// section 3 of fixtures/revisions/revisions-word15.docx, as check 16 recorded it
const BLOCKS = `<w:p><w:pPr><w:rPr><w:moveTo w:id="6" ${A}/></w:rPr></w:pPr><w:moveToRangeStart w:id="7" ${A} w:name="move241466837"/>`
  + `<w:moveTo w:id="8" ${A}><w:r><w:t>The second paragraph of the move: it goes before the first.</w:t></w:r></w:moveTo></w:p>`
  + '<w:moveToRangeEnd w:id="7"/>'
  + '<w:p><w:r><w:t>The first paragraph of the move: it stays where it is.</w:t></w:r></w:p>'
  + `<w:p><w:pPr><w:rPr><w:moveFrom w:id="9" ${A}/></w:rPr></w:pPr><w:moveFromRangeStart w:id="10" ${A} w:name="move241466837"/>`
  + `<w:moveFrom w:id="11" ${A}><w:r><w:t>The second paragraph of the move: it goes before the first.</w:t></w:r></w:moveFrom></w:p>`
  + '<w:moveFromRangeEnd w:id="10"/>'
  + '<w:p><w:r><w:t>The third paragraph of the move: it stays where it is.</w:t></w:r></w:p>';
const PACKAGE = `<pkg:package xmlns:pkg="${PKG}">`
  + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
  + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  + '</pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + `<pkg:xmlData><w:document xmlns:w="${W}"><w:body>${BLOCKS}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
const SCENARIOS = {
  'accept-destination': { act: 'accept', index: 0 },
  'accept-source': { act: 'accept', index: 1 },
  'reject-destination': { act: 'reject', index: 0 },
  'reject-source': { act: 'reject', index: 1 },
  'accept-all': { act: 'accept', all: true },
  'reject-all': { act: 'reject', all: true },
};
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const all = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
const xml = (node) => new XMLSerializer().serializeToString(node)
  .replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*)="[^"]*"/g, '');

async function blocks(context) {
  const ooxml = context.document.body.getOoxml();
  await context.sync();
  const doc = new DOMParser().parseFromString(ooxml.value, 'application/xml');
  const main = all(doc, PKG, 'part').find((p) => p.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const body = main && all(main, W, 'body')[0];
  return body ? Array.from(body.childNodes).filter((n) => n.nodeType === 1 && n.localName !== 'sectPr').map(xml) : [];
}

async function listed(context) {
  const changes = context.document.body.getTrackedChanges();
  changes.load('items/type,items/text,items/author');
  await context.sync();
  return { changes, list: changes.items.map((c, i) => ({ i, type: c.type, author: c.author, text: c.text })) };
}

async function check() {
  const report = { host: Office.context.diagnostics, scenarios: {} };
  for (const [name, scenario] of Object.entries(SCENARIOS)) {
    const entry = (report.scenarios[name] = {});
    try {
      await Word.run(async (context) => {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        context.document.body.clear();
        context.document.body.insertOoxml(PACKAGE, 'Start');
        await context.sync();
        entry.before = await blocks(context);
        const { changes, list } = await listed(context);
        entry.listedBefore = list;
        if (scenario.all) {
          if (scenario.act === 'accept') changes.acceptAll(); else changes.rejectAll();
        } else {
          const move = changes.items.filter((c) => c.text.includes(MOVED));
          const target = move[scenario.index];
          if (!target) throw new Error(`no change ${scenario.index} holding the moved text (${move.length} listed)`);
          entry.acted = { type: target.type, text: target.text };
          if (scenario.act === 'accept') target.accept(); else target.reject();
        }
        await context.sync();
        entry.after = await blocks(context);
        entry.listedAfter = (await listed(context)).list;
      });
    } catch (e) {
      entry.error = String(e && e.message || e);
    }
  }
  out.value = JSON.stringify(report, null, 2);
}
```

And the Script Lab snippet for 19 (Word, Office JS, WordApi 1.3), in a new blank document. The HTML
tab is check 15's (a **Run the check** button and a text box). The Script tab:

```js
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

// [id, linked first (target, address) or null, where the caret goes, the address given at the caret]
const CASES = [
  ['I', ['alpha beta gamma', 'https://example.com/i1'], 'inside beta', 'https://example.com/i2'],
  ['J', null, 'inside beta', 'https://example.com/j1'],
  ['K', ['alpha', 'https://example.com/k1'], 'after alpha', 'https://example.com/k2'],
  ['L', ['alpha beta gamma', 'https://example.com/l1'], 'inside beta', ''],
];
const all = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
const xml = (node) => new XMLSerializer().serializeToString(node)
  .replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*)="[^"]*"/g, '');

async function check() {
  const report = { host: Office.context.diagnostics, cases: {} };
  for (const [id, link, where, address] of CASES) {
    const entry = (report.cases[id] = {});
    try {
      await Word.run(async (context) => {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        const paragraph = context.document.body.insertParagraph(`Case ${id}: alpha beta gamma`, 'End');
        if (link) { paragraph.search(link[0], { matchCase: true }).getFirst().hyperlink = link[1]; await context.sync(); }
        const caret = where === 'after alpha'
          ? paragraph.search('alpha', { matchCase: true }).getFirst().getRange('End')
          : paragraph.search('et', { matchCase: true }).getFirst().getRange('Start');   // b|eta
        caret.load('hyperlink');
        await context.sync();
        entry.caretBefore = caret.hyperlink;
        caret.hyperlink = address;
        await context.sync();
        caret.load('hyperlink,text');
        const words = ['alpha', 'beta', 'gamma'].map((w) => paragraph.search(w, { matchCase: true }).getFirst());
        words.forEach((r) => r.load('hyperlink'));
        paragraph.load('text');
        await context.sync();
        Object.assign(entry, {
          caretAfter: caret.hyperlink, caretText: caret.text, paragraphText: paragraph.text,
          reads: Object.fromEntries(words.map((r, i) => [['alpha', 'beta', 'gamma'][i], r.hyperlink])),
        });
      });
    } catch (e) {
      entry.error = String(e && e.message || e);
    }
  }
  await Word.run(async (context) => {
    const ooxml = context.document.body.getOoxml();
    await context.sync();
    const doc = new DOMParser().parseFromString(ooxml.value, 'application/xml');
    const main = all(doc, PKG, 'part').find((p) => p.getAttributeNS(PKG, 'name') === '/word/document.xml');
    for (const p of main ? all(main, W, 'p') : []) {
      const m = /^Case ([IJKL]):/.exec(p.textContent);
      if (m) report.cases[m[1]].xml = xml(p);
    }
  });
  out.value = JSON.stringify(report, null, 2);
}
```

And the Script Lab snippet for 22 (Word, Office JS, WordApi 1.6), run in `22-deleted.docx`. The HTML
tab is check 15's (a **Run the check** button and a text box). The Script tab:

```js
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const SCENARIOS = [0, 1, 2].flatMap((i) => [[`accept-${i}`, i, 'accept'], [`reject-${i}`, i, 'reject']]);
const all = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
const xml = (node) => new XMLSerializer().serializeToString(node)
  .replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*)="[^"]*"/g, '');

/** The body's paragraphs, from the document's own body OOXML. */
async function paragraphs(context) {
  const ooxml = context.document.body.getOoxml();
  await context.sync();
  const doc = new DOMParser().parseFromString(ooxml.value, 'application/xml');
  const main = all(doc, PKG, 'part').find((p) => p.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const body = main && all(main, W, 'body')[0];
  return body ? Array.from(body.childNodes).filter((n) => n.localName === 'p').map(xml) : [];
}

async function listed(context) {
  const changes = context.document.body.getTrackedChanges();
  changes.load('items/type,items/text');
  await context.sync();
  return { changes, list: changes.items.map((c, i) => ({ i, type: c.type, text: c.text })) };
}

async function check() {
  const report = { host: Office.context.diagnostics, scenarios: {} };
  let original;
  await Word.run(async (context) => {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    const ooxml = context.document.body.getOoxml();
    await context.sync();
    original = ooxml.value;
    report.paragraphs = await paragraphs(context);
    report.listed = (await listed(context)).list;
    const ps = context.document.body.paragraphs;
    ps.load('items/text');
    await context.sync();
    const per = ps.items.map((p) => { const c = p.getTrackedChanges(); c.load('items/type,items/text'); return c; });
    await context.sync();
    report.byParagraph = ps.items.map((p, i) => ({ text: p.text, changes: per[i].items.map((c) => `${c.type}: ${c.text}`) }));
  });
  for (const [name, index, act] of SCENARIOS) {
    const entry = (report.scenarios[name] = {});
    try {
      await Word.run(async (context) => {
        context.document.body.insertOoxml(original, 'Replace');
        await context.sync();
        const { changes } = await listed(context);
        const target = changes.items[index];
        if (!target) { entry.skipped = `no change ${index}`; return; }
        entry.acted = { type: target.type, text: target.text };
        if (act === 'accept') target.accept(); else target.reject();
        await context.sync();
        entry.listedAfter = (await listed(context)).list;
        entry.paragraphs = await paragraphs(context);
      });
    } catch (e) {
      entry.error = String(e && e.message || e);
    }
  }
  await Word.run(async (context) => { context.document.body.insertOoxml(original, 'Replace'); await context.sync(); });
  out.value = JSON.stringify(report, null, 2);
}
```

And the Script Lab snippet for 23 (Word, Office JS, WordApi 1.6), in a new blank document. The HTML
tab is check 15's (a **Run the check** button and a text box). The Script tab:

```js
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const D1 = '2026-09-28T01:00:00Z';
const D2 = '2026-09-28T02:00:00Z';
const A = (date) => `w:author="Author A" w:date="${date}"`;
const B = (date) => `w:author="Author B" w:date="${date}"`;
let nextId = 100;
const id = () => `w:id="${nextId++}"`;
const p = (inner, mark) => `<w:p>${mark ? `<w:pPr><w:rPr>${mark}</w:rPr></w:pPr>` : ''}${inner}</w:p>`;
const r = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const ins = (who, text) => `<w:ins ${id()} ${who}>${r(text)}</w:ins>`;
const del = (who, text) => `<w:del ${id()} ${who}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const insMark = (who) => `<w:ins ${id()} ${who}/>`;
const delMark = (who) => `<w:del ${id()} ${who}/>`;
const CASES = {
  'ins-paragraph': p(r('One.')) + p(ins(A(D1), 'Two.'), insMark(A(D1))) + p(r('Three.')),
  'ins-across': p(r('Four ') + ins(A(D1), 'five.'), insMark(A(D1))) + p(ins(A(D1), 'Six ') + r('seven.')),
  'ins-mark-alone': p(r('Eight.'), insMark(A(D1))) + p(r('Nine.')),
  'del-paragraph-dates': p(r('One.')) + p(del(A(D1), 'Two.'), delMark(A(D2))) + p(r('Three.')),
  'del-paragraph-authors': p(r('One.')) + p(del(A(D1), 'Two.'), delMark(B(D1))) + p(r('Three.')),
  'del-same': p(r('Ten ') + del(A(D1), 'eleven ') + del(A(D1), 'twelve') + r('.')),
  'del-dates': p(r('Ten ') + del(A(D1), 'eleven ') + del(A(D2), 'twelve') + r('.')),
  'del-authors': p(r('Ten ') + del(A(D1), 'eleven ') + del(B(D1), 'twelve') + r('.')),
  'del-then-ins': p(r('Ten ') + del(A(D1), 'eleven') + ins(A(D1), 'twelve') + r('.')),
};
const packageOf = (blocks) => `<pkg:package xmlns:pkg="${PKG}">`
  + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
  + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  + '</pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + `<pkg:xmlData><w:document xmlns:w="${W}"><w:body>${blocks}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;

const all = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
const xml = (node) => new XMLSerializer().serializeToString(node)
  .replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*)="[^"]*"/g, '');

async function paragraphs(context) {
  const ooxml = context.document.body.getOoxml();
  await context.sync();
  const doc = new DOMParser().parseFromString(ooxml.value, 'application/xml');
  const main = all(doc, PKG, 'part').find((p) => p.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const body = main && all(main, W, 'body')[0];
  return body ? Array.from(body.childNodes).filter((n) => n.localName === 'p').map(xml) : [];
}

async function listed(context) {
  const changes = context.document.body.getTrackedChanges();
  changes.load('items/type,items/text,items/author,items/date');
  await context.sync();
  return { changes, list: changes.items.map((c, i) => ({ i, type: c.type, author: c.author, date: c.date, text: c.text })) };
}

async function check() {
  const report = { host: Office.context.diagnostics, cases: {} };
  for (const [name, blocks] of Object.entries(CASES)) {
    const entry = (report.cases[name] = {});
    for (const act of ['accept', 'reject']) {
      try {
        await Word.run(async (context) => {
          context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
          context.document.body.clear();
          context.document.body.insertOoxml(packageOf(blocks), 'Start');
          await context.sync();
          const { changes, list } = await listed(context);
          if (act === 'accept') { entry.before = await paragraphs(context); entry.listed = list; }
          const target = changes.items[0];
          if (!target) { entry[act] = 'nothing listed'; return; }
          if (act === 'accept') target.accept(); else target.reject();
          await context.sync();
          entry[act] = { after: await paragraphs(context), listedAfter: (await listed(context)).list };
        });
      } catch (e) {
        entry[act] = { error: String(e && e.message || e) };
      }
    }
  }
  out.value = JSON.stringify(report, null, 2);
}
```

And the Script Lab snippet for 27 (Word, Office JS, WordApi 1.6), in a new blank document. The HTML
tab is check 15's (a **Run the check** button and a text box). The Script tab:

```js
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const all = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
const xml = (node) => new XMLSerializer().serializeToString(node)
  .replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*|w16du:dateUtc)="[^"]*"/g, '');
const REVISIONS = /<w:(\w+Change|cellIns|cellDel|cellMerge|ins|del)\b/g;

// --- part A: Word writes the revision (tracking on, one Office JS change on a fresh table) ---
const VALUES = [['a1', 'b1', 'c1'], ['a2', 'b2', 'c2'], ['a3', 'b3', 'c3']];
async function freshTable(context) {
  const body = context.document.body;
  context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
  await context.sync();
  body.clear();
  body.insertParagraph('Before.', 'Start');
  body.insertTable(3, 3, 'End', VALUES);
  body.insertParagraph('After.', 'End');
  await context.sync();
  context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
  const tables = body.tables;
  tables.load('items');
  await context.sync();
  return tables.items[0];
}
const WORD_WRITES = {
  'table-alignment': (c, t) => { t.alignment = 'Centered'; },
  'table-width': (c, t) => { t.width = 300; },
  'table-style': (c, t) => { t.styleBuiltIn = 'GridTable4_Accent1'; },
  'cell-shading': (c, t) => { t.getCell(0, 0).shadingColor = '#FFFF00'; },
  'cell-width': (c, t) => { t.getCell(0, 0).columnWidth = 150; },
  'cell-valign': (c, t) => { t.getCell(0, 0).verticalAlignment = 'Center'; },
  'row-height': (c, t) => { t.rows.getFirst().preferredHeight = 40; },
  'add-column': (c, t) => { t.addColumns('End', 1, [['d1'], ['d2'], ['d3']]); },
  'delete-column': (c, t) => { t.deleteColumns(1, 1); },
  'merge-vertical': (c, t) => { if (typeof t.mergeCells !== 'function') throw new Error('Table.mergeCells is not in this build'); t.mergeCells(0, 0, 1, 0); },
  'mark-bold': (c) => { c.document.body.paragraphs.getFirst().font.bold = true; },
  'list-level': async (c) => {
    const p = c.document.body.paragraphs.getLast();
    p.startNewList();
    await c.sync();
    p.listItem.level = 1;
  },
};

// --- part B: each kind written by hand, put in with tracking off ---
const A = 'w:author="Author A" w:date="2026-09-28T01:00:00Z"';
let nextId = 500;
const id = () => `w:id="${nextId++}"`;
const run = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const p = (s, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${run(s)}</w:p>`;
const tc = (s, tcPr = '') => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcPr}</w:tcPr>${p(s)}</w:tc>`;
const tr = (cells, trPr = '') => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells}</w:tr>`;
const row = (n, trPr = '') => tr(tc(`a${n}`) + tc(`b${n}`) + tc(`c${n}`), trPr);
const GRID = '<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/>';
const tbl = (rows, tblPr = '', gridExtra = '') => p('Before.')
  + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${tblPr}</w:tblPr><w:tblGrid>${GRID}${gridExtra}</w:tblGrid>${rows}</w:tbl>` + p('After.');
const MARGINS = (m) => `<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="${m}" w:right="${m}" w:bottom="${m}" w:left="${m}" w:header="708" w:footer="708" w:gutter="0"/>`;
const WRITTEN = {
  'tblPrChange': tbl(row(1) + row(2) + row(3), `<w:jc w:val="center"/><w:tblPrChange ${id()} ${A}><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange>`),
  'tblGridChange': tbl(row(1) + row(2) + row(3), '', `<w:tblGridChange ${id()}><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/><w:gridCol w:w="3000"/></w:tblGrid></w:tblGridChange>`),
  'trPrChange': tbl(row(1, `<w:trHeight w:val="600"/><w:trPrChange ${id()} ${A}><w:trPr/></w:trPrChange>`) + row(2) + row(3)),
  'tcPrChange': tbl(tr(tc('a1', `<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:tcPrChange ${id()} ${A}><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr></w:tcPrChange>`) + tc('b1') + tc('c1')) + row(2) + row(3)),
  'cellIns': tbl(tr(tc('a1') + tc('b1') + tc('c1 inserted', `<w:cellIns ${id()} ${A}/>`)) + row(2) + row(3)),
  'cellDel': tbl(tr(tc('a1') + tc('b1') + tc('c1 deleted', `<w:cellDel ${id()} ${A}/>`)) + row(2) + row(3)),
  'cellMerge': tbl(tr(tc('a1', '<w:vMerge w:val="restart"/>') + tc('b1') + tc('c1')) + tr(tc('', `<w:vMerge/><w:cellMerge ${id()} ${A} w:vMerge="cont"/>`) + tc('b2') + tc('c2')) + row(3)),
  'numberingChange': p('Before.') + p('A numbered item.', `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:numberingChange ${id()} ${A} w:original="%1."/></w:numPr>`) + p('After.'),
  'mark-rPrChange': p('Before.') + p('A paragraph whose mark was made bold.', `<w:rPr><w:b/><w:rPrChange ${id()} ${A}><w:rPr/></w:rPrChange></w:rPr>`) + p('After.'),
  'sectPrChange': p('The end of section one.', `<w:sectPr>${MARGINS(720)}<w:sectPrChange ${id()} ${A}><w:sectPr>${MARGINS(1440)}</w:sectPr></w:sectPrChange></w:sectPr>`) + p('Section two.'),
};
const NUMBERING = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>`
  + '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
  + '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
const RELS = 'application/vnd.openxmlformats-package.relationships+xml';
const packageOf = (blocks) => `<pkg:package xmlns:pkg="${PKG}">`
  + `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="${RELS}"><pkg:xmlData>`
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
  + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  + '</pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + `<pkg:xmlData><w:document xmlns:w="${W}"><w:body>${blocks}</w:body></w:document></pkg:xmlData></pkg:part>`
  + `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="${RELS}"><pkg:xmlData>`
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
  + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'
  + '</pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">'
  + `<pkg:xmlData>${NUMBERING}</pkg:xmlData></pkg:part></pkg:package>`;
async function freshWritten(context, blocks) {
  context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
  await context.sync();
  context.document.body.clear();
  context.document.body.insertOoxml(packageOf(blocks), 'Start');
  await context.sync();
}

// --- recording ---
async function blocks(context) {
  const ooxml = context.document.body.getOoxml();
  await context.sync();
  const doc = new DOMParser().parseFromString(ooxml.value, 'application/xml');
  const main = all(doc, PKG, 'part').find((part) => part.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const body = main && all(main, W, 'body')[0];
  return body ? Array.from(body.childNodes).filter((n) => n.nodeType === 1).map(xml) : [];
}
async function listed(context) {
  const changes = context.document.body.getTrackedChanges();
  changes.load('items/type,items/text,items/author');
  await context.sync();
  return { changes, list: changes.items.map((c, i) => ({ i, type: c.type, author: c.author, text: c.text.slice(0, 80) })) };
}
async function measure(make) {
  const entry = {};
  for (const act of ['before', 'accept', 'reject', 'acceptAll', 'rejectAll']) {
    try {
      await Word.run(async (context) => {
        await make(context);
        await context.sync();
        if (act === 'before') {
          entry.markup = await blocks(context);
          entry.revisions = [...entry.markup.join('').matchAll(REVISIONS)].map((m) => m[1]);
          entry.listed = (await listed(context)).list;
          return;
        }
        const { changes } = await listed(context);
        if (changes.items.length === 0) { entry[act] = 'nothing listed'; return; }
        if (act === 'accept') changes.items[0].accept();
        else if (act === 'reject') changes.items[0].reject();
        else if (act === 'acceptAll') changes.acceptAll();
        else changes.rejectAll();
        await context.sync();
        const after = await blocks(context);
        entry[act] = {
          after: after.map((b, i) => (entry.markup && b === entry.markup[i] ? '=' : b)),
          listedAfter: (await listed(context)).list,
        };
      });
    } catch (e) {
      entry[act] = { error: String(e && e.message || e) };
    }
  }
  return entry;
}

async function check() {
  const report = { host: Office.context.diagnostics, wordWrites: {}, written: {} };
  for (const [name, op] of Object.entries(WORD_WRITES)) {
    out.value = `running ${name}...`;
    report.wordWrites[name] = await measure(async (context) => { const table = await freshTable(context); await op(context, table); });
  }
  for (const [name, markup] of Object.entries(WRITTEN)) {
    out.value = `running ${name}...`;
    report.written[name] = await measure((context) => freshWritten(context, markup));
  }
  out.value = JSON.stringify(report, null, 2);
}
```

And the Script Lab snippet for 28 (Word, Office JS, WordApi 1.6), run in `28a-changes.docx`. The
HTML tab is check 15's (a **Run the check** button and a text box). The Script tab:

```js
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => Word.run(async (context) => {
  const changes = context.document.body.getTrackedChanges();
  changes.load('items/type,items/author,items/text');
  await context.sync();
  out.value = JSON.stringify({
    host: Office.context.diagnostics,
    listed: changes.items.map((c, i) => ({ i, type: c.type, author: c.author, text: c.text })),
  }, null, 2);
}).catch((e) => { out.value = String(e && e.stack || e); }));
```

And the Script Lab snippet for 30 (Word, Office JS, WordApi 1.6; `Range.insertField` is WordApi 1.5),
in a new blank document. The HTML tab is check 15's (a **Run the check** button and a text box). The
Script tab (also `__tmp/30/check30-script.js`):

```js
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const all = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
const xml = (node) => new XMLSerializer().serializeToString(node)
  .replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*|w16du:dateUtc)="[^"]*"/g, '');
const REVISIONS = /<w:(\w+Change|cellIns|cellDel|cellMerge|ins|del)\b/g;

// The inserted content: each case ends with a plain paragraph, "(end of insertion)", to take
// whatever joining Word does with the body's last paragraph.
const run = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const p = (...items) => `<w:p>${items.join('')}</w:p>`;
const END = p(run('(end of insertion)'));
const BORDERS = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
  + '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>';
const table = (inner, w) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${BORDERS}</w:tblPr><w:tblGrid><w:gridCol w:w="${w}"/></w:tblGrid>`
  + `<w:tr><w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`;
const FRAGMENTS = {
  'hyperlink': p(run('before '), `<w:hyperlink w:anchor="target">${run('link')}</w:hyperlink>`, run(' after')) + END,
  'field': p(run('page '), `<w:fldSimple w:instr=" PAGE ">${run('7')}</w:fldSimple>`, run(' after')) + END,
  'control-inline': p(run('before '), `<w:sdt><w:sdtPr><w:alias w:val="Inline"/><w:tag w:val="inline"/><w:id w:val="101"/></w:sdtPr><w:sdtContent>${run('control')}</w:sdtContent></w:sdt>`, run(' after')) + END,
  'smart-tag': p(run('in '), `<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="City">${run('Sydney')}</w:smartTag>`, run(' after')) + END,
  'custom-xml': p(run('in '), `<w:customXml w:uri="urn:example" w:element="thing">${run('custom')}</w:customXml>`, run(' after')) + END,
  'control-block': p(run('before the control')) + `<w:sdt><w:sdtPr><w:alias w:val="Block"/><w:tag w:val="block"/><w:id w:val="102"/></w:sdtPr><w:sdtContent>${p(run('in a block control'))}</w:sdtContent></w:sdt>` + END,
  'table': p(run('before the table')) + table(p(run('cell')), 4000) + END,
  'nested-table': table(table(p(run('nested')), 2000) + p(run('outer cell')), 4000) + END,
};

// Part C: what @docx4j/core-ts (f96fa6c) writes for each fragment, inserted after "Kept." under tracking
// by Author A (w14 ids and w16du:dateUtc taken out; the rest as written).
const ENGINE = {
  'hyperlink':
    '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="2" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">before </w:t></w:r></w:ins><w:hyperlink w:anchor="target"><w:ins w:id="3" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">link</w:t></w:r></w:ins></w:hyperlink><w:ins w:id="4" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:ins></w:p><w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="6" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">(end of insertion)</w:t></w:r></w:ins></w:p>',
  'field':
    '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="2" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">page </w:t></w:r></w:ins><w:fldSimple w:instr=" PAGE "><w:ins w:id="3" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">7</w:t></w:r></w:ins></w:fldSimple><w:ins w:id="4" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:ins></w:p><w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="6" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">(end of insertion)</w:t></w:r></w:ins></w:p>',
  'control-inline':
    '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="2" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">before </w:t></w:r></w:ins><w:sdt><w:sdtPr><w:alias w:val="Inline"/><w:tag w:val="inline"/><w:id w:val="101"/></w:sdtPr><w:sdtContent><w:ins w:id="3" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">control</w:t></w:r></w:ins></w:sdtContent></w:sdt><w:ins w:id="4" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:ins></w:p><w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="6" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">(end of insertion)</w:t></w:r></w:ins></w:p>',
  'smart-tag':
    '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="2" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">in </w:t></w:r></w:ins><w:smartTag w:element="City" w:uri="urn:schemas-microsoft-com:office:smarttags"><w:ins w:id="3" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">Sydney</w:t></w:r></w:ins></w:smartTag><w:ins w:id="4" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:ins></w:p><w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="6" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">(end of insertion)</w:t></w:r></w:ins></w:p>',
  'custom-xml':
    '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="2" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">in </w:t></w:r></w:ins><w:customXml w:element="thing" w:uri="urn:example"><w:ins w:id="3" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">custom</w:t></w:r></w:ins></w:customXml><w:ins w:id="4" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:ins></w:p><w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="6" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">(end of insertion)</w:t></w:r></w:ins></w:p>',
  'control-block':
    '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="2" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">before the control</w:t></w:r></w:ins></w:p><w:sdt><w:sdtPr><w:alias w:val="Block"/><w:tag w:val="block"/><w:id w:val="102"/></w:sdtPr><w:sdtContent><w:p><w:pPr><w:rPr><w:ins w:id="3" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="4" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">in a block control</w:t></w:r></w:ins></w:p></w:sdtContent></w:sdt><w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="6" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">(end of insertion)</w:t></w:r></w:ins></w:p>',
  'table':
    '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="2" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">before the table</w:t></w:r></w:ins></w:p><w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/><w:tblBorders><w:top w:color="auto" w:space="0" w:sz="4" w:val="single"/><w:left w:color="auto" w:space="0" w:sz="4" w:val="single"/><w:bottom w:color="auto" w:space="0" w:sz="4" w:val="single"/><w:right w:color="auto" w:space="0" w:sz="4" w:val="single"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:trPr><w:ins w:id="3" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:trPr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr><w:p><w:pPr><w:rPr><w:ins w:id="4" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="5" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">cell</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:rPr><w:ins w:id="6" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="7" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">(end of insertion)</w:t></w:r></w:ins></w:p>',
  'nested-table':
    '<w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/><w:tblBorders><w:top w:color="auto" w:space="0" w:sz="4" w:val="single"/><w:left w:color="auto" w:space="0" w:sz="4" w:val="single"/><w:bottom w:color="auto" w:space="0" w:sz="4" w:val="single"/><w:right w:color="auto" w:space="0" w:sz="4" w:val="single"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:trPr><w:ins w:id="1" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:trPr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr><w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/><w:tblBorders><w:top w:color="auto" w:space="0" w:sz="4" w:val="single"/><w:left w:color="auto" w:space="0" w:sz="4" w:val="single"/><w:bottom w:color="auto" w:space="0" w:sz="4" w:val="single"/><w:right w:color="auto" w:space="0" w:sz="4" w:val="single"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:ins w:id="2" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:trPr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr><w:p><w:pPr><w:rPr><w:ins w:id="3" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="4" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">nested</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="6" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">outer cell</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:rPr><w:ins w:id="7" w:author="Author A" w:date="2026-09-29T11:00:00Z"/></w:rPr></w:pPr><w:ins w:id="8" w:author="Author A" w:date="2026-09-29T11:00:00Z"><w:r><w:t xml:space="preserve">(end of insertion)</w:t></w:r></w:ins></w:p>',
};

const RELS = 'application/vnd.openxmlformats-package.relationships+xml';
const packageOf = (blocks) => `<pkg:package xmlns:pkg="${PKG}">`
  + `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="${RELS}"><pkg:xmlData>`
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
  + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  + '</pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + `<pkg:xmlData><w:document xmlns:w="${W}"><w:body>${blocks}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;

// Every case starts from "Kept." and the body's last, empty paragraph, put in untracked.
async function fresh(context) {
  const body = context.document.body;
  context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
  await context.sync();
  body.clear();
  body.insertParagraph('Kept.', 'Start');
  await context.sync();
  const first = body.paragraphs.getFirst();
  first.load('text');
  await context.sync();
  return first;
}
const track = async (context, on) => {
  context.document.changeTrackingMode = on ? Word.ChangeTrackingMode.trackAll : Word.ChangeTrackingMode.off;
  await context.sync();
};

// Part B: Word's own gestures on content controls, links and fields, with tracking on.
const GESTURES = {
  'new-control-typed': async (c, first) => {
    const cc = first.getRange('End').insertContentControl();
    await c.sync();
    cc.insertText('typed in a new control', 'Replace');
  },
  'inserted-text-wrapped': async (c, first) => {
    const r = first.insertText(' inserted words', 'End');
    await c.sync();
    r.insertContentControl();
  },
  'existing-control-typed': async (c, first, entry) => {
    await track(c, false);
    first.getRange('End').insertContentControl();
    await c.sync();
    entry.setup = await blocks(c);
    const cc = first.contentControls.getFirst();
    await track(c, true);
    cc.insertText('typed into an existing control', 'Replace');
  },
  'field-inserted': async (c, first) => {
    const at = first.getRange('End');
    if (typeof at.insertField !== 'function') throw new Error('Range.insertField is not in this build (WordApi 1.5)');
    at.insertField('End', Word.FieldType.page);
  },
  'link-on-inserted-text': async (c, first) => {
    const r = first.insertText(' a link', 'End');
    await c.sync();
    r.hyperlink = 'https://example.com/';
  },
};

// --- recording ---
// blocks.last says what the read saw, for a run whose markup comes back empty.
async function blocks(context) {
  const ooxml = context.document.body.getOoxml();
  await context.sync();
  const s = ooxml.value;
  const doc = new DOMParser().parseFromString(s, 'application/xml');
  const parts = all(doc, PKG, 'part');
  const main = parts.find((part) => part.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const body = main && all(main, W, 'body')[0];
  const err = doc.getElementsByTagName('parsererror')[0];
  blocks.last = { ooxmlLength: s.length, head: s.slice(0, 160), parseError: err ? err.textContent.slice(0, 200) : null,
    parts: parts.map((part) => part.getAttributeNS(PKG, 'name')), bodyFound: !!body };
  return body ? Array.from(body.childNodes).filter((n) => n.nodeType === 1 && n.localName !== 'sectPr').map(xml) : [];
}
const withDiagnostics = (markup) => (markup.length ? { markup } : { markup, markupRead: blocks.last });
async function listed(context) {
  const changes = context.document.body.getTrackedChanges();
  changes.load('items/type,items/text,items/author');
  await context.sync();
  return { changes, list: changes.items.map((c, i) => ({ i, type: c.type, author: c.author, text: c.text.slice(0, 80) })) };
}
async function measure(make) {
  const entry = {};
  for (const act of ['before', 'rejectAll', 'acceptAll']) {
    try {
      await Word.run(async (context) => {
        const first = await fresh(context);
        await make(context, first, act === 'before' ? entry : {});
        await context.sync();
        if (act === 'before') {
          Object.assign(entry, withDiagnostics(await blocks(context)));
          entry.revisions = [...entry.markup.join('').matchAll(REVISIONS)].map((m) => m[1]);
          entry.listed = (await listed(context)).list;
          return;
        }
        const { changes } = await listed(context);
        if (changes.items.length === 0) { entry[act] = 'nothing listed'; return; }
        if (act === 'acceptAll') changes.acceptAll(); else changes.rejectAll();
        await context.sync();
        const body = context.document.body;
        body.load('text');
        await context.sync();
        entry[act] = { ...withDiagnostics(await blocks(context)), text: body.text, listedAfter: (await listed(context)).list };
      });
    } catch (e) {
      entry[act] = { error: String(e && e.message || e) };
    }
  }
  return entry;
}

// The probe: the hyperlink fragment inserted at the end with tracking on, and what the read saw of it. A run
// whose probe shows a paragraph mark per XML tag, or no blocks, is Word or the snippet's runtime, not the cases.
async function probe() {
  const sent = packageOf(FRAGMENTS.hyperlink);
  return Word.run(async (context) => {
    await fresh(context);
    await track(context, true);
    context.document.body.insertOoxml(sent, 'End');
    await context.sync();
    const markup = await blocks(context);
    const body = context.document.body;
    body.load('text');
    await context.sync();
    return { sentLength: sent.length, sentHead: sent.slice(0, 120), text: body.text, ...blocks.last, markup };
  });
}

async function check() {
  const report = { host: Office.context.diagnostics, probe: null, wordInserts: {}, gestures: {}, engine: {} };
  out.value = 'probe...';
  try { report.probe = await probe(); } catch (e) { report.probe = { error: String(e && e.message || e) }; }
  // Part A: the fragment put in with tracking ON - the markup Word writes for it, and what rejecting leaves
  for (const [name, fragment] of Object.entries(FRAGMENTS)) {
    out.value = `A ${name}...`;
    report.wordInserts[name] = await measure(async (c) => { await track(c, true); c.document.body.insertOoxml(packageOf(fragment), 'End'); });
  }
  // Part B: Word's own gestures, tracking on
  for (const [name, gesture] of Object.entries(GESTURES)) {
    out.value = `B ${name}...`;
    report.gestures[name] = await measure(async (c, first, entry) => { await track(c, true); await gesture(c, first, entry); });
  }
  // Part C: the engine's markup for the same fragments (Author A), put in with tracking OFF
  for (const [name, markup] of Object.entries(ENGINE)) {
    out.value = `C ${name}...`;
    report.engine[name] = await measure(async (c) => { c.document.body.insertOoxml(packageOf(markup), 'End'); });
  }
  out.value = JSON.stringify(report, null, 2);
}
```

And the Script Lab snippet for 31 (Word, Office JS, WordApi 1.6), in a new blank document. The HTML
tab is check 30's. The Script tab:

```js
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const all = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
const xml = (node) => new XMLSerializer().serializeToString(node)
  .replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*|w16du:dateUtc)="[^"]*"/g, '');

// Check 31: how Office JS lists an inserted table of several rows, with and without insertions
// touching it, and rows added to a table that was there (CR-002 section 37 item 7's open question).
const run = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const p = (...items) => `<w:p>${items.join('')}</w:p>`;
const row = (text) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(run(text))}</w:tc></w:tr>`;
const BORDERS = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
  + '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>';
const table = (...rows) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${BORDERS}</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${rows.join('')}</w:tbl>`;
const THREE = table(row('row one'), row('row two'), row('row three'));
const END = p(run('(end of insertion)'));
const FRAGMENTS = {
  'table-between': p(run('before the table')) + THREE + END,   // check 30's table case with three rows
  'table-alone': THREE,                                         // nothing of the insertion touching it
  'two-tables': table(row('a1'), row('a2')) + p(run('between')) + table(row('b1'), row('b2')) + END,
};

// Part C: the engine's markup (@docx4j/core-ts 426e7b5) for 'table-between', Author A; and variants.
const A = (id) => `w:id="${id}" w:author="Author A" w:date="2026-10-01T11:00:00Z"`;
const B = (id) => `w:id="${id}" w:author="Author B" w:date="2026-10-01T11:00:00Z"`;
const insP = (who, markId, runId, text) => `<w:p><w:pPr><w:rPr><w:ins ${who(markId)}/></w:rPr></w:pPr><w:ins ${who(runId)}>${run(text)}</w:ins></w:p>`;
const insRow = (who, rowId, markId, runId, text) => `<w:tr><w:trPr><w:ins ${who(rowId)}/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${insP(who, markId, runId, text)}</w:tc></w:tr>`;
const ENGINE = {
  'engine-table-between': insP(A, 1, 2, 'before the table') + table(insRow(A, 3, 4, 5, 'row one'), insRow(A, 6, 7, 8, 'row two'), insRow(A, 9, 10, 11, 'row three')) + insP(A, 12, 13, '(end of insertion)'),
  'engine-two-authors': insP(A, 1, 2, 'before the table') + table(insRow(A, 3, 4, 5, 'row one'), insRow(B, 6, 7, 8, 'row two'), insRow(A, 9, 10, 11, 'row three')) + insP(A, 12, 13, '(end of insertion)'),
  'engine-rows-added': p(run('Plain paragraph.')) + table(row('row one'), insRow(A, 1, 2, 3, 'row two'), insRow(A, 4, 5, 6, 'row three')) + p(run('Plain after.')),
  'engine-rows-apart': p(run('Plain paragraph.')) + table(insRow(A, 1, 2, 3, 'new first'), row('row one'), row('row two'), insRow(A, 4, 5, 6, 'new last')) + p(run('Plain after.')),
};

const RELS = 'application/vnd.openxmlformats-package.relationships+xml';
const packageOf = (blocks) => `<pkg:package xmlns:pkg="${PKG}">`
  + `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="${RELS}"><pkg:xmlData>`
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
  + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  + '</pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + `<pkg:xmlData><w:document xmlns:w="${W}"><w:body>${blocks}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;

async function fresh(context) {
  const body = context.document.body;
  context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
  await context.sync();
  body.clear();
  body.insertParagraph('Kept.', 'Start');
  await context.sync();
}
const track = async (context, on) => {
  context.document.changeTrackingMode = on ? Word.ChangeTrackingMode.trackAll : Word.ChangeTrackingMode.off;
  await context.sync();
};
// Part B's table that was there: "Kept.", then the three-row table, untracked
async function existingTable(context) {
  await fresh(context);
  context.document.body.insertOoxml(packageOf(THREE), 'End');
  await context.sync();
  const table = context.document.body.tables.getFirst();
  table.load('rowCount');
  await context.sync();
  return table;
}

// Part B: Office JS's own row insertions into a table that was there, tracking on
const GESTURES = {
  'addRows-end-two': async (c, table) => { table.addRows('End', 2, [['new one'], ['new two']]); },
  'addRows-start-two': async (c, table) => { table.addRows('Start', 2, [['new one'], ['new two']]); },
  'rows-apart': async (c, table) => {
    table.rows.getFirst().insertRows('After', 1, [['new after first']]);
    await c.sync();
    table.rows.getFirst().getNext().getNext().getNext().insertRows('After', 1, [['new after third']]);
  },
  'text-then-rows': async (c, table) => {
    // an inserted paragraph touching the table, and rows inserted at its start: one change or two?
    table.insertParagraph('typed before the table', 'Before');
    await c.sync();
    table.addRows('Start', 1, [['new first']]);
  },
  'rows-then-text': async (c, table) => {
    table.addRows('End', 1, [['new last']]);
    await c.sync();
    table.insertParagraph('typed after the table', 'After');
  },
};

// --- recording (check 30's) ---
async function blocks(context) {
  const ooxml = context.document.body.getOoxml();
  await context.sync();
  const s = ooxml.value;
  const doc = new DOMParser().parseFromString(s, 'application/xml');
  const parts = all(doc, PKG, 'part');
  const main = parts.find((part) => part.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const body = main && all(main, W, 'body')[0];
  const err = doc.getElementsByTagName('parsererror')[0];
  blocks.last = { ooxmlLength: s.length, head: s.slice(0, 160), parseError: err ? err.textContent.slice(0, 200) : null,
    parts: parts.map((part) => part.getAttributeNS(PKG, 'name')), bodyFound: !!body };
  return body ? Array.from(body.childNodes).filter((n) => n.nodeType === 1 && n.localName !== 'sectPr').map(xml) : [];
}
const withDiagnostics = (markup) => (markup.length ? { markup } : { markup, markupRead: blocks.last });
async function listed(context) {
  const changes = context.document.body.getTrackedChanges();
  changes.load('items/type,items/text,items/author');
  await context.sync();
  return { changes, list: changes.items.map((c, i) => ({ i, type: c.type, author: c.author, text: c.text.slice(0, 120) })) };
}
async function measure(make) {
  const entry = {};
  for (const act of ['before', 'rejectAll', 'acceptAll']) {
    try {
      await Word.run(async (context) => {
        await make(context);
        await context.sync();
        if (act === 'before') {
          Object.assign(entry, withDiagnostics(await blocks(context)));
          entry.listed = (await listed(context)).list;
          return;
        }
        const { changes } = await listed(context);
        if (changes.items.length === 0) { entry[act] = 'nothing listed'; return; }
        if (act === 'acceptAll') changes.acceptAll(); else changes.rejectAll();
        await context.sync();
        const body = context.document.body;
        body.load('text');
        await context.sync();
        entry[act] = { ...withDiagnostics(await blocks(context)), text: body.text, listedAfter: (await listed(context)).list };
      });
    } catch (e) {
      entry[act] = { error: String(e && e.message || e), debug: e && e.debugInfo ? JSON.stringify(e.debugInfo) : undefined };
    }
  }
  return entry;
}

async function check() {
  const report = { host: Office.context.diagnostics, wordInserts: {}, gestures: {}, engine: {} };
  for (const [name, fragment] of Object.entries(FRAGMENTS)) {
    out.value = `A ${name}...`;
    report.wordInserts[name] = await measure(async (c) => { await fresh(c); await track(c, true); c.document.body.insertOoxml(packageOf(fragment), 'End'); });
  }
  for (const [name, gesture] of Object.entries(GESTURES)) {
    out.value = `B ${name}...`;
    report.gestures[name] = await measure(async (c) => { const table = await existingTable(c); await track(c, true); await gesture(c, table); });
  }
  for (const [name, markup] of Object.entries(ENGINE)) {
    out.value = `C ${name}...`;
    report.engine[name] = await measure(async (c) => { await fresh(c); c.document.body.insertOoxml(packageOf(markup), 'End'); });
  }
  out.value = JSON.stringify(report, null, 2);
}
```

And the Script Lab snippet for 32 (Word, Office JS, WordApi 1.6; `body.getComments` is WordApi 1.4), in a
new blank document. The HTML tab:

```html
<button id="run">Run A and B-api</button>
<button id="setup">Set up typing (B)</button>
<button id="record">Record typing (B)</button>
<p>Copy this back:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const out = document.getElementById('out');
const guard = (fn) => () => fn().catch((e) => { out.value = String(e.stack || e) + (e && e.debugInfo ? '\n' + JSON.stringify(e.debugInfo) : ''); });
document.getElementById('run').addEventListener('click', guard(runAutomated));
document.getElementById('setup').addEventListener('click', guard(setupTyping));
document.getElementById('record').addEventListener('click', guard(recordTyping));

const all = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
const xml = (node) => new XMLSerializer().serializeToString(node)
  .replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*|w16du:dateUtc)="[^"]*"/g, '');

// Check 32. Case A: a pending insertion (or deletion) holding a comment's word, rejected (accepted);
// case B: typing into a link that was there, with tracking on.
const REF = (id) => `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
const CASES_A = {
  // A1: Word's own form (the editor's measure-word15.docx): the end marker and the reference after the w:ins
  'A1-word-form': { word: 'monthly', comment: 2, act: 'reject', blocks:
    `<w:p><w:pPr><w:rPr><w:ins w:id="10" w:author="Claude" w:date="2026-09-30T13:34:00Z"/></w:rPr></w:pPr><w:ins w:id="11" w:author="Claude" w:date="2026-09-30T13:34:00Z"><w:r><w:t xml:space="preserve">The tenant pays rent </w:t></w:r><w:commentRangeStart w:id="2"/><w:r><w:t>monthly</w:t></w:r></w:ins><w:commentRangeEnd w:id="2"/>${REF(2)}<w:ins w:id="13" w:author="Claude" w:date="2026-09-30T13:34:00Z"><w:r><w:t xml:space="preserve"> in advance.</w:t></w:r></w:ins></w:p>` },
  // A2: the editor's form: the end marker just inside the first w:ins
  'A2-editor-form': { word: 'beta', comment: 0, act: 'reject', blocks:
    `<w:p><w:ins w:id="11" w:author="A" w:date="2026-10-01T09:45:00Z"><w:r><w:t xml:space="preserve">alpha </w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>beta</w:t></w:r><w:commentRangeEnd w:id="0"/></w:ins>${REF(0)}<w:ins w:id="12" w:author="A" w:date="2026-10-01T09:45:00Z"><w:r><w:t xml:space="preserve"> gamma </w:t></w:r></w:ins><w:r><w:t>Some text.</w:t></w:r></w:p>` },
  // A3: a deletion in A1's shape, accepted
  'A3-deletion': { word: 'monthly', comment: 2, act: 'accept', blocks:
    `<w:p><w:del w:id="11" w:author="Claude" w:date="2026-09-30T13:34:00Z"><w:r><w:delText xml:space="preserve">The tenant pays rent </w:delText></w:r><w:commentRangeStart w:id="2"/><w:r><w:delText>monthly</w:delText></w:r></w:del><w:commentRangeEnd w:id="2"/>${REF(2)}<w:del w:id="13" w:author="Claude" w:date="2026-09-30T13:34:00Z"><w:r><w:delText xml:space="preserve"> in advance.</w:delText></w:r></w:del><w:r><w:t>Kept after.</w:t></w:r></w:p>` },
};
const COMMENTS = {
  2: 'On monthly.',
  0: 'On beta.',
};
// Case B's paragraph: an untracked external link on "lease"
const LINK_PARAGRAPH = '<w:p><w:r><w:t xml:space="preserve">Read the </w:t></w:r><w:hyperlink r:id="rIdLink" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>lease</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> today.</w:t></w:r></w:p>';

const RELS = 'application/vnd.openxmlformats-package.relationships+xml';
const commentsPart = (ids) => `<pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"><pkg:xmlData>`
  + `<w:comments xmlns:w="${W}">` + ids.map((id) => `<w:comment w:id="${id}" w:author="Reviewer" w:initials="R" w:date="2026-10-02T09:00:00Z"><w:p><w:r><w:t>${COMMENTS[id]}</w:t></w:r></w:p></w:comment>`).join('') + '</w:comments></pkg:xmlData></pkg:part>';
const packageOf = (blocks, { commentIds = [], link = false } = {}) => `<pkg:package xmlns:pkg="${PKG}">`
  + `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="${RELS}"><pkg:xmlData>`
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
  + `Type="${R}/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>`
  + `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="${RELS}"><pkg:xmlData>`
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + (commentIds.length ? `<Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/>` : '')
  + (link ? `<Relationship Id="rIdLink" Type="${R}/hyperlink" Target="https://example.com/lease" TargetMode="External"/>` : '')
  + '</Relationships></pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + `<pkg:xmlData><w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${blocks}</w:body></w:document></pkg:xmlData></pkg:part>`
  + (commentIds.length ? commentsPart(commentIds) : '') + '</pkg:package>';

async function fresh(context) {
  const body = context.document.body;
  context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
  await context.sync();
  body.clear();
  body.insertParagraph('Kept.', 'Start');
  await context.sync();
}
const track = async (context, on) => {
  context.document.changeTrackingMode = on ? Word.ChangeTrackingMode.trackAll : Word.ChangeTrackingMode.off;
  await context.sync();
};

// --- recording (checks 30 and 31's, plus the comments) ---
async function blocks(context) {
  const ooxml = context.document.body.getOoxml();
  await context.sync();
  const s = ooxml.value;
  const doc = new DOMParser().parseFromString(s, 'application/xml');
  const parts = all(doc, PKG, 'part');
  const main = parts.find((part) => part.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const body = main && all(main, W, 'body')[0];
  const comments = parts.find((part) => part.getAttributeNS(PKG, 'name') === '/word/comments.xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  blocks.last = { ooxmlLength: s.length, parseError: err ? err.textContent.slice(0, 200) : null, parts: parts.map((part) => part.getAttributeNS(PKG, 'name')), bodyFound: !!body };
  blocks.commentsPart = comments ? all(comments, W, 'comment').map((c) => ({ id: c.getAttributeNS(W, 'id'), text: c.textContent })) : null;
  return body ? Array.from(body.childNodes).filter((n) => n.nodeType === 1 && n.localName !== 'sectPr').map(xml) : [];
}
const withDiagnostics = (markup) => (markup.length ? { markup } : { markup, markupRead: blocks.last });
async function listed(context) {
  const changes = context.document.body.getTrackedChanges();
  changes.load('items/type,items/text,items/author');
  await context.sync();
  return { changes, list: changes.items.map((c, i) => ({ i, type: c.type, author: c.author, text: c.text.slice(0, 120) })) };
}
async function comments(context) {
  try {
    const list = context.document.body.getComments();
    list.load('items/content,items/authorName,items/resolved');
    await context.sync();
    const ranges = list.items.map((c) => { const r = c.getRange(); r.load('text'); return r; });
    await context.sync();
    return list.items.map((c, i) => ({ content: c.content, author: c.authorName, anchoredText: ranges[i].text }));
  } catch (e) {
    return { error: String(e && e.message || e) };
  }
}
async function snapshot(context) {
  const markup = await blocks(context);
  const body = context.document.body;
  body.load('text');
  await context.sync();
  return { ...withDiagnostics(markup), text: body.text, listed: (await listed(context)).list, comments: await comments(context), commentsPart: blocks.commentsPart };
}

// Case A: each form put in untracked; then that one change, reject all, accept all, each from a fresh start
async function caseA(name, spec) {
  const entry = {};
  const pkg = packageOf(spec.blocks, { commentIds: [spec.comment] });
  const put = async (context) => { await fresh(context); context.document.body.insertOoxml(pkg, 'End'); await context.sync(); };
  for (const act of ['before', 'thatOne', 'rejectAll', 'acceptAll']) {
    try {
      await Word.run(async (context) => {
        await put(context);
        if (act === 'before') { entry.before = await snapshot(context); return; }
        const { changes } = await listed(context);
        if (changes.items.length === 0) { entry[act] = 'nothing listed'; return; }
        if (act === 'thatOne') {
          const target = changes.items.find((c) => c.text.includes(spec.word));
          if (!target) { entry.thatOne = 'no change holds the word'; return; }
          entry.thatOneWas = { type: target.type, text: target.text, action: spec.act };
          if (spec.act === 'accept') target.accept(); else target.reject();
        } else if (act === 'acceptAll') changes.acceptAll(); else changes.rejectAll();
        await context.sync();
        entry[act] = await snapshot(context);
      });
    } catch (e) {
      entry[act] = { error: String(e && e.message || e), debug: e && e.debugInfo ? JSON.stringify(e.debugInfo) : undefined };
    }
  }
  return entry;
}

// Case B through the API: insertText at the three places of the typing, tracking on
async function caseBApi() {
  const entry = {};
  try {
    await Word.run(async (context) => {
      await fresh(context);
      context.document.body.insertOoxml(packageOf(LINK_PARAGRAPH, { link: true }), 'End');
      await context.sync();
      entry.setup = await snapshot(context);
      await track(context, true);
      const p = context.document.body.paragraphs.getFirst().getNext();
      // "x" in the middle of "lease": after "le"
      p.search('le', { matchCase: true }).getFirst().insertText('x', 'After');
      await context.sync();
      entry.afterX = await snapshot(context);
      // "y" at the link's end: after "lexase"
      p.search('lexase', { matchCase: true }).getFirst().insertText('y', 'After');
      await context.sync();
      entry.afterY = await snapshot(context);
      // "z" at its start: before "lexase"
      p.search('lexase', { matchCase: true }).getFirst().insertText('z', 'Before');
      await context.sync();
      entry.afterZ = await snapshot(context);
    });
  } catch (e) {
    entry.error = { error: String(e && e.message || e), debug: e && e.debugInfo ? JSON.stringify(e.debugInfo) : undefined };
  }
  return entry;
}

async function runAutomated() {
  const report = { host: Office.context.diagnostics, caseA: {}, caseBApi: null };
  for (const [name, spec] of Object.entries(CASES_A)) {
    out.value = `A ${name}...`;
    report.caseA[name] = await caseA(name, spec);
  }
  out.value = 'B api...';
  report.caseBApi = await caseBApi();
  out.value = JSON.stringify(report, null, 2);
}

// Case B by hand: the link paragraph, tracking on; type x, y, z; then Record.
async function setupTyping() {
  await Word.run(async (context) => {
    await fresh(context);
    context.document.body.insertOoxml(packageOf(LINK_PARAGRAPH, { link: true }), 'End');
    await context.sync();
    await track(context, true);
  });
  out.value = 'Ready: "Read the lease today." with "lease" a link, tracking ON (user name as set in File > Options).\n'
    + '1. Click between "le" and "ase" in the link and type x.\n'
    + '2. Click right after "lexase" (before the space) and type y.\n'
    + '3. Click right before the "l" of "lexase" (after the space) and type z.\n'
    + 'Then click "Record typing (B)" and copy the JSON back; then File > Save As check32-typed.docx into the shared __tmp/32/.';
}
async function recordTyping() {
  const report = { host: Office.context.diagnostics, typed: null };
  await Word.run(async (context) => { report.typed = await snapshot(context); });
  out.value = JSON.stringify(report, null, 2);
}
```

And the Script Lab snippet for 34's part 2 probe (Word, Office JS; `customXmlParts` is WordApi 1.4),
in a new blank document. The HTML tab is check 30's. The Script tab:

```js
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e) + (e && e.debugInfo ? '\n' + JSON.stringify(e.debugInfo) : ''); }));

const all = (root, ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
const xml = (node) => new XMLSerializer().serializeToString(node).replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*)="[^"]*"/g, '');

// Check 34, part 2: can Office JS map a content control to a custom XML node at all, and which side
// wins when it does? Probes each API before using it and records what is missing.
const DATA = '<data><name>Ann</name><empty/></data>';
const RELS = 'application/vnd.openxmlformats-package.relationships+xml';
const PARAGRAPH = '<w:p><w:r><w:t xml:space="preserve">The quick </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>brown</w:t></w:r><w:r><w:t xml:space="preserve"> fox.</w:t></w:r></w:p>';
const packageOf = (blocks) => `<pkg:package xmlns:pkg="${PKG}"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="${RELS}"><pkg:xmlData>`
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  + '</pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + `<pkg:xmlData><w:document xmlns:w="${W}"><w:body>${blocks}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;

async function markup(context) {
  const ooxml = context.document.body.getOoxml();
  await context.sync();
  const doc = new DOMParser().parseFromString(ooxml.value, 'application/xml');
  const parts = all(doc, PKG, 'part');
  const main = parts.find((p) => p.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const body = main && all(main, W, 'body')[0];
  const custom = parts.filter((p) => /customXml\/item\d+\.xml$/.test(p.getAttributeNS(PKG, 'name'))).map((p) => ({ name: p.getAttributeNS(PKG, 'name'), xml: new XMLSerializer().serializeToString(p).slice(0, 400) }));
  return { blocks: body ? Array.from(body.childNodes).filter((n) => n.nodeType === 1 && n.localName !== 'sectPr').map(xml) : [], customXmlParts: custom };
}
async function partXml(context, id) {
  try {
    const part = context.document.customXmlParts.getItem(id);
    const x = part.getXml();
    await context.sync();
    return x.value;
  } catch (e) { return { error: String(e && e.message || e) }; }
}
async function fresh(context) {
  context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
  context.document.body.clear();
  await context.sync();
  context.document.body.insertOoxml(packageOf(PARAGRAPH), 'Start');
  await context.sync();
}
async function addPart(context) {
  const parts = context.document.customXmlParts;
  parts.load('items/id');
  await context.sync();
  const before = parts.items.map((p) => p.id);
  const part = parts.add(DATA);
  part.load('id,namespaceUri');
  await context.sync();
  return { id: part.id, namespaceUri: part.namespaceUri, partsBefore: before.length };
}
const describe = (cc) => {
  const keys = [];
  for (const k of ['xmlMapping', 'type', 'subtype', 'placeholderText', 'tag', 'title']) keys.push(k + ':' + typeof cc[k]);
  return keys.join(' ');
};
async function mapCase(context, name, make) {
  const entry = { name };
  try {
    await fresh(context);
    entry.part = await addPart(context);
    const cc = await make(context);
    cc.load('text,type,subtype,tag,title');
    await context.sync();
    entry.controlBefore = { text: cc.text, type: cc.type, subtype: cc.subtype };
    entry.api = describe(cc);
    if (typeof cc.xmlMapping !== 'object' && typeof cc.xmlMapping !== 'function') {
      entry.mapping = 'ContentControl.xmlMapping is not in this Office JS build';
    } else {
      const m = cc.xmlMapping;
      entry.mappingApi = Object.getOwnPropertyNames(Object.getPrototypeOf(m)).join(',');
      const r = m.setMapping('/data/name');
      await context.sync();
      entry.setMapping = r && typeof r.value !== 'undefined' ? r.value : String(r);
      cc.load('text');
      await context.sync();
      entry.controlAfter = cc.text;
    }
    entry.markupAfter = await markup(context);
    entry.partXmlAfter = await partXml(context, entry.part.id);
  } catch (e) {
    entry.error = { error: String(e && e.message || e), debug: e && e.debugInfo ? JSON.stringify(e.debugInfo) : undefined };
  }
  return entry;
}
async function check() {
  const report = { host: Office.context.diagnostics, cases: [] };
  await Word.run(async (context) => {
    out.value = '2a...';
    report.cases.push(await mapCase(context, '2a-selection', async (c) => {
      const r = c.document.body.paragraphs.getFirst().search('quick brown').getFirst();
      let cc;
      try { cc = r.insertContentControl('PlainText'); await c.sync(); } catch (e) { report.plainTextError = String(e && e.message || e); cc = r.insertContentControl(); await c.sync(); }
      return cc;
    }));
    out.value = '2b caret...';
    report.cases.push(await mapCase(context, '2b-caret', async (c) => {
      const r = c.document.body.paragraphs.getFirst().search('quick').getFirst().getRange('End');
      let cc;
      try { cc = r.insertContentControl('PlainText'); await c.sync(); } catch (e) { cc = r.insertContentControl(); await c.sync(); }
      return cc;
    }));
    out.value = '2b Ann...';
    report.cases.push(await mapCase(context, '2b-already-Ann', async (c) => {
      const r = c.document.body.paragraphs.getFirst().search('quick brown').getFirst();
      let cc;
      try { cc = r.insertContentControl('PlainText'); await c.sync(); } catch (e) { cc = r.insertContentControl(); await c.sync(); }
      cc.insertText('Ann', 'Replace');
      await c.sync();
      return cc;
    }));
  });
  out.value = JSON.stringify(report, null, 2);
}
```

And the Script Lab snippet for 42 (Word, Office JS, WordApi 1.3), on a copy of
`fixtures/parity/numbering-stories.docx`. The HTML tab:

```html
<button id="run">Run the check</button>
<p>Copy this back:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

async function check() {
  await Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    const pictures = body.inlinePictures;
    paragraphs.load('items/text');
    pictures.load('items/width,items/height');
    body.load('text');
    await context.sync();
    out.value = JSON.stringify({
      host: Office.context.diagnostics,
      paragraphs: paragraphs.items.length,
      texts: paragraphs.items.map((p) => p.text),
      bodyTextHasBox: /P2[0-2]/.test(body.text),
      inlinePictures: pictures.items.map((p) => [p.width, p.height]),
    }, null, 2);
  });
}
```

And the Script Lab snippet for 43 (Word, Office JS, WordApi 1.3), on `__tmp/43/loadAndSave.docx`. The
HTML tab:

```html
<button id="run">Run the check</button>
<p>Copy this back:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

async function check() {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    const pictures = context.document.body.inlinePictures;
    paragraphs.load('items');
    pictures.load('items/width,items/height,items/altTextTitle,items/altTextDescription');
    await context.sync();
    // each picture's paragraph, as its 1-based position among the body's paragraphs
    const positions = pictures.items.map((p) => {
      const range = p.paragraph.getRange('Whole');
      return paragraphs.items.map((q) => range.compareLocationWith(q.getRange('Whole')));
    });
    await context.sync();
    const report = {
      host: Office.context.diagnostics,
      paragraphs: paragraphs.items.length,
      inlinePictures: pictures.items.map((p, i) => ({
        width: p.width,
        height: p.height,
        altTextTitle: p.altTextTitle,
        altTextDescription: p.altTextDescription,
        paragraph: positions[i].findIndex((c) => c.value === 'Equal') + 1,
      })),
    };
    // the image Office JS gives for each, on its own, in case a chart has none
    try {
      const images = pictures.items.map((p) => p.getBase64ImageSrc());
      await context.sync();
      images.forEach((image, i) => { report.inlinePictures[i].base64Length = image.value.length; });
    } catch (e) {
      report.base64Error = String(e.message || e);
    }
    out.value = JSON.stringify(report, null, 2);
  });
}
```

And the Script Lab snippet for 44 (Word, Office JS, WordApi 1.1, and `imageFormat` where this Word has
it), on `__tmp/44/loadAndSave.docx`. The HTML tab:

```html
<button id="run">Run the check</button>
<p>Copy this back:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

// word/media/image2.svg's bytes as base64 (760 bytes, 1016 characters)
const SVG = 'PHN2ZyB2aWV3Qm94PSIwIDAgOTYgOTYiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiIGlkPSJJY29uc19CYWRnZVRpY2sxIiBvdmVyZmxvdz0iaGlkZGVuIj48cGF0aCBkPSJNNDcuOTc5IDEwQzI3LjAwMzggMTAgMTAgMjcuMDAzOCAxMCA0Ny45NzkgMTAgNjguOTU0MiAyNy4wMDM4IDg1Ljk1OCA0Ny45NzkgODUuOTU4IDY4Ljk1NDIgODUuOTU4IDg1Ljk1OCA2OC45NTQyIDg1Ljk1OCA0Ny45NzkgODUuOTU4IDQ3Ljk3OCA4NS45NTggNDcuOTc3IDg1Ljk1OCA0Ny45NzYgODUuOTY1MiAyNy4wMDk2IDY4Ljk3NDQgMTAuMDA3MiA0OC4wMDggMTAgNDcuOTk4MyAxMCA0Ny45ODg3IDEwIDQ3Ljk3OSAxMFpNNTcuMTc5IDQ5Ljc5M0M1MS42NDU3IDU1LjMxNjMgNDYuMDc5IDYwLjg4ODcgNDAuNDc5IDY2LjUxIDM1LjcyMDMgNjEuNzI4NyAzMC45NTAzIDU2Ljk1ODcgMjYuMTY5IDUyLjJMMzAuOTUgNDcuNDE5IDQwLjQ3OSA1Ni45NDhDNDUuMTA1IDUyLjI1NTMgNDkuNzA1IDQ3LjU5NTcgNTQuMjc5IDQyLjk2OSA1OC44NSAzOC4zNDMgNjEuMzc5IDM1Ljg2IDY2LjEyIDMxLjI3OCA2Ni4yNTMgMzEuMTQ1IDY2LjM5NiAzMS4wMTMgNjYuNTUxIDMwLjg3OCA2Ni43MDEyIDMwLjc1MDYgNjYuODM1NSAzMC42MDU2IDY2Ljk1MSAzMC40NDZMNzEuNzk5IDM1LjIyN0M2Ni4xNjggNDAuODI3IDYyLjcxMiA0NC4yNjcgNTcuMTc4IDQ5Ljc5WiIvPjwvc3ZnPg==';

async function check() {
  await Word.run(async (context) => {
    const pictures = context.document.body.inlinePictures;
    pictures.load('items/width,items/height,items/altTextDescription');
    await context.sync();
    const report = { host: Office.context.diagnostics, inlinePictures: pictures.items.length };
    const picture = pictures.items[0];
    report.altTextDescription = picture.altTextDescription;
    // the image, on its own sync
    try {
      const image = picture.getBase64ImageSrc();
      await context.sync();
      report.base64Length = image.value.length;
      report.base64Start = image.value.slice(0, 16);
      report.equalsSvg = image.value === SVG;
    } catch (e) {
      report.base64Error = String(e.message || e);
    }
    // imageFormat, on its own sync, in case this Word does not have it
    try {
      picture.load('imageFormat');
      await context.sync();
      report.imageFormat = picture.imageFormat;
    } catch (e) {
      report.imageFormatError = String(e.message || e);
    }
    out.value = JSON.stringify(report, null, 2);
  });
}
```

And the Script Lab snippet for 45 (Word, Office JS, WordApi 1.1, and `imageFormat` where this Word has
it), on `__tmp/45/svg-and-plain.docx`. The HTML tab:

```html
<button id="run">Run the check</button>
<p>Copy this back:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

async function check() {
  await Word.run(async (context) => {
    const pictures = context.document.body.inlinePictures;
    pictures.load('items/width,items/height,items/altTextDescription');
    await context.sync();
    const report = {
      host: Office.context.diagnostics,
      inlinePictures: pictures.items.map((p) => ({ altTextDescription: p.altTextDescription, width: p.width, height: p.height })),
    };
    // each image, on its own sync
    try {
      const images = pictures.items.map((p) => p.getBase64ImageSrc());
      await context.sync();
      images.forEach((image, i) => {
        report.inlinePictures[i].base64Length = image.value.length;
        report.inlinePictures[i].base64Start = image.value.slice(0, 16);
      });
    } catch (e) {
      report.base64Error = String(e.message || e);
    }
    // imageFormat, on its own sync, in case this Word does not have it
    try {
      pictures.items.forEach((p) => p.load('imageFormat'));
      await context.sync();
      pictures.items.forEach((p, i) => { report.inlinePictures[i].imageFormat = p.imageFormat; });
    } catch (e) {
      report.imageFormatError = String(e.message || e);
    }
    out.value = JSON.stringify(report, null, 2);
  });
}
```

And the Script Lab snippet for 46 (Word, Office JS, WordApi 1.3, and `imageFormat` where this Word has
it), on `46-svg-binding-word365.docx`. The HTML tab:

```html
<button id="run">Run the check</button>
<p>Copy this back:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

async function check() {
  await Word.run(async (context) => {
    const pictures = context.document.body.inlinePictures;
    pictures.load('items/width,items/height');
    await context.sync();
    const controls = pictures.items.map((p) => { const c = p.parentContentControlOrNullObject; c.load('isNullObject,type,tag'); return c; });
    await context.sync();
    const report = {
      host: Office.context.diagnostics,
      inlinePictures: pictures.items.map((p, i) => ({
        width: p.width,
        height: p.height,
        control: controls[i].isNullObject ? null : { type: controls[i].type, tag: controls[i].tag },
      })),
    };
    // each image, on its own sync: an SVG begins PHN2Zy, a PNG iVBORw0KGgo
    try {
      const images = pictures.items.map((p) => p.getBase64ImageSrc());
      await context.sync();
      images.forEach((image, i) => {
        report.inlinePictures[i].base64Length = image.value.length;
        report.inlinePictures[i].base64Start = image.value.slice(0, 16);
      });
    } catch (e) {
      report.base64Error = String(e.message || e);
    }
    try {
      pictures.items.forEach((p) => p.load('imageFormat'));
      await context.sync();
      pictures.items.forEach((p, i) => { report.inlinePictures[i].imageFormat = p.imageFormat; });
    } catch (e) {
      report.imageFormatError = String(e.message || e);
    }
    out.value = JSON.stringify(report, null, 2);
  });
}
```

48. What Office JS's text holds for each kind of inline content (CR-012 D6 and section 12.3, 2026-10-21;
   Jason: "Write Word check 48"). Under D2 the Office JS readers keep Office JS's text, and phase 4 moves
   `Paragraph.text` off the objects package's `textOf` onto this package's own run reader, which then has
   to give what Office JS gives. Check 42 found a footnote or endnote reference read as U+0002 and a
   comment reference as nothing; check 49 found a `w:br` in a branch read as U+000B. This check measures
   every kind of inline content that reader maps, and what `search` and a range do across each. The file
   is `__tmp/48/48-text.docx` on the share, as `fixtures/check48/build.mjs` writes it (its parts as
   text). The file isn't committed: case 11's tracked change would make the weekly `RevisionsHarness`
   cut a CR-011 golden for it. Each case is one paragraph,
   `Case n: <what>.`, then two upper-case marker words with the item between them:
   - **1** `FOOTA`, a footnote reference, `FOOTB`; then `ENDA`, an endnote reference, `ENDB`.
   - **2** `COMA COMMENTED`, a comment on `COMMENTED`, its reference, `COMB`.
   - **3 to 6** `LINEA`/`LINEB` around a `w:br`; `PAGEA`/`PAGEB` a page break; `COLA`/`COLB` a column
     break; `CRA`/`CRB` a `w:cr`.
   - **7** `TABA`, a tab, `TABB`, a right-aligned `w:ptab`, `RIGHTB`.
   - **8** `HYPHA`, a non-breaking hyphen, `HYPHB`, an optional hyphen, `HYPHC`.
   - **9** `SYMA`, a Wingdings `w:sym` (F04A, a smiling face), `SYMB`.
   - **10** `HIDA`, a hidden run `HIDDENTEXT` (`w:vanish`), `HIDB`.
   - **11** `REVA`, a tracked deletion `DELETEDTEXT`, a tracked insertion `INSERTEDTEXT`, `REVB`.
   - **12** `FLDA`, a field with code `QUOTE "CODETWELVE"` and result `RESULTTWELVE`, `FLDB`; **13** the
     same as a `w:fldSimple` (`CODETHIRTEEN`, `RESULTTHIRTEEN`), between `SIMA` and `SIMB`.
   - **14** `LINKA`, a hyperlink `LINKTEXT` to a bookmark, `LINKB`; **15** `SDTA`, an inline content
     control holding `SDTTEXT`, `SDTB`.
   - **16** `PICA`, an inline picture (a green rectangle), `PICB`; **17** `MATHA`, an equation `x=1`
     (`m:oMath`), `MATHB`.
   - **18** a table of one row, cells `CELLA` and `CELLB`, then a paragraph `AFTERTABLE`.
   - **Steps, in Word 365 only.** Open `__tmp/48/48-text.docx`, and say whether Word reports a problem or
     offers a repair. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New
     snippet), paste the HTML and Script tabs of the snippet for 48 below over the defaults (both tabs
     are also on the share, `__tmp/48/check48-script-lab.html` and `check48-script-lab.js`), click
     **Run the check** and copy the JSON back. Nothing is edited; close without saving.
   - **What the JSON holds:** each paragraph's `text` and its whole range's text (`paragraphs`);
     `bodyText`; each query's hits, as their texts, or the error where Office JS refuses the query
     (`searches`); and the text of a range from one marker to the other, through `expandTo` (`spans`).
     Every character outside printable ASCII is written `\uXXXX`. The `^f`, `^e`, `^l`, `^m`, `^t`, `^~`,
     `^-` and `^g` queries are Word's Find codes for a footnote mark, an endnote mark, a line break, a
     page break, a tab, a non-breaking hyphen, an optional hyphen and a graphic; whether Office JS's
     `search` takes them is part of the question.
   - **What this package gives**, on the same file (`Paragraph.text`, through `textOf`):
     - note and comment references: nothing (`FOOTAFOOTB`, `ENDAENDB`, `COMMENTEDCOMB`);
     - every break, of any type, and the `w:cr`: `\n` (U+000A); the tab `\t`; the `w:ptab` nothing;
     - the hyphens U+2011 and U+00AD; the symbol U+F04A, its private-use code point;
     - hidden text: kept (`HIDAHIDDENTEXTHIDB`);
     - the revisions: the accepted view, the insertion in and the deletion out (`REVAINSERTEDTEXTREVB`);
     - a field: its result only, in both forms; the hyperlink's and the content control's text;
     - the picture: nothing (`PICAPICB`); the equation: nothing (`MATHA  MATHB`);
     - `Body.text`: the paragraphs joined by `\n`, each cell's paragraph a line of its own;
     - `search`: one hit across a note or comment reference and across the picture, and one each for
       the hidden word (and `HIDAHIDDENTEXTHIDB`, not `HIDAHIDB`), the inserted word and each field's
       result; none across any break, tab, hyphen or symbol, none for the deleted word, a field-code
       word or the equation, and none for a `^` code. There is no `expandTo`.
   - **What follows from it** (CR-012 phase 4, D6): each item's character in Office JS's text is what
     the run reader writes, and `search` and `Range` offsets count it. Where Office JS's `search` doesn't
     match across a character its `text` holds, this package's search keeps that rule too.
   - **Run 2026-10-21 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab).** His JSON is
     `fixtures/check48/48-text-word365.json`. The view Word was in (hidden text shown or not, the
     markup shown for review) wasn't recorded. `paragraph.text` and a range from one marker to the
     other agree in every case:

     | Case | Office JS's text | This package's |
     |---|---|---|
     | 1 footnote, endnote reference | U+0002 each | nothing |
     | 2 comment reference | nothing | nothing |
     | 3 `w:br` | U+000B | `\n` |
     | 4 page break | U+000C | `\n` |
     | 5 column break | U+000E | `\n` |
     | 6 `w:cr` | U+000D, within the paragraph | `\n` |
     | 7 tab; `w:ptab` | `\t`; `0` (U+0030) | `\t`; nothing |
     | 8 non-breaking, optional hyphen | U+001E, U+001F | U+2011, U+00AD |
     | 9 Wingdings `w:sym` F04A | `(` (U+0028) | U+F04A |
     | 10 hidden run | left out (`HIDAHIDB`) | kept |
     | 11 deletion, insertion | both in (`REVADELETEDTEXTINSERTEDTEXTREVB`) | the insertion only |
     | 12, 13 fields | the result | the result |
     | 14, 15 hyperlink, content control | the text | the text |
     | 16 inline picture | nothing (`PICAPICB`) | nothing |
     | 17 equation `x=1` | ` 𝑥=1 ` (U+1D465, mathematical italic x) | nothing (`MATHA  MATHB`) |
     | 18 a cell's paragraph | its text; its whole range ends with `\t` | its text |

     - **A paragraph's whole range ends with `\r`** (U+000D), and `body.text` joins the paragraphs with
       `\r`: `...paragraph.\rCELLA\tCELLB\t\tAFTERTABLE\r...`, each cell ending in `\t` and the row in
       another `\t`. Here `Body.text` joins with `\n`, each cell a line.
     - **`search`:**
       - It matches across a comment reference (`COMMENTEDCOMB`), and across an optional hyphen with
         the plain query (`HYPHBHYPHC`, its hit `HYPHB\u001fHYPHC`).
       - A non-breaking hyphen is matched by a plain `-` (`HYPHA-HYPHB`) and by `^~`.
       - Across a note reference, a break, a tab, the `w:ptab`, the symbol or the picture the plain
         query finds nothing. Word's Find codes do work there: `^f`, `^e`, `^l`, `^m`, `^t`, `^~`, `^-`
         and `^g` each match once, each hit's text holding the character above (`FOOTA\u0002FOOTB`).
         `^g`'s hit is `PICAPICB`, so the picture is a position in the search and nothing in the
         text. Nothing matched across the column break or the `w:cr` (no Find code was tried for
         the first, and `^l` doesn't match the second).
       - The hidden word is found, with a hit whose text is `""`. `HIDAHIDDENTEXTHIDB` is found, its
         text `HIDAHIDB`; `HIDAHIDB` is not.
       - The deleted word is found, and so is `REVADELETEDTEXTINSERTEDTEXTREVB`;
         `REVAINSERTEDTEXTREVB` is not.
       - A field's result is found and its code isn't. `x=1` isn't found (the text has `𝑥`), and neither
         is `CELLACELLB`.
     - So Office JS's text is Word's own character stream, and its search reads hidden and deleted
       words too. CR-012 section 6.2 has what follows for phase 4.
   - **Since CR-012 phase 4 (2026-10-21, its section 17)** this package gives Word's answers in Word's default
     view. `readers.test.mjs` builds this file in memory (`check48Bytes`) and holds `Paragraph.text`, the whole
     ranges, `Body.text` and every query to the JSON. Only case 11's deleted text and the six hidden- and
     deleted-word queries differ, which this run read in All Markup with hidden text shown.


   The Script Lab snippet for 48 (Word, Office JS, WordApi 1.3), on `__tmp/48/48-text.docx`. The HTML tab:

```html
<button id="run">Run the check</button>
<p>Copy this back:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

// every character outside printable ASCII as \uXXXX, so a control character shows
const vis = (s) => s.replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'));

const QUERIES = ['FOOTAFOOTB', 'FOOTA^fFOOTB', 'ENDAENDB', 'ENDA^eENDB', 'COMMENTEDCOMB',
  'LINEALINEB', 'LINEA^lLINEB', 'PAGEAPAGEB', 'PAGEA^mPAGEB', 'COLACOLB', 'CRACRB', 'CRA^lCRB',
  'TABATABB', 'TABA^tTABB', 'TABBRIGHTB', 'HYPHAHYPHB', 'HYPHA-HYPHB', 'HYPHA^~HYPHB', 'HYPHBHYPHC',
  'HYPHB^-HYPHC', 'SYMASYMB', 'HIDAHIDB', 'HIDDENTEXT', 'HIDAHIDDENTEXTHIDB', 'REVAREVB', 'DELETEDTEXT',
  'INSERTEDTEXT', 'REVAINSERTEDTEXTREVB', 'REVADELETEDTEXTINSERTEDTEXTREVB', 'RESULTTWELVE', 'CODETWELVE',
  'RESULTTHIRTEEN', 'CODETHIRTEEN', 'LINKTEXT', 'SDTTEXT', 'PICAPICB', 'PICA^gPICB', 'x=1', 'MATHAx=1',
  'CELLA', 'CELLACELLB'];

const SPANS = [['FOOTA', 'FOOTB'], ['ENDA', 'ENDB'], ['COMA', 'COMB'], ['LINEA', 'LINEB'], ['PAGEA', 'PAGEB'],
  ['COLA', 'COLB'], ['CRA', 'CRB'], ['TABA', 'RIGHTB'], ['HYPHA', 'HYPHC'], ['SYMA', 'SYMB'], ['HIDA', 'HIDB'],
  ['REVA', 'REVB'], ['FLDA', 'FLDB'], ['SIMA', 'SIMB'], ['LINKA', 'LINKB'], ['SDTA', 'SDTB'], ['PICA', 'PICB'],
  ['MATHA', 'MATHB'], ['CELLA', 'CELLB']];

async function check() {
  await Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load('items/text');
    body.load('text');
    await context.sync();
    const wholes = paragraphs.items.map((p) => p.getRange('Whole'));
    wholes.forEach((r) => r.load('text'));
    await context.sync();
    const report = {
      host: Office.context.diagnostics,
      paragraphs: paragraphs.items.map((p, i) => ({ text: vis(p.text), whole: vis(wholes[i].text) })),
      bodyText: vis(body.text),
      searches: {},
      spans: {},
    };
    // one query per sync, so that a query Office JS refuses fails alone
    for (const q of QUERIES) {
      try {
        const hits = body.search(q, { matchCase: true });
        hits.load('items/text');
        await context.sync();
        report.searches[q] = hits.items.map((h) => vis(h.text));
      } catch (e) {
        report.searches[q] = 'error: ' + (e.code || e.message);
      }
    }
    for (const [a, b] of SPANS) {
      try {
        const span = body.search(a, { matchCase: true }).getFirst().expandTo(body.search(b, { matchCase: true }).getFirst());
        span.load('text');
        await context.sync();
        report.spans[`${a}..${b}`] = vis(span.text);
      } catch (e) {
        report.spans[`${a}..${b}`] = 'error: ' + (e.code || e.message);
      }
    }
    out.value = JSON.stringify(report, null, 1);
  });
}
```

48b. Whether Office JS's text depends on Word's view (CR-012 section 6.2, question 2, 2026-10-21; Jason:
   "proceed as suggested"). Check 48 found hidden text left out of `paragraph.text` and deleted text
   kept, but not which view Word was in. This package has no view, so phase 4 needs one rule. The file
   is `__tmp/48/48b-text.docx` on the share, check 48's file with two cases more, written by the same
   `fixtures/check48/build.mjs` (and not committed, for check 48's reason):
   - **19** symbols in other fonts, between `SYM2A`/`SYM2B` to `SYM5A`/`SYM5B`: a Wingdings check mark
     (F0FC), Symbol's alpha (F061) and plus-minus (F0B1), and Segoe UI Symbol's U+263A. Check 48's
     Wingdings F04A read as `(`, not its low byte `J`; these say whether Word writes `(` for every
     symbol (the editor's question).
   - **20** `HIDSA`, a run hidden through its character style (`Check48Hidden`, `w:vanish`),
     `STYLEHIDDEN`, `HIDSB`: whether the style's `w:vanish` counts as a run's own does.
   - Each run reads every paragraph's `text` with `getReviewedText('Current')` and
     `getReviewedText('Original')` (WordApi 1.4, the editor's suggestion), and searches for the hidden
     and deleted words and across them.
   - **Steps, in Word 365 only.** Open `__tmp/48/48b-text.docx`, and say whether Word reports a problem
     or offers a repair. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code >
     New snippet) and paste the HTML and Script tabs of the snippet for 48b below over the defaults
     (both are also on the share, `__tmp/48/check48b-script-lab.html` and `check48b-script-lab.js`).
     Then run it five times, in five views. Each run adds its JSON to the box, so copy the box back
     once, after the fifth:
     1. **A, Word's defaults.** File > Options > Display: **Hidden text** unticked. Home > Show/Hide ¶
        off. Review > Display for Review (the list in the Tracking group): **Simple Markup**. Type
        `A simple markup, hidden not shown` in the snippet's box and click **Run the check**.
     2. **B, hidden text shown.** File > Options > Display: tick **Hidden text**, then OK. Type
        `B simple markup, hidden shown` and click **Run the check**. Then untick **Hidden text** again.
     3. **C.** Review > Display for Review: **All Markup**. Type `C all markup` and run.
     4. **D.** Review > Display for Review: **No Markup**. Type `D no markup` and run.
     5. **E.** Review > Display for Review: **Original**. Type `E original` and run.

     Copy the whole box back, set Display for Review to Simple Markup again, and close without saving.
     Nothing is edited.
   - **What this package gives**, on the same file: case 19 `SYM2A\uf0fcSYM2B SYM3A\uf061SYM3B
     SYM4A\uf0b1SYM4B SYM5A\u263aSYM5B` (each symbol's code point); case 20 `HIDSASTYLEHIDDENHIDSB`
     (kept, as case 10's direct `w:vanish`); and the rest as check 48.
   - **Run 2026-10-21 (Jason; Word 365, 16.0.20326.20232; Script Lab), on check 48's file.** The five
     runs were made on `48-text.docx`, not `48b-text.docx` (23 paragraphs, its first `Check 48.`), with
     the views as headings in the copied text and each run's `label` empty. So cases 19 and 20 are
     still to run. The views are what this check is for, and cases 10 and 11 are in both files. The
     result is `fixtures/check48/48-text-views-word365.txt`, five JSON objects in order A to E.
     - **`paragraph.text` never holds hidden text.** Case 10 reads `HIDAHIDB` in all five, hidden text
       shown (B) or not.
     - **`paragraph.text` follows Display for Review**, and only case 11 changes with it:
       - Simple Markup (A, B) and No Markup (D): `REVAINSERTEDTEXTREVB`, the accepted text;
       - All Markup (C): `REVADELETEDTEXTINSERTEDTEXTREVB`, both;
       - Original (E): `REVADELETEDTEXTREVB`, the original text.

       Check 48's run reads as C, so it was made in All Markup.
     - **`search` follows the display:**
       - the hidden word is found only with hidden text shown (B, as in check 48's run), its hit `""`;
         otherwise `HIDAHIDB` matches across it;
       - the deleted word is found in All Markup and Original;
       - the inserted word is found in all but Original.
     - **`getReviewedText('Current')` and `('Original')` don't depend on the view.** They are the
       accepted and the original text, each ending with `\r`. They are Word's raw stream:
       - hidden text kept (`HIDAHIDDENTEXTHIDB`);
       - a field as U+0013, its code, U+0014, its result, U+0015
         (`FLDA \u0013 QUOTE "CODETWELVE" \u0014RESULTTWELVE\u0015 FLDB`), a `w:fldSimple` alike;
       - the hyperlink as its field, `\u0013HYPERLINK \l "check48end" \h\u0001\u0014LINKTEXT\u0015`;
       - the content control as `<SDTTEXT>`;
       - the picture as `/`.

       Every other paragraph reads the same through them as through `text`.
     - So Office JS's `text` is the text Word displays for the markup shown, never hidden text, and
       `search` matches what is displayed. CR-012 section 6.2 has the rule this package takes. Since phase 4
       (CR-012 section 17) every paragraph of run A and of take 2 below reads here as it did in Word, and every
       query of both finds what Word's found.
   - **Take 2, the same day, on `48b-text.docx`, in Word's defaults (A).** The result is
     `fixtures/check48/48b-text-word365.json`. Every paragraph the two files share reads as in the first
     run's A, through `text` and both `getReviewedText`s.
     - **Case 19: every `w:sym` reads `(`**: `SYM2A(SYM2B SYM3A(SYM3B SYM4A(SYM4B SYM5A(SYM5B`, the
       Wingdings check mark, Symbol's alpha and plus-minus, and Segoe UI Symbol's U+263A alike, through
       `getReviewedText` too. So Word writes one constant character for a symbol, whatever its font or
       code, as for check 48's F04A.
     - **Case 20: a run hidden through its character style is left out** of `text` (`HIDSAHIDSB`), as a
       run's own `w:vanish` is; `search` matches across it and doesn't find its word; both
       `getReviewedText`s keep it. So the style's `w:vanish` counts, as the effective value.


   The Script Lab snippet for 48b (Word, Office JS, WordApi 1.4), on `__tmp/48/48b-text.docx`. The HTML tab:

```html
<p>View: <input id="label" size="40" placeholder="A simple markup, hidden not shown"></p>
<button id="run">Run the check</button>
<p>Copy all of this back after the last run:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value += '\n' + String(e.stack || e); }));

// every character outside printable ASCII as \uXXXX, so a control character shows
const vis = (s) => s.replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'));

const QUERIES = ['HIDDENTEXT', 'HIDAHIDB', 'HIDAHIDDENTEXTHIDB', 'STYLEHIDDEN', 'HIDSAHIDSB', 'HIDSASTYLEHIDDENHIDSB',
  'DELETEDTEXT', 'INSERTEDTEXT', 'REVAREVB', 'REVAINSERTEDTEXTREVB', 'REVADELETEDTEXTREVB', 'REVADELETEDTEXTINSERTEDTEXTREVB'];

async function check() {
  const label = document.getElementById('label').value;
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text');
    await context.sync();
    const report = { label, host: Office.context.diagnostics, paragraphs: [], searches: {} };
    let current = [];
    let original = [];
    try {
      current = paragraphs.items.map((p) => p.getReviewedText('Current'));
      original = paragraphs.items.map((p) => p.getReviewedText('Original'));
      await context.sync();
    } catch (e) {
      report.reviewedTextError = 'error: ' + (e.code || e.message);
      current = [];
      original = [];
    }
    report.paragraphs = paragraphs.items.map((p, i) => ({
      text: vis(p.text),
      current: current[i] ? vis(current[i].value) : null,
      original: original[i] ? vis(original[i].value) : null,
    }));
    // one query per sync, so that a query Office JS refuses fails alone
    for (const q of QUERIES) {
      try {
        const hits = context.document.body.search(q, { matchCase: true });
        hits.load('items/text');
        await context.sync();
        report.searches[q] = hits.items.map((h) => vis(h.text));
      } catch (e) {
        report.searches[q] = 'error: ' + (e.code || e.message);
      }
    }
    out.value += (out.value === '' ? '' : '\n') + JSON.stringify(report, null, 1);
  });
}
```

49. What Office JS reads over a kept `mc:AlternateContent`'s branch (CR-012 section 12, 2026-10-20).
   Under CR-012's D2 the Office JS readers keep Office JS's text, so the editor's ED-009 rule (read a
   run-level element as its selected branch's items, a paragraph-level one as its runs, and edit in
   the selected branch) is Office JS's only if Office JS reads the branch. (48, written later, is
   CR-012 D6's check.) The file is the editor's `mc-branches.docx`: eight kept elements, none edited,
   each under a paragraph naming its case. The snippet, `check49-script-lab.js`, is beside it on the
   share (`Office 2016/mc-fun/`); the editor's ED-009 section 9.2 item 4 has the cases.
   - **Run 2026-10-20 (Jason; Word 365, Script Lab).** His save is the editor's
     `packages/editor-model/test/fixtures/mc/mc-branches-check49-word365.docx` (editor `3189640`).
     - **`paragraph.text` gives the selected branch's words in every case**, never a Fallback's and
       never empty: case 1 `The Earth is flat`; case 2 `Before\tafter the tab\u000bafter the break`;
       case 3 `unknown choice` (a Choice requiring `x14ac`, understood); case 4 `nothing to select`
       (also `x14ac`); case 5 `A picture:  after it`; case 6 `I prefer w14`; case 7 `I prefer w16du`;
       case 8, the paragraph-level element, `bold ying plain ying`.
     - **`body.search`** finds each word that is only in a Choice, one hit each (`w16du` two: the
       branch and its case label), and none that is only in a Fallback (`round`, `Fallback text`,
       `fallback shown`, `yang`: zero each).
     - **A range from the first `flat` hit** reads `flat`, its paragraph `The Earth is flat`.
       `insertText("YY", "After")` through it lands after the hit, and the paragraph reads
       `The Earth is flatYY`. Word's save holds two plain runs, `The Earth is flat` and `YY`. As in
       ED-009 S6, the save has no `mc:AlternateContent` left.
   - **What this package gives**, on the same file in a kept load: case 1 `""`, case 2 `""`, case 6
     `I prefer `, and `search('flat')` nothing. A run-level element is one object segment of no text
     (CR-010 section 13.1). Case 8 reads `bold ying plain ying` through `textOf`, but an edit there throws
     (CR-012 section 12.4). For cases 3, 4 and 7, `selectedBranchOf` took another branch than Word's
     until CR-010 section 15 (`65de0fd`), which makes it take Word's.
     Resolved, they read `fallback shown`, `""` (the element is kept whole, nothing selectable) and
     `I prefer w14` (CR-010 section 15).
   - **Also found: a line break is U+000B in Office JS.** Case 2's `w:br` reads as `\u000b`, where
     `Paragraph.text` (`textOf`) gives `\n`. It was measured on a break inside a branch. Whether
     Office JS gives U+000B for every `w:br` isn't measured here; check 48 can take it.

50. Whether Word 365 reads a `Requires` prefix alike wherever it is declared (CR-010 section 14.9,
   2026-10-20; Jason: "yes", to building it). The ports declare a prefix only the source bound in
   different places: objects CR-011 on the part root, docx4j CR-052 D4 on the `mc:AlternateContent`,
   and Word 365's own saves on the `mc:Choice` that names it. The three are equivalent XML; this check
   asks whether Word treats them so. The file is `__tmp/50/50-declarations.docx` on the share, a copy
   of `fixtures/check50/50-declarations.docx` (built by `fixtures/check50/build.mjs`, which writes
   `word/document.xml` as text, so each declaration stays where it is put). It has five cases, each a
   paragraph `Case n: ` followed by one run-level `mc:AlternateContent`:
   - **The Choice holds a text box** (inline, blue outline) reading `n: the Choice (a text box)`; the
     Fallback is plain text, `n: the Fallback (plain text, no box)`. A box means Word took the Choice.
   - **Case 1** is Word's own form: `Requires="wps"`, declared on the root. Expected: a box.
   - **Cases 2 to 4** bind the `wps` namespace to another prefix and write the Choice's shape elements
     with it: `sroot` declared on the root (2), `sac` on the `mc:AlternateContent` (3), `schoice` on the
     `mc:Choice` (4). If Word resolves `Requires` through the namespace in scope, as Markup
     Compatibility says, all three are boxes.
   - **Case 5** requires `zz9`, a namespace Word does not know. Expected: plain text, which shows what
     taking a Fallback looks like.
   - **Steps, in Word 365 only.** Open `__tmp/50/50-declarations.docx`. Say whether Word reports a
     problem or offers a repair. Before anything else, File > Save As `50-declarations-word365.docx` in
     the same folder. For each case 1 to 5, say whether it shows a box or plain text, and the words it
     shows (or take a screenshot). Close without saving again.
   - **What this package gives**, in a kept load: `selectedBranchOf` takes the Choice in case 1 and the
     Fallback in cases 2 to 5, since it matches `Requires` by prefix (REQ-303) and `sroot`, `sac` and
     `schoice` are not in `READ_UNDERSTOOD_PREFIXES`. The keeping pass's `mce.branches_differ`
     warnings, which resolve by namespace, name the Choice as read in cases 2 to 4. So the two disagree
     on an aliased prefix, the reverse of CR-010 section 14.8's case (a Choice whose prefix doesn't
     resolve: the keeping pass takes the Fallback there, the readers the Choice).
   - **Run 2026-10-20 (Jason; Word 365).** Cases 1 to 4 show a box with a blue outline, each reading
     `n: the Choice (a text box)`; case 5 shows plain text, `5: the Fallback (plain text, no box)` (his
     screenshot). No problem on opening was mentioned. **Word reads a `Requires` prefix alike whether
     it is declared on the root, on the `mc:AlternateContent` or on the `mc:Choice`.** His save is
     `fixtures/check50/50-declarations-word365.docx` (copied from the share):
     - cases 1 to 4 are kept, each now `Requires="wps"` with `wps:wsp`, `wps` declared on the root.
       Word renamed the three aliases to its own prefix; `sroot`, `sac` and `schoice` are gone from the
       Choices;
     - case 5 is resolved: its Fallback's text is a plain run, and the Choice is gone;
     - each kept `mc:Fallback` carries unused copies of the source's declarations outside Word's fixed
       set: `zz9`, `sroot` and `a` from the root on all four, and `sac` on case 3's, from its
       `mc:AlternateContent`. `schoice`, declared on the Choice Word rewrote, isn't copied. This is the
       Fallback pattern the 2,728 saves show (CR-010 section 14.9).
     - So in cases 2 to 4 Word takes the Choice where this package's readers take the Fallback: a
       prefix it doesn't list, bound to a namespace it understands, is matched by name here (REQ-303).

51. What CR-012 phase 4 left unmeasured (its section 17.7, 2026-10-21; Jason: "we'd better measure those 3
   things against Word"). The file is `__tmp/51/51-text.docx` on the share, as `fixtures/check51/build.mjs`
   writes it (its parts as text; git ignores the `.docx`, and `check51Bytes` gives a test the same bytes). The
   snippet is beside it, `check51-script-lab.html` and `check51-script-lab.js` (WordApi 1.5, for the notes).
   Each case is a paragraph `Case n: <what>.`, then its item, between upper-case markers where it is inline:
   - **The joiner** (`Table.values`, `TableCell.value`, `ContentControl.text`, `Comment.content` hold several
     paragraphs: what goes between them?). **1** a table whose first cell holds `CELLP1`, `CELLP2`, `CELLP3`
     and whose second holds `OTHER`; **2** a block content control (rich text, titled `Check 51 block`) of
     `CCP1`, `CCP2`, `CCP3`; **3** an inline control (`Check 51 inline`) holding `INLINECC`, for its range;
     **4** a comment of `COMP1`, `COMP2`, `COMP3` on `COMMENTED`.
   - **A note's own mark** (`w:footnoteRef`, `w:endnoteRef`, at the start of the note's first paragraph).
     **5** a footnote of two paragraphs, `NOTEP1` and `NOTEP2`, then an endnote of one, `ENDP1`.
   - **The letters of an equation** (check 48 measured only `x`, U+1D465). Each equation stands between
     markers `MnA` and `MnB`:
     - **6** the default style: `abhxyzABHXYZ`; digits and operators `0123+-=<>()`;
     - **7** the default style: small Greek `αβγδεθλμπσςφω`, capital Greek `ΓΔΘΛΠΣΦΨΩ`, the variants and
       nabla `ϑϕϖϵϱϰ∂∇`;
     - **8** `m:sty` plain (`sinx2Γ`), bold (`bx2Γ`), italic (`ix2Γ`) and bold italic (`bix2Γ`);
     - **9** `m:scr` script (`BELeagz1`), fraktur (`CHIRZgz1`), double-struck (`CNRZz1`), sans-serif
       (`ab1α`), monospace (`az1`);
     - **10** bold script, bold fraktur, bold sans-serif, bold italic sans-serif;
     - **11** `m:nor` (normal text) `nx`, then `nx` followed by a math run `y`;
     - **12** the structures: a fraction a over b, x squared, the square root of x, a sum from i=1 to n of x,
       x in parentheses, sin x (`m:func`);
     - **13** a display equation (`m:oMathPara`) of two equations, `x=1` and `y=2`, alone in its paragraph.
   - **Steps, in Word 365 only.** Open `__tmp/51/51-text.docx`, and say whether Word reports a problem or
     offers a repair. Leave Word's view as it opens (Simple Markup, hidden text not shown). Open Script Lab
     (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet), paste the HTML and Script tabs
     from the share's two files over the defaults, click **Run the check** and copy the JSON back. Nothing is
     edited; close without saving. A screenshot of the equations (cases 6 to 13) as Word draws them helps
     tell a drawn letter from the text Office JS gives.
   - **What the JSON holds:** each body paragraph's `text` and whole range; `bodyText`; the table's
     `values`, and for each cell of its row its `value`, `body.text`, its body's whole range and its
     paragraphs; each control's `text`, its `Whole` and `Content` ranges and paragraphs; each comment's
     `content`; each footnote's and endnote's `body.text`, whole range, paragraphs and `reference` text.
     Every character outside printable ASCII is written `\uXXXX`, a surrogate pair as two.
   - **What this package gives**, on the same file (CR-012 phase 4, uncommitted when written):
     - the cell's value `CELLP1\nCELLP2\nCELLP3` (`values` the same), its body's text
       `CELLP1\rCELLP2\rCELLP3`; the block control's text `CCP1\nCCP2\nCCP3`, its `Whole` and `Content`
       ranges `CCP1`, its first paragraph's (a range is one paragraph's here); the inline control's
       `INLINECC` all three ways; the comment's content `COMP1\nCOMP2\nCOMP3`;
     - the footnote's paragraphs `\u0002 NOTEP1` and `NOTEP2`, the endnote's `\u0002 ENDP1`;
     - the equations in Unicode's mathematical alphabets: case 6 `𝑎𝑏ℎ𝑥𝑦𝑧𝐴𝐵𝐻𝑋𝑌𝑍` and `0123+-=<>()`; case 7
       small Greek italic (`𝛼𝛽𝛾...`), capital Greek upright, the variants italic and nabla upright; case 8
       `sinx2Γ`, `𝐛𝐱𝟐𝚪`, `𝑖𝑥2𝛤`, `𝒃𝒊𝒙2𝜞`; case 9 `ℬℰℒℯ𝒶ℊ𝓏1`, `ℭℌℑℜℨ𝔤𝔷1`, `ℂℕℝℤ𝕫𝟙`, `𝖺𝖻1α`,
       `𝚊𝚣𝟷`; case 11 `nx` and `nx𝑦`; case 12 each structure's runs in order with nothing between
       (`𝑎𝑏`, `𝑥2`, `𝑥`, `𝑖=1𝑛𝑥`, `𝑥`, `sin𝑥`); case 13 `𝑥=1𝑦=2`.
   - **Run 2026-10-21 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab), in Word's defaults**, with a
     screenshot of cases 6 to 13. His JSON is `fixtures/check51/51-text-word365.json`. Whether Word reported a
     problem on opening wasn't said.
     - **The joiner is the paragraph mark, `\r`**, as `body.text` has it:
       - a cell's `value` (and `values`) and its `body.text` are both `CELLP1\rCELLP2\rCELLP3`;
       - the block control's `text` is `CCP1\rCCP2\rCCP3`, its `Content` range the same, and its `Whole`
         range that and a last `\r`;
       - the comment's `content` is `COMP1\rCOMP2\rCOMP3`;
       - the inline control is `INLINECC` all three ways, with no mark.
     - **`body.text` ends a table row with `\r\n`.** The cells are a `\t` apart and there is no `\t` after the
       last: `CELLP1\rCELLP2\rCELLP3\tOTHER\r\nAFTERTABLE`. Check 48's single-row table read `CELLA\tCELLB\t\t`,
       but that run was in All Markup with hidden text shown. Checks 30 and 31's saved `body.text`, in Word's
       defaults, end every row `\r\n` too (`row one\r\nrow two\r\n`; a nested table `nested\r\nouter cell\r\n`).
       So the row's end depends on the view. A paragraph's whole range in a cell ends `\t` in both.
     - **A note's own mark reads U+0002**, as its reference does: the footnote's paragraphs are `\u0002 NOTEP1`
       and `NOTEP2`, its `body.text` `\u0002 NOTEP1\rNOTEP2`, and its reference `\u0002`; the endnote's
       `\u0002 ENDP1`.
     - **The equations:**
       - Every letter is in the alphabet Word draws it in, Greek capitals and nabla italic too (the screenshot
         shows them slanted).
       - A digit is upright in its script's family: bold under bold italic (`𝒃𝒊𝒙𝟐𝜞`), sans-serif under
         sans-serif (`𝘢𝘣𝟣α`, the alpha plain); script and fraktur have none of their own.
       - A hyphen-minus reads as the minus sign U+2212 (`0123+−=<>()`). `m:nor` text is as stored (`nx`).
       - **A structure is `,`, its arguments a `-` apart, then `.`**: the fraction `,𝑎-𝑏.`, x squared `,𝑥-2.`,
         the radical with its degree hidden `,𝑥.`, the sum `,𝑖=1-𝑛-𝑥.` (its `∑` not in the text), the
         parentheses `,𝑥.` (their characters not in the text), `sin x` `,sin-𝑥.`.
       - The display equation's two parts read with nothing between them, `𝑥=1𝑦=2`.
     - **This package since** (CR-012 sections 17.11 and 17.12): every paragraph and whole range, `body.text`,
       the cell, the controls' text and ranges, the comment and the notes read as Word's. The block control's
       range is a `SpanningRange`: all of its text, its writes refused until CR-014 makes a `Range` span
       paragraphs. `readers.test.mjs` holds them all to the JSON. *Since CR-014 phase 1 it is one `Range` across the
       paragraphs, reading the same, its writes refused until phase 2.*

   The Script Lab snippet for 51 (Word, Office JS, WordApi 1.5), on `__tmp/51/51-text.docx`. The HTML tab:

```html
<button id="run">Run the check</button>
<p>Copy this back:</p>
<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>
```

The Script tab:

```js
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

// every character outside printable ASCII as \uXXXX (a surrogate pair as two), so a control character shows
const vis = (s) => String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'));
const failed = (e) => 'error: ' + (e.code || e.message);

async function check() {
  await Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load('items/text');
    body.load('text');
    const tables = body.tables;
    tables.load('items/values');
    const controls = body.contentControls;
    controls.load('items/title,items/text');
    await context.sync();
    const wholes = paragraphs.items.map((p) => p.getRange('Whole'));
    wholes.forEach((r) => r.load('text'));
    await context.sync();
    const report = {
      host: Office.context.diagnostics,
      paragraphs: paragraphs.items.map((p, i) => ({ text: vis(p.text), whole: vis(wholes[i].text) })),
      bodyText: vis(body.text),
      tables: [], contentControls: [], comments: [], footnotes: [], endnotes: [],
    };
    // (1) a cell, a control and a comment of several paragraphs
    for (const table of tables.items) {
      const entry = { values: table.values.map((row) => row.map(vis)), cells: [] };
      const cells = table.rows.getFirst().cells;
      cells.load('items/value');
      await context.sync();
      for (const cell of cells.items) {
        const cellParagraphs = cell.body.paragraphs;
        cellParagraphs.load('items/text');
        cell.body.load('text');
        const whole = cell.body.getRange('Whole');
        whole.load('text');
        await context.sync();
        entry.cells.push({ value: vis(cell.value), bodyText: vis(cell.body.text), bodyWhole: vis(whole.text), paragraphs: cellParagraphs.items.map((p) => vis(p.text)) });
      }
      report.tables.push(entry);
    }
    for (const control of controls.items) {
      const whole = control.getRange('Whole');
      const content = control.getRange('Content');
      whole.load('text');
      content.load('text');
      const controlParagraphs = control.paragraphs;
      controlParagraphs.load('items/text');
      await context.sync();
      report.contentControls.push({ title: control.title, text: vis(control.text), whole: vis(whole.text), content: vis(content.text), paragraphs: controlParagraphs.items.map((p) => vis(p.text)) });
    }
    try {
      const comments = body.getComments();
      comments.load('items/content');
      await context.sync();
      report.comments = comments.items.map((c) => ({ content: vis(c.content) }));
    } catch (e) {
      report.comments = failed(e);
    }
    // (2) the notes: their text, their paragraphs and their reference (WordApi 1.5)
    for (const kind of ['footnotes', 'endnotes']) {
      try {
        const notes = body[kind];
        notes.load('items');
        await context.sync();
        for (const note of notes.items) {
          note.body.load('text');
          const noteParagraphs = note.body.paragraphs;
          noteParagraphs.load('items/text');
          const whole = note.body.getRange('Whole');
          whole.load('text');
          const reference = note.reference;
          reference.load('text');
          await context.sync();
          report[kind].push({ bodyText: vis(note.body.text), bodyWhole: vis(whole.text), paragraphs: noteParagraphs.items.map((p) => vis(p.text)), reference: vis(reference.text) });
        }
      } catch (e) {
        report[kind] = failed(e);
      }
    }
    out.value = JSON.stringify(report, null, 1);
  });
}
```

52. Ranges across paragraphs (CR-014 phase 0, 2026-10-22; Jason: "implement CR-014 phase 0"). What Office JS does
   with a `Word.Range` that spans paragraphs, where a `Range` here is one paragraph's: CR-014 section 2's seven
   questions and the editor's five additions (its section 6). Cases 86 to 95 measure CR-002 section 55's D2 as well:
   what Office JS's edits leave in a table cell (Word's table rule). The file is `__tmp/52/52-ranges.docx` on the share, as
   `fixtures/check52/build.mjs` writes it (its parts as text; git ignores the `.docx`, and `check52Bytes` gives a test
   the same bytes). The snippet is beside it, `check52-script-lab.html` and `check52-script-lab.js` (WordApi 1.6),
   copies of `fixtures/check52/`'s, which are too long to repeat here.
   - **A region per case.** Each of the 95 cases is a region of its own: a paragraph `BEGIN<n> <kind> <operation>
     <tracking> <range>`, the kind's content, and a paragraph `END<n>`. So one run measures every case on content no
     other case has touched, where a fresh copy per case would be 95 runs. The range is written in lower case in the
     `BEGIN` paragraph, and the content's words in upper case, so the snippet's whole-word, case-sensitive search
     finds a word (`HEAD7B`) in the content only. The paragraphs are told apart by their formatting:
     - a heading (Heading 1, `HEAD<n>A HEAD<n>B`);
     - a body paragraph (centred, its runs italic, `BODY<n>A BODY<n>B`);
     - a list item (List Paragraph, numbered `1.` in each case, its runs underlined, `LIST<n>A LIST<n>B`);
     - a right-aligned paragraph (`NEXT<n>A NEXT<n>B`) and a plain one (`PLAIN<n>A PLAIN<n>B`);
     - a table of two rows of two cells, each cell's words `T<n>R1C1A T<n>R1C1B` and so on.
   - **The kinds of range.** "From a word to a word" is the first word's search hit `expandTo` the second's, so it
     runs from the start of the first to the end of the second:
     - **HB**: from the heading's second word to the body paragraph's first; **BH**: the same with the body paragraph
       first, from its second word to the heading's first;
     - **LB**: from the list item's second word to the body paragraph's first; **BL**: from the body paragraph's
       second word to the list item's first;
     - **HMB**: as HB, with a plain paragraph between the two;
     - **TBL**: as HB, with the table between (the whole table in the range);
     - **INTO**: from the heading's second word to the first word of the table's second row (ending inside the
       table);
     - **CELL**: from the first cell's second word to the first word of the body paragraph after the table (starting
       inside the table);
     - **HID**: as HB, the heading holding a hidden run (`HIDDEN<n>`) and a third word after its second;
     - **FN**: as HB, the heading holding a footnote reference after its second word (the note reads `NOTE<n>`), and
       a third word;
     - **CC**: the content (`getRange('Content')`) of a block content control tagged `c52-<n>`, which holds a
       heading, a body paragraph and a list item;
     - **W1**: a body paragraph's whole range (`getRange('Whole')`, its mark included), a right-aligned paragraph
       after it;
     - **W2**: from the heading's second word to the end of the body paragraph after it, its mark included, a
       right-aligned paragraph after that.
   - **The cells of cases 86 to 95** (CR-002 section 55). Each is a table of one row, whose second cell holds
     `T<n>R1C2A T<n>R1C2B`, and whose first cell holds a list item (List Paragraph, numbered, underlined, `LIST<n>A
     LIST<n>B`), so that what Word keeps of a paragraph's properties shows:
     - **TC1**: the list item alone;
     - **TC2**: a nested table of one cell (`N<n>A N<n>B`), then the list item;
     - **TC3**: a block content control tagged `c52-<n>` around the list item.

     Their `BEGIN` paragraphs name `cell list<n>a` (the cell holding the word, through `parentTableCell`),
     `paragraph list<n>a`, or `control c52-<n>`.
   - **The operations:**
     - `delete()` on every kind, and `insertText('NEWTEXT', 'Replace')` on all but HMB, HID and FN, each untracked and
       tracked;
     - `font.bold = true` on HB, TBL, CELL and CC, untracked and tracked;
     - `insertParagraph('NEWPARA', 'Before')` and `('NEWPARA', 'After')` on HB, TBL, CELL, CC, W1 and W2 untracked,
       and on HB tracked;
     - `insertText('NEWSTART', 'Start')` and `insertText('NEWEND', 'End')` on HB, CC, W1 and W2, untracked;
     - `insertComment('NEWCOMMENT')` on HB, TBL, CELL, CC, W1, W2 and FN, untracked;
     - for CR-002 section 55: the cell's `body.insertTable(1, 1, 'End', [['NEWCELL']])` on TC1;
       `Paragraph.delete()` of the list item on TC1 (the cell's only paragraph) and TC2 (its last, after the nested
       table); the cell's `body.clear()` on TC1, each untracked and tracked; the cell's `body.insertOoxml(…,
       'Replace')` on TC1, untracked, with a flat OPC package whose body is a one-cell table (`NEWOOXML`) and nothing
       after it (the snippet's `TABLE_OOXML`); and `ContentControl.delete(false)` on TC3, untracked.

     Tracked is `document.changeTrackingMode` set to `TrackAll` by the snippet for the case; it sets it back to
     `Off` at the end.
   - **Two regions are only read.** **VIEWS** (case 1): a tracked deletion from the heading's second word through
     its mark into the body paragraph, a comment from the body paragraph's second word to the list item's first, a
     tracked insertion from the list item's second word through its mark into a plain paragraph, and a table. It
     shows what `getTrackedChanges()`, `getComments()` and their `getRange()` read when they span paragraphs.
     **SEARCH** (case 2): `^p` alone, `^p^p`, and `^p` between two paragraphs, between a list item and the table,
     after a cell's last word, between two cells, at a row's end, and between the table and the paragraph after it.
   - **Steps, in Word 365 only.**
     1. Open `__tmp/52/52-ranges.docx`, and say whether Word reports a problem or offers a repair. Leave Word's view as
        it opens (Simple Markup, hidden text not shown).
     2. File > Save As `52-ranges-word365.docx`, in the same folder, before anything else, so that the run's changes
        don't go into `52-ranges.docx`.
     3. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet), and paste the HTML
        and Script tabs from the share's `check52-script-lab.html` and `check52-script-lab.js` over the defaults.
     4. Click **Run the check**, once. It changes the document case by case, showing which case it is on, and may
        take a few minutes. When it says Done, copy the box into a new file `__tmp/52/result-run.txt`.
     5. Ctrl+S, which saves `52-ranges-word365.docx` as the run left it.
     6. File > Save As `52-ranges-word365-accepted.docx`, then Review > Accept > Accept All Changes. Click
        **Report only** (it changes nothing), copy the box into a new file `__tmp/52/result-accepted.txt`, then
        Ctrl+S and close.
   - **What the JSON holds.** In `reads`, for each kind, what its range reads before anything is written: `text`,
     `isEmpty`, `style`, `font`; `getRange` of Whole, Start, End, Content and After, with each one's text and its
     first paragraph's; `paragraphs`; `tables`; the hits of `^p` searched in it; and, as `swapped`, `expandTo` with its
     ends the other way round. In `cases`, for each case: the range's text just before the operation; what the
     operation returned (the new range's, paragraph's or comment's text, a new table's values); and the region after it. The region is its
     text, each paragraph's text, style, alignment, list string and level, table nesting level and font (for bold,
     each word's bold), its tables' values and whole range, its tracked changes (type, text, range), comments
     (content, range), content controls and footnotes. `result-accepted.txt` holds the regions only. A step Office JS
     refuses reads `error: <code>` in place of its value. Every character outside printable ASCII is written
     `\uXXXX`.
   - **What this package gives:** since CR-014 phase 1 (2026-10-22), the reads. `ranges.test.mjs` runs the same
     snippet (`CHECK52.runCheck`, its `reads` and `report` modes) through the Word shim, on `check52Bytes` and on
     Word's saves, and holds it to Word's JSON case for case. The differences that remain are asserted as they stand
     (CR-014 section 9.4): `^p` (phase 3), Word's style `Other` in three places, `Range.footnotes` and `getTextRanges`
     not implemented, a control showing its placeholder, and the tracked-change listing's grouping. Since CR-014
     phase 2 (2026-10-22), the writes too. The snippet's `run` mode goes through the shim on `check52Bytes`, and
     cases 1 to 85 are held to Word's run, to Word's save after the run (each region's markup), and, with Accept
     All, to Word's report after Accept All (CR-014 section 10.4). The differences left are `getTextRanges` (the bold
     cases' words), Word's extra empty paragraph on a control's content range, the mark formatting Word gives a
     paragraph `insertParagraph('After')` goes after (cases 67 and 68, a departure), and the listing's grouping.
     Since phase 3 (2026-10-22), `^p` (case 2 and each kind's `marks`) and `insertComment` (cases 79 to 85) are
     Word's too: the hits, and the comments' markers and ranges, but for the reference run's formatting and case 82's
     start marker (CR-014 section 11.2). Since CR-002 section 55 (2026-10-23, section 55.7), cases 86 to 95 are held
     to Word's the same way: Word's but for the new table's own paragraph in 86 and 87, and 94's returned value. When this was
     written, before phase 1, there was nothing to compare: a `Range` here was one paragraph's and had no
     `expandTo`. Until then `ranges.test.mjs` holds the file and the
     snippet to each other: each `BEGIN` paragraph parses with the snippet's own parser to the case it was built as,
     and each word the snippet searches for is in the file once. A stand-in for Office JS, where every member exists,
     ran all 95 cases through the snippet with no error in the script (not committed). For CR-002 section 55,
     nothing either: it is proposed, and these cases decide its D2, whether its repair follows Word's edits or Word's
     repair on opening.
   - **Run 2026-10-22 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab), in All Markup.** His JSON is
     `fixtures/check52/52-ranges-word365-run.json`, all 95 cases. The text reads under tracking include deleted text
     (case 1's `HEAD1B`), which check 48 found Word's default view, Simple Markup, leaves out; only those reads depend
     on the view.
     - **Reads**, for every kind, before any write:
       - A range's `text` is in `body.text`'s form between its ends: `HEAD3B\rBODY3A`, a table's rows ended
         `\r\n` and its cells a `\t` apart, a footnote reference `\u0002`, hidden text left out.
       - `getRange('Whole')` is the range itself, not widened. `'Content'` is the same, without a last mark where
         the range ends with one (W1, W2).
       - `'Start'` and `'End'` are empty, in the first and the last paragraph. `'After'` is empty, in the last
         paragraph, or in the next where the range ends with a mark.
       - `paragraphs` lists every paragraph the range touches, a table's cells' too. `expandTo` with its ends swapped
         gives the same range.
       - `font` reads `null` (`''` for the name, `Mixed` for underline) where the paragraphs differ.
     - **`^p`** (case 2) finds the marks of the paragraphs outside the table, and nothing else:
       - no cell's or row's end;
       - no hit for `^p` between a paragraph and a table's first cell, or after the table's last cell;
       - `HEAD2B^pBODY2A` finds one hit across the two paragraphs.
     - **Views that span paragraphs** (case 1):
       - the tracked deletion through the heading's mark is one `Deleted` change, its text and range
         `HEAD1B\rBODY1A `;
       - the tracked insertion through the list item's mark is one `Added` change;
       - the comment's range is `BODY1B\rLIST1A`.
     - **`delete()`, untracked:**
       - The paragraphs join. The joined paragraph has the properties of the paragraph the range ends in (HB: centred
         Normal; BH: Heading 1; LB: centred Normal; BL: the list item, numbered), and each run keeps its own
         formatting. HMB's paragraph between and TBL's whole table go.
       - A range ending with a paragraph's mark takes the mark: W1's paragraph goes whole, and W2's head joins the
         next paragraph, in the next paragraph's properties.
       - A range starting or ending inside a table (INTO, CELL) throws `GeneralException` and changes nothing,
         tracked or not. So does `Replace`.
       - FN's footnote goes with its reference.
       - CC's control stays, holding one empty paragraph in its last paragraph's properties (the list item's,
         numbered).
     - **`delete()`, tracked:** one `Deleted` change across the paragraphs (`HEAD16B\rBODY16A`), with TBL's table
       inside it. The paragraphs stay until accepted.
     - **`insertText('NEWTEXT', 'Replace')`, untracked:** as `delete()`, then the text at the join (`HEAD29A NEWTEXT
       BODY29B`, centred Normal).
       - A range ending with a mark keeps that mark: W1 leaves `NEWTEXT` in the body paragraph's properties; W2 leaves
         `HEAD38A NEWTEXT` in them, the next paragraph apart.
       - CC's control holds one paragraph, `NEWTEXT`, in the list item's properties.
     - **`insertText('NEWTEXT', 'Replace')`, tracked:** every case lists only `Added: NEWTEXT`, inserted before the
       range's first word (`HEAD39A NEWTEXTHEAD39B`), and no deletion. The save shows the deletion is there (below).
     - **`font.bold`** bolds every word in the range, a table's cells too. Tracked, it gives `Formatted` changes,
       none for the heading's word that was bold already.
     - **`insertParagraph`** splits the paragraph:
       - `'Before'` at the range's start (`HEAD57A`, `NEWPARA`, `HEAD57B`, all Heading 1), inside the cell for CELL;
       - `'After'` at its end (`BODY63A`, `NEWPARA`, ` BODY63B`, centred);
       - where the range starts at a paragraph's start or ends with its mark, the new paragraph goes before or after
         that paragraph, in its paragraph properties;
       - on CC's content Word adds an empty paragraph as well: before `NEWPARA` for `'Before'` (in the control,
         Heading 1), after it for `'After'` (numbered 2. and 3.);
       - tracked: `Added: \rNEWPARA\r`.
     - **`insertText` at `'Start'` and `'End'`** goes at the range's first and last character, never past a last
       mark (W1's `'End'`: `BODY77A BODY77BNEWEND`).
     - **`insertComment`** works on every kind, CELL's too. Its range is the range's text, W1's without the mark
       (`BODY83A BODY83B`).
     - **CR-002 section 55's cells (cases 86 to 95): Word never leaves a cell without its last paragraph.**
       - `Paragraph.delete()` of the cell's only paragraph, or of its last after a nested table, leaves that
         paragraph empty, keeping its properties (List Paragraph, numbered `1.`). `body.clear()` does the same.
       - `ContentControl.delete(false)` of a control around the cell's paragraph leaves an empty paragraph in that
         paragraph's properties, outside any control.
       - Tracked, `Paragraph.delete()` and `body.clear()` each give one `Deleted` change of the text alone
         (`LIST90A LIST90B`), not of the mark.
       - `body.insertTable(…, 'End')` puts the table after the cell's last paragraph, and an empty paragraph after
         the table in that paragraph's properties (numbered `3.`). The new table's own paragraph takes them too (List
         Paragraph, `2.`). Tracked: `Added: \rNEWCELL\r\n`.
       - `body.insertOoxml(table, 'Replace')` leaves the inserted table and, after it, the cell's paragraph, emptied,
         numbered `1.`.
     - **Word's saves and the report after Accept All** (the same day): `52-ranges-word365.docx`, saved after the
       run, and `52-ranges-word365-accepted.docx`, after Review > Accept All Changes, are in `check52/`; the report is
       `fixtures/check52/52-ranges-word365-accepted.json`. Display for Review was **All Markup** (Jason), so the text
       reads under tracking include deleted text. The markup:
       - **A tracked deletion across paragraphs** marks every mark in the range deleted (`w:pPr/w:rPr/w:del`) and
         puts each paragraph's text in `w:del`. The last paragraph's mark stays, and **no `w:pPrChange` is written**
         (the editor's addition 1). A whole table in the range has every row's `w:trPr/w:del`, and every cell's mark and
         text deleted. A hidden run and a footnote reference go inside the `w:del`.
       - **Accept All gives the untracked answer**: one paragraph in the last paragraph's properties, TBL's table gone,
         W1's paragraph gone, W2's head joined to the next paragraph. This package's `acceptAllRevisions()` on the save
         after the run gives the same paragraphs as Word's accepted save in 94 of the 95 regions (each paragraph's text,
         style, alignment, list and table nesting). The 95th, case 26, differs by Word's filler of five spaces in the
         emptied control. `ranges.test.mjs` holds it so.
       - **Untracked `delete()`** takes the hidden run with it, and the footnote leaves `footnotes.xml`. Emptying CC's
         control leaves one paragraph in the list item's `w:pPr`, holding a bold run of five spaces (Word's filler for
         an empty control).
       - **Tracked `Replace`** writes `w:ins` with the new text, then `w:del` with the range, the marks deleted as
         `delete()` deletes them, except that a last mark in the range stays. So the deletion is there, though
         `getTrackedChanges()` listed only the insertion. Accepted, the new text has no formatting of its own.
       - **Untracked `Replace`** gives the new text the replaced range's first character's formatting, as direct run
         properties: from a heading, Heading 1's font, size, colour and bold (`HEAD29A NEWTEXT BODY29B`, in a centred
         Normal paragraph); W1's the body run's italic.
       - **`font.bold`** writes `w:b` on each run in the range that isn't bold already. A heading's run, bold through
         its style, gets nothing, and under tracking gives no change. A mark in the range is bolded too (CC's middle
         paragraph's `w:pPr/w:rPr/w:b`). Tracked, each run has a `w:rPrChange`.
       - **`insertComment`** puts `w:commentRangeStart` before the range's first run and `w:commentRangeEnd` after its
         last character, before any mark, then the reference run. CELL's start is inside the cell. On CC's content
         range the start goes at the start of `w:sdtContent`, before the first paragraph
         (`w:displacedByCustomXml="prev"`), and the end in the last paragraph.
       - **`insertParagraph`**: both halves of a split paragraph keep its `w:pPr`. `'After'` gives the new paragraph
         the run properties at the split (italic). Tracked, the split's new marks are inserted (`w:pPr/w:rPr/w:ins`):
         the first half's and `NEWPARA`'s.
       - **Cases 86 to 95**: the cell's kept paragraph keeps its `w:pPr` whole (`w:pStyle`, `w:numPr`).
         - Tracked `Paragraph.delete()` and `body.clear()` put the text in `w:del` and leave the mark unmarked.
         - `insertTable(…, 'End')` splits the cell's paragraph. The table goes between the halves; the first half, which
           keeps the text, gains its last run's `w:u` as mark properties. Tracked, that first half's mark is the
           inserted one, with the table's rows; the empty paragraph after the table is the cell's own mark, unmarked.
         - Accept All leaves each cell ending in that paragraph.
     - **The saves are `check52/`** (Jason: "test/check52/, outside fixtures/ like test/mc-fun/"), since under
       `fixtures/` the weekly `RevisionsHarness` would cut CR-011 goldens for them. Its `manifest.json` gives each one's
       SHA-1. `ranges.test.mjs` holds this package's `acceptAllRevisions()` on the first to the second, region for
       region.

53. CR-014 phase 1's choices where Word was not measured (CR-014 section 9, 2026-10-22; Jason: "let's measure in Word
   the Choices you made where Word wasn't measured"). Read only: nothing is changed. The file is `__tmp/53/53-reads.docx`
   on the share, as `fixtures/check53/build.mjs` writes it (git ignores the `.docx`; `check53Bytes` gives a test the same
   bytes). The snippet is beside it, `check53-script-lab.html` and `check53-script-lab.js` (WordApi 1.6), copies of
   `fixtures/check53/`'s. Each of the 35 cases is a region, `BEGIN<n> <op> <words>` then its content then `END<n>`, and
   the snippet reads what the op names:
   - **A range's style across paragraphs** (CR-014 section 9.1 reports their shared style, else `Normal`), and where
     Word's `Other` comes from:
     - cases 1 to 3: Heading 1 into Heading 2; a centred paragraph into a plain one (both Normal); Heading 1 into
       Heading 1;
     - cases 4 to 6: a heading holding a footnote reference, a range before the reference, a range across it, and the
       paragraph itself;
     - cases 7 to 9: a run in the Emphasis character style, alone, with plain text, and the plain text alone;
     - cases 10 to 14: a paragraph after a nested table in a cell (a list item, a plain paragraph, a heading), a list
       item in a cell with no nested table, and a paragraph before a nested table.
   - **Content controls' ranges** (section 9.2 ends a row-level control's range, and a control's whose content ends with
     a table, with `\r\n`): case 15 a row-level control around a table's first row; case 16 a block control of a
     paragraph and then a table; case 17 a block control holding only a table; case 18 a cell-level control holding
     two paragraphs. Each control's text, paragraphs, and `Whole`, `Content`, `Start` and `End` ranges.
   - **The controls a range lists** (section 9.1 lists those whose content the range holds whole): case 19 a range from
     a paragraph into a block control's first paragraph; case 20 a range inside a control; case 21 a range over a
     paragraph holding an inline control.
   - **A table's ranges and the tables a range lists**: case 22 a table's `Whole`, `Start`, `End` and `After` ranges
     (only `Whole` was measured, in check 52); case 23 a cell's body's ranges, the cell holding a nested table; case 24
     a range inside a nested table's cell.
   - **Tracked changes' ranges, in each view** (section 9.2: a deletion through a mark covers the mark, `\r`; a
     group's range runs from its first piece to its last; a lone mark change's range is its paragraph's): case 25 a
     deletion through a heading's mark into the next paragraph; case 26 an insertion through a list item's mark; case
     27 an inserted mark alone; case 28 a deleted mark alone; case 29 a whole inserted paragraph; case 30 a whole
     deleted paragraph. For each: the region's changes (type, text, range text, how many paragraphs the range
     touches), and each paragraph's `getTrackedChanges()` and its whole range's.
   - **Font across runs that differ in colour and highlight** (case 31), **a range across two stories** (case 32: a
     heading's word `expandTo` a footnote's whole range), and **the body's ranges** (case 33).
   - **A control showing its placeholder** (cases 34 and 35, CR-012's: check 52 found Word reading an emptied control
     marked `w:showingPlcHdr` over five spaces as no text): one with Word's own placeholder text in the Placeholder
     Text style, one with five spaces. Its text, `placeholderText`, its paragraph's text and its whole range.
   - **Steps, in Word 365 only.** Two runs of the same snippet, one in each view.
     1. Open `__tmp/53/53-reads.docx`. Say whether Word reports a problem or offers a repair.
     2. Review > Tracking > Display for Review: choose **Simple Markup**.
     3. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet), and paste the HTML
        and Script tabs from the share's `check53-script-lab.html` and `check53-script-lab.js` over the defaults.
     4. In the snippet's box, leave the label as `Simple Markup`, click **Run the check**, and when it says Done copy
        the box into a new file `__tmp/53/result-simple.txt`.
     5. Review > Tracking > Display for Review: choose **All Markup**.
     6. Change the label to `All Markup`, click **Run the check** again, and copy the box into a new file
        `__tmp/53/result-all.txt`.
     7. Close without saving.
   - **What this package gives:** the same snippet runs through the Word shim on `check53Bytes` (no Word needed), so its
     answers are compared with Word's case for case when the run is back. Since phase 1, it gives:
     - Heading 1 into Heading 2 `Normal`, Heading 1 into Heading 1 `Heading 1`; every paragraph after a nested table its
       own style;
     - the row control's ranges `ROW15A\tROW15B\r\n`, the cell control's `Whole` `CELL18A\rCELL18B\t`;
     - the deletion through a mark `\r`, a lone mark's change its paragraph's range (`ONE27A\r`), the inserted
       paragraph `NEW29A NEW29B\r`;
     - case 31's colour the first run's (`#FF0000`; a single paragraph's font reads its first run, CR-002);
     - case 35's text the five spaces.

     `Range.parentTable` (case 22), `Range.footnotes` (case 32) and `ContentControl.showingPlaceholderText` are not
     implemented here.
   - **Run 2026-10-22 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab), in both views.** The JSONs are
     `fixtures/check53/53-reads-word365-simple.json` and `53-reads-word365-all.json`. The second run's label was left
     at `Simple Markup`, but it is the All Markup run: it differs from the first exactly in the tracked changes' text
     and ranges and in `body.text`'s length. CR-014 section 9.7 has what each case answered. In short:
     - **A range's style:** a range holding its paragraphs whole reads their style (shared, else `Normal`); a partial
       range reads its runs' character style (`Emphasis`), `Other` (named `''`) where they differ (a footnote
       reference among plain runs), and else the paragraph style. A paragraph after a nested table in a cell reads
       `Other` unless it is Normal (cases 10 and 12).
     - **Controls:**
       - a row control's text is its cells a `\t` apart (`ROW15A\tROW15B`);
       - a row control's, and a control's ending with a table, 'Content' has no row end and 'Whole' ends `\t\t`;
       - a cell control's 'Whole' ends `\t`;
       - a range lists the controls it overlaps, not one holding the range.
     - **Tables:** 'Start', 'End' and 'After' are collapsed at the first cell's start, the last cell's end and the
       next paragraph.
     - **Tracked changes** in Simple Markup:
       - a deletion's text and range read `''`;
       - an inserted mark's range is `\r` in its own paragraph;
       - each paragraph lists its part of a change across paragraphs, and a wholly deleted paragraph lists none.

       In All Markup the deleted text is read.
     - **Other cases:** a colour or highlight differing across runs reads `''`. A range `expandTo` another story throws
       `ItemNotFound`. A control showing its placeholder reads as empty, Word's own placeholder text or five spaces
       alike.
     - **This package since:** `ranges.test.mjs` runs the snippet through the Word shim and holds it to the Simple
       Markup JSON. The differences left are each listed and explained in the test (Heading 1's theme colour, which is
       CR-008 phase 2's; Word's `Other` after a nested table; the tracked-change listing and the Simple Markup readers;
       the placeholder; `Range.footnotes` and `showingPlaceholderText` not implemented). A single paragraph's font,
       read where its runs differ, and an automatic colour are Word's since CR-002 section 59 (2026-10-23).
54. Writes of a range, for CR-014 phase 3 and phase 2's choices check 52 didn't measure (CR-014 section 10.6, 2026-10-22;
   Jason: "do it", after check 54 first and phase 3's comments and `^p` meanwhile were proposed). What Office JS does
   with each write CR-014 section 10.6 lists as unmeasured, and with `insertContentControl` on a range across
   paragraphs, which phase 3 brings. The file is `__tmp/54/54-writes.docx` on the share, as `fixtures/check54/build.mjs`
   writes it (its parts as text; git ignores the `.docx`, and `check54Bytes` gives a test the same bytes). The
   snippet is beside it, `check54-script-lab.html` and `check54-script-lab.js` (WordApi 1.6), copies of
   `fixtures/check54/`'s.
   - **A region per case**, as in check 52: a paragraph `BEGIN<n> <kind> <operation> <tracking> <range>`, the
     kind's content, and a paragraph `END<n>`, the range's words in lower case there and in upper case in the
     content. The paragraphs:
     - a heading (Heading 1, three words, `HEAD<n>A HEAD<n>B HEAD<n>C`);
     - a body paragraph (centred, its runs italic, `BODY<n>A BODY<n>B`);
     - a list item (List Paragraph, numbered `1.`, its runs underlined, `LIST<n>A LIST<n>B`);
     - a plain paragraph (`PLAIN<n>A PLAIN<n>B`), and a table of two rows of two cells (`T<n>R1C1A T<n>R1C1B` ...);
     - a paragraph in the style Check 54 Italic, which makes its runs italic (`QUOTE<n>A QUOTE<n>B`);
     - a paragraph in the style Check 54 Styled, which underlines its runs, capitalises them, sets them at 16
       points and colours them red (`STYL<n>A STYL<n>B`);
     - block content controls tagged `c54-<n>`, a comment (`INSIDE<n>`, `ACROSS<n>`) and a bookmark
       (`Check54Mark<n>`).
   - **The ranges.** "From x to y" is x's search hit `expandTo` y's; "to-mark y" ends with y's paragraph, its mark
     included; "to-table y" ends with the whole of the table holding y (`table.getRange('Whole')`); "start x" starts
     at x's paragraph's start; "whole x" is x's paragraph's 'Whole' range; "word x" is x's search hit alone.
   - **The cases** (the number is CR-014 section 10.6's item each answers):
     - 1 to 3 (item 1): `insertParagraph('NEWPARA', 'Before')` and `('NEWPARA', 'After')` on the heading's middle
       word alone, and 'Before' tracked: is the heading split, or the paragraph put before or after it?
     - 4 (item 7): `insertText('NEWTEXT', 'Replace')` on the heading's middle word, tracked: is the `w:ins` before
       the `w:del` or after it?
     - 5 to 7 (item 2): from the heading's middle word to the whole table after it, deleted (untracked, tracked)
       and replaced: does the heading's head join the plain paragraph after the table?
     - 8 to 13 (item 3): the 'Whole' range of a body paragraph followed by a table, of a cell's last paragraph (a
       body paragraph before it in the cell), and of a content control's last paragraph, deleted untracked and
       tracked: is the paragraph removed, or emptied?
     - 14 to 20 (item 4): a range holding a whole content control (one plain paragraph) between the heading and the
       body paragraph, deleted and replaced, untracked and tracked; a range from the heading into a control, deleted
       untracked and tracked; a range from a control's paragraph out to the paragraph after it, deleted.
     - 21 to 25 (item 5): `font.italic = true` across a Check 54 Italic paragraph and a plain one (untracked,
       tracked); `font.size = 14` across a heading (14 points through its style) and a plain paragraph (untracked,
       tracked); `font.underline = 'Single'` across an underlined list item and a plain paragraph. Does Word write
       the property on runs that read it already?
     - 26 (item 6): `insertText('NEWTEXT', 'Replace')` from a plain paragraph into a Check 54 Styled one: what
       NEWTEXT's run properties undo of the style.
     - 27 to 30 (item 8): from the heading's first word to the body paragraph's first, deleted: a comment over the
       heading's middle word, wholly inside (untracked, tracked); a comment and a bookmark each starting inside the
       range and ending after it (untracked).
     - 31 (item 9): under tracking, a paragraph inserted after the heading (`OWN<n>A OWN<n>B`, by whoever runs the
       check), then a range from the heading's middle word into it deleted: is the inserted paragraph taken back?
     - 32 to 35 (item 10): `styleBuiltIn = 'Heading2'` across a body paragraph and a list item (untracked,
       tracked); `hyperlink = 'https://example.com/check54'` and `insertOoxml(paragraph, 'Replace')` across a heading
       and a body paragraph.
     - 36 to 43 (phase 3): `insertContentControl()` on a range from the heading's middle word to the body
       paragraph's first (untracked, tracked); from the middle word to the body paragraph's end, its mark included;
       over both paragraphs whole, with the mark (untracked, tracked); across a table; into a table; out of one.

     Tracked is `document.changeTrackingMode` set to `TrackAll` by the snippet for the case; it sets it back to
     `Off` at the end.
   - **Steps, in Word 365 only.**
     1. Open `__tmp/54/54-writes.docx`. Say whether Word reports a problem or offers a repair. Leave Word's view as
        it opens.
     2. File > Save As `54-writes-word365.docx`, in the same folder `__tmp/54`, before anything else, so that the
        run's changes don't go into `54-writes.docx`.
     3. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/54/check54-script-lab.html` over the HTML tab and `__tmp/54/check54-script-lab.js` over the Script
        tab.
     4. Click **Run the check**, once. It changes the document case by case, showing which case it is on. When it
        says Done, copy the box into a new file `__tmp/54/result-run.txt`.
     5. Click into the document's text, then File > Save, which saves `54-writes-word365.docx` as the run left it.
        (Ctrl+S right after the run goes to Script Lab's pane, which has the focus, and the document isn't saved: both
        of the run's saves in the first runs held the document from before the run.)
     6. File > Save As `54-writes-word365-accepted.docx`, in `__tmp/54`, then Review > Accept > Accept All Changes.
        Click **Report only** (it changes nothing), copy the box into a new file `__tmp/54/result-accepted.txt`,
        then Ctrl+S and close.
   - **What the JSON holds.** For each case: the range's text just before the operation; what the operation
     returned (the new range's, paragraph's or control's text, and a new control's type); and the region after it:
     its text, each paragraph's text, style, alignment, list string and level and table nesting level, its tables'
     values, its tracked changes (type, text, range), comments (content, range), content controls (tag, type, text),
     bookmarks (`getBookmarks`, hidden and adjacent included) and `hyperlink`. A step Office JS refuses reads
     `error: <code>` in place of its value.
   - **What this package gives**, before the run (CR-014 section 10). `ranges.test.mjs` runs the snippet through
     the Word shim, and every case runs with no error of the script's own.
     - Cases 1 to 3 put the new paragraph before or after the heading, unsplit; case 4 writes its `w:del` first.
     - Cases 5 to 7 keep the heading's mark, so its head doesn't join the paragraph after the table.
     - Cases 8 to 13 empty the paragraph and keep it.
     - Case 14 takes the control and case 16 replaces it, untracked; cases 15 and 17 to 20 refuse with
       `NotImplemented`.
     - Cases 21 and 22 write `w:i` on the plain runs only, the Check 54 Italic ones reading it already. Cases 23 to
       25 write the size and the underline on every run and on the marks the range holds, the heading's and the
       list item's among them.
     - Case 26 gives NEWTEXT `w:u w:val="none"`, `w:caps w:val="false"`, 11 points and automatic colour.
     - Cases 27 to 30 leave the comments' and the bookmark's markers where they fall: case 27's comment, its range
       emptied, and its reference end up in the joined paragraph.
     - Case 31 marks the heading's mark deleted and takes back the inserted words in the range, the inserted
       paragraph's own mark kept.
     - Cases 32 to 43 refuse with `NotImplemented`, and so does `getBookmarks`, which the shim doesn't have.
   - **Run 2026-10-22 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab), in All Markup** (the text reads under
     tracking show deleted text). The JSONs are `fixtures/check54/54-writes-word365-run.json` and
     `54-writes-word365-accepted.json`, Word's save after Accept All is `check54/54-writes-word365-accepted.docx`
     (a SHA-1 manifest), and CR-014 section 12 has every case. The save after the run, `check54/54-writes-word365.docx`,
     is a re-run's the same day, whose JSON was the first run's case for case: the first two saves held the document
     from before the run, Ctrl+S having gone to Script Lab's pane, and the third was made with File > Save (step 5 now
     says so). In short:
     - `insertParagraph` splits a paragraph within one paragraph's range too (cases 1 to 3).
     - A range from a paragraph's tail to a whole table's end: deleted, the head joins the paragraph after the
       table (5, 6); replaced, `GeneralException` (7).
     - A 'Whole' range with no paragraph after it: before a table or at a control's end the paragraph goes (8, 12);
       a cell's last stays, tracked its text alone deleted (10, 11).
     - A whole control in a range goes with it, tracked too (14 to 17), tracked between `w:customXmlDelRangeStart`
       and `End` markers, check 30's form for an inserted control. Into a control or out of one: the texts go,
       nothing joins across its edge, the start's paragraph takes the end's properties (18 to 20).
     - A font property goes only where a run doesn't read it already, size too (21 to 25). `Replace`'s look is
       this package's (26). A tracked `Replace` within one paragraph puts the `w:ins` first (4), which this package
       follows since 2026-10-23 (Jason).
     - A comment whose reference goes with a deletion goes (27, 28); one whose reference stays keeps its markers
       at the join, as a bookmark does (29, 30).
     - Across paragraphs: `styleBuiltIn` sets each paragraph's style and drops direct paragraph formatting (32,
       33); `hyperlink` writes a HYPERLINK field across them (34); `insertOoxml` replaces as `Replace` does (35).
     - `insertContentControl` across paragraphs wraps paragraphs held whole (38, 43) and throws `GeneralException`
       for any other range (36, 37, 39 to 42).
     - Accept All, against what this package's (docx4j's) resolution does: an emptied paragraph whose deleted mark
       has a table after it goes (9); a control's emptied last paragraph keeps a bare mark (13); a paragraph whose
       deleted mark has a control after it takes that control's first paragraph's properties (19).
     - **This package since** (CR-014 section 12): `ranges.test.mjs` holds the shim's run to Word's run, and its
       Accept All to Word's report and save, and its markup to Word's save after the run, each difference listed
       with its reason. Those are a style's resets, the hyperlink field, and the three Accept All rules, which go to
       docx4j. A deleted comment's reference run, which this package left beside the `w:del` (case 28), goes inside it
       since CR-014 section 12.6.
55. What Office JS's writers do with the characters its readers give (CR-002 section 60, 2026-10-23; Jason: "write
   check 55"). Check 48 measured only the reads: a `w:br` reads as U+000B, a note reference as U+0002, and so on.
   Section 60 decided that this package's writers turn each character back into the element it reads as, and refuse
   U+0002 and the characters XML 1.0 excludes; this check says whether Word's own writers do the same. The file is
   `__tmp/55/55-characters.docx` on the share, as `fixtures/check55/build.mjs` writes it (its parts as text; git
   ignores the `.docx`, and `check55Bytes` gives a test the same parts). The snippet is beside it,
   `check55-script-lab.html` and `check55-script-lab.js` (WordApi 1.6), copies of `fixtures/check55/`'s.
   - **A region per case**, as in check 54: a paragraph `BEGIN<n> <operation> <tracking> <characters>`, the case's
     paragraph, and a paragraph `END<n>`. The characters are code points (`U+000B`, or `U+000D,U+000A` for two); the
     snippet writes `x`, them, then `y`. The case's paragraph is `TARGET<n>`, or for a bound control `BOUND<n> ` and a
     plain-text content control tagged `c55-<n>`, holding `start<n>`, bound to `value<n>` in the custom XML part whose
     namespace is `urn:check55`.
   - **The characters**: tab U+0009, line feed U+000A, carriage return U+000D, the two together, the line break
     U+000B, the page break U+000C, the column break U+000E, the non-breaking hyphen U+001E, the optional hyphen
     U+001F, the note reference U+0002; U+0001 and U+0007, which XML 1.0 excludes too; and U+2011 and U+00AD, the
     hyphens as docx4j's text gives them.
   - **The cases:**
     - 1 to 14: `paragraph.insertText('x' + c + 'y', 'End')`, for each character in that order.
     - 15 to 28: `paragraph.insertParagraph('x' + c + 'y', 'After')`, the same.
     - 29 to 42: `insertComment('x' + c + 'y')` on the paragraph's word, the same.
     - 43 to 46: a comment made with `COMMENT<n>`, then its `content = 'x' + c + 'y'`, for U+000D, U+000A, the two
       together and U+000B.
     - 47 to 51: `insertText('x' + c + 'y', 'End')` under tracking, for U+0009, U+000A, U+000D, U+000B and U+001E.
     - 52 to 56: the bound control, multi-line, given `insertText('x' + c + 'y', 'Replace')`, for U+000B, U+000A,
       U+000D, U+001E and U+0002.
     - 57 and 58: the same on a single-line bound control, for U+000B and U+000A.

     The questions: does each character become the element it reads as, or stay a character in the `w:t`? Do `\n`,
     `\r` or the two make a paragraph? What does Word throw for U+0002, U+0001 and U+0007, if it throws? Do U+2011 and
     U+00AD stay text? What does a bound control's node hold? Tracked is `document.changeTrackingMode` set to
     `TrackAll` by the snippet for the case; it sets it back to `Off` at the end, and the author is whoever runs it.
   - **Steps, in Word 365 only.**
     1. Open `__tmp/55/55-characters.docx`. Say whether Word reports a problem or offers a repair. Leave Word's view as
        it opens.
     2. File > Save As `55-characters-word365.docx`, in the same folder `__tmp/55`, before anything else, so that the
        run's changes don't go into `55-characters.docx`.
     3. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/55/check55-script-lab.html` over the HTML tab and `__tmp/55/check55-script-lab.js` over the Script
        tab.
     4. Click **Run the check**, once. It changes the document case by case, showing which case it is on. When it
        says Done, copy the box into a new file `__tmp/55/result-run.txt`.
     5. Click into the document's text, then File > Save, which saves `55-characters-word365.docx` as the run left it.
        (Ctrl+S right after the run goes to Script Lab's pane, which has the focus, and the document isn't saved.)
     6. File > Close. Then open `__tmp/55/55-characters-word365.docx` again, and say whether Word reports a problem or
        offers a repair. If it offers a repair, say No, say so, and stop there.
     7. Open Script Lab again (Script Lab > Code); the snippet is still there. Click **Report only** (it changes
        nothing), copy the box into a new file `__tmp/55/result-reopened.txt`, then close the document without saving.
   - **What the JSON holds.** For each case: the text written (`written`), what the operation returned (the range's
     or the new paragraph's text), and the region after it: its text, each paragraph's text, the OOXML of each
     paragraph between BEGIN and END (the body's content with its rsids dropped, and the comments where the OOXML holds
     them), the comments' content and the tracked changes (type, text). For a bound case, also the control's text and
     its node's value as the custom XML part's XML holds it, with whether that XML parses (`wellFormed`, by the add-in's
     `DOMParser`). A step Office JS refuses reads `error: <code>: <message>` in place of its value.
   - **What this package gives**, before section 60 is implemented. `characters.test.mjs` runs the snippet through
     the Word shim, and every case runs with no error of the script's own; the custom XML part's read fails, since the
     shim has no `customXmlParts`.
     - Cases 1 to 14 and 47 to 51 keep every character in the `w:t`, the tracked ones inside the `w:ins`.
     - Cases 15 to 28 make a tab a `w:tab` and a line feed a `w:br` (the objects package's `r()`), and keep the rest in
       the `w:t`; the two together are a `\r` in the `w:t` and a `w:br`.
     - Cases 29 to 42 make a line feed a second paragraph in the comment and keep the rest in the `w:t`; cases 43 to
       46 make a paragraph of `\r`, `\n` and the two together, and keep U+000B in the `w:t`.
     - Cases 52 to 58 keep the characters in the control's `w:t`, multi-line or not, and in its node (read from the
       package after the run: U+000B, U+001E and U+0002 among them, which XML 1.0 excludes).
   - **Run 2026-10-23 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab).** The JSONs are
     `fixtures/check55/55-characters-word365-run.json` and `55-characters-word365-reopened.json`, and Word's save after
     the run is `check55/55-characters-word365.docx` (a SHA-1 manifest). Word reported no problem opening the file, nor
     opening its own save again; that save's document part holds no character XML 1.0 excludes. CR-002 section 60.3 has every case. In short:
     - **Word refused nothing**, and threw no error in any case.
     - **Tab, U+000B, U+000C, U+000E, U+001E and U+001F** became `w:tab`, `w:br`, a page break, a column break,
       `w:noBreakHyphen` and `w:softHyphen`, in their own run after the text, through every writer: `insertText`,
       tracked too (inside the `w:ins`), `insertParagraph`, `insertComment` and `comment.content`.
     - **`\n`, `\r` and `\r\n`** each made a new paragraph through `insertText` (`\r\n` one), its returned range reading
       `x\ry`, and through `insertComment` and `comment.content`. Tracked, the new mark is inserted
       (`w:pPr/w:rPr/w:ins`). Through `insertParagraph` they became their escapes as text: a backslash and `n`, `r`,
       or `r` and `n`.
     - **U+0002 and U+0001** left nothing in the OOXML but a run boundary where they stood. The paragraph's text read
       them until the save, and not after it. **U+0007** went at once.
     - **U+2011** stayed a character, in a run of its own in Cambria Math. **U+00AD** stayed a character in the `w:t`.
     - **The bound controls (52 to 58) weren't measured.** `body.contentControls.getByTag('c55-<n>')` found none of
       them (`ItemNotFound`), so nothing was written. Their nodes still held `start<n>`, and the custom XML part's XML
       parsed.
   - **This package since** (CR-002 section 60, implemented 2026-10-23): `characters.test.mjs` runs the snippet through
     the Word shim and holds it to Word's run case for case. 42 cases agree. Cases 16 to 18 are E2's departure (one
     paragraph, a `w:br` or `w:cr`, not Word's escape text). Cases 10, 11, 24, 25, 38 and 39 drop U+0001 and U+0002 at
     once, as Word's save does, so they are held to Word's reopened report. Cases 52 to 58 are check 55b's.
55b. The bound controls check 55 couldn't find (CR-002 section 60.5, 2026-10-23; Jason: "E1 to E4 as recommended", E4
   being this check). Check 55's `body.contentControls.getByTag` found none of its seven bound plain-text controls, so
   what Word writes into a bound control's custom XML node is still unmeasured. The file is `__tmp/55b/55b-bound.docx`
   on the share, as `fixtures/check55/build.mjs` writes it (`check55bBytes`; git ignores the `.docx`). The snippet is
   beside it, `check55b-script-lab.html` and `check55b-script-lab.js` (WordApi 1.5), copies of `fixtures/check55/`'s.
   - **A region per case**, as in check 55: `BEGIN<n> <operation> <tracking> <characters>`, a paragraph `BOUND<n> `
     and a plain-text content control tagged `c55b-<n>`, bound to `value<n>` in the custom XML part whose namespace is
     `urn:check55b`, then `END<n>`. The control holds `start<n>`, as its node does.
   - **Each control is looked for five ways, each recorded:** `getContentControls({ types: ['PlainText'] })` on the
     paragraph's whole range and on the body, and `contentControls` on the paragraph, on its whole range and on the
     body. The write goes through the first way that finds it.
   - **The cases:**
     - 1 to 6: a multi-line control, given `insertText('x' + c + 'y', 'Replace')` for U+0041 (`A`, a plain character,
       to show a write reaches the node at all), U+000B, U+000A, U+000D, U+001E and U+0002.
     - 7 to 9: a single-line control, the same for U+0041, U+000B and U+000A.
     - 10: nothing written. The control holds `stale10`, and its node holds `line1`, a line feed and `line2`: what Word
       shows of a node's line feed when it opens the file.

     The questions: which ways find a plain-text control? What does the node hold after each write? Does the
     control's text and markup follow? And what does Word make of a line feed in a node?
   - **Steps, in Word 365 only.**
     1. Open `__tmp/55b/55b-bound.docx`. Say whether Word reports a problem or offers a repair. Leave Word's view as it
        opens.
     2. File > Save As `55b-bound-word365.docx`, in the same folder `__tmp/55b`, before anything else, so that the
        run's changes don't go into `55b-bound.docx`.
     3. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/55b/check55b-script-lab.html` over the HTML tab and `__tmp/55b/check55b-script-lab.js` over the Script
        tab.
     4. Click **Run the check**, once. When it says Done, copy the box into a new file `__tmp/55b/result-run.txt`.
     5. Click into the document's text, then File > Save, which saves `55b-bound-word365.docx` as the run left it.
        (Ctrl+S right after the run goes to Script Lab's pane, which has the focus, and the document isn't saved.)
   - **What the JSON holds.** For each case: the text written, each lookup (how many controls it lists, whether it
     finds the case's, and the control's type there), the write's way and the range it returned, and the region after
     it: its paragraphs' text, the bound paragraph's OOXML, the control's text and type, and its node's value as the
     custom XML part's XML holds it, with whether that XML parses (`wellFormed`). A step Office JS refuses reads
     `error: <code>: <message>`.
   - **What this package gives**, before section 60 is implemented. `characters.test.mjs` runs the snippet through
     the Word shim. Every case's control is found through `paragraph.contentControls`, the first way the shim has:
     it has no `getContentControls`, and no `customXmlParts` to read the node through. Each write keeps its
     characters in the control's `w:t`. Case 10's control reads `stale10`, since a package read here isn't bound until
     `applyBindings` runs.
   - **Run 2026-10-23 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab).** The JSON is
     `fixtures/check55/55b-bound-word365-run.json`, and Word's save after the run is `check55/55b-bound-word365.docx`
     (in check 55's SHA-1 manifest). Word reported no problem opening the file. CR-002 section 60.6 has every case. In
     short:
     - **The lookups.** `getContentControls({ types: ['PlainText'] })` found every control, on the range (1 listed) and
       on the body (10 listed). `contentControls` on the paragraph, the range and the body listed none: 0 of the
       body's 10 plain-text controls.
     - **A line break (U+000B) into a multi-line control** became a `w:br` in it, and `\r\n` in its node. A
       non-breaking hyphen became `w:noBreakHyphen`, and `-` in the node. U+0002 went from the control (its text read
       it until the save), and became a space in the node. `A` went into both as it is.
     - **`\n` or `\r` into the control** split its paragraph and made the control a block-level one, its paragraph
       holding `x`, a `w:br` and `y` (multi-line) or `xy` (single-line). Its node read `x\r\ny` (multi-line) or `xy`.
     - **A single-line control** dropped U+000B from its text (`xy`), but its node read `x\r\ny`.
     - **Case 10:** a node holding `line1`, a line feed and `line2` showed as `line1`, a `w:br` and `line2`, and Word's
       XML of it read `line1\r\nline2`, in the session and in the save.
   - **This package since** (CR-002 section 60): `characters.test.mjs` runs the snippet through the shim, after
     `applyBindings`, which stands in for Word's open, and holds it to Word's run. The differences: cases 3, 4 and 9 are
     refused (F4); case 6 drops U+0002 at once (E3). The nodes hold `\n` for a break (F1) and `-` for a non-breaking
     hyphen (F2).
56. A bound block control of plain text and a second block (CR-002 section 61, 2026-10-23; Jason: "write check 56").
   docx4j-python's check found that Word can't open a file holding a bound block control with no type element and two
   paragraphs. This check opens the same shape as this package writes it, and the `w:text` and `w:richText` kinds beside
   it, and then measures what Office JS's own block writes do to each kind. The files are on the share in `__tmp/56/`, as
   `fixtures/check56/build.mjs` writes them (git ignores the `.docx`; `check56Bytes` and `check56OpenFiles` give a test
   the same parts). The snippet is beside them, `check56-script-lab.html` and `check56-script-lab.js` (WordApi 1.5),
   copies of `fixtures/check56/`'s.
   - **The four files to open**, each one block-level control tagged `c56-1`, bound to a custom XML node holding
     `start1`, given a second block by this package's own writers today:
     - `56a-untyped-two-paragraphs.docx`: no type element, `start1` and `second paragraph` (docx4j-python's 3a);
     - `56b-text-two-paragraphs.docx`: `w:text w:multiLine`, the same two paragraphs;
     - `56c-richtext-two-paragraphs.docx`: `w:richText`, the same two paragraphs;
     - `56d-untyped-table.docx`: no type element, a one-row table before its paragraph, so that the control still ends
       with a paragraph.
   - **The snippet's file**, `56-writes.docx`: a region per case, `BEGIN<n> <operation> <kind>`, a block-level control
     tagged `c56-<n>` holding one paragraph `start<n>` and bound to `value<n>` (namespace `urn:check56`), then `END<n>`.
     The kinds: `untyped` (no type element), `text-multi` (`w:text w:multiLine`), `text-single` (`w:text`) and `richtext`
     (`w:richText`). Each kind gets each write, at the control's end: `insertParagraph('NEWPARA<n>', 'End')`,
     `insertTable(1, 2, 'End', ...)`, `insertText('x\ny', 'End')`, and `insertOoxml` of two paragraphs. That makes 16
     cases. Each control is looked for through `getContentControls()` (all, then plain text, then rich text) and
     `contentControls` on the body, and through its paragraph's `parentContentControlOrNullObject`, each recorded. The
     write goes through the first that finds it.
   - **The questions:** does Word open each of the four files, or repair it, and how? Does Office JS refuse each write,
     or what does it write: a second paragraph, a table, a line break? Does Word then open its own save?
   - **Steps, in Word 365 only.** First the four files, one at a time:
     1. Open `__tmp/56/56a-untyped-two-paragraphs.docx`. Say whether it opens, reports a problem, or offers a repair.
     2. If Word offers to repair it, click Yes, then File > Save As `56a-untyped-two-paragraphs-repaired.docx` in
        `__tmp/56`, and close it. If it opens, close it without saving.
     3. The same for `56b-text-two-paragraphs.docx`, `56c-richtext-two-paragraphs.docx` and `56d-untyped-table.docx`,
        a repaired one saved as its name with `-repaired` before `.docx`.

     Then the snippet:
     4. Open `__tmp/56/56-writes.docx`. Say whether Word reports a problem or offers a repair.
     5. File > Save As `56-writes-word365.docx`, in the same folder `__tmp/56`, before anything else.
     6. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/56/check56-script-lab.html` over the HTML tab and `__tmp/56/check56-script-lab.js` over the Script tab.
     7. Click **Run the check**, once. When it says Done, copy the box into a new file `__tmp/56/result-run.txt`.
     8. Click into the document's text, then File > Save. (Ctrl+S right after the run goes to Script Lab's pane, which
        has the focus, and the document isn't saved.)
     9. File > Close. Open `__tmp/56/56-writes-word365.docx` again, and say whether Word reports a problem or offers a
        repair. If it offers one, click Yes, then File > Save As `56-writes-word365-repaired.docx` in `__tmp/56`.
   - **What the JSON holds.** For each case: each lookup (how many controls it lists, whether it finds the case's, its
     type), the write's way and what it returned, or Office JS's error as `error: <code>: <message>`; and the region
     after it: its text, its paragraphs' text, its OOXML (the body's content), the control's type, text, paragraphs and
     table count, and its node's value with whether the custom XML part's XML parses.
   - **What this package gives**, before section 61 is implemented. `bound-controls.test.mjs` runs the snippet through
     the Word shim, every control found through `body.contentControls`. `insertParagraph` gives every kind a second
     paragraph, and `insertText` keeps its `\n` in the `w:t` (section 60 not implemented either). The shim has no
     `insertTable` or `insertOoxml` on a control, no `getContentControls`, no `parentContentControlOrNullObject` and no
     `customXmlParts`, so those steps read `NotImplemented`. The four files to open are this package's own writes.
   - **Run 2026-10-23 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab).** The JSON is
     `fixtures/check56/56-writes-word365-run.json`, and Word's saves are in `check56/` with a SHA-1 manifest: the four
     files' repairs, the save after the run, and that save's repair. CR-002 section 61.3 has every case. In short:
     - **All four files to open** were unreadable to Word (56a, 56b, 56c and 56d alike: `w:text`, `w:richText` and a
       table too). Each repair made the control plain text (`<w:text/>`, multi-line kept for 56b, a `w:richText`
       replaced) and refilled it from its node (`start1`); the second paragraph, or the table, was lost.
     - **On opening `56-writes.docx`** (without a problem), Word read every bound control as plain text: the untyped and `w:richText`
       controls became `<w:text/>`, and every lookup gave `PlainText`. `getContentControls()` with no types found all
       16, as did `parentContentControlOrNullObject`; `getContentControls` with `RichText` and `contentControls` found
       none.
     - **`insertTable`** threw `InvalidArgument` and **`insertOoxml`** of two paragraphs `GeneralException`, in every
       kind, and nothing changed.
     - **`insertParagraph('NEWPARA<n>', 'End')`:** in a single-line control, the text joined its paragraph
       (`start1NEWPARA1`) and its node read `start1\r\nNEWPARA1`. In the multi-line control (case 5), a second
       paragraph inside the control, its node `start5\r\nNEWPARA5`.
     - **`insertText('x\ny', 'End')`:** a single-line control dropped the line feed (`start3xy`, its node the same); the
       multi-line one wrote a `w:br`, its node `start7x\r\ny`.
     - **Step 9:** Word found unreadable content in its own save and repaired it, keeping `document.xml` byte for byte
       (only the revision count, dates, rsids and paragraph ids changed). Case 5's two paragraphs are the likely cause,
       56b's shape, though the save also holds single-line controls whose nodes hold `\r\n`.
   - **This package since** (CR-002 section 61, implemented 2026-10-23): `bound-controls.test.mjs` holds the shim's run
     to Word's. Case 5 (the multi-line control's `insertParagraph`) is refused (H2). The table and OOXML cases need
     members the shim's `ContentControl` hasn't; through the content API they refuse with Word's `InvalidArgument` and
     `GeneralException`. The four files to open are written as text since, this package's writers no longer making them.
57. A paragraph break inside a field's result, and a tracked delete of an equation (CR-002 section 62, 2026-10-23;
   Jason: "write check 57"). docx4j-python's review found the two ports differ here, and neither has measured Word: this
   package splits a complex field's result and refuses a `w:fldSimple`'s, where docx4j-python refuses both; and its
   tracked `Paragraph.delete` of a paragraph holding an equation wraps the equation in one `w:del`, where its
   `Range.delete` and docx4j-python refuse. The file is on the share in `__tmp/57/`, as `fixtures/check57/build.mjs`
   writes it (git ignores the `.docx`; `check57Bytes` gives a test the same parts). The snippet is beside it,
   `check57-script-lab.html` and `check57-script-lab.js` (WordApi 1.6), copies of `fixtures/check57/`'s.
   - **The file**, `57-fields-equations.docx`: a region per case, `BEGIN<n> <operation> <shape>`, the case's paragraph,
     then `END<n>`. Ten cases:
     - 1 and 2: `insertText('x\ny', 'After')` on the word `FA<n>` inside a field's result (`QUOTE "FA<n> FB<n>"`),
       untracked: a complex field (1), a `w:fldSimple` (2);
     - 3 and 4: `insertParagraph('NEWPARA<n>', 'After')` on the same word, the same two fields;
     - 5 and 6: nothing done. The file holds what this package's tracked `Paragraph.delete` writes today: one `w:del`
       around a display equation (`m:oMathPara`, E=mc²) (5), and around an inline one with words either side (6), the
       paragraph mark deleted, author `docx4j`;
     - 7 and 8: `paragraph.delete()` under `TrackAll`, on the same two shapes;
     - 9 and 10: under `TrackAll`, the display equation's paragraph `Content` range deleted (9), and the range from the
       word `L10` to the word `R10`, the inline equation between them (10).
   - **The questions:** does Word split a field's result, refuse, or put the break elsewhere, and is the field then one
     field? Does Word show cases 5 and 6 as deleted, and do Accept All and Reject All take and keep their equations?
     What does Word write for cases 7 to 10, and how does it accept and reject that?
   - **Steps, in Word 365 only:**
     1. Open `__tmp/57/57-fields-equations.docx`. Say whether Word reports a problem or offers a repair. If it can't
        open the file, stop there and say so: cases 5 and 6 are then the likely cause, and a file without them follows.
     2. Review > Display for Review > **All Markup**. Look at regions 5 and 6 (`BEGIN5 open display-equation` and
        `BEGIN6 open inline-equation`), and say for each whether the equation shows as deleted (struck through, in the
        revision colour), as the words around it in region 6 do.
     3. File > Save As `57-fields-equations-word365.docx`, in the same folder `__tmp/57`, before anything else.
     4. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/57/check57-script-lab.html` over the HTML tab and `__tmp/57/check57-script-lab.js` over the Script tab.
     5. Click **Run the check**, once. When it says Done, copy the box into a new file `__tmp/57/result-run.txt`.
     6. Click into the document's text, then File > Save. (Ctrl+S right after the run goes to Script Lab's pane, which
        has the focus, and the document isn't saved.)
     7. Review > Accept (the lower half of the button) > **Accept All Changes**. File > Save As
        `57-fields-equations-word365-accepted.docx` in `__tmp/57`, then File > Close.
     8. Open `__tmp/57/57-fields-equations-word365.docx` again. Review > Reject (the lower half of the button) >
        **Reject All Changes**. File > Save As `57-fields-equations-word365-rejected.docx` in `__tmp/57`, then File >
        Close.
   - **What the JSON holds.** For each case: the write's tracking mode and what it returned, or Office JS's error as
     `error: <code>: <message>`; and the region after it: its text, its paragraphs' text, its OOXML (the body's
     content), its tracked changes (type, text, author) and its fields (code and result).
   - **What this package gave**, before the run. `fields-equations.test.mjs` runs the snippet through
     the Word shim. Cases 1 and 3 split the complex field's result, the field running across the paragraphs; 2 and 4
     are refused (`NotSupportedError`, a `CTSimpleField` can't be cut in two). Cases 7 and 8 write what 5 and 6 hold,
     each listed as one deletion without the equation's text; 9 and 10 are refused (a tracked edit may not take an
     equation). The shim has no `Range.fields`, which reads `NotImplemented`.
   - **Run 2026-10-23 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab; All Markup).** The JSON is
     `fixtures/check57/57-fields-equations-word365-run.json`. Word's saves are in `check57/` with a SHA-1 manifest:
     after the run, after Accept All, and after Reject All. CR-002 section 62.2 has every case. In short:
     - **Step 1:** Word opened the file without a problem. **Step 2:** regions 5 and 6 showed as deleted, the equation
       struck through in the revision colour, like the words either side of it in region 6.
     - **Cases 1 and 3:** Word split the complex field's result as this package does, the field running across the
       paragraphs (one field, result `FA1x\ry FB1`).
     - **Cases 2 and 4:** Word made the `w:fldSimple` a complex field and split it the same way. This package refuses.
     - **Cases 5 and 6:** Word read this package's one `w:del` as its own form: OOXML, save, listing, Accept All and
       Reject All as for cases 7 and 8. Word's form has a `w:del` in each `m:r` and in the structure's `m:ctrlPr`.
     - **Cases 9 and 10:** Word deleted the equation in its own form, the mark kept. This package refuses.
     - **This package on Word's saves:** Reject All matches Word's. Accept All leaves an empty `m:oMath` wherever a whole
       equation was deleted in Word's form (F1, docx4j's too). `Paragraph.text` reads such an equation's structure
       characters `,-.` (F2).
   - **This package since** (CR-002 section 62.3, implemented 2026-10-23): D1, a `w:fldSimple` made a complex field
     and split; D2, an equation deleted under tracking in Word's form; F1 and F2 fixed. `fields-equations.test.mjs`
     holds the shim's run to Word's case for case: cases 1 to 4 and 7 to 10 write Word's markup, and Word read cases 5
     and 6 as cases 7 and 8 now write them. This package's save after the run, and Word's, each accepted and rejected
     here, match Word's Accept All and Reject All region for region.
58. What CR-002 section 58, CR-012 sections 18 and 19 and CR-002 section 32's item 1 left to measure (2026-10-23; Jason:
   "one check 58 first", before implementing them). Four questions in one run: what a paragraph's tracked-change listing
   holds and what accepting a paragraph's part of a change does; a content control showing its placeholder at run level,
   in `search`, and written into; a range across a deleted paragraph mark, and a deleted mark at a cell's end; and Word's
   wildcards over an emoji and an Arabic letter with its mark. The file is on the share in `__tmp/58/`, as
   `fixtures/check58/build.mjs` writes it (git ignores the `.docx`; `check58Bytes` gives a test the same bytes). The
   snippet is beside it, `check58-script-lab.html` and `check58-script-lab.js` (WordApi 1.6), copies of
   `fixtures/check58/`'s.
   - **The file**, `58-measures.docx`: a region per case, `BEGIN<n> <op> <args>`, its content, then `END<n>`. The
     tracked changes are by `Check 58` (and `Check 58 other` in case 3). 22 cases:
     - **A paragraph's listing** (CR-002 section 58; check 53's case 30 listed no change for a wholly deleted
       paragraph): 1 a wholly deleted paragraph between two plain ones; 2 two of them in a row; 3 the same by two
       authors; 4 a paragraph whose text is deleted and whose mark isn't; 5 a wholly inserted paragraph. For each: the
       region's changes, and each paragraph's `getTrackedChanges()`, its whole range's and its content range's.
     - **A range across a deleted mark** (CR-012 section 19), its text and paragraphs: 6 a deletion through a heading's
       mark into the next paragraph; 7 a mark deleted alone; 8 a wholly deleted paragraph between the range's ends.
     - **A deleted mark at a cell's end** (CR-012 section 19), which Word never writes: 9 the cell's only paragraph; 10
       its last of two. The table's values and whole range, and the cell's value, text, whole range and paragraphs.
     - **A control showing its placeholder** (CR-012 section 18): 11 a run-level one between two words, its text and its
       paragraph's; 12 a block one, a run-level one, and a plain paragraph of the same words, searched for
       `Click or tap here to enter text.`, `tap` and `PLAIN12A Click`, paragraph by paragraph.
     - **Wildcards** (CR-002 section 32 item 1): 13 a paragraph `x😀y a😀b كَتَبَ z`, searched with wildcards for `?`,
       `[!a-z ]`, `a?b`, `x?y` and `[!a-z ]@`.
     - **Edits** (the **Edit** button only; **Read** reads them as they are): 14 to 17, a deletion through a heading's
       mark into the next paragraph: the first change paragraph 1 lists accepted (14), paragraph 2's accepted (15),
       paragraph 1's rejected (16), paragraph 2's rejected (17); 18 and 19, an insertion through a list item's mark into
       the next paragraph, paragraph 2's part accepted (18) and rejected (19). Then the region's paragraphs and changes.
       20 to 22, a control showing its placeholder written into with `insertText('FILL<n>', …)`: a block one with
       `'Replace'` (20) and `'End'` (21), a run-level one with `'End'` (22); then its text and its paragraph's.
   - **Steps, in Word 365 only:**
     1. Open `__tmp/58/58-measures.docx`. Say whether Word reports a problem or offers a repair: cases 9 and 10 hold a
        deleted mark at a cell's end, which Word never writes.
     2. File > Save As `58-measures-word365.docx`, in the same folder `__tmp/58`, before anything else (the Edit run in
        step 7 changes the document, and AutoSave would write it over the file you opened).
     3. Review > Tracking > Display for Review: choose **Simple Markup**.
     4. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/58/check58-script-lab.html` over the HTML tab and `__tmp/58/check58-script-lab.js` over the Script tab.
     5. Leave the label as `Simple Markup` and click **Read**. When it says Done, copy the box into a new file
        `__tmp/58/result-read-simple.txt`.
     6. Review > Tracking > Display for Review: choose **All Markup**. Change the label to `All Markup`, click **Read**,
        and copy the box into a new file `__tmp/58/result-read-all.txt`.
     7. Still in All Markup, with the label `All Markup`, click **Edit** once. Copy the box into a new file
        `__tmp/58/result-edit.txt`.
     8. Click into the document's text, then File > Save. (Ctrl+S right after the run goes to Script Lab's pane, which
        has the focus, and the document isn't saved.) Then File > Close.
   - **What this package gives**, before the run. `readers.test.mjs` runs the snippet through the Word shim, in both
     modes:
     - a wholly deleted paragraph lists its change, where check 53 found Word listing none;
     - a range across a deleted mark reads the mark (`HEAD6A \rBODY6B`);
     - a deleted mark at a cell's end reads as any other cell end;
     - a control showing its placeholder reads its text, and `search` finds it;
     - the wildcards' `?` and `[!a-z ]` find no half of the emoji's surrogate pair and no Arabic letter without its
       mark (section 32's grapheme filter);
     - a paragraph after a mark deleted into it lists no part of that change, so cases 15, 17, 18 and 19 have nothing to
       accept or reject; paragraph 1 lists its part (`HEAD14B\r`), and accepting or rejecting it resolves that part
       only, the next paragraph's `BODY14A ` staying a deletion;
     - `insertText(…, 'End')` on a control showing its placeholder writes after the placeholder text.
     `ContentControl.showingPlaceholderText` isn't implemented here.
   - **Run 2026-10-23 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab).** Word opened the file without a problem or a
     repair, cases 9 and 10's deleted cell-end marks among it. The JSONs are `fixtures/check58/58-measures-word365-read-simple.json`,
     `-read-all.json` and `-edit.json`. Word's save after the three runs is `check58/58-measures-word365.docx`, with a SHA-1
     manifest. In short:
     - **A paragraph's listing** (cases 1 to 5, 14 to 19), in both views:
       - a wholly deleted paragraph lists its change in its own paragraph (case 1), where check 53's case 30, of the same
         shape, listed none;
       - two deleted paragraphs in a row, by one author or two, list nothing in either paragraph (cases 2 and 3). Case 3's
         region lists only the first author's deletion;
       - a paragraph whose text is deleted and whose mark isn't lists nothing (case 4); an inserted paragraph lists its
         change (case 5);
       - a deletion or insertion through a mark is listed by its first paragraph (`HEAD14B\r`) and not by the next, in
         every run, where check 53's case 25, built the same way, listed the next paragraph's part too;
       - a content range lists a change cut to it (`NEW5A NEW5B` without the mark), the whole range and the paragraph
         with the mark.
     - **Accepting or rejecting a paragraph's part** (cases 14 and 16) resolves that part only, `BODY14A ` staying a
       deletion, as this package does. Cases 15, 17, 18 and 19 had nothing listed to resolve, in Word as here.
     - **A range across a deleted mark** in Simple Markup reads nothing for the mark: `HEAD6A BODY6B`, `TWO7ATWO7B`, and
       `PLAIN8A\rPLAIN8C` across a wholly deleted paragraph; its paragraphs are listed apart, the deleted one as `''`.
       All Markup reads the marks. **A deleted mark at a cell's end** reads as any cell end (`CELL9A\t`), as here.
     - **A control showing its placeholder** reads as empty, run-level too: its text `''`, its paragraph's
       `PLAIN11A  PLAIN11B`, `placeholderText` `null`. `search` matches the placeholder's text, block and run-level,
       but each hit reads `''`; `PLAIN12A Click` matches, reading `PLAIN12A `. Writing into one replaces the placeholder,
       with `'End'` as with `'Replace'` (`FILL21`, not after it).
     - **Wildcards:** `?` matches the emoji as one character (`x😀y` for `x?y`, `a😀b` for `a?b`), and matches the
       paragraph's mark at its end. `?`, `[!a-z ]` and `[!a-z ]@` match nothing in the Arabic word, and the negated class
       nothing in the emoji: `[!a-z ]` and `[!a-z ]@` find only the mark.
     - **This package since** (2026-10-23, K1 to K6 as recommended; CR-002 sections 58.1 and 32, CR-012 section 18.1):
       the listing kept (K1); a deleted mark read as nothing (CR-012 section 19); a placeholder read as no text and no
       text to search (K3), written over at any location; `?` a code point (K4). `readers.test.mjs` holds the shim's run
       to Word's, the differences listed: cases 2 to 4 (K1), 12 (K3), 13 (K5, K6). Check 59 measures where a wildcard
       matches a mark (K5).
59. Where Word's wildcards match a paragraph mark (CR-002 section 32's K5, 2026-10-23; Jason: "let's do small check 59
   now"). Check 58 found `?` matching a paragraph's mark at its end, in one body paragraph. A wildcard search that finds a
   mark could be followed the way a plain `^p` query is (CR-014 phase 3: `search` finds it, `replaceText` refuses). The
   places `^p` needed care are unmeasured for wildcards: across two paragraphs, a cell's end, a row's end, a mark inside a
   block control. Read only. The file is `__tmp/59/59-wildcard-marks.docx` on the share, as `fixtures/check59/build.mjs`
   writes it (git ignores the `.docx`; `check59Bytes` gives a test the same bytes). The snippet is beside it,
   `check59-script-lab.html` and `check59-script-lab.js` (WordApi 1.6), copies of `fixtures/check59/`'s.
   - **The file:** five regions, `BEGIN<n> <shape>`, its content, `END<n>`. Each content paragraph holds one capital
     letter, so a search for `?` lists every character and mark Word matches:
     - 1 `body`: two body paragraphs, `Q` and `Z`;
     - 2 `cell-only`: a one-cell table whose cell holds `Q`, then a paragraph `Z`;
     - 3 `cell-two`: a one-cell table whose cell holds `Q` and `Z`, then `W`;
     - 4 `cells-rows`: a table of two rows of two cells, `Q` `Z` and `W` `X`, then `Y`;
     - 5 `control`: a block content control holding `Q` and `Z`, then `W`.
   - **The searches**, all with wildcards: across the region's range, `Q?Z`, `Q^13Z`, `Q*Z` and `^13` (Word's wildcard
     code for a paragraph mark); in each paragraph, in the first cell's body, over the table's whole range and in the
     control, `?`, `[!A-Z]` and `^13` (and `Q?Z`, `Q*Z` over the table and the control). Each hit is reported by its
     text.
   - **Steps, in Word 365 only:**
     1. Open `__tmp/59/59-wildcard-marks.docx`. Say whether Word reports a problem or offers a repair.
     2. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/59/check59-script-lab.html` over the HTML tab and `__tmp/59/check59-script-lab.js` over the Script tab.
     3. Click **Run the check**, once. When it says Done, copy the box into a new file `__tmp/59/result.txt`.
     4. Close the document without saving.
   - **What this package gives**, before the run (`readers.test.mjs` runs the snippet through the Word shim): no wildcard
     matches a mark anywhere, as K5 decided. `?` lists each paragraph's letter; a cell's body and the table list their
     letters; `[!A-Z]`, `^13`, `Q?Z`, `Q^13Z` and `Q*Z` find nothing.
   - **Run 2026-10-23 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab).** The JSON is
     `fixtures/check59/59-wildcard-marks-word365.json`; nothing was saved (read only). Word's wildcards match a paragraph
     mark where a plain `^p` does (CR-014 phase 3), and nowhere else:
     - **Matched:** the mark between two body paragraphs (`Q?Z`, `Q^13Z` and `Q*Z` each find `Q\rZ`); a mark inside a
       cell (case 3); the marks inside a block control, its last paragraph's included (case 5). A paragraph's own search
       holds its mark (`?` gives `Q`, `\r`). `^13` is Word's wildcard code for the mark.
     - **Not matched:** a cell's end (case 2's cell gives `Q` alone), a cell's or a row's end in a table of two rows
       (case 4's table gives `Q`, `Z`, `W`, `X`), anything across a table's edge (`Q?Z` finds nothing in cases 2 and 4).
     - **A control's range** omits its last paragraph's mark (case 5's control gives `Q`, `\r`, `Z`), as the paragraph's
       own range doesn't.
     - **This package**: every hit but the marks was Word's, scope for scope. Since K7 (the same day, CR-002 section 32),
       the marks are too: `readers.test.mjs` holds the shim's run to Word's in every scope, and `replaceText` with
       wildcards refuses where a hit holds a mark, as a plain `^p` query is refused.

60. What Office JS's `Font` reports for run properties Word reads leniently (CR-015 D4, 2026-10-24; Jason: "implement
   CR-015"). Read only: nothing is changed. The file is `__tmp/60/60-runprops.docx` on the share, as
   `fixtures/check60/build.mjs` writes it (git ignores the `.docx`; `check60Bytes` gives a test the same bytes). The snippet
   is beside it, `check60-script-lab.html` and `check60-script-lab.js` (WordApi 1.3), copies of `fixtures/check60/`'s.
   - **The file** is 25 of docx4j's probes (`runprops-word365/`, rounds 1 to 6), one after another, each starting with a
     paragraph `PROBE <name>`; every other paragraph is `label: MARKER (end)`, the marker run carrying the property tested
     (on/off values such as `yes`, `FALSE` and `" "`; `w:u` values outside its list; colour names and hex spellings; values
     over styles). Word opened each of these probes without a repair. The three `props-enum` probes are left out: Word
     repairs them, and what they show is its repair's doing.
   - **What it reads:** for each marker, Office JS's `font.bold`, `italic`, `underline`, `strikeThrough`,
     `doubleStrikeThrough`, `superscript`, `subscript`, `color` and `highlightColor`, then `allCaps`, `smallCaps`, `hidden`,
     `emboss`, `engrave`, `outline`, `shadow` and `underlineColor` where the host has them (each read alone, so one it lacks
     reads as an error and the rest still come back). A hidden marker (`w:vanish`) is recorded by its paragraph's text alone.
   - **Steps, in Word 365 only:**
     1. Open `__tmp/60/60-runprops.docx`. Say whether Word reports a problem or offers a repair.
     2. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/60/check60-script-lab.html` over the HTML tab and `__tmp/60/check60-script-lab.js` over the Script tab.
     3. Click **Run the check**, once. When it says Done, copy the box into a new file `__tmp/60/result.txt`.
     4. Close the document without saving.
   - **What this package gives**, before the run (`runprops-word365.test.mjs` runs the snippet through the Word shim):
     each marker reads as the probe's own `Font` does, which is Word's save's reading: an invalid on/off value on, an
     invalid `w:u` no underline, a colour name its hex (`red` `#FF0000`), an invalid colour automatic (`#000000`), an
     invalid highlight none (`null`), a character style's highlight none. A green highlight reads `#008000`, the objects
     package's value for Word's `green`, which Office JS's own is the check's to say. `allCaps` and the rest after
     `highlightColor` are not in the shim's `Font`.
   - **Run 2026-10-24 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab), no repair.** The JSON is
     `fixtures/check60/60-runprops-word365.json`. Office JS reads every marker as the shim does, but Word's green highlight:
     `#00FF00` in Office JS, `#008000` in the objects package's table (docx4j's `Highlight.getHexVal()`), the objects
     package's to fix. `allCaps`, `smallCaps`, `hidden`, `emboss`, `engrave`, `outline` and `shadow` agree with the
     resolver on every marker, and `underlineColor` is the `w:u`'s own colour where it states one (even with no underline
     drawn), else the text's under an underline, else `#000000`. CR-015 section 11.6 has the detail.

61. Word's highlight palette through Office JS, read and written (CR-015 section 12, 2026-10-24; Jason: "write check 61").
   Check 60 found Office JS reading Word's green highlight `#00FF00` where the objects package's table (docx4j's) has
   `#008000`. This measures all sixteen names before that table changes. The file is `__tmp/61/61-highlights.docx` on the
   share, as `fixtures/check61/build.mjs` writes it (git ignores the `.docx`; `check61Bytes` gives a test the same bytes).
   The snippet is beside it, `check61-script-lab.html` and `check61-script-lab.js` (WordApi 1.3), copies of
   `fixtures/check61/`'s.
   - **The file:** a paragraph `READ <name>: R_<name> (end)` for each of the sixteen highlight names, `none` and a character
     shading, the marker run carrying it; then 45 paragraphs `WRITE <n> <value>: W<nn> (end)`, each marker plain (the
     last highlighted yellow, for the null write).
   - **What it does:** reads each READ marker's `font.highlightColor`; sets each WRITE marker's `font.highlightColor` to the
     value its label shows (the names; Word's usual palette; the objects table's other values; `#00ff00`, `Yellow`;
     `#123456` and `#FF8000`; `null`), then reads it back and takes the run's `w:highlight` and `w:shd` from the range's
     `getOoxml()`. **The writes change the document**: close it without saving.
   - **Steps, in Word 365 only:**
     1. Open `__tmp/61/61-highlights.docx`. Say whether Word reports a problem or offers a repair.
     2. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/61/check61-script-lab.html` over the HTML tab and `__tmp/61/check61-script-lab.js` over the Script tab.
     3. Click **Run the check**, once. When it says Done, copy the box into a new file `__tmp/61/result.txt`.
     4. Close the document **without saving**.
   - **What this package gives** (`runprops-word365.test.mjs` runs the snippet through the Word shim): the objects
     package's table. Each name reads its value there (`green` `#008000`, `darkBlue` `#00008B`, `darkYellow` `#FFD700`,
     ...), a value of the table writes its name, and the other 12 values are refused.
   - **Run 2026-10-24 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab), no repair, closed without saving.** The JSON is
     `fixtures/check61/61-highlights-word365.json`. Office JS reads Word's palette (`green` `#00FF00`, `darkBlue`
     `#000080`, `darkCyan` `#008080`, `darkGreen` `#008000`, `darkMagenta` `#800080`, `darkRed` `#800000`, `darkYellow`
     `#808000`, `darkGray` `#808080`, `lightGray` `#C0C0C0`; the bright ones and black and white as the table). Written, a
     name is a CSS colour name (`green` stores `darkGreen`; `darkYellow` is refused) and any hex is stored as the nearest
     highlight by a rule not yet known (`#A9A9A9` `darkGray`, `#123456` `darkCyan`), never as a shading. The shim
     differs in 39 of the 63 cases, all the objects package's table; CR-015 section 12 has the detail.

62. Word's rule for the highlight a written colour becomes (CR-015 section 12.2, 2026-10-24; Jason, through the objects
   session: "write check 62", "so we can see whether there is a simple rule"). Check 61 found Office JS storing any
   `#RRGGBB` as one of the sixteen highlight colours, but not always the nearest. The file is
   `__tmp/62/62-highlight-snap.docx` on the share, as `fixtures/check62/build.mjs` writes it (git ignores the `.docx`;
   `check62Bytes` gives a test the same bytes). The snippet is beside it, `check62-script-lab.html` and
   `check62-script-lab.js` (WordApi 1.3), copies of `fixtures/check62/`'s.
   - **The file:** a title, then 7,225 paragraphs, each holding one colour as six hex digits (a 16-level grid, the grey
     ramp, single-channel ramps, and the neighbourhoods of `#A9A9A9` and `#123456`).
   - **What it does:** sets each paragraph's `font.highlightColor` to its colour, 400 at a time, reads each back, and for
     every 97th takes the stored `w:highlight` from `getOoxml()`. **The writes change the document**: close it without
     saving. It takes a few minutes; the status line counts the batches.
   - **Steps, in Word 365 only:**
     1. Open `__tmp/62/62-highlight-snap.docx`. Say whether Word reports a problem or offers a repair.
     2. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/62/check62-script-lab.html` over the HTML tab and `__tmp/62/check62-script-lab.js` over the Script tab.
     3. Click **Run the check**, once, and wait for Done (about 7,000 colours). Copy the box into a new file
        `__tmp/62/result.txt`. The box holds one long line, so use Ctrl+A in it, then Ctrl+C.
     4. Close the document **without saving**.
   - **What this package gives** (`runprops-word365.test.mjs` runs the snippet through the Word shim): a colour of the
     objects package's table is stored as its name, and any other is refused. What it should do is objects CR-014's H4,
     which the run decides.
   - **Run 2026-10-24 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab), no repair, closed without saving.** The JSON is
     `fixtures/check62/62-highlight-snap-word365.json`. Every colour is stored as one of the sixteen. The grey ramp is four
     bands (`00` black, `40` darkGray, `A0` lightGray, `E0` white), but for `#A9A9A9`, CSS's `darkgray`, which is darkGray. A
     single channel over `00` is its dark colour up to `C0` and its bright one above; a single channel under `FF`, with the
     others at `FF`, is the colour, never white. It is the nearest in RGB for only 61% of the colours. The rule is the
     objects session's to fit (CR-015 section 12.2).


63. Writing into a content control showing its placeholder (CR-002 section 63, 2026-10-24; Jason: "write check 63"). From
   docx4j-python's port of CR-012 section 18 (its CR-016 sections 11.6 and 11.11):
   - Only `ContentControl.insertText` clears `w:showingPlcHdr`, here and there.
   - Every other write into such a control leaves the flag on, so the text it writes reads as no text.
   - `Paragraph.text =` on a paragraph holding a run-level placeholder control drops the control and its binding. That's
     here; docx4j-python refuses the write.

   Nobody has measured Word for any of these. The file is `__tmp/63/63-placeholders.docx` on the share, as
   `fixtures/check63/build.mjs` writes it (git ignores the `.docx`; `check63Bytes` gives a test the same parts). The snippet
   is beside it, `check63-script-lab.html` and `check63-script-lab.js` (WordApi 1.5), copies of `fixtures/check63/`'s.
   - **The file:** a region per case, `BEGIN<n> <op> <shape> <mode>`, one control tagged `c63-<n>` showing Word's own
     placeholder text (`Click or tap here to enter text.`, in the Placeholder Text style, with `w:showingPlcHdr`; no glossary
     part, as check 58's), then `END<n>`.
     - **The shapes:** `block`, a block-level control holding one paragraph; `run`, a run-level control between
       `PLAIN<n>A` and `PLAIN<n>B`; `block-bound` and `run-bound`, the same as `w:text` controls bound to an empty node.
     - **The writes**, each `NEW<n>`, untracked (cases 1 to 26) and tracked (27 to 52):
       - through the control: `insertText` with `'Replace'` and `'End'`, `insertParagraph` at its end;
       - through its `Content` range: `insertText` with `'Replace'` and `'End'`, and `delete()`;
       - through its `Whole` range: `insertText` with `'Replace'` (run-level);
       - through its paragraph, or the paragraph holding it: `insertText` with `'Replace'`, `'Start'` and `'End'`,
         `insertParagraph` `'After'` and `'Before'`, and the `Content` range's `delete()`.
     - **Cases 53 and 54** are written nothing. They hold what this package leaves after a delete: a control showing its
       placeholder with no placeholder text, run-level with an empty `w:sdtContent`, and block-level with an empty
       paragraph.
   - **The questions:**
     - After each write, does the control stay?
     - Does Word clear `w:showingPlcHdr` and drop the Placeholder Text style?
     - Where does the text land, and what do `text`, `placeholderText` and `showingPlaceholderText` read?
     - Under tracking, what is marked?
     - For a bound control, what does its node hold?
     - What do cases 53 and 54 show?
   - **Steps, in Word 365 only:**
     1. Open `__tmp/63/63-placeholders.docx`. Say whether Word reports a problem or offers a repair. Then say what the
        controls in cases 53 and 54 show on the page: nothing, `Click or tap here to enter text.`, or something else.
     2. File > Save As `63-opened-word365.docx`, in the same folder `__tmp/63`, before anything else.
     3. File > Save As again, `63-run-word365.docx` in `__tmp/63`. This is the copy the run changes.
     4. On the Review tab, leave Track Changes off, and set Display for Review to Simple Markup.
     5. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/63/check63-script-lab.html` over the HTML tab and `__tmp/63/check63-script-lab.js` over the Script tab.
     6. Click **Report only**. When it says Done, copy the box into a new file `__tmp/63/result-report.txt`.
     7. Click **Run the check**, once. It turns Track Changes on for each tracked case, and off again after it. When it says
        Done, copy the box into a new file `__tmp/63/result-run.txt`.
     8. Click into the document's text, then File > Save. (Ctrl+S right after the run goes to Script Lab's pane, which has
        the focus, and the document isn't saved.)
     9. File > Close. Open `__tmp/63/63-run-word365.docx` again, and say whether Word reports a problem or offers a repair.
        If it offers one, click Yes, then File > Save As `63-run-word365-repaired.docx` in `__tmp/63`.
   - **What the JSON holds.** For each case:
     - the write's result: the way the control was found, and the text of the range the write returned, or Office JS's
       error as `error: <code>: <message>`;
     - the region after it: its paragraphs' text and its OOXML (the body's content);
     - the control's type, `text`, `placeholderText`, paragraphs and `showingPlaceholderText` (`undefined` where Word
       gives none);
     - for a bound control, its node.
   - **What this package gives**, before the run (`placeholders.test.mjs` runs the snippet through the Word shim and pins
     each case):
     - Only `ContentControl.insertText` clears the flag. Every other write into the control leaves it on, so its text
       reads as no text.
     - The untracked `insertText('Replace')` on the paragraph holding a run-level control (cases 19 and 26) empties the
       control and keeps it.
     - Three defects show (CR-002 section 63):
       - a tracked `insertText` on a bound block control is refused only after it has changed the control (case 47);
       - text written to a run-level control's `Content` or `Whole` range lands outside it (15, 16, 18; 41, 42, 44), and
         deleting its `Content` range deletes nothing;
       - a tracked write into a block control still showing its placeholder reads as text, where untracked it reads as
         none (33, 49).
     - The shim has no `customXmlParts`, so a bound control's node reads `NotImplemented`.
   - **Run 2026-10-24 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab).** No problem opening the file or Word's save
     after the run. The JSONs are `fixtures/check63/63-placeholders-word365-report.json` and `-run.json`. Word's saves are
     in `check63/` with a SHA-1 manifest: the file as opened (saved as `63-placeholders-word365.docx`) and after the run.
     CR-002 section 63.2 has every case. In short:
     - **Opened:** cases 53 and 54 show an empty control, kept as written.
     - **Writes into the control:** every text write that reaches the control's content replaces the placeholder and
       clears `w:showingPlcHdr`, at any location, through the control, its `Content` range or a block control's
       paragraph. The control's `insertParagraph` replaces it but leaves the flag on.
     - **Formatting:** the Placeholder Text style stays in a rich text control (with `w:color w:val="auto"` added for an
       `'End'`, and for a paragraph's `'Start'`), and goes in a bound `w:text` control.
     - **Replaces and deletes:** the paragraph's `insertText(…, 'Replace')` or `Content` delete takes the control, and so
       does a run-level control's `Whole` Replace. Tracked, Word marks the control deleted with `w:customXmlDelRange`,
       and puts the insertion before the deletion. A paragraph inserted beside a block control goes outside it.
     - **Nodes:** a bound control's node stayed empty after every write. Reopened, Word refilled each bound control from
       its node, so cases 21 to 25 and 47 to 51 show empty controls: Word's writes were lost (Jason's screenshots).
     - **This package** differs in 44 of the 52 writes; D1 to D7 are proposed in section 63.3.

64. Check 63's writes on controls in Word's own form, and a bound control's node (CR-002 section 63.4, 2026-10-24; Jason:
   "let's do check 64").
   - **Why:** check 63's controls had no `w:placeholder` and the file no glossary part, so Word had only the content's run to
     format a write from. That may be why the Placeholder Text style stayed, with `w:color w:val="auto"` added for some
     writes (D2). And Word left every bound node empty, then lost the writes when it reopened the file (D7). This check
     asks whether that comes of the placeholder or of the binding.
   - **The file** is `__tmp/64/64-placeholders-word.docx` on the share, as `fixtures/check64/build.mjs` writes it (git ignores
     the `.docx`; `check64Bytes` gives a test the same parts). The snippet is beside it, `check64-script-lab.html` and
     `check64-script-lab.js` (WordApi 1.5), copies of `fixtures/check64/`'s.
     - Every control is in Word's form, as Word saved one in check 34 (`fixtures/revisions/check34/34a-mapped.docx`):
       `w:placeholder` naming `DefaultPlaceholder_-1854013440`, with that file's glossary parts copied unedited.
     - Regions as check 63's, `BEGIN<n> <op> <shape> <mode>` ... `END<n>`, the control tagged `c64-<n>`.
   - **The cases** (52):
     - **1 to 9 and 22 to 30:** a block control showing its placeholder, given check 63's writes (and the control's
       `insertText` at `'Start'`), untracked and tracked.
     - **10 to 14 and 31 to 35:** a run-level control showing its placeholder, the same.
     - **15 to 21 and 36 to 42:** no control in the file. The snippet makes one with `insertContentControl`, with Track
       Changes off: around the region's empty paragraph, or at the point after `PLAIN<n>A`. It reports the control as Word
       made it, then gives it a write.
     - **43 to 52**, untracked: a bound `w:text` control showing its placeholder over an empty node (43 to 45 block, 49 and
       50 run-level), and one bound to a node holding `VALUE<n>`, which it shows (46 to 48 block, 51 and 52 run-level). Each
       gets `insertText(…, 'Replace')` through the control, its `Content` range and, for a block control, its paragraph.
   - **Steps, in Word 365 only:**
     1. Open `__tmp/64/64-placeholders-word.docx`. Say whether Word reports a problem or offers a repair.
     2. File > Save As `64-opened-word365.docx`, in the same folder `__tmp/64`, before anything else.
     3. File > Save As again, `64-run-word365.docx` in `__tmp/64`. This is the copy the run changes.
     4. On the Review tab, leave Track Changes off, and set Display for Review to Simple Markup.
     5. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/64/check64-script-lab.html` over the HTML tab and `__tmp/64/check64-script-lab.js` over the Script tab.
     6. Click **Report only**. When it says Done, copy the box into a new file `__tmp/64/result-report.txt`.
     7. Click **Run the check**, once. It turns Track Changes on for each tracked case, and off again after it. When it
        says Done, copy the box into a new file `__tmp/64/result-run.txt`.
     8. Click into the document's text, then File > Save. (Ctrl+S right after the run goes to Script Lab's pane.)
     9. File > Close. Open `__tmp/64/64-run-word365.docx` again, and say whether Word reports a problem or offers a repair.
        Then say, or screenshot, what cases 43 to 52 show: `NEW<n>`, `VALUE<n>`, or nothing.
   - **What the JSON holds:** check 63's, and for a made case what making the control gave: where it was made, or
     Office JS's error, and the region's OOXML with the control as Word made it.
   - **What this package gives** (`placeholders.test.mjs` runs the snippet through the Word shim and pins each case):
     - The `w:placeholder` and the glossary change nothing: check 63's answers, case for case.
     - A made control shows no placeholder.
     - A bound control showing a value takes every write, and `ContentControl.insertText` writes its node. The shim has no
       `customXmlParts`, so the node isn't read here.
   - **Run 2026-10-24 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab).** No problem opening the file or Word's save
     after the run. The JSONs are `fixtures/check64/64-placeholders-word-word365-report.json` and `-run.json`. Word's saves
     are in `check64/` with a SHA-1 manifest. Reopened, every bound control showed `NEW<n>`. CR-002 section 63.5 has every
     case. In short, in Word's form the placeholder is not content:
     - every write, `insertParagraph` too, takes it away and clears the flag, leaving plain text;
     - under tracking, only the insertion is recorded;
     - every bound node, empty or not, receives the write;
     - Word's own `insertContentControl` makes this form.

     Check 63's oddities all come of the bare form, which docx4j's binding and this package write. This package differs
     in 28 of the 52 writes. The first control blinked during the run; the results are as expected.

65. What Word needs of a placeholder in its own form (CR-002 section 63.7, 2026-10-24; Jason: "write check 65", for D8).
   - **Why:** check 64's controls named Word's default glossary entry, `DefaultPlaceholder_-1854013440`, and the file had
     Word's whole glossary, six parts. D8 has this package write that form; this check settles how little of it Word needs.
   - **The files** are on the share in `__tmp/65/`, as `fixtures/check65/build.mjs` writes them (git ignores the `.docx`;
     `check65Bytes` gives a test the same parts). The snippet is beside them, `check65-script-lab.html` and
     `check65-script-lab.js` (WordApi 1.5), copies of `fixtures/check65/`'s.
     - **`65a-reference-only.docx`:** the controls name the default entry, and the file has no glossary part.
     - **`65b-minimal-glossary.docx`:** a glossary written here, the glossary document part alone. It holds the default
       entry and one of its own text (`Enter the name.`, under a name of Word's kind, 32 hex digits).
   - **The cases**, regions as check 63's, each control tagged `c65-<n>`:
     - **In both files:**
       1. the control's `insertText('NEW1', 'End')`, which in the bare form kept the style;
       2. its `insertParagraph`, which in the bare form left the flag on;
       3. a tracked `insertText(…, 'Replace')`, which in the bare form marked the placeholder deleted;
       4. a bound `w:text` control's `insertText` over an empty node;
       5. a run-level bound control's `Content` range `insertText`;
       6. setting `placeholderText` to `Enter the name 6.`, which Word keeps in a glossary;
       7. and 8. a control and a bound control left alone;
       9. a control made by `insertContentControl`, nothing written;
       - the last case (10 in 65a, 14 in 65b): a bound `w:text` control in the bare form, check 63's, with no
         `w:placeholder`. The snippet writes nothing there; you type into it by hand (docx4j CR-063's question: does a
         typed value reach the node, where Office JS's didn't).
     - **In 65b only:** a control showing the glossary's own entry (10 written, 11 left alone), and one naming an entry
       the glossary hasn't (12 written, 13 left alone).
   - **Steps, in Word 365 only.** First `65a-reference-only.docx`:
     1. Open `__tmp/65/65a-reference-only.docx`. Say whether Word reports a problem or offers a repair, and what case 7's
        control shows on the page.
     2. File > Save As `65a-opened-word365.docx`, in the same folder `__tmp/65`.
     3. File > Save As again, `65a-run-word365.docx` in `__tmp/65`. This is the copy the run changes.
     4. On the Review tab, leave Track Changes off, and set Display for Review to Simple Markup.
     5. Open Script Lab (Insert > Get Add-ins > Script Lab, then Script Lab > Code > New snippet). Paste the share's
        `__tmp/65/check65-script-lab.html` over the HTML tab and `__tmp/65/check65-script-lab.js` over the Script tab.
     6. Click **Report only**. When it says Done, copy the box into a new file `__tmp/65/result-a-report.txt`.
     7. Click **Run the check**, once. When it says Done, copy the box into a new file `__tmp/65/result-a-run.txt`.
     8. Click into case 10's control (it shows `Click or tap here to enter text.`), type `TYPED10`, then click outside it.
        Then File > Save. (Ctrl+S right after the run goes to Script Lab's pane.)
     9. File > Close. Open `__tmp/65/65a-run-word365.docx` again. Say whether Word reports a problem or offers a repair,
        and what cases 4, 5 and 10 show: `NEW4`, `NEW5` and `TYPED10`, or nothing.

     Then the same for `__tmp/65/65b-minimal-glossary.docx`, with `b` for `a` in every name: `65b-opened-word365.docx`,
     `65b-run-word365.docx`, `result-b-report.txt`, `result-b-run.txt`. At step 1, also say what cases 11 and 13 show on
     the page. At steps 8 and 9 the typed case is 14: type `TYPED14`.
   - **What this package gives** (`placeholders.test.mjs` runs the snippet through the Word shim on both files and pins
     each case): check 63's answers again, whatever the glossary holds. `placeholderText` reads the control's runs, the
     custom control's `Enter the name.` among them.
   - **Run 2026-10-25 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab).** No problem opening either file or Word's
     saves. The JSONs are `fixtures/check65/65a-reference-only-word365-report.json` and `-run.json`, and the same for
     `65b-minimal-glossary`. Word's saves are in `check65/` with a SHA-1 manifest. Reopened, 65a showed `NEW4`, `NEW5`
     and `TYPED10`, and 65b `NEW4` and `NEW5`. CR-002 section 63.8 has every case. In short:
     - the reference alone gives Word's form while the file is open: 65a's writes all give check 64's answers;
     - Word's save drops a reference whose entry no glossary holds, so the file is back in the bare form;
     - Word keeps a glossary of the document part alone, but its entries lose the Placeholder Text style, which the
       glossary's own styles lack;
     - `placeholderText` set through Office JS is a new entry, named with 32 hex digits;
     - typed by hand, 65a's bare-form bound control took `TYPED10` and wrote its node. In 65b the typing at step 8 went
       into case 10. `TYPED14` typed into case 14 after reopening (on screen inside the placeholder text) was saved over
       the share's `65b-run-word365.docx`; the copy is `check65/65b-reopened-typed-word365.docx`. It holds `TYPED14`
       alone, in the control and the node, and Word gave that control, and four others, glossary entries of their own.

     This package differs in writes 2, 3, 5 and 9 of each file, and reads `placeholderText` from the runs.

66. How much glossary a placeholder needs (CR-002 section 63.10, 2026-10-25; Jason: "yes", to writing check 66, for D9).
   - **Why:** check 65's 65b had a glossary of the document part alone. Word kept it, but its entries lost the
     Placeholder Text style, which the glossary's own styles lacked. And a reopened save, typed into, moved controls off
     the default entry. D9 chooses what this package writes where a file has no glossary: (a) the document part and a
     styles part, or (b) Word's six parts. This check gives each to Word.
   - **The files** are on the share in `__tmp/66/`, as `fixtures/check66/build.mjs` writes them (git ignores the `.docx`;
     `check66Bytes` gives a test the same parts). The snippet is beside them, `check66-script-lab.html` and
     `check66-script-lab.js`, check 65's with check 66's names.
     - **`66a-glossary-with-styles.docx`:** D9 (a). The glossary document part holds the default entry in the Placeholder
       Text style, and a styles part holds that style and the one it is based on, as Word defines them.
     - **`66b-word-glossary.docx`:** D9 (b). Check 34's glossary, Word's six parts, as check 64 had it.
   - **The cases**, the same in both files, each control tagged `c66-<n>`:
     1. a control showing its placeholder, left alone;
     2. a bound control showing its placeholder, typed into by hand after reopening;
     3. a control holding `VALUE3`; the snippet deletes its content (does Word bring the placeholder back, grey?);
     4. a control the snippet makes with `insertContentControl` (is its placeholder grey?);
     5. `placeholderText` set to `Enter the name 5.`;
     6. a bound control in the bare form (no glossary reference), typed into by hand after reopening.
   - **Steps, in Word 365 only.** First `66a-glossary-with-styles.docx`:
     1. Open `__tmp/66/66a-glossary-with-styles.docx`.
     2. File > Save As `66a-opened-word365.docx`, in the same folder `__tmp/66`.
     3. File > Save As again, `66a-run-word365.docx` in `__tmp/66`. This is the copy the run changes.
     4. In Script Lab, replace check 65's snippet, or start a new one: paste `__tmp/66/check66-script-lab.html` over the
        HTML tab and `__tmp/66/check66-script-lab.js` over the Script tab.
     5. Click **Report only**. When it says Done, copy the box into a new file `__tmp/66/result-a-report.txt`.
     6. Click **Run the check**, once. When it says Done, copy the box into a new file `__tmp/66/result-a-run.txt`.
     7. Say whether the placeholders in cases 3 and 4 are grey or black. Then File > Save. (Ctrl+S right after the run goes
        to Script Lab's pane.)
     8. File > Close. Open `__tmp/66/66a-run-word365.docx` again. Straight away, typing nothing, File > Save As
        `66a-reopened-word365.docx` in `__tmp/66`.
     9. Click into case 2's control (it shows `Click or tap here to enter text.`) and type `TYPED2`. Then click into case
        6's control and type `TYPED6`, and click outside it. Say whether the typed text replaced each placeholder or went
        inside it. Then File > Save As `66a-typed-word365.docx` in `__tmp/66`, and File > Close.

     Then the same for `__tmp/66/66b-word-glossary.docx`, with `b` for `a` in every name: `66b-opened-word365.docx`,
     `66b-run-word365.docx`, `result-b-report.txt`, `result-b-run.txt`, `66b-reopened-word365.docx`,
     `66b-typed-word365.docx`. The typed cases are 2 and 6 again.
   - **What this package gives** (`placeholders.test.mjs` runs the snippet through the Word shim on both files and pins
     each case): the same in both files. The control emptied (3) and the control made (4) show no placeholder.
   - **Run 2026-10-25 (Jason; Word 365, 16.0.20326.20232, PC; Script Lab).** No problem opening either file or any of
     Word's saves. The JSONs are `fixtures/check66/66a-glossary-with-styles-word365-report.json` and `-run.json`, and the
     same for `66b-word-glossary`. Word's eight saves are in `check66/` with a SHA-1 manifest. Cases 3 and 4 showed their
     placeholders grey, and the typed text replaced each placeholder, in both files. CR-002 section 63.11 has it all. In
     short, the two files gave the same answers at every step:
     - Word kept 66a's entry in the style and completed the glossary to six parts;
     - every reference stayed through the run, a reopening and typing;
     - the bare-form control typed into was given no entry, and both typed nodes were written.

     So D9 (a), the document and styles parts, is enough (decided, Jason, 2026-10-25). This package differs in cases 3 and 4.
