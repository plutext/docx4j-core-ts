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

The images in `content-c.test.mjs` are base64 constants rather than files: a 4 x 3 PNG at 96 dpi
(a real one, deflated with `node:zlib`), a 2 x 2 GIF87a, a 2 x 2 24-bit BMP at 3780 px/m, and a
JPEG header (SOI, a JFIF APP0 at 300 dpi, SOF0, EOI) for the header reader.

## Parity fixtures (`fixtures/parity/`, CR-001 Phase B step 0, 2026-09-19)

Thirty-seven documents the Java harness (`test/java/`, step 1) turns into goldens under
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

## Numbering indent fixtures (`fixtures/ind/`, CR-001 Phase B step 3)

The seven flat OPC documents of docx4j's `ListNumberIndTest`, copied unchanged from
`docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/ind/` at `VERSION_17_1_1`
(Apache-2.0): `abstract_style_with`, `abstract_style_without`, `abstract_nostyle_ppr`,
`abstract_nostyle_noppr`, `override_nostyle_ppr`, `abstract_style_ind_only` and
`abstract_style_basedon`. Each is one `w:num` 1 whose level 0 states its indent, or its linked
`w:pStyle` does, or neither; `numbering.test.mjs` asserts the `w:ind/@w:left`
`NumberingDefinitionsPart.getIndOf('1', '0')` answers (2880, 12, 13, none, 23, 11, 31).

## Parity goldens (`golden/`, CR-001 Phase B step 1)

What docx4j answers for every fixture above and for the eight `.docx` in `fixtures/`, written
by the Java harness in `java/` and read by `parity.test.mjs`. 45 goldens, harness version 3,
from docx4j `VERSION_17_1_1` at `7fba7a150`; the resolver, the emulator and the selector
reproduce every recorded answer. `golden/README.md` is the review guide (counts per golden, and
what was found producing them); `java/README.md` says how to build docx4j, run the harness and
read a golden. `.github/workflows/parity.yml` reruns it weekly against docx4j's head and opens
a pull request when an answer changes.

`golden/xpath/` holds CR-006's oracle, made by the same project's `XPathHarness` and read by
`xpath-parity.test.mjs`: what docx4j's `getJAXBNodesViaXPath` selects for 47 expressions over each
fixture's main document part, from the docx4j 17.2.1 release. `selectObjects` returns the same hits
in the same order, with objects of the same classes, on all 43 fixtures that have hits.

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
   re-marshal, so Excel does not need the wrapper and resolve-on-load stands (CR-004 section 7). Use
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

   **Part 1 run 2026-10-03 (current Word), `fixtures/revisions/check35/35a-word15.docx`; part 2 (Word
   2010, revised above) to come.** Jason's notes, then the file:
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
