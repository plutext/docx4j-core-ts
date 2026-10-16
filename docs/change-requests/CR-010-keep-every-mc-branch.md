# CR-010: Keep every branch of `mc:AlternateContent` (docx4j CR-021's policy)

**Status:** Proposed 2026-10-17, at Jason's request ("Shouldn't we be preserving all branches of
mc?"). Not scheduled; each phase waits for Jason's go. The policy is docx4j's CR-021 (done
2026-09-19, in 17.2.0). Phase 2 needs work in the objects package, as a CR there.

## 1. Why

This package resolves `mc:AlternateContent` when it unmarshals a part (CR-001 section 5.6). The
first `mc:Choice` whose `Requires` namespaces are all in `UNDERSTOOD_NAMESPACES` replaces the
element, else its `mc:Fallback` does. The branch not taken is gone from that part, and a save that
re-marshals the part writes only the branch taken. Untouched parts are copied as bytes, so they
keep every branch.

CR-001 section 5.6 says this "is docx4j's behaviour (its preprocessor)". That stopped being true
with docx4j CR-021 (`docs/developer/change-requests/CR-021-mc-alternate-content-policy.md`):

- docx4j keeps the element wherever it occurs, and saves every branch.
- Readers see one branch, the one `McSelection` selects (`McMode.READ`, `TraversalUtil`'s
  default since 17.2.0).
- Mutators (data binding, field update, find and replace, ...) see every branch (`McMode.ALL`), so
  an edit lands in each, and the saved document says the same thing whichever branch a reader
  understands.

So this package now departs from docx4j here. It costs three things:

1. **Every Fallback** is lost from any part that is re-marshalled: a text box's VML, a chart's
   fallback picture (CR-009 section 8, `c36d5f2`), an equation's picture in PowerPoint. A consumer
   that doesn't understand the Choice then has nothing to show.
2. **Two outright losses**, where the Choice's content is bound but its parent doesn't accept it
   once promoted (CR-001 sections 21 and 21.1):
   - Excel's `x15ac:absPath` in the workbook;
   - `c14:style` in a chart part. Its Fallback `c:style` goes too, so a re-marshalled chart loses
     its style altogether.
3. **Parity.** `test/xpath-parity.test.mjs` excludes the four fixtures with `mc:AlternateContent`
   (`loadAndSave.docx`, `numbering-stories-coretests.docx`, `numbering-stories.docx`,
   `tables-textbox-in-cell.docx`). docx4j's tree keeps the element and this package's doesn't, so
   their selections can't be compared.

CR-001 section 15.4 recorded why the default stayed resolve-on-load once keeping became lossless
(objects 0.1.5): a kept branch's content is mostly DOM, which the typed views don't enter. It left
the decision "to revisit when a consumer needs a byte-faithful re-marshal of both branches".
Section 21 said a fix "is a CR of its own, and the measurement above is what it should start from".
This is that CR.

## 2. Measured (2026-10-17)

Corpus: the 459 files the release checks hash (this package's `test/fixtures/` and docx4j's
`docx4j-core-tests/src/test/resources/charts/`, both recursive), objects 0.3.2, jsonix 3.4.1.
241 of them hold 286 `mc:AlternateContent` elements, 2 of them nested.

**What resolve-on-load takes and loses.** Every element's Choice is understood, so the Choice is
taken every time.

| part | host | `Requires` | elements | lost when the part is re-marshalled |
|---|---|---|---|---|
| `charts/chartN.xml` | `c:chartSpace` | `c14` | 227 | the Fallback's `c:style`, and the Choice's `c14:style` with it (section 1 item 2) |
| `xl/workbook.xml` | `workbook` | `x15` | 31 | no Fallback; the Choice's `x15ac:absPath` (section 1 item 2) |
| `word/document.xml` | `w:r` | `wps` | 9 | the Fallback (`w:pict`) |
| `word/document.xml` | `w:r` | `cx1` | 2 | the Fallback (a picture of the chart) |
| `slides/slideN.xml` | `p:spTree` | `cx1` | 2 | the Fallback (a picture of the chart) |
| `drawings/drawingN.xml` | `xdr:twoCellAnchor` | `cx1` | 4 | the Fallback (a "not available" shape) |
| `drawings/drawingN.xml` | `xdr:wsDr`, `xdr:twoCellAnchor`, `xdr:oneCellAnchor` | `a14` | 4 | the Fallback |
| `slides/slideN.xml` | `p:spTree` | `a14` | 1 | the Fallback (a picture of the equation) |
| `drawings/drawingN.xml` | `xdr:twoCellAnchor` | `tsle`, `sle15` | 2 | the Fallback (an "Excel 2010 or later" box, CR-004 section 6) |
| `worksheets/sheetN.xml` | `worksheet`, `controls` | `x14` | 4 | nothing: no Fallback, and the controls are typed where CR-004 phase B wants them |

**What keeping costs, today.** Loaded with `mcePreprocess: false`, unmarshalled and saved, all
281 parts that hold the element write every `mc:AlternateContent` back whole. The check compared
namespaces, names, attributes and children, leaving prefix declarations aside. So keeping loses
nothing. What it loses is reach: inside a kept branch the model holds

| typed | DOM |
|---|---|
| `c14:style` 227, `x15ac:absPath` 31 | `c:style` 227, `w:drawing` 13, `w:pict` 9, `xdr:sp` 9, `xdr:graphicFrame` 7, `p:graphicFrame` 2, `p:pic` 2, `p:sp` 2, `controls` 2, `xdr:twoCellAnchor` 2 |

A `w:drawing` in a kept branch is DOM, so `InlinePicture`, text, search and the rest of the
content API can't see it. docx4j's JAXB types it, because its classes bind `w:drawing`, `w:pict` and
`w:tbl` as root elements (`@XmlRootElement`), and the lax wildcard types any root element. The
objects mapping declares those three only within their parents (`scope: '.R'`, `'.Body'`, ...),
while `w:p` and `w:r` are global, so the wildcard leaves them DOM.

## 3. Proposal (for Jason's decision)

1. **Keep the element wherever it occurs**, as docx4j does. The tree holds
   `org_docx4j_mce.AlternateContent`, and every branch is saved. Resolve-on-load stays available as
   a load option, for a caller that wants one flat tree.
2. **One selection rule for reading**: the objects package's `mcBranchOf`, its port of docx4j's
   `McSelection`, which `textOf` already uses. Which branch it should select by default is decision
   D1 (section 5).
3. **Two traversal modes, as docx4j's `McMode`.**
   - Readers (text, search, lists, `outline`, the effective formatting reads, `select`) walk the
     selected branch.
   - Mutators walk every branch: `replaceText`, content-control and data-binding updates
     (`applyBindings`, the XHTML and picture passes), tracked-change accept and reject, and
     anything that removes content. Each walker in `src/` is classified, as docx4j CR-021 section
     8.3 classified its 51.
4. **Typed branches.** Content the content API reads must be typed inside a kept branch. There are
   two routes, both in the objects package; phase 2 chooses:
   - the generated mapping declares, as global elements, what docx4j binds as root elements
     (`w:drawing`, `w:pict`, `w:tbl`, ...), so the lax wildcard types them at load;
   - or the facade types a DOM branch child in its host's container at load, the way objects CR-002
     unmarshals a fragment.
5. **Save writes every branch.** An edit made through a mutator lands in every branch, so a
   Fallback is never left saying something the Choice doesn't. Content the API inserts is written as
   given.

## 4. Phases

0. **The audit** (half a day):
   - each walker in `src/` classified as reader, mutator or indifferent;
   - what the editor (`../docx4j-ts-editor`) reads through the tree, since its views assume one
     flat tree today;
   - which of the four excluded XPath fixtures come back once the element is kept, and what the
     parity goldens (lists in text boxes) need from D1.
1. **Keep where nothing reads the content** (half a day). The preprocessor leaves the element alone
   where its host is `c:chartSpace` or the SpreadsheetML workbook. The model binds the element in
   both, and nothing in this package reads `c14:style` or `absPath`. This ends section 1 item 2's
   two losses, 227 and 31 elements in the corpus, with no change to the content API.
   - The worksheet's `x14` controls stay resolved until phase 3, since CR-004 phase B reads them
     typed.
   - Test: a chart part's `c14:style` and `c:style`, and the workbook's `absPath`, survive
     `unmarshalAll` and a save.
2. **Typed branches**, in the objects package: a CR there for section 3 item 4's route. Its test is
   this CR's survey: every DOM entry in section 2's second table that the content API reads comes
   back typed, and every part still writes back whole.
3. **The content API over kept branches** (three to five days):
   - the walkers through `mcBranchOf` (readers) and every branch (mutators);
   - `Body.paragraphs` and the story walks deciding how a text box's paragraphs count, as docx4j's
     READ visits them once;
   - `InlinePicture`, search, `replaceText`, tracked changes, bindings.
4. **The default flips to keep** (a day, plus the editor's side): `mcePreprocess` defaults to false.
   - The four XPath fixtures rejoin `xpath-parity.test.mjs`.
   - CR-001 section 5.6 is corrected.
   - The editor takes the release in step.

The estimates are rough, and phase 2's depends on the route.

## 5. Decisions (for Jason)

- **D1: which branch a reader takes by default.**
  - This package and `mcBranchOf` take the first Choice whose prefixes are all understood, which is
    Word's rule (ECMA-376 Part 3, 10.2.1).
  - docx4j's `McSelection` takes a Choice only for prefixes listed in `docx4j.jaxb.mc.preferChoice`,
    and that list is empty by default, so docx4j reads the Fallback. Its reasons (CR-021 section
    8.5): preferring `wps` made two documents of its corpus worse, and its FO exporter doesn't draw
    an inline `wps` text box.
  - Keeping Word's rule keeps today's answers and the content API's Office JS shape. Following
    docx4j's default would make the trees agree for parity.
  - Recommended: Word's rule, with the understood set as an option. The parity harness would then
    run docx4j with `preferChoice` set to match.
- **D2: mutators edit every branch** (docx4j's ALL). Recommended. The alternative leaves a stale
  Fallback.
- **D3: phase 1 alone, now?** It fixes the two outright losses without touching the content API.
  Recommended either way.

## 6. Risks

- **The tree's shape changes for consumers.** An `AlternateContent` appears where the resolved
  content stood (`P.content`, `R.content`, the drawing anchors). Code that walks the tree itself,
  the editor's first, has to go through `mcBranchOf`, or through the content API.
- **A misclassified walker:** a mutator walking one branch leaves the others stale; a reader
  walking all of them counts a text box twice (the bug that started docx4j CR-021).
- **A kept branch is lossless only while the model binds its whole content** (docx4j CR-021
  section 8.6 item 10). Measured lossless for all 281 parts today; section 2's survey is re-run
  before the default flips.

## 7. References

- docx4j CR-021 (`VERSION_17_3_3`, `docs/developer/change-requests/`), `org.docx4j.jaxb.McSelection`,
  `McMode`; CR-024 (`mc:AlternateContent` in DrawingML).
- This package: CR-001 sections 5.6, 15.4, 17.5, 21 and 21.1; CR-004 section 6 (the controls, the
  slicers); CR-009 section 8 (the chartex namespaces); `src/opc/mce/`.
- The objects package: `builders/wml`'s `mcBranchOf`, `childValues` and `textOf`; its CR-002 (fragments
  unmarshalled in their container).
- The survey script is not committed. It loads each file twice, resolved and with
  `mcePreprocess: false`, and compares every element's normalised form; phase 0 makes it a test.
