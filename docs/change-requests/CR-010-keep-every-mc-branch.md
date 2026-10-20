# CR-010: Keep every branch of `mc:AlternateContent` (docx4j CR-021's policy)

**Status:** Proposed 2026-10-17, at Jason's request ("Shouldn't we be preserving all branches of
mc?"). Phase 1 implemented 2026-10-17 (section 9; D2, D3 and D4 decided by Jason that day, and
D1 settled by the text extraction paper's D4). Phase 2 implemented 2026-10-18: the objects
package's CR-009, released in its 0.4.0, which this package now requires (section 10). Phase 0
implemented 2026-10-18 (section 11): the walker audit, the editor's exposure, the XPath and parity
measurements, and the survey as a test in `mce.test.mjs`; D5 and D6 decided by Jason that day,
and D5 implemented (section 11.8).
Phase 3 implemented 2026-10-18 at Jason's request ("do CR-010 phase 3"; section 13). Phase 4
implemented 2026-10-18 at Jason's request ("implement CR-010 phase 4, coordinating with the editor as
appropriate"; section 14): load keeps the element by default. It needs `@docx4j/jsonix` 3.4.2,
published the same day, which objects 0.4.1 requires (section 14.4). This package requires objects
0.4.2 since 2026-10-19, which also declares the prefixes a kept `mc:Choice`'s `Requires` names
(section 14.8). The read-time warning moved to CR-012, the text-model CR, on 2026-10-19 (section 14.7). Phases 0 to 4 **released in 0.5.0** (2026-10-19). Section 14.9 (2026-10-20, found by the editor): a `Requires` prefix the prefix table lacks loses
its declaration on save, an objects package defect, sent there as a CR at Jason's word. Section 15
(2026-10-20, the editor's ED-009 question 2): the `Requires` gate alone picks the readers' branch
for a Choice whose content is in bound namespaces, taken as the editor asked (Jason: "do as the
editor asks re Q1, Q2"); proposed, not scheduled. The policy is docx4j's CR-021 (done
2026-09-19, in 17.2.0).

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
3. **Parity.** `test/xpath-parity.test.mjs` excluded, until D5 (section 11.8), the four fixtures
   with `mc:AlternateContent` (`loadAndSave.docx`, `numbering-stories-coretests.docx`,
   `numbering-stories.docx`, `tables-textbox-in-cell.docx`). docx4j's tree keeps the element and
   this package's doesn't, so their selections can't be compared.

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
6. **One selection rule in both selectors**, the preprocessor's and `mcBranchOf`: the text
   extraction paper's D4, which Jason accepted on 2026-10-17 (section 7). Today:
   - **Namespaces.** The preprocessor resolves each `Requires` prefix to its URI at the Choice.
     `mcBranchOf` compares prefixes with the conventional prefix table. D4 item 1 wants URIs, with
     prefixes only where the declarations in force are not kept; that is an objects change.
   - **An empty `Requires`** is not understood by either, as D4 item 3 says. Nothing to do.
   - **Choices only, none understood.** `mcBranchOf` yields the first Choice (D4 item 4, a
     departure confined to reading). The preprocessor removes the element, as ISO does. Section 5
     D4 says what load should do instead.
   - **The understood set** is D4 item 2's: one list published once, as a resource in docx4j as
     the strict fixtures are, with `UNDERSTOOD_NAMESPACES` asserted against it by a test. That
     doesn't mean shrinking this package's set: its ~110 URIs are the widest of the three engines
     and the natural first content. Additions (cx1 and cx2, CR-009 section 8) go to the published
     list.

## 4. Phases

0. **The audit** (half a day; implemented 2026-10-18, section 11):
   - each walker in `src/` classified as reader, mutator or indifferent;
   - what the editor (`../docx4j-ts-editor`) reads through the tree, since its views assume one
     flat tree today;
   - which of the four excluded XPath fixtures come back once the element is kept, and what the
     parity goldens (lists in text boxes) need from D1.
1. **Keep where nothing reads the content** (half a day; implemented 2026-10-17, section 9). The
   preprocessor leaves the element alone where its host is `c:chartSpace` or the SpreadsheetML
   workbook. The model binds the element in both, and nothing in this package reads `c14:style` or
   `absPath`. This ends section 1 item 2's two losses, 227 and 31 elements in the corpus, with no
   change to the content API.
   - The worksheet's `x14` controls stay resolved until phase 3, since CR-004 phase B reads them
     typed.
   - Test: a chart part's `c14:style` and `c:style`, and the workbook's `absPath`, survive
     `unmarshalAll` and a save.
   - Added 2026-10-17, from the paper's D4 (section 7): an element nothing can be selected from
     is kept whole, not removed (D4 below).
   - Also added: the `mce.branches_differ` warning in its load-time form. Before the preprocessor
     resolves an element, it compares the normalised `w:t` text of the branch it takes with the
     branches it drops, and records the part and the branch taken on a load report that readers
     surface. That report is new API, shaped in this phase. This form has no revision view and
     no hidden-text rule, and needs neither to answer "does the other branch say something
     different". The read-time form, with the read rules, comes with phase 3. It needs no
     phase 4. (Phase 3 did not carry it. Phase 4 keeps the load-time form, run without resolving
     (section 14.1), and the read-time form moved to CR-012 (section 14.7).)
2. **Typed branches**, in the objects package: a CR there for section 3 item 4's route (objects
   CR-009, released in its 0.4.0 on 2026-10-18; section 10). Its test is
   this CR's survey: every DOM entry in section 2's second table that the content API reads comes
   back typed, and every part still writes back whole.
3. **The content API over kept branches** (three to five days):
   - the walkers through `mcBranchOf` (readers) and every branch (mutators);
   - `Body.paragraphs` and the story walks deciding how a text box's paragraphs count, as docx4j's
     READ visits them once (section 12: D7 and D8);
   - `InlinePicture`, search, `replaceText`, tracked changes, bindings.
   - Implemented 2026-10-18 (section 13). Search, `replaceText` and bindings reach no text box in
     either load, and go to the text-model CR with the text-box view (section 13.4).
4. **The default flips to keep** (a day, plus the editor's side): `mcePreprocess` defaults to false.
   - The four XPath fixtures rejoin `xpath-parity.test.mjs`. Done early, by D5 (section 11.8).
   - CR-001 section 5.6 is corrected.
   - The editor takes the release in step.
   - Implemented 2026-10-18 (section 14), with the warning kept by a pass that doesn't resolve, and
     the editor's side on its branch `mce-keep`. It needs `@docx4j/jsonix` 3.4.2 (section 14.4).

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
  - **Settled 2026-10-17** by the text extraction paper's D4, which Jason accepted as written
    (`docx4j-portfolio` 2d8fd36). D4 applies the ISO rule on URIs to every reader in all three
    engines; it is Word's rule. docx4j will select by purpose: text by the rule, while drawing
    keeps the Fallback default until docx4j's FO output draws an inline `wps` box. Until docx4j's
    split lands, the parity harness runs docx4j with `preferChoice` set to match, as recommended
    above. Word's evidence is in section 7.
- **D2: mutators edit every branch** (docx4j's ALL). Recommended. The alternative leaves a stale
  Fallback. The paper assumes it throughout, and docx4j's CR-043 lister already reads every
  branch. **Decided 2026-10-17** (Jason: "D2: yes").
- **D3: phase 1 alone, now?** It fixes the two outright losses without touching the content API.
  Recommended either way. The paper takes no position on timing. Phase 1 now also carries the
  two additions of 2026-10-17. **Decided 2026-10-17** (Jason: "D3: start phase 1 now, with the
  two additions above"); implemented, section 9.
- **D4: what load does with an element nothing can be selected from** (Choices only, none
  understood, no Fallback; added 2026-10-17). The decided rule (D4 item 4) is that a reader takes
  its first Choice, and the portfolio session proposed the preprocessor do so too, as
  `mcBranchOf` does. Measured the same day, on a `w:r` holding such an element with a text box in
  an unknown namespace, saved after an edit:

  | load | `getBody()` | the paragraph's text | the save after an edit |
  |---|---|---|---|
  | removes it (today, ISO) | opens | excludes the box | the run empty: the element is lost |
  | promotes the first Choice | **throws** `PartUnmarshalException`: the run's property "does not allow DOM content" | | |
  | keeps the element whole | opens | excludes the box | the element back whole |

  - Promoting fails because the preprocessor's output is what the model unmarshals and a save
    writes, so promoting would make the writer rely on the reading departure, which D4 item 4
    rules out.
  - Keeping is lossless today. The text becomes readable when phase 3 reads kept branches
    through `mcBranchOf`.
  - Recommended: keep the element whole, in phase 1. With this package's understood set the
    case is rare: the corpus has none, and Word's Choice-only boxes require `wps`, which is
    understood. It arises under a caller's narrower `understood` set.
  - **Decided 2026-10-17**, with D3: keeping it is one of phase 1's two additions.
- **D5 and D6**, from phase 0, are in section 11.8: the four XPath fixtures rejoining now, loaded
  kept, and the parity harness's `MC_PREFER_CHOICE` refreshed. Both decided 2026-10-18.
- **D7 and D8**, for phase 3, are in section 12.5: whether `Body.paragraphs` holds a text box's
  paragraphs, and where the text-box view goes. Both decided 2026-10-18, subject to check 42 in Word.

## 6. Risks

- **The tree's shape changes for consumers.** An `AlternateContent` appears where the resolved
  content stood (`P.content`, `R.content`, the drawing anchors). Code that walks the tree itself,
  the editor's first, has to go through `mcBranchOf`, or through the content API.
- **A misclassified walker:** a mutator walking one branch leaves the others stale; a reader
  walking all of them counts a text box twice (the bug that started docx4j CR-021). The paper
  finds the same duplication in other tools: Gemini's app and docx2txt-based loaders emit every
  branch (its sections 7.2 and 7.5).
- **A kept branch is lossless only while the model binds its whole content** (docx4j CR-021
  section 8.6 item 10). Measured lossless for all 281 parts today; section 2's survey is re-run
  before the default flips. Re-run 2026-10-18 over 483 files: all 292 parts whole (section 14.3).

## 7. The text extraction paper (2026-10-17)

The portfolio's `docs/text_extraction_behaviour_model.md` (written 2026-10-17; its section 9
decisions were taken the same day, Jason accepting every recommendation, `docx4j-portfolio`
2d8fd36) defines one text model for docx4j, this package and python. It names this CR twice:
D1 is the rule it wants in all three engines, and phase 3 fills its text-box address. What it
decided, and what it found, as it bears on this CR:

- **Decided (its D4):**
  - the ISO rule on namespace URIs, for every reader;
  - one understood set, published once;
  - an empty `Requires` is not understood;
  - a Choice-only element nobody understands yields its first Choice, for reading only;
  - never both branches;
  - the `mce.branches_differ` warning.

  These settle D1 here. Section 3 item 6 and section 5 D4 apply them.
- **Word's evidence** (Word 365 build 16.0.20326, Jason, 17 October 2026; the paper's section 8):
  - Re-saving the extraction probe, Word resolved every non-shape `AlternateContent` to the branch
    the ISO rule selects, in all eight cases. That includes a rebound prefix (matched by URI) and
    a Choice requiring two prefixes, one unknown (the Fallback was taken).
  - `Range.Text` over `StoryRanges` gives a text box as its own story holding only the drawn
    branch's text. No Fallback text is in any story unless the Fallback is the branch drawn.
- **Fallbacks on Word's save** (scripted rounds by docx4j's session and Jason, in docx4j CR-021
  section 8.5, edited there and not yet committed when this was written):
  - Word keeps an authored Fallback unless it edits the shape, or the document holds an
    `altChunk` anywhere, in which case every Fallback is regenerated.
  - A bare `wps` box gains a Fallback on every save.
  - So in a Word-saved file whose boxes weren't edited, and with no `altChunk`, the Fallbacks are
    still the author's. A reader on the Fallback can read text Word never showed. That is the
    plainest case for phase 4's flip.
- **The text-box address** `body/3/textbox:0/1` (the paper's section 4.3) is a proposal, not
  decided. The address grammar is shared with python's CR-003 section 3.4 and the editor's
  ED-001, so phase 3 agrees a form with those sessions before it lands.
- **The text CR.** The paper's section 10 proposes a core-ts text CR (`getText`, text boxes after
  this CR's phase 3). It consumes this CR's selection rule and doesn't define one.

## 8. References

- docx4j CR-021 (`VERSION_17_3_3`, `docs/developer/change-requests/`), `org.docx4j.jaxb.McSelection`,
  `McMode`; CR-024 (`mc:AlternateContent` in DrawingML).
- This package: CR-001 sections 5.6, 15.4, 17.5, 21 and 21.1; CR-004 section 6 (the controls, the
  slicers); CR-009 section 8 (the chartex namespaces); `src/opc/mce/`.
- The objects package: `builders/wml`'s `mcBranchOf`, `childValues` and `textOf`; its CR-002 (fragments
  unmarshalled in their container).
- The portfolio: `docs/text_extraction_behaviour_model.md`, sections 4, 8 and 9 (D4); decisions
  taken at 2d8fd36.
- The survey script is not committed. It loads each file twice, resolved and with
  `mcePreprocess: false`, and compares every element's normalised form. Phase 0 made it a test over
  this repository's fixtures (section 11.6).

## 9. Phase 1 (implemented 2026-10-17)

### 9.1 What changed

- **`src/opc/mce/preprocessor.mts`.** Load leaves two kinds of `mc:AlternateContent` whole, branches
  and all, including any element nested in them:
  - under a host in `KEPT_HOSTS` (`c:chartSpace` and SpreadsheetML's `workbook`), or under the
    caller's own `keptHosts`, as `{namespace}localName`; an empty set resolves everywhere, as
    before;
  - where nothing can be selected (Choices only, none understood), which ISO, and this package
    until now, removes (D4).
- **Outermost first.** The pass now resolves each element before what it contains, then visits
  the content it leaves in place; before, it went innermost first. Resolution comes out the same,
  but an element nested in a dropped branch is no longer visited, so it reports nothing.
- **New entry points.** `preprocessAlternateContent(root, options)` returns
  `{ resolved, kept, warnings }`, and `resolveAlternateContent` still returns the number resolved.
- **The warning.** `mce.branches_differ` (`McBranchesDifferWarning`) is recorded when the text of
  the branch taken differs from that of a branch dropped:
  - The text is a branch's `w:t`, `a:t` and `m:t`, with whitespace collapsed. A nested element
    contributes the branch the same rule selects, or its first Choice when nothing can be
    selected (the reading rule).
  - It carries the part name, the line and column (where the DOM records them, as xmldom does),
    the branch taken (`Choice(wps)`, `Fallback`), every branch's text and a message.
- **Where warnings collect.** `XmlPart.warnings` holds what the preprocessor returned the last
  time the part was parsed, by `getContents()` or a private read, so a reader that leaves the part
  untouched gets them too. `setContents` clears them. `OpcPackage.warnings` gathers the warnings of
  every part parsed so far. A `DomPreprocessor` may now return warnings, and one that returns
  nothing leaves none.
- **Exports** from `./opc`: `KEPT_HOSTS`, `preprocessAlternateContent`, and the types
  `McBranchesDifferWarning`, `PartWarning` and `MceResult`.
- **Tests:**
  - five new ones in `mce.test.mjs`: a chart part's `c14:style` and its `c:style` Fallback
    through `unmarshalAll` and a save; `KEPT_HOSTS` and an empty `keptHosts`; an element nothing
    can be selected from, through an edit and a save; the warning on a text box, with
    `setContents` clearing it; the warning from a private read of docx4j's waterfall;
  - two changed: `create.test.mjs`'s workbook test now asserts that `absPath` is kept, where it
    asserted the loss, and `mce.test.mjs`'s first test asserts the kept element, where it
    asserted removal.

  Run against the build of `3d609c0` (with the two new names shimmed), all seven fail and the rest
  pass. Suite 838 passed.

### 9.2 Measured over the 459 files

- **Re-marshalled** after `unmarshalAll` and a save, 258 of the 8,235 parts change: exactly the 227
  chart parts (172 in xlsx, 32 in docx, 23 in pptx) and the 31 workbooks that hold the element.
  Every other part is byte-identical to `3d609c0`'s output.
- **In those 258**, every element comes back whole (names, attributes and children; prefix
  declarations aside). At `3d609c0`, none of the 258 did.
- **No reader is affected.** Nothing in `src/` reads a chart part's tree (`ChartPart` is the class
  alone) or a workbook's `alternateContent`.
- **The warning:** 11, from private reads of every XML part; no file failed to load. The nine
  `wps` text boxes' branches agree.

  | branches | warnings | where |
  |---|---|---|
  | the Choice without text, the Fallback with Excel's notice | 7 | cx1 4 ("This chart isn't available in your version of Excel"), the a14 slicer, sle15, tsle |
  | the Choice with text, the Fallback without | 3 | two a14 checkboxes ("Check Box 1"), a pptx slide's a14 text box |
  | both with text: the same equation in OMML and in its linear form | 1 | `loadAndSave.xlsx`'s a14 equation |

  - A reader on the Fallback would read the first seven's notices, as docx4j's default and the
    paper's Fallback-only tools do, and nothing for the next three. So those ten are what the
    warning is for.
  - The equation is a false positive of the text comparison. The Choice's `m:t` runs and the
    Fallback's linear `a:t` are two notations of one formula. The read-time form (CR-012's since
    2026-10-19, section 14.7) can compare math by its own rule, and this one is left as it is.

### 9.3 Not done in phase 1

- **The worksheet's `x14` controls** stay resolved (section 4).
- **The content API** doesn't read a kept branch, so the text of an element nothing can be
  selected from appears nowhere until phase 3.
- **No `warnings: false` option** (the paper's section 4.1 item 6). The load-time comparison is
  one pass over branches that are dropped in any case.
- **The warnings cover only parts parsed so far.** A caller wanting a whole package's warnings
  reads its parts first; private reads keep them untouched.

## 10. Phase 2: the objects package's CR-009 (implemented 2026-10-18)

Jason asked for phase 2 on 2026-10-17. Phase 2 is objects work, so this package's part is the
route and its evidence. Both are in `../docx4j-generated-objects-ts`'s
`docs/change-requests/CR-009-type-kept-mc-branches.md`. It was written here and taken by that
repository's session the same day (`1e4da5a`, pushed), which added a verification section of its
own (its section 7). Jason gave the go there on 2026-10-18, and it was released as objects 0.4.0
the same day (section 10.4).

### 10.1 The route: section 3 item 4's first

- **What the wildcard types.** Inside a kept branch the runtime types an element only when the
  context declares it globally.
- **The gap.** docx4j declares 640 elements globally (`@XmlRootElement` and unscoped
  `@XmlElementDecl`, its generated objects at `VERSION_17_3_4`). Of those, this mapping declares
  46 WordprocessingML elements only within parents and 15 only as properties; the rest are global
  here too, outside WordprocessingML, or ambiguous in docx4j itself (`w:sdtContent`).
- **The proposal.** Declare those 61 globally, each with docx4j's root class's type.
  - As the objects session refined it, the recommended form is one more hand-written mapping
    beside the generated modules. It depends on `org_docx4j_wml`, declares no types, and names its
    element declarations' types by qualified name. `modulesFor` includes it with
    `org_docx4j_wml`, and the context isn't mutated.
  - The alternative is the generated mapping, through a compiler CR.

The second route, typing a DOM branch child in its host's container at load, would reach the same
types by more machinery. Its scope-correct typing would differ from docx4j's only where docx4j
itself types globally. So the CR doesn't propose it.

### 10.2 Measured, with the 61 registered in a live context (no module edited)

- **Kept branches** (`mcePreprocess: false`): `w:drawing` 13 and `w:pict` 9 come back typed. Every
  other DOM entry in section 2's second table stays DOM, since it is outside WordprocessingML or
  `c:style`.
- **The round trip.** 277 of the 281 parts write every element back whole. The four others
  (`loadAndSave.docx` and the three `cr043-*.docx`) lose `w14:anchorId` on a Fallback's
  `v:shapetype`, which `CTShapetype` doesn't bind, and nothing else. The CR's item 2 asks for that.
  - docx4j 17.3.2 loses the attribute too: loading `loadAndSave.docx` and marshalling its main
    document part, the Fallback is kept and the attribute gone. It also writes `o:spt="202"` as
    `"202.0"`. A Java check from the scratchpad against the harness's jars, not committed.
  - Two rewrites are lexical, and a resolved Choice gets them today: `xsd:boolean` `1` as `true`,
    and `o:gfxdata`'s base64 without its line breaks. The comparison treats both as equal.
- **Default load** (resolving): none of the 8,235 parts changes.
- **Closed since.** docx4j's CR-048 (`acd0defc9`, reported from here) declares `w14:anchorId` on
  every VML shape and `v:group`, and `o:gfxdata` where it was missing. The objects package
  regenerated from it (`97c43b9`), and that is in 0.4.0, so the four parts are whole (section
  10.4). docx4j's own fixtures lost more than this survey found: `anchorId` on `v:shape` and
  `v:rect`, `o:gfxdata` on `v:oval` and `v:rect`.

Two findings from the objects session's own verification (its CR-009 section 7):
- **`o:spt` is written back `"202"` here**, where docx4j writes `"202.0"`.
- **Each of the 61 becomes a valid root,** so `unmarshalString('<w:drawing/>')` succeeds where it
  threw. That matters for phase 3: a fragment of these elements no longer needs wrapping in its
  container.

### 10.3 Here

- **The test that announced the release.** `mce.test.mjs`'s preprocessing-off test asserted that
  the kept `w:drawing` and `w:pict` were DOM, and it failed on CR-009's build, as intended. It now
  asserts the typed form: the `w:drawing` with its `wp:anchor`, and the `w:pict` with its
  `v:shapetype`, whose `w14:anchorId` the save writes back. On 0.3.2 it fails.
- **The range** is `^0.4.0`, and section 10.2's survey, re-run, is phase 2's acceptance (10.4).
  The survey scripts are in the scratchpad; phase 0 made them a test (section 11.6).

### 10.4 Acceptance (2026-10-18)

Checked twice: before the release, against `npm pack` of the objects checkout at `abcb881`; after
it, against 0.4.0 from npm, whose tarball is file for file the same except its version. The
"before" side is this package's HEAD on 0.3.2. Both runs gave the same results.

- **The suite.** Only the announcer failed (section 10.3). Parity, XPath parity and the typechecks
  passed.
- **Kept branches** (`mcePreprocess: false`):
  - `w:drawing` (13) and `w:pict` (9) are typed. Section 2's other DOM entries stay DOM, since
    they are outside WordprocessingML or are `c:style`.
  - All 281 parts write every element back whole, the four `w14:anchorId` parts included.
  - Eight parts differ from their source outside the element, identically on 0.3.2. They predate
    CR-009: `iterateDelta="1E-4"` written `"0.0001"` in three workbooks, and `c:strRef`'s
    `c:extLst` put in schema order in five filtered charts.
- **Bytes, every entry after `unmarshalAll` and a save** (459 inputs, 8,235 entries):
  - Default load: none changes.
  - Kept branches: nine `word/document.xml` parts change. Compared element by element, every
    difference is lexical:
    - `xsd:boolean` `1`/`0` written `true`/`false`;
    - `o:gfxdata` without its line breaks;
    - attribute order and namespace declarations;
    - whitespace-only text between elements under element-only parents.
    No attribute or element is added or dropped, and no text changes.
- **An element marshalled on its own** (`Paragraph.getXml`, `ContentControl.getXml`,
  `toApiScript`'s fallback, which `insertXml` takes back). Over the docx fixtures' 8,089 such
  elements:
  - `w:p`, `w:sdt` and `w:tr` are unchanged.
  - `w:tbl`, `w:tc`, `w:drawing`, `w:bookmarkStart` and `w:bookmarkEnd` no longer carry
    `xsi:type`, now that they are declared roots, as docx4j's are. Nothing else changes.
  - Both forms, with and without `xsi:type`, read back and marshal to the same new output, so a
    script written before still loads.
- **Not done:** the content API over kept branches (phase 3), and the survey as a test (phase 0,
  done since: section 11.6).

## 11. Phase 0: the audit (2026-10-18)

Jason asked for phase 0 on 2026-10-18 ("do CR-010 phase 0"). This section answers section 4's three
questions and makes the survey a test, as sections 8 and 10.3 said phase 0 would. Phase 3 starts from
section 11.7's list. The test is in `mce.test.mjs` (section 11.6). The probes and the harness
copy behind sections 11.2, 11.4 and 11.5 are in the session's scratchpad, not committed.

### 11.1 How a walker meets a kept element

- **Where the element can stand.** The model admits `mc:AlternateContent` in a paragraph's content
  and that of the run holders in it (`w:hyperlink`, `w:fldSimple`, run-level `w:sdt` and
  `w:customXml`, `w:smartTag`, `w:dir`, `w:bdo`), in a run's content, and in a math run's. Block
  containers (body, cell, block content control) can't hold it. Every WordprocessingML instance in
  the corpus is in a run (section 2).
- **Three kinds of walk**, and the kind decides what a walker does with a kept element:
  - *Generic walks* descend every property: the objects package's `walk`, `find`, `walkAll`,
    `nextSdtId` and `linkParents`, and this package's own recursions over `Object.keys`. They reach
    every branch, which is docx4j's ALL.
  - *Named walks* follow the model's content properties: `childrenOf`, `runItemsOf`, `segmentsOf`,
    `runsOf`, `rowsOf`, `cellsOf`. An `AlternateContent` has none of them (its branches are `choice`
    and `fallback`), so they step over it and see no branch. docx4j has no counterpart, since
    `TraversalUtil` always descends.
  - *`textOf`* takes `mcBranchOf`'s branch, which is docx4j's READ. It is the only walker that does.
- So a reader written as a generic walk counts a text box twice once the element is kept. A mutator
  written as a named walk misses the element altogether. Phase 3's work is those two lists.

### 11.2 The walkers

Measured on the eight `.docx` fixtures that hold the element, each loaded twice, resolved and
kept, with every reader and mutator below run on both loads and the results compared. The
fixtures are `loadAndSave`, the two `numbering-stories`, `tables-textbox-in-cell`,
`charts/waterfall` and the three `cr043-*`. Each walker was also read. A row groups the walkers that
share a primitive.

**Readers: one branch (READ).**

| walker | with the element kept, today | phase 3 |
|---|---|---|
| `List.mts` `walkStory`, under `Body.listLabels`, `ListItem.listString`, `level` and `siblingIndex` | generic. A text box's list paragraphs are numbered once per branch: `numbering-stories` gives 12 labels where resolving gives 9. The body's own numbers don't move, because a text box counts in a fresh state | `mcBranchOf` at the element |
| `TableContextTracker.walk` (public; the parity test uses it) | generic. A text box's paragraphs are visited once per branch, each with no context | `mcBranchOf` |
| `Paragraph.inlinePictures`, under `Body.inlinePictures` | named. An inline drawing in the Choice is no longer listed: 1 to 0 in each `numbering-stories`, 3 to 0 in `tables-textbox-in-cell`, 1 to 0 in `waterfall`. Check 42 found since that Office JS lists no inline picture for `numbering-stories`' inline text box, so for a text box 0 is Office JS's count, and the resolved load's 1 is the defect (section 12.6) | pictures only (`pic:pic`), from the selected branch (section 12.6) |
| the `segmentsOf` and `runsOf` readers: search, the text views, `Range` reads | named. Same results either way, since a drawing adds no text to a paragraph's | none, until the text CR (section 7) puts a text box's text in |
| `textOf`: `Paragraph.text`, `Body.text`, the outline's text, `toApiScript`'s cell text | READ already. A drawing gives no text in either load | none |
| `placeholdersOf` (a pptx layout's placeholders) | named. A placeholder inside an element would be missed. The corpus has none there | `mcBranchOf`, low priority |
| `Body.select`, `selectObjects` (XPath over a marshalled snapshot) | every branch, as docx4j's `getJAXBNodesViaXPath` gives: 188 of 188 agree (section 11.4) | none. The expression decides, as with docx4j's finders (CR-021 section 8.3) |

**Mutators: every branch (ALL).**

| walker | with the element kept, today | phase 3 |
|---|---|---|
| everything through `segmentsOf({ objects: true })` (CR-002 section 48): `Paragraph.delete` and `Range.delete` under tracking, `insertText(..., 'Replace')`, an untracked `getRange('Whole').delete()` | named. The element is not an `OBJECT_ITEMS` segment, so a tracked `Paragraph.delete` leaves its run outside the `w:del`, and Accept All then leaves the box behind. That is section 48's gap again, on all five fixtures tried (`loadAndSave`, `numbering-stories`, `tables-textbox-in-cell`, `waterfall`, `cr043-tracked`). An untracked `getRange('Whole').delete()` or `insertText('', 'Replace')` leaves it in place, where resolving removes the drawing. `Range.delete` under tracking shares the path and was read, not run | the element as one object segment. Removing it removes every branch |
| `pkg.acceptAllRevisions`, `rejectAllRevisions`: the text box stories (`textBoxesIn`, generic) | every branch's `w:txbxContent` is resolved, which is right. The count returned counts a box's changes once per branch: `cr043-tracked` returns 39 where resolving returns 37 | resolve every branch; count the selected one |
| removal's relationship cleanup (`relationshipIdsIn`, generic) | every branch. With the element kept, the Fallback's picture relationship goes with the chart's: `waterfall` drops both. Resolving leaves the image relationship behind, its reference already dropped at load | none |
| `insertOoxml` and `insertXml` reference rewriting (`visitReferences` over `walkAll`) | every branch | none |
| what a resolution checks around it: `bookmarksIn`, the broken-bookmark pass, `noteReferencesIn`, `commentReferencesIn` (all generic) | every branch. A note or comment referenced only from a Fallback stays referenced | none |
| id allocation: `highestAnnotationId` (the tracker and `seedAnnotationIds`), `nextSdtId`, `nextDrawingId` (all generic) | every branch. An id used in any branch is taken | none |
| `linkParents` | every branch | none |
| `fontsInUse` and `getStylesInUse` (`walkStories`, generic) | every branch, as docx4j's `FontsAnalysis` and `FontAndStyleFinder`: a face used only in a Fallback is still used | none |
| an untracked `Paragraph.delete`, `Body.clear` | the paragraph goes whole, every branch with it | none |
| the picture binding's `firstBlipOf` (`find`) | the first `a:blip` in any branch, so the Choice's. A Fallback's `v:imagedata` would keep the old image. No picture control in the corpus wraps its drawing in mc | every branch's picture |
| `paraIdTaken`, `Body.assignParaId` | block walks. A text box's paraIds are not seen in either load, which predates this CR | ALL, with the text boxes |
| `applyBindings`, `updateFromContentControls` (controls from `collectControls`, named) | a control in a text box is not reached in either load, which predates this CR. docx4j's `BindingHandler` reaches every branch | ALL, if text box controls join |
| `replaceText` | segment-based. A text box's text is not searched in either load (predates) | ALL, if text boxes join search (the text CR, section 7) |

**Indifferent.** These are block walks, or they stop where a drawing begins:
- `Body.paragraphs`, `tables`, `outline`, `rowElements`, the addresses (`elementAt`, `pathOf`,
  `addressOf`);
- the story walk of `getTrackedChanges` and `trackedChangeTokensOfParagraph` (a text box is a story
  of its own, reached through `resolveStory`), `movePiecesOf`, `resolveUnlisted`, the
  `insertedFieldAt` scan;
- `Table`, `tableColumns`, `tableRevisions`;
- the offsets that count text up to an element: the comment markers' (`collect`),
  `acceptedOffsetOf`, a content control's empty position (the element adds no text, so a segment
  of no width for it, section 11.7 item 2, moves none of them);
- `collectControls` as a listing, `hyperlinkHolders`;
- `PropertyResolver.cellContextOf` (up the `PARENT` chain, stopping at `w:txbxContent`);
- `toApiScript`: an element in a run can't be expressed, so the paragraph falls back to its
  marshalled XML, every branch kept;
- `html/convert` (it walks HTML), and the preprocessor itself.

### 11.3 The editor

This is a read-only survey of `../docx4j-ts-editor` at `a7d67e4` (2026-10-17), with its main claims
checked in the code. The editor loads with this package's default everywhere. Nothing in it names
`mcePreprocess`, `KEPT_HOSTS` or the warnings.

- **What it reads through the tree.** Its ProseMirror import (`editor-model/src/import.mts`) walks
  paragraphs and runs itself: `runItemsOf`, then a run's content by local name. Its export writes
  back, by identity, whatever it did not rebuild.
  - **Nothing breaks.** A kept element in a run becomes an opaque inline atom of kind
    `AlternateContent`, which is written back whole. So an edited part keeps both branches. Today
    every text box in that part loses its wrapper and its Fallback.
  - **What it shows changes.** Today any inline `w:drawing` becomes an image node, whatever its
    graphic. `numbering-stories.docx`'s inline `wps` box is drawn as a missing image, and kept it
    would be the opaque glyph. A Fallback picture behind a Choice nobody understands would also go
    from an image to the glyph. The import needs `mcBranchOf` where it meets the element, and so do
    its image lookups.
  - **A paragraph-level element** holding runs would turn read-only, and its text would be missed
    by the editor's find, its offsets, the MCP read and the outline. The corpus has none.
- **Its own generic walks:**
  - `labels.mts` `walkStory`, a copy of this package's, numbers both branches. Inside a text box
    that is harmless (fresh counters), and only the test suite uses it as an oracle; the app uses
    `liveLabels`.
  - `AskBox`'s `textOfElement` collects every `w:t`, so both branches' text.
  - `applyReferenceIds`, `hasPageField`, `usedSdtIds` and the `sources` path ids see both branches
    and stay correct, since they rewrite or collect sets.
- **What depends on phase 3 here:**
  - Its XPath hit placement and the MCP find climb `PARENT` through the branches. Load links them:
    a Fallback paragraph's chain runs `w:txbxContent`, `v:textbox`, `v:shape`, `w:pict`, `Fallback`,
    `AlternateContent`, `w:r`, `w:p`, `w:body`.
  - `preserved.mts` numbers `getTrackedChanges()` entries within a block, so phase 3 must not list a
    change once per branch.
  - Its find and XPath tests compare its answers with `search` and `selectObjects`.
- **Tests that pin today's tree:** `diagnostics.expected.json` (`loadAndSave`'s opaque `drawing`,
  `numbering-stories`' image node) and `editor/test/diagnostics.test.tsx`. They change in step with
  phase 4.
- **An inconsistency, for the editor to record.** ED-002 says the XML pane's Apply runs the MCE
  preprocessor, but `app/xml.ts` uses plain `unmarshalString`. So an element pasted there stays in
  the tree, while the save's self-check reload resolves it. Keeping on load removes the mismatch.

### 11.4 The four excluded XPath fixtures

docx4j's `XPathHarness` was run from a scratchpad copy that records hits for every fixture. It used
the jar the XPath goldens were made with (`2e381ef26`, `f805ae4be89897e3`). Its hits were then
compared with `selectObjects` as `xpath-parity.test.mjs` compares them, with the element kept
and with it resolved:

| fixture | kept | resolved |
|---|---|---|
| `loadAndSave.docx` | 47 of 47 | 39 of 47 |
| `numbering-stories-coretests.docx` | 47 of 47 | 33 of 47 |
| `numbering-stories.docx` | 47 of 47 | 33 of 47 |
| `tables-textbox-in-cell.docx` | 47 of 47 | 36 of 47 |

- **All four come back** once the element is kept, with objects 0.4.0's typed branches.
- **Resolved, they differ** where a path crosses the element: `//w:p`, `//w:r`, `//w:t`,
  `//w:drawing`, `//a:graphic`, `//w:p[.//w:t]`, `//w:p[last()]` and `//w:t/text()` in all four;
  `//w:pPr` and `//wp:inline` in the other three; `//w:rPr`, `//w:numPr` and their predicates in the
  two `numbering-stories`; `//w:tbl//w:p` in `tables-textbox-in-cell`. docx4j selects a text box's
  paragraphs under both branches, and the resolved tree has them once, under `w:r/w:drawing`.
- They need not wait for phase 4. The harness drops its exclusion and records their hits, and the
  test loads those four with `mcePreprocess: false`. Then `selectObjects` over a kept tree is held
  to docx4j while phase 3 changes the walkers around it (D5).

### 11.5 What the parity goldens need from D1

- **Nothing new.** Since harness version 3 (CR-021 phase 1), the harness has set docx4j's
  `preferChoice` to this package's understood prefixes, so docx4j already reads the Choice the
  TypeScript side resolves to. The four parity fixtures holding the element all require `wps`.
- **Measured.** `parity.test.mjs` was run from a scratchpad copy that loads with
  `mcePreprocess: false`. 8 of 249 tests fail: the four fixtures, two tests each ("17 paragraphs,
  the golden has 14"). The test's own two walks, `visit` and `walkNumbering`, are generic. With
  each taking `mcBranchOf`'s branch at the element, all 249 pass. So phase 4's change to the parity
  test is those two lines, and the goldens stand.
- **A drift in the harness.** `MC_PREFER_CHOICE` names 84 prefixes, and `UNDERSTOOD_NAMESPACES` has
  100 with prefixes. Seventeen understood here are missing from it (`c16r3`, `cei`, `cx1`, `cx2`,
  `sle`, `sle15`, `tsle`, `w16`, `w16cex`, `x12ac`, `x14`, `x15`, `x15ac`, `x16`, `xcalcf`, `xm`,
  `xr`). It also names `c173`, which is not understood. Its comment says it is that set. No parity
  fixture has a Choice requiring any of them, so no golden would move. It is refreshed when the
  goldens are next regenerated, which CR-011's phase 0 does (D6).
- **For phase 3's readers.** `mcBranchOf`'s default set is `NAMESPACE_PREFIXES`' 136 prefixes, 36
  more than load understands (`x14ac`, `w16du`, `xr2`, `cx3` to `cx8`, ...). A reader on the
  default would take a Choice the loader rejects, so phase 3 passes the understood set itself
  (`UNDERSTOOD_NAMESPACES` through `NAMESPACE_PREFIXES`) until `mcBranchOf` matches URIs (D4 item
  1, an objects change). Seven understood namespaces have no conventional prefix, the SpreadsheetML
  main among them. No `Requires` in section 2's corpus names one.

### 11.6 The survey as a test

`mce.test.mjs` gains "kept branches: each fixture part holding mc:AlternateContent writes every one
back whole, its WordprocessingML typed". It walks `test/fixtures/`: every zip package, and every
flat OPC package. Each is loaded with `mcePreprocess: false`, unmarshalled and saved, and then:

- **Whole.** Every part that held the element writes each one back: names, attributes and
  children, with prefix declarations and section 10.4's three lexical rewrites left aside. That is
  28 parts: 27 in the zips, and the flat OPC fixture's header.
- **Typed.** No WordprocessingML element in a kept branch is DOM, and typed `w:drawing` (12) and
  `w:pict` (10) are both found. What stays DOM is outside WordprocessingML: `xdr:sp` 6,
  `xdr:graphicFrame` 4, `c:style` 4, `controls` 2, `xdr:twoCellAnchor` 2, `p:sp` 2,
  `p:graphicFrame` 1, `p:pic` 1.
- **A floor of 28 parts**, so the walk can't pass by finding nothing.
- **It bites.** Loading the default way (resolving) fails it, first on `charts/waterfall.docx`.
- **Cost.** `mce.test.mjs` takes under two seconds with it.

The release checks' run over the 459 files stays a script, since docx4j's charts directory is not
in this repository.

### 11.7 For phase 3

1. **One selection helper**: `mcBranchOf` with the understood prefixes (section 11.5). Every READ
   walker in section 11.2 goes through it.
2. **`segmentsOf`** lists the element as one `OBJECT_ITEMS` segment, where its `w:drawing` stood.
   That closes section 48's gap for a tracked delete, a range deletion and the text setter, since
   removing the element removes every branch. `inlinePictures` lists pictures only (`pic:pic`), in a
   drawing or in the selected branch, since Office JS lists no inline text box (section 12.6).
3. **`walkStory` and `TableContextTracker.walk`** take the selected branch at the element.
4. **`resolveStory`** goes on resolving every branch's text boxes, and counts the selected branch.
5. **Predating this CR, and decided with the text CR** (section 7): text boxes in content controls
   and bindings, in search and `replaceText`, and their paraIds.
6. **The editor**, before phase 4: `mcBranchOf` in its run projection and image lookups, its label
   oracle and `AskBox`'s text walk, and its diagnostics fixtures (section 11.3).
7. **Tests**: the XPath four rejoin (D5); the parity test's two walks (section 11.5).

### 11.8 Decisions

- **D5: the four XPath fixtures rejoin now, loaded kept** (section 11.4). The harness records their
  hits, their four goldens are regenerated from the same docx4j, and `xpath-parity.test.mjs` loads
  them with `mcePreprocess: false`. Recommended. The kept tree's XPath answers are then held to
  docx4j through phase 3, not only from phase 4. **Decided 2026-10-18** (Jason: "D5: ok").
  **Implemented the same day**, after phase 0's commit:
  - `XPathHarness` records hits for every fixture; the header's `alternateContent` still counts the
    elements, and its log line names the fixtures the test keeps them in.
  - The four goldens were regenerated with it from the jar the others were made with (`2e381ef26`,
    `f805ae4be89897e3`). Regenerating all 62 the same way changed nothing in the other 58 but their
    dates, so only the four are replaced. Each gains its 47 expressions' hits and the prefixes its
    paths use, and they equal section 11.4's run.
  - `xpath-parity.test.mjs` loads a fixture whose golden counts an element with
    `mcePreprocess: false`; its first test asserts every golden has hits and that the fixtures
    holding the element are the four listed (`KEPT`, formerly `EXCLUDED`). 64 tests, 4 more; loaded
    the default way, those 4 fail.
- **D6: `MC_PREFER_CHOICE` refreshed** to `UNDERSTOOD_NAMESPACES`' prefixes (section 11.5), with
  CR-011's phase 0, whose regeneration touches every golden's header anyway. Recommended. No answer
  moves. **Decided 2026-10-18** (Jason: "D6: ok"); it is done with CR-011's phase 0.

## 12. Phase 3: does a text box's paragraph count in `Body.paragraphs`? (2026-10-18)

Jason asked for this before any phase 3 code ("write the short phase 3 section on whether text-box
paragraphs count in Body.paragraphs, with a recommendation, before writing any code"). Section 4
gives phase 3 the question of how a text box's paragraphs count. Once the element is kept, a text
box exists twice, once in each branch, so the answer decides what every count is over.

### 12.1 What stands

- **`Body.paragraphs` is the story's block walk**: the body, tables, cells and content controls. It
  doesn't enter a run, so a text box's paragraphs are not in it, whether the element is resolved or
  kept. Nor are they in what is built on it: `text`, `getText`, `search`, `replaceText`, `outline`,
  `inlinePictures`, the addresses and `select`'s views. That last one already says a run in a text
  box "stays a `Bound`".
- **Every other rule here makes a text box a story of its own**:
  - list numbering counts it in a fresh state (CR-001 section 15.2, docx4j's CR-014 probe P7);
  - the cell context stops at it (CR-007, docx4j's CR-030 probe T3);
  - `pkg.acceptAllRevisions` resolves it as a story, while `Body.acceptAll` stays its host story's
    (CR-002 section 49).

  `Body.listLabels` is a story walk, not `Body.paragraphs`. It lists a text box's numbered
  paragraphs, each counted in the box's own state: 3 of `numbering-stories`' 9 labels.
- **The other engines and the specification:**
  - docx4j-python's `Body.paragraphs` and `Body.text` leave text boxes out ("Word's main story, no
    text boxes", its CR-003 decided question 8).
  - docx4j has no `Body.paragraphs`. Its `TraversalUtil` walks (READ) visit a text box's paragraphs
    once, inside the anchoring paragraph.
  - The text extraction specification WD 0.9 (portfolio `a195eae`) emits a text box's paragraphs
    once, after the anchoring paragraph, from the selected branch (REQ-518). It addresses them under
    that paragraph, `body/3/textbox:0/1`, never as body ordinals (REQ-519). Its note says the ports
    adopt these segment kinds with their text-model CRs.
- **Word** counts a text box as a story of its own (section 7, from `StoryRanges`). What Office JS's
  `body.paragraphs` gives over a document with a text box was not measured when this was written;
  check 42 measured it since (section 12.6): it leaves the box out.

### 12.2 The options

- **A. Leave them out.** `Body.paragraphs` stays the story's own paragraphs, whether the element is
  kept or resolved. A text box's paragraphs are a story, reached on their own (section 12.4).
- **B. Put them in once**, from the selected branch, after the anchoring paragraph: docx4j's walk,
  in REQ-518's order.
- **C. Put them in from every branch.** Rejected: each text box would appear twice, which is the bug
  that started docx4j's CR-021.

### 12.3 Recommendation: A

- **No answer moves.** Phase 3 is to make a kept tree answer as a resolved one does. B would change
  `Body.text`, `outline`, `search`, `lists` and `inlinePictures` in every document with a text box,
  under the default load too.
- **The addresses stay ordinals.** `body/3` and `body/3/1/0/2` count block children. B would put
  paragraphs in the list that have no block ordinal. REQ-519 addresses them under their anchor
  instead, which A leaves room for.
- **One rule.** Numbering, the cell context and tracked changes already treat a text box as a story,
  and so do Word and python. B would make `Body.paragraphs` the one exception.
- **Writes.** Under B, `delete` or `insertText` on a text box's paragraph would edit only the
  selected branch and leave the Fallback saying something else, against D2. Under A, nothing in the
  content API writes to a text box until a view designed to edit every branch exists.
- **Consumers.** The editor's projection, its console and the MCP outline go on seeing the
  paragraphs they see today.

A changes nothing in `Body.paragraphs`. What phase 3 still owes is in the story walks, section 11.7
items 3 and 4.

### 12.4 What follows from A

- **The story walks count a text box once**: `listLabels` (in the box's own state, as now),
  `TableContextTracker.walk`, fonts and styles in use, and CR-011's listing (after its story).
  Readers take the selected branch; mutators (resolution, removal's cleanup) take every branch.
- **Reaching a text box** is a view of its own. It would be a `Body` per box anchored in a
  paragraph, in anchor order, over the selected branch's `w:txbxContent`, addressed
  `…/textbox:N/…` (REQ-519).
  - Reading through it is straightforward.
  - Writing through it has to edit every branch's twin (D2). The corpus's nine boxes hold the same
    text in both branches (section 9.2), but nothing guarantees the same runs, so matching the
    twins is a design of its own.
  - REQ-519's note gives the segment kinds to the text-model CR, which takes the next free number
    (CR-011 is tracked changes now).

### 12.5 Decisions

- **D7: `Body.paragraphs` leaves a text box's paragraphs out, whether kept or resolved** (option A).
  Recommended. Phase 3 then changes nothing that `Body.paragraphs` and the views built on it report.
- **D8: the text-box view goes to the text-model CR**, not phase 3. That includes `textbox:N`
  addresses, reading, and writes that land in every branch. Recommended. It keeps phase 3 to
  making a kept tree answer as a resolved one, and puts the view with the rest of REQ-518 and
  REQ-519.
- **Decided 2026-10-18, both, subject to a check in Word** (Jason: "D7 and D8 agreed, subject to a
  check in word"). The check is `test/README.md`'s check 42, a Script Lab snippet on a copy of
  `numbering-stories.docx`. It reads Office JS's `body.paragraphs` (11 here, 14 with the box's),
  whether `body.text` holds the box's text, and `body.inlinePictures` (1 here: the box's inline
  drawing). If Office JS counts the box as the body's, D7 is revisited before phase 3 relies on it.
  Run the same day: it doesn't, and D7 and D8 stand without condition (section 12.6).

### 12.6 Check 42 (2026-10-18)

Jason ran check 42 in Word 16.0.20326.20232 (PC) on a copy of `numbering-stories.docx`.

- **D7 stands.** `body.paragraphs` holds 11 paragraphs, the same as `Body.paragraphs`, and
  `body.text` leaves out the box's text (P20 to P22). Office JS leaves a text box's paragraphs out of
  the body, as option A does. D7 and D8 hold without condition.
- **Found: Office JS lists no inline picture for an inline text box.** `Paragraph.inlinePictures`
  lists every `w:drawing` holding a `wp:inline`, whatever its graphic, so it lists the box. It gives 1
  where Office JS gives 0.
  - This predates this CR (CR-002's `InlinePicture`) and shows in the default load.
  - Section 11.2 read the kept load's 0 as a loss. For a text box, 0 is Office JS's count.
  - Phase 3 therefore lists pictures only, a drawing whose `a:graphicData` holds a `pic:pic`, in the
    drawing itself or in the selected branch (section 11.7 item 2). Resolved and kept then agree
    with Office JS, and `tables-textbox-in-cell`'s three boxes and `waterfall`'s chartex frame are
    no inline pictures in either load.
  - Whether Office JS lists an inline chart is not measured. `loadAndSave.docx` has an inline
    `c:chart` beside an inline picture, and this package listed both. Check 43 measured it since:
    Office JS lists no inline chart either (section 13.4).
- **Found: Office JS puts U+0002 in a paragraph's text for each note reference.** The two footnote
  references in "P11 Two footnotes and" and the endnote reference in "P15 An endnote" each read
  `\u0002`. P19's comment reference reads as nothing. `Paragraph.text` (`textOf`, docx4j's
  `TextUtils`) gives nothing for any of them. This predates this CR and is not phase 3's. It is a
  rule of the text model, for the text-model CR to settle against the specification.


## 13. Phase 3: the content API over kept branches (2026-10-18)

Jason asked for phase 3 on 2026-10-18 ("do CR-010 phase 3"). It works from section 11.7's list, with
D7 and D8 (section 12) and check 42's finding on inline pictures (section 12.6). Implemented the same
day. A kept tree now answers as a resolved one does, and the default load stays resolving until
phase 4.

### 13.1 What changed

- **One selection helper** (section 11.7 item 1):
  - `UNDERSTOOD_PREFIXES` (`./opc`) is `UNDERSTOOD_NAMESPACES` through `NAMESPACE_PREFIXES`, 100
    prefixes.
  - `selectedBranchOf(element)` (`./model`) is `mcBranchOf` over that set, so a reader takes the
    branch load would have resolved the element to. A Choice requiring `w16du`, `x14ac` or `cx3`
    (conventional, not understood) gives the Fallback here, where `mcBranchOf`'s own default gives
    the Choice.
  - `isAlternateContent(value)` (`./model`) recognises the element or its value.
  - Every READ walker below goes through the helper.
- **`segmentsOf({ objects: true })`** lists the element in a run as one object segment of no text
  (item 2). A tracked `Paragraph.delete` or `Range.delete` puts its run in the `w:del`, and Accept All
  removes it. An untracked whole-range delete and `insertText(..., 'Replace')` remove it. Every branch
  goes, since the element goes whole. That closes CR-002 section 48's gap on kept trees.
  - Offsets don't move, since the element adds no text.
  - An element in a paragraph's content, outside a run, is not listed. Word writes none, and the
    corpus has none (section 11.3).
- **`inlinePictures` lists pictures only.** That is a `w:drawing` whose `wp:inline` holds a `pic:pic`,
  in a run or in the selected branch of a kept element (`picturesIn`, section 12.6).
  - A picture from a branch carries the element as `InlinePicture.holder`. Its `delete` removes the
    element whole, untracked or tracked, and its image relationship is queued with every branch's.
  - Reads and setters are the selected branch's drawing (section 13.4).
  - **This changes answers in the default load too**, since a chart, a text box or a group is no
    longer listed:

    | fixture | before | now | what is no longer listed |
    |---|---|---|---|
    | `loadAndSave.docx` | 2 | 1 | the inline `c:chart` (check 43: Office JS gives 1) |
    | `numbering-stories.docx`, `numbering-stories-coretests.docx` | 1 | 0 | the inline `wps` text box (check 42: Office JS gives 0) |
    | `tables-textbox-in-cell.docx` | 3 | 0 | three inline `wps` text boxes |
    | `charts/waterfall.docx` | 1 | 0 | the chartex frame |

- **`walkStory`** (`List.mts`, under `listLabels`, `ListItem.listString`, `level`, `siblingIndex`) and
  **`TableContextTracker.walk`** take the selected branch at the element (item 3), so a text box is
  numbered and visited once.
- **`resolveStory`** (`pkg.acceptAllRevisions`, `rejectAllRevisions`) still resolves every branch's
  text box, and counts the changes of the selected one (item 4).
- **`Body.paragraphs` is unchanged** (D7), and so is everything built on it. No text-box view (D8).

### 13.2 Measured

The eight `.docx` fixtures of section 11.2 were each loaded resolved and kept.

| | before, kept | before, resolved | now, both |
|---|---|---|---|
| `numbering-stories` (both): list labels | 12 | 9 | 9 |
| the same: `TableContextTracker.walk` visits | 17 | 14 | 14 |
| `tables-textbox-in-cell`: visits | 11 | 8 | 8 |
| `cr043-tracked`: `acceptAllRevisions()` | 39 | 37 | 37 |
| each fixture: the element left after a tracked `Paragraph.delete` and Accept All, or a whole-range delete | 1 to 3 | 0 | 0 |
| inline pictures | section 11.2's losses | section 13.1's table, before | section 13.1's table, now |

A wider probe found no other difference between the two loads on the eight fixtures. It compared:

- reads: paragraph texts and `Body.text`; every segment with its offsets, and the object segments' offsets;
  `outline`, `search`, `getTrackedChanges`; content controls, tables, lists, addresses; the paragraph
  formatting reads; list labels;
- edits: the listing after a tracked delete, Reject All restoring the element, `insertText('X', 'Replace')`,
  and a save and reload.

`fontsInUse` and `getStylesInUse` walk every branch (section 11.2), and they agree on all eight too.

### 13.3 Tests

- **`test/kept-branches.test.mjs`**, 6 tests:
  - `UNDERSTOOD_PREFIXES` and `selectedBranchOf` against `mcBranchOf`'s default;
  - the eight fixtures' readers, resolved and kept, deep-equal: `Body.paragraphs` and `Body.text`,
    the labels, the tracker's visits, the pictures, both resolve-all counts;
  - pictures only, in both loads;
  - a picture in a branch: listed once from the Choice, or the Fallback when the Choice's prefix is
    not understood; deleted outright, the element and its image part go; deleted under tracking,
    Accept All removes the element and Reject All restores it whole;
  - the three mutators on section 11.2's five fixtures: no element left, and the kept answers equal
    the resolved ones;
  - `cr043-tracked`'s text box: resolved in both branches, counted once.

  Against a build of `3a0050e`, with the two new names shimmed into the test, five fail. The first
  passes only through the shims, and fails on its import without them.
- **`test/parity.test.mjs`.** The four fixtures whose XPath goldens count the element (`KEPT`) run each
  comparison twice, resolved and kept, so the kept tree is held to docx4j's answers ahead of phase 4.
  - The test's walks `visit` and `walkNumbering` take `selectedBranchOf` (section 11.5).
  - A new check: `TableContextTracker.walk` visits the paragraphs the harness's walk lists, each once,
    in order.
  - 261 tests, 12 more. Against `3a0050e`'s build, the four kept "styles, paragraphs, runs and tables"
    tests fail on that check.
- Suite 866 passed; typecheck passed.

### 13.4 Not done

- **Predating this CR**, for the text-model CR (item 5): text boxes in content controls and bindings,
  in `search` and `replaceText`, and their paraIds. The same goes for check 42's U+0002 rule for note
  references.
- **A picture's setters write the selected branch only.** `width`, `height` and the alt text leave a
  Fallback's copy as it was. Matching a branch's twin is section 12.4's design, the same as for a
  text box. No fixture has a picture inside a branch, and Word writes a picture as a plain
  `w:drawing`.
- **The picture binding's `firstBlipOf`** still points the first `a:blip` in any branch at the new
  image (section 11.2), for the same reason.
- **`placeholdersOf`.** A kept `p:spTree` branch holds `p:sp` as DOM (section 2), so there is nothing
  typed to read until the objects package types it.
- **Readers use load's default understood set.** A caller that passes its own `understood` to a
  preprocessor gets no matching option on the readers.
- **Check 43: Office JS lists no inline chart.** `test/README.md`'s check 43 ran on a copy of
  `loadAndSave.docx` (Jason, 2026-10-18: "we'd better check whether Office JS counts an inline chart as
  a picture"; Word 16.0.20326.20232, PC).
  - Office JS gave one inline picture: 72 by 72 points, in the 30th of 41 paragraphs, alt text "Badge
    Tick1 with solid fill". `Body.inlinePictures` gives the same in both loads.
  - The chart in the 41st paragraph is not listed, so leaving it out is Office JS's count, and the
    pictures-only rule stands.
- **Found by check 43: Office JS's image for an SVG picture is the SVG.** The picture's `a:blip` embeds
  `image1.png`, and its `a:extLst` holds an `asvg:svgBlip` embedding `image2.svg`.
  - `getBase64ImageSrc()` gave 1016 characters, exactly the length of `image2.svg` in base64.
    `InlinePicture.getBase64()` gives the PNG's, 5652. Only the lengths were compared.
  - This predates this CR (CR-002's `InlinePicture`) and is not phase 3's.
  - Fixed the same day as CR-002 section 52 (Jason: "on second thoughts, do the svg fix now"):
    `getBase64` gives the `svgBlip`'s part where there is one. Whether Office JS's `imageFormat` then
    says `Svg` is not measured.
- **The editor** (item 6, before phase 4): `mcBranchOf` (now `selectedBranchOf`) in its run projection
  and image lookups, its `labels.mts` oracle and `AskBox`'s text walk, and its diagnostics fixtures. It
  doesn't read `inlinePictures`, so section 13.1's change doesn't reach it there. A script in its console
  that reads `body.inlinePictures` through the shim gets the new count.

## 14. Phase 4: the default flips to keep (2026-10-18)

Jason asked for phase 4 on 2026-10-18 ("implement CR-010 phase 4, coordinating with the editor as
appropriate"). Implemented the same day.
- The default load now keeps every `mc:AlternateContent`, as docx4j does, and a save writes every
  branch. The content API already answered the same over a kept tree (section 13).
- The editor ran its suites against this build and found a defect in jsonix, which the kept load
  is the first to reach (section 14.4). jsonix fixed it in 3.4.2, published the same day, which the
  lockfile now holds.

### 14.1 What changed

- **`mcePreprocess` defaults to false.** `OpcPackage.load` installs `mceKeep` unless the caller passes
  `mcePreprocess: true`, which installs the resolving `mcePreprocess` as before, `KEPT_HOSTS` and the
  nothing-selectable rule included. `{ mcePreprocess: false }` is now the default spelled out.
- **The warning channel stays open.** Before this phase, `mcePreprocess: false` meant no pass at all,
  so a kept load recorded no `mce.branches_differ`. `cr022-checkbox.xlsx` gave 1 warning resolved
  and 0 kept. Phase 1 had left the read-time form to phase 3 (section 4), and phase 3's list
  (section 11.7) didn't carry it. The flip would have closed the channel, so the load-time form now
  runs without resolving:
  - `preprocessAlternateContent(root, { resolve: false })` visits the elements resolving would, in
    the same order. That is outermost first, then into the branch the rule selects, and no element
    `KEPT_HOSTS` or the nothing-selectable rule would keep.
  - It reports the same warnings and leaves the DOM as it was. `kept` counts the elements it met.
  - `mceKeep` is that pass with load's defaults. `createMcePreprocessor({ resolve: false, ... })`
    takes a caller's own set.
  - Exported from `./opc`: `mceKeep`, and `resolve` on `McePreprocessOptions`.
- **Inserted content is written as given** (section 3 item 5). `Body`, `Paragraph` and `Range`'s
  `insertOoxml` and `insertXml` run the package's preprocessor over the fragment. For a loaded
  package that is now `mceKeep`, so an inserted `mc:AlternateContent` is kept, as a created
  package, which never had a preprocessor, always kept it. Under `mcePreprocess: true` it is
  resolved, as before.
- **`clone()`** loads with its original's options, so a clone keeps or resolves as the original did.
- **`selectObjects` reaches inside DOM content** (`binder.mts`). The marshal reports a wildcard
  child's imported copy, but nothing inside it. Kept, a DrawingML frame is DOM from the branch down,
  so `cr022-slicers-timelines.xlsx`'s slicer, formerly the wildcard child, is now three levels
  inside one. An element inside an imported copy now maps to the node at the same place under the
  original. Below a typed ancestor an unreported element stays unmapped, as before. A hit's path
  now crosses the element: `/xdr:wsDr/xdr:twoCellAnchor[1]/mc:AlternateContent[1]/mc:Choice[1]/xdr:graphicFrame[1]/...`.

### 14.2 What a consumer sees

- **WordprocessingML.** The element stands in run content with both branches, a text box's
  `w:drawing` in the Choice and its `w:pict` in the Fallback, both typed. The content API's answers
  don't move (section 13.2): `Body.paragraphs`, text, search, lists, `inlinePictures`, tracked
  changes, the addresses.
  - A walk of the tree written outside this package has to choose: `selectedBranchOf` to read,
    every branch to write.
  - The test's own label oracle (`lists.test.mjs`) is such a walk. It counted 17 paragraphs where the
    golden has 14 until it took the selected branch, as parity's walks had since phase 3.
  - A re-marshalled document holds a text box twice. `cr043-tracked.docx`'s box keeps its insertion
    and deletion in each branch after `Body.acceptAll`, so its XML has 4 revision elements where the
    resolved save had 2, as Word's own file does.
- **SpreadsheetML.**
  - A worksheet's form controls stay in Excel's `mc:AlternateContent`. `x:controls` is DOM in the
    `x14` Choice, since it is a global element neither here nor in docx4j (`CTControls` has no
    `@XmlRootElement`). So `worksheet.controls` is typed only under `mcePreprocess: true` (CR-004
    section 7). The control's `r:id` still resolves to its `ControlPropertiesPart`, and a save
    writes Excel's markup back, which Excel accepts as it accepts the bare form (docx4j CR-021
    section 8.13).
  - The slicer, timeline and `a14` drawings keep their Fallback placeholders, with the frame and
    the shapes DOM in their branches.
- **PresentationML.** The `a14` shape that holds `loadAndSave.pptx`'s equation is DOM in its
  Choice.
  - docx4j types it: `p:sp` is a root element there (`Shape`, `@XmlRootElement(name = "sp")`),
    while the objects package's CR-009 declared WordprocessingML's globals only.
  - `placeholdersOf` doesn't look inside the element. The corpus has no placeholder there.
  - The OMML test (`omml.test.mjs`) reads the typed equation from the resolved load, and checks
    that the kept load writes it back whole.
- **What saves keep.** Every branch, so no Fallback is lost from a re-marshalled part.
  - CR-001 section 21's open kind (a) is closed by default. `cr022-checkbox-linked.xlsx`'s drawing
    holds two `a14:legacySpreadsheetColorIndex` attributes. Resolved, the Choice is typed and
    `a:srgbClr` doesn't bind them, so they go. Kept, they are written back (`sml-extensions.test.mjs`
    pins both).

### 14.3 The survey re-run (section 6)

The corpus is this repository's fixtures plus docx4j's `docx4j-core-tests/src/test/resources/charts/`
(483 packages, 8,568 entries; 459 when section 2 measured it). Each package was loaded, run through
`unmarshalAll`, saved, and hashed entry by entry. The runs were:
- this build, three ways: default, `mcePreprocess: false` and `mcePreprocess: true`;
- HEAD's build (`a2a048a`), two ways: its default and `mcePreprocess: false`.

The scripts are in the session's scratchpad, not committed.

| comparison | result |
|---|---|
| `mcePreprocess: true` against HEAD's default | every entry byte-identical: the option is the old load |
| `mcePreprocess: false` against HEAD's | every entry byte-identical: the keeping pass changes no output |
| the default against `mcePreprocess: true` | 27 entries differ: `word/document.xml` 9, a header 1, slides 3, `xl/drawings` 12, worksheets 2. Exactly the re-marshalled parts that hold the element outside `KEPT_HOSTS` |
| kept: every element back whole (section 11.6's comparison) | 297 elements in 292 parts, all whole |
| kept: parts that differ from their source anywhere | 8, the same as section 10.4's and on the resolved load too: `iterateDelta` in 3 workbooks, `c:extLst` order in 5 filtered charts |
| `mce.branches_differ` from private reads of every XML part | 14 in each load, identical in every package |
| errors on load, `unmarshalAll` or save | none, in any of the five runs |

What the kept branches hold:

| typed | DOM |
|---|---|
| `c14:style` 229, `x15ac:absPath` 36, `w:drawing` 14, `w:pict` 10 | `c:style` 229, `xdr:sp` 12, `xdr:graphicFrame` 10, `p:graphicFrame` 2, `p:pic` 2, `p:sp` 2, `x:controls` 2, `xdr:twoCellAnchor` 2 |

No WordprocessingML element is DOM.

### 14.4 jsonix's undeclared `input` (found by the editor; fixed in jsonix 3.4.2)

- **The defect.** `@docx4j/jsonix` 3.4.1's `Base64Binary.decode` assigns an undeclared `input`
  (`jsonix-factory.js` line 4489).
  - Node loads jsonix as sloppy CommonJS, where the assignment makes a global and nothing fails.
  - A strict ESM bundle (the editor's Vite build, vitest's jsdom) throws `ReferenceError: input is
    not defined`, and the part doesn't open: "/word/document.xml: v:shape/@o:gfxdata=...: input is
    not defined".
  - Kept, a Word text box's VML Fallback is unmarshalled, and its `v:shape`'s or `v:oval`'s
    `o:gfxdata` (`xsd:base64Binary`) is decoded. Resolved, the Fallback is dropped first. So phase 4
    is the first load to reach it.
  - Found by the editor (`docx4j-ts-editor-50`) on its first run against this build. It sent jsonix
    the fix, and hand-patched this repository's installed copy for its own runs. That copy is
    restored (`npm ci`), and every number here is on the published file.
- **Measured over the 483 packages:**
  - Globals leaked by load, `unmarshalAll` and save: `input` from 5 packages in the kept load
    (`loadAndSave.docx`, `mc-alternate-content-header.xml`, the three `cr043-*.docx`), none resolved.
  - Strict mode: `node --use-strict` doesn't reach CommonJS, and 3.4.1 passed under it, so it was no
    instrument. A `Module.prototype._compile` hook that prepends `"use strict"` to
    `jsonix-factory.js` makes 3.4.1 throw on exactly those 5 packages, in the kept load only.
- **The fix**, jsonix `231679e` (version `7fa5002`, which adds only the version bump). It is three
  `var`s in the factory: `input`, and `that` in `unmarshalFile` and `unmarshalURL`. It was checked
  here as a candidate with `npm install --no-save` (`dependency-release-verification`):
  - the suite passes but for the test below, which fails as meant;
  - typecheck passes;
  - all 25,704 entry hashes (483 packages, three loads) are unchanged, and so are the warnings;
  - no global leaks in either load;
  - under the strict hook: 0 errors over the 483 packages in either load, and the whole suite passes
    but for the test below.

  The editor's five browser specs over the text-box fixtures pass in its production build with it,
  17 of 17. `npm ci` restored 3.4.1 afterwards. The report went to `jsonix-d2`.
- **Published as 3.4.2** the same day (tag at `7fa5002`). The npm artifact checks out:
  - its integrity matches the lockfile's;
  - its `jsonix-factory.js` and `jsonix.js` are the checkout's at `7fa5002`;
  - with the factory compiled strict, it gives 0 errors over the 483 packages in all three loads,
    with all 25,704 entry hashes and every warning unchanged from published 3.4.1;
  - no global leaks.
- **The lockfile** moved from 3.4.1 to 3.4.2 (`npm update @docx4j/jsonix`; `package.json` unchanged).
- **The test.** `mce.test.mjs` asserted the defect as it was (CLAUDE.md): the kept load leaked `input`
  from those five fixtures, and the resolved load leaked nothing. Against a build of HEAD it failed,
  and it failed again when the candidate was installed, as meant. It now asserts that no load of the
  packages holding the element leaks a global, which a strict consumer would get as a ReferenceError.
- **The floor.** core-ts has no jsonix dependency of its own; it gets jsonix through the objects
  package, whose range is `^3.4.1`. A consumer whose lockfile holds 3.4.1 keeps it after taking
  phase 4, and every Word text box with VML `o:gfxdata` then fails to open in a strict-mode consumer,
  every browser consumer among them. Two ways to make 3.4.2 the floor, for Jason:
  - a direct `"@docx4j/jsonix": "^3.4.2"` in this package's dependencies. It takes effect with this
    release, and npm installs the one copy that satisfies both ranges;
  - the objects package raising its range in its next release, and this package's objects floor
    following it. The objects package's `721eab4` makes its range `^3.4.2` and adds a strict-mode
    `o:gfxdata` round trip to its esbuild bundle test, which fails on 3.4.1.

  Recommended: the second, the way every jsonix floor has come here, with the direct dependency
  only if this package released first.

  **Decided by Jason, 2026-10-18** ("Raise core-ts's objects range to ^0.4.1"): the second.
  - **Objects 0.4.1** was published the same day. Its tag is `10b0d39`: `d888cd5` plus the version
    line. `d888cd5` is `721eab4`, two doc commits and a proposed objects CR, so `modules/` and
    `src/` are as in 0.4.0.
  - **The candidate**, checked here before the release (`dependency-release-verification`): the
    tarball packed from `d888cd5`. Its 846 files other than `package.json` were byte-identical to
    0.4.0's, and the jsonix range was the only difference in `package.json`. With it the suite passed
    (874), typecheck passed, and the 25,704 entry hashes, the 14 warnings and the 297 whole elements
    were unchanged, with the factory compiled strict and 0 errors.
  - **The published artifact**: it differs from that tarball only in its version line, and the same
    checks give the same results.
  - **The raise.** `package.json` requires `^0.4.1`, and the lockfile holds 0.4.1 (integrity
    `sha512-TXfq0s7d…FRGDQ==`) and jsonix 3.4.2. After `npm ci` the suite passes (874), typecheck
    passes, and the survey is unchanged from the locked 0.4.0, entry for entry.

  So a consumer that takes this release gets jsonix 3.4.2 or later.

### 14.5 The editor

- **Told as soon as the flip built**, with section 11.3's list. It ran its suites against this
  working tree, found section 14.4, and did its side on branch `mce-keep` (`53c0ea0` in
  `../docx4j-ts-editor`). The branch merges when it installs the release carrying phase 4:
  - **Import.** A run's element goes through `selectedBranchOf`. An inline drawing in the selected
    branch is an image node, with the element as its source, so the export writes every branch back
    by identity. Anything else is the opaque `AlternateContent` atom, as before.
  - **Image lookups** go through one helper. `walkStory` and `textOfElement` take the selected branch.
  - **Its diagnostics pin** is regenerated: `loadAndSave`'s `drawing: 1` is `AlternateContent: 1`.
    It also gains a test with a picture inside a Choice.
  - **The XPath placement** needed nothing. Its companion's find had a defect that predates this phase
    (a hit in a text box climbed to the box's paragraph, which `addressOf` doesn't address), and it is
    fixed. ED-005 section 12.106 records that the XML pane's mismatch is gone (section 11.3).
- **Its numbers**, with jsonix 3.4.2:
  - typecheck clean;
  - model 1327 (3 new);
  - editor unit 214 of 218 (the four failures known from the 0.5.0 review);
  - companion 32;
  - browser specs over the text-box fixtures 17 of 17.
- **Release order, its view.** Phase 4 inside 0.5.0 if jsonix 3.4.2 ships this week: one install, one
  floor raise. Otherwise 0.5.0 first and phase 4 as 0.6.0, since nothing in the editor needs phase
  4 to ship. Jason decides.

### 14.6 Tests

- **`mce.test.mjs`, 5 new:**
  - the default keeps, `mcePreprocess: false` is the same, and `mcePreprocess: true` resolves; a
    clone loads as its original;
  - the keeping pass reports what resolving reports, nested elements included, and leaves the DOM
    as it was;
  - over the 16 fixture packages holding the element, a kept load's warnings equal a resolved
    load's (8 warnings);
  - `insertOoxml` into a loaded document writes the element as given;
  - section 14.4's leak, asserted as it was on 3.4.1, and since 3.4.2 that no load leaks a global.

  Against HEAD's build, with `mceKeep` shimmed, the first four fail, and so does the flat OPC
  header test, which now covers both loads. The leak test failed there in its 3.4.1 form. In its
  final form it passes there, since HEAD's default load resolves and decodes no Fallback.
- **Changed to say which load they mean:**
  - `kept-branches.test.mjs`. Its resolved side was the default, so after the flip it compared kept
    with kept. It is now `mcePreprocess: true`.
  - `parity.test.mjs`. The four fixtures holding the element run the default (kept) and once more
    resolved: 261.
  - `xpath-parity.test.mjs`. Every fixture loads the default way: 64.
  - the label oracle's walk in `lists.test.mjs`.
  - In `mce.test.mjs`, the header and the three chartex frames run in both loads.
  - `omml.test.mjs`, `sml-extensions.test.mjs` (the controls, the slicer drawings, the legacy
    colours), `content.test.mjs` (the slicer's path) and `tracking.test.mjs` (CR-043's count).
- **Suite:** 874 passed; typecheck passed.

### 14.7 Not done

- **The release.** Both floors are in place: jsonix 3.4.2 (section 14.4) and the objects fix of
  section 14.8, through objects `^0.4.2`. Then the editor installs, and the objects session reruns
  its sample project on the release.
- **`p:sp` and the `xdr:` shapes in a kept branch stay DOM.** docx4j types `p:sp` there. The
  objects package's CR-010 (proposed 2026-10-18, `d888cd5`) declares it and `p:cmAuthor` globally,
  as CR-009 did WordprocessingML's. Its `xdr:` anchors wait on docx4j binding
  `a14:legacySpreadsheetColorIndex`. `placeholdersOf` would then read a kept `p:spTree` through
  `selectedBranchOf`, in the release that takes it. Nothing reads them today, and no consumer has
  asked.
- **`worksheet.controls`** is typed only resolved (section 14.2).
- **The read-time warning moves to CR-012**, the text-model CR, proposed the same day (Jason, 2026-10-19:
  "yes, move the read-time warning to CR-012 and commit"). That is the paper's section 4.1 form,
  with a revision view and a hidden-text rule. Those rules are the text model's: hidden text is the
  extraction specification's REQ-507, and the views are the model's. The math comparison section 9.2
  leaves open goes with it. The load-time form covers what the paper's warning asks, so CR-010 closes
  without it.
- **Readers use load's default understood set** (section 13.4). A caller's own
  `createMcePreprocessor({ understood })` warns by its set while readers keep theirs.
- **Text boxes in search, `replaceText`, bindings and paraIds**, and the `textbox:N` view, go to the
  text-model CR (D8, section 13.4).

### 14.8 A kept `mc:Choice` loses the declaration its `Requires` names (found 2026-10-19)

- **Found** answering docx4j-25's question from its CR-049 (`ba40c53f4`): does a marshal declare a
  root's `mc:Ignorable` prefixes and its Choices' namespaces? The first half is fine. Each marshal
  copies its own prefix table, and `mc:Ignorable` comes from the root being written, so no state is
  shared between marshals. The second half isn't.
- **The defect is the objects package's.** Its `fixRootNamespaceDeclarations` removes each root
  declaration that no element or attribute uses, and keeps only what `mc:Ignorable` names. A
  Choice's `Requires` names prefixes too, and doesn't count as a use. So where the Choice's content
  is in another namespace than the one `Requires` names, the declaration goes. Example:
  `waterfall.docx`'s `Requires="cx1"`, whose content is `cx:` (2014/chartex). The saved part then
  names a prefix nothing declares.
- **Measured** over the 483 packages: every `mc:Choice` in a re-marshalled part, checked with
  `lookupNamespaceURI`.

  | Load | `Requires` prefixes | Undeclared | In |
  |---|---|---|---|
  | `mcePreprocess: true` | 265 | 0 | |
  | default (kept) | 297 | 18: `cx1` 11, `x14` 4, `a14`, `sle15` and `tsle` 1 each | 11 packages |

  The `cx1` cases are the waterfall docx, pptx and xlsx, the chartex `manual-bin*` files and the
  three CR-047 pareto files. The `x14` cases are the worksheet controls of the two `cr022-checkbox`
  files; the `a14`, `sle15` and `tsle` cases are `cr022-slicers-timelines`.

  So this is new with phase 4. Before it, these elements were resolved and no Choice was written.
  Phase 4's survey (section 14.3) compared the elements canonically, without namespace
  declarations, so it couldn't see this.
- **What it costs.**
  - On reload, the keeping pass treats such a Choice as not understood, since its prefix doesn't
    resolve. The content API's readers match the `Requires` string by prefix and still take the
    Choice, so the two disagree.
  - docx4j has declared Choice prefixes on save since 6.0.0. `JaxbXmlPart`'s `mcChoiceNamespaces`
    javadoc says: "These must be declared, or Word can't open the docx". That is docx4j's record, not
    a measurement here. No Office application has opened one of these files.
  - docx4j-python has the same gap (its CR-011 Q11).
- **The fix goes to the objects package** (sent 2026-10-19): count a Choice's `Requires` prefixes as
  used, and declare at the root any still missing, from the prefix table. That is docx4j CR-049's
  contract. Every one of the 18 prefixes is in `NAMESPACE_PREFIXES`.
- **The test.** `mce.test.mjs` asserted the defect as it was (CLAUDE.md): 10 of the 33 prefixes the
  fixtures' kept load writes (`cx1` 3, `x14` 4, `a14`, `sle15` and `tsle` 1 each), and none in the
  resolved load. The fix announced itself by failing it, and since the raise below it asserts none
  undeclared in either load.
- **The candidate, `33a7f44`** (objects CR-001 section 3.2): a Choice's `Requires` prefixes count as
  used, and any still undeclared is declared from the table. A `Requires` prefix nothing resolves is
  kept, with a warning. Checked here on 2026-10-19 with the tarball packed from it:
  - 0 of 297 undeclared in the kept load, and 0 of 265 resolved;
  - of the 25,704 entry hashes, 30 differ: 15 parts in each of the default and kept loads, none in
    the resolved load. Each of those parts is byte-identical to 0.4.1's apart from its root start
    tag, which gains exactly the missing declarations (`cx1` 11, `x14` 2, `sle15` 1, `a14` with
    `tsle` 1);
  - 0 errors with the factory compiled strict, and the warnings unchanged;
  - the suite passes but for the test above, which fails as meant; typecheck passes.

- **Objects 0.4.2** was published the same day. Its tag is `29cd1cc`: `33a7f44` plus the version
  bump. Checked here on the npm artifact:
  - its tarball's 847 files are the candidate's, apart from the version in `package.json`;
  - of the 25,704 entry hashes, none differs from the candidate's, and the same 30 differ from
    0.4.1's; 0 errors with the factory compiled strict;
  - 0 of 297 undeclared in the kept load, and 0 of 265 resolved.
- **The raise.** `package.json` requires `^0.4.2`, and the lockfile holds 0.4.2 (integrity
  `sha512-rRINQ9FA…GShLAlA==`) with `@docx4j/jsonix` 3.4.2. A lockfile holding 0.4.1 no longer
  satisfies the range, so no install of this package's release keeps the defect. The test above
  asserts none undeclared; the suite passes (875) and typecheck passes.

### 14.9 A `Requires` prefix the prefix table lacks loses its declaration (found 2026-10-20)

- **Found by the editor** (`docx4j-ts-editor-50`), building ED-009's fixture. The fixture's first
  version had a Choice with `Requires="zz9"`, and `xmlns:zz9` declared on `w:document`. A save
  through the engine wrote the root without the declaration. The fixture now uses `x14ac`, which the
  prefix table has.
- **Reproduced** on `f8617e8`, with objects 0.4.2. The source root carries
  `xmlns:zz9="urn:example:zz9"` and `mc:Ignorable="zz9"`, and a run holds an element with
  `Requires="zz9"`. After one `insertParagraph` and a save:
  - the root declares `w` and `mc` only;
  - `Requires="zz9"` is kept, so the part names a prefix nothing declares;
  - `zz9` is dropped from `mc:Ignorable`;
  - the marshal warns of both ("...which no declaration or NAMESPACE_PREFIXES entry resolves;
    kept", "...; dropped").

  With `mcePreprocess: true` the Fallback replaces the element, so no Choice is written. The
  Ignorable entry is still dropped, as it was before 0.5.0. So the `Requires` half is new with
  phase 4.
- **A real fixture has it.** docx4j's `text-extraction/extraction-probe.docx` (`4a89a5967`) has
  four Choices whose prefix the table lacks: `zz9` three times, and `shape2010`, a second prefix for
  the `wps` namespace. It declares each on its own `mc:AlternateContent`, not on the root. Kept,
  edited and saved, all four resolve nowhere. CR-012's phase 0 copies these probes into this
  repository.
- **Over the corpus**, 834 elements with a Choice in 719 packages (this repository's fixtures and
  docx4j-core-tests' resources). Only those probes have a `Requires` prefix that
  `UNDERSTOOD_PREFIXES` lacks: `zz9` 12, `shape2010` 4 and `wpi` 1. The `wpi` probe saves clean.
- **The cause is the objects package's.** The object tree keeps none of the source's namespace
  declarations. `fixRootNamespaceDeclarations` declares a `Requires` or `mc:Ignorable` prefix only
  from `NAMESPACE_PREFIXES` and `IGNORABLE_PREFIX_ALIASES`, so a binding only the source had is
  lost. Section 14.8's fix, the objects package's CR-001 section 3.2, can't reach it. For
  `shape2010`, the table can't help even though it knows the namespace, since it gives that
  namespace's prefix as `wps`.
- **Sent to the objects package** (`docx4j-generated-objects-ts-82`, 2026-10-20) at Jason's word ("Send
  this declaration defect to the objects session as a CR there"). The request is to keep the declarations
  of the part it unmarshalled, on the root and below it, and fall back to them for the prefixes a
  `Requires` or an `mc:Ignorable` names. It's preferably all inside the facade, since this package
  touches no prefixes.
- **Objects CR-011** (`e311ef1`, proposed 2026-10-20) takes it. The facade records the source's
  bindings at unmarshal: what is in scope at the root, and at each `mc:Choice` the binding of every
  prefix its `Requires` names. At marshal it finds the record from the element or its nearest
  `PARENT`. A prefix resolves from the source first, then the table, then the aliases. The source
  comes first because three VML parts in its corpus bind `x` to the Excel office namespace, where the
  aliases give SpreadsheetML. Its prototype brings back this section's repro and the probe with
  every prefix bound, and leaves its 341 corpus hashes as 0.4.2's.
- **Two things for this package, from the objects session:**
  - **`XmlPart.recordDeclarations` and `declareIgnorablePrefixes`** (CR-001 section 17) have had
    nothing to do since objects 0.1.5. `marshalNode`'s own pass declares or drops every
    `mc:Ignorable` prefix before `marshalToNode` calls the second, so it finds none undeclared.
    Confirmed here: the repro logs the facade's "dropped", never this package's own warning. Both
    go when the range is raised to the release carrying objects CR-011.
  - **Inserts.** `insertOoxml` brings in a Choice whose prefix only the inserted OOXML declares, as a
    bare fragment and as flat OPC. Each is unmarshalled on its own and its content moved into the
    destination tree, so a record keyed by the fragment's root isn't found from there. Measured:
    both forms save the Choice with `zz9` bound nowhere. Suggested to the objects session: key the
    per-Choice record by the Choice's own value, so it moves with it and `deepCopy` can carry it.
    Its own suggestion, an `adoptSourceBindings(from, to)` that `ooxml.mts` would call, works too.
    Its call.
- **docx4j**, read, not run: `Docx4jUnmarshallerListener` records each Choice's `Requires` as
  prefixes on the part (`JaxbXmlPart.addMcChoiceNamespace`, since 6.0.0), and the marshal declares
  them through its prefix mapper. Whether that mapper can declare a prefix its own table lacks
  wasn't checked.
- **The test.** `mce.test.mjs` asserts the defect as it is (CLAUDE.md), over `zz9` on the root,
  `zz8` on the element and `shape2010`. Each `Requires` is kept and nothing declares it, and `zz9`
  leaves `mc:Ignorable`. The fix announces itself by failing it.

## 15. The readers' branch: the `Requires` gate alone (the editor's ED-009, 2026-10-20)

**Status:** Proposed 2026-10-20 at the editor's request (ED-009 section 3.8, its proposal P7). Taken
as the editor asked (Jason, 2026-10-20: "do as the editor asks re Q1, Q2"). Not scheduled.

### 15.1 What the editor measured

Jason's probe paragraph has one run-level element with ten Choices and a Fallback. The Choices
require `w14`, `w15`, `w16se`, `w16cid`, `w16`, `w16cex`, `w16sdtdh`, `w16sdtfl`, `w16du` and `wp14`,
in that order, and each holds its own prefix as text.
- `selectedBranchOf` takes `w14`, the Choice Word 365 shows.
- `UNDERSTOOD_PREFIXES` lacks `w16sdtdh`, `w16sdtfl` and `w16du`, which the editor reports Word 365
  understands. All three are in `NAMESPACE_PREFIXES`. No generated module binds them, which is why
  `UNDERSTOOD_NAMESPACES` doesn't have them.
- So with one of the three first, Word shows it and this package takes the next Choice it
  understands. The editor's `mc-branches.docx` case 7 (`w16du` first) reads "I prefer w14" here, resolved.
  Word's display is the editor's report from Jason's probe; it wasn't checked here.

### 15.2 The rule

For the branch the readers take (`selectedBranchOf`), a Choice is understood when:
- every namespace its `Requires` names is one Word 365 understands; and
- everything inside it is in namespaces the model binds, so it unmarshals typed and round-trips.

So the gate alone decides, and section 11.5's round-trip guard (docx4j CR-021 section 8.9) moves
from the `Requires` namespaces onto the Choice's content. In the probes the prefix is only a gate,
and the content is plain WordprocessingML.

- **Kept loads, the default.** The choice decides only what the readers read and the mutators edit.
  The save writes every branch as loaded, so a wider set can't lose content.
- **`mcePreprocess: true`.** Today it uses the same set to choose which branch replaces the element,
  and that converts DOM into typed content (CLAUDE.md). Under the content condition, the guard is
  checked on what is converted. Content the model binds only partly still loses the unbound part
  (CR-004 section 7). So resolving keeps `UNDERSTOOD_NAMESPACES` until the round trip of such a
  Choice is measured. The rule here is for kept loads first.
- **Which namespaces Word understands.** Only the three are measured. The list goes with section 3's
  item 6, D4 item 2's published list.
- **Matching by URI** (D4 item 1, an objects change) is related. `shape2010` in docx4j's extraction
  probe is the `wps` namespace under another prefix (section 14.9), and today's prefix match
  doesn't take it.

### 15.3 What it changes

- **In the corpus, nothing.** No `Requires` in its 834 elements names `w16du`, `w16sdtdh` or
  `w16sdtfl` (section 14.9's count), so no fixture and no parity golden changes branch.
- **docx4j.** It's a departure from CR-021 section 8.9, to raise with docx4j through the portfolio
  when this is scheduled. The parity harness's `preferChoice` (section 11.5) has to follow, so docx4j
  reads the branch this package reads.
- **The editor** shows whatever `selectedBranchOf` selects, and records the difference meanwhile in
  its `docs/word-divergences.md`.
