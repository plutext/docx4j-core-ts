# CR-010: Keep every branch of `mc:AlternateContent` (docx4j CR-021's policy)

**Status:** Proposed 2026-10-17, at Jason's request ("Shouldn't we be preserving all branches of
mc?"). Phase 1 implemented 2026-10-17 (section 9; D2, D3 and D4 decided by Jason that day, and
D1 settled by the text extraction paper's D4). Phases 0, 2, 3 and 4 are not scheduled; each waits
for Jason's go. The policy is docx4j's CR-021 (done 2026-09-19, in 17.2.0). Phase 2 needs work in
the objects package, as a CR there.

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

0. **The audit** (half a day):
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
     phase 4.
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
  before the default flips.

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
  `mcePreprocess: false`, and compares every element's normalised form; phase 0 makes it a test.

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
    Fallback's linear `a:t` are two notations of one formula. The read-time form (phase 3) can
    compare math by its own rule, and this one is left as it is.

### 9.3 Not done in phase 1

- **The worksheet's `x14` controls** stay resolved (section 4).
- **The content API** doesn't read a kept branch, so the text of an element nothing can be
  selected from appears nowhere until phase 3.
- **No `warnings: false` option** (the paper's section 4.1 item 6). The load-time comparison is
  one pass over branches that are dropped in any case.
- **The warnings cover only parts parsed so far.** A caller wanting a whole package's warnings
  reads its parts first; private reads keep them untouched.
