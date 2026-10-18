# CR-010: Keep every branch of `mc:AlternateContent` (docx4j CR-021's policy)

**Status:** Proposed 2026-10-17, at Jason's request ("Shouldn't we be preserving all branches of
mc?"). Phase 1 implemented 2026-10-17 (section 9; D2, D3 and D4 decided by Jason that day, and
D1 settled by the text extraction paper's D4). Phase 2 implemented 2026-10-18: the objects
package's CR-009, released in its 0.4.0, which this package now requires (section 10). Phase 0
implemented 2026-10-18 (section 11): the walker audit, the editor's exposure, the XPath and parity
measurements, and the survey as a test in `mce.test.mjs`; D5 and D6 decided by Jason that day,
and D5 implemented (section 11.8).
Phase 3 implemented 2026-10-18 at Jason's request ("do CR-010 phase 3"; section 13). Phase 4 is not
scheduled and waits for Jason's go. The policy is docx4j's CR-021 (done 2026-09-19, in 17.2.0).

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
     phase 4.
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
