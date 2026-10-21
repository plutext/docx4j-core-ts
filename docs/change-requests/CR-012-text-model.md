# CR-012: The text model: docx4j CR-046's text extraction (views, scope, plain text, Markdown, the block model), held to the specification's conformance suite

**Status:** Proposed 2026-10-19, at Jason's request ("open a text-model CR here (That would be
CR-012)"), on docx4j CR-046 phase 4's hand-off (the docx4j session, 2026-10-18). Decisions D1 to D9
(section 6) were taken by Jason on 2026-10-21, all as recommended ("CR-012 all decisions as
recommended"; section 6.1). Phases 0 and 1 implemented 2026-10-21 (Jason: "do CR-012 phases 0 and
1"), not released: the suite copied from docx4j `198d015a3` (WD 0.28's) with its runner and the agreement
test's, every row pending (section 13); REQ-510's level owner in the `Emulator` (section 14). Phase 2
implemented 2026-10-21 (Jason: "start CR-012 phase 2"), not released: the plain-text model, `Texts.of` and the
block model in `src/model/text/`, its views a plan over the unresolved tree (`ResolvedView`), held to the
suite's 735 rows it claims and the agreement test's 588 comparisons with no difference (section 15). Phase 3
implemented 2026-10-21 (Jason: "start CR-012 phase 3"), not released: Markdown, `MarkdownExporter` with
`blocks`, docx4j's `WmlToMarkdown` over the block model and a port of commonmark-java's renderer, so the
Markdown is docx4j's byte for byte; 786 suite rows pass, and 1,320 of 1,420 exports over `test/` equal
docx4j's, the rest differing only where phase 5 or REQ-303 does (section 16). Phase 4 implemented
2026-10-21 (Jason: "start CR-012 phase 4"), not released: the readers, `Paragraph.text` and `search` in
Office JS's characters and rules in Word's default view, equal to Word 365's answers on checks 48, 48b and 51;
ED-009's branch content read and edited; text boxes at `textbox:` addresses, read-only; `pkg.paragraphAt`
over notes and comments; `outline()` on the model and async (the editor asked, Jason chose); three defects
in 0.5.0 found and fixed (section 17). Word
check 48, run 2026-10-21, widens D6 (section 6.2): `Paragraph.text` and `search` follow Office JS in
all of it, in Word's default view, Simple Markup with hidden text not shown (Jason, the same day, after
check 48b). D2 starts from his decision of 2026-10-19, that the
Office JS readers keep calling the objects package's `textOf` and that it is not deprecated. It
takes over CR-010's text-box view (D8), the U+0002 rule (section 12.6) and the read-time
`mce.branches_differ` warning (section 14.7). The editor reviewed it the same day (section 11), and
its asks are in section 4.4. Section 12 (2026-10-20) takes the editor's ED-009 rule for run content
in a kept `mc:AlternateContent`, reading and editing, as the editor asked (Jason: "do as the editor
asks re Q1, Q2"), in phase 4. It also records a defect in 0.5.0 that the rule closes: edits over a
paragraph-level element's text land outside it. Until then such an edit throws `NotSupportedError`
(section 12.4, Jason: "yes to all 3 questions"), implemented 2026-10-20, not released; superseded by phase 4,
which reads and edits the branch (section 17.3).

## 1. Why

The text extraction specification (portfolio `docs/text_extraction_specification_wd.md`, WD 0.16,
`dae90bf`, 2026-10-19) says what text a document holds:
- which branch of an `mc:AlternateContent` is read;
- the three revision views;
- which stories are read and in what order;
- how characters, tables, fields, notes, hidden text, text boxes and embedded content become text;
- the plain-text, Markdown and block-model forms.

docx4j implements it as its CR-046 (done 2026-10-18 for 17.3.4: `org.docx4j.model.text`, 27 files).
docx4j-python's CR-011 ports it, proposed. The specification becomes 1.0 when three implementations
pass its conformance suite. This package is meant to be one of them (the paper's section 10).

This package has no text model in the specification's sense. It has two text readers, and neither
is one:
- **The objects package's `textOf`**, docx4j `TextUtils`'s counterpart. It reads the accepted view
  only and joins paragraphs with `\n`. It reads no text box and no note, keeps hidden text, and has
  no notion of fields beyond their results. Its `mc:AlternateContent` choice uses its own prefix
  set (`CONVENTIONAL_PREFIXES`), not `UNDERSTOOD_PREFIXES`.
- **`segmentsOf`** (`src/model/content/tree.mts`). It reads one paragraph's runs in the accepted or
  original view. It has no `markup` view, doesn't join a paragraph across a deleted mark, and
  doesn't enter a text box.

The paper's section 1 table records what that costs today:
- `Body.text` is accepted only, doesn't join a deleted mark, reads no text boxes, emits nothing for
  math, and has no notes;
- a cell is a line.

The editor's ED-008 (proposed) is to put the model's options on its `outline`, `find`, `read` and
`extract_text` verbs, and it has "no text model of its own". It depends on this CR.

## 2. What docx4j built (read 2026-10-19, docx4j `VERSION_17_3_4` at `3dcbb3764`, unpushed)

- **The facade, `Texts`:**
  - `of(pkg[, TextOptions])`, plus overloads for one part (main, header or footer) and for a content
    object;
  - `ofResolved(pkg, opts)`, the copy route of the agreement test;
  - `plainText(pkg, opts)`;
  - `normalize(text)`, REQ-1201's normalisation.
- **The walk,** `Extractor`: one walk does the four layers (resolve, view, scope, serialise) in
  REQ-201's order.
- **`TextView`:** `ACCEPTED` (the default), `ORIGINAL` and `MARKUP`.
- **`TextOptions`:** the specification's section 10 names, through `of(Map)` and `set(name, value)`.
  The defaults:

  | Option | Default |
  |---|---|
  | `view` | `accepted` |
  | `include` | none (it takes `headers`, `footers`, `comments`, `glossary`) |
  | `text_boxes`, `notes`, `note_references`, `placeholders`, `charts`, `smartart`, `altchunks`, `warnings` | true |
  | `comments`, `hidden`, `field_codes`, `toc`, `list_labels`, `embedded`, `alt_text`, `metadata`, `page_markers` | false |
  | `fields` | `cached` (or `update`) |
  | `bindings` | `bound` (or `surface`) |
  | `symbols` | `asis` |
  | `math` | `latex` |
  | `tables` | `auto` (or `tsv`, `grid`) |
  | `markup_detail` | `type` |
  | `layout` | `none` (or `indent`) |

  It refuses `symbols: unicode` and `math: unicodemath`, naming REQ-1001.
- **The result, `TextResult`:** `blocks`, `comments`, `warnings`, `metadata` and `options`.
  - A `TextBlock` is a paragraph, table or chart. Its story is main, text box, footnote, endnote,
    header, footer, comment, glossary, SmartArt or embedded.
  - A block carries:
    - address and paraId, style, runs, fields, rows;
    - the hidden, placeholder and decorative flags, and the revision count;
    - indents and alignment;
    - `listLabel`, outline level, and `TextNumbering` (`numId`, `level`, `format`, `levelText`,
      `start`, `value`, `abstractNumId`);
    - for a note, its type, number and id; for a chart, its title, header, rows and description.
  - A `TextRun` is text, a field code, a note reference or mark, a comment reference, a list label,
    an image, an object or a page marker. In the markup view it carries
    `TextRevision{type, author, date}`.
  - `TextWarning{kind, address, branch, contentType, partName}` has six kinds:
    `mce.branches_differ`, `field.no_result`, `sdt.binding_differs`, `altchunk.not_expanded`,
    `content.not_read` and `image.not_described`.
- **Serialisers:**
  - `PlainText.of(result)`;
  - Markdown in `docx4j-markdown`'s `MarkdownExporter.export(result[, pkg])`;
  - HTML and FO take a `TextView` setting, not the result;
  - JSON is in docx4j-mcp only.
- **Elsewhere in docx4j:**
  - `McPurpose.TEXT` and `McSelection.selectedBranch`;
  - the published understood set, `docx4j-core/src/main/resources/org/docx4j/jaxb/understood-namespaces.txt`;
  - `ResolvedView`, the accepted and original rules;
  - `OmmlToLatex`, moved into core.
- **`TextUtils`** is deprecated with its output unchanged (CR-046 section 2.7). No shipped reader
  uses it.
- **What needs the engine:**
  - **Hidden text** (REQ-507): through `PropertyResolver.getEffectiveRPr` and
    `getEffectiveParagraphMarkRPr`.
  - **`fields: update`** (REQ-506): `FieldValues`, over `FieldUpdater`'s resolvers and
    `FormattingSwitchHelper`.
  - **List labels** (REQ-510): `Emulator.getNumber` per story, and `Emulator.levelStyleId` for
    ownership (`d9acdf5e2`).
  - **Embedded content** (REQ-512): xlsx4j's `DataFormatter` for a sheet's used range, pptx4j for
    slides, and `Ole10Native` for a `.txt` package.
  - **The TOC** (REQ-520), and **notes**, numbered in reference order with footnotes and endnotes
    apart.
  - **Text boxes** after their anchor, from the TEXT branch.
- **The suite** (`docx4j-core-tests/src/test/resources/text-extraction/`):
  - Fixtures in `phase0/`:
    - `extraction-probe.docx`: cases 1 to 49, 53 and 54;
    - `extraction-probe-external.docm`: cases 50 to 52;
    - `extraction-probe-structure.docx`: cases 55 to 57.
  - Twelve goldens: per fixture, `.accepted.txt`, `.original.txt`, `.markup.txt` and `.md`.
  - Five TSVs:

    | File | Rows |
    |---|---|
    | `expected-tokens.tsv` | 188 |
    | `case-statements.tsv` | 57 |
    | `expected-findings.tsv` | 26 |
    | `expected-order.tsv` | 2 |
    | `goldens.tsv` | 12 |

  - `GOLDENS.md` is the adapter contract: a scenario's options are `; `-separated and each needs its
    statement, and a `default` golden needs "xhtml importer absent".
  - The engine contract is `ExtractionConformance.Engine`:
    `implementedStatements()` and `extract(fixture, view, scenario, format)`, returning
    `{text, findings}`.
  - A row whose statements aren't all implemented is pending.
  - docx4j's own run: 972 pass, 0 fail, 51 pending. 42 of the pending need the inspector and 9 an
    XHTML importer.
- **`ViewAgreementTest`** runs over the 128 CR-043 fixtures (`revisions/`) and the 13
  `text-extraction` documents. For the accepted and original views, in the default scope and with
  everything included, it requires the walk's plain text to equal, exactly, the text of a copy
  resolved by `TrackedChanges.acceptAll` or `rejectAll`.
- **`LevelOwnerTest`** holds three Word-checked documents, with Word's labels read off its PDFs.

## 3. What this package has to build on

- **The engine:**
  - `PropertyResolver.getEffectiveRPr(rPr, pPr?, cellContext?)` resolves `w:vanish` through the
    style chain as a toggle, and `getEffectiveParagraphMarkRPr` resolves the mark;
  - the `Emulator`, with per-story `NumberingStates` (`walkStory` in `List.mts`);
  - `cellContextOf`;
  - `selectedBranchOf`, CR-010's reading rule.
- **Stories:**
  - `revisedStories()`, which gives docx4j's order: main, each footnote, each endnote, each comment,
    then the headers and footers;
  - `main.headerParts` and `footerParts`;
  - `footnotesPart`, `endnotesPart` and `commentsPart`;
  - `glossaryDocumentPart`, which is unread today.
- **Package-level verbs:**
  - `pkg.clone()`;
  - `pkg.acceptAllRevisions()` and `rejectAllRevisions()` over every story. These are the agreement
    test's copy route.
- **Private reads.** `XmlPart.readContents()` reads a part without marking it unmarshalled. A save
  after extraction then stays byte-identical, which is what REQ-204 ("never modify the package")
  means here.
- **Parts and readers that exist but don't do what the model needs:**
  - `AlternativeFormatInputPart`, `EmbeddedPackagePart`, `OleObjectBinaryPart` and
    `GlossaryDocumentPart` exist as classes. Nothing reads them.
  - Field codes are not parsed anywhere. `instrText` has no text, and `fldSimple` is a run holder,
    so every reader sees a field's result.
  - Nothing reads note numbers.
- **The understood set.** `UNDERSTOOD_NAMESPACES` holds 107 URIs, exactly the first 107 of docx4j's
  published 118, in order (measured 2026-10-19). The other 11 are:
  - WordprocessingML's `sdtdatahash`, `word16du` and `sdtformatlock`;
  - `drawing/2016/11/diagram`;
  - SpreadsheetML's `ac` and six revision namespaces.

  Readers match `Requires` by prefix (`UNDERSTOOD_PREFIXES`). The object tree keeps no namespace
  declarations in scope, so the prefix is all there is.
- **Strict documents** are not converted (CR-001 section 17.6).
- **Addresses** (CR-002 section 3.3):
  - `body/3`, `body/3/1/0/2`, `header:rId3/0`, a paraId, or `{contains}`.
  - `pkg.paragraphAt` resolves `body`, `header:` and `footer:` only. A `footnote:`, `endnote:` or
    `comment:` address gives undefined, although the bodies carry those prefixes.
  - There is no `textbox:`.
- **What Word was seen to do** (check 42, Word 16.0.20326.20232):
  - Office JS's `body.paragraphs` leaves a text box out, and so does `body.text`.
  - Each footnote or endnote reference reads as U+0002 in `paragraph.text`. A comment reference
    reads as nothing.
  - Here `Paragraph.text` gives nothing for either.
  - Hidden text in Office JS's `text` has not been measured.

## 4. Proposal

### 4.1 The model

A module `src/model/text/` (docx4j's `org.docx4j.model.text`), exported from `./model` and `.`:

```ts
const result = await Texts.of(pkg, { view: 'markup', include: ['headers', 'footers'], hidden: true });
result.blocks;    // TextBlock[]: address, paraId, story, style, runs, rows, numbering, flags
result.warnings;  // TextWarning[]: mce.branches_differ, field.no_result, ...
PlainText.of(result);           // REQ-701: the plain text
MarkdownExporter.export(result); // REQ-703 and REQ-704
```

The overloads (section 4.4): `Texts.of(part, options?, context?)`, for a main, header, footer, notes
or comments part, and `Texts.of(pkg, content, options?, context?)`, for a content object.

- **`Texts.of` is async**, because parts unmarshal lazily. Every part it reads that wasn't already
  unmarshalled is read privately, so the package is unchanged (REQ-204).
- **`Texts.ofResolved(pkg, options)`** is the copy route: `clone()`, then `acceptAllRevisions()` or
  `rejectAllRevisions()`, then a walk with no view logic.
- **The walk** is one pass in REQ-201's order. Concretely:
  1. The branch is `selectedBranchOf`'s (REQ-301 to 305), so the text model and the content API read
     the same branch.
  2. The view is a walk over revision holders, with the paragraph join at a deleted mark (REQ-402
     to 405).
  3. The scope (REQ-501 to 520):
     - the main story, with text boxes after their anchor;
     - footnotes, then endnotes;
     - headers, footers, comments and the glossary under `include`;
     - notes as `[n]`;
     - hidden text out by the effective `w:vanish`;
     - fields cached or as codes, and the TOC under `toc`;
     - placeholders and bindings, images, properties;
     - tables per REQ-605.
  4. The characters per REQ-601 to 606, with no trimming.
- **The result and its parts** use docx4j's names: `TextResult`, `TextBlock`, `TextRun`, `TextRow`,
  `TextCell`, `TextField`, `TextNumbering`, `TextRevision`, `TextWarning` and `ImageDescriber`. Their
  fields are camelCase, and the block model's JSON (REQ-706) is the same objects.
- **Options** take the specification's section 10 names verbatim (D1). An option this package
  doesn't implement yet is refused with the statement named (REQ-1001). Section 5 lists them.
- **The read-time `mce.branches_differ`** (from CR-010 section 14.7) is a warning of the walk.
  - It compares the branches' text under the view and the hidden rule, so a branch whose only
    difference is hidden or deleted text no longer warns.
  - Math compares by its own rule, so the equation CR-010 section 9.2 recorded (an OMML Choice
    against its linear Fallback) is no longer a false positive.
  - Load's warning stays as it is.
  - *Superseded 2026-10-21 (section 15.6, item 2): Jason ruled docx4j's rule for WD 0.29, every `w:t`,
    `m:t` and `a:t` whatever the view, hidden text included. The equation is a workbook's, which the
    text model doesn't read.*
- **The understood set** is published as `UNDERSTOOD_NAMESPACES` (REQ-306). Since CR-010 section 15
  the readers also count `READ_ONLY_UNDERSTOOD_NAMESPACES`, and the walk takes `selectedBranchOf`'s
  branch, so the extractor's set is the two together. A test asserts it against docx4j's list.
  - It equals 111 of the list's 118 URIs: the first 107, then lines 109 to 112 (`w16sdtdh`, `w16du`,
    `w16sdtfl` and `x14ac`), which Word 365 understands.
  - Phase 0 measured, over the corpus, whether any `mc:Choice` requires one of the other 7. None
    does (section 13.4), so the 7 are documented as the difference, and a test names them.
  - Had one done, adding it to the readers would have needed a Word check, and adding it to load
    CLAUDE.md's round-trip test first.
  - REQ-303's prefix table is documented as this package's matching.

### 4.2 What the existing readers do (D2 and D3)

| Reader | After this CR |
|---|---|
| `Paragraph.text`, `Body.text`, `Range.text`, `Table.text`, `TableCell.text`/`value` and `values`, `ContentControl.text`, `Comment.content`, `TrackedChange.text`, `search`, `replaceText`, and `getText({ view })` | Office JS's text, unchanged: `textOf` and `segmentsOf`, held to Word checks. Documented as not the text model (D2). |
| `Body.outline()`, `pkg.outline()` | The text model's blocks (D3). The paragraph text is the model's. A text box's paragraphs are listed after their anchor, flagged and addressed `body/3/textbox:0/1`. It takes `TextOptions`. |
| `Body.listLabels()`, `ListItem.listString` | The Emulator, with REQ-510's level ownership (D9). |

### 4.3 Addresses (D4 and D5)

REQ-519's segment kinds after the anchoring paragraph's address:
- `body/3/textbox:0/1` (legacy WordArt too);
- `body/3/chart:0`;
- `body/3/smartart:0/1`;
- `body/3/object:0/…` and `body/7/altchunk:0/…`, each followed by the inner address (`body/0`,
  `sheet:0`, `sheet:0/table`, `slide:0/2`);
- `glossary:0/1`.

They are the block model's addresses. Under D4:
- `paragraphAt`, `elementAt` and `addressOf` accept `textbox:`;
- `pkg.paragraphAt` resolves `footnote:`, `endnote:` and `comment:` as well, which it doesn't today.

A `Paragraph` at a text-box address reads the selected branch's `w:txbxContent`. Under D5 its
mutators throw `NotSupportedError` in this CR.

### 4.4 What the editor's review adds (section 11)

- **Source references.** Under D2, a block's text and `Paragraph.text` differ: hidden text, `[n]`, a
  joined mark, a text box. So a span counted in `Paragraph.text`, which is what the editor's verbs
  take, isn't an offset in the block.
  - Each `TextBlock` carries `source`: its paragraph's element. A joined block's `source` is the
    first paragraph, as its paraId is (REQ-405).
  - Each `TextRun` carries `source`: its run element and paragraph. Where the run contributes to
    Office JS's text it also carries `officeOffset`, its offset in that paragraph's
    `Paragraph.text`.
  - A run only the model writes has no `officeOffset`. That covers a `[n]`, a list label, a page
    marker and a field code.
  - Text that `Paragraph.text` holds and the model leaves out, such as hidden text, is the gap
    between two runs' offsets. So a hit in a block maps back to a span, or is known not to be one.
  - These fields are not part of the block model's JSON (REQ-706), which holds no tree objects.
- **The overloads, and one context.**
  - `Texts.of(part, …)` reads one story part, and `Texts.of(pkg, content, …)` reads a content object.
  - What a story needs from the rest of the document is a `TextContext`:
    - the note references' order, for `[n]`;
    - the numbering states per story, for labels and `value`;
    - the bookmarks, for `fields: update`.
  - `Texts.contextOf(pkg)` builds one. A call given a context uses it, and a call without one builds
    its own.
  - `TextResult.concat(…)` assembles several calls' results into one, in the order given.
  - Phase 2 ships all three forms.
  - The editor's session exports only the main story before its reads. Reading a header or a note
    through the part overload then reads the part as the session holds it, if the editor has
    exported it. Flushing its stories is the editor's own work.
- **Markdown per block.** `MarkdownExporter.blocks(result)` returns the front matter apart, and each
  block with its Markdown. `export` joins them. The editor writes an address comment before each
  block and caps the result, and its own writer goes (ED-008 P3).
- **Cost, measured.** Phase 2 measures `Texts.of` on the editor's 255-page benchmark,
  `ms-docx-250.docx` (375,340 bytes, in the editor's repository at
  `packages/editor-model/test/fixtures/large/`). A read at an address goes through the
  content-object overload and doesn't walk the whole document.
  - The file is Microsoft's [MS-DOCX] specification saved as a Word document, its body doubled.
    Whether it may ship in a public repository is an open question for Jason, which the editor's
    CLAUDE.md records.
  - Jason, 2026-10-21, by way of the editor's session: it may not be committed here. So phase 2
    reads it by path from the editor's checkout, a test over it skips where the file is absent, and
    nothing of its content goes into a golden or a snapshot here.
- **D6 and check 48.** The result goes to the editor. If `Paragraph.text` gains U+0002, every span
  in a paragraph with a note reference moves by one character per reference, so the editor changes
  its own `paragraphText` in the same install.

## 5. What this package claims, and what stays pending (D7)

The specification's conformance classes (section 12.1) are claimed separately. This CR claims:
- the plain-text extractor (sections 2 to 6, 7.1, 7.5 and 8);
- the Markdown extractor (with 7.2);
- the block-model extractor (with 7.4).

It doesn't claim the inspector (section 9), which docx4j hasn't built either.

Declared as not implemented when phase 3 lands, each pending in the suite until its phase:

| Statement | What it needs | Phase |
|---|---|---|
| REQ-506 `fields: update` | field-code parsing, the property and bookmark resolvers, the `\@` and `\#` switches (docx4j `FieldUpdater`, `FormattingSwitchHelper`) | 5a |
| REQ-604 math as LaTeX | a port of `OmmlToLatex`. Until then the walk writes each equation's `m:t` text inside the delimiters, REQ-604's form for an equation that can't be converted, and the statement stays pending | 5b |
| REQ-511 charts | `ChartTable` and `ChartDescription`: this package's CR-009, docx4j CR-047's 33 goldens | 5c, after CR-009 |
| REQ-511 SmartArt | the data part's text | 5c |
| REQ-512 embedded | a docx through the model, a pptx's slides, an xlsx's used range through its number formats (a formatter port) | 5d |
| REQ-513 altChunk | text and docx/Flat OPC chunks. XHTML, MHT and HTML chunks need an importer. The adapter advertises "xhtml importer absent", as the goldens were cut | 5e |
| REQ-203 Strict | a conversion on load, which CR-001 section 17.6 left out | not in this CR |
| `symbols: unicode`, `math: unicodemath` | refused, as docx4j refuses them | not in this CR |

## 6. Decisions (taken 2026-10-21, all as recommended)

1. **D1, the names.** Recommended:
   - docx4j's class names (`Texts`, `TextOptions`, `TextResult`, `PlainText`, `MarkdownExporter`),
     in `src/model/text/`;
   - options keyed by the specification's section 10 names verbatim (`note_references`,
     `field_codes`), not camelCase.

   The verbatim keys let the suite's scenarios, the editor's ED-008 and the MCP tools pass options
   through unchanged. docx4j-python takes the same names (its CR-011 Q1). This is a recorded
   departure from CLAUDE.md's camelCase rule.
2. **D2, the Office JS readers stay Office JS's.** This follows Jason's decision of 2026-10-19
   ("we may continue to call objects textOf as it is (meaning it wont be deprecated)").
   - Every reader in section 4.2's first row keeps what Office JS returns, held to Word checks, and
     is documented as not the text model.
   - **That conflicts with REQ-202.** It says every shipped text reader goes through the same four
     layers, and only a getter that is deprecated, documented and unused may stay outside.
   - Recommended: ask the specification for an exemption for a reader whose contract is another API
     an implementation emulates (Office JS's `paragraph.text`), documented as that API's text.
     Claim REQ-202 for every reader that has no such contract.
   - The alternative, the model behind `Paragraph.text`, would make `text` disagree with what Word's
     own Office JS returns for the same document (check 42's U+0002, a text box left out).
   - docx4j-python's CR-011 Q4 goes the other way: its `Body` and `Paragraph` text become the model,
     a departure from Office JS. The two ports should agree, and that is Jason's call for both.
   - The editor accepts D2 on one condition: the model's runs map back to `Paragraph.text`
     (section 4.4). The exemption's contract then says which text each field counts in.
3. **D3, `outline()` on the model.** Recommended. It is this package's own reader, not Office JS's,
   and the paper's section 10 names it with `getText` as what goes structural.
   - Its paragraph text becomes the model's: hidden text out, a deleted mark joined.
   - It lists text boxes after their anchor, and takes `TextOptions`.
   - Its shape stays: `paragraphs`, `tables`, `headers`, `footers`.
   - The editor's outline verb reads it, so this moves the editor's answers. That is what ED-008
     asks for.
4. **D4, the text-box address is `body/3/textbox:0/1`.** Recommended: keep the proposal unchanged.
   - docx4j implements it (`608e40023`), and python and the editor wait on it.
   - REQ-519's NOTE 2 gives the final spelling to this CR. Deciding it lets the specification mark
     the segment decided.
   - With it, `pkg.paragraphAt` resolves the `footnote:`, `endnote:` and `comment:` prefixes its
     bodies already use.
5. **D5, text-box views are read-only here.** Recommended.
   - A box exists once in each branch of its `mc:AlternateContent`: `wps:txbx` in the Choice and
     `v:textbox` in the Fallback.
   - CR-010's D2 has mutators edit every branch. That needs a mapping between the two boxes'
     paragraphs, which nothing has today.
   - So writes into a text box (`insertText`, `replaceText` inside one, bindings, paraIds) are a
     later CR. This CR's views throw `NotSupportedError` from every mutator.
6. **D6, U+0002 in the Office JS readers.** Check 42 measured that Office JS's `paragraph.text`
   reads each footnote and endnote reference as U+0002, where this package gives nothing.
   - Recommended: `Paragraph.text`, and the readers built on it, give U+0002 for each reference. In
     the model a reference is `[n]` (REQ-503).
   - The cost:
     - `Paragraph.text` moves off `textOf` to this package's own run reader, the case named when
       `textOf` was kept;
     - `Range` offsets and `search` count the character.
   - So a Word check comes first, check 48: whether Office JS's `search` matches across a reference,
     and what a `Range` over one reads.
7. **D7, the order and the first claim.** Recommended:
   - phases 0 to 3 make the three classes claimable, with section 5's statements declared and
     pending;
   - phase 5 fills them one at a time, each with Jason's go.
   - The alternative, claiming nothing until every statement passes, leaves the editor's ED-008
     waiting on a number-format port it doesn't need.
8. **D8, the oracle.** Recommended, all three of:
   - **(a) The conformance suite.** Copied into `test/fixtures/text-extraction/` at a recorded
     docx4j commit, with a provenance README (as `test/golden/`), and run by a Node mirror of
     `ExtractionConformance`.
   - **(b) The agreement test** (REQ-406 and REQ-1202) over the 128 CR-043 fixtures and the probes.
     These are copied too: 4.3 MB.
   - **(c) Block-model parity goldens.** A `TextHarness` in `test/java` records what docx4j's
     `Texts.of` answers per fixture: blocks, addresses, text per view, and a few option sets.
     They are compared as `parity.test.mjs` compares and held to zero differences.
     - These are docx4j's answers, as the paper's dependency rule requires ("the ports port the Java
       goldens").
     - They reach what the suite doesn't: REQ-519's addresses are pinned only by docx4j's own
       `TextsTest`.
     - (c) waits for docx4j to push `VERSION_17_3_4`, or release 17.3.4, so the weekly workflow can
       rebuild them. That needs CR-011 phase 0's move of the parity workflow and the pom.
     - They are cut at docx4j `8d32bf928` or later. Before it, every header or footer block's address
       began with the part content's Java `toString()` (`org.docx4j.wml.Ftr@50bb1c1ffooter:rId9/0`,
       a local variable in `Extractor.headerOrFooter` shadowing the prefix). It now reads
       `footer:rId9/0`, which `TextsTest.headersBySection` asserts and which this package's
       `HeaderPart` and `FooterPart` bodies already give. Section 2 was read at `3dcbb3764`, before
       the fix; nothing here was taken from those addresses (docx4j-25, by way of the portfolio,
       2026-10-19).
9. **D9, REQ-510 in the Emulator, first.** Recommended: port `Emulator.levelStyleId`.
   - A numbering level belongs to one paragraph style: the one its `w:lvl/w:pStyle` names, else the
     first paragraph style in `styles.xml` whose own `w:numPr` names it.
   - Any other style carrying the level gives no label and no count, but keeps the indent.
   - Word's labels from `LevelOwnerTest`'s three documents are the test.
   - It changes `Body.listLabels`, `ListItem.listString` and the parity goldens wherever two
     paragraph styles carry one unowned level. docx4j-25 says Word-saved documents never have that.
   - It is independent of the rest, and python's parity waits on it (its CR-011 risk 3). So it is
     phase 1. It is separable as its own CR if Jason prefers.

### 6.1 What the decisions settle, and what they leave (2026-10-21)

Jason took all nine on 2026-10-21: "CR-012 all decisions as recommended". Each stands as written
above. Read against what has moved since the CR was written:

- **D2 decides this package's readers, not python's.** docx4j-python's CR-011 Q4 makes its `Body`
  and `Paragraph` text the model. Until that is decided the same way, which is Jason's call there,
  the two ports' `paragraph.text` disagree for the same document (section 8's last risk).
  - The exemption is still to be asked. REQ-202, as WD 0.26 reads it, admits only a deprecated
    getter. It goes to the specification through the portfolio session (section 9).
  - **Asked and granted in WD 0.27** (portfolio `804007d`, 2026-10-21, Jason: "draft 27: good"). REQ-202
    now ends: "A reader whose contract is another API's that the implementation emulates (Office JS's
    `paragraph.text`) MAY likewise remain outside the layers if it is documented as that API's text
    and not the text model, and each of the text model's runs records where its text lies in that
    API's text, so that the two map onto each other." `TextRun.officeOffset` (section 4.4) meets the
    last condition, so it is phase 2's to ship with the blocks.
- **D4.** REQ-519's second NOTE in WD 0.26 still calls the text-box segment a proposal. The
  specification is asked to mark it decided.
  - **Decided in WD 0.27.** REQ-518 says "the form is `body/3/textbox:0/1`, REQ-519", and REQ-519's NOTE
    records the segment as adopted by this CR's D4 (Jason, 2026-10-21), as docx4j implements it.
- **D6 waits on check 48.** Check 48 isn't written yet. It is a Word check:
  - whether Office JS's `search` matches across a note reference, and what a `Range` over one reads;
  - whether every `w:br` reads as U+000B (section 12.3).

  It comes before phase 4 moves `Paragraph.text`. Until it is run, `Paragraph.text` is unchanged.
  - **Written 2026-10-21 and run the same day** (`test/README.md` check 48,
    `test/fixtures/check48/build.mjs`; Jason: "Write Word check 48"); section 6.2 has the answers. It measures every kind of inline content the run reader will map,
    not only the two above: note and comment references, the four breaks, tabs, hyphens, a symbol,
    hidden text, both revision kinds, both field forms, a hyperlink, a content control, a picture, an
    equation and a table. For each it reads `paragraph.text`, the paragraph's range, `body.text`,
    `search` across the item (with Word's `^` Find codes) and a range from one side to the other. This
    package's answers on the same file are recorded beside it.
- **D8 (c) no longer waits.** docx4j's `VERSION_17_3_4` is public: on GitHub at `5c691fc1f`
  (committed 2026-10-20) as of this writing. It holds `8d32bf928`'s address fix and `d9acdf5e2`'s
  level owner, and the parity workflow checks it out since CR-011 phase 0. So the block-model
  goldens can be cut from a branch commit with phase 2 rather than after it.
  - `test/java/pom.xml` resolves the 17.3.3 release for a hand-run, and that has no
    `org.docx4j.model.text`. So until 17.3.4 is released, a `TextHarness` runs against docx4j built
    from the commit it records, in a separate Maven repository.
- **D9 moves none of the 62 parity goldens.** docx4j at `5c691fc1f` already reads the unlinked
  level's owner in its labels (`Emulator.levelStyleId`, through `styleLinkedElsewhere`), and gave
  the same answers on all 214 goldens (CR-011 section 10.1). So it moves no label in those fixtures,
  and `LevelOwnerTest`'s three documents and Word's labels are the test, as D9 says.
- **The specification has moved on.** This CR read WD 0.16. The portfolio's copy is WD 0.26
  (2026-10-21). 0.17 to 0.26 are mostly:
  - the inspector, not in scope;
  - `altChunk` (REQ-513, phase 5e): an RTF chunk through an RTF reader, and `display: none` in an
    HTML, XHTML or MHT chunk as hidden text;
  - REQ-504 leaving out a comment with no anchor in the main document, a header or footer, or a
    note (0.20), which is phase 2's.

  Phase 0 starts by reading those changes, and the suite at the docx4j commit it copies. It did
  (section 13.1), and by then the specification was WD 0.27.
- **Still open, and not one of the nine:** whether the editor's `ms-docx-250.docx` may be committed
  here (section 4.4). Until Jason says so, phase 2 reads it from the editor's checkout. **Answered
  2026-10-21: no** (Jason, by way of the editor's session; section 4.4).

### 6.2 Check 48's answers, and what they ask of phase 4 (2026-10-21)

Jason ran check 48 in Word 365 (16.0.20326.20232, PC) on 2026-10-21. His JSON is
`test/fixtures/check48/48-text-word365.json`, and `test/README.md` check 48 has the whole table.
`paragraph.text` and a range over the same content agree in every case. Office JS's text is Word's own
character stream, and it differs from `Paragraph.text` here (through `textOf`) in most of the 18 cases:

| Content | Office JS | Here |
|---|---|---|
| Footnote or endnote reference | U+0002 | nothing |
| `w:br`, page break, column break, `w:cr` | U+000B, U+000C, U+000E, U+000D | `\n` each |
| `w:ptab` | `0` | nothing |
| Non-breaking, optional hyphen | U+001E, U+001F | U+2011, U+00AD |
| `w:sym`, any font or code (check 48b) | `(` | the code point (U+F04A) |
| Hidden run, its own `w:vanish` or its style's (check 48b) | left out | kept |
| Tracked deletion | in, in All Markup, check 48's view (out in Word's default, check 48b) | out |
| Equation `x=1` | ` 𝑥=1 `, Word's math italic | nothing |
| Between paragraphs (`body.text`), after a cell | `\r`; `\t` | `\n`; `\n` |

The comment reference, the tab, the insertion, both field forms, a hyperlink, a content control and a
picture agree: nothing for the reference and the picture, the text for the rest.

Office JS's `search` follows its own rules over that stream:
- a plain query matches across a comment reference and an optional hyphen, and `-` matches a
  non-breaking hyphen;
- it doesn't match across a note reference, a break, a tab, a symbol or a picture;
- Word's Find codes do (`^f`, `^e`, `^l`, `^m`, `^t`, `^~`, `^-`, `^g`);
- the picture is a position in a search (`^g`) but nothing in the text;
- it finds a hidden word, its hit's text `""`, and a deleted word.

So D6 was too narrow. U+0002 is one of eight characters, and hidden text goes the other way from this
package's. Check 48b (below) found that deleted text, and what `search` finds, depend on Word's view. Phase 4 needs Jason's answer to three questions first:

1. **How far `Paragraph.text` follows Office JS.** Recommended: all of it.
   - Every character as the table gives it: hidden text out, deleted text in, the equation as Word's
     math text. In `Body.text`, `\r` between paragraphs and `\t` after each cell.
   - D2's contract is Office JS's text, and an offset the editor counts in `Paragraph.text` is then
     the one Word counts.
   - A reader that wants clean text has the text model, from phase 2.
   - The alternative, Office JS's characters with this package's content rules (hidden kept, deleted
     out), is neither Office JS's text nor the model's.
2. **Which view.** Whether Office JS's text drops hidden text and keeps deleted text may depend on
   Word's view: whether hidden text is shown (File > Options > Display), and Review > Display for
   Review (Simple Markup, All Markup, No Markup, Original). The run didn't record it. This package has
   no view, so it needs one rule.
   - Recommended: a short follow-up, check 48b. The same file and snippet, run with hidden text shown
     and not, under each of the four markup displays.
   - Then take Word's default view: hidden text not shown, and Simple Markup, Word 365's default, as
     check 48b confirms.
3. **Search.** Recommended: phase 4's `search` takes Office JS's rules, Word's Find codes included,
   each matching its character in the table. A `^g` matches a picture's position, which takes no
   character in the text and so needs an offset map of its own.

**Jason's answers, 2026-10-21: questions 1 and 3 as recommended** ("q1 and q3: i agree"); question 2
below, after check 48b.
`Paragraph.text` follows Office JS in all of it, and phase 4's `search` takes Office JS's rules, Word's
Find codes included. Question 2, the view, waits on check 48b (`test/README.md`; written the same day,
with the editor's two refinements and a run hidden through its character style, five views in one run
of the snippet).

**Check 48b's answer to question 2** (run 2026-10-21; `test/README.md` check 48b). The five views were
measured on check 48's file. A second run in Word's defaults, on 48b's own file, took its cases 19 and
20: every `w:sym` reads `(`, whatever its font or code (the editor's question), and a run hidden through
its character style is left out as a run's own `w:vanish` is.
- **`paragraph.text` never holds hidden text**, shown or not.
- **It follows Review > Display for Review:**
  - Simple Markup, Word 365's default, and No Markup give the accepted text;
  - All Markup gives deleted and inserted text both;
  - Original gives the original text.
- **`search` matches what is displayed.** A hidden word is found only while hidden text is shown;
  otherwise a query matches across it. A deleted word is found in All Markup and Original.
- **`getReviewedText('Current')` and `getReviewedText('Original')` don't depend on the view.** They
  give Word's raw stream of the accepted or original text: hidden text kept, a field as U+0013 code
  U+0014 result U+0015, a hyperlink as its `HYPERLINK` field, a content control as `<...>`, a picture
  as `/`.

**Decided 2026-10-21 (Jason: "Word's default view: Simple Markup, with hidden text not shown"), as
recommended.** The rule is Word's default view, Simple Markup with hidden text not shown:
- `Paragraph.text` is the accepted text, leaving hidden text out, in check 48's characters.
- `search` matches over that: it neither finds a hidden word nor a deleted one, and matches across
  both.
- That is what an add-in reads in Word's default view. In content it is the accepted view that
  `getText({ view: 'accepted' })` reads, less hidden text; only the characters differ.
- `getReviewedText` is Office JS's own view-independent reader. It could be added to the content API
  later, as Word's stream with its field and object characters, but phase 4 doesn't need it.

The editor accepts the rule (2026-10-21, its ED-008 section 8.3, editor `d10f116`). Its find bar and its
agent's spans read the accepted view already, so only check 48's characters and hidden text move for it
at phase 4's install. Whether its find bar matches deleted text in All Markup, as Word's Find does, is
its own choice, stated as a departure from `body.search` if made.

**The editor's answers** (2026-10-21, its ED-008 section 8.3, editor `1505b8e`; for Jason):
1. Yes, all of it, as D2's contract. Its `atomText` moves with the install that carries it: the
   characters above, a `w:cr` told apart from a `w:br` at import (today both are one break there), the
   equation read through this package. Hidden text and deletions are not a character table; it pins
   them as known differences until it models `w:vanish` (effective, styles included) and question 3.
2. Check 48b, as recommended, and two refinements to it:
   - A second Wingdings symbol and a `Symbol`-font one. F04A reads as `(`, which is not its low byte
     (`J`), so one sample doesn't say whether Word maps each symbol or writes one constant character.
   - `getReviewedText('Current')` and `getReviewedText('Original')` per paragraph (Office JS, WordApi
     1.4). With them one run can say whether `text` is the displayed text or everything, without a
     run per markup display.
3. Yes. Its find bar holds to the engine's search, and finding a deleted word fits what it intends for
   ED-008 step 2: the find bar matches what its markup view shows, and the agent's find keeps counting
   in the accepted text (ED-008 section 3.1).

It also notes that a `w:cr` reads as U+000D, the character `body.text` puts between paragraphs, so in
`body.text` a `w:cr` can't be told from a paragraph's end.

**`officeOffset`** (section 4.4) counts in this stream: a `TextRun`'s offset in `Paragraph.text` steps
over each U+0002 and break character, and over hidden text where Office JS leaves it out, as the
portfolio session noted (REQ-202's emulated-API exemption asks only that the two map onto each other).

**What it costs.** In any paragraph holding one of these (a break, a note reference, a hyphen, a
symbol, hidden text, a tracked deletion, an equation), the editor's spans move. So its `paragraphText`
changes in the same install (section 4.4, the editor's fifth ask). Phase 4 grows by about a day: the
character table, the view rule, the search rules and the offset map. `TrackedChange.text`, `Table.values`
and the other readers of section 4.2's first row aren't measured here; they follow `Paragraph.text`
unless a check says otherwise.

## 7. Phases and sizes

| Phase | What | Size |
|---|---|---|
| 0 | The suite and the oracle (D8 a and b). Copy the fixtures, TSVs and goldens with provenance; the Node conformance runner with an engine that implements nothing (everything pending); the agreement test's runner; the understood-set assertion and the measurement of the 11 (7 since CR-010 section 15). **Done 2026-10-21, section 13** | 1 day |
| 1 | REQ-510's level ownership in the Emulator (D9), with `LevelOwnerTest`'s documents and Word's labels. **Done 2026-10-21, section 14** | ½ to 1 day |
| 2 | The plain-text model (sections 4.1 and 4.4): `Texts.of`, its part and content overloads with `TextContext`, and `ofResolved`; the source references; the three views and the agreement test passing; the stories, notes, comments, text boxes and addresses (D4); fields cached, codes and the TOC; hidden text; placeholders and bindings; images; properties; tables; the characters; CriticMarkup; the block model; the warnings, the read-time `mce.branches_differ` among them; `Texts.of` measured on `ms-docx-250.docx`. **Done 2026-10-21, section 15** (`paragraphAt`'s `footnote:`, `endnote:` and `comment:` with phase 4's views) | 7 to 9 days |
| 3 | Markdown (REQ-703, REQ-704): headings, emphasis, lists by the numbering model with REQ-703's interrupted-list rules, notes, GFM tables, front matter; `MarkdownExporter.blocks`. **Done 2026-10-21, section 16** | 2 to 3 days |
| 4 | The readers (D3, D5, D6): `outline()` on the model; text-box `Paragraph` views at their addresses; U+0002 after check 48; ED-009's run content in a kept branch, read and edited, check 49 run 2026-10-20 (section 12). **Done 2026-10-21, section 17** | 3½ to 4½ days |
| 5 | Section 5's statements, each its own go: a `fields: update`, b math as LaTeX, c charts (after CR-009) and SmartArt, d embedded, e altChunk | each 1 to 4 days |
| 6 | Block-model parity goldens (D8 c). docx4j's 17.3.4 work is now public, so this can go with phase 2 (section 6.1) | 1 day |

The estimates are rough. Phase 2 is the bulk, and it is where the agreement test will find the
deleted-mark and rejected-field cases the paper's REQ-406 NOTE records ("both ports' text views of
17 October fail it").

## 8. Risks

- **Labels move** (D9). Wherever a document has two paragraph styles on one unowned level,
  `listString` changes for every reader, Office JS's included. The three Word documents are the
  evidence that the new answer is Word's.
- **The suite is docx4j's.** Its scenarios encode docx4j's option handling, and its
  `textRequirements` makes exceptions for named cases. The Node runner mirrors them as written, and
  any difference goes back to docx4j rather than into the runner.
- **Embedded xlsx text needs Excel's number formats** (REQ-512). Nothing here formats numbers. A port
  of a formatter is the largest single item in phase 5, and it is why D7 doesn't wait for it.
- **Size.** The walk reads every story on each call. `Texts.of` is a whole-document read, and the
  editor calls it per verb. Phase 2 measures it on the editor's 255-page benchmark, and a read at an
  address takes the content overload (section 4.4). If the whole read is still too slow, a
  per-package cache invalidated by edits is a later measure, not part of this CR.
- **Two ports, one rule** (D2). If python's Q4 and this CR's D2 are decided differently, the two
  ports' `paragraph.text` disagree for the same document.

## 9. Not in scope, and hand-offs

- **The inspector** (section 9, REQ-901 to 904). docx4j plans it on its anonymiser's walker. It
  comes after that.
- **HTML export** (section 7.3). This package has no HTML exporter.
- **Strict conversion** (REQ-203). CR-001 section 17.6 left it out.
- **Writes into text boxes** (D5).
- **Text boxes in `search` and `replaceText`.** That needs a Word check of what Office JS's search
  does in a box, and D5's mapping for the replacement.
- **The editor** (ED-008) consumes the options, the blocks, the Markdown and the warnings. Its
  review is section 11. Its step 2 installs the release that carries phases 0 to 3.
- **docx4j-python** (CR-011) is to agree with D2 and D4.
- **The specification** is to settle the REQ-202 exemption (D2) and mark REQ-519's text-box segment
  decided (D4), through the portfolio session.

## 10. References

- docx4j CR-046, `docs/developer/change-requests/CR-046-text-extraction-behaviour-model.md`,
  `VERSION_17_3_4`, done 2026-10-18. Commits:
  - phase 0 `c0c72ecef` and `65a527d47`;
  - phase 1 `f71cdd299` and `8da15bb8c`;
  - phase 2 `608e40023` and `187e5a367` (with CR-047);
  - phase 3 `f59828e62`, with `d9acdf5e2` after it (the level owner and the goldens as the WD cites
    them);
  - suite changes `a41de133b` (several options in a row) and `3dcbb3764` (the third fixture, the
    capability rules).

  All unpushed when read.
- The package `org.docx4j.model.text`. The harness is
  `docx4j-core-tests/src/test/java/org/docx4j/textextraction/`: `ExtractionConformance`,
  `TextsAdapter`, `ExtractionConformanceTest`. Also
  `org/docx4j/model/text/ViewAgreementTest.java` and
  `org/docx4j/model/listnumbering/LevelOwnerTest.java`.
- The suite: `docx4j-core-tests/src/test/resources/text-extraction/` (`GOLDENS.md`, `phase0/`).
- The specification: portfolio `docs/text_extraction_specification_wd.md`, WD 0.16 (`dae90bf`,
  2026-10-19). The paper: `docs/text_extraction_behaviour_model.md` (section 10, section 10a
  item 4).
- docx4j-python CR-011, `docs/change-requests/CR-011-one-text-model.md`.
- This package:
  - CR-002 section 3.3 (addresses);
  - CR-010 sections 12 (D7, D8), 12.6 (check 42) and 14.7;
  - CR-001 section 17.6 (Strict);
  - CR-009 (the chart read model);
  - CR-011 (tracked changes in every story, which shares the story order);
  - `test/README.md` check 42.

## 11. The editor's review (2026-10-19)

The editor (`docx4j-ts-editor-50`) read the CR against what ED-008's `read`, `find` and `outline`
need. It accepted D3, and D1, D4, D5, D7 and D8 as recommended. It made five asks, all taken into
section 4.4:

1. **Two texts per paragraph under D2.** The editor's spans (`set_format`, `delete`, `comment`), its
   `if_text` and its change report count in `Paragraph.text`. Under D2 a block's text differs
   wherever hidden text, a note reference, a joined mark or a text box is involved. So a hit taken
   from the blocks would edit the wrong characters.
   - Asked: each block and run carries its source, and each run its offset in Office JS's text where
     it contributes. Taken as `source` and `officeOffset`.
   - Without that the editor would keep `find` on `Paragraph.text` and use the model for snippets
     only, two readers for one verb.
   - The REQ-202 exemption is fine with the editor on this condition.
2. **The call site isn't the package.** The editor's reading tools run inside `session.readTree`,
   which exports only the main story before the callback, so `Texts.of(pkg)` would read stale notes
   and headers and miss the session's comments.
   - Asked: docx4j's overloads over one part and over a content object, and a way to assemble
     several calls into one result with note numbers and list values right.
   - Taken as the overloads with `TextContext` and `TextResult.concat`.
3. **Markdown per block.** The editor puts an address comment before each block and caps the
   result, and one joined string doesn't allow that. Taken as `MarkdownExporter.blocks`.
4. **Cost per call.** The editor calls per verb, after each of the agent's edits, on documents up to
   its 255-page benchmark.
   - Asked: phase 2 measures on it, and a read at an address doesn't walk the whole document. Taken.
   - The editor hands the fixture over.
5. **D6 and check 48.** U+0002 in `Paragraph.text` moves every span in a paragraph with a reference
   by one character per reference. Asked: check 48's result, including whether Office JS's search
   matches across the character, before the change. Taken.

Smaller points, recorded:
- D3 doesn't move the editor's own outline tool by itself. That tool reads `Paragraph.text`, and
  ED-008 moves it onto the blocks. The in-app outline panel is unchanged.
- Under D4 and D5 the editor's `Address` and `addressOf` follow at the install. Its XPath find then
  stops at a box's paragraph rather than the anchor.
- D7 suits ED-008. Its step 2 installs the release carrying phases 0 to 3. Phase 4 isn't on its
  path, and phase 5 fills in behind.
- The warnings' block-grammar addresses are what `read` needs.
- The editor's pending insertions in a room are `w:ins` by the agent and read as accepted text. No
  special case is needed.

## 12. Run content in a kept branch (the editor's ED-009, 2026-10-20)

The editor's ED-009 (accepted by Jason on 2026-10-20, editor main `0a2ec4d`) covers Jason's probes,
`mc-simple.docx`, `mc-run.docx` and a paragraph with ten Choices. They put bare `w:t` in the branches
of a run-level `mc:AlternateContent`, and runs in those of a paragraph-level one. Word 365 shows the
first Choice it understands. The editor now projects the selected branch's items as the run's own,
editable. Its question 1 (ED-009 section 6, proposal P4) is for this CR. Jason, 2026-10-20: "do as the
editor asks re Q1, Q2". Question 2 is CR-010 section 15.

### 12.1 What this package reads and edits today

Measured on `f8617e8` over the editor's `mc-branches.docx`. It is built from `mc-simple.docx`,
case by case.

| Case | Shape | `Paragraph.text`, kept | resolved (`mcePreprocess: true`) |
|---|---|---|---|
| 1 | `w:r/mc:AlternateContent`, a `w:t` in each branch | `""` | `"The Earth is flat"` |
| 2 | the same, with a tab and a break in the Choice | `""` | `"Before\tafter the tab\nafter the break"` |
| 6 | ten Choices after a run `I prefer ` | `"I prefer "` | `"I prefer w14"` |
| 8 | `w:p/mc:AlternateContent`, runs in each branch | `"bold ying plain ying"` | `"bold ying plain ying"` |

- **Run-level** (cases 1, 2 and 6). The element is one object segment of no text (CR-010 section
  13.1), so `search('flat')` finds nothing in case 1. That is the difference the editor pins in its
  `alternate-content.test.mts`, naming ED-009.
- **Paragraph-level** (case 8). `Paragraph.text` and `search` go through the objects package's
  `textOf`, which reads the branch. The mutators go through `segmentsOf`, which doesn't list the
  element ("Word writes none, and the corpus has none", CR-010 section 13.1). So a `Range` that
  `search` returns inside the branch lies past every segment, and on a kept load:
  - `insertText('very ', 'Start')` writes a new run after the element, at the paragraph's end;
  - `insertText('PLAIN', 'Replace')` writes `PLAIN` there and leaves ` plain ying` in place, so
    the text reads `bold ying plain yingPLAIN`;
  - `delete()` and `font.italic = true` change nothing.

  None of them throws. This is a defect in 0.5.0, whose load keeps the element by default.
  Resolved, the element is gone and all four edit as they should.
- **Where the shape occurs.** Not in a Word-written file in the corpus. Over its 490 `.docx` files
  (this repository's fixtures and docx4j-core-tests' resources at `VERSION_17_3_4`), 13
  paragraph-level elements hold runs in their branches. All 13 are in docx4j's hand-built
  text-extraction probes (`extraction-probe.docx` and its four copies under `phase0/`), which phase 0
  copies here. CR-010 section 13.1's "the corpus has none" was measured before those probes. Jason's
  probe was built by hand too.
- **Through the editor's agent.** The companion's `find` returns spans from `Paragraph.search`, and
  its `edit` deletes or replaces through the engine's `Range`. So an agent that finds text in such a
  branch and edits it gets a success for an edit that went wrong or did nothing (the editor,
  2026-10-20; its `docs/upstream-limitations.md` item 23).

### 12.2 The rule (the editor's, as asked)

1. **A run-level element** whose selected branch (`selectedBranchOf`) holds run items reads as
   those items. Its `w:t` are text segments at the run's offsets, its tabs and breaks their
   characters, and its objects objects.
2. **A paragraph-level element** reads as the runs its selected branch holds. That goes for
   `segmentsOf` too, so the text and the mutators agree.
3. **An edit** through the content API over such a paragraph lands in the selected branch. The
   other branches are kept as they were read.
4. **A deletion** that takes the whole element takes every branch with it, as now (CR-010 section
   13.1).
5. **CR-010's one-object rule stays** for the shape Word writes, a text box. A selected branch
   holding a `w:drawing` or a `w:pict` reads as one object, and D5's text-box views are unchanged.

### 12.3 What follows

- **Item 3 departs from CR-010's D2** (mutators edit every branch, decided 2026-10-17) for this
  shape. Nothing maps a Fallback's text to the Choice's, as D5 says of text boxes. So after an edit
  the branches disagree. That is what the editor's ED-009 writes, and the next load's
  `mce.branches_differ` reports it. Taken with the rule; Jason's to reverse.
  - **Office JS edits the same way.** Check 49: an `insertText` through a range over a branch's text
    lands where it says, and Word's save holds the edit.
  - **Word does the same, and more.** Jason's Word 365 save of the editor's check file (ED-009
    section 9.2, 2026-10-20) has no `mc:AlternateContent` left. Every element, edited or not,
    run-level or paragraph-level, became its selected branch's items as plain runs, and every Fallback
    was dropped. So a stale branch matters only to a reader that doesn't understand the Choice, until
    Word next saves the file. docx4j CR-021 section 8.1's "Word keeps a paragraph-level element under
    `w:p`" holds for a bare drawing, the case it measured; with runs in the Choice, Word resolves it.
- **D2, the Office JS readers.** Items 1 and 2 change `Paragraph.text`, `search`, `Range` offsets
  and `replaceText` for these shapes. Under D2 those keep Office JS's text, and what Office JS's
  `paragraph.text` reads for a branch's text isn't measured. Word displays the selected Choice;
  that's the probe. So **check 49** came first. It read `paragraph.text`, `search` and a range over a
  branch's text in Word 365 (`test/README.md`). If Office JS read the branch, the rule would be Office
  JS's and D2 would hold; if not, the question would go back to Jason.
  - **Run 2026-10-20 (Jason, Script Lab, over the editor's `mc-branches.docx`): Office JS reads the
    branch.** `paragraph.text` gives the selected branch's words in all eight cases, run-level and
    paragraph-level, and never a Fallback's. `search` finds a word only a Choice holds, and never one
    only a Fallback holds. So items 1 and 2 are Office JS's text, and D2 holds.
  - **Found with it.** Office JS gives a `w:br` in a branch as U+000B, where `textOf` gives `\n`.
    Check 48 can measure whether that holds for every break, before `Paragraph.text` moves off
    `textOf` (D6).
- **`textOf`.** `Paragraph.text` calls it (D2), and it reads a paragraph-level branch but not a
  run-level one. Its branch choice is its own prefix set (section 1), and its output is kept stable
  (objects `c31cb83`). So item 1 is the same move as D6: `Paragraph.text` onto this package's run
  reader, which then takes `selectedBranchOf`'s branch. Phase 4 does both together.
- **The model** takes the selected branch wherever the element is (section 4.1, item 1). So the
  blocks read these shapes whatever the Office JS readers do, and the agreement test (D8 b) covers
  the probes.
- **The defect in 12.1.** Item 2 closes it. If it shouldn't wait for phase 4, there are two smaller
  fixes. One is `segmentsOf` listing a paragraph-level element's selected runs, items 2 and 3 for
  that shape alone (about half a day). The other is a mutator throwing `NotSupportedError` for a
  range its segments can't place. The editor would take the error first: its companion maps an
  engine error to a refusal with a reason, where today a silent no-op reports success. Jason's call.
- **Size.** Phase 4 grows by 1½ to 2½ days: the run reader under item 1 (shared with D6), items
  2 and 3, and the tests over the editor's eight cases.

### 12.4 Until then: an edit the runs can't place throws (implemented 2026-10-20)

*Superseded 2026-10-21 by phase 4 (section 17.3): the segments read the branch, so every offset is a place
in the runs, and these edits land in the selected branch.*

Jason, 2026-10-20, asked whether to fix the defect now: "yes to all 3 questions". Of section 12.3's
two smaller fixes, this is the error, the one the editor asked for first. Making the branch editable
now would take items 2 to 4 with it: a deletion emptying the branch, tracked changes inside it, the
run listings. Those are phase 4.

- **The rule.** `Paragraph.splitAt`, `splice` and `insertItemsAt` compare `Paragraph.text` with the
  segments' text. Where the two differ, they refuse an offset past the first difference, and an
  insertion at it, with `NotSupportedError` (`Word.ErrorCodes.notImplemented`), naming CR-012
  section 12. Every text edit goes through one of the three: `insertText`, `delete`, `replaceText`,
  `Range.font` writes, `insertOoxml`, content controls, hyperlinks and comments. Each refuses before
  it changes anything. A tracked `insertInlinePictureFromBase64(..., 'Replace')` checks before it adds
  its image part, which it otherwise does before its `splice`. A break or a picture at the paragraph's
  `Start` or `End` goes to the ends of its content, so it is placed right and not refused.
- **What still works.** An edit before the element. `Paragraph.text = value` with change tracking
  off, which replaces the whole content, every branch with it, as item 4 says. Every edit in a
  resolved load.
- **Where it applies.** Over the corpus's 14,236 body paragraphs (486 of its 490 `.docx` files load),
  `Paragraph.text` and the segments differ in 9. All 9 are in docx4j's text-extraction probes, and
  all are this shape. So no edit in a Word-written file is newly refused.
- **One consequence.** `Body.replaceText` works paragraph by paragraph, so a refusal in a later
  paragraph leaves the earlier paragraphs' replacements made.
- **Measured.** On the editor's case 8, kept: replace, insert at a range's start, delete, italic and
  `insertText('!', 'End')` each throw. Resolved, each edits as it says. `acceptAllRevisions` and
  `rejectAllRevisions` give the same trees and counts as before on the 488 files of CR-002 section 53.
- **Tests.** `kept-branches.test.mjs`: seven edits past the element each throw and leave the part
  unchanged, and so do a tracked replace and a tracked picture Replace, which adds no part. Edits
  before it work and keep every branch, the same seven edit the resolved paragraph exactly, and an
  edit beside a run-level element works. With the picture's check after `addImage`, the test fails
  ("no part added"). Against `8d90308` it fails ("Missing
  expected exception: replace"). The suite passes (878) and so does typecheck.

## 13. Phase 0: the suite and the oracle (implemented 2026-10-21)

Jason, 2026-10-21: "do CR-012 phases 0 and 1".

### 13.1 The specification since WD 0.16

Read from WD 0.17 to WD 0.27 (portfolio `804007d`, committed the same day, carrying this CR's two asks).
For this package:
- **0.17 to 0.22, the inspector** (section 9, REQ-901 to 907, REQ-1204). Not in scope (section 5).
  Its rows in the suite pend as "the inspector (CR-050)".
- **0.20, REQ-504.** A comment with no anchor in the main document, a header or footer, or a note is
  left out, as Word drops it on save. Phase 2's.
- **0.23 to 0.27, `altChunk` (REQ-513).** An RTF chunk is read through an RTF reader where there is
  one, and left out with `altchunk.not_expanded` where there isn't. An HTML, XHTML or MHT chunk is
  read through an HTML reader (a parser of [HTML]'s algorithm and a reader of its tree), and is not
  expanded where there is none. What `display: none` hides in one is hidden text. A chunk of more
  than 50,000 elements is not expanded. The suite's capability is now `html reader present` or
  `absent`, in place of 0.16's XHTML importer.
  - This package has neither reader. So its adapter advertises `html reader absent`, the
    capability the goldens are cut under, and case 36's and case 63's HTML rows, which need a
    reader, stay pending.
  - All of it is phase 5e's. Nothing in 0.17 to 0.27 changes phases 2 to 4 beyond REQ-504's rule.
- **0.27, this CR's asks.** REQ-202's exemption and REQ-518/519's text-box segment, section 6.1.

### 13.2 The copy

`test/text-extraction/` (its `README.md` is the provenance):
- `suite/` is docx4j's `docx4j-core-tests/src/test/resources/text-extraction/`, all 80 files,
  unedited.
- `revisions/` holds the 9 of docx4j's 128 CR-043 fixtures that `test/fixtures/` didn't already
  have. The other 119 were there byte for byte, from this package's own checks and CR-011.
- `understood-namespaces.txt` is docx4j's published list.
- `manifest.json` gives every file's docx4j path, SHA-1 and path here.

That is 1.7 MB in all, where section 6 estimated 4.3 MB. `scripts/copy-text-extraction-suite.mjs
<docx4j> <commit>` makes the copy from the commit with `git archive`, never from a working tree.

**The commit is `198d015a3`** (`VERSION_17_3_4`, docx4j CR-053 under WD 0.28, Word's rounds 11 and 12),
committed 2026-10-21 and not yet pushed. Its suite is WD 0.28's (portfolio `f2046e5`, whose `[PROBE]`
cites `4a8833eea`) with only `support/inspect_oracle.py` changed: the oracle's decoder drops its
UTF-16 step, and no table, fixture, golden or document moves. The next WD revision cites it.
- The commit before it, `d18693760` (docx4j CR-043 section 15: a comment whose range a resolution
  empties goes whole, wherever its markers were, as Word 365 does), adds four fixtures under
  `revisions/`: `python/artefact8/` (three) and `editor/review/review-for-word-word365-rejected.docx`.
  They are copied too, which makes 13 of docx4j's 132 copied here and 119 found under
  `test/fixtures/`.
- The first copy, committed as `77c5e1e`, was at `1d108c0af` (CR-053 phase 2), the newest committed
  suite WD 0.27 cited. That suite differed from the pushed `5c691fc1f`, which is WD 0.26's
  `86449c013` unchanged, in three ways: case 36's and case 63's HTML rows need an HTML reader;
  their warning rows follow; and three goldens are recut without the chunk's text.
- `4a8833eea` adds a fifth fixture, `extraction-probe-html.docx`: four chunks for an inspector with
  an HTML reader, its findings in `expected-inspection-html.tsv` and its status in
  `expected-inspection-html-status.tsv`. The other four fixtures' inspection status now skips their
  HTML and MHT chunks as `not_decoded`, the oracle script (`inspect_oracle.py`) does the same, and
  the token rows' notes cite WD 0.27. No token row, golden or case statement moved.
- Jason, 2026-10-21: "wait for the WD 0.28 commit", then "WD 0.28 should be committed - ask the
  portfolio agent". The portfolio session's answer: copy at `4a8833eea`, since WD 0.28 cites it and
  every table and fixture is final there, and again at Word's round 11 commit, which would change only
  the oracle's decoder. Both were done the same day. At `198d015a3` the manifest showed exactly that
  one file changed in the suite, as the portfolio and docx4j-25 said.

**One departure from D8 (a)'s wording.** The copy is at `test/text-extraction/`, not
`test/fixtures/text-extraction/`. The weekly parity workflow's `RevisionsHarness` writes a golden
for every `.docx` with revision markup at any depth under `test/fixtures/`, and the probes hold
such markup. Under `test/fixtures/` they would have become CR-011 goldens, the 90 would have
moved, and differences there are CR-011's to investigate, not this phase's.

### 13.3 The runners

**`test/text-extraction/conformance.mjs`** is docx4j's `ExtractionConformance` as it stands at
`4a8833eea`, mirrored as written:
- the tables and their column check;
- REQ-1201's normalisation and counting, with the longer-marker rule;
- the markings;
- the requirements, with docx4j's named exceptions (case 21, case 25, `CHT34_AXIS_TITLE`), the
  scenario statements and the HTML reader's capability;
- the phases;
- the supplemental scenarios;
- the findings, order, inspection, counts and golden checks;
- since `4a8833eea`, the HTML reader's inspection of the fifth fixture, which compares `parsed` as well,
  leaves out a row that waits for a ruling (its `pending`) and reports it pending for that reason;
  and the other fixtures' whole inspection, which needs `html reader absent` as well.

An engine is the adapter `TextsAdapter` is in docx4j. Here it may be asynchronous, since `Texts.of`
will be. There is none until phase 2.

With no engine the suite has **1,584 rows, all pending**. By the phases of docx4j's CR-046, which
the report names: 618 are its phase 1 (the plain-text model), 697 its phase 2, 54 its phase 3
(Markdown) and 215 the inspector's. Here the first two are phase 2's, with phase 5's statements
among them, and the third is phase 3's.

The port was checked against docx4j's own run (its
`docx4j-core-tests/target/text-extraction/conformance.tsv`, written at 08:00 on 2026-10-21 after
`4a8833eea`: 1,584 rows, 1,565 passing and 19 pending; `198d015a3` moves no table, so no row):
- The 1,584 rows here are exactly docx4j's rows, in the same order.
- Each of docx4j's 19 pending rows pends here for the same phase, the fifth fixture's inspection
  among them (docx4j's run has no HTML reader).
- The first copy was checked the same way against docx4j's run of that morning: its 1,583 rows were
  docx4j's less the fifth fixture's 7.

The runner's own tests (`text-extraction.test.mjs`) port `ExtractionConformanceTest`'s checks of the
protocol: normalisation, marking, the golden reader, Markdown image destinations, an engine's rows
and a golden checked and a broken one failing, the order rows, the inspection compared by
occurrence, and the case statements against Appendix A and the README. One test is this package's
own: the HTML reader's inspection, whose `parsed` and `pending` docx4j's tests reach only through
its engine.

Not ported: the three tests of docx4j's builder output (`corpusHasEveryCaseAndValidPayloads`,
`inspectionOracleMatchesTheFixtures`, `docxAltChunkRequiresItsImportContentType`). They check the
fixtures docx4j builds. Here the manifest's hashes hold the copy to docx4j's files.

One Java detail is spelled out. Normalisation strips Java's `\s`, which is the six ASCII white-space
characters, and `\p{Z}`. JavaScript's `\s` adds U+FEFF, the only character it holds outside those,
so the runner names the characters. A test pins U+FEFF as kept.

**`test/text-extraction/agreement.mjs`** is docx4j's `ViewAgreementTest`. It covers 147 fixtures:
docx4j's 132 under `revisions/` and the 15 `.docx` under `text-extraction/`, in docx4j's order.
Each gets four comparisons (accepted and original, in the default scope and with every story
included), so 588 in all, every one pending.
- The copy route, the reference, runs now: `acceptAllRevisions` and `rejectAllRevisions` on each
  of the 147 complete without an exception, in about 4.5 seconds.
- docx4j's `d18693760` changes what `rejectAll` leaves of a comment on rejected inserted text, where
  this package still follows Word 2010 and 15 (CR-002 section 38). That changes the reference's
  comments, which the agreement test reads only with every story included. It is not this CR's to
  port: it belongs with CR-011's resolve-all verbs, and it waits on Jason. *Ported 2026-10-21 as
  CR-002 section 54, with `a8bc6cda9`'s range across paragraphs. The views (`ResolvedView`) make the
  same comments candidates; that half comes with this CR's views.*
- As docx4j does, the copy is the fixture loaded a second time, not `clone()`. docx4j's reason, a
  save dropping an unconventional `Requires` prefix (probe case 6), no longer holds here since
  objects 0.4.3 (CR-010 section 14.9). But a second load is the cheaper copy, so section 4.1's
  `Texts.ofResolved` may take a loaded package rather than clone one; phase 2 decides.

### 13.4 The understood set (REQ-306)

A test asserts this package's sets against docx4j's list:
- `UNDERSTOOD_NAMESPACES` is the list's first 107 URIs, in order;
- `READ_ONLY_UNDERSTOOD_NAMESPACES` is lines 109 to 112;
- the 7 others are named: line 108, `drawing/2016/11/diagram`, and lines 113 to 118, six SpreadsheetML
  revision namespaces.

**The measurement.** The corpus was:
- this repository's `test/`;
- the docx4j checkout, every module;
- the editor's checkout;
- the Enterprise chartex Office files, read in place;
- `~/vbShares`.

It held 14,531 Office files, of which 10,838 distinct packages read (217 didn't), with 189,750 XML
parts and 11,415 `mc:Choice` elements. **No `Requires` names a prefix bound to any of the 7.** They
are declared often, as `mc:Ignorable` namespaces on part roots: the diagram namespace in 30,215
parts, each revision namespace in 2,990 to 5,128. None is ever a Choice's requirement. So the
readers' set stays at 111, and the 7 are the documented difference. The scan is
`requires-scan.mjs` in the session's scratchpad. It was not committed, being a one-off over
machine-local corpora.

### 13.5 Tests

`text-extraction.test.mjs` has 15 tests:
- the copy against the manifest, with no file added;
- the case statements;
- eight protocol tests;
- the suite with no engine;
- the agreement fixtures, the copy route and the runner's report of a difference;
- the understood set.

The whole suite passes (999, with section 14's seven), and so does typecheck.

### 13.6 What phase 0 leaves

- **Copy again** whenever WD's `[PROBE]` moves. Recount the pending rows and port any runner
  change each time, as the README says.
- **D8 (c)**, the block-model goldens, goes with phase 2 or 6 (section 6.1).
- **Phase 2** installs the engine in `text-extraction.test.mjs` (`ENGINE`) and the agreement test's.

### 13.7 docx4j CR-053's hand-off, for phase 5e (2026-10-21)

docx4j-25 handed off docx4j CR-053 (phase 4) on 2026-10-21. docx4j is the oracle under WD 0.28, at
`35e304770`, `1d108c0af`, `4a8833eea` and `198d015a3`, which the suite copy carries. Nothing changes
here now:
- this package has no extractor;
- it writes no altChunk: its `html` module converts XHTML into WordprocessingML;
- `AlternativeFormatInputPart` is typed on load by its relationship, keeping the part's declared
  content type.

For phase 5e, and for any future writer of HTML chunks:
- **Writing** (docx4j `AltChunkHtml`):
  - strip a leading XML declaration, after a BOM and white space too, since Word 365 refuses one;
  - keep the declared encoding: UTF-8 with a BOM, or another encoding transcoded to UTF-8 with a BOM
    and its `meta` rewritten;
  - type a loaded chunk by its declared content type, the extension being the fallback (CR-053 D6).
- **Reading** (REQ-513):
  - An HTML, XHTML or MHT chunk is read only through an HTML reader: an [HTML]-algorithm parser and
    a reader of its tree. XHTML is read as HTML, well-formed or not. Without a reader the chunk is not
    expanded, with `altchunk.not_expanded` naming its content type.
  - Decoding (WD 0.28): a BOM; else a `meta` `charset` or `http-equiv`, by [ENCODING]'s labels
    (ISO-8859-1 and US-ASCII are windows-1252); else UTF-8 where valid; else windows-1252. There is no
    XML-declaration step and no UTF-16 detection without a BOM.
  - The hiding rule is WD 0.28 REQ-513's subset of CSS, with docx4j's `HtmlChunks` as the reference:
    - a selector list loses its whole rule on any selector Word ignores;
    - a rule inside any at-rule does not count;
    - a class matches only a host style of exactly its name;
    - elements Word doesn't know are transparent;
    - an explicit `display` other than `none` or `inherit` shows its content;
    - a table's row groups are read in source order.
  - A chunk over 50,000 parsed elements is not expanded.
- **Inspecting** (REQ-901, 902, 907; section 5 does not claim the inspector):
  - Without a reader, a chunk is skipped as `not_decoded`, its scripts and references still found.
  - With one, the inspector runs the extraction's import after REQ-907's limits. It gives one
    `text.hidden` per outermost hidden element holding text that isn't all white space, at its path
    in the parsed tree, and `parsed: html` on the `altchunk` finding.
  - Over 50,000 elements the inspector reports `limit.elements`.
- **The harness:** `htmlInspections` and `pend`, ported in section 13.3.

## 14. Phase 1: REQ-510, the level's owner (implemented 2026-10-21)

docx4j `d9acdf5e2` (CR-046 phase 3's follow-ups), ported:
- **`Emulator.levelStyleId(numId, ilvl)`**, and the static `Emulator.levelStyleId(numberingPart,
  numId, ilvl)` as docx4j has it. A level's owner is the paragraph style its `w:lvl/w:pStyle`
  names. Failing that, it is the first paragraph style in styles.xml whose own `w:numPr` names a
  `w:num` on the same `w:abstractNum` at that level (its `w:ilvl`, else 0). Otherwise there is none.
- **`styleLinkedElsewhere`** asks it in place of `getLinkedStyleId`. A paragraph naming no style is
  the default paragraph style's in the check.
- **`PropertyResolver.getStylesInOrder()`** is new: the styles part's `w:style` elements in document
  order, from the tree the resolver reads (its private copy, or the part's own once that is
  unmarshalled). docx4j reads the styles part's JAXB element there. The resolver held the order
  only in its id map until now.
- **`getNumber` with no `w:pPr`** resolves as an empty `w:pPr` does, against the default paragraph
  style (docx4j's change in the same commit). `peek` follows.
- **`NumberingResult.count`** was already docx4j's new `getValue()`. Its comment says so now.

**One docx4j inconsistency, found by the port and fixed in docx4j.** At `d9acdf5e2`, docx4j's
`numRefFor(pkg, null)` still answered "no pPr", not numbered. Its own comment calls it "the answer
getNumber acts on", and `getNumber` now numbered that paragraph. So in a document whose default
paragraph style is numbered, a paragraph with no `w:pPr` would have had a label from
`Body.listLabels` and `ListItem.listString`, while `Paragraph.isListItem` (through `numRefFor`) said
false.
- The parity harness records `numRefFor` per paragraph, so the first port (`77c5e1e`) kept docx4j's
  answer and pinned it in a test.
- Reported to docx4j the same day. Jason: "make the one-line fix so it matches getNumber" (docx4j
  `c283dffdb`, pinned by its `ParityAccessorsTest`: numId 1, ilvl 0, from the style).
- Ported. `numRefFor` resolves a missing `w:pPr` as an empty one, against the default paragraph style.
  As in docx4j, a package with no numbering part now answers "no numbering part" before anything
  else. Two tests changed with it: the pinned one now asserts numId 1, ilvl 0, not direct; and an
  older one that held "no pPr" now holds "no numbering part" and the default style's reason.
- The 62 parity goldens give 0 differences with it, so none of their fixtures has the case, and none
  is regenerated.

**Tests** (`numbering.test.mjs`) port the Emulator half of docx4j's `LevelOwnerTest`, with its three
documents in `test/fixtures/listnumbering/`:
- the three documents' labels, which are Word's, read off its PDFs;
- the same labels through `Body.listLabels`;
- the owners;
- the style that lost the level keeps its indent (720 and 720) and has no label;
- in `DefaultStyleNumberedTest`'s document, a paragraph with no `w:pPr` is numbered, and `numRefFor`
  agrees.

Against the code before the change all seven fail, the linked document's because its list C names
no style. With it, all seven pass.

Where the fixtures live: the parity `Harness` reads only `test/fixtures/` and its `parity/`, and the
three hold no revision markup, so neither weekly harness reads them.

**Parity.** The 62 goldens give 0 differences, as section 6.1 expected: their fixtures have no
style that loses a level. The text model's half of `LevelOwnerTest` (`theTextModelNumbersAsWordDoes`,
`theTextModelsCountIsTheLabels`) is phase 2's, with `TextNumbering`.

**Who sees it.** `Body.listLabels`, `ListItem.listString` and every other caller of
`Emulator.getNumber`, in a document where two paragraph styles carry one level that names no style.
The second style's paragraphs lose their label and their count, as in Word.

## 15. Phase 2: the plain-text model (implemented 2026-10-21)

Jason, 2026-10-21: "start CR-012 phase 2". The same day, on the result: "decisions accepted". So REQ-303 and
REQ-1203 stay unclaimed as section 15.3 says, and the departures of sections 15.5 and 15.6 stand. REQ-303 by
namespace URI is not asked of the objects package. REQ-307's wording went to the specification through the
portfolio session, and phase 6 follows separately. On REQ-307 Jason then chose docx4j's rule (section 15.6,
item 2).

### 15.1 What was built

- **`src/model/text/`**, exported from `./model` and `.`, docx4j's `org.docx4j.model.text` at `198d015a3` (it hasn't
  changed since; `VERSION_17_3_4` is at `4b9c71288`):
  - `Texts`: `of` for a package, for one story part, for content in a package and for content with no package;
    `ofResolved`, the copy route; `plainText`; `contextOf`; `normalize`.
  - `TextContext` and `PackageTrees` (section 15.4).
  - `TextOptions`, keyed by the specification's names (D1), with `TextOptionError` naming the statement of a
    refusal, and `TEXT_OPTION_NAMES`.
  - `TextResult` (with `concat`), `TextBlock`, `TextRun`, `TextRow`, `TextCell`, `TextField`, `TextNumbering`,
    `TextRevision`, `TextWarning`, `TextComment`, `TextImage`, `ImageDescription`, `ImageDescriber` and
    `AltTextDescriber`. docx4j's enum values are kept as strings (`'PARAGRAPH'`, `'TEXT_BOX'`,
    `'NOTE_REFERENCE'`), as CR-011 kept its change kinds, so a block model passed as JSON reads the same from
    either implementation.
  - The walk, `Extractor`, and the serialiser, `PlainText`: the listed, TSV and grid table forms,
    CriticMarkup with `markup_detail`, the brackets of REQ-707, the metadata block and `layout: indent`.
- **`src/model/content/ResolvedView.mts`**: docx4j's `org.docx4j.model.revisions.ResolvedView` (at `a8bc6cda9`, with
  U23's comment ranges across paragraphs), over this package's own listing (`trackedChangesOfStory`) and
  resolution. The moves are replayed on `TrackedChange.mts`'s `movePiecesOf`, which it now exports with
  `storyOf` and `MovePiece`. This is the half of docx4j `a8bc6cda9` that CR-002 section 54 left to the views: a
  comment a resolution would take whole goes from the `accepted` and `original` views too.

### 15.2 The package is never changed (REQ-204)

- Every tree a reading walks is read before the walk starts (`PackageTrees.read`): the main document, every
  note, comment, header and footer, the settings, the document properties, and the glossary when it is
  included. Each comes from `readContents()`: the live tree where its part is unmarshalled, a private parse
  where it isn't. The resolver and the numbering definitions were already read privately.
- The plan is made over those same trees, by identity, so content given to `Texts.of(pkg, content)` must be
  the package's own tree: its part unmarshalled. A private parse is a new tree each time.
- A binding's store is the custom XML part's DOM (`parseDocument`, which doesn't adopt it), or the core or
  application properties' XML parsed apart from the part.
- Tested: after a reading in every view with every option phase 2 takes, the three probes save byte for
  byte as before, and no part is unmarshalled. A package with no styles part isn't given one.

### 15.3 What the suite and the agreement test say

**The agreement test (REQ-406, REQ-1202): 588 comparisons, no difference.** Over its 147 fixtures, the views'
walk equals the copy that `acceptAllRevisions` or `rejectAllRevisions` leaves, read with no view logic, in
both views, in the default scope and with everything included. 94 of the 147 read differently in the two
views and 85 have a different number of lines, so the comparison is not empty.

**The suite: 735 rows pass and 849 pend.** The engine is `test/text-extraction/adapter.mjs` (docx4j's
`TextsAdapter`). It claims REQ-201 to 202, 204, 301 to 302, 304 to 307, 401 to 406, 501 to 505, 507 to 510, 514
to 518, 520, 601 to 603, 605 to 606, 701 to 702, 706 to 709, 801 to 802, 1001 and 1202, with "html reader
absent". What pends, by what it waits for:

| Waits for | Rows |
|---|---|
| The inspector: REQ-901 to 907 (186), with REQ-1204 (5), REQ-903 alone (24) | 215 |
| REQ-512, embedded (phase 5d) | 165 |
| REQ-506, `fields: update` (5a) | 153 |
| REQ-513, altChunk (5e) | 81 |
| REQ-512 and REQ-513 | 61 |
| REQ-511, charts and SmartArt (5c) | 54 |
| REQ-703, Markdown (phase 3) | 51 |
| REQ-513 with an HTML reader present | 15 |
| REQ-604, LaTeX (5b) | 12 |
| REQ-506 and REQ-512 | 12 |
| REQ-303, `Requires` by namespace URI (case 6) | 9 |
| REQ-1203, the whole-document goldens | 9 |
| REQ-511 and CR-047 (a chart's description) | 6 |
| REQ-512, REQ-513 and an HTML reader | 3 |
| REQ-703 and REQ-1203 | 3 |

- **REQ-1203.** The goldens hold the whole document, phase 5's content with it, so the statement is claimed
  when phase 5 is done. Until then a test of this package's own holds them as they can be held. The
  external and structure probes' six goldens are exact in every view, with the trailing line ends trimmed to
  one, as docx4j's `ExtractionGoldens` writes them. The main probe's three differ only on case 6's line and
  phase 5's lines: the equation not as LaTeX (case 32), the chart (34), the SmartArt (35) and the chunks (43
  to 47).
- **REQ-303.** Case 6 binds the `wps` namespace to the prefix `shape2010` and expects the URI to decide. This
  package matches `Requires` by prefix (`selectedBranchOf`, CR-010), which REQ-303 allows where the object
  model keeps no declarations, if documented. So case 6 reads the Fallback, and its rows pend. The objects
  package now keeps each `mc:Choice`'s source-only bindings (its CR-011, objects 0.4.3, for the marshal), but
  privately. Matching by URI would need a reader for them there, and `selectedBranchOf` taking it, so that the
  content API reads the same branch. That would be an objects CR, and is Jason's call; nothing else in the
  corpus rebinds a prefix (section 13.4).

### 15.4 The context and the overloads (section 4.4)

- **`TextContext`** is what a story needs from the rest of the document: the note references' order (`[n]`),
  the numbering states per story, the bookmarks (for phase 5a), the page markers so far, and each view's plan,
  made once. It is made by `Texts.contextOf(pkg)`, which reads every story once (`PackageTrees`). A call
  given one uses it and adds to it, and a call without one makes its own.
  - A context is one reading in one view. Its note numbers and list counts are that view's, so a call in
    another view is refused (`TextOptionError`), and so is a call over another package.
  - It is a snapshot: after the package changes, make another.
- **The part overload** reads the main document part (its story and text boxes; the notes it references are
  numbered in the context, not read), a header or a footer (addressed `header:rId3/...`), the footnotes or
  endnotes part, or the comments part. docx4j's takes the first three; the notes and comments parts are this
  package's, as the editor asked.
  - A notes part read under a context that hasn't read the main story reads it first, for the references'
    order alone, so a notes part read on its own numbers its notes as the package does.
  - Read in REQ-501's order under one context, the parts give the package's text: a test holds the
    concatenation equal to `Texts.of(pkg)` in all three views.
- **The content overload** reads in the package's view, addressed from `content`. It finds the part whose
  tree holds the content, for a hyperlink's relationship and a picture's image part, where docx4j passes
  none. With a context whose plan is made, one paragraph of the 255-page benchmark reads in under a
  millisecond.
- **`TextResult.concat`** joins results in the order given: the blocks one after another, the comments
  together, each warning once, the first's metadata and options.

### 15.5 Source references and `officeOffset` (section 4.4, REQ-202)

- A block's `source` is its paragraph's `w:p` (a joined block's first), a table's `w:tbl`, or a bound control's
  `w:sdt`. A run's `source` is its `w:r` and `w:p`.
- A run's `officeOffset` is where its text starts in its paragraph's `Paragraph.text`. That is the objects
  package's `textOf` today (D2), so the offset is computed by `textOf`'s own walk: its containers, its branch
  choice (the conventional prefixes) and its characters. Phase 4 moves `Paragraph.text` to check 48's
  characters, and this computation moves with it.
- **Within a run the two texts agree character for character.** A run that would join another across text
  `Paragraph.text` holds and the model leaves out stays apart. Examples are hidden text, a table of contents'
  result under `toc: false`, or a `w:ptab` (a tab here, nothing in `textOf`). docx4j joins such runs. The
  text is the same either way: only the block model's run boundaries differ, and only there.
- What only the model writes has no offset: a note reference or mark, a list label, a page marker, a field
  code, an image, a bound value. Text `Paragraph.text` holds and the model leaves out, such as hidden text,
  is the gap between two runs' offsets.
- These are not in the block model's JSON (`toJSON`), which holds no tree object (REQ-706).

### 15.6 Departures from docx4j, and what phase 2 leaves out

1. **`TextModelView`**, not `TextView` (D1's exception). The content API has exported a `TextView` since CR-002
   phase F: `getText`'s two views, `accepted` and `original`.
2. **`mce.branches_differ` reads each branch as docx4j does**: every `w:t`, `m:t` and `a:t` in it, whatever the
   view, hidden text included. The first implementation read each branch by the walk's rules, as section 4.1
   said and the paper's section 4.1 item 6 has it: revisions by the view, hidden runs left out unless
   `hidden: true`.
   - Over the 284 `.docx` files of `test/` (this repository's fixtures and the suite), in all three views,
     the two rules gave the same warnings, and no suite row moved either way.
   - REQ-307 didn't say which. The portfolio session put it to Jason for WD 0.29 and recommended docx4j's
     rule. A comparison REQ-307 lets an implementation make at load can't depend on a view. And a branch
     that differs only in hidden text is one an agent should hear of.
   - **Jason, 2026-10-21: "choose docx4j's rule."** `branchText` is docx4j's `normalized` again. A test holds
     that branches differing only in hidden text, or only in inserted text, warn in every view, with
     `hidden` on or off. Deleted text is `w:delText` and is never compared.
   - **WD 0.29** (portfolio `b76886a`, 2026-10-21) writes the ruling into REQ-307: a branch's text is "every
     `w:t`, `m:t` and `a:t` it holds, in document order, in every view and whether or not it is hidden", so
     the warning depends on neither the view nor `hidden`. Its `[PROBE]` cites docx4j `198d015a3`, the suite
     copied here, so no copy is needed.
3. **Run boundaries** follow `officeOffset` (section 15.5).
4. **A move half its story has no piece for** (one in a text box, whose paragraphs' story is the body
   around it) is resolved as the plain insertion or deletion it stands for, as `acceptAllRevisions` resolves
   it. docx4j's plan leaves such a half where it is.
5. **Not yet** (section 5's statements, pending in the suite):
   - `fields: update` (REQ-506) and `embedded: true` (REQ-512) are refused with a `TextOptionError`, reason
     `not_implemented`. docx4j's `math: unicodemath` and `symbols: unicode` are refused as docx4j refuses
     them.
   - An equation is its `m:t` text in document order inside `$...$` or `$$...$$`, REQ-604's form for one
     that can't be converted, until phase 5b ports a converter.
   - A chart or SmartArt contributes nothing, and raises nothing, until phase 5c. Legacy WordArt is read: a
     one-paragraph text box.
   - An `altChunk` is never expanded until phase 5e. It raises `altchunk.not_expanded` with its content type,
     under `altchunks: true` as under `false`.
6. **Strict documents** aren't read (REQ-203, CR-001 section 17.6). The four in docx4j's test resources fail to
   unmarshal, as they do for every reader here.
7. **Dates** are written as the package wrote them, from the calendar's fields (docx4j's `toXMLFormat`).

### 15.7 Cost (section 4.4's measurement)

Measured on the editor's `ms-docx-250.docx`, read from its checkout and not copied here (section 4.4); Node 22,
the median of five:

| Reading | Time |
|---|---|
| `Texts.of(pkg)`, a fresh context, parts not unmarshalled | 1.44 s (accepted), 1.54 s (original), 1.34 s (markup) |
| of which `Texts.contextOf`, the private parse of the main part above all | 1.23 s |
| `Texts.of(pkg)` with the main part unmarshalled, as the editor holds it | 314 ms (accepted), 306 ms (original), 203 ms (markup) |
| `Texts.of(pkg)` under a context made once | 198 ms |
| `Texts.of(pkg, paragraph)` under a context with its plan made | under 1 ms |
| `Texts.of(pkg, paragraph)` with no context, the main part unmarshalled | 112 ms (the plan) |
| `PlainText.of` | 3 ms |

The document reads as 6,684 blocks and 498,215 characters, and gives 304 warnings, all `field.no_result`
(index entries, which have none). Unmarshalling the main part (`getBody()`) takes 1.40 s, so a reading costs
the editor about a fifth of what the document's own load does. Section 8's per-package cache isn't needed.

### 15.8 Tests

- **`text-model.test.mjs`** (42 tests):
  - docx4j's `TextsTest`, less `strictAsTransitional`;
  - its `TextsTocTest`;
  - its `TextsPhase2Test` where the statement is phase 2's: the table forms, the listed format's headers,
    metadata, the glossary, images and describers, WordArt, bound values, page markers, `layout: indent`, the
    package unchanged, and an equation as its text;
  - this package's own: the byte-for-byte save; `officeOffset` against `Paragraph.text` over the probe
    (every run with an offset, in two views); hidden text as the gap between offsets; the part overloads
    under one context concatenated equal to the package's text, the view a context is bound to; the
    content overload; the JSON; REQ-307's rule as Jason ruled it (section 15.6, item 2).
- **`text-extraction.test.mjs`**: the suite with the engine (the counts above, and every pending row naming a
  statement the engine doesn't claim); the whole-document goldens as section 15.3 holds them; the agreement
  test over 147 fixtures. The phase 0 test with no engine stays, and checks the runner's pending report.
- **`numbering.test.mjs`**: `LevelOwnerTest`'s text-model half, `theTextModelNumbersAsWordDoes` and
  `theTextModelsCountIsTheLabels`. The three documents' labels from the text model are Word's, and each
  `TextNumbering`'s value is its label's count.
- **A sweep**, not committed: the 579 `.docx`, `.docm` and `.dotx` files of this repository's `test/` and
  docx4j's test resources, in all three views with every option phase 2 takes, through `PlainText.of` and
  `JSON.stringify`. In 15.6 s, 42,664 blocks, no exception but the Strict four.

### 15.9 What phase 2 leaves

- **Phase 3**, Markdown (`MarkdownExporter`, `blocks`): the 54 rows waiting on REQ-703. *Done 2026-10-21,
  section 16.*
- **Phase 4**, the readers:
  - `outline()` on the model;
  - text-box `Paragraph` views at `textbox:` addresses, and `pkg.paragraphAt` resolving `footnote:`,
    `endnote:` and `comment:` (D4);
  - `Paragraph.text` in check 48's characters, `officeOffset` with it;
  - ED-009. *Done 2026-10-21, section 17.*
- **Phase 5**, a to e, each on Jason's go.
- **Phase 6** (D8 c): the block-model parity goldens. They will meet sections 15.5 and 15.6's departures: run
  boundaries where `Paragraph.text` holds more, and the content overload's part.
- **REQ-303 by URI**, if Jason wants it (section 15.3): a reader in the objects package for each `mc:Choice`'s
  source bindings, then `selectedBranchOf` over it.

## 16. Phase 3: Markdown (implemented 2026-10-21)

Jason, 2026-10-21: "start CR-012 phase 3".

### 16.1 What was built

All three modules are in `src/model/text/`, exported from `./model` and `.`:
- **`MarkdownExporter.mts`**: docx4j-markdown's `MarkdownExporter` and `MarkdownExportOptions`, at
  `VERSION_17_3_4`. Their main sources haven't changed since `d9acdf5e2`, before the suite's `198d015a3`.
  - `export(pkg)` reads the package under `options.textOptions`.
  - `export(result, pkg?)` writes a result under the options it was read with. The package, where given,
    is asked for the quote style's indent.
  - `blocks(...)` (section 16.4, item 5).
  - The static forms, `MarkdownExporter.export(...)` and `MarkdownExporter.blocks(...)`, use the default
    options, section 4.1's form.
  - `MarkdownExporter.frontMatter(result)` is REQ-704's YAML.
  - `MarkdownExportOptions` holds `textOptions` and `view`, its alias, and `imageDestination` (section
    16.4, item 2).
- **`WmlToMarkdown.mts`**: docx4j's `WmlToMarkdown`, which builds the Markdown tree from the block model:
  - headings by outline level, not style name;
  - lists by the numbering model, with REQ-703's interrupted-list rules: a margin paragraph ends the list,
    which resumes at Word's count; an indented paragraph continues the deepest item it reaches; an empty
    paragraph leaves the list alone; an adjacent list takes the other delimiter;
  - a label level as a bold label;
  - code blocks from the code styles or a mono paragraph style; inline code, bold, italic and strikethrough
    where a run adds them to its style;
  - links; pictures with the describer's text;
  - notes as `[^1]` and `[^e1]` with their definitions after the body;
  - GFM tables: the `w:tblHeader` rows as the header, alignment from the header cells, a span padded, a
    cell's paragraphs `<br>` apart;
  - block quotes from the Quote styles, nested by the indent beyond the style's own;
  - thematic breaks; display math;
  - CriticMarkup in `markup`, with `markup_detail`.
- **`commonmark.mts`**: a port of commonmark-java 0.30.0's Markdown renderer, which docx4j serializes
  with (section 16.2). It covers `MarkdownWriter`, `CoreMarkdownNodeRenderer` and the Markdown renderers of
  its GFM tables, strikethrough and footnotes extensions, plus docx4j-markdown's math renderer, for the
  nodes `WmlToMarkdown` builds. commonmark-java is BSD-2-Clause. Its notice is kept in the file, and
  `NOTICE` names it.

### 16.2 The renderer: a port, not a dependency

docx4j writes a commonmark-java tree and renders it with commonmark-java's `MarkdownRenderer`. So the
escaping, the list markers, the block separators and the line prefixes in docx4j's output are
commonmark-java's.
- **The alternative** was a JavaScript Markdown library. It would write its own escaping and markers, not
  commonmark-java's. The suite would mostly not see the difference, since REQ-1201's normalisation strips
  `*`, `_`, `\`, `{`, `}`, `$` and white space. But two ports' Markdown would differ for the same document.
  And `package.json` would gain a dependency where the runtime has `fflate` alone.
- **Ported:** the renderer's five classes docx4j uses are 1,032 lines of Java with their comments. The
  port leaves out the nodes docx4j never builds (indented code, HTML blocks, soft breaks, inline footnotes)
  and is 649 lines of TypeScript, the licence included. No dependency was added.
- **The result is docx4j's Markdown byte for byte** wherever the block models agree (section 16.5).

### 16.3 What the suite says

**786 rows pass and 798 pend.** The adapter writes Markdown for format `md`, with the package, as docx4j's
adapter does, and claims REQ-703 and REQ-704 besides phase 2's statements. Every one of the 51 rows that
waited on REQ-703 alone passes:
- case 21's six Markdown rows: a link's target, from the relationship and from the field code;
- case 56's 30 rows: the interrupted list's tokens;
- case 64's 15 rows: the links' texts.

What still pends is section 15.3's table less those 51 rows. REQ-1203 now holds 12: the three Markdown
goldens wait on it as the nine plain-text ones do.

**The goldens.** The test holds the Markdown goldens as it holds the plain-text ones (section 15.3):
- the external and structure probes' Markdown is exact;
- the main probe's differs only in case 6's block (REQ-303) and phase 5's: the equation (32), the chart
  (34), the SmartArt (35) and the chunks (43 to 47).

So 8 of the 12 goldens are exact.

### 16.4 Departures from docx4j

1. **Async.** `export` and `blocks` return promises, because a picture's bytes are read from its part
   (`getBytes()`) and `Texts.of` is async. `PlainText.of` stays synchronous.
2. **`imageDestination`, not `imageDirPath` and `imageTargetUri`.** Nothing outside `./node` writes a file
   (CLAUDE.md's rule). Where docx4j writes each image to a directory, a host here passes a function. It
   is given the bytes, the content type, the part, the picture's number and docx4j's file name
   (`image1.png`), and it returns the link. The default is docx4j's: a data URI.
3. **An SVG picture's data URI holds the part's bytes as stored.** docx4j's holds the SVG re-serialized
   from its `XmlPart`, with the attributes in another order. REQ-703 leaves destinations to the
   implementation, and the suite's runner compares them as `…`.
4. **One `export` for a package or a result.** docx4j's three overloads are `export(source, pkg?)`.
   docx4j's deprecated `TrackedChangesPolicy` isn't ported: the view is the option.
5. **`blocks(...)`** is this package's, the editor's ask (sections 4.4 and 11). It gives the front matter,
   each block with its Markdown, and the whole. Joined, the front matter and the blocks' Markdown are
   `export`'s.
   - A block's Markdown is whole lines, from the start of the line its first character is on to the
     next block's, with the blank line after it.
   - A block that writes nothing has `''`: an empty paragraph, or a code paragraph merged into the fence
     before it.
   - Where a block begins is recorded as the renderer writes. Each node made for a block carries the block
     as its `tag`, and the writer notes the line its first character lands on.
6. **Not here:** docx4j's Markdown importer (`MarkdownImporter`, `MarkdownToWmlVisitor` and `LatexToOmml`).
   It turns Markdown into WordprocessingML, which isn't text extraction. Math as LaTeX waits on phase 5b,
   and until then an equation is its `m:t` text between the delimiters (section 15.6, item 5).

### 16.5 Checked against docx4j itself (a sweep, not committed)

docx4j-core and docx4j-markdown were built from a read-only export of `198d015a3`, with `mvn package` and
nothing installed into the shared repository. A runner wrote docx4j's Markdown of each of the 284 `.docx`,
`.docm` and `.dotx` files of `test/` in five readings:
- `accepted`, `markup` and `original`;
- `accepted` with `metadata: true`;
- `markup` with `markup_detail: author_date` and `comments: true`.

That is 1,420 exports, each read with `export(result, pkg)` as the adapter reads it. This package wrote
the same in the same way.
- **1,290 are byte-identical.**
- **1,320 are identical once image destinations are masked.** Of the 497 pictures, 467 point to the same
  data URI. The other 30 are one SVG picture in each of the three files of check 38b and the three of check
  47, in five readings each (section 16.4, item 3).
- **The other 100 differ only in content that phase 5 or REQ-303 governs**, which the plain text leaves
  pending too:
  - LaTeX (5b): docx4j-python's artefact 8 in six copies, CR-043's three equation fixtures, and
    `loadAndSave.docx`, which also has a chart;
  - charts (5c): `charts/waterfall.docx`;
  - REQ-303: check 50's rebound prefixes;
  - the probe in its eight copies: case 6, and phase 5's cases.

The tests' own expectations were checked the same way: block quotes with and without the package, the
thematic break, emphasis, the escaping case, a mono document default and a code fence. docx4j writes the
same strings.

### 16.6 Cost

On the editor's `ms-docx-250.docx` (section 15.7; Node 22, the median of five, a result read once under a
context), the Markdown takes 62 ms with `export(result, pkg)` and 53 ms with `blocks(result, pkg)`, against
3 ms for `PlainText.of`. That is 693,135 characters of Markdown, and 4,109 of the 6,684 blocks write any.
Joined, the blocks are the export.

### 16.7 Tests

- **`markdown.test.mjs`** (43 tests):
  - docx4j's `MarkdownTextModelTest`, all 19 of its tests;
  - its `MarkdownExportTest`, all 11. `markdownSignificantTextIsEscaped` reads its Markdown back through
    docx4j's importer, so here it compares docx4j's string instead;
  - its `MarkdownExportExtensionsTest`, its 9 tests in 5. The image tests import Markdown, so here they
    build the picture;
  - its `MarkdownAlternateContentTest`: built, loaded kept, and loaded with `mcePreprocess: true`;
  - this package's own: block quotes with and without the package, the thematic break, emphasis, two of
    `blocks()`, REQ-204, and the options.
- **REQ-204** is tested on the three probes. After a Markdown export in every view, with metadata, every
  story, hidden text, list labels and field codes, no part is unmarshalled and each probe saves byte for
  byte as before.
- **`text-extraction.test.mjs`**: the counts above. The goldens test reads the Markdown goldens too.
- 1,097 tests pass.

### 16.8 What phase 3 leaves

- **The editor** can replace its own writer (ED-008 P3) with `blocks()` once a release carries phase 3.
- **Phase 6** could hold the Markdown too: section 16.5's runner as a harness, its outputs as goldens.
- Phases 4 and 5 are as section 15.9 has them.

## 17. Phase 4: the readers (implemented 2026-10-21)

Jason, 2026-10-21: "start CR-012 phase 4". It went as section 7's row and section 15.9 have it. One question
came up first: `outline()` going async. The editor was asked ("ask the editor for its opinion on outline()")
and Jason decided it (section 17.6).

### 17.1 One run reader: `Paragraph.text` is Office JS's (D2, D6, section 6.2)

`Paragraph.text` no longer calls the objects package's `textOf`. It is the text of its segments
(`segmentsOf`, `tree.mts`), the same reader every edit, `Range` and `search` use, so there is one text and
one set of offsets.
- **Office JS's characters** (`officeItemText`, check 48):

  | Item | Character |
  |---|---|
  | footnote or endnote reference | U+0002 |
  | the note's own mark (`w:footnoteRef`, `w:endnoteRef`) | U+0002 (check 51) |
  | `w:br` | U+000B |
  | page break | U+000C |
  | column break | U+000E |
  | `w:cr` | U+000D |
  | `w:ptab` | `0` |
  | non-breaking hyphen | U+001E |
  | optional hyphen | U+001F |
  | any `w:sym` | `(` |

  A comment reference, a field's code and a picture add nothing.
- **Word's default view** (section 6.2, Jason's question 2): the accepted text, with hidden text left out.
  Hidden means a run's own `w:vanish`, else its paragraph style's and character style's: the text model's
  rule (REQ-507), and Word's in check 48b case 20. It is decided by `PropertyResolver.isVanishedByStyles`,
  cached per style pair, and `hiddenRunsOf`.
- **An equation** (`m:oMath`, `m:oMathPara`) is one segment, its text as Word's (`mathTextOf`, `mathText.mts`;
  checks 48 case 17 and 51):
  - each letter is in Unicode's mathematical alphabet for its run's `m:sty` and `m:scr`, Greek capitals and
    nabla included, and each digit upright in its script's family;
  - a hyphen-minus is the minus sign U+2212;
  - `m:nor` text is as stored;
  - each structure is `,`, its arguments a `-` apart, then `.` (a fraction a over b `,𝑎-𝑏.`), its operator and
    delimiter characters not written, and a degree `m:degHide` hides left out.
- **`Body.text`** is `body.text`'s form in Word's default view (checks 30, 31 and 51):
  - each paragraph's text is followed by `\r`, the last paragraph's left out;
  - in a table a row's cells are a `\t` apart, a cell's last mark is not written, and each row ends with
    `\r\n`. Check 48's `\t` after every cell and row was All Markup's, with hidden text shown (section 17.11).

  `getText({ view: 'original' })` has the same form. A paragraph's `getRange('Whole').text` ends with its
  mark, `\r`, or `\t` for a cell's last paragraph. That mark is no offset: edits and `search` work on the
  text before it. `'Content'` ends with the text.
- **The other readers in section 4.2's first row** use `Paragraph.text` per paragraph:
  - `TableCell.text` and `value` (so `values`), a block `ContentControl.text` and `Comment.content` are their
    content's `Body.text`, the paragraphs a mark `\r` apart, as check 51 measured;
  - `Table.text`, this package's extension, keeps a line per paragraph;
  - a block control's `getRange()` is its first paragraph's, the mark at the end of 'Whole' (a `Range` is one
    paragraph's; Office JS's holds the whole control, check 51).
- **Agreement with Word.**
  - All 23 paragraphs of check 48's file equal Word 365's `paragraph.text` in check 48b's run A, Word's
    default view.
  - All 25 of 48b's file equal its take 2.
  - Each 'Whole' range and `Body.text` equal check 48's, but for case 11, whose deleted text check 48 read
    in All Markup.

### 17.2 `search` by Office JS's rules (section 6.2, question 3)

`officeMatches` (`search.mts`) is the one search behind `Paragraph.search`, `searchMatches`, `replaceText`
and `Range.search`. It reads `Paragraph.text` with these rules:
- a picture is U+FFFC, a place a plain query doesn't match across and `^g` matches;
- hidden text is not there;
- an optional hyphen is passed over unless the query names it (`^-`).

A plain query takes Word's Find codes: `^f`, `^e`, `^l`, `^m`, `^t`, `^~`, `^-`, `^g` and `^^`. `^f` and `^e`
each match only their own kind of reference, though both read U+0002. A plain `-` matches a non-breaking
hyphen too. Wildcards read the same text with the codes untranslated. A regular expression
(`matchRegExp`, this package's extension) reads `Paragraph.text` itself. A paragraph with nothing but text
is searched as it stands, a fast path.

Word's own queries on its files:
- 12 of 12 hits equal run A's and 12 of 12 equal 48b take 2's;
- check 48's 41 give 35, every Find code among them. The other six are the hidden and deleted words, which
  check 48 read with hidden text shown, in All Markup.

### 17.3 ED-009's run content in a kept branch (section 12.2)

The segments read a kept `mc:AlternateContent` as its selected branch:
- **In a run**, a branch holding run items reads as those items at the run's offsets. One holding a
  drawing, a picture or an object, or content the model doesn't bind (DOM), is one object, as before.
- **Among the runs**, the element reads as the runs its branch holds.
- **The segments carry `alternate`**, so edits land in the branch, and the other branches are kept as they
  were read:
  - a run split inside a run-level branch splits the element as the editor's ED-009 section 3.4 (P2) does:
    the first piece keeps the element and its other branches, and a later piece gets an element of its own,
    the same branches with the selected one holding the rest and the others empty;
  - a tracked deletion writes the selected branch's `w:delText` and keeps the Fallback's `w:t` (ED-009 P3,
    `toDeletedText`);
  - an edit that empties the branch takes the element, every branch with it (item 4), at run and paragraph
    level.
- **Section 12.4's `NotSupportedError` is gone.** Each of its seven edits now gives the text a resolved
  load gives, the Choice holding the edit and the Fallback as read. `kept-branches.test.mjs` asserts this.
- **`runsOf`** reads a paragraph-level branch's runs too.

### 17.4 Text boxes at `textbox:` addresses (D4, D5), and the notes' and comments' addresses

- **`textBoxesOf(p)`** (`textBoxes.mts`) numbers a paragraph's text boxes as the text model's `anchored`
  walk does:
  - the accepted view's drawings, VML pictures and kept elements, the selected branch only;
  - a hidden run's and a field code's left out;
  - nothing in a chart, SmartArt, a picture or an embedded object;
  - legacy WordArt counted, with no paragraphs.
- **The addresses resolve both ways.**
  - `Body.elementAt` and `paragraphAt` resolve `body/3/textbox:0/1`. `addressOf` gives it back, stepping
    from a box's paragraph to its anchor.
  - An address that names nothing is `undefined`, never an error, as the editor asked.
  - Over every `.docx` in `test/` holding a text box, 129 of the 132 text-box paragraphs the text model
    reports resolve to the same `w:p` and give back the same address. WordArt is another 8, with no
    paragraph.
  - The other 3 are in a block-level element (`mc-fun`'s `txbx` probes), which `elementAt` doesn't step
    into: CR-013 phase 2's addresses.
- **D5: a text box's paragraph is read-only.** It reads as any paragraph does (`text`, `search`, `font`).
  `Paragraph.inTextBox` says it is one. Every write through it, a `Range` on it or its `Font` throws
  `NotSupportedError` before changing anything (`assertWritable`). Text boxes are not among the body's
  paragraphs, as check 42 found Office JS's are not.
- **`pkg.paragraphAt`** resolves `footnote:2/0`, `endnote:1/0` and `comment:5/0`: the note's or comment's
  body by its `w:id`, as the text model and `Comment` address them.

### 17.5 What was found on the way (defects in 0.5.0, fixed)

1. **`Range.font` writes formatted more than the span.** A span ending at a `w:t`'s end in a run holding more
   (a tab, more text) formatted the rest of the run too: bold on `lat` in `flat⇥really` bolded `⇥really`. A
   write now isolates the span's runs, as a tracked deletion does (`splitRunsAround`).
2. **A `Font` read split runs.** Reading `range.font.bold` split the runs at the range's edges. `Font`'s
   holders now take a write flag, and a read splits nothing.
3. **`Paragraph.insertContentControl` linked the new control's PARENT into itself.** The builders' `sdt()`
   relinked the paragraph before its parent was read, so the control's PARENT was its own content. Every
   PARENT climb through such a control ran out its guard, so `addressOf` of its paragraph was `undefined`.
   An unguarded climb looped, which is how phase 4's `inTextBox` found it.

### 17.6 `outline()` on the model (D3), async

The text model's reading is async: building a `TextContext` reads parts privately. `Body.outline()` was
sync, and the editor called it synchronously four times.
- **Asked of the editor** (2026-10-21). Its answer (ED-008 section 8.7) was (a), both async: one definition
  of a paragraph's text, and no second behaviour hidden behind a context argument. Its calls are in
  `readTree` callbacks, which take async functions, and its panel moves to a `useEffect`.
- **Its conditions:**
  - a text-box flag on each entry;
  - `elementAt` over a `textbox:` address returns `undefined` rather than throwing;
  - the cost reported (section 17.8).
- **Jason took (a).**
- **What it is now.**
  - `Body.outline(options?, context?)` and `pkg.outline(options?, context?)` return Promises. They take
    `Texts.of`'s options and read under one context.
  - Each paragraph's text is the model's (`TextBlock.text`): hidden text out unless `hidden: true`, a
    deleted mark joined, the view asked for.
  - A text box's paragraphs follow their anchor, or the table it is in, flagged `textbox: true`.
  - The shape is otherwise the same.
- **No import cycle.** `Body` can't import the text model, which imports it. So the text model registers
  its reader (`model/text/outline.mts`), as the comment API registers itself, and the package class loads
  that module.

### 17.7 Departures and choices, recorded

- **`TrackedChange.text` keeps docx4j's characters.** The revisions goldens hold docx4j's text (`\n` for a
  break, nothing for a note reference), so a tracked change's text, a row's and a section's are read with
  `segmentsOf`'s `characters: 'textOf'`, the reader as it was before this phase. Section 6.2 had them
  follow `Paragraph.text` unless a check said otherwise. The goldens are that check.
- **An equation in an edit.**
  - A span may take an equation whole, untracked.
  - A span that cuts one, an insertion inside one, or any tracked edit taking one throws
    `NotSupportedError`: its math runs' own revisions are nothing this package writes.
  - A `Font` write over a span leaves the equation's runs as they are.
- **Hidden text in an edit** is an object to the span, as a picture is (CR-002 section 48): a deletion takes
  the hidden text strictly inside it, and an insertion beside it goes into the visible run. Word's
  behaviour is not measured.
- **`textbox:N` counts per paragraph.** The text model numbers a joined chain's boxes on from its first
  paragraph, so a box in a paragraph whose mark the view deletes is addressed differently by the two. None
  of the fixtures has one.
- **Not measured.** A note's own mark as U+0002, the math alphabets beyond case 17, a multi-paragraph
  cell's, control's or comment's joiner, and Office JS's search inside a text box. A short follow-up check
  could take them; nothing here waits on one.
  - **Check 51 takes the first three** (`test/README.md`, written 2026-10-21; Jason: "we'd better measure
    those 3 things against Word"). It also reads a block control's `Whole` and `Content` ranges, here its
    first paragraph's only, and the math structures. The text box search stays unmeasured. Run the same
    day: section 17.11.
- **Public behaviour that moves** with a release carrying this phase:
  - `Paragraph.text` and everything counted in it: offsets, `search`, `Range` and the editor's spans (its
    install, section 4.4);
  - `Body.text`'s joiners;
  - `segmentsOf`'s and `textOfView`'s default characters;
  - `outline()` becomes async.

### 17.8 Cost, on the editor's `ms-docx-250.docx` (Node 22, the median of five)

| Reader | 0.5.0 | Phase 4 |
|---|---|---|
| `Body.text` | 14 ms | 16 ms |
| `Paragraph.text`, all 9,178 | 13 ms | 14 ms |
| `body.search('the')` | 34 ms | 38 ms |
| `Body.outline()` | 16 ms (sync) | 198 ms under a context made once; 355 ms with a fresh context |

`outline()` costs what the text model's reading costs (section 15.7's 198 ms under a context). It is the
"tens of ms or more" case the editor named. Its remedies are its own `touches()` gate, or an outline per
block under the context: its choice. The sweep took 8.7 s over the 608 loadable `.docx` files of `test/` and
docx4j's test resources, every paragraph's text and whole range, `Body.text` in both views, three searches,
the headers and footers, and `pkg.outline()`, with no exception. The other four are Strict, which nothing
here reads.

### 17.9 Tests

- **`readers.test.mjs`** (12 tests):
  - check 48's and 48b's files against Word 365's JSON, paragraph for paragraph, the whole ranges and
    `Body.text`;
  - check 51's file against its JSON: every paragraph and whole range, `Body.text`, the cell, the controls,
    the comment, the notes, the equations;
  - search, against all of Word's queries;
  - hidden text read and edited;
  - equations: the alphabets, a whole deletion, the refusals, a format;
  - the `textbox:` agreement over every fixture with a text box;
  - D5's refusals, one per kind of write, and addresses naming nothing;
  - the notes' and comments' addresses;
  - `outline()`: its text boxes, options, view and context, and REQ-204;
  - the three defects of section 17.5.
- **`kept-branches.test.mjs`**: section 12.4's refusal test became ED-009's: the seven edits in the branch
  (equal to a resolved load), a tracked replacement, an emptied branch, and the run-level element's
  replace, split, tracked deletion, whole deletion and drawn branch.
- **Moved because the measured text moved:**
  - `content`, `content-c`, `office-js` and `tracking`: `Body.text`'s `\r` and `\t`;
  - `html`: a break U+000B;
  - `text-model`: `officeOffset` in Office JS's characters, hidden text outside `Paragraph.text`;
  - `tracking`: an equation's range is its text;
  - `tracking` again: a search across a picture takes `^g`;
  - `mc-fun`: a block-level element in a cell is in neither the cell's paragraphs nor its value, where
    `textOf` read it into the value.
- **`check48/build.mjs`** exports its files' bytes (`check48Bytes`) and writes them only when run, so the
  tests read check 48's file without it being committed.
- No suite row, agreement comparison or golden moves. 1,117 tests pass.

### 17.10 What phase 4 leaves

- **The editor's install.** `Paragraph.text`'s characters, hidden text and `outline()`'s Promise (sections
  4.4, 6.2 and 17.6).
- **CR-013 phase 2.**
  - A block-level element's paragraphs, rows and cells for the content API, and their addresses: the three
    text boxes of section 17.4.
  - The audit's `textOf` question is moot for `TableCell.text`, `Table.text` and `ContentControl.text`,
    which no longer call it.
- **Text boxes in `search` and `replaceText`, and writes into text boxes** (section 9).
- **Ranges across paragraphs**: CR-014, proposed (section 17.12).
- **Phases 5 and 6**, as section 15.9 has them.

### 17.11 Check 51's answers (2026-10-21)

Jason ran check 51 in Word 365 the same day, in Word's defaults, with a screenshot of the equations (`test/README.md`
check 51; his JSON is `test/fixtures/check51/51-text-word365.json`). It answered section 17.7's three, and two
things more:
- **The joiner is `\r`.** A cell's `value` and `body.text`, a block control's `text` and a comment's
  `content` are their content's `body.text`. They moved from `\n`.
- **A note's own mark is U+0002**, as implemented.
- **The equations.**
  - Word draws Greek capitals and nabla italic too. I had them upright.
  - A digit takes its script's upright family: bold under bold italic, sans-serif under sans-serif.
  - A hyphen-minus is U+2212.
  - A structure is written `,`, its arguments a `-` apart, `.`. All of it is in `mathText.mts` now.
- **`body.text` ends a table row with `\r\n`** in Word's default view: cells a `\t` apart and nothing after the
  last. Checks 30 and 31's saved `body.text` agree, every row `\r\n`. Check 48's `\t\t` was read in All Markup
  with hidden text shown, so the row's end depends on the view, and section 6.2's rule, Word's default view,
  takes `\r\n`. `Body.text` moved.
- **A block control's range** in Office JS holds all its paragraphs (`CCP1\rCCP2\rCCP3\r` for 'Whole'). Here a
  `Range` is one paragraph's, so it held the first. That is the content API's shape since CR-002, not a
  character question. Section 17.12 makes it read as Word's and refuse writes, and CR-014 proposes ranges across
  paragraphs. A one-paragraph control's 'Whole' range ends with its mark, as Word's does.

Every paragraph and whole range, `body.text`, the cell, the controls' text, the comment and the notes of check
51's file now read as Word's. `readers.test.mjs` holds them to the JSON. The tests that read the joiners and
`body.text` over a table moved with them: `comments`, `content-c`, `customxml`, `tracking`. An equation's
range in `tracking` is its structure's text, `,𝑥-2.`. 1,117 tests pass.

### 17.12 A block control's range across paragraphs: read as Word's, writes refused (2026-10-21)

Jason asked why a `Range` here is one paragraph's where check 51's is not, and whether we should align. The answer:
CR-002 scoped it so, for phase B's edits, without a measurement, and a control's range was the first paragraph's,
silently. He accepted both recommendations: CR-014 proposes ranges across paragraphs, and this stopgap went in
here ("yes please").
- **`SpanningRange`** (`Range.mts`) is what `ContentControl.getRange('Whole'|'Content')` gives for a block control
  of more than one paragraph:
  - its `text` is the control's `Body.text`, 'Whole' with the last paragraph's mark: check 51's
    `CCP1\rCCP2\rCCP3\r` and `CCP1\rCCP2\rCCP3`;
  - `paragraphs` are all of them, `search` finds over all of them, and the tracked changes and comments are all
    of theirs;
  - `font` reads the first run's.
- **Every write throws `NotSupportedError` naming CR-014**, before changing anything: `text`, `insertText`,
  `delete`, `font`, the style, `insertParagraph`, `insertOoxml`, `replaceText`, `insertComment`,
  `insertContentControl` and `hyperlink`. The control's own verbs (`text`, `insertText`, `insertParagraph`,
  `clear`, `delete`) act on all of it, as they did.
- **`getRange('End')`** is the last paragraph's end, for any control, where it was the first paragraph's.
- **Left as they are** (CR-014's): `Comment.getRange()`'s array, and `TrackedChange.getRange()` for a change across
  paragraphs (its first paragraph's part).
- `readers.test.mjs`'s check 51 test holds the range to Word's text and each write to its refusal. 1,117 tests
  pass.
