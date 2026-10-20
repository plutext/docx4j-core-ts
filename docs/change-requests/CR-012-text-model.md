# CR-012: The text model: docx4j CR-046's text extraction (views, scope, plain text, Markdown, the block model), held to the specification's conformance suite

**Status:** Proposed 2026-10-19, at Jason's request ("open a text-model CR here (That would be
CR-012)"), on docx4j CR-046 phase 4's hand-off (the docx4j session, 2026-10-18). Not scheduled.
Decisions D1 to D9 (section 6) are Jason's. D2 starts from his decision of the same day, that the
Office JS readers keep calling the objects package's `textOf` and that it is not deprecated. It
takes over CR-010's text-box view (D8), the U+0002 rule (section 12.6) and the read-time
`mce.branches_differ` warning (section 14.7). The editor reviewed it the same day (section 11), and
its asks are in section 4.4. Section 12 (2026-10-20) takes the editor's ED-009 rule for run content
in a kept `mc:AlternateContent`, reading and editing, as the editor asked (Jason: "do as the editor
asks re Q1, Q2"), in phase 4. It also records a defect in 0.5.0 that the rule closes: edits over a
paragraph-level element's text land outside it.

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
- **The understood set** is published as `UNDERSTOOD_NAMESPACES` (REQ-306). A test asserts it
  against docx4j's list.
  - It equals the list's first 107 URIs today.
  - Phase 0 measures, over the corpus, whether any `mc:Choice` requires one of the other 11. If none
    does, the 11 are documented as the difference.
  - If one does, adding it goes through CLAUDE.md's round-trip test first.
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
  - So it is read from the editor's checkout for the measurement, and not committed here unless
    Jason agrees.
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

## 6. Decisions (for Jason)

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

## 7. Phases and sizes

| Phase | What | Size |
|---|---|---|
| 0 | The suite and the oracle (D8 a and b). Copy the fixtures, TSVs and goldens with provenance; the Node conformance runner with an engine that implements nothing (everything pending); the agreement test's runner; the understood-set assertion and the measurement of the 11 | 1 day |
| 1 | REQ-510's level ownership in the Emulator (D9), with `LevelOwnerTest`'s documents and Word's labels | ½ to 1 day |
| 2 | The plain-text model (sections 4.1 and 4.4): `Texts.of`, its part and content overloads with `TextContext`, and `ofResolved`; the source references; the three views and the agreement test passing; the stories, notes, comments, text boxes and addresses (D4); fields cached, codes and the TOC; hidden text; placeholders and bindings; images; properties; tables; the characters; CriticMarkup; the block model; the warnings, the read-time `mce.branches_differ` among them; `Texts.of` measured on `ms-docx-250.docx` | 7 to 9 days |
| 3 | Markdown (REQ-703, REQ-704): headings, emphasis, lists by the numbering model with REQ-703's interrupted-list rules, notes, GFM tables, front matter; `MarkdownExporter.blocks` | 2 to 3 days |
| 4 | The readers (D3, D5, D6): `outline()` on the model; text-box `Paragraph` views at their addresses; U+0002 after check 48; ED-009's run content in a kept branch, read and edited, after check 49 (section 12) | 3½ to 4½ days |
| 5 | Section 5's statements, each its own go: a `fields: update`, b math as LaTeX, c charts (after CR-009) and SmartArt, d embedded, e altChunk | each 1 to 4 days |
| 6 | Block-model parity goldens (D8 c), once docx4j's 17.3.4 work is public | 1 day |

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
- **D2, the Office JS readers.** Items 1 and 2 change `Paragraph.text`, `search`, `Range` offsets
  and `replaceText` for these shapes. Under D2 those keep Office JS's text, and what Office JS's
  `paragraph.text` reads for a branch's text isn't measured. Word displays the selected Choice;
  that's the probe. So **check 49** comes first. It reads `paragraph.text`, `search('flat')` and a
  range over the branch's text in Word 365, on the editor's check file (`mc-for-word.docx`, ED-009
  step S6) or on `mc-branches.docx`, for cases 1, 2, 6 and 8. If Office JS reads the branch, the rule
  is Office JS's and D2 holds. If it doesn't, the question goes back to Jason.
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
