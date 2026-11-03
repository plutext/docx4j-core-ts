# CR-021: A `w:br` outside a run, read as Word reads it (the port of docx4j CR-073)

**Status:** Proposed 2026-11-03, from docx4j CR-073's hand-off (its phase 2, D5), relayed by the portfolio session
(docx4j-portfolio-10, registry `74da802`). Measured here the same day (section 1), and docx4j's three rules tried as a
prototype outside `src/` (section 2). D1 to D6 for Jason (section 5). docx4j's CR-073 is complete (phase 1
`95c6b42d3`, on `VERSION_17_3_6`, unreleased: its rules are in `mc-preprocessor.xslt`, `BrOutsideARunTest`); docx4j-python's
and the editor's checks are their own entries.

## 0. Summary

docx4j's load survey found a `w:br` outside a `w:r` in 206,351 of 3,722,083 documents. It appears directly in a `w:p`,
directly in a `w:hyperlink`, or inside a `w:t`:
- 89% are one web-to-docx generator's blank line, `<w:p><w:br/></w:p>`;
- PHPWord writes `<w:t>a<w:br/>b</w:t>`;
- a template filler nests `w:t` in `w:t`.

Word 365 opens every such shape without a repair, and reads each as a break at that point, all the text kept (docx4j's
probe: thirteen one-case files and Word's saves of them).

Here it is worse than in docx4j 17.3.5, in two ways:

1. **A `w:t`, `w:instrText`, `w:delText` or `w:delInstrText` holding an element makes its whole part unreadable.**
   xmldom parses it. The runtime's unmarshal then refuses the text element's content ("Element text content may not
   contain START_ELEMENT"), and `getContents()` throws `PartUnmarshalException`:
   - for the main part, every reader of the body throws;
   - for a header, the text model reports `content.not_read` and leaves the header out.
   docx4j at least kept the text after the break; here the part is lost. This covers five of the thirteen files.
2. **A `w:br` directly in a `w:p`, `w:hyperlink` or `w:body` is kept, but behaves as misplaced content.** It is typed
   `Br` where it was, and is a misplaced child for CR-016:
   - a `target.misplaced` warning on load;
   - unseen by `Paragraph.text`, `search` and the text model, so the text either side runs together;
   - **a save of its part, once unmarshalled, is refused** (`MisplacedContentException`);
   - untouched, it round-trips byte for byte.
   This covers the other eight files. An editor that changes anything in such a document cannot save it.

The fix is docx4j's three rules as a DOM pass before unmarshalling, through the hook this package already has for the
MCE pass. As a prototype, it reads twelve of the thirteen files as Word's saves read, structure and run properties
included. The thirteenth differs only by docx4j's own D6 (a `w:instrText` outside any field, which Word shows as text).
It also reproduces Word's saves of check 72's two bare-break files (CR-016 section 7.1).

## 1. Measured here (2026-11-03, the working tree over `6eb4f50`, objects 0.7.0)

Each probe file is from `../docx4j/docx4j-core-tests/src/test/resources/br-placement/` (its `README.md` has Word's
table). Each was loaded here and every part unmarshalled. "Word" is this package's reading of Word's save of the same
file (`*-word365.docx`), which loads with no warning in every case. Breaks are shown as `⏎` (U+000B, Office JS's) and
`⇟` (a page break).

| File | The `w:br` | Load | `Paragraph.text` here / Word | Save after unmarshalling |
|---|---|---|---|---|
| `br-1-p-alone` | all of a `w:p` | `target.misplaced` | `""` / `"⏎"` | refused |
| `br-2-p-between-runs` | in a `w:p`, between runs | `target.misplaced` | `"BR2 beforeBR2 after"` / `"BR2 before⏎BR2 after"` | refused |
| `br-3-t-mixed` | in a `w:t` | **`PartUnmarshalException`** | / `"BR3 before⏎BR3 after"` | |
| `br-4-t-nested` | between two `w:t` nested in a `w:t` | **`PartUnmarshalException`** | / `"BR4 before⏎BR4 after"` | |
| `br-5-hyperlink` | in a `w:hyperlink` | `target.misplaced` | joined / a break | refused |
| `br-6-cell` | in a cell's `w:p` | `target.misplaced` | joined / a break | refused |
| `br-7-footer` | in a footer's `w:p` | `target.misplaced` | (the text model) joined / a break | refused |
| `br-8-p-page` | `w:type="page"` in a `w:p` | `target.misplaced` | `"BR8 before"` / `"BR8 before⇟"` | refused |
| `br2-1-run-properties` | D1's eight cases (one in a `w:t`) | **`PartUnmarshalException`** | / Word's twelve paragraphs | |
| `br2-2-instrtext-field` | in a field's `w:instrText` | **`PartUnmarshalException`** | / `"I1 start ⏎I1 result I1 end"` | |
| `br2-3-instrtext-bare-header` | in a bare `w:instrText`, in a header | the header: **`PartUnmarshalException`** | the text model: `content.not_read` for the header | |
| `br2-4-body-br` | in `w:body` | `target.misplaced` | `"B1 after"` / `"⏎B1 after"` | refused |
| `br2-5-body-br-page` | `w:type="page"` in `w:body` | `target.misplaced` | `"B2 after"` / `"⇟B2 after"` | refused |

The text model (`PlainText` with headers and footers) differs from its reading of Word's save in the same places. An
untouched save of each is byte for byte; the package loads in every case, and only the part fails.

The editor session's run of the same files on 0.6.0 (ED-008 section 8.8 item 17, editor `b651abc`) agrees:
- four do not open (br-3, br-4, br2-1, br2-2);
- in the other nine the break is an opaque atom, with no break drawn.

It took the thrown error for the parser's, and concluded that no rule could see these files until the parser tolerated
element children. Here that is not so. The parser (xmldom) reads all thirteen; the refusal is the runtime's unmarshal,
which runs after `XmlPart.preprocessor`, the DOM hook. So a rule in that hook sees the DOM first, and no change to
`@docx4j/jsonix` or the objects package is needed (section 2).

## 2. A prototype of docx4j's rules (scratch, not in `src/`)

docx4j CR-073 section 8's three rules, written as a `DomPreprocessor` and passed as the load's `preprocessor`, followed by
`mceKeep`:
1. a `w:br` in a run container gets a run of its own, with the mark's properties while no run precedes it;
2. a text element holding elements is split into its text and its breaks;
3. a `w:br` in `w:body` goes into the next paragraph.

The prototype is about 100 lines, in the session's scratchpad.

- **All thirteen load.** `Paragraph.text` and the text model equal their readings of Word's saves in twelve.
  `br2-3`'s header differs only by D6: Word shows its bare `w:instrText` as text, and docx4j, by its D6, keeps it a
  field code, so the text model hides it.
- **The saved structure, held to Word's.** This is docx4j `BrOutsideARunTest`'s comparison: text, breaks and their
  types, field characters and paragraph ends, in order, each with its run's bold, italic and size, after a save that
  re-marshals.
  - 14 of the 15 story parts equal Word's saves; the 15th is `br2-3`'s header (D6).
  - D1's run properties are Word's in all eight of `br2-1`'s cases. A break first in its paragraph (R3) or alone (R5)
    takes the mark's `w:sz w:val="28"`. A break after a run (R1, R2, R4, R7, R8) takes none. A break in a bold run's
    `w:t` (R6) stays in the bold run.
  - In `br2-2`, the field code is split round its break, as Word's save splits it.
- **Check 72 agrees.** CR-016 section 7.1 has Word's saves of 72e (a `w:br` in `w:body`) and 72j (a `w:br` in a `w:p`).
  The prototype's reading of each original equals them, structure for structure. That is a third Word run giving the
  same rule.
- **Nothing else moves.** Over the 622 documents under `test/` (fixtures, `mc-fun`, the check saves, the text-extraction
  suite, ...), only 72e and 72j change, and no load fails.
- **Cost.** The prototype makes five tag scans per WordprocessingML part. Over all 622 documents it took 235 ms of
  `unmarshalAll`'s 7,363 ms (3%). One walk would do it.

## 3. Proposal

### 3.1 The pass (`src/opc/repair/breaks.mts`, `repairBreaksOutsideRuns`)

docx4j's three rules exactly (its CR-073 section 8), over a part's DOM, for a part whose document element is in the
WordprocessingML namespace:

1. **A `w:br` in a run container** gets a `w:r` of its own, its attributes kept.
   - The parent must be `w:p`, `w:hyperlink`, `w:smartTag`, `w:customXml`, `w:fldSimple`, `w:ins`, `w:del`,
     `w:moveFrom`, `w:moveTo`, `w:sdtContent`, `w:dir` or `w:bdo`. They are named because a math run holds a `w:br`
     legally.
   - Its nearest paragraph-or-block ancestor must be a `w:p`.
   - The run has the mark's properties (the `w:pPr/w:rPr`, without `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo` and
     `w:rPrChange`) where no earlier sibling on the break's path from its paragraph holds a `w:r`. Otherwise it has
     none (D1).
2. **A `w:t`, `w:instrText`, `w:delText` or `w:delInstrText` holding elements** becomes its stretches of text, each
   under the same name with `xml:space="preserve"`, and its `w:br`s, in order, in the same run.
   - A nested `w:t` gives its text.
   - Text in any other child element is kept.
   - Whitespace between children is kept only where `xml:space` preserves it.
   - A `br` in no namespace is not a break, and only its text is kept: docx4j's survey sample `01af91808eac`.
3. **A `w:br` directly in `w:body`** goes into the next `w:p` as its first run, after its `w:pPr`, with that mark's
   properties. Where no paragraph follows, it becomes a paragraph of its own (unmeasured, as docx4j).

### 3.2 Where it runs (D1)

The repair is its own step in `XmlPart`, run before the package's hook (`mceKeep`, `mcePreprocess` or a caller's own),
on every route that parses a part:
- `getContents()`;
- the private reads;
- `setBytes`;
- `insertOoxml`'s two routes. The fragment route takes the package's hook today, and a created package has none.

A caller's own hook therefore sees the repaired DOM, as docx4j's other templates run before anything else does.

### 3.3 The report (D2)

One part warning per part, as docx4j's one WARN per part (its D3), code `br.outside_run`, shared with docx4j-python
(settled 2026-11-03 through the portfolio session, provisional on Jason's word on each CR; it follows the shared codes'
`subject.what` pattern: `target.misplaced`, `value.out_of_range`, `altchunk.not_expanded`).
- Its fields are docx4j-python's, in this package's camelCase, as `value.out_of_range`'s `word_reads` is `wordReads` here:
  - `partName`;
  - `breaks`: the `w:br`s found outside a run;
  - `textElements`: the `w:t`, `w:instrText`, `w:delText` and `w:delInstrText` holding elements.
- The message is docx4j's WARN: "w:br outside a w:r: n outside a run, m text elements holding elements; read as Word
  reads them".
- It joins the `PartWarning` union, so `XmlPart.warnings` and `pkg.warnings` carry it.

### 3.4 What the save writes (D3)

- **A part that is re-marshalled** writes the repaired shape, which is Word's save's shape (docx4j's D4).
- **A part never unmarshalled** is copied from its source byte for byte, as every untouched part is here. That is also
  the editor's condition: it writes an unedited main part from its source bytes.
- **A private read** (the text model's, `checkContent()`) sees the repair and leaves the part untouched.

### 3.5 CR-016 (D4)

A repaired break is in a run, so it is no longer a misplaced child:
- no `target.misplaced` warning;
- no refusal at save.

Check 72's 72e and 72j then read and save as Word's saves of them. `check72.test.mjs` and `misplaced-content.test.mjs`
pin their old fates: a load warning of `Br`, the readers not seeing the break, the save refused. Those pins move to the
new report and to Word's saves; the other sixteen files are unchanged.

### 3.6 What the readers gain, and what is left

Every reader and mutator works on the tree, so all of these see the breaks and the text with no change of their own:
`Paragraph.text`, `Body.text`, `Range`, `search`, `replaceText`, the text model and Markdown, the `Emulator`,
`selectObjects`, the tracked-change listing (a deleted line break in a `w:delText` kept), and the Word shim.

Not here, as in docx4j:
- **A `w:instrText` outside any field** (D5; docx4j's D6). Word shows it as text.
- **A `w:br` directly in another block container** (`w:tc`, `w:txbxContent`, a block `w:sdtContent`, a header's
  root). docx4j's rule 3 is `w:body` only, and none was seen.
- **`w:tab`, `w:cr` or `w:noBreakHyphen` outside a run.** That is docx4j's catalogue query, not yet made.
- **A Strict document.** This package does not convert Strict (CR-001 section 17.6).

## 4. Tests and the oracle

- **The oracle.** `test/br-placement/` holds docx4j's 26 files (the thirteen probes and Word's saves) and its
  `README.md`, copied unedited from `95c6b42d3` with a SHA-1 manifest, outside `fixtures/` like `test/store-items/`.
- **`test/br-placement.test.mjs`:**
  - each probe against Word's save: the structure (section 2's comparison), `Paragraph.text`, the text model, and the
    report;
  - docx4j's six edge tests mirrored: a math run's `w:br` and a legal `w:r/w:br` left alone; whitespace without
    `xml:space`; the mark's own elements dropped; deleted text; a body `w:br` before a paragraph with no `w:pPr`, and
    before a table;
  - an untouched save byte for byte;
  - a save after unmarshalling, not refused and holding the repaired shape;
  - `insertOoxml` of a misplaced break into a created package;
  - a caller's own hook seeing the repaired DOM.
- **Check 72:** 72e's and 72j's pins move (section 3.5).
- **The gate in `misplaced-content.test.mjs`** stays as it is: no fixture warns.
- **The full suite**, since other tests pin totals.

## 5. Decisions (for Jason)

| # | Question | Options | Recommendation |
|---|---|---|---|
| D1 | Where the pass runs | (a) its own step in `XmlPart` before the package's hook, and in `insertOoxml`'s routes; (b) in the default hook only, beside `mceKeep` (a caller's own hook, or a created package's `insertOoxml`, would go without it); (c) in the objects facade's unmarshal (an objects CR and release; it would reach the facade's other callers too) | (a). It is a reading of the file, not an option. The MCE pass is this package's by the same precedent, and it needs no release elsewhere. |
| D2 | The report | (a) one part warning per part, `br.outside_run` with docx4j-python's fields (`partName`, `breaks`, `textElements`; section 3.3); (b) one warning per break | (a), as docx4j's D3 and docx4j-python's. |
| D3 | What the save writes | (a) the repaired shape where the part is re-marshalled, an untouched part byte for byte; (b) an option to keep the original | (a), as docx4j's D4, under this package's untouched rule. |
| D4 | CR-016's check 72 | (a) a repaired break is no longer misplaced; 72e's and 72j's pins move to Word's saves; (b) keep warning `target.misplaced` beside the repair | (a). The repair is Word's reading, and three Word runs agree. |
| D5 | A `w:instrText` outside any field (`br2-3`), which Word shows as text | (a) out, as docx4j's D6, following docx4j's field model when it decides; (b) read as text here | (a). |
| D6 | Release | (a) with the next release; (b) on its own | (a). |

## 6. Phases and size

One phase:
- the pass, about 150 lines;
- the report;
- the oracle's copy and its test, about 300 lines;
- check 72's pins.

It depends on nothing unreleased: docx4j's rules are the specification, and its probe and Word's saves are the oracle.

## 7. Registry

`core-ts/CR-021`: proposed, `depends_on: [docx4j/CR-073.1]`, sent to the portfolio session.

## 8. Findings beyond this CR (recorded, nothing changed)

- **A break in a field's code.** `Paragraph.text` holds U+000B for a `w:br` in a complex field's code, between `begin`
  and `separate`. Word's save of `br2-2` reads `"I1 start ⏎I1 result I1 end"`. The Office JS reader leaves out the
  `w:instrText` but does not follow the field's code region. Whether Office JS's `text` holds that break is not
  measured; the text model leaves it out (field codes off). A Script Lab read of `paragraphs.items[0].text` on
  `br2-2-instrtext-field-word365.docx` would say.

## 9. Sources

- docx4j:
  - `docs/developer/change-requests/CR-073-br-outside-a-run-read-as-word-reads-it.md` (sections 3, 5 and 8);
  - `docx4j-core-tests/src/test/resources/br-placement/` and its `README.md` (round 1 `a2a971416`, round 2
    `61147016e`, phase 1 `95c6b42d3`);
  - `mc-preprocessor.xslt`;
  - `BrOutsideARunTest`.
- The portfolio: `docs/broken_producer_repairs_across_ports.md` (the family: this package had no repair pass in
  September); `docs/docx_corpus_use.md` section 8.14 (the survey).
- Here:
  - `src/parts/XmlPart.mts` (`preprocessor`, `parse`, `unmarshal`);
  - `src/opc/mce/preprocessor.mts`;
  - CR-016 sections 7 and 7.1 (check 72);
  - CR-012 section 25 (`content.not_read`).
