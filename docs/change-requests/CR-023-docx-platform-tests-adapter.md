# CR-023: An adapter for docx-platform-tests, the public conformance matrix

**Status:** Proposed 2026-11-03, from Jason's ruling F2 of the same day (portfolio `mcp_strategy.md` section 14; registry
`portfolio/docx-platform-tests`, `48ba846`). He wants docx4j, this package and docx4j-python measured locally against
`open-agreements/docx-platform-tests`, and adapters submitted upstream for whichever he chooses. docx4j's adapter is
written (`058ec1d4b`, `docx4j-platform-tests-adapter/`). The protocol and the scenarios were read here from the clone at
`~/git/untrusted/docx-platform-tests-untrusted` (head `33b4fff`), read only, nothing in it run. D1 to D5 for Jason
(section 4).

## 0. Summary

docx-platform-tests is "cross-implementation conformance tests for WordprocessingML in the web-platform-tests
tradition": each scenario cites an ECMA-376 clause, runs one operation through a library's adapter, and asserts on
the output. Seven libraries have columns today (Safe Docx, python-docx, LibreOffice, Open XML SDK, docx-js, SuperDoc,
docx-rs); no docx4j and no port.

The suite's protocol (`docs/benchmarking.md` section 3 in the portfolio has it in full):
- An adapter is a CLI the runner calls once per scenario:
  `<cmd> --protocol-version 1 --operation op.json --input in.docx --output out.docx`.
  - `op.json` is the operation alone; the adapter never sees the assertions.
  - Exit 0 means the output was written; 1 an error, on stderr; 2 declined, with one line on stdout ("never
    approximate"); 3 a protocol mismatch.
- The bar is "glue, not algorithms". An operation needing an algorithm the library lacks is declined.

The 44 scenarios use 21 operations, and every one but the cell merge maps onto this package's existing API (section 1).
The adapter is a small Node CLI over the content API, kept outside the published package.

## 1. The operations, mapped (read 2026-11-03)

| Group (scenarios) | Operation | Here |
|---|---|---|
| tracked changes (15) | `acceptAllTrackedChanges` (7), `rejectAllTrackedChanges` (8) | `pkg.acceptAllRevisions()`, `rejectAllRevisions()` (CR-011), held to docx4j's revisions goldens |
| find-replace (2), content controls (4), tables (2) | `replaceFirstTextOccurrence` | `body.search(findText)` and `insertText(replaceText, 'Replace')` on the first hit. Search matches across runs, so docx4j's declined case (the match across runs) should pass here |
| comments (3) | `addCommentOnFirstTextOccurrence` (2) | the first hit's `insertComment(text)`, with the package's author set to `commentAuthorName` |
| | `removeAllComments` (1) | `body.getComments()` and `delete()` each; the comment parts go with the last, as Word does (CR-002 section 54) |
| generation (9) | `appendParagraphWithText` (2), `insertParagraphAfterAnchorText` | `body.insertParagraph(text, 'End')`; the hit's paragraph's `insertParagraph(text, 'After')` |
| | `composeDocumentWithParagraphs` (2), `...WithTable`, `...WithHyperlink`, `...WithCompatibilityMode` | `createPackage()` (its settings write `compatibilityMode` 15), `insertParagraph`, `font`, `insertTable(rows, columns, 'End', values)`, `Range.hyperlink` |
| | `composeDocumentWithHeaderText` | no Office JS `getHeader` here: glue over the parts layer (a `HeaderPart`, its relationship, and the `sectPr`'s `w:headerReference`), or declined (D3) |
| styles and numbering (5) | `applyParagraphStyleToAnchor`, `formatFirstTextOccurrence` | `paragraph.style` (`styles.ensure`'s built-ins, CR-002 section 41); the hit's `font.bold` |
| | `applyNumberingToAnchorParagraph`, `composeDocumentWithNumberedList` | `attachToList(numId, level)`; `startNewList()` with the level's format |
| | `setDefaultFooterText` | as the header (D3) |
| tables (6) | `appendTableRow`, `deleteTableRowAtIndex`, `setTableCellText` | `table.addRows('End', 1, [cellTexts])`; `rows.items[i].delete()`; `getCell(r, c).value` |
| | `mergeTableCellsInRow` | no merge API here (CR-002 section 42 adds and removes columns only): declined, as docx4j declines it |

**The four content-control scenarios don't load here today** (measured 2026-11-03 on their `input.docx`, read as data):
- Each holds an element in its `mc:Ignorable` namespace inside `w:sdtPr` (`ext:opaqueExtension`). The runtime's
  unmarshal refuses it ("could not be unmarshalled as is not known in this context and the property does not allow
  DOM content"), so the main part is unreadable. That holds by default, under `mcePreprocess`, and under CR-021's pass.
- docx4j 17.3.5 failed the same four, for another reason; it was fixed in `86a203f16`.
- The untouched save is byte for byte.
- An ignorable attribute (`ext:opaqueAttribute` on `w:sdt`, or one on `w:p`) is dropped silently on a re-marshal, so
  the scenarios' preservation rows would fail even once the part is read.
- Both are load and preservation gaps of their own, not the adapter's: a finding for Jason, outside this CR. MCE lets a
  consumer ignore what it doesn't understand, but this package refuses the part where Word reads it.

`mc:Ignorable`'s private prefix stays declared on a re-marshal on every path (default, `mcePreprocess`, CR-021's pass,
including a part CR-021 repaired), so docx4j's second finding, the declaration dropped after its preprocessor, doesn't
occur here.

The assertions are XPath counts and existence checks, text at offsets, numbering formats, a comment and its anchor, a
hyperlink's target, a declared `mc:Ignorable` namespace, and `canonicalXmlEquals` (11, graded pass-divergent where only
it fails).

## 2. Proposal

- **Where:** `platform-tests-adapter/` at the repository's root.
  - `adapter.mjs` is the CLI. It imports the package by name, from `dist/` for the local column and from the published
    release for an upstream submission.
  - `README.md` says what is library API and what is glue, with a maintenance policy, as the suite asks.
  - `install.sh` uses no dependency beyond this package.
  - It is outside `src/` and the `exports` map, and outside `package.json`'s `files`, so nothing published changes.
- **The local column (D2).** The suite's runner is code from an untrusted clone. Running it (`npm ci --ignore-scripts`,
  `npx tsx src/run.ts --registry <ours> --results <ours>`) needs Jason's word in this session, as it did in docx4j's.
  - Without it, a small evaluator of our own reads each `scenario.json` as data and checks its assertions on the output,
    as docx4j's session did. That covers every assertion but `canonicalXmlEquals`, whose rows are then reported as
    "pass*".
  - The results are recorded in this CR and sent to the portfolio for `benchmarking.md` section 3's table.
- **What it shows us.** Every row is held against docx4j's column and python's when they exist. A difference between
  docx4j's and ours is investigated: the port is held to docx4j's answers, and Word's checks decide where they differ.

## 3. Registry

`core-ts/CR-023`: proposed, `depends_on: [portfolio/docx-platform-tests]`, sent to the portfolio session.

## 4. Decisions (for Jason)

| # | Question | Options | Recommendation |
|---|---|---|---|
| D1 | Write the adapter now | (a) yes, the local column first; (b) after docx4j's column is run by the suite's own runner | (a). It is glue over existing API, and the evaluator needs no permission. |
| D2 | How the local column is run | (a) our own evaluator over the scenarios' assertions, as docx4j's; (b) the suite's runner, with your permission to run the clone's code; (c) both | (a) now, and (b) when you run it, or say this session may, with the command in section 2. |
| D3 | A header or footer made where the document has none | (a) glue over the parts layer, recorded in the README as glue; (b) declined until the content API has Office JS's `sections.getFirst().getHeader('Primary')` | (a) for the local column; the README names it. A content API for headers is a CR-002 item of its own. |
| D4 | The merge of cells in a row | (a) declined, as docx4j declines it; (b) a merge in the content API first | (a). |
| D5 | Upstream | (a) submit when you choose, from a released version; (b) local only | (a), your call, as F2 rules. |

## 5. Sources

- The portfolio: `docs/mcp_strategy.md` section 14 (F2); `docs/benchmarking.md` section 3 (the protocol, docx4j's
  preliminary column).
- The clone (read only, not run): `README.md`; `scenarios/*/*/scenario.json` (44); `registry/adapters.json`;
  `adapters/python-docx/`.
- docx4j: `docx4j-platform-tests-adapter/` (`058ec1d4b`).
