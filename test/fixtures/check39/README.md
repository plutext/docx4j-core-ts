# Word check 39: a column inserted in each of Word's table layout modes (CR-002 section 42)

Built by the editor's `packages/editor-model/test/fixtures/build/table-autofit-probe.mjs`
(docx4j-ts-editor ED-005 section 12.78): one document, three two-column tables of the text width
(9,026 twips), each under a paragraph naming its mode.

| Table | `w:tblLayout` | `w:tblW` | `w:tcW` |
|---|---|---|---|
| Fixed column width | `fixed` | 9026 dxa | 4513 dxa each |
| AutoFit to window | absent | 5000 pct | 2500 pct each |
| AutoFit to contents | absent | 0 auto | 0 auto each |

**Run in Word 365, 2026-10-09 (Jason):** the caret in the first cell of each table, Table Tools >
Layout > Insert Right, Save As `table-autofit-probe-word365.docx`. What it wrote:

| Table | `w:tblW` | `w:tblGrid` | `w:tcW` of the first row |
|---|---|---|---|
| Fixed | 13539 dxa | 4513, 4513, 4513 | 4513 dxa each |
| Window | 5000 pct | 3005, 3006, 3006 | 1666, 1667, 1667 pct |
| Contents | 0 auto | 593, 222, 597 | 0 auto each |

The editor holds its own `table.addColumns` to this save (`test/check39.test.mts` there; the files
are the same as its `fixtures/tables/check39/`). Copied here for CR-002 section 42, the engine's
`addColumns` / `insertColumns`. A Word 2010 save is optional (below compatibility mode 15 a
percentage width resolves against a column widened by a cell margin at each end).
