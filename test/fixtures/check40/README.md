# Word check 40: columns, the three points section 42.5 left open (CR-002)

Built by `node test/fixtures/check40/build.mjs` from the repository root after `npm run build`;
copies in the shared `__tmp/40/`. Word 365. Three parts; for each file, **Save As** the name given
before doing anything else (AutoSave would overwrite the original), do the one action, save, close.

| File | What it holds | What to do |
|---|---|---|
| `40a1-window-unequal.docx` | an AutoFit-to-window table, columns 3,000 and 6,026 twips (1661 and 3339 pct), cells W1A W1B / W2A W2B | caret in **W1A** (the narrow first cell); **Table Layout > Insert Right**; save as `40a1-word365.docx` |
| `40a2-window-unequal.docx` | the same table | caret in **W1B** (the wide last cell); **Table Layout > Insert Left**; save as `40a2-word365.docx` |
| `40b1-span.docx` | a fixed table of three equal columns; row 1 is a cell spanning columns 1 and 2 (`S1AB`) then `S1C`; row 2 is `S2A S2B S2C` | caret in **S2A** (row 2, first cell); **Insert Right**; save as `40b1-word365.docx` |
| `40b2-span.docx` | the same table | caret in **S1AB** (the spanning cell); **Insert Right**; save as `40b2-word365.docx` |
| `40c-officejs.docx` | two fixed 3 x 2 tables, D and E | Script Lab: `check40-script.js` as the script, `check40.html` as the HTML; **Run the check**; copy the JSON back into a file `40c-result.json`; then save as `40c-word365.docx` |

The questions, and what the engine (9152cd2) writes meanwhile:

- **40a1 and 40a2, the window-mode rescale with unequal columns.** Check 39's columns were
  equal, so it could not tell a proportional rescale from an equal share. The engine copies the
  neighbour's grid width and rescales proportionally: 40a1 (neighbour 3,000) gives a grid of
  2252 / 2251 / 4523 and cells 1247 / 1247 / 2506 pct; 40a2 (neighbour 6,026) gives 1799 / 3614 /
  3613 and 996 / 2002 / 2002 pct. An equal share would give about 3009 / 3009 / 3008 and 1666 /
  1667 / 1667 either way. What `w:tblGrid` and the first row's `w:tcW` say in each save answers it.
- **40b1 and 40b2, a column inserted beside and inside a spanning cell.** 40b1's boundary (after
  `S2A`, between columns 1 and 2) lies inside `S1AB`'s span: the engine widens the span to 3 (its
  width to 9,027 twips) and inserts a cell only in row 2; the grid becomes 3009 / 3009 / 3009 /
  3008 and `w:tblW` 12,035. Does Word do the same, or split the spanning cell? 40b2's boundary
  (after `S1AB`, between columns 2 and 3) is at a cell edge in both rows: the engine inserts a cell
  in each row, row 1's copying the spanning cell's properties (less its span) and the grid gaining
  the width of column 2 (the one beside the boundary, 3,009), not the spanning cell's 6,018. What
  width does Word give the new column, and what does row 1's new cell look like?
- **40c, through Office JS with tracking on.** `Table.deleteColumns(1, 1)` on table D, then
  `Table.addColumns('End', 1, [['E1D'], ['E2D']])` on table E; the snippet records each table's
  XML before and after and what `getTrackedChanges()` lists, and the save shows what Word keeps.
  Word's own Delete Columns is not tracked (known issues, entry 2); the engine's `deleteColumns`
  removes untracked. Does Office JS's deletion track anything, refuse, or remove? For the add: does
  Word write a `w:tcPrChange` on the new cells as well as on the old (the engine writes none on
  the new), and does the add take the Insert Right form of check 39?

Then say what you saw, and the saves are compared here against the engine's output, part by part.
