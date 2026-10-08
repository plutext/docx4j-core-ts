# Word check 41: columns, the two points check 40 left (CR-002 section 42.6)

Built by `node test/fixtures/check41/build.mjs` from the repository root after `npm run build`;
copies in the shared `__tmp/41/`. Word 365. For each file, **Save As** the name given before doing
anything else (AutoSave would overwrite the original), do the one action, save, close.

| File | What it holds | What to do |
|---|---|---|
| `41a-properties.docx` | a fixed two-column table; column 1 yellow, centred vertically, a thick red left border, centred bold text; column 2 blue, bottom-aligned, a dashed green right border, right-aligned italic text | caret in **P1A**; **Table Layout > Insert Right**; save as `41a-word365.docx` |
| `41b-properties.docx` | the same table | caret in **P1B**; **Table Layout > Insert Left**; save as `41b-word365.docx` |
| `41c-officejs.docx` | table F AutoFit to contents, table G AutoFit to window, 3 x 2 each | Script Lab: `check41-script.js` as the script, `check41.html` as the HTML; **Run the check**; copy the JSON back into a file `41c-result.json`; then save as `41c-word365.docx` |

The questions:

- **41a and 41b: whose properties the new column takes.** Check 40 showed the new column's
  *width* is the column to the right of the boundary's; the probes' cells were plain, so the cell
  properties (`w:tcPr`: shading, vertical alignment, borders) and the empty paragraph's properties
  (`w:pPr`, its mark's `w:rPr`) were not measured. The engine copies the right-hand cell's `w:tcPr`
  less span and merges, and gives the new paragraph no properties. In 41a the boundary is between
  the yellow and the blue column with the caret on the yellow side; in 41b the same boundary with
  the caret on the blue side. Does the new middle column come out blue both times (the right-hand
  cell), yellow both times (the left), or the caret's colour (yellow then blue)? And its empty
  paragraph: centred bold, right-aligned italic, or plain?
- **41c: Office JS's add on a contents-mode and a window-mode table, tracking on.** Check 40c
  measured a fixed table. For AutoFit to contents the engine records the table change too, but
  since nothing in `w:tblPr` or the grid record differs it lists no table change, only the text
  insertions (checks 24 and 26's rule). What does Office JS write (`w:tcW` `auto` on the new
  cells? a record on them?) and list for F? And for G, the window table: does Word rewrite every
  cell's percentage as check 39 saw for a hand insert, and list one `Formatted` change?

Then say what you saw, and the saves are compared here against the engine's output, part by part.
