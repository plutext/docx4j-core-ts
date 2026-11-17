# Office JS shadingColor in Word 365: CR-002 section 67's round

These are the portfolio's Word bridge round `core-ts-table-shading`, and its later sets `-r`, `-2a` and `-3`.
- **Where they ran:** bridge `958f35f`, Word 365 16.0.20326.20318, 2026-11-17.
- **What the jobs are:** each is a `check` job that runs Office JS on `ts-template.docx`, then saves the document (`formats: ["docx"]`).
- **Where these copies come from:** unedited from `bridge/portfolio/<job>/`, under `manifest.json` (SHA-1).
- **What `test/table-shading.test.mjs` does:** runs each job's own script through the Word shim on the same template, and holds the getters, the sets and the save to Word's.

## The files

| File | What it is |
|---|---|
| `ts-template.docx` | Every job's input: ten labelled tables (below). |
| `base/shd1-style-forms-word365.docx` | The template's base: the editor's Word 365 save from its round `row-shading` (its `packages/editor-model/test/fixtures/row-shading/` at `a99c6d1`). |
| `build.mjs` | Writes the template and the staged sets. Its output is these files byte for byte: `node test/table-shading/build.mjs test/table-shading/base/shd1-style-forms-word365.docx <an empty directory>`. |
| `<job>.check-script.js` | The script the job ran: the page's JOB and the common code. |
| `<job>.result.json` | The bridge's record: Word's version, the dialogs, whether the script answered, what the save stripped. |
| `<job>.check-result.json` | What the script reported, for the jobs Word answered: every getter before and after, each set's outcome, each table's XML from `getOoxml()`. |
| `<job>.ts-template.word.docx` | Word's save after the script, for the same jobs. |
| `ts04-row-clear.dialog-1.png` | The dialog every clear raised: WD-004, "There is not enough memory or disk space to update the display.". The same image for all 13 clears. |

## The tables

| Table | What it holds |
|---|---|
| T1 | plain, four rows, no shading |
| T2 | `ShadedTable`, whose own `w:tcPr` shades every cell orange (`FFC000`); rows 1 and 2 their own yellow |
| T3 | `BandedTable`, odd rows orange; row 1 its own yellow |
| T4 | plain; column 2 merged down rows 1 and 2 |
| T5 | the same, the restart yellow |
| T6 | row 1 yellow and blue; row 2 `w:themeFill="accent1"` over a stored `FF0000` (accent1 is `156082`); row 3 `fill="auto"` and none |
| T7, T10 | the table's own `w:tblPr/w:shd` yellow; row 2 cell 2 its own blue |
| T8 | plain |
| T9 | `ShadedTable`; row 1 cell 1 its own yellow |

## The jobs

| Job | Set | What it does | Word |
|---|---|---|---|
| `ts01-gets` | 1 | every getter | answered, saved |
| `ts02-row-hex` | 1 | rows: `#FFFF00`, `FFFF00`, `#00b0f0`, `#123456`, and a theme-filled row | answered, saved |
| `ts03-row-names` | 1 | rows: `yellow`, `Red`, `lightgreen`, `nonsense` | answered (`nonsense`: `InvalidArgument`), saved |
| `ts04-row-clear`, `ts04r-row-clear` | 1, `-r` | six row clears, `null` among them | WD-004, no answer |
| `ts04a` to `ts04f` | `-2a` | one clear each | WD-004, no answer |
| `ts05-row-merged-continuation` | `-3` | the row a merge reaches from above | answered, saved |
| `ts05b-row-merged-clear` | `-3` | that row cleared, under a yellow restart | WD-004 |
| `ts06-row-merged-restart` | `-3` | the row a merge starts in | answered, saved |
| `ts07-cell` | `-3` | cells: hex, a name, a continuation's place (`ItemNotFound`), a theme fill | answered, saved |
| `ts07b-cell-clear-styled`, `ts07c-cell-clear-banded` | `-3` | a cell cleared | WD-004 |
| `ts08-table` | `-3` | tables: plain, styled, over its own shading, over mixed cells | answered, saved |
| `ts08b-table-clear` | `-3` | a table cleared | WD-004 |
| `ts09-tracked` | `-3` | rows, a cell and a table with tracking on | answered, saved |
| `ts09b-tracked-clear` | `-3` | a row cleared with tracking on | WD-004 |

What Word answered, and what this package decided from it, are in CR-002 sections 67.3 and 67.4.
