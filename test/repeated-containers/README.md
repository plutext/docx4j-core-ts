# A child given more than once (CR-028 phase 2)

docx4j's `docx4j-core-tests/src/test/resources/repeated/` at the docx4j 17.3.6 release (tag `docx4j-17.3.6`, `4afcf4be8`), all
326 files copied unedited, with `manifest.json`'s SHA-1 of each. 112 sources: `<name>.docx` is the source, `<name>-word365.docx`
Word 365's save of it and `<name>-word365.txt` its Save As text. Word refused `nd-10` to `nd-14`, which nest a `w:document`, so
they have neither.

| Round | Files | What is repeated |
|---|---|---|
| `rc-` | 16 | Two copies: `w:body`, `w:sdtContent`, `w:txbxContent`, `w:pPr`, `w:rPr`, `w:sectPr`, `w:tblPr`, and on/off and valued leaves |
| `rt-` | 20 | Three copies of the same, four of `w:txbxContent`, bodies with and without a `w:sectPr`, and two `mc:AlternateContent` roots |
| `sm-` | 40 | WD REQ-322's merge round: `w:sectPr`'s children and attributes, references, `w:tblPr`, `w:ind`, `w:spacing`, `w:rFonts`, `w:vanish`, and one `mc:AlternateContent` root (`sm-52`) |
| `ft-` | 12 | Header and footer references across copies and joined bodies |
| `nd-` | 11 | A `w:body` nested in a `w:body`, and a nested `w:document` (refused by Word) |
| `op-` | 6 | A child given twice in one parent |
| `c79-` | 7 | docx4j CR-079's phase 0: `w:sdtContent` at run, row and cell level, `wps:txbx`, a mark's `w:rPr`, `w:numPr`, `w:pPrChange` |

`test/repeated-containers.test.mjs` runs docx4j's `RepeatedChildrenTest` comparisons, which are python's `test_repeated.py`, on this
package's reads and saves. Then it holds this package's own readers on each source to the same readers on Word's save: the text
model, `Body.text`, `Paragraph` and `Font`. It also checks every load route, the save and its gate, and a write.
