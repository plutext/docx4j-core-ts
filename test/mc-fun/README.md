# mc-fun: `mc:AlternateContent` at every level (CR-013)

docx4j's `mc-fun` probes, copied unedited from `docx4j-core-tests/src/test/resources/mc-fun/` at
docx4j `69af8cf42` (`VERSION_17_3_4`; the files are unchanged since `1cff2cd5c`, docx4j CR-051 phase 1,
which added them; Apache-2.0). `manifest.json` names the commit and each file's SHA-1, and
`mc-fun.test.mjs` fails on a file that isn't listed or doesn't match: copy again, never edit.

They are Jason's and docx4j's session's (docx4j CR-051 section 3). `<name>-word365.docx` is Jason's save
of `<name>.docx` in Word 365, from docx4j CR-051's Word checks. Word resolves each element on save,
keeping only a run-level text box's (`mc-txbx-run`, Word's own form), so a save shows the branch Word
read. The probes are:

| Probe | Where the element stands | Its branches |
|---|---|---|
| `mc-text-newest-first` | in a `w:r` | a `w:t` each |
| `mc-run-newest-first` | in a `w:p` | a `w:r` each |
| `mc-sdt-run-newest-first` | in a run-level `w:sdtContent` | a `w:r` each |
| `mc-block-newest-first` | in `w:body` | a `w:p` each |
| `mc-block-nofallback` | in `w:body`, with no Fallback | a `w:p` each |
| `mc-sdt-block-newest-first` | in a block-level `w:sdtContent` | a `w:p` each |
| `mc-tc-block-newest-first` | in a `w:tc` | a `w:p` each |
| `mc-row-newest-first` | in a `w:tbl` | a `w:tr` each |
| `mc-sdt-row-newest-first` | in a row-level `w:sdtContent` | a `w:tr` each |
| `mc-cell-newest-first` | in a `w:tr` | a `w:tc` each |
| `mc-sdt-cell-newest-first` | in a cell-level `w:sdtContent` | a `w:tc` each |
| `mc-mixed` | one element each in a run-level control, a hyperlink, a paragraph and the body | `w14` Choices and Fallbacks |
| `mc-txbx-block`, `mc-txbx-sdt-block`, `mc-txbx-tc-block` | in the body, a block-level control, a `w:tc` | a paragraph holding a text box each |
| `mc-txbx-run` | in a `w:r`, Word's own form | a text box each |

"Newest first" means ten Choices requiring `w16sdtfl`, `w16du`, `w16sdtdh`, `w16cex`, `w16`, `w16cid`,
`w16se`, `w15`, `w14` and `wp14`, in that order, then a Fallback. Word 365 takes `w16sdtfl`. Each branch's
text names the namespace it requires.

`mc-run-word365.docx` is Word's save of Jason's original-order `mc-run.docx`, which isn't in docx4j's
resources (it is on his share, `Office 2016/mc-fun/`). So 13 probes here have a Word save.

They are kept out of `fixtures/` so that the parity, XPath and revisions harnesses, and the tests that
walk `fixtures/`, don't take them up.
