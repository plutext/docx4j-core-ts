# A relationships part that isn't well formed (CR-032)

Three files of docx4j's broken-part family (`docx4j-core-tests/src/test/resources/text-extraction/broken-part/` at the
docx4j 17.3.6 release, tag `docx4j-17.3.6`, `4afcf4be8`, built by its `build_broken_part.py`), copied unedited by name with
Word's saves and texts, and the family's two tables whole, with `manifest.json`'s SHA-1 of each. The rest of the family
(styles, numbering, settings, properties and stores that can't be read; CR-012 section 28 here) isn't copied: it comes with
the suite when that is copied at a release that holds it.

| Case | File | What is broken | Word 365 |
|---|---|---|---|
| 100 | `broken-part-intact.docx` | Nothing: the twin each broken file is read against | Silent |
| 119 | `broken-part-document-rels.docx` | `word/_rels/document.xml.rels` cut after its third relationship, inside `Relationships` | Unreadable content; repaired "Comments". The styles, numbering and settings relationships kept, the comments' lost, their references left as empty comments |
| 120 | `broken-part-package-rels.docx` | `_rels/.rels` cut after its first relationship | Unreadable content; repaired. The main part kept, the three properties parts lost |

`<name>-word365.docx` is Word's save of each, and `<name>-word365-utf8.txt` its Save As text. `expected-tokens.tsv` and
`expected-findings.tsv` are docx4j's tables for the whole family: `test/rels-not-read.test.mjs` reads their rows for cases
119 and 120, the tokens in the three views by the suite's own rule (`conformance.mjs`'s `verify`, the adapter's reading), and
`part.not_read` naming the relationships part in each.
