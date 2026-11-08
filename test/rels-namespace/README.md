# A relationships part in another namespace (CR-025 section 8)

The portfolio's `rels-namespace` round, as docx4j holds it (`docx4j-core-tests/src/test/resources/package-wiring/` at `598fe9212`,
the `rn-*` files, built by its `build_rels_namespace.py`), copied unedited with `manifest.json`'s SHA-1 of each. Word
16.0.20326.20232, through the bridge, each file twice, the runs equal. `<name>-word365.docx` is Word's save of each.

| File | Holds | Word |
|---|---|---|
| `rn-00-control-plain`, `rn-01-control-rich` | The controls: a plain body naming no relationship, and a rich one naming a header, with a note | Silent |
| `rn-02-plain-wrong`, `rn-03-rich-wrong` | `word/_rels/document.xml.rels` wholly in another namespace | Repaired, every relationship read |
| `rn-04-plain-no-namespace` | The same, in no namespace | The same |
| `rn-05-plain-root-wrong` | Its root alone in another | The same |
| `rn-06-plain-children-wrong`, `rn-07-plain-one-wrong` | All its children in another, or the second alone | The same |
| `rn-08-rich-no-part` | No relationships part at all: the comparison, not a case of the rule | Repaired; the header, the style, the settings and the note lost |
| `rn-09-package-rels-wrong` | `_rels/.rels` in another namespace | Repaired, the main part read |
