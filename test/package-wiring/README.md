# Two main parts: docx4j's probes and Word 365's saves (CR-024 section 9)

Seven files of docx4j CR-075's `package-wiring` family and Word's save of each, copied unedited from
`docx4j-core-tests/src/test/resources/package-wiring/` at `27d94d04b`, with `manifest.json`'s SHA-1 of each. The family's
own README has the rest: Word 16.0.20326.20232, 7 November 2026, through the bridge, each file twice, the two runs agreeing.

- `two-main-rels-document-first.docx`, `-other-first.docx`: one `_rels/.rels` with two `officeDocument` relationships.
- `two-main-rels-zip-other-first.docx`, `-types-other-first.docx`, `-names-swapped.docx`: one factor changed at a time: the
  zip entries, the content-type Overrides, the names.
- `one-main-rel-other-override-last.docx`, `one-main-rel-document-override-last.docx`: one relationship, and a second main
  part no relationship names, its Override last.

`<name>-word365.docx` is Word's save. The five `two-main-rels` files offer repair, and the saves here are each a first open
in a Word session, which reads the last relationship's target whatever the format or the Overrides. A later repaired open
in the same session alternates, so these saves are what a fresh Word reads, not what Word always reads. The two
`one-main-rel` files open silently, the related part read.
