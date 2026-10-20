# The text extraction conformance suite (CR-012 phase 0)

What the text model (CR-012) is held to: the text extraction specification's conformance suite
(D8 a) and its agreement test (D8 b), copied from docx4j, which cut them (its CR-046, CR-050 and
CR-053). The specification is the portfolio's `docs/text_extraction_specification_wd.md`; its
`[PROBE]` reference names the docx4j commits the suite is normative at.

| Path | What |
|---|---|
| `suite/` | docx4j's `docx4j-core-tests/src/test/resources/text-extraction/`, whole and unedited: the fixtures (`phase0/*.docx`, `*.docm`), the tables (`expected-tokens.tsv`, `expected-findings.tsv`, `expected-order.tsv`, `case-statements.tsv`, `goldens.tsv`, `expected-inspection*.tsv`), the goldens (`phase0/*.txt`, `*.md`), the contracts (`GOLDENS.md`, `INSPECTION.md`), the probes' Word evidence and the builder's support files. Its own READMEs describe it. |
| `revisions/` | The CR-043 fixtures of docx4j's `revisions/` that `test/fixtures/` does not already hold (9 of 128), at their paths there. |
| `understood-namespaces.txt` | docx4j's published understood set (`docx4j-core/src/main/resources/org/docx4j/jaxb/`, REQ-306). |
| `manifest.json` | The docx4j commit, and for every file docx4j's path, its SHA-1 and its path here. The 119 revisions fixtures this repository already held are listed at their `test/fixtures/` paths. |
| `conformance.mjs` | The runner: docx4j's `ExtractionConformance`, mirrored. |
| `agreement.mjs` | The agreement test's runner: docx4j's `ViewAgreementTest`, mirrored. |

`test/text-extraction.test.mjs` checks every file against the manifest, the protocol (the
runner's own tests, as docx4j's `ExtractionConformanceTest` has them), the suite with no engine
(every row pending, with its phase), the agreement test's fixtures and copy route, and the
understood set against docx4j's list.

The suite lives here rather than under `test/fixtures/`: the weekly parity workflow's
`RevisionsHarness` writes a golden for every `.docx` with revision markup at any depth under
`test/fixtures/`, and these files are held to the suite, not to that harness.

## Provenance

Copied 2026-10-21 from docx4j `1d108c0af` (`VERSION_17_3_4`, CR-053 phase 2, committed 2026-10-21 and
not yet pushed), the newest committed suite WD 0.27 (portfolio `804007d`) cites. Its suite differs
from the pushed `5c691fc1f`, which equals WD 0.26's `86449c013`, in case 36's and case 63's HTML
chunk rows, which need an HTML reader, and in three goldens recut without one. WD 0.27 also cites
CR-053 phase 3, a fifth fixture (`extraction-probe-html.docx`), not yet committed in docx4j: copy
again once it is.

## Copying again

```
node scripts/copy-text-extraction-suite.mjs ../docx4j <commit>
```

It reads the commit with `git archive`, never the working tree, replaces `suite/` and `revisions/`
whole, and rewrites `manifest.json`. Review the difference before committing it, as `GOLDENS.md`
asks of a recut: a changed golden or table changes what the text model is held to. Then port any
change to docx4j's `ExtractionConformance` or `ViewAgreementTest` into the two runners, recount
the pending rows in `text-extraction.test.mjs`, and compare the runner's rows with docx4j's own
report (`docx4j-core-tests/target/text-extraction/conformance.tsv`): the same rows, in the same
order. Nothing here is edited by hand; a fault in the suite goes to docx4j.
