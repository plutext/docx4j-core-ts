# A main document part whose root isn't `w:document` (CR-028 D10 and D11, WD REQ-318 to REQ-320)

The files WD REQ-318 to REQ-320 rest on, with Word 365's save of each it opened as `<name>-word365.docx`, copied unedited with
`manifest.json`'s SHA-1 of each and where it came from. Word 16.0.20326.20232, through the bridge, each file twice, the runs
agreeing. `manifest.json` gives each file's root and what Word did: `silent`, `repaired` or `refused`.

- `or-*`, `mu-*`, `mr-*` and `rw-10` to `rw-13`: from docx4j's `docx4j-core-tests/src/test/resources/mc-root/` at `4b26e7422`,
  each named, since that directory holds rounds this set doesn't take. `or-*` were built here (CR-028 section 1.1), `mu-*` here
  (section 1.4), `mr-*` by docx4j (its CR-079 section 1.1), `rw-*` by the objects package.
- `rq-*`: the `cr079-root-questions` round, built by docx4j (its `CR-079-notes/build_rq.py`), from the share's
  `bridge/staging/` and Word's saves from `bridge/portfolio/`, the `-a` run of each.

What a root here reads as is `test/main-part-root.test.mjs`'s table. `mc:AlternateContent` as the root is REQ-317's, not this
set's. `or-14-comments` was offered a repair, and Word then kept the comment; REQ-320 gives no text, so that file is held to the
statement and not to Word's save.
