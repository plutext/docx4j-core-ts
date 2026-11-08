# An `External` officeDocument relationship beside an internal one (CR-031)

The portfolio's `x-ext` round, from the share's `bridge/portfolio/x-ext-*-a` (each `-b` job holds the same bytes), copied
unedited on 2026-11-09 with `manifest.json`'s SHA-1 of each. `manifest.json` also records Word 16.0.20326.20232's two opens of
each through the bridge: in all eight, a repair offered ("Word found unreadable content ..."), pressed Yes, and then "Word
experienced an error trying to open the file", with no sign of a fetch. So Word has no reading of these files to follow, and
WD REQ-206's NOTE refuses them at `repair: refuse` and `repair: word`.

| File | `_rels/.rels` holds |
|---|---|
| `x-ext-first-source` | An `External` officeDocument relationship to `http://example.org/x.docx`, then the internal one to `word/document.xml` |
| `x-ext-last-source` | The internal one, then the `External` one |
| `x-ext-first-saved`, `x-ext-last-saved` | The same two orders, in python's save of each source (CR-031 section 1) |

At `repair: all` this package reads `word/document.xml`, drops the `External` relationship and says so
(`package.external_main_part`). CR-031 D7's Word round opened those saves: `ctra-<stem>.docx` is this package's
`repair: 'all'` save of each, sent as `bridge/staging/core-ts-repair-all` and opened twice on 2026-11-09 (bridge `4dcb436`,
with the portfolio's control, any dialog stopping the run). All eight opened with no message, and Word's text of each was
`DOCUMENT`. `ctra-<stem>-word365.docx` and `-word365.txt` are Word's save and text from the first run, the second's text the
same; `manifest.json` records both runs.
