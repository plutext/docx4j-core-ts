# An undeclared prefix in `mc:Ignorable` (CR-030, WD REQ-323)

The portfolio's `undeclared-ignorable` round, with Word 365's save of each file it opened as `<name>-word365.docx`, copied
unedited from the share (sources from `bridge/staging/undeclared-ignorable/`, saves from `bridge/portfolio/`, the `-a` run of
each), with `manifest.json`'s SHA-1 of each. Word 16.0.20326.20232, each file twice, the runs agreeing, 8 November 2026.

- `ui-00`: the control. `ui-01` to `ui-03`: the main part's `mc:Ignorable` lists a prefix it doesn't declare, unknown, known,
  and several. `ui-04`, `ui-05`: the prefix declared only below the root, known and unknown. `ui-06`: an unknown prefix declared
  on the root. `ui-07`: a leading space in the list. `ui-10` to `ui-13`: the same as `ui-01` in a header, the footnotes, the
  styles and the settings part. `ui-14`: a known prefix undeclared in all five.

Word repairs `ui-01` to `ui-04` and `ui-10` to `ui-14`, losing nothing, opens `ui-00`, `ui-06` and `ui-07` silently, and refuses
`ui-05`, whose element in that namespace isn't ignorable.
