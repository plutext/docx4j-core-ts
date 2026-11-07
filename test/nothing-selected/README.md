# An element that selects nothing: Word's rounds (CR-029)

The sources of the rounds WD REQ-305 and REQ-306 rest on, with Word's text save of each, copied unedited from the share's
`bridge/portfolio/`, with `manifest.json`'s SHA-1 of each. Word 16.0.20326.20232, through the bridge, each file twice, the
runs agreeing. `<name>.txt` is Word's text save and `<name>.word.docx` Word's save, kept for `rw-*` and `te-*`.

- `ul-*`: the 118 URIs of the old published list, and a made-up control, each a Choice with no Fallback. `understood-list.index.tsv` gives each file's URI.
- `uc-*`: 11 namespaces that list lacked, and two controls.
- `c305-*`: block and run level, namespaces Word knows and doesn't.
- `rw-24`, `rw-25`: an unknown namespace at the root and in a body.
- `te-*`: what a removal leaves. `te-00` is the control. `te-01` to `te-05` hold the element, and `te-11` to `te-15` are docx4j-python's `word`-mode saves of them with it taken out. A `w:tc` left with no block is repaired, a `w:tr` left with no cell refused (no save), and a `w:tbl` left with no row dropped silently.

Every file is Choice-only, in a main document part, on one Word build.
