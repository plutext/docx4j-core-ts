# Word fragments: the editor's proposal 46 check (CR-005 section 11)

These files are copied unedited from docx4j-ts-editor `c71ebd6`, `packages/editor-model/test/fixtures/fragments/`
(ED-005 proposal 46, part 6(a)). `manifest.json` gives each file's SHA-1. The editor's README, which describes how
the files were built and what Word did with them, is copied as `EDITOR-README.md`. The PDFs and the Compatibility
Checker screenshot aren't copied.

The editor built the inputs by hand as XML. Word 365 ran them through the portfolio's Word bridge (`22479f7`) on
2026-11-11, in compatibility mode 15 (`-word365`) and mode 12 (`-word365-mode12`). Every file opened silently.

| File | What it is |
|---|---|
| `46a-altchunk.docx` | docx4j's form: the `od:progid=Word.Document` control holding a `w:altChunk` to the fragment |
| `46a-word365*.docx` | Word's save of it: the chunk resolved, the fragment's styles and lists reconciled |
| `46b-altchunk-matchsrc.docx` | the same, with `w:matchSrc` |
| `46b-word365*.docx` | Word's save of it: the fragment's styles written as direct formatting |
| `46c-template.docx` | the template as authored: the `od:progid` control, a plain text control and a rich text control mapped by `w15:dataBinding`, all three bound to the node holding the fragment |
| `46c-word365*.docx` | Word's save of it: the `w15` control filled from the node, the other two showing their placeholders |

`test/fragments.test.mjs` runs `46c-template.docx` through `applyBindings` and holds the result to Word's saves.
