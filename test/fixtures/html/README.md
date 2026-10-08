# The HTML corpus (CR-005 section 8; the editor's paste corpus, ED-003 section 8.3)

What Word puts on the clipboard for a selection, as `text/html`, captured by Jason on
2026-09-26 with Word 15 (the current release; the reference machine's Word 2010 is the round-trip
oracle, not the clipboard's) and committed under `results/`, one file per capture. The converter in
`src/html/convert.mts` (moved here from the editor 2026-10-08, ED-001 decision 26) is tested over these
files (`test/html.test.mjs`); each has an `.expected.json` beside it, the converted form pinned
(regenerate with `UPDATE_HTML=1` when a deliberate change moves them, and say why in the commit
message). The captures are in this directory, the terminal's under `terminal/`.

The captures are UTF-8 with a byte-order mark and Windows line ends, and their non-ASCII text
arrived doubly encoded (`Get-Clipboard` read the clipboard's UTF-8 as Latin-1 and `Out-File`
wrote it as UTF-8 again: a non-breaking space is "Â" followed by one); the test's loader undoes
that reading, since a browser hands the editor the text decoded once. Word 15 writes a tab as a
span with `mso-tab-count` holding non-breaking spaces, and lists as `<ul>` and `<ol>` with
`mso-list` on the items.

## The source document

`paste-samples.docx` is built by `../build/paste-samples.mjs` with the engine (so it is ours to
redistribute): a title and an instruction paragraph, then six numbered sections, each one capture.
Open it in Word on the reference machine.

## Capturing

For each section, select from its heading down to the line before the next heading, copy
(Ctrl+C), then in PowerShell save the clipboard's HTML flavour, which carries Word's dialect
(`class=MsoNormal`, `mso-list`, the conditional comments, the `mso-` styles):

```powershell
Get-Clipboard -TextFormatType Html | Out-File -Encoding utf8 headings.html
```

The file begins with the `CF_HTML` header (`Version:`, `StartHTML:`, `StartFragment:` and so on)
before `<html>`; keep it, the converter skips it and the header's fragment markers say what Word
counted as the selection. Save the captures under these names, in `results/`:

| File | Section of the document | What it proves |
|---|---|---|
| `headings.html` | 1. Headings and body text | heading levels by style name, plain paragraphs |
| `inline.html` | 2. Inline formatting | bold, italic, underline, strike, superscript and subscript, a run with two marks, a line break, a tab |
| `bullets.html` | 3. A bulleted list | a bulleted list two levels deep, the paragraph after it |
| `numbering.html` | 4. A numbered list | a numbered list two levels deep, the paragraph after it |
| `table.html` | 5. A simple table | a three-by-three table with a header row, the paragraph after it |
| `dropped.html` | 6. What the editor drops on paste | alignment, indent, font, size, colour and highlight dropped and reported |
| `all.html` | the whole document (Ctrl+A) | everything in one paste |

The title and the instruction paragraph belong to `all.html` only.

The `inline` and `all` captures were redone on 2026-09-26 from the rebuilt document, whose
inline runs each carry one mark (the first build's inherited the previous run's formatting, so
its marks nested); the mark stack itself is exercised by `<b><i>bold italic together</i></b>`.

## Terminals (ED-005 section 12.67)

`terminal/konsole-26.04.html` is what Konsole 26.04.3 (Manjaro, Wayland) put on the clipboard as
`text/html` for a shell's output selected and copied with Ctrl+Shift+C, its profile's "Copy text as
HTML" on (the default). It was saved by `wl-paste -t text/html` on 2026-10-06. It is an XHTML
document whose body is one span in `font-family:monospace`, holding a span per coloured run, each
with `color` and `background-color`, and a `<br>` at each line's end. `terminal/konsole-26.04.txt`
is the same copy's `text/plain`, as Jason pasted it into the session. The editor pastes such HTML
as its plain text (`isTerminalHtml`); the converter alone would make a paragraph of each span.

A copy made inside Claude Code carried `text/plain` alone the same day, `/copy` and a selection
alike, whatever Konsole's setting (its type names suggest Claude Code puts it on the clipboard
itself; not confirmed), so it needs no rule.
