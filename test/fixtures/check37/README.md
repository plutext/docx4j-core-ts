# Word check 37: escaped XHTML bound through an `od:xpath` entry (CR-005 section 8.2)

Built by `node test/fixtures/check37/build.mjs` from the repository root after `npm run build`
(docx4j-ts-editor ED-005 proposal 44, part 5). Two files, each with three controls:

1. a block-level rich text control tagged `od:xpath=x1&od:ContentType=application/xhtml+xml`;
2. a run-level rich text control tagged `od:xpath=x2&od:ContentType=application/xhtml+xml`, inside a
   paragraph that goes on after it;
3. for contrast, a plain text control with a `w:dataBinding` to the same node as 1.

The first two carry no `w:dataBinding`: their XPath is the XPaths part's entry, as docx4j's
`bind.xslt` reads it. The data part holds the markup escaped: a heading, a paragraph with bold and a
link, a two-item list and a two-by-three table for `x1`; two paragraphs for `x2`.

| File | What it is |
|---|---|
| `37a-template.docx` | as authored: the three controls hold placeholder text |
| `37b-bound.docx` | after the engine's `applyBindings({ html: { parser } })`: 1 holds the converted blocks, 2 the first paragraph's text (the second noted as dropped), 3 the escaped markup as text |

## What to do, in Word 2010 and in Word 15

For each file: open it, look, then save it beside the original as `37a-word2010.docx`,
`37a-word15.docx`, `37b-word2010.docx`, `37b-word15.docx` (Save As, the same format).

What to look for:

- **37a, controls 1 and 2:** the placeholder text, unchanged. Word has no binding to act on. Does
  either Word complain about the tag, or change the control (its title, its content) on open or
  save?
- **37a, control 3:** the escaped markup as text (`<h1>Delivery terms</h1><p>...`), which is Word
  filling a text binding from the node's string value.
- **37b, control 1:** a heading "Delivery terms" in Heading 1, "within 14 days" bold, "the terms" a
  working link (hover shows https://example.org/terms), a bulleted list of two, a table of three rows
  with a header row. Does every piece survive the save? Does Word renumber or restyle anything?
- **37b, control 2:** "First paragraph in italics" with the italics, inside its paragraph, "end of
  the paragraph" after it.
- **37b, control 3:** the markup as text, as in 37a.
- **Developer tab, Properties** on control 1 in either Word: what the dialog shows for a control
  bound only through the tag (no XML mapping), and whether OK without changes rewrites anything.

Then say what you saw, and the saves are compared here against the originals, part by part.
