# Word check 38: picture bindings (CR-005 section 9)

Built by `node test/fixtures/check38/build.mjs` from the repository root after `npm run build`
(docx4j-ts-editor ED-005 proposal 45, part 1). Two shapes bind a control to a node holding base64
image data:

- **(a)** Word's own picture content control, `w:sdtPr/w:picture` with a `w:dataBinding`;
- **(b)** `od:Handler=picture` on the tag of a rich text control, bound through the XPaths part's
  `od:xpath` entry and carrying no `w:dataBinding` (the engine's bind step fills it; Word has no
  binding to act on).

| File | What it is |
|---|---|
| `38a-picture-template.docx` | (a) as authored: the control shows a **red 200 x 100** placeholder; its node holds a **blue 300 x 150** image |
| `38a-picture-bound.docx` | (a) after the engine's bind: the control shows the blue image inside the placeholder's drawing (2.08 x 1.04 in, the placeholder's size) |
| `38b-handler-template.docx` | (b) as authored: a floating red placeholder (control 1), two text placeholders (2 and 3) and one in a 3000-twip table cell (4) |
| `38b-handler-bound.docx` | (b) after `applyBindings()`: 1 shows the blue image floating at the same place and size; 2 a wide green image (800 x 200) scaled to 3.125 in; 3 the same scaled to the text width; 4 the same fitted to the cell |

## What to do, in Word 2010 and in Word 15

For each file: open it, look, then save it beside the original as `38a-picture-template-word2010.docx`
and so on (Save As, the same format, the file's name plus `-word2010` or `-word15`).

What to look for:

- **38a template.** The question is whether Word fills a picture control from its binding on
  open: does the control show the **blue** image (Word filled it from the node) or the **red**
  placeholder (Word left it)? Either way, save.
- **38a template, then change the picture by hand, in Word 15 only:** on a fresh copy of the
  template, click the control's picture and use its picture icon or **Change Picture** to put in
  any image from disk, then save as `38a-picture-changed-word15.docx`. The question is what Word
  writes to the node: the new image's base64, or nothing. The engine reads that file afterwards and
  compares the node with the picture; until then `updateFromContentControls` leaves a picture
  control alone.
- **38a bound.** Blue on open in both Words; does the save keep the control, its binding and the
  image?
- **38b template.** Controls 1 to 4 unchanged (the red floating placeholder beside its text; the
  three placeholder texts). Does either Word complain about the tag, or alter a control on save?
- **38b bound.** Control 1: the blue image floating where the red one was, the same size (2.08 x
  1.04 in), the text wrapping around it on both sides. Control 2: a green image 3.125 in wide.
  Control 3: the green image as wide as the text area. Control 4: the green image 2.08 in wide,
  inside the cell. Do both Words open the file without a message (the floating picture's `wp:anchor`
  attributes are written as `true`/`false`, which Word itself writes as `1`/`0`)? Does the save keep
  every picture at its size?

Then say what you saw, and the saves are compared here against the originals, part by part.
