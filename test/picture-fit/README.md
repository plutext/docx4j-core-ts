# Word 365 fitting a node's picture into a mapped picture control (CR-005 section 10.2)

The portfolio's bridge rounds `core-ts-picture-fit` (2026-11-12, the VM's go 208) and `core-ts-picture-fit-2`
(2026-11-13, go 214), on bridge `22479f7`, in Word 365 16.0.20326.20318. Each file is a block-level picture control
mapped to `/data/pic`, its node holding a picture the control doesn't show yet, so Word fills it when it opens the file.
Each job opened the file and saved it as `.docx` and PDF. Every job ran twice (`-a`, `-b`), and with a control set; all
opened silently and each pair's extents and image agreed.

Copied unedited from `bridge/portfolio/<case>-a/`:

- `<case>.docx`: the input, built by this package (`build.mjs` records how; the bytes here are the ones the bridge ran);
- `<case>.word.docx`: Word's save;
- `<case>.result.json`: the bridge's record of the job.

`manifest.json` holds each file's SHA-1, and `test/picture-binding.test.mjs` checks them.

| Case | The drawing before | The node | Word's save |
|---|---|---|---|
| `ctpp-00-form-wide` | This package's form, 1270000 square, flag | 120 × 60 | 1270000 × 635000, flag cleared |
| `ctpp-01-form-tall` | The same | 60 × 120 | 635000 × 1270000, flag cleared |
| `ctpp-02-plain-wide` | A red 64 × 64 picture, 1270000 square | 120 × 60 | 1270000 × 635000 |
| `ctpp-03-plain-tall` | The same | 60 × 120 | 635000 × 1270000 |
| `ctpp-04-wide-box-tall` | A red 128 × 64 picture, 1270000 × 635000 | 60 × 120 | 317500 × 635000 |

Word fits the node's picture inside the drawing's box, keeping its aspect ratio, flag or not. It also writes an
`effectExtent` of 6350 on the shortened side, which this package doesn't.

## The editor's files (`pdu*`)

Five of the editor's files from its rounds `picture-dataurl` and `picture-placeholder` (2026-11-12, bridge `22479f7`, the
same Word), each with Word's save as `-word365.docx`. They are copied unedited from the editor's
`packages/editor-model/test/fixtures/pictures-dataurl/` at `513fa39`, whose README describes them. Each is a picture control
mapped to a node holding a 120 × 60 PNG as bare base64.

| File | The control holds | Word | This package's bind |
|---|---|---|---|
| `pdu4-base64-filled` | a 120 × 60 picture | the picture swapped, 1143000 × 571500 kept | the same |
| `pdu6-base64-placeholder-flag` | this package's form | 1270000 × 635000, flag cleared | the same (E1) |
| `pdu1-base64-empty` | placeholder text, flag | left, silently | the picture at its natural size, flag cleared (E2) |
| `pdu5-base64-text-noflag` | placeholder text, no flag | left, after a repair | the same (E2) |
| `pdu7-base64-empty-content` | nothing | left, silently | the same (E2) |

E2 is a recorded difference from Word (Jason, 2026-11-13): the bind fills a control Word leaves, since the bind produces the
file, and a picture it writes is then shown by every reader and refilled by Word on each open.
