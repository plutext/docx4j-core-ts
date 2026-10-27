# CR-017: Integer attribute values Word 365 reads differently, reported on load and save

**Status:** Proposed 2026-10-26, at Jason's request ("open the read-side CR"): the read side of Word's integer ranges,
whose write side is CR-002 section 65. Objects 0.6.0, which brings its CR-015 and CR-016, was released 2026-10-26 and is the
floor since `b19d935` (2026-10-27), which did P1: `ValueKeptWarning`'s two new kinds, a decimal loading as Word reads it.
P2 and P3, and section 65.3's table move, not started. E1 to E5 open (section 5).

## 1. Why

docx4j CR-065 and CR-066 measured how Word 365 reads an integer attribute, over six rounds of probes Jason ran
(2026-10-26). Two things matter here:
- **A decimal or an odd form.** Word reads `720.5` as 720, cutting toward zero. It reads `""` and the forms with no
  digit (`.`, `+`, `-`, `+.`) as 0, and refuses a document holding anything else (`" 240 "`, `0x10`, `1e3`).
- **A value past the attribute's storage.** Each attribute is kept in 16 bits, signed or unsigned, and a value past
  that reads wrapped, with no error. So `w:ind w:left="7205001"` opens as -3959 twips.

The objects package follows both in its next release:
- its CR-015 reads a decimal as Word does and reports it through `onWordReading`, with two new kinds,
  `'value.truncated'` (a numeric value) and `'value.refused'`;
- its CR-016 adds `onOutOfRange(report)` on unmarshal and marshal: for each value Word would read differently, it
  reports the element, the attribute, the value, `wordReads`, `storage` and, on save, the path. It never rewrites a
  value.

This package has to take both up at the range raise, or it won't compile: `ValueKeptWarning.code` copies the reading's
kind. And it can then tell a caller, on the part and on the save, what Word would show.

## 2. Today (0.6.0, objects 0.5.0)

- **A decimal fails its part.** Over CR-065's probes, every part holding one throws
  `PartUnmarshalException`, for example "/word/document.xml: w:spacing/@w:before="120.5" (line 3, column 264):
  Argument [120.5] must be an integer". So `getBody()` throws for the main part. Nothing fails silently, and a save
  copies an unread part's bytes as they were.
- **A value past its storage is read as written, with no warning.** `w:left="8000000"` reads as `leftIndent` 400000,
  where Word shows 230.4 pt. A save writes it as held.
- **`ValueKeptWarning`** has the codes `'value.kept.exact'` and `'value.kept.invalid'` (CR-015 D8), and a value of
  `boolean | string | undefined`.

## 3. Proposal

- **P1, the readings.** `ValueKeptWarning.code` gains `'value.truncated'` and `'value.refused'`, and its value gains
  `number`. Each kind gets its message:
  - a truncated value: "read as 259, as Word 365 reads it" (today's form);
  - a refused one: "dropped: Word 365 refuses a document with this value".

  A decimal no longer fails its part: the part loads, holding Word's reading, with a warning per value.
- **P2, on load.** `XmlPart` passes `onOutOfRange` to the unmarshal. Each report becomes a `PartWarning` on the part's
  `warnings`, and so on `OpcPackage.warnings`. The warning has the code `'value.out_of_range'`, the part, the line and
  column where the DOM has them, and the fields agreed with docx4j-python (E1): `element` (`w:ind`), `attribute`
  (`w:left`), `value` as held, `wordReads` (null where no reading is known), `storage` (`int16`, `uint16`, `8 bits`,
  `5 bits`, `int32`, with a note for the special rules: the shared first-line and hanging field, the sign as the rule
  for a row's height and an exact line, the clamps on tab stops and page width) and `range`, the `[min, max]` kept as
  written, in the attribute's own unit. The message
  reads, for example: "/word/document.xml: w:ind/@w:left=7205001 (line 3) is past what Word 365 stores there (int16);
  Word reads it as -3959". The value stays as written.
- **P3, on save.** The marshal of each re-marshalled part passes `onOutOfRange` too, and `pkg.lastSave` gains
  `outOfRange`, as it has `normalised`: each value written past its storage, with its path. Nothing is refused or
  rewritten. A value set through a setter can't be one since CR-002 section 65; a direct edit of the tree can.
- **P4, the readers** stay as they are, on the value held (E2).

## 4. With other work

- **CR-002 section 65** guards the setters, the write side. Its `WORD_STORAGE` moves onto objects' `WORD_INTEGER_STORAGE`
  at the same raise, with the setters' own policy kept on top (section 65.3).
- **The same raise** takes objects' CR-014 (the highlight palette, with checks 61 and 62's tests) and the jsonix 3.9.0
  floor.
- **CR-016** (misplaced content) is separate: it is about which element sits where, not a value.

## 5. Decisions

| # | Question | Options | Recommendation |
|---|---|---|---|
| E1 | The warning's code and fields | (a) `'value.out_of_range'`, with `element`, `attribute`, `value`, `wordReads`, `storage` and `range` (P2), the code docx4j-python's CR-022 raises at its setters and its CR-020 validator will report under (python-cd, 2026-10-26: one code for the one condition, refused at a setter or found in a file; its fields the same, `word_reads` in Python's case; written as docx4j-python CR-020 section 4.7, which records this package's `wordReads` and numeric `value` as the TypeScript spelling); (b) a name of this package's own | (a): one code and one set of fields across the ports. `value.kept.*` is this package's naming (CR-015 D8; only its reasons, exact and invalid, are docx4j-python CR-018's), so this would be the first code the two share. |
| E2 | What the content API's readers give for a value past its storage | (a) the value held, the warning saying what Word shows (docx4j CR-066 D3, objects CR-016 section 5, docx4j-python); (b) Word's reading, as Office JS's readers give it (check 71: `leftIndent` read back 230.4 after Word wrapped 400000) | (a). The tree, the text model, the readers and every port agree on one number, and the warning carries Word's. Emulating the wrap would swap one wrong number for another and hide the evidence. The departure from Office JS is recorded. |
| E3 | Where the save's report goes | (a) `pkg.lastSave.outOfRange`, beside `normalised`; (b) part warnings | (a): it is about what that save wrote, as `normalised` is. |
| E4 | One warning per value, or per part with a count | (a) per value, as `ValueKeptWarning`; (b) per part | (a). Objects reports per value, and a document with many such values is rare: none of the 282 fixtures holds one (scanned 2026-10-26). |
| E5 | Tests | (a) CR-065's and CR-066's probe files (rounds 2, 5 and 6) and Word's saves copied unedited into `test/integer-attributes/`, with a SHA-1 manifest; each value held to Word's save through the warnings, as `runprops-word365/` holds CR-015's; (b) objects' own tests only | (a). It holds this package's warnings to Word's saves, not to the objects table alone. |

## 6. Phases

| Phase | Work | Status |
|---|---|---|
| 0 | This CR | done 2026-10-26 |
| 1 | At the objects range raise: P1, the decimals probes (round 2) loading, each truncated or refused value reported; `WORD_STORAGE` onto objects' table (CR-002 section 65.3) | P1 done at the raise, `b19d935` (2026-10-27; objects 0.6.0 released 2026-10-26); the table move not started |
| 2 | P2 and P3: the out-of-range warnings on load and the save's report; rounds 5 and 6's probes held to Word's saves | with phase 1, in the same release |
