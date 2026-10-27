# CR-017: Integer attribute values Word 365 reads differently, reported on load and save

**Status:** Proposed 2026-10-26, at Jason's request ("open the read-side CR"): the read side of Word's integer ranges,
whose write side is CR-002 section 65. Objects 0.6.0, which brings its CR-015 and CR-016, was released 2026-10-26 and is the
floor since `b19d935` (2026-10-27), which did P1: `ValueKeptWarning`'s two new kinds, a decimal loading as Word reads it.
**P2 and P3 implemented 2026-10-27** (Jason: "start CR-002 section 65.3 then CR-017 P2 and P3"), with section 65.3's table
move, E1 to E5 taken as recommended (section 7); unreleased.

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
| 1 | At the objects range raise: P1, the decimals probes (round 2) loading, each truncated or refused value reported; `WORD_STORAGE` onto objects' table (CR-002 section 65.3) | P1 done at the raise, `b19d935` (2026-10-27; objects 0.6.0 released 2026-10-26); the table move done 2026-10-27 (section 7, CR-002 section 65.7); round 2's probes held to Word's saves (section 7) |
| 2 | P2 and P3: the out-of-range warnings on load and the save's report; rounds 5 and 6's probes held to Word's saves | done 2026-10-27 (section 7) |

## 7. Implemented (2026-10-27)

Jason: "start CR-002 section 65.3 then CR-017 P2 and P3". E1 to E5 were open; the instruction was taken as going with the
recommendations, and each is implemented as recommended: E1 (a) the shared code and fields, E2 (a) the readers on the value
held, E3 (a) `pkg.lastSave.outOfRange`, E4 (a) one warning per value, E5 (a) the probes and Word's saves as the tests.

**P2, on load.** `XmlPart.unmarshal` passes `onOutOfRange` beside `onWordReading`, and each report becomes a
`ValueOutOfRangeWarning` on the part's `warnings` (so on `OpcPackage.warnings`): `code` `'value.out_of_range'`, `partName`,
`line` and `column` where the DOM has them, `element` and `attribute` as the part writes them (`w:ind`, `w:left`), `value`
as held, `wordReads` (null for an unmeasured attribute, and for a column width Word computes), `readAs` where the shared
field's sign moved the value to the other attribute (`w:hanging` for a `w:firstLine` of 32768), `rule` where a line's or a
height's sign carries one, `storage` (the objects package's wording) and `range`, the `[min, max]` Word keeps as written,
in the attribute's unit. The message: `/word/document.xml: w:ind/@w:left=7205001 (line 3, column 1245) is past what Word
365 stores there (int16, -32768 to 32767); Word reads it as -3959`; for the shared field `Word reads it as 32768 in
w:hanging`; for a rule `Word reads it as atLeast 1`; for an unmeasured attribute `is outside -32768 to 65535, where Word
365 changes every measured value of its kind (w:gridCol/@w:w is unmeasured); what Word reads is not known`.

**P3, on save.** `PartBytesOptions` gains `onOutOfRange`; `XmlPart.marshalToNode` passes it to the marshal, and `Save`'s
`Saver` collects each report as the same warning, with `path` from the part's root in place of the line, onto
`pkg.lastSave.outOfRange` beside `normalised`. Nothing is refused or rewritten, under `normaliseOnSave` or not. One thing
the facade's contract made necessary: the marshal reports a value as its attribute is written, before the element's other
attributes are, and the warning's range depends on the `lineRule` or `hRule` beside a line or a height; so `marshalToNode`
holds the reports until the element is whole, then builds the warnings.

**The range (E1's `range`, and section 65.3's bounds), one source.** `src/model/wordIntegers.mts`: `keptRange(typeName,
property, rule?)` derives from the objects table's entry (`bits`, `signed`, `pair`, `signRule`, `heightFlag`, `min`, `max`,
`negative`) the range Word keeps as written: a signed field `-2^(bits-1)` to `2^(bits-1) - 1`, unsigned `0` to `2^bits - 1`;
`firstLine` `0` to `32767` and `hanging` `0` to `32768`; a line or a row height under `exact` `1` to `32768`, under `atLeast`
`0` to `32767`, a line under `auto` the plain signed field, a row height under `hRule="auto"` null (ignored); a frame's height
`0` to `32767`; then the clamps, and a column's width or space `0` and up (Word replaces a negative). 65535 where Word writes
it back as -1 is in the range; -1 there, and a column space's -1, Word keeps too, its own spellings, and the objects
package's `kept` says so; the test states both. `WORD_STORAGE`'s ten rows are now `row(attribute, unit, typeName,
property, policy)`: `keptRange`'s bounds, narrowed only where the policy says why (`spacingBefore` and `spacingAfter` stop
at 65534, "Word's save writes 65535 back as -1"; a width's `min` 0), and `reads` is `wordIntegerReading`'s answer in the
setters' words (`a hanging indent of 32768`, `atLeast 1`). A row whose entry the table lacks fails at load, so a table change
announces itself. `measurements.test.mjs`'s pins hold, one message changed (`spaceBefore` 3276.75 now says why).

**The report names no entry.** The facade's `OutOfRange` carries the element, the attribute and the reading, not the
table's `TYPE_NAME` and property, and an element's class is not recoverable from its name (`CTBorder` is `w:top`, `w:bdr`,
...; `TblWidth` is `w:tblW`, `w:tcW`, ...). So `rangeOfReading` finds the range back from the reading's `storage` string,
with the attribute's local name for the shared field and the rule written beside a line or a height, from a map of every
string the table gives (21 keys, each one range; a key two ranges shared would be dropped, none is). It holds because the
strings are the table's own, generated from the same fields; `integer-attributes.test.mjs` asserts the lookup agrees with
`keptRange` for every entry and rule. An entry reference on the report would make the lookup unnecessary: asked of the
objects package as a small addition to its CR-016, not blocking.

**Tests** (`test/integer-attributes.test.mjs`, 9 tests; 1,626 pass): the corpus `test/integer-attributes/` (114 files,
1.9 MB: CR-065's round 2 to 4 files whole, CR-066's round 5 groups, round 6 groups and the settings singles, with Word's
saves; the family and single files Word was never given left on the share), its SHA-1 manifest; `keptRange` against
`wordIntegerReading` at every boundary and 40 samples for each of the table's 51 entries under every rule, and
`rangeOfReading` agreeing; `WORD_STORAGE`'s rows within Word's range and narrowed only with a reason; rounds 5 and 6 as the
objects package's `word365-ranges.mjs` holds them, through this package's loads and saves: 660 integer cases, 614 compared
to Word's save, 16 left out by Word as a 0, 30 stated (`w:gridCol` recomputed, `w:wBefore` and `w:wAfter` dropped), a
warning on load and on save for exactly the 432 cases Word's save changed (the objects package's own count), with
`wordReads`, `readAs` and `rule` what Word wrote, every value written as held, and Word's own saves reporting nothing; group
E's no-digit forms; the message and fields of each kind; a value set by code reported on save with its path, under the
normaliser or not, a clean package and an untouched one reporting `[]`; round 2's 64 cases read as Word's saves have them,
50 equal, the six known differences stated strict both ways (a value truncated to the default Word leaves out; the grid
Word recomputes), the eight refused forms dropped and reported, the all-cases file loading every part (55 truncated, 5 read
as 0, 8 refused), and Word's saves needing no reading of Word's; and every fixture's main part holding nothing out of range
(E4). `test/nodenext/consumer.mts` uses `ValueOutOfRangeWarning` and `keptRange` by the package's name.

**Not done.** P4 is nothing to do (E2). Check 71b (CR-002 section 65.5 D2) is still Jason's. The editor and docx4j-python are to
be told the field names as implemented (`readAs` and `rule` are this package's additions beside the agreed six).

**docx4j's side (hand-off from docx4j-12, 2026-10-27).** docx4j CR-066 D6 is built in its inspector (docx4j `338beec18`, CR-066
section 10): a `value.out_of_range` finding per attribute occurrence, over every WordprocessingML part, unselected `mc` branches and
an embedded docx one level deep, with `attribute`, `type`, `value`, `storage`, `range` and `word_reads`. Its widths are one resource,
`docx4j-core/src/main/resources/org/docx4j/model/text/word-integer-storage.tsv` (101 rows, CR-066 sections 3.1 and 3.4 row for row,
with a context column: `type=auto|nil`, `lineRule=exact`, `hRule=exact`), which docx4j's inspector and docx4j-python's oracle read;
the objects package's `WORD_INTEGER_STORAGE` is the third copy, and the one this package reads. Where the shapes differ: docx4j's
`word_reads` is a string in the attribute's own terms (`hanging 3959`, `atLeast 1`, `dropped`, `computed`, `changed`), this package's
`wordReads` a number with `readAs` and `rule` beside it (and `describeReading` gives the string for the message); docx4j does not check
`w:fitText`, a row height under `auto`, a table width under `auto` or `nil`, or `w:legacy`, and merges a run property's finding into the
adjacent run's as its other run findings merge. Nothing here changes for it: the code and the six fields agree. If this package ever
reads a table rather than the objects package's, it is to be that TSV, row for row, not a fourth encoding of the sections.
