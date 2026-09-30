# Known (to us) issues in Word

Behaviour of Microsoft Word we have measured and consider a defect or a gap, not merely a
convention to follow. Each entry says what was done, what Word wrote, the evidence, and what this
package does about it. Word's conventions that the engine follows on purpose (a revision's local
`w:date`, a joined paragraph's surviving id, and so on) are recorded in the change requests, not
here.

## 1. A list-definition edit is not tracked

**Found** 2026-09-28, Word 16.0.20326.20158 (Windows), `test/README.md` check 25, cases 25c/25d;
files in `test/fixtures/revisions/check25/`.

**What was done.** Track Changes on, a numbered list "One.", "Two.", "Three." (the second at the
third level). With only a caret in "One." (nothing selected), Home > Bullets; then, the caret at
the very start of "One.", Tab twice.

**What Word did.** It changed the list's own definition in `word/numbering.xml`, not the
paragraphs. Bullets made level 0 of the definition a bullet, so every level-0 item of the list,
"Three." too, turned to bullets. Tab at the start of the list's first item moved every level of
the definition 1080 twips to the right, so the whole list shifted. The paragraphs kept their
`w:numPr` unchanged (`25c-demote.docx` against `25b`).

**The issue.** None of it was tracked. The document visibly changed, with numbers turned to
bullets and the list indented, and yet:

- the Review pane listed nothing, and the saved file has no revision markup at all;
- Review > Reject > Reject All Changes changed nothing (`25d-demote-reject.docx` is `25c` again);
- a reviewer has no way to see that the list was changed, by whom, or to undo it.

Word's revisions attach to paragraph and run properties (`w:pPrChange`, `w:rPrChange`), and these
edits went into the definition, where Word records none. The same commands with the paragraphs
**selected** change the paragraphs instead (new `w:numId`, `w:ilvl`) and are tracked as
`w:pPrChange` (25a, 25e). Whether Microsoft regards the untracked case as a defect is not known to
us.

**What this package does.** The same as Word, knowingly:

- The content API's list writes (`attachToList`, `detachFromList`, `startNewList`,
  `ListItem.level`) set the paragraph's `w:numPr`, so under tracking they are recorded as
  `w:pPrChange` and can be rejected.
- Office JS's list-level methods, `List.setLevelNumbering`, `setLevelBullet`, `setLevelIndents` and
  `setLevelAlignment`, edit the existing definition in `word/numbering.xml` in place and are not
  tracked, as Word's own definition edits are not. (What those methods do under tracking in Word
  itself, through Office JS, has not been measured.) `pkg.numbering`'s `newList` and `restart` add
  definitions and list instances rather than editing one.
- The editor (`../docx4j-ts-editor`) never edits an existing definition from its list commands. Its
  one path that does, a raw edit of the numbering part, is untracked, as in Word (its ED-005
  section 12.7 item 6).

A caller that edits definitions under tracking should expect no revision and no way to reject the
edit, in this package and in Word alike.

## 2. Deleting a table column and merging cells are not tracked; an added column cannot be rejected away

**Found** 2026-09-28, Word 16.0.20326.20158 (Windows), through Office JS, `test/README.md` check 27
(part A, cases `delete-column`, `merge-vertical`, `add-column`).

With Track Changes on, `table.deleteColumns(1, 1)` removed the column outright and
`table.mergeCells(0, 0, 1, 0)` merged two cells vertically (their text moved into one cell), and
neither left any revision: nothing was listed, and there was nothing to reject. `table.addColumns`
was tracked, but as a formatting change of the table plus a `w:ins` of the new cells' text - Word
wrote no `w:cellIns` - so rejecting everything left the new column in place, empty. None of the three
can be seen or undone by a reviewer as a table change. (Measured through Office JS only; whether
Word's own Table Tools commands behave the same, or warn that the action will not be tracked, has
not been checked.)

## 3. Office JS cannot list a cell merge, and failed on a list change at the end of the body

**Found** 2026-09-28, the same run (check 27). With a `w:cellMerge` in the document,
`body.getTrackedChanges()` threw `GeneralException` (as it does over a move, checks 16 and 17). A list
made and moved a level down under tracking, on the body's last paragraph, was listed as one
`Formatted` change, and then `accept()`, `reject()`, `acceptAll()` and `rejectAll()` each threw
`GeneralException`.

## 4. `insertOoxml` does not carry table and numbering revisions faithfully

**Found** 2026-09-28, the same run (check 27, part B). Put in through `insertOoxml` with tracking off,
a `w:tblPrChange` and a `w:tblGridChange` were dropped (the table came in with its current
properties and no record), and so was a `w:numberingChange` (the list renumbered, no record). A
`w:trPrChange` or `w:tcPrChange` came in with records Word made up - a `w:tblPrChange`, a
`w:tblGridChange` over an odd grid, a `w:tblPrExChange`, cell widths recorded as `auto` - and
rejecting them produced a mangled table (five grid columns, `w:gridAfter`, spanning cells). An
inserted or deleted cell (`w:cellIns`, `w:cellDel`), a mark's `w:rPrChange` and a section break's
`w:sectPrChange` came in intact. A test that needs Word's resolution of table-property revisions has
to open a file (check 18), not insert one.

## 5. A resolved comment is not kept in a compatibility-mode document

**Found** 2026-09-29, current Word (Microsoft 365), by Jason, relayed by the editor (its ED-005
section 12.23 item 7, where the saves are kept as fixtures; not measured in this repository).

A comment thread marked resolved (`w15:done="1"` in `word/commentsExtended.xml`) showed as resolved
in Word, but saving the document - a compatibility-mode document, as Word 2013 writes it - wrote
`w15:done="0"`. It did so for the editor's file and for Word 2013's own serialisation of the same
comments, which declares `mc:Ignorable="w14 w15 wp14"` on the comment parts, so the declarations are
not the cause: current Word does not keep a resolved state it did not set itself in a
compatibility-mode document, whoever wrote it. Word 15 honours `w15:done` in the same files.

**What this package does.** Nothing different: it writes and reads `w15:done` as the format says.
A caller that needs a thread to stay resolved through a save in current Word should not rely on a
compatibility-mode document.

## 6. Rejecting text typed into an existing content control leaves its placeholder in a `w:ins`

**Found** 2026-09-30, Word 16.0.20326.20158 (Windows), through Office JS, `test/README.md` check 30
(gesture `existing-control-typed`; the JSON is `test/fixtures/revisions/check30/result.json`).

With tracking off, an empty rich text control was put at the end of "Kept."; with tracking on as
Author B, text was typed into it (`insertText`, "Replace"), which Word tracked as a `w:ins` inside
`w:sdtContent` with no marker on the control - the form that says the control was there before.
Reject All Changes then, rightly, kept the control and put its placeholder back ("Click or tap here
to enter text.", `w:showingPlcHdr`), but wrote the placeholder run **inside a `w:ins` of Author B's**,
so that `getTrackedChanges()` still listed one `Added` change with empty text, and the document was
not clean after rejecting everything. Accept All had no such leftover.

**What this package does.** Rejecting the insertion removes the `w:ins` and leaves the control's
content empty, with no placeholder run and nothing listed (`tracking.test.mjs`, "a control that was
there before and is typed into under tracking stays, empty"); an empty control is Word's own state
for one cleared by hand. Nothing is written to imitate the leftover.

## 7. Office JS leaves another author's inserted row out of the listing

**Found** 2026-10-01, Word 16.0.20326.20158 (Windows), through Office JS, `test/README.md` check 31
(engine case `engine-two-authors`; the JSON is `test/fixtures/revisions/check31/result.json`).

A three-row table inserted whole, its first and third rows Author A's and its second Author B's
(`w:trPr/w:ins`, every mark and run inserted, the same form for all three), between two paragraphs
inserted by Author A. `body.getTrackedChanges()` listed Author A's two halves, `"before the table\rrow
one\r\n"` and `"row three\r\n(end of insertion)"`, and **nothing for Author B's row**: the row is in
the markup, `rejectAll()` removed it and `acceptAll()` kept it, but no `TrackedChange` was given for
it. A reviewer working from the listing would not see B's row.

**What this package does.** Lists it: `[A "row one\r\n", B "row two\r\n", A "row three\r\n"]` for that
markup (`tracking.test.mjs`, "an inserted table groups with the rows and the inserted text touching
it"). Whether Word's own Reviewing Pane shows the row has not been checked.

## 8. `insertParagraph('Before')` on a table under tracking leaves the new mark plain

**Found** 2026-10-01, the same run (check 31, gesture `text-then-rows`).

With tracking on, `table.insertParagraph('typed before the table', 'Before')` wrote the paragraph's
run in a `w:ins` but no `w:ins` on its mark (`<w:p><w:ins>...</w:ins></w:p>`, no `w:pPr`), so the new
paragraph itself is not an insertion: rejecting everything left an empty paragraph between "Kept."
and the table, and the listing kept the text apart from the rows added after it (a plain mark is a
break). `insertParagraph('After')` on the table marked the mark inserted, as typing Enter does.

**What this package does.** `Body.insertParagraph` and `Paragraph.insertParagraph` mark the new mark
inserted wherever the paragraph goes (CR-002 section 13), so rejecting leaves nothing.
