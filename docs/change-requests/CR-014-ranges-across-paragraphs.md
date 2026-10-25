# CR-014: Ranges across paragraphs, as Office JS's `Word.Range` is

**Status:** Proposed 2026-10-21, at Jason's request. He asked why a `Range` here is one paragraph's where
Word's is not, and whether we should align. To the answer, open a CR with the editor's review and a Word check
first, and stop the silent partial results meanwhile, he said "yes please". The stopgap went in with CR-012
phase 4: a multi-paragraph block control's range reads as Office JS's and refuses every write (CR-012 section
17.12). Decisions D1 to D4 (section 5) were decided 2026-10-22 as recommended
(Jason: "CR-014 decisions D1 to D4: recommendations accepted."): one `Range` with an end paragraph;
`Comment.getRange()` returns one `Range`, the per-paragraph array kept as `getRanges()`; whole tables read and
deleted, a range starting or ending inside a table read and its writes refused unless check 52 shows a simple rule;
and the editor's answer stands, so phase 1 needn't wait on it. Check 52 (phase 0) was written 2026-10-22 (Jason:
"implement CR-014 phase 0"): 85 cases, a region each, in one file and one Script Lab run (section 6.1), and ten
more for CR-002 section 55's D2 (Word's table rule). Jason ran
it the same day, in All Markup, and saved Word's file after the run and after Accept All (section 6.2): every question of
section 2 and every addition of the editor's has its answer. **Phase 0 is done.** **Phase 1, the reads, implemented
2026-10-22** (Jason: "implement CR-014 phase 1"; section 9): one `Range` with two ends, read as Word's, the views
returning one, `SpanningRange` gone; held to check 52 by running its own snippet through the Word shim. Its writes
refuse until phase 2. Its choices where Word was unmeasured were then measured by check 53 (Jason ran it the same day,
in both views) and changed to Word's answers where they differed (section 9.7). Committed `ef57729`. **Phase 2, the writes,
implemented 2026-10-22** (Jason: "implement CR-014 phase 2"; section 10): `delete`, `insertText` (`Replace`, at the
ends), `font`, `insertParagraph` and `replaceText` across paragraphs, untracked and tracked, as Word writes them;
held to check 52 by running its snippet's `run` mode through the Word shim against Word's run, its save after the
run and its save after Accept All; a range starting or ending inside a table refused as Word refuses it (D3).
Committed `53db5bc`. **Phase 3 begun 2026-10-22** (Jason: "do it"; section 11): check 54 written first, for phase 2's
unmeasured choices (section 10.6) and content controls across paragraphs, one Word run; comments across paragraphs
and `^p` in `search` implemented meanwhile, held to check 52. Check 54 run the same day (section 12): phase 2's
choices changed to Word's answers where it differed, content controls across paragraphs done on its answers (whole
paragraphs only, any other range Word's `GeneralException`), and a tracked deletion over a whole control written in
Word's markup and resolved as Word does. A tracked `Replace` within one paragraph puts its insertion first, as
Word's (Jason, 2026-10-23; section 12.1, item 7), and a deletion takes a comment's reference with it, its comment
going untracked and on Accept All (section 12.6). Left: three rules of Word's Accept All, docx4j's to take first
(section 12.3). Committed `87bbb38`. Section 13 (2026-10-23, Jason: "the shim refuses getOoxml - fix this please"):
the shim's `getOoxml` and `toApiScript` across paragraphs give the range's blocks, and under tracking no longer mark a
copy's trimmed text deleted. Section 14 (2026-10-23, docx4j-python's finding): a range holding one table whole, nothing
either side, deleted the block after it too; fixed, committed `222cc20`. Section 12.3's three Accept All rules (cases 9, 13
and 19) followed since 2026-10-23, as docx4j's CR-043 section 17 (`29dbf5964`) adopted them, committed `222cc20`; and
as docx4j revised them the same day from a text-left probe (`e55a83b8e`), with Word's Reject All of check 54, committed `4837d04`. Phases 0 to 3 and sections 12 to 14 **released in 0.6.0** (2026-10-25).

## 1. Why

**A `Range` here is one paragraph's.** CR-002 (2026-09-09) defined it as "a run of text within one paragraph, or
a whole paragraph". The reason was scope: phase B's edits all work on one paragraph's text segments and their
offsets, and a search never crosses a paragraph mark. It was not a fidelity decision, and nothing was measured.
CR-002 section 7 lists it as a departure only for `ContentControl.getRange`.

**Office JS's `Word.Range` spans what it covers**: paragraphs, tables and content controls. Check 51 (2026-10-21)
measured it on a block content control of three paragraphs. Its 'Whole' range reads `CCP1\rCCP2\rCCP3\r` and its
'Content' range `CCP1\rCCP2\rCCP3`, where this package gave `CCP1\r` and `CCP1`.

**Where a one-paragraph range stands in for a longer one today:**
- `ContentControl.getRange()` on a block control: the first paragraph. Since CR-012 section 17.12 it is a
  `SpanningRange` when the control holds more than one paragraph: it reads all of them and refuses every write.
- `Comment.getRange()`: a `Range[]`, one per paragraph the comment touches (CR-002 phase G, an extension), where
  Office JS returns one `Range`.
- `TrackedChange.getRange()` for a change whose pieces span paragraphs: the part in its first paragraph
  (CR-002 section 29). Its `text` is the whole change's, so the two disagree.
- `Range.paragraphs` always has one item. There is no `Body.getRange`, `Table.getRange` or `Range.expandTo`.

**What it costs.**
- The results are silently partial. Before the stopgap, `cc.getRange().text` read one paragraph of three, and
  `cc.getRange().delete()` or a `font` write changed one, with no error. An add-in, or the editor's agent, trusts
  the result.
- Code written against Office JS that takes a range across paragraphs, from a control, a comment, a change, a
  `^p` search or `expandTo`, can't run here.

## 2. What Office JS does

**Measured:**
- **Check 51:** a block control's ranges, as above. An inline control's range is its text, with no mark.
- **Checks 48, 51, 30 and 31:** the text a range across paragraphs would read is `body.text`'s form. Each
  paragraph's text is followed by its mark `\r`. In Word's default view a table row's cells are a `\t` apart and
  the row ends `\r\n`.

**Not measured; check 52 (section 6) asks:**
1. A range from one paragraph's middle to another's: its `text`, and what `delete()` leaves. Are the paragraphs
   joined, and whose properties does the joined paragraph keep? Under change tracking, are the marks deleted?
2. `insertText(…, 'Replace')` over it: is the result one paragraph, and in whose formatting?
3. A `font` write over it, and `insertParagraph('Before'|'After')`: before its first paragraph or after its last?
4. A range that holds a whole table, and one that starts or ends inside a cell: `text`, `delete()`, a
   replacement. Is a partial table refused?
5. `expandTo` between two paragraphs, and `search` with `^p` (Word's paragraph mark), whose hits span paragraphs.
6. `insertComment` over a range across paragraphs: where Word puts the markers. This package's `Comment` reads
   such a comment already.
7. `getRange('Start')`, `('End')` and `('Content')` of such a range.

## 3. Proposal

- **One `Range` class, its ends in two paragraphs.** `paragraph` and `start` stay the start, as now.
  - A new `endParagraph`, defaulting to `paragraph`, holds `end`, and both are in one story.
  - A range within one paragraph is exactly today's: nothing changes for a caller that never makes another kind,
    the editor's per-paragraph spans included.
  - `SpanningRange` goes, its reads becoming `Range`'s.
- **Reads:**
  - `text` in `body.text`'s form between the two ends, with a 'Whole' range's last mark;
  - `paragraphs` all of them;
  - `font` the first run's;
  - `search` and `replaceText` within the span, a paragraph at a time unless the query holds `^p`;
  - `getTrackedChanges` and `getComments` over every paragraph;
  - `getRange` and `expandTo`.
- **Writes**, per check 52's answers:
  - `delete` and `Replace` take the tail of the first paragraph, the whole of every block between, and the head
    of the last, then join the first and last paragraphs. Under tracking each removed mark is a deleted mark,
    with the machinery `Paragraph.delete` and Accept All already use (CR-002 phase F);
  - `font` writes each paragraph's span;
  - `insertText` at the ends;
  - `insertParagraph` before the first or after the last paragraph;
  - `insertComment` puts its markers in the first and last paragraphs.
- **Views that return one:** `ContentControl.getRange()`; `Comment.getRange()` (D2); `TrackedChange.getRange()`;
  a new `Body.getRange()` and `Table.getRange()`; `Range.expandTo(range)`; a `^p` search hit.
- **Tables in a range:** whatever check 52 shows. A range holding a whole table takes it whole. A partial one is
  refused with `NotSupportedError` until a use asks for more (D3).
- **Out of scope:** ranges across stories (a body and a note), and across a text box's boundary. Ranges through
  a block-level `mc:AlternateContent` follow CR-013's D4 (read-only).

## 4. Phases and sizes

| Phase | What | Size |
|---|---|---|
| 0 | Check 52 (section 6), and the editor's review (D4) | ½ day, and Jason's run |
| 1 | The reads: `Range` with two ends, `text`, `paragraphs`, `search`, `getRange`, `expandTo`; the views returning such ranges; `SpanningRange` retired | 1½ to 2 days |
| 2 | The writes across paragraphs: `delete`, `Replace`, `font`, `insertText`, `insertParagraph`, untracked and tracked; tables per D3 | 2 to 3 days |
| 3 | Comments and content controls on such a range; `^p` in `search` | 1 day |

## 5. Decisions

All four were decided 2026-10-22 as recommended (Jason: "CR-014 decisions D1 to D4: recommendations accepted.").

1. **D1, the shape** (decided 2026-10-22, as recommended). Recommended: one `Range` class with an end paragraph, a single-paragraph range unchanged,
   rather than a second class beside `Range`. Office JS has one `Range`, and two classes would make every caller
   ask which it has.
2. **D2, `Comment.getRange()`** (decided 2026-10-22, as recommended). Office JS returns one `Range`; this package has returned `Range[]` since CR-002
   phase G. Recommended: return one `Range` as Office JS does, and keep the per-paragraph array as
   `getRanges()`. It breaks a caller that indexes the array, which on 0.x is allowed.
3. **D3, tables** (decided 2026-10-22, as recommended). Recommended: a range holding whole tables reads and deletes them. One that starts or ends
   inside a table reads in `body.text`'s form, and its writes throw `NotSupportedError` unless check 52 shows a
   rule simple enough to follow.
4. **D4, the editor** (decided 2026-10-22, as recommended: its answer below stands). Its spans are per paragraph, so this is additive for it, but its agent's verbs (`delete`,
   `replace`, `comment` over a span) could take ranges across paragraphs. Ask it, as with CR-012, before phase 1.
   - **Its answer (2026-10-21, ED-008 section 8.8):** additive. Its agent's span stays inside one paragraph's text
     (its `docs/agent.md` section 2.3); several whole blocks go through its own verbs. The first uses it would
     take, both after this CR's phase 3:
     - `comment` over an address range, a reviewer's common case, which its margin anchors already;
     - `delete` of a span crossing a paragraph's end, the mark deleted as its tracker does for Backspace.

     A span crossing paragraphs would be a second shape in its contract, not a change to the first. **Phase 1
     needn't wait on it.**
   - **On D2:** nothing of the editor's indexes `Comment.getRange()`'s array, and one `Range` as Office JS
     has is fine.

## 6. Check 52

A file with paragraphs of distinct formatting (a heading, a body paragraph, a list item), a table of two rows,
and a three-paragraph block control. A Script Lab snippet:
- makes ranges with `expandTo` (one paragraph's middle to another's; one around the table; one from inside a cell
  to the paragraph after the table) and reads their `text`;
- applies `delete`, `insertText('X', 'Replace')`, `font.bold`, `insertParagraph` and `insertComment` to each, on
  a fresh copy each time (tracking off, then on);
- reports each result's paragraphs (text, style, list) and the document's `body.text`;
- searches `^p` and reads the hits.

Word's saves are kept beside the JSON, as checks 30 and 31's were.

The editor's additions (2026-10-21, its D4 answer):
1. The same deletion with its ends swapped (heading first, body first) and one starting in the list item, in
   both tracking modes: which paragraph's properties and numbering the joined paragraph keeps, and, tracked, the
   `w:pPrChange` the deleted mark leaves. Its ED-005 section 12.44 carry rule is its reading of Word 15 for a mark
   deleted alone; across paragraphs it is unmeasured.
2. A range crossing a hidden run, and one crossing a footnote reference: is the note's text in `text`, and what
   does `delete` do to the note?
3. `insertComment` on a 'Whole' range ending at a paragraph mark: where the end marker goes.
4. `^p` against a cell's end mark and a row's end.
5. `expandTo` from inside a cell to the paragraph after the table, under tracking: which rows are marked deleted,
   and whether the table survives. The editor's Accept All removes a table whose every row is deleted (its ED-005
   section 12.109).

### 6.1 As written (2026-10-22)

`test/README.md` check 52 is the check: its steps, what the JSON holds, and its run once it is back. The file is
`test/fixtures/check52/build.mjs`'s (`check52Bytes`; the `.docx` git-ignored, Word reading the share's copy), the
snippet `test/fixtures/check52/check52-script-lab.js` and `.html`.

- **A region per case, not a fresh copy per case.** Office JS can't reopen a file, and 85 copies would be 85 runs.
  Each case is a region, `BEGIN<n> <kind> <operation> <tracking> <range>`, the kind's content, `END<n>`, so the
  operations of one case never touch another's content, and one run measures them all. Cases 1 to 85 are this
  CR's. Cases 86 to 95 were added the same day for CR-002 section 55's D2 (what Office JS's edits leave in a table
  cell), since check 52 had not run: one run for both. The document says what each
  case is, and the snippet reads it from there: the file is the only list of cases.
- **The kinds of range** (thirteen): heading to body and the ends swapped (HB, BH; the editor's addition 1), from and
  to a list item (LB, BL), a whole paragraph between (HMB), the whole table between (TBL), ending inside the table
  (INTO), starting inside it (CELL; addition 5), across a hidden run (HID) and a footnote reference (FN; addition 2),
  a block control's content (CC), a paragraph's whole range (W1) and a range ending at a mark (W2; addition 3).
- **The operations:** `delete` and `Replace` (section 2, items 1, 2 and 4), `font.bold` and `insertParagraph`
  (item 3), `insertText` at Start and End, `insertComment` (item 6), each in both tracking modes where tracking can
  change the answer. Each kind's reads come first (items 5 and 7: `text`, `font`, `getRange` of each location,
  `paragraphs`, `^p` within it, `expandTo` with its ends swapped). Two regions are read only: `^p` against
  paragraphs, a cell's end and a row's end (addition 4), and a comment and tracked changes that span paragraphs,
  with their `getRange()` (section 1's `Comment.getRange()` and `TrackedChange.getRange()`).
- **Word's two saves** come back with the JSON, after the run and after Accept All: the first shows the markup Word
  writes (the deleted marks, `w:pPrChange`, where `insertComment` puts its markers), the second what accepting it
  leaves (whether CELL's table survives, addition 5).
- **Phase 1 runs the same snippet here.** `CHECK52.runCheck(run, mode)` takes `Word.run` as an argument and nothing
  of Script Lab, so phase 1's test can run it through the Word shim (`Word.run(pkg, …)`) on `check52Bytes` and
  compare the result with Word's JSON case for case, once the shim has `expandTo`. `test/ranges.test.mjs` holds the
  file and the snippet to each other meanwhile.

### 6.2 The run (2026-10-22)

Jason ran check 52 on Word 365 (16.0.20326.20232, PC) the same day, in All Markup, then saved Word's file, accepted all
and saved again. `test/README.md` check 52 has the run case by case, and Word's markup. The saves are `test/check52/`,
outside `fixtures/` as `test/mc-fun/` is, so that the weekly `RevisionsHarness` cuts no golden for them (Jason).
`test/fixtures/check52/52-ranges-word365-run.json` is the run and `52-ranges-word365-accepted.json` the report after
Accept All. The text reads under tracking include deleted text, as All Markup shows it.

**What it settles for phase 1 (the reads):**
- `text` is `body.text`'s form between the ends, as section 3 proposed: a mark `\r`, a table's rows ended `\r\n`
  and cells a `\t` apart, a note reference `\u0002`, hidden text left out.
- `getRange('Whole')` is the range itself, not widened to its paragraphs. `'Content'` is the same without a last mark
  where the range ends with one. `'Start'`, `'End'` and `'After'` are empty, `'After'` in the next paragraph where the
  range ends with a mark.
- `paragraphs` is every paragraph the range touches, a table's cells' too. `expandTo` doesn't care which end is
  first.
- **`font` reads `null` where the paragraphs differ** (`''` for the name, `Mixed` for underline), not the first
  run's as section 3 proposed. Phase 1 follows Word.
- **A tracked change across paragraphs is one change** (case 1: `Deleted`, `HEAD1B\rBODY1A `), and its `getRange()`
  spans it. This package already lists it as one; its `getRange()` gives the first paragraph's part (section 1), and
  phase 1 widens it. **A comment's `getRange()` spans its paragraphs** (`BODY1B\rLIST1A`), as D2 has it.
- **`^p`** matches a paragraph's mark outside a table, and never a cell's or row's end. It doesn't match across the
  boundary between a paragraph and a table either, so phase 3's rule is that simple.

**What it settles for phase 2 (the writes):**
- **`delete()` joins the paragraphs in the properties of the paragraph the range ends in**, whichever kind it is
  (heading, body, list item), each run keeping its own formatting. Blocks between go whole, a table too.
- **A last mark in the range goes with it**: a paragraph's whole range deletes the paragraph, and a range ending with a
  mark joins its head to the next paragraph, in that paragraph's properties.
- **`Replace` is `delete()` and then the text, but keeps a last mark.**
- **A range starting or ending inside a table: Word refuses `delete()` and `Replace`** (`GeneralException`, nothing
  changed). D3's refusal is Word's own. Word does apply `font`, `insertParagraph` and `insertComment` to such a range,
  so phase 2 refuses only the two Word refuses.
- **`font` writes every run in the range**, a table's cells too.
- **`insertParagraph` splits the paragraph** at the range's start (`'Before'`) or end (`'After'`), the new paragraph
  in that paragraph's properties. Where the range starts at a paragraph's start or ends with its mark, nothing is
  split. This is not section 3's "before the first or after the last paragraph". Today's `Range.insertParagraph`, on a
  one-paragraph range, goes before or after the whole paragraph; whether Word splits a one-paragraph range the same
  way isn't measured, and phase 2 leaves it as it was (section 10.6, item 1).
- **`insertText` at `'Start'` and `'End'`** goes at the first and last character, before a last mark.
- **`insertComment`**'s range is the range's text, a last mark not included.
- **A footnote reference in a deleted range takes its note.**
- **Under tracking**, a deletion is one `Deleted` change, and the paragraphs stay until accepted. Every mark in the
  range is marked deleted (`w:pPr/w:rPr/w:del`), and the last paragraph's mark stays. **No `w:pPrChange` is written**:
  the editor's addition 1, whose carry rule (its ED-005 section 12.44) is for a mark deleted alone. A table in the
  range has every row and every cell's mark deleted.
- **Accept All gives the untracked answer**: the paragraph the range ends in keeps its properties, and the table goes.
  This package's `acceptAllRevisions()` on Word's save after the run agrees with Word's save after Accept All in 94 of
  the 95 regions, read through the content API (each paragraph's text, style, alignment, list and table nesting). The
  95th, case 26, differs only by Word's filler: Word's emptied control holds five spaces, this package's nothing.
  `test/ranges.test.mjs` holds it so.
  Phase 2's tracked deletions write Word's markup, and the existing resolution already joins as Word does.
- **Tracked `Replace`** writes the insertion before the deletion, the marks deleted as `delete()` deletes them,
  except a last mark. Office JS's `getTrackedChanges()` listed only the insertion, but the deletion is in the markup.
- **Untracked `Replace` gives the new text the first replaced character's formatting as direct run properties**: from
  a heading, Heading 1's font, size, colour and bold, in a joined body paragraph. Tracked, it has none.
- **`font.bold` adds `w:b` only where a run isn't bold already** (a heading's run, bold through its style, gets none,
  and no tracked change). A mark inside the range is formatted too.
- **`insertComment`**: the start before the range's first run, the end after its last character before any mark,
  the reference after the end. On a control's content range, the start goes before the first paragraph, inside
  `w:sdtContent`.
- **A hidden run and a footnote reference in the range** go with it: untracked, the note leaves `footnotes.xml`;
  tracked, both stand inside the `w:del` (the editor's addition 2).
- **The editor's addition 5**, from inside a cell to after the table under tracking, is refused by Word, so no row
  is marked deleted. Addition 4 (`^p` against a cell's or row's end) never matches. Addition 3 (a comment on a whole
  range ending at a mark) ends before the mark.
- Two quirks of Word's are not proposed for copying: on a control's content range, `insertParagraph` adds an empty
  paragraph as well, before `NEWPARA` for `'Before'`, after it for `'After'`.

## 7. Risks

- **Joins under tracking.** A deleted mark between paragraphs of different properties is where CR-002 sections 29
  and 38 found Word's subtleties. Phase 2 reuses that code rather than writing new joins.
- **The editor's model.** If the editor wants spanning ranges in its verbs, its own span type changes. That is its
  CR, after D4.
- **Offsets.** A range's `start` is an offset in its first paragraph and `end` in its last. Code that compares
  `end` with `paragraph.text.length` assuming one paragraph has to look at `endParagraph`. All of it is in
  `src/model/content/`, and phase 1 audits it. *Audited in phase 1 (section 9.6).*

## 8. References

- CR-002 section 3.4's subset and section 7, the one-paragraph `Range`; phase G's `Comment.getRange`; section 29's
  `TrackedChange.getRange` for groups.
- CR-012 sections 17.11 (check 51) and 17.12 (the `SpanningRange` stopgap).
- `test/README.md` checks 30, 31, 48 and 51.
- Office JS: `Word.Range` (`expandTo`, `getRange`, `paragraphs`, `search` with `^p`), `Word.ContentControl.getRange`,
  `Word.Comment.getRange`.

## 9. Phase 1: the reads (2026-10-22)

*Jason, 2026-10-22: "implement CR-014 phase 1". Implemented the same day; committed `ef57729`.*

### 9.1 One `Range`, two ends

D1 as decided. `Range` keeps `paragraph` and `start` and gains `endParagraph`, the paragraph holding `end`, which is
`paragraph` itself for a range within one paragraph. Nothing changes for such a range: every one this package made
before is one. A range across paragraphs is made by `expandTo` or by a view (section 9.2), never by a caller's
constructor call (an internal one: its two paragraphs are in one story, the second after the first).

- **The story** (`story.mts`): `storyOf(p)` is the outermost array of block content above a paragraph. A text box's
  `w:txbxContent` and a kept `mc:AlternateContent`'s unread branch are stories of their own (out of scope, section
  3). `storyLayout` is that story's text in `body.text`'s form, with where each paragraph's text stands in it.
  `Body.text` is now that layout's text, its last mark dropped: one walk for both, and the sweep (section 9.5) found
  the text the same in every file.
- **`text`** is the layout's slice between the two ends: a mark `\r` between paragraphs, a table's cells a `\t`
  apart and its rows ended `\r\n`, a note reference `\u0002`, hidden text left out.
  - With `mark` (a 'Whole' range), the last paragraph's mark follows, `markText` (`\r`, or `\t` for a cell's last),
    as check 51 measured a cell's body's whole range (`CELLP1\rCELLP2\rCELLP3\t`, `OTHER\t`).
  - Or `endMark` in its place: a table's 'Whole' range ends `\r\n`, as Word gave it (check 52 case 1).
- **`paragraphs`** are every paragraph between the ends in the layout, a table's cells' among them. **`isEmpty`** is
  `text === ''`.
- **`style` and `styleBuiltIn`** (the latter new on `Range`), as checks 52 and 53 measured them.
  - A range holding its paragraphs whole reads their paragraph style where they share one, else `Normal`.
  - Any other range reads its runs' character style where they share one (`Emphasis`), and `''` (`Other`) where
    their character styles differ (a footnote reference among plain text, an Emphasis run among plain ones); with
    no character style, the paragraph style as above.
  - This holds within one paragraph too, which until check 53 read its paragraph's style.
  - Word's other `Other`, a paragraph after a nested table in a cell whose style isn't Normal, is not followed
    (section 9.7).
- **`font`** across paragraphs is a `MixedFont`: each run read in its own paragraph, through its own `w:pPr` and
  cell context. A property where the runs differ reads `null`; `''` for the name, the colour and the highlight;
  and `Mixed` for the underline, as Office JS's do (checks 52 and 53; their types say `boolean` too).
  - A range within one paragraph reads its first run, as CR-002 decided. Office JS reads `null` there too, but
    changing that is CR-002's and not proposed here (section 9.4 counts the difference).
- **`search`** goes a paragraph at a time, the first from `start` and the last to `end`. `^p` is phase 3's.
- **`getTrackedChanges()`** across paragraphs filters the story's listing by overlap, rather than taking each
  paragraph's listing. A paragraph's listing cuts a change at its edges (`HEAD1B\r` where Word has
  `HEAD1B\rBODY1A `); the story's groups it as Word does.
  - A range within one paragraph keeps the paragraph's listing. Its overlap test now handles a change whose range
    starts or ends in another paragraph.
- **`getComments()`** is each paragraph's, deduplicated.
- **`getRange`** gains `'After'`, as check 52 measured:
  - 'Whole' is the range itself, not widened;
  - 'Content' is the same without a last mark;
  - 'Start' and 'End' are collapsed at the two ends;
  - 'After' is collapsed after the end, at the next paragraph's start where the range holds its last mark.
- **`expandTo(range)`** (new): from whichever start comes first to whichever end comes last, a mark-holding end
  counting as later. It refuses two ranges in different stories with Office JS's `ItemNotFound` (check 53).
- **`tables`** (new) are the tables the range covers some of; a range inside one cell looks at that cell's tables
  instead (check 52: a range across a table, or into or out of one, lists it; a cell's whole range lists none).
- **`contentControls`** (new) are the controls the range overlaps, except one holding the whole range (check 53).
  So a range from a paragraph into a control lists it, one inside a control doesn't, and one over an inline control
  lists it. **`parentTableCell`** and **`parentTable`** (new) are its start's.
- **`hyperlink`** reads the first link among its paragraphs.
- **Every write refuses** (until phase 2, section 10) with `NotSupportedError` naming phase 2, before anything changes: `text`, `insertText`,
  `insertParagraph`, `insertOoxml`, `delete`, `replaceText`, `insertComment`, `insertContentControl`, `hyperlink`,
  `style`, `styleId` and `font`. So does `getOoxml` through the shim, and `toApiScript`, whose one `w:p` can't hold
  several blocks.
- `Paragraph.tableNestingLevel` is new (Office JS's), counted through `parentTableCell`.

### 9.2 The views

- **`ContentControl.getRange()`** of a block, row or cell control is one `Range` from its first paragraph's start to
  its last's end, 'Whole' with the last mark. `SpanningRange` (CR-012 section 17.12, never released) is gone; its
  reads are `Range`'s, and its writes refuse as before.
  - A row-level control, or one whose content ends with a table: 'Content' ends with the last cell's text, and
    'Whole' with `\t\t`, as Word 365 gives them (check 53). Its `text` leaves out that last row's end, and a row
    control's text has its cells a `\t` apart.
- **`Comment.getRange()`** is one `Range` (D2), `undefined` when the comment has no markers in the body. Check 52
  case 1's comment reads `BODY1B\rLIST1A`, as Word's does.
  - A start marker between blocks (Word writes one at a control's content's start, case 82) starts at the next
    paragraph; an end marker between blocks ends at the previous paragraph's end.
  - **`getRanges()`** is the per-paragraph array `getRange()` was. Its answers are the same in all 806 files of the
    sweep; the tests that indexed `getRange()` now call it.
- **`TrackedChange.getRange()`** of a group runs from its first piece to its last, across paragraphs.
  - A paragraph mark ends its own paragraph's range, holding the mark. So an inserted paragraph's range reads
    `…first.\r` in that paragraph alone, as the change's own text does (check 53, case 29). Until now it was the
    group's part in its first paragraph.
  - A deletion is collapsed where it starts, as Word's default view reads it (check 53: `''` in Simple Markup).
  - A lone mark change is the mark: an inserted mark `\r`, a deleted one collapsed at its paragraph's end
    (check 53). Until now it was its paragraph's whole range.
- **`Body.getRange()`** (new): 'Whole' from the first paragraph to the last, with its mark (check 51: a cell's
  `…CELLP3\t`, a footnote's `… NOTEP2\r`); 'Content' without it; 'Start' and 'End' collapsed.
- **`Table.getRange()`** (new): 'Whole' ending `\r\n` (check 52 case 1); 'Start' and 'End' collapsed; 'After' at
  the next paragraph.
- `test/office-js-subset.ts` lists the new members, so `Word.supported` names them.

### 9.3 The shim

Phase 1's test runs check 52's own snippet through the shim (section 9.4), which needed four things of it:
- a `ClientResult` holding a collection (`range.getComments()`, asynchronous here) has `load()` and `items`, so that
  Office JS code's `load`, `context.sync()`, `items` reads it;
- a collection of content controls has `getByTag` and `getByTitle`, as Office JS's has;
- an Array method on a collection (`map`, `filter`) gives what it gives. Before, an empty result became a collection
  that `JSON.stringify` threw on;
- the package's null `ListItem` takes `load()`, which Office JS code calls before reading `isNullObject`.

The snippet gained a `reads` mode: each kind's reads alone, nothing written.

### 9.4 Held to Word 365

`test/ranges.test.mjs`, five tests (ten in all with phase 0's):
- **The reads.** `CHECK52.runCheck(…, 'reads')` through the shim, on `check52Bytes`, against Word's run JSON. All 16
  kinds agree in every field (`text`, `isEmpty`, `style`, `font`, each `getRange`, `paragraphs`, `tables`, `swapped`)
  with two exceptions. `marks`, the `^p` hits, are phase 3's: none here yet. And TC2's style is Word's `Other`,
  asserted as it stands; FN's `Other` follows since check 53 (section 9.7).
- **The report on Word's save after Accept All.** Through the shim, against Word's report after Accept All: each
  region's `text`, paragraphs (text, style, alignment, table nesting, list), tables (values, whole range), tracked
  changes, comments (content and range) and content controls agree, region for region, with these exceptions:
  - `footnotes`: `Range.footnotes` is not implemented;
  - the bold cases' paragraphs: `getTextRanges` is not implemented;
  - a list paragraph after a nested table in a cell (cases 89, 91, 94): Word's `Other` again;
  - **cases 13 and 26**: Word, emptying a control, marks it `w:showingPlcHdr` over a run of five spaces. Office JS
    reads that as no text, in the paragraph and the control alike, where this package reads the spaces. That is the
    text reader's (CR-012), recorded here, not changed;
  - paragraph `font`, not compared: a paragraph reads its first run (CR-002). Of 1,269 reads, 40 differ from Word's:
    39 where Word reads `null` or `Mixed`, and one empty heading (case 60) where Word reads the mark's run properties
    (`w:b w:val="0"`) and this package the style's bold.
- **The spanning views on Word's save after the run.** The tracked changes each region's range lists agree with
  Word's in 19 regions, case 1's deletion and insertion through a mark among them (a deletion compared by its type,
  since Word's text was All Markup's). The other 11 differ by the listing's own grouping, CR-002's, asserted as they
  stand:
  - after a tracked `Replace` (8 cases), Office JS lists the insertion alone, where this package lists the deletion
    the markup holds too;
  - Word lists formatting across a table's cells as one change (cases 54, 55), this package a run at a time;
  - Word lists a tracked table insertion at a cell's end as one change (case 87), where this package lists the mark.
- **The API itself:** a range across paragraphs from `expandTo` (text, paragraphs, each `getRange`, the mixed font,
  `search`, every write refused with the text unchanged, a range ending at a mark, two stories refused), and the
  views (a block control, a comment and `getRanges()`, a tracked insertion through a mark, a table's and a cell's
  and the body's range, `tables`, `tableNestingLevel`).

### 9.5 Before and after

A sweep of the 806 `.docx` here, in docx4j's test resources and in the editor's, before (`96f735b`) and after.
**Unchanged** in every file:
- `Body.text`;
- each paragraph's text and whole range;
- each cell's text;
- `Comment.getRanges()` (the old `getRange()`);
- each paragraph's range's tracked-change count.

The same four files throw before and after. **Changed, as intended:**
- **24 control ranges in 22 files.** 18 row-level controls now range over their row. 6 block controls ending with
  a table now end as section 9.7 says, where `SpanningRange`'s 'Whole' had `\r\n\t`.
- **59 comment ranges in 47 files.** A comment across paragraphs reads `\r` where `getRanges()` has pieces.
  `invoice.docx`'s comment, whose markers stand between blocks, has a range at all.
- **132 tracked-change ranges in 78 files**: groups across paragraphs and groups with a mark, as section 9.2 says.
  These counts are from before check 53. Section 9.7 has the counts after it, which moved lone marks and deletions
  too.

### 9.6 The audit, and what phase 1 leaves

Section 7's offsets audit found each place that took a `Range` as one paragraph's:
- `Range.getTrackedChanges`'s overlap test, which now handles a change across paragraphs;
- `commentsOf` and `insertComment`, which are given a paragraph's piece, or refuse;
- `paragraphElementOfRange` (`toApiScript`, the shim's `getOoxml`), which refused until section 13.

`TrackedChange.getRange` (section 9.2) is a view, not a consumer.

Left for later:
- the writes (phase 2);
- `^p` and comments and controls inserted on such a range (phase 3);
- `Range.footnotes` and `Paragraph.getTextRanges`;
- a single-paragraph `font` reading `null` where runs differ (CR-002's);
- a lone mark change's range;
- the listing's grouping where Word's differs (section 9.4);
- Word's empty reading of a control showing its placeholder (CR-012's).

### 9.7 Check 53: the choices measured (2026-10-22)

Jason asked to measure in Word every choice section 9 made where Word was not measured ("let's measure in Word the
Choices you made where Word wasn't measured"). Check 53 (`test/README.md`) is read only: 35 cases, run on Word 365
(16.0.20326.20232) once in Simple Markup and once in All Markup. The JSONs are
`test/fixtures/check53/53-reads-word365-simple.json` and `-all.json`; the second run's label was left at Simple
Markup, but the file is the All Markup run, and it differs from the first exactly where the view matters.

**Confirmed:**
- A range across paragraphs of differing paragraph styles reads `Normal`, and of one style, that style (cases 1–3).
- A table's 'Start', 'End' and 'After' ranges (case 22).
- A cell control's ranges (case 18), a cell body's ranges and the tables they list, a nested table among them
  (case 23), and none from inside a nested table's cell (case 24).
- A range over an inline control lists it (case 21).
- An insertion through a mark (case 26) and a whole inserted paragraph's text (case 29).

**Changed to Word's answer:**
- **The style rule** (section 9.1). Word reads a partial range's runs' character style (`Emphasis`, case 7), and
  `Other` where they differ (case 8, and case 5's footnote reference, which explains check 52's FN). A range
  holding its paragraphs whole reads the paragraph style, even across a footnote reference (case 6).
- **`MixedFont`**: a colour or highlight where runs differ reads `''`, not `null` (cases 1 and 31).
- **Controls**:
  - a row control's text is its cells a `\t` apart (case 15);
  - a row control's, or a table-ending control's, 'Content' has no row end and 'Whole' ends `\t\t` (cases 15 to
    17), where section 9.2 had guessed `\r\n`;
  - their text leaves out the last row's end.
- **`contentControls`**: a range into a control lists it, and one inside a control doesn't (cases 19 and 20). The
  rule had been the controls the range holds whole.
- **Tracked changes** (section 9.2): a deletion's range is empty in Simple Markup (cases 25, 28 and 30), where it
  had become `\r`. An inserted mark is `\r` in its own paragraph (cases 26, 27 and 29). A lone mark's range is no
  longer the paragraph.
- **`expandTo`** across stories throws with Office JS's code, `ItemNotFound` (case 32).
- **`Range.parentTable`** is new, which the snippet's case 22 reads.

**Differences left, each outside phase 1.** `ranges.test.mjs` asserts each:
- **A single paragraph's `font`** reads its first run (CR-002; cases 8 and 31), where Office JS reads `null` or `''`
  where the runs differ.
- **Word's `Other` for a paragraph after a nested table in a cell** whose style isn't Normal (cases 10 and 12, as
  check 52's TC2). A paragraph before one, or a Normal one after, reads its style (cases 11 and 14). Not followed:
  it hides the style the paragraph has. Recorded as an intentional departure in CR-002 section 57 (Jason,
  2026-10-22).
- **The tracked changes in Simple Markup:**
  - a change's text here is docx4j's, which is All Markup's (`HEAD25B\rBODY25A `), where Simple Markup reads a
    deletion's text as `''`;
  - Word joins the paragraphs a deleted mark ends, so its collapsed deletion touches two paragraphs (case 25);
  - Word's paragraph listing gives each paragraph its part of a change (case 25's second paragraph lists
    `BODY25A `), and a wholly deleted paragraph none (case 30).

  The listing is CR-002's; the joined paragraphs are the readers' (CR-012).
- **`Body.text` in Simple Markup** (case 33) joins the paragraphs at a deleted mark, leaves a deleted paragraph out,
  and reads a control showing its placeholder as empty: 40 characters fewer than this package's here. This is
  CR-012's readers.
- **A control showing its placeholder** (cases 34 and 35): Word's text, its paragraph's and the control's, is `''`
  for Word's own placeholder text as for check 52's five spaces, and `placeholderText` reads `null`. This package
  reads the placeholder's text (CR-012's readers); `showingPlaceholderText` is not implemented.
- **`font.color`**: an automatic colour reads `#000000` in Word and `''` here. A theme colour resolves differently:
  Heading 1's reads `#0F4761` in Word and `#365F91` here. Both are `Font`'s (CR-002 and CR-008), outside the
  comparison except where colours differ across paragraphs, which agrees.

**Before and after, again** (section 9.5's 806 files plus check 53's): control and comment ranges as there, and 472
tracked-change ranges in 223 files:
- 213 lone inserted marks read `\r` where they read their paragraph;
- 198 deletions read `''` where they read their paragraph (a lone deleted mark's) or `\r`;
- 60 insertions read exactly the change's text;
- one formatting change.

Nothing new throws.

## 10. Phase 2: the writes (2026-10-22)

*Jason, 2026-10-22: "implement CR-014 phase 2". Implemented the same day.*

The writes follow check 52's answers (section 6.2). Each goes through `src/model/content/spanEdits.mts`, which
`Range` calls for a write across paragraphs, and for a delete (or a tracked `Replace`) of a range holding its
paragraph's mark.

### 10.1 What each write does

- **Where a range stands** (`extentOf`): the innermost list of blocks holding both ends, and in it the block holding
  each end. That block is the end's paragraph, or a table or block content control the range holds whole: it
  starts at the block's first paragraph's start, or ends with its last paragraph's mark or a table's row end
  (`Table.getRange()`, a body that starts with a table). A range starting or ending inside a block it doesn't hold
  whole is refused, before anything changes:
  - inside a table with `GeneralException`, which is Word's own refusal (check 52's INTO and CELL, tracked or not;
    D3);
  - inside a content control with `NotSupportedError`, since Word wasn't measured there.
- **`delete()`**, untracked (cases 3 to 15):
  - the first paragraph loses its text from the range's start, and every block between goes, a table too;
  - the last paragraph loses its text to the range's end, and the two join in the **last paragraph's properties**,
    each run keeping its own formatting. This is the resolution side's `joinWithNext`, which Accept All has always
    used, so the joined paragraph also keeps the last paragraph's `w14:paraId`.
  - Where the range holds the last paragraph's mark (a 'Whole' range), the mark goes too, and the paragraph after
    is joined in, in its properties (W1, W2). With no paragraph after it in the same list (a cell's last paragraph,
    a control's, the body's, one before a table) the mark stays. CR-002 section 55's rule holds: a cell keeps its
    last paragraph.
  - A footnote or endnote whose reference goes is queued to leave the notes part on save (case 12), unless a
    reference to it still stands then.
  - A section break in the first paragraph's mark goes with the mark, and the relationships its headers and footers
    held are released. A block-level range marker between the paragraphs stays, and the join brings it into the
    joined paragraph, as Accept All's does.
- **`delete()`**, tracked (cases 16 to 28): nothing is joined.
  - The text is a `w:del`, a hidden run and a note reference inside it.
  - Every mark in the range is marked deleted, except a last one the range doesn't hold, and **no `w:pPrChange`**
    is written (the editor's addition 1).
  - A table in the range has its rows' `w:trPr/w:del`, and its cells' text and marks deleted.
  - A mark another author's deletion already marked is left as it is, and its paragraph's text is still deleted.
    (`Paragraph.markDeletedInPlace` skipped that text until now; a table row's deletion inherits the fix.)
  - Accept All gives the untracked answer, Reject All the original (section 10.5).
- **`insertText(…, 'Replace')`** is `delete()` keeping a last mark, then the text where the paragraphs joined:
  - **untracked** (cases 29 to 38), in the first replaced character's look (`lookOf`): its run's own properties,
    and each property its paragraph's style gave it that the joined paragraph's doesn't, written direct. A heading's
    font, size, colour and bold come into a body paragraph (case 29). A body paragraph's run gets `w:b w:val="0"`
    and `w:color w:val="auto"` in a heading (case 30), the neutral value of what the heading's style sets.
  - **tracked** (cases 39 to 48), the `w:ins` before the deletion, in the first replaced run's direct formatting,
    as Word writes it. A range within one paragraph holding its mark is replaced so too (case 47); one that doesn't
    hold its mark keeps CR-002's order, the `w:del` first.
- **`font`** (cases 49 to 56): every run in the range, a table's cells' too, split at the range's ends, and each
  mark the range holds.
  - A toggle (bold, italic, strikethrough, double strikethrough, subscript, superscript) is written only where the
    run or mark doesn't read it already: no `w:b` on a heading's runs, bold through its style, and no tracked change
    there.
  - Tracked, each run and mark written has its `w:rPrChange`. A mark's is a `w:pPr/w:rPr/w:rPrChange` (case 56), new
    in the tracker (`recordMarkRPrChange`, `settleMarkRPrChange`); the listing already read it as a formatting change.
- **`insertText` at 'Start' or 'End'** (cases 71 to 78) goes at the first character and at the last, before a last
  mark ('Before' and 'After' alike).
- **`insertParagraph`** (cases 57 to 70) goes where the range starts ('Before') or ends ('After'):
  - **inside a paragraph**, the paragraph is split there (`splitParagraph`). The text before the split becomes a new
    paragraph in front, in a copy of the properties, with a mark of its own in the formatting of the run before the
    split. The original keeps the rest, its own mark, its section break and its identity. The new paragraph goes
    between the halves, its mark and text in that same formatting (case 63's italic).
  - A link, a revision or a smart tag the split falls inside is cut into two of the same, a revision's second half
    under a new id. A content control, a field, an equation or a kept `mc:AlternateContent` there is refused, since
    each half would be one of it.
  - At a paragraph's start the new paragraph goes before the paragraph, and at its end after it, as
    `Paragraph.insertParagraph` puts one.
  - Tracked, the first half's mark and the new paragraph are inserted (cases 69 and 70). Reject All joins them back.
- **`replaceText`** across paragraphs goes a paragraph at a time, as `search` does (section 3).
- **The range after a write**: collapsed where a deletion closed up (tracked, where it starts, as Simple Markup
  reads it); over the new text after `Replace`; after `insertParagraph` and `insertText`, still over its text.

### 10.2 Changed for a range within one paragraph

- **A range holding its paragraph's mark deletes the mark**: `paragraph.getRange('Whole').delete()` removes the
  paragraph, joining the paragraph after it in, where it emptied the paragraph before (check 52 case 14). The same
  goes for a single-paragraph block control's 'Whole' range and for an inserted mark's range. Where no paragraph
  follows in the same list, the paragraph is emptied as before.
- A tracked `Replace` of such a range writes its insertion before the deletion (case 47).
- `Paragraph.markDeletedInPlace` deletes a paragraph's text even where its mark is already deleted (section 10.1).

### 10.3 Refused under change tracking

*Since check 54 (section 12.4) a content control or custom XML element held whole is deleted in Word's markup; an
`w:altChunk` and a kept `mc:AlternateContent` are still refused.* As phase 2 had it: a tracked `delete` or `Replace` whose range holds a whole block that isn't a paragraph or a table is refused with
`NotSupportedError`, before anything changes (`assertTrackable`). Such blocks are a content control, a custom XML
element, an `w:altChunk` and a kept `mc:AlternateContent`.
- A control's content can be marked deleted, but what Word writes for the control itself isn't measured: check 30
  measured an inserted control's markers only.
- An `w:altChunk` has no deleted form.
- Accept All would leave such a block, and an empty paragraph in a control, where the untracked deletion takes it,
  the paragraphs either side unjoined. The sweep (section 10.5) found exactly that before the refusal.

Untracked, such a block goes whole, as a table does.

### 10.4 Held to Word 365

`test/ranges.test.mjs`, seven tests (eighteen in all with phases 0 and 1):
- **The run.** Check 52's snippet in its `run` mode, through the Word shim on `check52Bytes`, against Word's run JSON.
  Compared for each of cases 1 to 85: the error, the result, the range read first, the region's paragraphs, its
  tracked changes' types and its comments; untracked, also its text, tables and controls. Under tracking Word's text
  is All Markup's, so a tracked region's paragraphs are compared without their text. Every refusal is Word's
  (cases 9, 10, 22, 23, 34, 35, 44, 45). The differences left, asserted as they stand:
  - after a tracked `Replace`, Office JS lists the insertion alone, and this package lists the deletion too (CR-002's
    listing, phase 1's finding);
  - the bold cases' report reads each word through `getTextRanges`, which is not implemented, and Word lists a
    formatting change across a table's cells as one (cases 54 and 55);
  - Word's extra empty paragraph on a control's content range (cases 60 and 66), which section 6.2 decided not to
    copy;
  - `insertComment` across paragraphs, phase 3's (cases 79 to 85; 83 inserts its one-paragraph comment, and the
    snippet's chained `getRange()` fails on the asynchronous result).
- **Accept All.** Accept All after the shim's run, saved and reloaded, reported through the shim against Word's
  report after its Accept All. Every case from 1 to 85 is the same, the tracked deletions and replacements among them.
  The exceptions are the categories above: the bold cases, 60 and 66, and the comments.
- **The markup.** Each region's markup after the shim's run, a line a paragraph, held to Word's save after its run.
  A line gives the paragraph's style, alignment and numbering, its mark's properties and revisions, a `w:pPrChange`,
  then its runs merged by revision and properties, and a table's rows. Every case from 1 to 85 is Word's markup but:
  - 1, read only: Word's save gave the comment reference's run `w:u` and `w:sz` of its own;
  - 13: Word's filler for an emptied control, a bold run of five spaces;
  - 60 and 66: Word's extra paragraph;
  - **67 and 68**: Word gives a paragraph's mark its last run's formatting when `insertParagraph('After')` falls at
    the paragraph's end. Not followed, a departure: the mark is left as it was, as `Paragraph.insertParagraph`
    leaves it, and under tracking carrying it would need a formatting change on the mark that Reject All undoes;
  - 79 to 85: the comments.

  The test also pins Word's own lines for case 16 (the tracked deletion's marks), 39 (`w:ins` before `w:del`), 29
  (the heading's look), 56 (a mark's tracked bold) and 70 (a tracked split).
- **The API**: refusals before anything changes (a table, a control, under both modes), Reject All after each
  tracked write giving back the original, a footnote leaving on save, a table's, a cell's, a control's and the
  body's whole range deleted, a split inside a link and refused inside a control, and `insertText` and
  `replaceText` across paragraphs.

### 10.5 The sweep

The 808 `.docx` of section 9.5, each with one range from the middle of a paragraph to the middle of one up to four
later in reading order. Each of six writes was made untracked, then tracked on a fresh load: `delete`, `Replace`,
`bold`, `insertParagraph` 'Before' and 'After', and a 'Whole' range's `delete`. Each document was saved and reloaded
after every write. 685 documents had such a range; 119 had too few paragraphs, and 4 (strict namespace) don't load.
Accept All after the tracked write was held to the untracked write, Accept All'd too, in each paragraph's text,
style and alignment. Reject All was held to the original, Reject All'd, so that changes a document held already
cancel out.
- **Every write that wasn't refused agreed both ways, in every document.**
- Refused: 224 deletes and as many replacements, all the ranges starting or ending inside a table (167) or a
  control, holding a control or `w:altChunk` under tracking, or cutting an equation; 30 paragraph splits inside a
  control or an equation; 4 bolds cutting an equation; one 'Whole' delete taking an equation under tracking.
- Before section 10.3's refusal and the `markDeletedInPlace` fix, 30 documents differed on Accept All. A control or
  `w:altChunk` between the ends was left behind and blocked the join. A paragraph between whose mark another author
  had deleted kept its text.

### 10.6 Choices check 52 didn't measure

Each could go to a check 54, as check 53 measured phase 1's:
1. A range within one paragraph: `insertParagraph` still goes before or after the paragraph (CR-002), where Word
   splits a paragraph a range across paragraphs starts or ends in.
2. A range ending with a whole table after a paragraph's tail: the paragraph keeps its mark, not joined to the one
   after the table. A W2 range's head is joined that way.
3. A 'Whole' range of a paragraph followed by a table, or last in its cell, control or body: emptied, its mark kept.
4. A range holding a whole block control or `w:altChunk`: taken whole untracked, refused tracked (section 10.3). A
   range starting or ending inside a control: refused.
5. The toggle rule for the toggles other than bold. Name, size, colour, highlight and underline are written to every
   run.
6. `lookOf`'s neutral values beyond `w:b w:val="0"` and automatic colour (`u`, `highlight`, `vertAlign`, `sz` 20).
7. A tracked `Replace` within one paragraph, not holding its mark: CR-002's `w:del` first, where Office JS wrote
   `w:ins` first for each shape check 52 measured.
8. Comments whose reference or markers a deletion takes, and bookmarks it breaks: left as they fall. Notes go
   (measured).
9. A tracked deletion over a mark this author inserted: marked deleted (inserted and deleted), not taken back.
10. A range across paragraphs whose `style`, `styleId` or `hyperlink` is set, or given `insertOoxml`: still refused,
   as section 3 didn't propose them.

### 10.7 Left

- Phase 3: comments and content controls on such a range, and `^p` in `search`.
- `Range.footnotes` and `getTextRanges`; CR-002 section 55 (cases 86 to 95; implemented 2026-10-23, its section 55.7).
- Deleting a range within one paragraph, `Paragraph.delete` and `Body.clear` leave a note whose reference they take
  in the notes part (CR-002's).

## 11. Phase 3: comments, `^p`, and check 54 (2026-10-22)

*Jason, 2026-10-22: "do it", to the order proposed: check 54 first, then comments and `^p`, which check 52 measured,
while it waits for his run, then content controls on its answers. Comments and `^p` implemented the same day; not
committed.*

### 11.1 Check 54

`test/README.md` check 54, `test/fixtures/check54/` (`check54Bytes`, the snippet `check54-script-lab.js`), and the
share's `__tmp/54/`. Its 43 cases measure the ten choices of section 10.6 and `insertContentControl` across paragraphs
(cases 36 to 43), a region each as in check 52, in one run with Word's saves after it and after Accept All. Each
case names the section 10.6 item it answers. `ranges.test.mjs` holds its file and snippet together: each case parses
to the case it was built as, each word its range names is in the file once, and the snippet runs through the shim
with no error of its own. The README entry records what this package gives before the run.

### 11.2 Comments across paragraphs

`Range.insertComment` on a range across paragraphs puts the markers as Word did (check 52, cases 79 to 85):
- the start before the range's first character, in its first paragraph;
- the end and the reference run after its last character, in its last paragraph, before a mark the range holds;
- a table between, or a cell at either end, stays as it is (cases 80 and 81).

An insertion at either end is handled as within one paragraph: the start goes inside it, and the end splits it.
The comment's `getRange()` is the range commented.

Two differences from Word's markup, not followed:
- **Word gives the reference run the text's size** (`w:sz 22`, over the CommentReference style's 8 points) and,
  where text follows it, that text's italic. The reference run here has the style alone, as CR-002's has within one
  paragraph.
- **On a control's content range Word put the start marker before the control's first paragraph**, in
  `w:sdtContent` (case 82). Here it is inside that paragraph, which reads the same (section 9.2).

**Found by the sweep (section 11.5), fixed within one paragraph too.** A marker went beside a whole run where the
range began after a tab or break in that run: `splitAt` cut only inside a `w:t`. The comment's range then started a
character early. The runs are now split as `font` splits them (`splitRunsAround`), which `placeAround`'s own doc
had claimed since CR-002 phase G. `ranges.test.mjs` holds both shapes, and the test fails with the old split.

### 11.3 `^p` in `search`

A plain query holding `^p`, Word's paragraph mark, matches across paragraphs, as Word found (check 52, case 2, and
each kind's `marks`):
- the query is split at its marks (`paragraphMarkPieces`; `^^` stays a caret). A hit runs from the first piece's match
  ending at a paragraph's end, through each middle piece matching a whole paragraph's text, to the last piece's match
  from the next paragraph's start;
- a query ending with `^p` ends with that paragraph's mark, which the hit holds: `^p` alone is a range of `\r`, its
  paragraph the one whose mark it is;
- a mark matches only a paragraph's own mark that the range holds: never a cell's end or a row's end (the last
  paragraph of a cell reads `\t`), and never across a table's edge, so every paragraph of a hit is in the same cell or
  in none (`LIST2B^pT2R1C1A` found nothing). A mark between two paragraphs of one cell is a paragraph's, and matches;
- `^p^p` is an empty paragraph's mark after another mark.

`Paragraph.search` and `Body.search` with `^p` search their 'Whole' range, which holds their marks. Wildcards and
regular expressions don't read `^p` as a mark: wildcards read it as the two characters, as they did.

**`replaceText` with `^p` in its query is refused** (`NotSupportedError`). The hits are found, but what Office JS's
replacement of a lone mark leaves isn't measured, and phase 2's `Replace` keeps a last mark. Until now such a query
matched the two characters `^p`.

### 11.4 Held to Word 365

`test/ranges.test.mjs`, twenty-one tests in all now:
- **Phase 1's reads test** compares `marks`, each kind's `^p` hits, which it had left out: all 16 kinds agree with
  Word's.
- **Phase 2's run test** compares case 2's eleven queries, each hit's text and paragraphs: Word's. Cases 79 to 85 now
  differ from Word's run only in the snippet's own error: it calls `getRange()` at once on what `insertComment`
  returns, which is asynchronous here. The comments and their ranges are Word's.
- **The Accept All test** finds the comments' regions Word's too.
- **The markup test** holds each comment's markers to Word's save, the two differences of section 11.2 aside.
- **A phase 3 test**: comments across a table, from a paragraph's end, to a paragraph's start, read back after a
  save; `^p` against marks, a hit across a mark, a table's edge, a cell's paragraphs and an empty paragraph, the body
  and a paragraph, wildcards, and the `replaceText` refusal. Another test holds the split runs of section 11.2.

### 11.5 The sweep

The 808 files of section 9.5. In each, one comment over a range from a paragraph's middle to one up to four
paragraphs later, saved and reloaded, and `body.search('^p')`:
- 680 comments read back as the range they were made over;
- 5 were refused, the range starting inside an equation (CR-012 phase 4's rule);
- 119 files have no such range, and 4 (strict namespace) don't load;
- every `^p` hit is a lone `\r`. Each file's count is its paragraphs outside tables plus those not last in their cell,
  the `mc-fun` probes' selected branches among them.

Before section 11.2's fix, two files' comments started a character early.

### 11.6 Left for phase 3

- Content controls on a range across paragraphs, on check 54's answers (cases 36 to 43): done, section 12.2.
- Section 10.6's choices, changed to Word's where check 54 differs: done, section 12.1.
- `replaceText` with `^p`, should a use ask: replacing a mark is check 54's untested neighbour.


## 12. Check 54's run (2026-10-22)

Jason ran check 54 on Word 365 (16.0.20326.20232, PC) the same day, in All Markup as check 52's (deleted text shows in
the run's text reads), and accepted all. `fixtures/check54/54-writes-word365-run.json` is the run and
`54-writes-word365-accepted.json` the report after Accept All. Word's save after Accept All is `test/check54/` (a
SHA-1 manifest, outside `fixtures/` as check 52's are). **Word's save after the run came from a second run.** The
first run's save, and a re-run's, held the document from before the run: Ctrl+S had gone to Script Lab's pane,
which has the focus after the run. The re-run's JSON was the first run's, case for case. Saved with File > Save,
`test/check54/54-writes-word365.docx` holds the run, Word's tracked markup among it (section 12.4).

### 12.1 What Word did, and what changed here

Section 10.6's items, by number:
1. **`insertParagraph` on a range within one paragraph splits the paragraph** where the range starts ('Before') or
   ends ('After'), tracked too (cases 1 to 3), as across paragraphs. `Range.insertParagraph` now does so for every
   range (`insertParagraphAt`). Until now a range within one paragraph put the paragraph before or after it.
2. **A range from a paragraph's tail to the end of a whole table**: deleted, the paragraph's head joins the paragraph
   after the table, in its properties (case 5), and tracked its mark is deleted (case 6). Replacing it is refused
   with `GeneralException` (case 7). Both followed: a mark the range holds goes with it, whatever ends the range.
3. **A 'Whole' range with no paragraph after it in its list**:
   - before a table (case 8) and at a control's end (case 12), the paragraph goes;
   - a cell's last paragraph is emptied and kept (case 10), and tracked its text is deleted and not its mark
     (case 11, CR-002 section 55's rule);
   - a story's last paragraph keeps its mark too (unmeasured, Word can't delete the last mark).

   Followed (`dropMark`, `deleteMark`).
4. **A whole content control in the range**:
   - untracked, it goes and the paragraphs join (cases 14 and 16), as before;
   - tracked, Word deletes it too (cases 15 and 17), with check 30's markers for an inserted control turned to
     deletions: a `w:customXmlDelRangeStart` before the control, its end first inside its content, a second start
     last inside, its end after the control, and the content deleted between them. Accept All took the control and
     joined the paragraphs. **Followed** (section 12.4): section 10.3's refusal is lifted for controls.
   - **A range into a control or out of one** (cases 18 to 20): the texts go, nothing joins across the control's
     edge, and the start's paragraph takes the end paragraph's properties, its mark in its last run's formatting.
     Followed for `delete()`, one control level down (`extentOf`'s `crossControls`, `takeProperties`). Tracked, the
     start's mark and the texts are deleted (case 19's listing). `Replace` across a control's edge stays refused,
     unmeasured.
5. **Every property, not toggles alone**: Word wrote no `w:sz` on a heading's runs already at 14 points (cases 23
   and 24), and no `w:i` on runs italic through their style (cases 21 and 22). It underlined the list item's mark,
   whose runs were underlined directly and the mark not (case 25). `formatSpan` now writes a property only where a
   run or mark doesn't read it already, whatever the property.
6. **`lookOf`'s neutral values**: `w:caps w:val="0"`, automatic colour, the old size, `w:u w:val="none"` (case
   26), as this package wrote them.
7. **A tracked `Replace` within one paragraph**: Word put the `w:ins` first (case 4), as across paragraphs.
   **Followed** (Jason, 2026-10-23: "Tracked Replace order within one paragraph: follow Word (insertion first)"):
   every tracked `Range.insertText(…, 'Replace')` now goes through `replaceAcross`, the insertion first. CR-002
   wrote the `w:del` first, and the editor's agreement tests pin that order (its message of 2026-10-22), so they
   move. Inside another author's insertion the new text still splits it and isn't nested in it (CR-002 section 37);
   the deletion now stands in the half after the new text.
8. **Comments and bookmarks a deletion reaches**:
   - a comment whose reference goes with the range goes, untracked (case 27) and, tracked, on Accept All (case 28);
   - one whose reference stays keeps its markers at the join (case 29), as a bookmark does (case 30).

   Followed: `cut` takes a comment reference in the span and queues the comment, as a resolution queues one.
9. **A tracked deletion into a paragraph this author inserted** (case 31): Word took its own inserted mark back and
   joined the two paragraphs, its model having marked the first paragraph's mark inserted (CR-002 section 13 marks
   the new paragraph's). Accept All and Reject All give the same document here as Word's; not changed.
10. **Across paragraphs:**
    - **`style`** (cases 32 and 33): Word gave each paragraph Heading 2, and took away the body paragraph's centring
      and italic and the list item's numbering, keeping its underline. Followed in part: `style`, `styleId` and a
      new `styleBuiltIn` setter set each paragraph's style. Word's resets aren't followed: they are its paragraph
      style's rules, which `Paragraph.style` doesn't have either, and were not measured within one paragraph
      (CR-002's).
    - **`hyperlink`** (case 34): Word wrote a HYPERLINK field from the first paragraph into the last. **Still
      refused**: this package's links are `w:hyperlink` elements, one paragraph's each, which its getter reads.
    - **`insertOoxml(…, 'Replace')`** (case 35): one paragraph's runs in place of the range, where the paragraphs
      join, as `Replace` (`HEAD35A NEWOOXML BODY35B`). Followed; 'Before' and 'After' put one paragraph's runs at the
      range's start and end, and several blocks before its first paragraph or after its last. Several blocks in
      place of a range across paragraphs stay refused, unmeasured.

`insertOoxml`'s `Replace` within one paragraph now keeps a last mark, as a replacement does, where it called
`delete()`, which since phase 2 takes the mark of a 'Whole' range.

### 12.2 Content controls across paragraphs (phase 3)

**Word puts a control around paragraphs held whole and refuses any other range** with `GeneralException` (cases
36, 37 and 39 to 42: a range starting or ending inside a paragraph, across a table, into one, out of one, tracked
or not). On a range from the first paragraph's start to the last's mark it wrapped both in one block control
(cases 38 and 43). `Range.insertContentControl` across paragraphs now does the same (`wrapParagraphs`). Tracked,
the control is an insertion with check 30's markers, as a paragraph's control is, and Word's save after the run has
the same markers (case 43). Word's run listed them as two insertions of no text, where this package's listing lists
none.

### 12.3 Held to Word 365

`test/ranges.test.mjs`, five tests:
- **The file and the snippet agree**, the snippet running through the shim with no error of its own.
- **The JSONs are of this file**, and the saves are Word's, their SHA-1s in the manifest.
- **The run**: the snippet's `run` mode through the shim against Word's run, case for case, as check 52's run test
  compares. The differences left are each listed with their reason:
  - case 4's order (12.1, item 7);
  - Word's `Other` for an emptied list item in a cell (cases 10 and 11, CR-002 section 57's family);
  - the listing's grouping (cases 13 and 43), and a tracked `Replace` listing its deletion too (17);
  - All Markup's comment range (28);
  - Word's model of an inserted mark (31);
  - a style's resets (32, 33);
  - the hyperlink (34);
  - `insertOoxml`'s asynchronous result (35).
- **The markup**: each region's markup after the shim's run against Word's save after the run, as check 52's markup
  test compares. Every case is Word's markup, case 4 among them since item 7, but a style's resets (32, 33), the
  hyperlink (34), the reference run's size (29, and 28, where Word's save wrote it on the file's own run), and Word's
  model of an inserted mark (31). Case 28 also found that a tracked deletion left a comment's reference run beside
  the `w:del`, where Word puts it inside, fixed by section 12.6. The test pins Word's lines for the deleted control
  (15) and the inserted one (43).
- **Accept All**: Accept All after the shim's run, saved and reloaded, against Word's report after Accept All and
  its save, each region's markup line by line. **Every case is Word's**, the tracked control deletions among them,
  but a style's resets (32, 33), the hyperlink (34), the reference run's size (29), the `Other` style (10, 11, the
  report alone), and three rules of Word's Accept All (cases 9, 13 and 19):
  - **an emptied paragraph whose deleted mark has a table after it goes** (case 9);
  - **a control's emptied last paragraph keeps a bare mark** (case 13);
  - **a paragraph whose deleted mark has a control after it takes the control's first paragraph's properties**, its
    content not joined (case 19).

  This package's Accept All is docx4j's resolution (CR-011), held to docx4j's answers by the revisions goldens, and
  docx4j has none of the three. They go to docx4j (its CR-043) as findings. The goldens are docx4j's, so this
  package's resolution changes when docx4j's does.

  *Followed since 2026-10-23.* docx4j adopted the three, and the deleted control of cases 15 and 17 as this package
  resolves it, as its CR-043 section 17 (`29dbf5964`, `VERSION_17_3_5`, not yet released; Jason's decision there). It
  holds them to Word's save of this check, copied into its core-tests as `revisions/core-ts/check54/`. This package
  ports them:
  - **`joinWithNext`** (`TrackedChange.mts`), where no paragraph follows a mark that goes, does one of three things,
    as docx4j's `Resolver.joinWithNext` does: an emptied paragraph with a table next goes (case 9); one with a block
    control next takes the control's first paragraph's properties, nothing joined (case 19); a control's emptied last
    paragraph keeps a bare mark (case 13). The rules come after this package's own fallback for an inserted mark
    rejected at a container's end (CR-002 section 13), which docx4j hasn't.
  - **`ResolvedView`**, docx4j's `ResolvedView` as changed with them. It removes case 9's paragraph. It removes a table
    whose every row it removes (case 6, which this package's resolution already did). It gives a paragraph its
    properties through `paragraphPropertiesOf` (cases 13 and 19). And it takes a control deleted whole, accepted,
    with everything in it, where before a deleted control rejected read as unwrapped. The text model's `Extractor`
    reads a paragraph's style and properties through `paragraphPropertiesOf`, as docx4j's `viewPPr`.
  - **The filter** takes `CONTENT_CONTROL_DELETION` among its kinds. It was missing since `87bbb38`, so a filter
    naming it was refused.

  The Accept All test lists none of the three now: every case is Word's, but for 10, 11 and 29, the style resets
  (32, 33) and the hyperlink (34). A new test in `text-model.test.mjs`, after docx4j's `ViewAgreementTest`, reads
  Word's save after the run: its accepted view, this package's Accept All copy, and, region by region for cases 6, 9,
  13, 15, 17 and 19, Word's own Accept All save, block by block. The view's half fails without the `ResolvedView`
  change. docx4j at `29dbf5964` (built into a separate Maven repository) answers the same as 17.3.4 for all 96
  revisions goldens and all 282 text goldens: no fixture has these shapes, so the goldens stay as recorded.
  `npm test`: 1,490 pass. Committed `222cc20`.

  *Revised the same day, docx4j `e55a83b8e` (Jason: "yes").* Jason ran Word 365's Reject All of check 54's save: the
  original everywhere, as docx4j's and this package's. A second probe, docx4j's text-left probe, put the paragraph's text
  left beside the mark, and changed three of the rules. docx4j's order where no paragraph follows a mark that goes is now
  this, and this package's `joinWithNext` follows it:
  1. an emptied paragraph with a table or a block control next goes (the control is new: the probe's case 3);
  2. text left with a table next joins the table's first cell's first paragraph, at its front, the range markers after it
     with it; that paragraph keeps its own properties (new: case 1, where the paragraph had stayed before the table);
  3. text left with a block control next takes the control's first paragraph's properties (check 54 case 19, the probe's
     case 5);
  4. a control's last paragraph keeps a bare mark, text left or not (before, only when emptied: case 2);
  5. otherwise the mark is only dropped.

  `ResolvedView` follows (`removes`, `joinsIntoTable`, `paragraphPropertiesOf`), and so does the text model's walk, which
  holds a chain joining into a table until the table's first cell's first paragraph, as docx4j's Extractor does. Word's
  saves are copied: check 54's Reject All into `check54/` (its manifest), the probe with Word's Accept All and Reject All
  saves and docx4j's builder into `s17/`, with a manifest. `ranges.test.mjs` holds Reject All of check 54's save to
  Word's, every region, and the probe's Accept All and Reject All to Word's, region for region but case 4 (a deleted
  mark on a cell's last paragraph, which Word doesn't write; docx4j's section 11, not followed). `text-model.test.mjs`
  holds the probe's accepted view to its Accept All copy and to Word's save. docx4j at `e55a83b8e` answers the same
  as 17.3.4 for the 96 revisions goldens and the 282 text goldens. `npm test`: 1,497 pass.

### 12.4 What the save after the run settled

- **A tracked deletion over a whole content control** (cases 15 and 17): Word's `w:customXmlDelRange*` markers
  (12.1, item 4). The tracked deletion writes them (`markHolderDeleted`, check 30's `markHolderInserted` with the
  Del elements). The `holder` tracked change, which listed and resolved an inserted control, now takes a deleted
  one too (`holderChangeAt`): it is listed as `Deleted` with the deletions around it (`CONTENT_CONTROL_DELETION`, an
  extension of docx4j's kinds, as the insertion's handling already was); Accept takes the control away; Reject, the
  markers. A Word document holding them now resolves as Word does too. section 10.3 refuses only an `w:altChunk`
  and a kept `mc:AlternateContent` now.
- **A tracked control around paragraphs** (case 43): Word's markers are check 30's, which this package writes.
- **A tracked style** (case 33): a `w:pPrChange` on each paragraph, as this package writes; the resets aside.

### 12.5 The sweep, again

Section 10.5's sweep with check 54's rules, 808 files, the tracked control deletions included: 24 deletions and
replacements over whole controls that section 10.3 refused now pass both ways. Every write that wasn't refused agreed
with the untracked write after Accept All and with the original after Reject All, but for two kinds of difference:
- **74 'Whole' deletions** of a paragraph followed by a table or a control: untracked the paragraph goes, as Word's
  case 8; tracked and accepted it stays empty, Word's Accept All rule of case 9 that docx4j's resolution lacks;
- **two** across a control's edge (out of a table of contents into the heading after it, `loadAndSave.docx`; from a
  title into the second of several bound controls, `hyperlink-binding-test.docx`): case 19's Accept All rule.

Reject All gave the original everywhere.

### 12.6 A comment's reference in a deletion

Section 12.3's case 28 found that a tracked deletion left a comment's reference run beside the `w:del`, where Word
puts it inside. In case 28 the two gave the same document once accepted or rejected, but only because the deletion
took all of the comment's text: Accept All removes a comment whose text a resolution empties (CR-002 section 54).
Tried here (Jason asked, 2026-10-23), a comment over "B C" with its reference after "C", and "C D" deleted:
- **tracked**, the reference stayed outside the `w:del`, so Accept All kept the comment over "B", where Word's markup
  takes the reference with the deletion;
- **untracked, within one paragraph**, the deletion kept the comment whose reference it removed: a reference has no
  text, so the splice's segments never held it. Across paragraphs `cut` took it (case 27).

**Fixed the same day** (Jason: "yes make it"), in the paragraph's own primitives, so every deletion has it, within
one paragraph and across:
- `Paragraph.splice`, untracked, takes a comment reference run in the span (strictly inside it, or at an edge the
  span holds) with the text, and queues its comment to go on save, its markers and replies with it, as a resolution
  queues one (`commentReferencesIn`, `queueCommentRemoval`);
- `deleteText`, tracked, puts the reference run into the `w:del` with the text, where it stands among the runs,
  as Word's case 28 has it; taking back this author's own insertion takes such a comment too;
- **the resolution**, docx4j's `Removals` rule that the port lacked: a comment whose reference or range marker the
  resolved content held is judged when the resolution ends, and goes whole where no reference to it is left
  (Word keeps no comment without one, core-ts check 32), or loses the half of its range left where only half is
  (`settleCommentAnchors`). docx4j has had it since its CR-043 phase 1, before the commit its revisions goldens were
  recorded at, and the 96 goldens still agree: none of their fixtures had a comment's reference resolved away apart
  from its text.

`cut` no longer handles references itself. Check 54's case 28 is Word's markup now but for the reference run's
`w:sz`, which Word's save wrote on the file's own run (as check 52's case 1). A sweep of the 265 files with comments
that load: the paragraph holding a comment's reference had its content deleted, untracked and tracked. The comment went
in every one, Accept All gave the untracked result and Reject All the original in every one. Before the fix of an
empty span's edges (a paragraph holding nothing but the reference, `invoice.docx`), five of them differed.

## 13. `getOoxml` and `toApiScript` across paragraphs (2026-10-23)

*Jason, 2026-10-23: "the shim refuses getOoxml - fix this please". Implemented the same day.*

**Before.** The shim's `Range.getOoxml` and `toApiScript` of a `Range` made one `w:p`, so a range across paragraphs
refused with `NotSupportedError`. The message still read "until CR-014 phase 2", though phase 2 had not changed it
(section 9.6).

**Now** (`fragment.mts`, `blocksOfRange`):
- A range's blocks are detached copies: the first paragraph from the range's start, every block between whole (a
  table, a content control), and the last paragraph to the range's end. The last is left out where the range ends at
  its start.
- The blocks are the ones `delete()` takes (`extentOf`). So a range starting or ending inside a table or a content
  control it doesn't hold whole is refused, with `NotSupportedError`: a fragment of part of a table isn't written.
  What Word's `getOoxml` gives there isn't measured.
- `getOoxml` gives the blocks' XML one after another, the shim's bare fragment form (Office JS's is a `pkg:package`,
  as before), which `insertXml` takes back. `toApiScript` writes each block as it writes a body's.

**A defect fixed with it, since CR-002 phase I.** The copy was trimmed through `Paragraph.splice` on a view in the
range's own body. Under tracking that splice is tracked, so `getOoxml` of one word gave the paragraph's other words as
`w:del` deletions, and `toApiScript` wrote them through its `insertXml` fallback. The copy is now trimmed through a
body that reads as the range's body does (its package's styles decide hidden text, and so the offsets) and writes
untracked. What the trim queues on the package, a picture's relationship or a comment whose reference the copy lost,
is settled at save against the document, which still holds them.

**Tests.** `office-js.test.mjs`:
- within one paragraph, and across a table, the blocks' text and order;
- the fragment back through `insertXml`;
- a range into a table refused;
- under tracking, no `w:del` and the document unchanged;
- `toApiScript` over the same range.

`ranges.test.mjs`: check 52's `HEAD3B` to `BODY3A`, where it asserted the refusal, now gives the two pieces.


## 14. A range holding one block whole deleted the block after it too (2026-10-23)

**Found by** docx4j-python's Phase C port (its CR-013 section 14.5 item 5), which read this package's code at `906d5e3` and
hit the same shape in its own. Confirmed here: `table.getRange('Whole').delete()` on a table of two cells, with a
paragraph after it, removed the table and that paragraph. A one-cell table's whole range is within one paragraph and
doesn't come here.

**Why.** A range holding one block whole, with no paragraph at either end, has `first === last` and no head or tail.
`deleteUntracked` removed the last block (`removeBlock(…, last)`), then the first, at the same index, which by then
held the block after it. Tracked, `deleteTracked` deleted the same block twice. That did nothing visible for a table,
since a row already deleted is left as it is, but a control would have taken Word's markers twice. A control's whole
range isn't this shape: its paragraphs are the head and the tail, and the delete takes its content.

**Fixed:** the block is removed, or marked deleted, once. `ranges.test.mjs` deletes a two-cell table's whole range
untracked (the paragraphs either side stay) and tracked (one change; Accept All takes the table alone). The untracked
half fails on `906d5e3`'s build.

## 15. The matchers across marks exported, for the editor (2026-10-26)

**Asked by** the editor (its ED-008 section 8.8, install item 11). Its Finder is to follow Word's Find across paragraph
marks, as `Body.search` and `Range.search` have since 0.6.0: `^p` in a plain query (section 11) and wildcards
(CR-002 section 32, K7). The Finder searches the editor's own edit layer, not this package's tree, so it needs the
two matchers over its own segments. 0.6.0 had them in `search.mts`, but the content index re-exported only
`searchPattern`, `findAll`, `matchesOf`, `expandReplacement` and `officeMatches`, and the exports map allows no deep
import. Until a release exports them, the editor carries a verbatim copy, under `UPSTREAM-24`.

**Exported** from `model/content/index.mts`, and so from `.` and `./model`, unchanged, as `officeMatches` is:
- `wildcardMatchesAcross(pieces, query, options)`. Each piece is a paragraph's `MarkedSegments`: its segments in
  `officeMatches`' `SearchSegment` shape, the `[from, to)` of its text that is read, its text's length, and whether its
  mark follows. Each hit is a `MarkedMatch`, from `start` in piece `first` to `end` in piece `last`. It holds the mark
  where `mark` is set.
- `paragraphMarkPieces(query)`: a plain query split at `^p` (`^^` stays an escaped caret), or `undefined` where the
  query holds no `^p`.

The caller still chooses which paragraphs a hit may span and which marks it may hold. That is `Range.searchWildcards`'s
rule, which isn't exported: one cell's paragraphs or none's, and a paragraph's own mark, never a cell's or a row's end.

**Tests.** `readers.test.mjs` checks that both entry points export the matchers. It runs `wildcardMatchesAcross` over
segments built by hand in the shape a caller passes, and holds it to `Body.search` over the same paragraphs.
`test/nodenext/consumer.mts` uses the two functions and their types through the package's name. 1,608 tests pass.
