# CR-014: Ranges across paragraphs, as Office JS's `Word.Range` is

**Status:** Proposed 2026-10-21, at Jason's request. He asked why a `Range` here is one paragraph's where
Word's is not, and whether we should align. To the answer, open a CR with the editor's review and a Word check
first, and stop the silent partial results meanwhile, he said "yes please". The stopgap went in with CR-012
phase 4: a multi-paragraph block control's range reads as Office JS's and refuses every write (CR-012 section
17.12). Nothing of this CR is implemented. Decisions D1 to D4 (section 5) were decided 2026-10-22 as recommended
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
in both views) and changed to Word's answers where they differed (section 9.7). Not committed.

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
  way isn't measured, and phase 2 asks.
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

*Jason, 2026-10-22: "implement CR-014 phase 1". Implemented the same day; not committed.*

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
- **Every write refuses** with `NotSupportedError` naming phase 2, before anything changes: `text`, `insertText`,
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
- `paragraphElementOfRange` (`toApiScript`, the shim's `getOoxml`), which refuses.

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
