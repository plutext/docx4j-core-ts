# CR-014: Ranges across paragraphs, as Office JS's `Word.Range` is

**Status:** Proposed 2026-10-21, at Jason's request. He asked why a `Range` here is one paragraph's where
Word's is not, and whether we should align. To the answer, open a CR with the editor's review and a Word check
first, and stop the silent partial results meanwhile, he said "yes please". The stopgap went in with CR-012
phase 4: a multi-paragraph block control's range reads as Office JS's and refuses every write (CR-012 section
17.12). Nothing of this CR is implemented. Decisions D1 to D4 (section 5) are open.

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

## 5. Decisions (open)

1. **D1, the shape.** Recommended: one `Range` class with an end paragraph, a single-paragraph range unchanged,
   rather than a second class beside `Range`. Office JS has one `Range`, and two classes would make every caller
   ask which it has.
2. **D2, `Comment.getRange()`.** Office JS returns one `Range`; this package has returned `Range[]` since CR-002
   phase G. Recommended: return one `Range` as Office JS does, and keep the per-paragraph array as
   `getRanges()`. It breaks a caller that indexes the array, which on 0.x is allowed.
3. **D3, tables.** Recommended: a range holding whole tables reads and deletes them. One that starts or ends
   inside a table reads in `body.text`'s form, and its writes throw `NotSupportedError` unless check 52 shows a
   rule simple enough to follow.
4. **D4, the editor.** Its spans are per paragraph, so this is additive for it, but its agent's verbs (`delete`,
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

## 6. Check 52 (to write)

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

## 7. Risks

- **Joins under tracking.** A deleted mark between paragraphs of different properties is where CR-002 sections 29
  and 38 found Word's subtleties. Phase 2 reuses that code rather than writing new joins.
- **The editor's model.** If the editor wants spanning ranges in its verbs, its own span type changes. That is its
  CR, after D4.
- **Offsets.** A range's `start` is an offset in its first paragraph and `end` in its last. Code that compares
  `end` with `paragraph.text.length` assuming one paragraph has to look at `endParagraph`. All of it is in
  `src/model/content/`, and phase 1 audits it.

## 8. References

- CR-002 section 3.4's subset and section 7, the one-paragraph `Range`; phase G's `Comment.getRange`; section 29's
  `TrackedChange.getRange` for groups.
- CR-012 sections 17.11 (check 51) and 17.12 (the `SpanningRange` stopgap).
- `test/README.md` checks 30, 31, 48 and 51.
- Office JS: `Word.Range` (`expandTo`, `getRange`, `paragraphs`, `search` with `^p`), `Word.ContentControl.getRange`,
  `Word.Comment.getRange`.
