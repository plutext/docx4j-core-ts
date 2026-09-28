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
