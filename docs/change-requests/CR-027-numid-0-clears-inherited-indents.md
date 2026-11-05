# CR-027: `w:numId 0` clears every inherited left, hanging and first-line indent, docx4j's rule of 2026-11-05

**Status:** Proposed 2026-11-06, at Jason's "w:numId 0: CR", on the portfolio session's hand-off of docx4j `9e3100351` (`VERSION_17_3_6`, local and
unpushed when this was written). Measured here 2026-11-05 and again 2026-11-06 (section 1). Docs only, nothing implemented. D1 open.

## 0. Summary

Where a paragraph or a style turns numbering off with `w:numId w:val="0"`, the `PropertyResolver` reads the indent wrong. It drops only what equals the
level's indent, so a numbered style's own `w:ind` stays. Word 365 clears every inherited left, hanging and first-line indent and keeps the right one.
docx4j took that rule on 2026-11-05 from Word's goldens on its probe `numbering-off-style-ind`, and this package still holds the rule docx4j withdrew.

The fix is a few lines in `styleUtil.mts`. No golden moves, because no fixture has the shape, so parity would never have flagged it. docx4j's own test,
mirrored here, has to hold it.

## 1. Here today (measured on `33bae48`'s `src/`)

docx4j's `NumIdZeroIndentTest` cases through `getEffectivePPr`, in twips. The level is left 1440, hanging 1152. `OwnInd` is a numbered style stating left
454, hanging 454, right 57. `Switch` is a style based on it holding `w:numId 0`. `OnBase` is a numbered style based on one stating left 720, hanging 360.

| Case | Shape | docx4j's rule | Here |
|---|---|---|---|
| N1 | `OwnInd`, direct `w:numId 0` | right 57 | **Wrong:** left 454, hanging 454, right 57 |
| N4 | `OnBase`, direct `w:numId 0` | nothing | Right |
| N7 | `Switch` | right 57 | **Wrong:** left 454, hanging 454, right 57 |
| N8 | `Switch` with a `w:ind` of its own, left 300, hanging 100 | left 300, hanging 100, right 57 | Right |
| N10 | N1 with a direct left 300 | left 300, right 57 | **Wrong:** left 300, hanging 454, right 57 |
| N11 | N1 with a direct hanging 200 | hanging 200, right 57 | **Wrong:** left 454, hanging 200, right 57 |
| N14 | `Switch` with a direct left 142 | left 142, right 57 | **Wrong:** left 142, hanging 454, right 57 |

The test's two earlier cases read right here: a numbered paragraph, and `w:numId 0` with a direct left 720. Python's run of the seven gave the same five
wrong (the portfolio's hand-off). On a scratch copy of `dist` patched with section 2's rule, all nine read right.

Every effective read of a paragraph's indent goes through this: `Paragraph.leftIndent`, `firstLineIndent` and `formatting`, and whatever else calls `getEffectivePPr`.

The rule rests on docx4j's measurement, in `docx4j-export-fo/docs/word-layout-rules.md` at anchor `s28numid0`: Word's drawn positions on the probe, in
modes 12 and 15 alike. Word wasn't run for this CR. That document leaves one shape unmeasured: a level with no `w:ind` of its own.

## 2. Proposal

In `src/model/properties/styleUtil.mts`, as docx4j's `StyleUtil`:

- `dropNumberingInd` goes, and with it the test of each component against the level's. In its place a function deletes `left`, `leftChars`, `start`,
  `startChars`, `hanging`, `hangingChars`, `firstLine` and `firstLineChars`, and leaves the rest.
- So a right indent that equals the level's now stays, where the old rule dropped it.
- The branch in `applyPPrBase` stops asking `numbering.getInd`. It asks whether the inherited `w:numPr` names a list that exists, through
  `numbering.instanceListDefinitions.get(String(numId))`. Where it names none the rule doesn't apply, as in docx4j's document. A level with no `w:ind` no
  longer stops the branch, which follows docx4j on the shape section 1 calls unmeasured.
- `NumberingIndents` gains `instanceListDefinitions`, which both its implementors already have: `NumberingDefinitions` and `NumberingDefinitionsPart`.
- The comment saying a `w:ind` the style states itself survives is the withdrawn reading, and it goes.

Only what is inherited so far clears. A `w:ind` stated with the `w:numId 0` or below it still applies, because the merge follows the branch.

`dropLevelIndent` in `TrackedChange.mts` is a different mechanism and doesn't change.

## 3. Tests and oracles

- `test/numid-zero-indent.test.mjs`: docx4j's `NumIdZeroIndentTest` mirrored, its styles and numbering built in memory, all nine cases through
  `getEffectivePPr` and through `Paragraph.leftIndent`.
- Beside them: an inherited `w:numPr` naming no list leaves the indent alone, the `*Chars` forms clear, `w:right` and `w:end` stay, and a level with no `w:ind` clears as docx4j's does.
- Parity: nothing moves. Over the full suite on the patched copy the branch ran 11 times, and the two rules agreed all 11. So a fixture of this
  shape joins the goldens when `test/java/pom.xml` moves to the docx4j release holding `9e3100351`, and docx4j's own answer then holds it. Until then the
  pom resolves 17.3.5, which answers by the old rule.

## 4. Phases and size

One phase, ¼ day. It waits on nothing here. The rule is in docx4j's history, though that commit is unpushed.

## 5. Decisions

| | Question | Recommendation |
|---|---|---|
| D1 | Which release? | With 0.7.0 if built before it. It's small, it moves no golden, and it corrects a wrong read. It shouldn't hold 0.7.0 up |

## 6. Registry

`core-ts/CR-027`, proposed, no dependencies. Sent on commit.

## 7. Sources

- docx4j `9e3100351`: `StyleUtil.java`, `NumIdZeroIndentTest.java`, `docx4j-export-fo/docs/word-layout-rules.md` at `s28numid0`, `CHANGELOG.md`.
- The portfolio session's hand-off of 2026-11-05.
- `src/model/properties/styleUtil.mts`, `src/model/listnumbering/definitions.mts`, `src/parts/wml/index.mts`.
