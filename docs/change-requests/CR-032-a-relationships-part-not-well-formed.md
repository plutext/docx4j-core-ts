# CR-032: A relationships part that isn't well formed, read as Word's repair reads it (docx4j CR-075 D5, D6 and D8)

**Status:** Decided and implemented 2026-11-10 (Jason: "commit, push, then fix the Two things 17.3.6 does that this package doesn't
yet"; for this one, "docx4j's reading (Recommended)"). Unreleased. The port of docx4j CR-075's phase 4, released in docx4j
17.3.6 (`e2d51d1e2`), which CR-001 section 27 found unported.
**Counterparts:** docx4j CR-075 D5, D6 and D8 (17.3.6); WD 0.76's ruling on WD 0.73's item (i) (portfolio `7c5ee35`); CR-031 here,
whose `repair` decides the save; CR-025 here, whose `SaveLossException` and `checkLossOnSave` the save's gate shares. docx4j-python
refused such a part when docx4j's D5 was written (its CR-031); whether it still does is its own.

## 0. Summary

A relationships part that isn't well formed refused the whole load here: `word/_rels/document.xml.rels` cut off after its third
relationship gave `PartUnmarshalException: … Malformed XML`, and nothing of the document could be read. WD 0.76 ruled that this
conforms, and so does the other reading: the relationships parsed before the error, `part.not_read` naming the relationships part
with the count kept, and the parts only a lost relationship named read as absent. That is Word's repair, and docx4j 17.3.6 reads it
so at every value of its `RepairMode`, refusing a save unless the mode follows Word's. Jason took docx4j's reading for this package.

Now, at every value of `repair`:
- a relationships part that isn't well formed is read short: each whole `Relationship` before the break, as a part in the right
  namespace, with a `RelationshipsNotReadWarning` (`part.not_read`, the part's name, the count kept, the parse error and its
  position) on the part, so on `pkg.warnings` and `pkg.relationshipsNotRead`, and in the text model as `part.not_read`, the part's
  name the selector and the count the detail;
- a part only a lost relationship named isn't read, and its content-type Override goes;
- `_rels/.rels` that lost the main part's relationship so is refused, `PartUnmarshalException` naming the part and the parse error.

A save of such a package is refused, `SaveLossException` with a `relationships` loss naming the part and the count kept, unless the
load's `repair` is `'word'` or `'all'`, when it writes the repaired form as Word's save does, or `pkg.checkLossOnSave` is false.
Either way `lastSave.relationshipsNotRead` lists the parts written so.

## 1. What Word and docx4j do

docx4j CR-075's table, rows 5 and 6, and its broken-part family's README (Word 365):

| Shape | Word 365 | docx4j 17.3.5 | docx4j 17.3.6 | Here before |
|---|---|---|---|---|
| A part's relationships part not well formed (`broken-part-document-rels`: cut after its third relationship) | "Unreadable content", repaired "Comments": the three relationships before the break kept, the comments' lost, their references left as empty comments | No relationships, no warning, a stack trace; the save dropped styles, numbering, settings and the comments | The three kept, `part.not_read`, the save refused unless `WORD` | Refused at the load |
| The package's `_rels/.rels` not well formed (`broken-part-package-rels`: cut after its first relationship) | Repaired: the main part kept, the three properties parts lost | Refused, its cause a `NullPointerException` | The main part's relationship kept, `part.not_read`; refused naming the part where the main part's is lost | Refused at the load |

## 2. Decisions

| | Question | Decided |
|---|---|---|
| D1 | How such a part reads | **At every value of `repair`, Word's reading** (Jason, 2026-11-10: "docx4j's reading (Recommended)"). The other choice was CR-024's pattern: refused under `'refuse'`, read only under `'word'` and `'all'`. CR-031 section 2.1.1 allows either, since the read is the same at every value; a read-only use now gets nearly all of such a file, as the repair policy's P2 has it |
| D2 | The save | **Refused unless the load's `repair` is `'word'` or `'all'`, or `checkLossOnSave` is false.** docx4j lifts its gate under `WORD` alone. Here `checkLossOnSave` lifts it too, being this package's switch for a save that loses what a read didn't keep (CR-025); `repair` deciding the save as well is what CR-024's shapes do under `'word'`, where the save writes the repaired package. CR-031 D6 stands: `repair` never changes `checkLossOnSave`. The refusal is CR-025's `SaveLossException`, `save.content_lost`, a `ContentLoss` of kind `relationships` |
| D3 | `_rels/.rels` | As docx4j's D6: read the same way, and refused, naming `_rels/.rels` and the parse error, where the main part's relationship isn't among those kept, at every value |
| D4 | The lost parts' Overrides | Removed after the load's walk, where any relationships part was read short: each Override naming a part the load didn't reach. A save then writes the package Word's save of its repair writes, as CR-024 settles a dropped main part's Override |
| D5 | What counts as not well formed | The parser's. xmldom refuses a duplicate attribute, an unescaped `<` in a value, a mismatched or missing end tag, and reads through an undefined entity, an unquoted value and a lone `&`, which docx4j's StAX reader stops at. Not changed here: it is how every part has always parsed |
| D6 | `clone()` | Refused as a save is, as CR-025 refuses it, since it goes through one. docx4j spares `clone()`, as a copy and not a save; here a clone that wrote the short read would carry no warning, and its later save no gate |
| D7 | `repaired` | `none` for a short read alone, as docx4j's `LoadFinding.repaired` counts it: the reading is the same at every value, not a repair the option makes |
| D8 | A caller who puts the relationships back | Says so with `setContents` on the relationships part, which clears its warnings, so the save goes. docx4j's caller says so with `setRepairMode` |

## 3. As built (2026-11-10)

- `src/opc/repair/relationshipsBeforeTheBreak.mts`: the reading. A DOM parser has no streaming read, so the part's text is cut after
  each child of the root, found by a scan that knows quotes, comments, processing instructions and CDATA, and the longest cut that
  parses closed with the root's end tag holds the relationships before the break. The parser decides, never the scan. The last cut
  is tried first, the one more parse a part cut off takes; then the last cut at the parse error's position (xmldom's locator puts
  it on the character it stopped at); then a binary search. As docx4j has it, a `Relationship` is taken by its local name in any
  namespace, kept with `Id`, `Type`, `Target` and `TargetMode`, one without any of the first three passed over, and returned as a
  part in the right namespace.
- `XmlPart.readMalformed`, a hook in `parse`: undefined by default, so every other part still throws. `RelationshipsPart` overrides
  it, keeps the parse error on `readFailure` and says so first in `readRoot`, `notRead` the warning.
- `Load.mts`: `_rels/.rels`'s refusal (D3), the Overrides (D4). `OpcPackage.relationshipsNotRead`.
- `Save.mts`: the gate in `settleLosses` (D2), `SaveReport.relationshipsNotRead`; `ContentLoss.kind` `relationships`, and
  `SaveLossException`'s message saying the way past for each kind.
- The text model: `TextPackage.relationshipsNotRead`, and `part.not_read` with the count as its detail.

On docx4j's two files every value of `repair` reads what docx4j 17.3.6 reads: the three relationships of `document-rels`, the main
part's alone of `package-rels`.

**Not measured:** whether Word opens this package's save of the repaired form with no message. Word's own save rebuilds what its
repair lost: a comments part holding empty comments for the references left in the body, and its own properties parts. This
package's save, like docx4j's, writes the package without the lost parts and leaves the body's comment references as they are. A
Word check of the two saves would settle it.

## 4. Tests

`test/rels-not-read.test.mjs`, ten tests, over `test/rels-not-read/`: docx4j's broken-part files 100, 119 and 120 from the 17.3.6
release with Word's saves and texts, and the family's two tables, copied unedited with a SHA-1 manifest.
- Each file at each value of `repair`: the part and the count, the reason and its position, on `pkg.warnings`, `repaired` none, the
  parts the intact twin's less what only a lost relationship named, and no Override for them.
- docx4j's tables, cases 119 and 120: every token in the three views by the suite's own rule (`conformance.mjs`'s `verify`, the
  adapter's reading), 84 checks, and `part.not_read` naming the part in each. The `with_option` column held to Word's Save As text,
  and this package's `profile: word-text` with list labels held to it on every row that isn't about a field (`fields: update` isn't
  implemented here, CR-012 phase 5a).
- The save refused at the default, naming the part and the count and the way past, nothing written, and `clone()` refused; under
  `'word'`, `'all'` and `checkLossOnSave` off the same save, `lastSave.relationshipsNotRead`, the save read back at the default with
  nothing said, the lost parts and their Overrides not written, the main part copied byte for byte, and the same text as the short
  read; `setContents` lifting the gate.
- `_rels/.rels` cut before the main part's relationship, first and second, refused at every value.
- The shapes: a cut inside the root's start tag (nothing kept); a duplicate attribute in the second relationship, and an unescaped
  `<` in the third's value; a `Relationship` without a `Target`; a comment holding a `Relationship`'s text and a value holding `>`;
  what follows the root's end; relationships in another namespace before a break, read at the default and saved in the right one;
  a part that parses whole.
- The text model's warning, its selector and detail.

## 5. Registry

`core-ts/CR-032`, done, depending on nothing in another repository (docx4j's CR-075 phase 4 is its source, released). Sent to the
portfolio session with the commit's hash.

## 6. Sources

- docx4j `e2d51d1e2` (CR-075 phases 2 and 4), at the 17.3.6 release, tag `docx4j-17.3.6`, `4afcf4be8`:
  `RelationshipsPart.readRelationshipsBeforeTheBreak`, `Load3`'s two callers, `OpcPackage.refuseSaveThatLoses`, `LoadFinding`.
- docx4j's `docs/developer/change-requests/CR-075-packages-word-repairs-at-load.md`, sections 0, 3 (WD 0.76's ruling), 5 (D5, D6,
  D8) and 6 (phase 4).
- `docx4j-core-tests/src/test/resources/text-extraction/broken-part/README.md`, its Word table.
- CR-001 section 27, which found it unported.
