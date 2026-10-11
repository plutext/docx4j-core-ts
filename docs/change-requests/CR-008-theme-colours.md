# CR-008: Theme colours - the merge rules docx4j CR-034 changed, and its colour resolver

**Status:** Proposed 2026-10-12, on docx4j CR-034's hand-off (the docx4j session, 2026-10-12). Not
scheduled; Jason decides. Phase 1 (the merge rules) is small and parity-bound; phase 2 (the
resolver) is new scope.

## 1. Why

docx4j CR-034 (`docs/developer/change-requests/CR-034-theme-colours-resolved.md`, section 11 for
Office's readings) changed `StyleUtil`'s merge rules for colours and added `org.docx4j.model.color`,
a resolver from a theme reference, tint and shade to the colour Word, PowerPoint or Excel draws. It
is on docx4j's `VERSION_17_3_3`, not pushed when handed off: phase 1 `3dbe4eee5`
(`org.docx4j.model.color`), phase 2 `2e381ef26` (the outputs and `StyleUtil`).

This package mirrors `StyleUtil` in `src/model/properties/catalogue.mts` (CR-001 Phase B), and the
parity goldens are docx4j's answers: once docx4j pushes the branch, the weekly parity run turns
every golden whose effective properties hold a theme colour (`fonts-theme-lang` has fifteen, among
others) and opens a pull request. CR-001's rule is that the port follows unless the Java is plainly
wrong, and here the Java follows Word (corpus document 3229 draws the runs black).

## 2. Phase 1: the merge rules

What `StyleUtil` now does, and what the catalogue does today:

1. **Emptiness counts the theme reference.** `isEmpty` and `areEqual` of `w:color` and `w:u` now
   count `themeColor`, `themeTint` and `themeShade`. A style's `w:color` that states only a theme
   colour is inherited; it used to be dropped as empty. Here `isEmptyColor` reads `w:val` alone and
   `isEmptyU` `w:val` and `w:color` alone (`catalogue.mts`), so such a colour is dropped the same way.
2. **A colour inherits as one.** Where a layer states the value (`w:val`, or `color` / `fill`) or the
   theme reference, its value, reference, tint and shade replace the ones below *together*; a tint
   or shade stated alone still inherits on its own. This applies to `w:color`; to `w:u`'s colour
   (its `w:val` still separately); to `w:shd`'s pattern-colour group (`color`, `themeColor`,
   `themeTint`, `themeShade`) and fill group (`fill`, `themeFill`, `themeFillTint`,
   `themeFillShade`), each its own; and to a border's colour. Why: Word draws the theme reference
   over the value, so attribute by attribute a run's own `w:color w:val="auto"` under a style's
   `accent1` kept the style's reference and drew in the accent. Here `mergeColor`, `mergeU`,
   `mergeShd` and `mergeBorder` merge attribute by attribute (`attr` per attribute, `themeColor`
   for the enum), which is exactly the behaviour CR-034 left. CR-015's rule that a `themeFill`
   survives a direct `w:shd` stating no fill still holds, since the fill group is its own.
3. **Tests in docx4j:** `StyleUtilMergeRulesTest` (three new cases) and `ThemeColourOutputTest`;
   the catalogue's own tests (`test/properties.test.mjs` or wherever the merge rules are held) gain
   the same three, and the parity goldens are regenerated at the docx4j commit once it is pushed
   (`test/java/README.md`'s recipe), with zero differences the target.

Estimate: half a day, including the goldens.

## 3. Phase 2: the resolver

`org.docx4j.model.color`: `ResolvedColor`, `ThemeColors`, `DrawingMLColors`, `WordColors`,
`SpreadsheetColors`, `StyleFill`; its `package-info` summarises it. Test vectors under
`docx4j-core-tests/src/test/resources/theme-colors/` with a `README.md`: `cases.tsv` (1,068 probe
cases in seven probe files), `readings.tsv` (what Word, PowerPoint and Excel drew for each),
`word-tint-shade.tsv` (119 corpus triples: base, operation, byte, Word's value). The rules in brief,
from the hand-off:

- **Word:** the theme reference wins over the value beside it. `themeTint` and `themeShade` work in
  HSL luminance (L·t + 1 − t, and L·t), and the tint wins if both are present. Word's own value is
  returned where it is within one unit of the computation (a current cache).
- **DrawingML:** transforms apply in document order. `tint`, `shade`, `inv`, `gamma` and the channel
  transforms work in linear RGB; `hue`, `sat` and `lum` in HSL, with `sat` not clamped; `gray` is
  Rec. 709 luma. `sysClr` is the machine's colour, not `lastClr`.
- **Excel:** `tint` uses Windows' integer HLS (HLSMAX 240), with L truncated after adding 0.02.
  Theme indexes 0 to 3 are `lt1`, `dk1`, `lt2`, `dk2`.

Where it would sit here: `src/model/color/`, beside `src/model/fonts/` (`ThemeFonts` already reads
the theme part); the vectors as a parity oracle of their own, read by a test as the goldens are.
Consumers: the editor's rendering of theme colours (it asked nothing yet), `Font.color` as an
effective read (today it reports the `w:color` value, which for a theme colour is Word's cached
value or `auto`). Estimate: two days with the vectors; not before a consumer asks.

## 4. Noted, not in scope: a hyperlinked TOC's entries

Coming in docx4j, not yet committed: Word draws the entries of a hyperlinked TOC (runs in a
`TOC \h` field result) without the `Hyperlink` character style at all - colour, underline, weight
and size. Measured on the corpus goldens, about 1,780 such entries, none in the Hyperlink colour;
corpus document 373 shows the weight. docx4j's outputs will drop the style there in a preprocess
step (`TocHyperlinkStyle`). Relevant here only to an effective read of such a run (`Font`'s reads
through the `PropertyResolver` would report the style's colour where Word draws none); nothing is
proposed until a consumer describes TOC formatting.

## 5. Decisions

None yet.
