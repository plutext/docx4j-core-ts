# CR-008: Theme colours - the merge rules docx4j CR-034 changed, and its colour resolver

**Status:** Proposed 2026-10-12, on docx4j CR-034's hand-off (the docx4j session, 2026-10-12).
Phase 1 (the merge rules) implemented 2026-10-12, at Jason's request relayed by the docx4j-python
session (section 6), **released in 0.5.0** (2026-10-19). Phase 2 (the resolver) implemented 2026-10-24 at
Jason's request, the whole package with `Font.color` drawn (section 7; D1 and D2 decided the same day), **released in 0.6.0** (2026-10-25).

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

**Follow-up handed off 2026-10-13: luminance rounding** (docx4j `691620fb7`, pushed; CR-034
section 11, "Follow-up: luminance rounding"). A port of the resolver takes this rule with it. It
applies where a colour's last transform that is not an alpha transform is `lum`, `lumOff` or
`lumMod`. For such a colour, `DrawingMLColors.transform` and `ColorMath` now:

- compute `lumMod` as `l * val / 100000.0`, in that order;
- convert back with the Foley-van Dam form (`ColorMath.hslToRgbFoley`). Here `m2` is
  `l <= 0.5 ? l*(1+s) : l + s - l*s` and `m1` is `2l - m2`. Each channel is `hue(m1, m2, h + 1/3 | h |
  h - 1/3)`, where `hue` first takes `h -= floor(h)` and then returns
  `h < 1/6 ? m1+(m2-m1)*6h : h < 1/2 ? m2 : h < 2/3 ? m1+(m2-m1)*(2/3-h)*6 : m1`;
- round each channel half down, `ceil(clamp01(v)*255 - 0.5)` (`ColorMath.channelHalfDown`).

Every other transform still converts with `hslToRgb` and rounds half up. `WordColors` and
`SpreadsheetColors` are unchanged.

The exact form matters because luminance results often land on exactly half a unit, and Office
rounds those halves both ways. With these expressions in this order, the direction follows the
floating-point error, which predicts 29 of the 30 halves in CR-034's P3 readings. A port must keep
the same double arithmetic and the same order of operations, with no algebraic simplification and
no early rounding. JS numbers are IEEE doubles like Java's, so the same expressions in the same
order should give the same bits.

`readings.tsv` is unchanged. Against it, the luminance and picker cases of P3
(`tc-p3-transforms.docx` and `.pptx`) are exact in 97 of 98 in each host, up from 85. The one
still off is T239 (`tx2` with `lumMod` 25000), whose red computes as 3.5000000000000018 where Office
gives 3. docx4j's `ThemeColourProbesTest` holds the luminance families to at most one case off
(`Tally.luminanceOff() <= 1`), the bar a vector test here would take. A case where it shows: an
Excel chart style's `tx1` text with `lumMod` 50000 and `lumOff` 50000 is `7F7F7F`, not `808080`.

**Measured since** (check 53, 2026-10-22; CR-002 section 59): Office JS's `font.color` on Heading 1 resolves its
`w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"` through the theme, the created package's Office
2023 theme (accent1 `156082`), to `#0F4761`. This package's `Font.color` reads `#365F91`, the stored `w:val`. That
is the case this phase's resolver answers.

## 4. Noted, not in scope: a hyperlinked TOC's entries

Coming in docx4j, not yet committed: Word draws the entries of a hyperlinked TOC (runs in a
`TOC \h` field result) without the `Hyperlink` character style at all - colour, underline, weight
and size. Measured on the corpus goldens, about 1,780 such entries, none in the Hyperlink colour;
corpus document 373 shows the weight. docx4j's outputs will drop the style there in a preprocess
step (`TocHyperlinkStyle`). Relevant here only to an effective read of such a run (`Font`'s reads
through the `PropertyResolver` would report the style's colour where Word draws none); nothing is
proposed until a consumer describes TOC formatting.

## 5. Decisions

- **2026-10-12: phase 1 scheduled**, at Jason's request, relayed by the docx4j-python session, which
  ported the same rules the same day (its CR-009 phase A, `e0b68b9` and `a849af9`) and reads this
  repository's goldens from the sibling checkout. It asked for the goldens at docx4j `2e381ef26` or
  later, regenerated in a commit of their own.

- **2026-10-24: phase 2 scheduled** (Jason: "implement CR-008 phase 2"), and two questions decided the
  same day, both as recommended:
  - **D1, scope:** the whole of `org.docx4j.model.color` as section 3 lists it, all three hosts and
    `StyleFill`, held to all 1,068 vectors and the corpus triples; not the Word half first.
  - **D2, the direct view:** `getFont({ direct: true }).color` keeps `w:val` as written; only the
    effective view reads a theme colour through the theme.

## 6. Phase 1 as implemented (2026-10-12)

`catalogue.mts` follows `StyleUtil` at docx4j `2e381ef26` (`VERSION_17_3_3`, local and not pushed
when this was written):

- `isEmptyColor` and `isEmptyU` count `themeColor`, `themeTint` and `themeShade`, as docx4j's
  `isEmpty(Color)` and `isEmpty(U)` do (`themeColor` counts unless absent, as in the Java, where
  `statesColour` below treats `none` as not stated).
- `statesColour(value, theme)`: the value is not empty, or the theme reference is neither absent
  nor `none`. Where it holds, `mergeColor`, `mergeU` (its colour; `w:val` still on its own),
  `mergeShd` (the pattern colour group and the fill group, each on its own) and `mergeBorder` (its
  colour) replace value, reference, tint and shade together, absent ones included. Otherwise only
  a stated tint or shade is taken. The per-attribute `themeColor` helper is gone.
- docx4j's `areEqual(Color)` and `areEqual(U)` changed too, and have no counterpart here (the
  `areEqual` overloads are not ported, `styleUtil.mts`).

Tests: docx4j's three `StyleUtilMergeRulesTest` cases in `test/properties.test.mjs`, less their
`areEqual` assertions, plus CR-015's rule that a direct `w:shd` stating no fill leaves an inherited
`w:themeFill`. All three fail on the catalogue before the port.

**The goldens** were regenerated at `2e381ef26` (built from a `git archive` export with a scratch
Maven repository; jar `docx4j-core-17.3.3-SNAPSHOT.jar`, SHA-256 `f805ae4be89897e3…`), all 62 and
the 62 XPath goldens, and taken in full, in a commit of their own. Header fields aside, two
answers moved, the two docx4j-python had measured: `fonts-modesOfApplication`, main[5] run 3,
whose `effectiveRPr` colour (`SubtleReference`'s plain `accent2` `C0504D`) no longer carries
`Heading1`'s `themeShade` `BF`; and `invoice2013`, table 0, whose `effectiveTableStyle` borders
take the table's own `color="auto"` in place of `TableGrid`'s `themeColor` `text1` on all six
edges. No XPath hit moved. Parity is zero on all 124. Between `856afc5da` and `2e381ef26` the
classes ported here changed only by this commit, issue 593 (ported, CR-001 section 22) and
`8928aacc5` (line endings), and nothing after `2e381ef26` on the local branch touches them.

**A departure from section 2:** that section said to regenerate once docx4j pushes. The goldens are
at an unpushed commit because docx4j-python asked for them now. As in CR-007 section 7, until
docx4j pushes `2e381ef26` the weekly workflow measures the pushed `VERSION_17_3_3`, which has the
attribute-by-attribute merge, and will open a pull request moving these two answers back. Close it
unmerged, or push docx4j first.

Later on 2026-10-12 docx4j pushed `VERSION_17_3_3` to `88ce2aacb`, which contains `2e381ef26`, so the
weekly run will not open that pull request. The thirteen commits after `2e381ef26` don't touch
the classes ported here. They are charts (CR-035 phases 2 to 4, CR-036 phase 1), the
`xlsx4j` cell writers, output-side work (HTML, and section 4's TOC hyperlink style as an output
preprocess step, `c24c676ad`), and line endings in `EmbeddedPackagePart`.

## 7. Phase 2 as implemented (2026-10-24)

Ported from docx4j `VERSION_17_3_5` at `1fffc20a3`, where `org.docx4j.model.color` is unchanged since
`691620fb7` (the luminance rounding, 2026-10-13), class for class, under `src/model/color/` and exported from
`.` and `./model`: `ResolvedColor`, `ThemeColors`, `DrawingMLColors`, `WordColors`, `SpreadsheetColors`,
`StyleFill`, with `ColorMath` and `ColorTables` internal as in docx4j.

**The arithmetic** is docx4j's expression for expression, in the same order (Java's integer division is
`Math.trunc` of the quotient; nothing simplified), as section 3's follow-up asks. JS numbers are IEEE doubles,
and the result bears that out: every probe's tally, family by family, is the one docx4j's own
`ThemeColourProbesTest` prints at `1fffc20a3` (section 7.2).

### 7.1 Names and departures

docx4j's names, with the CLAUDE.md naming rule's changes:

- getters are properties (`ResolvedColor.red`, `alpha`, `isOpaque`; `StyleFill.kind`, `color`, `fill`;
  `ThemeColors.formatScheme`, `placeholder`), and docx4j's `null` is `undefined`. `StyleFill.Kind` is its
  names as strings (`'SOLID'`, ...), as the text model's enums are.
- `ThemeColors.of(Theme)` is `ThemeColors.ofTheme(theme, defaultTheme?)`, synchronous. `of(WordprocessingMLPackage)`,
  `of(SpreadsheetMLPackage)` and `of(Part)` are one `await ThemeColors.of(pkgOrPart)`, asynchronous because it
  reads parts, with `readContents()`, so a theme or slide nobody edited still saves byte for byte. `forChart(chart)`
  is asynchronous for the same reason. A themeless package's theme is the bundled one `pkg.fonts.defaultTheme`
  names (docx4j's `docx4j.fonts.defaultTheme`); parsing it needs the objects context, which
  `ThemeColors.loadDefaultTheme()` builds, so `ofTheme(undefined)` works synchronously once anything has been read.
- docx4j's overloads by type become names: `WordColors.of(Color)`, `of(CTBorder)`, `of(U)` are `of`, `ofBorder`,
  `ofUnderline` (the three objects can't be told apart by shape), and `PropertyResolver.resolveColor(Color)`,
  `(CTBorder)`, `(U)` are `resolveColor`, `resolveBorderColor`, `resolveUnderlineColor`; `resolveFill`,
  `resolvePatternColor` and `getThemeColors` keep their names.
- `WordColors.luminance` and `SpreadsheetColors.tint`, package-private in docx4j, are public here: the corpus
  test calls the first, as docx4j's test does.
- `DrawingMLColors` knows an object by its `TYPE_NAME` (which unmarshalled and factory-built objects carry), by its
  element name where it comes as `{ name, value }` (a `JAXBElement`'s counterpart), and a holder of a colour (a
  scheme slot, a solid fill, a style or font reference) also by its `srgbClr`, `schemeClr`, ... member. docx4j
  uses `instanceof`, which has no counterpart for plain objects.

### 7.2 The PropertyResolver and `Font.color`

`PropertyResolver.create` reads the theme part privately, as it reads the styles, numbering and settings parts,
and parses the default theme the package names, so `getThemeColors()` is synchronous: the theme part (or the
default theme) with the settings' `w:clrSchemeMapping`, made once per `refresh()`. `ResolverSource` gains `theme`
and `defaultTheme`. The effective properties themselves are unchanged (they keep the attributes as written, as
docx4j's do), so no parity golden moves.

`Font.color` in the effective view reads a run's effective `w:color` through `WordColors`: where it has a theme
reference, the colour Word draws (Word's own value where the `w:val` beside it is within one unit, a current
cache; else the computation). An automatic colour still reads `#000000` (CR-002 section 59), and a value with no
theme reference reads as before. The direct view is unchanged (D2). A range across paragraphs reads the same way.
So check 53's cases 3 to 5, Heading 1's `w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"` over the
created package's Office 2023 theme, read `#0F4761`, as Word 365 does; `ranges.test.mjs` no longer lists them
among check 53's differences, and asserts the colour. Nothing else in the content API reads a colour Office JS
reports that a theme can give (the subset has no shading, border or underline colour reads).

### 7.3 Tests

`test/theme-colors/` is docx4j's `docx4j-core-tests/src/test/resources/theme-colors/`, copied unedited at
`1fffc20a3` (added in `3dbe4eee5`, Apache-2.0), with a `manifest.json` of SHA-1s; outside `fixtures/`, so no
fixture-wide test or harness counts its probes. `test/theme-colors.test.mjs` (18 tests):

- docx4j's `ThemeColourProbesTest`: P3 in Word and in PowerPoint (403 cases each), P1, P2, P5, P4 and the corpus
  triples, at docx4j's bar (exact, or within one unit where CR-034 records it; the luminance families exact but
  T239; the saturation set on a grey the one family off) and, beyond it, **docx4j's tallies family by family**,
  pinned from its run at `1fffc20a3`: P3 the same 40 families in both hosts (`lumMod` 16 exact and 1 within one,
  `sat` 11, 2 and 2, ...), P1 exact but P1-10, P2, P4 and P5 exact, the corpus 76 exact and 41 within one of 117.
- docx4j's `ThemeColorsTest`: a new docx has the Office 2023 theme; a themeless package takes the default theme
  its `fonts.defaultTheme` names (2013's and 2007's both `4472C4`, CR-034 D2); the sheet indexes; a chart drawn
  in its host's theme, with its `c:clrMapOvr` and its theme override part (built in memory: docx4j's
  `charts/bar-chart.docx` is not among the fixtures); a slide, layout and master, and a part of a docx.
- docx4j's `ThemeColourOutputTest`, its resolver cases: stale cache, a direct value, a theme reference with no value,
  a current cache, `auto` beside a theme colour, shading, a border, an underline, made once, read again on refresh.
- `Font.color`: the same document through the content API, effective and direct; and every P1 and P2 paragraph's
  text read through `Body` against what Word drew (17 and 6; P2's `w:clrSchemeMapping` included). A themeless
  package's resolver at the 2013 default theme.

### 7.4 Findings

1. **hexBinary is declared `string` and unmarshalled as bytes.** `a:sysClr/@lastClr` and SpreadsheetML's `rgb`
   are `[0, 0, 0]` and `[255, 217, 87, 87]` at run time, where the objects package's declarations say `string`:
   jsonix-schema-compiler maps hexBinary to `string` (its CR-005 table), and jsonix's `HexBinary` type
   unmarshals to an array of bytes. The resolver reads either (`bytesOf`), and a test asserts the bytes, so a
   change on either side shows. Not raised with the compiler or the objects package yet; Jason decides.
2. docx4j's `WordColors.luminance` comment says "all 119 distinct ones in four corpora, 77 exact"; its vectors
   hold 117 and its own test prints 76 exact and 41 within one, which is what is pinned here.
3. docx4j's P2 test counts 21 "run text" readings: its loop takes every `w:p` whose text, text boxes included,
   starts `P2-`, which is the six runs, the paragraph anchoring the shapes (its text starts with `P2-07`), and the
   seven shapes' label paragraphs in each branch of their `mc:AlternateContent`. The probe's runs are the six the
   `Font.color` test counts; not a defect.

Full suite: 1,515 tests pass (1,497 before), typecheck clean.
