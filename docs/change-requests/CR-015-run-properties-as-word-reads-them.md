# CR-015: Run properties read as Word reads them (on/off, enumerations, colours)

**Status:** Proposed 2026-10-24, drafted at Jason's request relayed by the portfolio session ("yes please" to drafting
one CR each in docx4j, core-ts and docx4j-python, against the same rules and probes; docx4j's is CR-060, python's
CR-018). **D1 to D11 decided 2026-10-24 as recommended** (Jason: "CR-015 D1 to D11 decided as now recommended";
section 7), after probe rounds 5 and 6 and a review against CR-018 (section 10). **Implemented 2026-10-24, phases 0 to
2b** (Jason: "implement CR-015"; section 11), on objects 0.5.0 (its CR-013 phases A to C, published 2026-10-24; the range
raised to `^0.5.0`); committed `76fb45d`. **Phase 3's check 60 run the same day** (Jason; no repair): Office JS reads
every marker as this package does, but Word's green highlight, which it reports `#00FF00` and the objects package's
table `#008000` (section 11.6). docx4j's `82c64bf14` (`w:framePr`'s `anchorLock` read as stated) is ported with
it, and the parity goldens regenerated at docx4j `5163c2986`.

## 1. Why

Jason asked on 2026-10-24 whether `<w:b/>` reads as bold on. It does here, in docx4j and in docx4j-python. He then had
four rounds of probes run through Word 365. Word turned out lenient where this package is strict: it reads a run
property whose value the schema doesn't allow (`w:b w:val="yes"`, `w:u w:val="Single"`, `w:color w:val="red"`), and
this package either refuses the whole document part over it or reports the raw value.

Refusing the part is the worse half. One such attribute anywhere in `document.xml` and `getBody()`, `Texts.of`, every
reader and every edit throw `PartUnmarshalException`. The package loads (parts load lazily), and an untouched save
copies the part back byte for byte, but nothing can be read. Word opens the same file with no repair prompt.

## 2. Word's rules (Word 365, Jason's runs, 2026-10-24)

The probes, Jason's PDFs and Word's saves (`*-word365.docx`) are in `~/vbShares/Office 2016/onoff-fun/`, with a
`README.txt`. They were built by the portfolio session (its `onoff/build.py` to `build4.py`). Each probe paragraph is
`label: MARKER (end)`, and only the marker run carries the property tested. The rules below are the portfolio's
summary of Jason's runs. This session checked each one against Word's saves (marker run by marker run, 2026-10-24);
the saves agree.

1. **On/off (`ST_OnOff`):** `b`, `i`, `caps`, `smallCaps`, `strike`, `dstrike`, `outline`, `shadow`, `emboss`,
   `imprint` and `vanish` were measured, and by extension `bCs`, `iCs` and every other `CT_OnOff`.
   - Off only for exactly `false`, `off` or `0`, case-sensitive and untrimmed.
   - Everything else is on: bare, `true`, `on`, `1`, `""`, `" "`, `TRUE`, `FALSE`, `OFF`, `False`, `Off`, `" false"`,
     `"false "`, `00`, `no`, `yes`. There is no repair prompt.
   - Word saves an on value bare (`<w:b/>`; `onoff-*`, `props-b-case`, `props-toggles-*`). It drops an off value,
     or keeps it as `w:val="0"` where it overrides a style (`props-b-styles` S2 to S4, S6).
   - The toggle rule holds: ParaBold plus CharBold is not bold (S7).
2. **Enumerations (`u`, `highlight`, `vertAlign`, `em`):** a value outside the enumeration is unset; the case is exact.
   - A bare `<w:u/>`, or one with only `w:color`, is no direct underline: over CharUnderline the style's underline
     shows (S11), and with no style there is none (U1, U7). `"none"` over a style removes the underline (S10).
   - An invalid `w:u` (`""`, `SINGLE`, `Single`, `NONE`, `bogus`, `true`, `1`, `on`, `off`, `0`, `false`) opens with
     no repair and no underline. Word saves it dropped.
   - A missing or invalid value on `highlight`, `vertAlign` or `em` (`props-enum-missing`; `props-enum-other` N1 to
     N5: `YELLOW`, `bogus`, `Superscript`, `bogus`, `DOT`) makes Word repair the file, dropping the element.
3. **`w:color`:**
   - Six hex digits, any case, an optional leading `#`, surrounding whitespace trimmed (`colour-hex` X1, X2, X4,
     X9, X10).
   - Or a name from the classic HTML (Netscape) colour table, case-insensitive. Measured from Word's saves: 41
     distinct names, every one at its CSS value (`red` FF0000, `darkGray` A9A9A9, `lightgrey` D3D3D3,
     `lightgoldenrodyellow` FAFAD2, ...).
   - Not names, so automatic (18 measured): `darkYellow` and `lightGray` (the highlight names the table lacks),
     `grey`, `darkgrey`, `dimgrey`, `slategrey`, `lightslategrey`, `darkslategrey` (the table spells grays with an a,
     except `lightgrey`), CSS 4's `rebeccapurple`, `transparent` and `currentcolor`, X11-only `lightgoldenrod`,
     `violetred` and `navyblue`, `"dark red"` (a space), and the system colours `window`, `buttonface` and `windowtext`.
     That fits the CSS 3 (SVG) keyword list, 147 names, less six of its seven `grey` spellings (Word keeps
     `lightgrey`) and `lightgray`: 140 names.
   - Anything else draws automatic with no style beneath: `F00`, `0xFF0000`, five or seven digits, `GG0000`, `""`.
     So do `auto`, `Auto` and `AUTO`. Over a style only exactly `auto` is automatic (round 6, rule 5).
   - Word saves a name or lower-case hex as upper-case hex (`colour-names-*`, `colour-case`, `colour-table-*`).

4. **Over a style (round 5, D5's probe; built and run 2026-10-24, docx4j CR-060 section 12).**
   - **`w:u`:** over a character or paragraph style underlining double, a bare `w:u` and the values `bogus`,
     `Single`, `SINGLE`, `""` and `NONE` read as absent. The style's underline shows, and Word's save drops the
     direct `w:u`. A colour with no valid value gives the style's underline in that colour.
   - **But `w:u w:val="true"` removes the underline:** Word draws none and saves `w:val="none"`. Whether `1`, `on`,
     `false`, `0`, `off` or `TRUE` act as `true` does over a style is unmeasured.
   - **`w:color`:** over a style's colour, `F00`, `bogus`, `""`, a bare `<w:color/>`, `grey`, `0xFF0000` and seven
     digits read as absent: the style's colour shows, and the save drops them. `auto` overrides (black); `Auto` and
     `AUTO` were settled by round 6 (rule 5).
   - **`highlight`, `vertAlign`, `em`:** an invalid or bare value over a style is repaired away and the style's value
     shows. Word drew no highlight from the character style in that (repaired) file, though it drew the paragraph
     style's; whether that is Word's rule or the repair's doing is not known.
   - **Colour names elsewhere:** Word reads names in `w:u/@w:color`, `w:shd/@w:fill` and a border's `w:color`, and
     saves them as hex. An invalid `F00` there: the underline's colour is automatic, the shading is dropped (no
     shading), and the border's colour is automatic (saved `auto`).
5. **Round 6 (run 2026-10-24 with no repairs; docx4j CR-060 section 13).**
   - **`w:u w:val` set to an on/off word:** exactly `true`, `false`, `on`, `off`, `1` or `0` (case-sensitive) reads
     as **none**. `true` and `false` alike remove a style's underline, and Word saves `w:val="none"` (in its colour,
     where a `w:color` is given). `True`, `TRUE`, `yes` and `no` are invalid, so absent, and the style's underline
     shows. This settles D6.
   - **`auto`:** only exactly `auto` is automatic and overrides a style's colour. `Auto`, `AUTO`, `" auto"`, `"auto "`
     and `#auto` are invalid, so absent, and the style's colour shows. (Round 3 could not tell these apart, having no
     style; the portfolio's "auto in any case" was corrected the same day, and docx4j's CR-060 rule 3 with it.) Hex
     stays as round 3 measured it: trimmed, `#` accepted, any case.
   - **Highlight:** Word draws no highlight from a character style (alone, with bold, or through `basedOn`), with no
     repair involved, and its save writes `w:highlight w:val="none"` on those runs. A paragraph style's highlight and
     a direct one are drawn. A character style's shading is drawn.
   - **Unusable values over a style, by where they sit:**
     - an unusable `w:shd/@w:fill` is automatic (no shading) and **overrides** the style (saved `fill="auto"`);
     - an unusable border colour is automatic (black) and **overrides** (saved `auto`);
     - an unusable `w:u/@w:color` is **absent**: the style's underline colour shows. `auto` there draws black over a
       style's red.
   - **Two of Word's own saves change meaning,** so an oracle comparing probe and save states these two from the
     PDFs:
     - H7: direct green over a highlighting character style is drawn green, but the save drops the green;
     - U4: `w:u w:color="auto"` over a red underline style is drawn black, but the save drops the `w:u`.

## 3. What this package does (measured 2026-10-24, main `4837d04`)

The portfolio session had both rounds run through this package. The rounds' probe files were read in place, and each
marker paragraph was also loaded on its own in an in-memory copy, to see which values fail.

- **Whole-part failures** (`PartUnmarshalException`, from jsonix at unmarshal):
  - Any on/off value but exactly `true`, `false`, `1`, `0`, `on` or `off` ("Either [true], [1], [on], [false], [0]
    or [off] expected as boolean value"), for all eleven toggles. That is every value in rule 1's "on" list but
    bare, `true`, `on` and `1`.
  - Any `w:u` outside `UnderlineEnumeration` (including the on/off words), and any `vertAlign` or `em` outside
    `STVerticalAlignRun` or `STEm` ("Value [..] is invalid for the enum type [..]").
- **Raw readings:** `highlight` and `color` are plain strings in the model, so they load.
  - `Font.highlightColor` reads `YELLOW` and `bogus` raw, where Word has no highlight.
  - `Font.color` reads `red` raw, where Word has `#FF0000`. Lower-case hex reads correctly (`#FF0000`).
  - A bare `highlight`, `vertAlign`, `em` or `color` reads as Word reads it after its repair: none, baseline, no
    mark, automatic.
- **What a save writes**, where the part is re-marshalled (an untouched part is copied byte for byte):
  - An on value is written `w:val="true"`, but a bare one stays bare.
  - An off value is written `w:val="false"`, kept even where it overrides nothing.
  - `red`, `YELLOW` and lower-case hex are written as they were read.
  - The bare enumeration elements are written back, which Word repairs again.
- **What already matches Word:** every valid value, for all eleven toggles. Also bare means on; off is exactly
  `false`/`off`/`0`; the toggle rule (S7) and the style cases S1 to S11; bare and colour-only `w:u`; `"none"` over a
  style; vanish hidden.

## 4. Where the readings are decided

- **On/off:** every reader tests `val !== false`.
  - Here: the catalogue's `isTrue` (`src/model/properties/catalogue.mts:111`), with the toggles in `styleUtil.mts`,
    `PropertyResolver`'s vanish test, `tree.mts` (hidden runs, a control's flags), `Table`, `ContentControl`, the
    custom XML kinds and XHTML, `RunFontSelector`, the numbering's `isLgl`, the text model's `on`/`onOf`
    (`Extractor.mts:87`), and `WordprocessingMLPackage`'s tracking flag; `mathText.mts:135` already tests the strings.
  - In the objects package: `Font`'s reads go through its `readRunOptions` (`on()`, `builders/wml`).
  - So if the parse gives Word's boolean, every one of them is right with no change. That is the case for a lenient
    `ST_OnOff` at unmarshal (section 6).
- **Enumerations:**
  - `u`, `vertAlign` and `em` are jsonix enum types. An invalid value never reaches a reader today; the part fails.
  - Once it loads, the catalogue's `isEmptyU`/`isEmptyVal` and merges decide whether it overrides a style, and
    `readRunOptions` maps it for `Font`. That mapping is `WML_TO_UNDERLINE`, which reads an unknown value as `Mixed`,
    and `vertAlign === 'superscript'`.
  - `highlight` is a string: `isEmptyHighlight` and `readRunOptions`' `highlightHexValue` (which returns an unknown
    name raw).
- **Colour:**
  - `readRunOptions` (six hex digits only, else raw).
  - CR-008's `ResolvedColor.ofHex`, used by `WordColors` and the theme path of `Font.color` (six digits, optional
    `#`, no trimming, no names).
  - The catalogue's `isEmptyColor`/`statesColour`.
- **The Word shim:** its `font` is `Font`, so Office JS's readers follow `Font`.
  - Office JS's own answers for these values are not measured: Word's rendering is, `font.color` and
    `font.highlightColor` as Office JS reports them are not.
  - `toApiScript` and `spanEdits`' "already reads so" both go through `readRunOptions`.

## 5. Proposal

1. **On/off at the parse** (jsonix or the objects facade, section 6). Every `ST_OnOff` value reads as Word reads it:
   false for exactly `false`, `off` or `0`, true for anything else. Nothing here changes but tests. The jsonix
   change needs scoping: Word's rule is measured for WordprocessingML's `ST_OnOff`, not for every `xs:boolean`, and
   Excel's and PowerPoint's readers are unmeasured.
2. **An invalid enumeration value loads and reads as unset.** jsonix (or the facade) keeps the part loading. Then
   either the value is dropped at the parse (the attribute absent), or it is kept as its string and every reader
   treats a value outside the enumeration as unset. The dependency's design decides which (section 6, D1).
   - The readers here then make "unset" hold through the catalogue: an invalid value doesn't override a style, if D5
     confirms that.
   - So does `Font` (through `readRunOptions`): underline `None` rather than `Mixed`, no superscript.
   - The one exception measured: `w:u w:val="true"` is not absent but removes the underline, as `none` does (round 5).
     The other on/off words in a `w:u` over a style are unmeasured; until measured they take `true`'s reading
     (D6).
   - `highlight` gets the same rule in its readers: a value outside `ST_HighlightColor`, case-exact, is unset.
3. **`w:color` as Word reads it.** The rule is trim, an optional `#`, six hex digits in any case, or a name from the
   table, case-insensitive; else automatic.
   - It is used in `readRunOptions` (objects) and `ResolvedColor.ofHex`/`WordColors` here (CR-008's port, where
     docx4j's CR-060 normalises the same way). It is also used in `isEmptyColor`/`statesColour`: round 5 shows an
     unusable value is absent, so the style's colour shows, and only `auto` overrides it. docx4j reads an unusable
     value that isn't empty as automatic (CR-060 section 12), so parity waits for docx4j to follow Word there.
   - The same reading of a value applies to `w:u/@w:color`, `w:shd/@w:fill` and border colours (round 5). What an
     unusable one does differs (round 6). An underline colour is absent, so a style's underline colour shows. A fill
     is automatic (no shading) and a border colour automatic (black), and both override a style (D7).
   - The table is the 140 classic names, values from CSS 3: docx4j's `ColorTables.HTML` (parity). It is not CR-008's `PRESET`, which is DrawingML's and spells and values some names differently
     (`ltGoldenrodYellow` FAFA78 against Word's `lightgoldenrodyellow` FAFAD2).
4. **What a save writes** (D2): a save never changes meaning (the portfolio's rule). Only re-marshalled parts are
   affected; an untouched part is still copied byte for byte.
5. **The parity goldens:** where docx4j's CR-060 changes `StyleUtil` or `PropertyResolver` (an invalid value no longer
   overriding a style, say), the catalogue follows and the goldens are regenerated at that docx4j commit, as CR-008
   phase 1 did. A scan of the 268 `.docx` under `test/fixtures/` (2026-10-24) found none of the values this CR
   changes the reading of, but for 20 lower-case hex colours in `onoff/pandoc-code.docx`'s styles part, which
   already read as Word reads them.

## 6. Dependencies

- **jsonix** (jsonix-d2's repository, `../jsonix`):
  - **The precedent:** jsonix-CR-007 (3.4.1, 2026-10-16) made the runtime's `Boolean` accept `on` and `off`, at this
    package's request (CR-001 section 25: a Pandoc document, `test/fixtures/onoff/pandoc-code.docx`, would not
    open). It put the fix in the runtime because docx4j's `wml.xsd` declares `CT_OnOff`'s `val` as `xsd:boolean`
    (52 declarations, and `ST_OnOff` only 8). A compiler type for `ST_OnOff` would miss the toggles, and the
    1,108 `Boolean` properties across the Office mappings share the one type.
  - **What is needed:** a parse that takes Word's whole rule, where today's boolean type refuses any value outside
    `true`/`false`/`1`/`0`/`on`/`off`. Unlike CR-007's two forms, this changes values that were errors into
    readings (`FALSE` is on). Also an enum value outside its enumeration, kept or dropped, rather than refused.
  - **jsonix-CR-010**, proposed 2026-10-24 (jsonix `4864682`, unpushed; registered as `jsonix/CR-010`), answers it
    with callbacks rather than Word's rules in the runtime. Jason decided its J1 to J3 as recommended, so
    `onUnprintableValue` ships with `onInvalidValue`, as D2's source-spelling alternative needs. It was implemented the same day on jsonix master
    `5df33ad`, verified by objects for CR-013.B and pushed (CI green), and **released in `@docx4j/jsonix` 3.7.0** (tag
    at `d4e8af8`, with provenance; npm's latest). objects' prototype of CR-013.B on that commit read all 174 run-property markers in the 20
    probes as Word's saves do: every source probe loads, and a save and reload reads the same. Until the objects
    facade sets the callbacks, nothing changes here.
    - `createUnmarshaller({ onInvalidValue(text, propertyInfo, classInfo, node, error) })` is consulted only after
      an attribute value fails its type's parse. It returns the value to store, or `undefined` to leave the property
      unset, or throws to abort, located exactly as today.
    - `createMarshaller({ onUnprintableValue(value, propertyInfo, classInfo, error) })` returns the string to write,
      or `undefined` to omit the attribute; it is needed only for a kept enum string.
    - **Scope:** Word's rule can't be global (1,051 of the 1,099 `Boolean` attribute properties are outside WML:
      DrawingML 189, SpreadsheetML 764, PresentationML 91). It can't be per part either, since `document.xml` holds
      DrawingML booleans (`wp:anchor`'s, the locks). The callback gets the class info, so the objects facade can
      take `org_docx4j_wml`'s classes and rethrow everything else, keeping today's behaviour where Word is unmeasured.
    - **It covers these cases exactly:** every value the strict parse accepts is one Word reads the same way, so
      the callback sees precisely the values whose reading this CR decides. For on/off the facade returns `true`,
      and a marshal writes `"true"`, which Word reads the same.
    - **Keep or drop:** the runtime supports both. jsonix-d2 recommends keep until D5 is measured, a point taken in
      D2 below.
- **objects** (docx4j-generated-objects-ts-82's, `../docx4j-generated-objects-ts`): its **CR-013**, proposed
  2026-10-24 (objects `e0af904`, unpushed) and registered by phase. Nothing is implemented; it waits for Jason.
  - **CR-013.A** (this CR's phase 2, and `highlight`), no dependency:
    - `readRunOptions` reads a `w:highlight` outside `ST_HighlightColor`, case-exact, as `null`.
    - `w:color` goes through a new `helpers/wml` export: `WORD_COLOR_NAMES` (the 140 names) and `wordColor(value)`.
      It returns `'#RRGGBB'`, `'auto'` or `undefined`: it trims, takes an optional `#` and hex in any case, or a
      name case-insensitively, reads `auto` in any case, and gives `undefined` for an unusable value. Round 6 makes
      only exactly `auto` automatic: `Auto` and `AUTO` are unusable, and absent over a style (rule 5).
    - Its **O1**: an unusable colour reads `''` (as if not set) in `readRunOptions` until D5's probe, then follows
      it, as this CR's interim "unset, as if absent".
    - Its **O2** proposes that `ResolvedColor.ofHex` and `WordColors` here call `wordColor` rather than keep a second
      table. Recommendation: yes, with docx4j CR-060's reading of the same values held to it for parity.
  - **CR-013.B** (this CR's phase 1), waiting on jsonix's half:
    - If the runtime drops an invalid enumeration value, as docx4j CR-060's D3 does, `readRunOptions` already reads
      underline `None` and no superscript.
    - If the runtime keeps the string, underline's fallback becomes `None`, not `Mixed`. Every valid `ST_Underline`
      value has an Office JS name, so `Mixed` is reached only by an invalid value.
    - `on()` needs no change if the parse gives Word's boolean.
    - For the facade's part of the scoping, objects prefers marking the WordprocessingML mappings lenient over
      setting it per part, since a `document.xml` also holds DrawingML and VML booleans, for which Word's rule
      isn't measured.
  - **Implemented 2026-10-24: phases A and B** (objects main `7feac7f`, then planned for 0.4.4 and released in 0.5.0; floor `@docx4j/jsonix`
    `^3.7.0`). That is everything this CR needs from objects but phase C.
    - `helpers/wml`: `WORD_COLOR_NAMES` (the 140 names, frozen, held equal to docx4j's) and `wordColor(value)`.
      Exactly `auto` gives `'auto'`; hex, trimmed, with an optional `#`, or a name in any case gives `'#RRGGBB'`; else
      `undefined`, which the caller reads by place (absent in `w:color` and `w:u/@w:color`, automatic in a fill or
      border).
    - `readRunOptions`: `color` is `wordColor(val) ?? ''`, `highlightColor` is `null` outside `ST_HighlightColor`, and
      `underline` is `'None'` (no longer `'Mixed'`) outside `ST_Underline`.
    - The facade's `unmarshalString`, `unmarshalNode` and `unmarshalPackage` pass `wordInvalidValue` as
      `onInvalidValue` unless given their own. It is scoped to `wml`, `w14`, `w15`, `w15symex`, `w16`, `w16cex` and
      `w16cid` (D10): a refused boolean is `true`, a refused enumeration value is dropped, and `w:u`'s six on/off words
      give `'none'`. Anything else still fails the part.
    - D8's `onWordReading(reading)` option on `unmarshalString` and `unmarshalNode` gives `{ element, attribute, text,
      value, kind }`. `w:u="true"` reports `value.kept.invalid` with value `'none'`. It isn't called when a caller
      passes its own `onInvalidValue` (`wordInvalidValue` is exported to compose), and `unmarshalPackage` takes no
      options, so it reads by the rule and reports to no one.
    - Over Jason's 28 probes (rounds 1 to 6), 373 of 376 markers read as Word's saves do. The 3 differences are
      `colour-elsewhere-styles` U2 to U4, where Word's save drops a direct `w:u w:val="single"` that repeats the
      style's. The 341 corpus hashes are identical to 0.4.3's.
    - Phase C waits on jsonix-CR-011, which objects is verifying for jsonix-d2.
  - **CR-013.C:** the save normaliser, revised for this CR's decisions (objects `7009f54`, unpushed; it depends on .A
    and .B, and waits for Jason's go).
    - It follows docx4j CR-060 with CR-061's table, value for value, as a marshal-only pass over the marshalled DOM,
      `w:rPrChange`'s former properties included (D9).
    - **The opt-out:** `setNormaliseOnSave(false)` process-wide, and a per-call marshal option `normaliseOnSave`, which
      `pkg.normaliseOnSave` here passes. Off writes colours as read and keeps the `highlight`, `vertAlign` and `em`
      elements. On/off and `w:u` are settled at the parse, so they are written the same either way.
    - A scan of objects' 341 fidelity parts found no colour outside the normal form, so Office's own output doesn't
      move.
- **docx4j** (CR-060, in parallel; parity, not a registry dependency, since neither waits on the other). Its D1 to
  D6 were accepted as recommended (Jason, 2026-10-24, `5d7564778`), and it was implemented the same day on
  `VERSION_17_3_5` (`87e6ace56`, `893658f68`, `868a1797b`; phase 4 `1d1b824a9`). With CR-061 it was **pushed at
  `5163c2986`** the same day, after a full clean build (core-tests 2072 run, 0 failures, 30 skipped; export-fo 858, 0,
  9). As the oracle, what it decided:
  - **On/off:** an adapter on the `Boolean` fields; off for exactly `false`, `off` or `0`, else on; written `true`
    or `false` (its D2).
  - **`w:color`:** `WordColors.normalise` trims, takes six hex digits in either case with an optional `#`, else one
    of the 140 names in `ColorTables.HTML`, case-insensitive; anything else is automatic. It is applied to `w:color`
    and `w:u/@w:color` only (`WordColors.of(Color)`, `of(U)`); `w:shd` and border colours are untouched, names there
    being unmeasured. A source colour keeps its spelling on save (its D4).
  - **An invalid direct `w:color` over a style's colour:** docx4j keeps the direct value, so it reads automatic.
    `StyleUtil`'s colour emptiness is unchanged, which the catalogue here already matches. That is not "as if absent";
    it is still unmeasured in Word (D5).
  - **Highlight:** outside `ST_HighlightColor`, case-exact, is absent, the resolver included. `StyleUtil.isEmpty(Highlight)`
    tests membership (`WordRepairs.isHighlightColor`), so a style's highlight beneath shows.
  - **`w:u` without a valid value** (bare, colour-only, invalid, which JAXB reads as `null`): no underline drawn.
  - **On save (its D3):** a `highlight`, `vertAlign` or `em` with a missing or invalid value is left out of `w:rPr`
    and a paragraph mark's `w:rPr` (the object in memory keeps it); a revision's former `w:rPr` is written as it is.
    `w:u` is written bare, as JAXB writes an unreadable value. It records the D5 caveat.
  - **Out of its scope:** `w:tblLook`'s `firstRow` and the rest, an enumeration of the six spellings there.
  - **Its fixtures:** `docx4j-core-tests/src/test/resources/runprops/word365/` (the 20 probes and Word's 20 saves, with
    a README). Its `RunPropertiesWord365Test` takes Word's save as the oracle: each marker reads in the probe as it
    reads in Word's save.

  **What parity takes here:**
  - The catalogue's `isEmptyHighlight` follows `StyleUtil` (membership, case-exact).
  - `WordColors.of` and `ofUnderline` follow `normalise`. CR-008 phase 2 ported `WordColors` at `1fffc20a3`, before
    it; this CR's phase 2 brings it, through objects' `wordColor` (O2), held to `ColorTables.HTML`.
  - An invalid direct colour still states a colour (automatic) in the effective properties, as now.
  - The parity goldens are regenerated at docx4j's pushed commit. The scan above found none of these values in the
    fixtures, so no answer should move.

## 7. Decisions (decided 2026-10-24, as recommended)

Jason, 2026-10-24: "CR-015 D1 to D11 decided as now recommended". Each decision is its recommendation below, as it
stood then: D2 as docx4j's rule (CR-060 with CR-061) value for value, D10 as revised on python's answer, D5 and D6 as
rounds 5 and 6 answered them, D7 following docx4j's phase 4, D11 settled.

- **D1. Where the leniency lives.**
  - (a) At the parse, in jsonix and the objects facade: on/off to Word's boolean, an invalid enum value dropped or
    kept as a string. Every reader here, the text model and every consumer agree for free.
  - (b) In the readers here and in `readRunOptions`, over raw strings: the parse would have to keep the strings
    anyway, so it doesn't save the dependency.
  - Recommendation: (a) for on/off and the enumerations, with `readRunOptions` and this package's readers taking
    `highlight` and `color`, which are strings already.
- **D2. What a save writes for a value read leniently.** Recommendation: **docx4j's rule, value for value**
  (CR-060 with CR-061, both decided by Jason on 2026-10-24).
  - **The principle** (Jason's direction, given to python and docx4j alike): "they are entitled to assume we don't
    write dodgy content". A save never changes Word's reading and never writes content the schema rejects. A value
    is rewritten only when its spelling is invalid, never merely to match Word's own spelling.
  - **What is written, on by default:**
    - **On/off:** `true`/`false`, as jsonix's boolean writes it and as docx4j does. It is valid in every edition,
      Strict included, and Word reads it exactly ("No need to switch on/off to bare or 0"). A bare value stays bare.
    - **`w:u w:val`:** exactly `true`, `false`, `on`, `off`, `1` or `0` is written `none` (D6). Any other invalid
      value is dropped and the `w:u` written bare, which Word reads as absent (round 5). A colour on it stays.
    - **`highlight`, `vertAlign`, `em`:** an element whose value is missing or invalid is dropped, as Word's repair
      does.
    - **Colours** (docx4j CR-061):
      - an accepted value (a name, hex in any case, `#`, padding) is written as upper-case `RRGGBB`, and `auto` as
        `auto`;
      - a rejected `w:color/@w:val` drops the `w:color` only where it has no theme attribute at all. Beside a
        `w:themeColor`, or a `w:themeTint` or `w:themeShade` without one, the element is kept and its `val` written
        `"auto"` (`w:val` is required, and the theme wins). This is docx4j's code as it landed (`aaadd93f9`);
      - a rejected `w:u/@w:color` drops the attribute;
      - a rejected `w:shd/@w:fill` or border colour is written `auto`. So is a rejected `w:shd/@w:color`, a shading
        pattern's colour, which follows the fill's rule (round 7: a name drawn as hex; `F00` automatic and saved
        `auto`, over a style too).
      - `w:background/@w:color`, the page colour, reads names too. Where its colour is rejected, Word drops the whole
        `w:background`, and docx4j writes `w:color="auto"`. That is valid, and both presumably mean no page colour,
        though it is unmeasured. This port writes docx4j's form, value for value; matching Word exactly would mean
        dropping the element.
    - **A revision's former properties** (`w:rPrChange`'s, which Reject All restores) are treated like live ones (D9).
  - **How:** marshal-only, as docx4j's (its D6). The tree in memory keeps the source's value, and only the output is
    normalised. That is objects CR-013.C grown into the normaliser.
  - **The opt-out**, as docx4j's (its D1, D2): a process-wide default and a per-package setting
    (`pkg.normaliseOnSave = false`, beside `pkg.fonts.defaultTheme`'s pattern). Off means colours are written as read
    and the `highlight`/`vertAlign`/`em` drop is off; on/off stays `true`/`false` either way.
  - **A part never unmarshalled** is copied byte for byte, as now.
  - **Across ports:** docx4j-python's CR-018 section 4.7 writes the same values, and differs only in on/off's
    spelling (bare and `"0"`; D11: free).
  - **docx4j's implementation:** CR-061 is complete on `VERSION_17_3_5` (`aaadd93f9`, then `5163c2986`, unpushed),
    tested by `ColourNormalFormTest` and `ColourNormalFormOptOutTest`.
    - Its phase 3 is Jason's Word check. docx4j's saves of 15 probes opened with no repair, Word's saves of them equal
      its saves of the original probes at all 251 markers, and the PDFs are pixel-identical, so the normal form keeps
      Word's reading.
    - It also ran round 7, whose fixtures are in docx4j's `runprops/word365/` (the rows above).
  - **The alternatives**, kept for the record:
    - the source's spelling throughout, through jsonix-CR-010's `onUnprintableValue`, which changes no reader's
      meaning but keeps writing what Word repairs;
    - CR-060 alone, before CR-061 (colours as read).
- **D3. The colour table:** the 140 classic names with CSS 3's values, the 41 measured names held exactly to Word's
  saves, the rest taken from the table. Recommendation: yes. docx4j has it as `ColorTables.HTML` (its D5), which
  objects' `WORD_COLOR_NAMES` and this package's tests are held to.
- **D4. Office JS's own readers:** the shim promises Office JS's answers, and Office JS's `font.color`,
  `font.highlightColor`, `font.underline` and `font.superscript` over these probes are unmeasured. Recommendation: a
  Script Lab check (check 60) reading the probes' markers, before implementing `Font`'s side.
- **D5. An invalid value over a style: answered by round 5** (section 2, rule 4). An invalid enumeration value or
  highlight is "as if absent", as docx4j reads it, so dropping it on save keeps meaning, as every option of D2 needs. Two
  answers differ from today's interim readings:
  - an unusable colour is absent too (the style's colour shows), where docx4j and the catalogue here read it
    automatic;
  - `w:u w:val="true"` removes the underline, where docx4j reads it absent and a bare save would turn the style's
    underline back on.
  Both are docx4j's to follow first (CR-060), and this port's with it.
- **D6. The on/off words in a `w:u`: answered by round 6** (section 2, rule 5). Exactly `true`, `false`, `on`,
  `off`, `1` or `0` reads as `none`, and a save writes `w:val="none"` (in its colour, if any). Any other value reads
  as absent and is written bare (D2).
- **D7. Round 6's other findings, docx4j's first, then this port's.**
  - A character style's highlight is never drawn: the resolver here would leave `w:highlight` out of what a
    character style contributes, while a paragraph style's and a direct one still apply.
  - An unusable value's reading depends on the property. A run colour, a highlight, `vertAlign`, `em`, `w:u/@w:val`
    (but for the on/off words) and an underline colour are absent. A shading fill and a border colour are automatic,
    and override the style.
  - Only exactly `auto` is automatic.
  - Recommendation: take each as docx4j's CR-060 does, for parity; the goldens are regenerated at that commit.
    docx4j's phase 4 (`1d1b824a9`, unpushed) follows round 6 in every reading above (section 10), so this waits only
    for its push.

- **D8. A record of what was read by Word's rule, and what a save rewrote** (CR-018's Q16 and Q17). Each value the
  facade's callback reads by Word's rule becomes a `PartWarning` on `XmlPart.warnings` (which already carries
  `mce.branches_differ`), and `OpcPackage.warnings` gathers them. A save under D2's normaliser reports each rewrite (part,
  path, before, after). Recommendation: yes, with names that read as CR-018's (python's answer, 2026-10-24).
  - `value.kept.exact` for a value that would be a lexical form once trimmed (`" false"`, `" yellow"`).
  - `value.kept.invalid` for one that wouldn't (`"yes"`, `"Yellow"`, `"bogus"`).
  - `value.normalised` for a save's rewrite.
  - CR-018's record is `KeptValue(qname, parent_class, path, text, type, reason)`, surfaced as `part.kept` and
    `package.kept`; its save report is `package.last_save.normalised`. jsonix does not collapse whitespace, so both
    kinds reach the callback here as invalid; the facade tells them apart by trimming.
  - docx4j logs each rewrite at DEBUG and has no report object (its CR-061 D5). The record is this package's choice;
    a warning per value costs nothing where there is none.
  - **Where the data comes from** (objects CR-013 O6, decided by Jason as recommended, objects `bdfc52e`): the facade
    hands it over rather than raising anything or writing to the console, having no warnings model, and a
    `console.warn` per value would be noise in an Office JS add-in. Its O4 is decided too: every enumeration in the
    lenient scope is lenient.
    - An unmarshal option, `onWordReading(record)`, gives the source element, the attribute's name, the text as read,
      the value the model holds instead (`true`, `'none'`, or none if dropped), and the kind (`value.kept.exact` or
      `value.kept.invalid`). This package makes each a `PartWarning` on the part being read.
    - A marshal option, `onNormalised(record)`, gives the written element's path, the attribute (or the element, if
      dropped), and the values before and after. This package adds the part and keeps the save's report.
- **D9. A revision's former properties** (`w:rPrChange/w:rPr`, a mark's `w:rPrChange`). Under docx4j's adapters
  they get the type-level rules: on/off `true`/`false`, `w:u` on/off words `none`. Its phase 3 drop doesn't reach them
  (a list there). CR-018 normalises them, since Reject All restores them and Word reads them by the same rules, and
  docx4j-25 agrees they should be treated as live properties. Recommendation: the same rules as live properties,
  whichever D2 Jason takes.
- **D10. The scope beyond `org_docx4j_wml`.** The facade's scoping by module (jsonix-CR-010) leaves the `w14` and
  `w15` modules strict, so a `w14:checked w:val="TRUE"` would still fail its part. These are WordprocessingML too, and
  their on/off readers are unmeasured (CR-018's Q12 lists a bare `w14:checked`). CR-018's parse keeps every
  `xs:boolean` in every namespace as text and reads it by Word's rule, so python reads that value as on.
  - Recommendation (revised on python's answer): include Word's own WordprocessingML modules (`w14`, `w15` and the
    later `w16*`) in the facade's scope by extension, as python does. Refusing a whole part is worse than an
    unmeasured reading by Word's rule.
  - DrawingML, VML, SpreadsheetML and PresentationML booleans stay strict (jsonix-CR-010's reason).
  - This is objects' to apply, if Jason agrees. Its CR-013 O5 (objects `bc043fe`, unpushed) gives the detail.
    - Of the six mappings only `w14` has on/off values.
    - `CT_OnOff.val`, a `Boolean` (`w14:checkbox`/`w14:checked`, `w14:conflictMode`, `w14:discardImageEditingData`),
      takes the `Boolean` rule, so `"TRUE"` reads true.
    - Two on/off values are typed as an anonymous enumeration of `true`, `false`, `0` and `1`:
      `CT_LinearShadeProperties/@scaled`, and `CT_StyleSet/@val` in `w14:stylisticSets`. O5 proposed reading them by
      the `Boolean` rule mapped onto the members. On implementing, objects found both are strings in the mapping (as in
      docx4j), never refused and held as written, so the rule has nothing to apply to there. `w14`'s `CT_OnOff`
      (`w14:checked`) is a real boolean and does get it.
    - `w14`'s 13 other enumerations, the text effects among them, would become lenient too. `w15`, `w15symex`,
      `w16`, `w16cex` and `w16cid` have no booleans, and one enumeration between them.
- **D11. The spelling of on and off, across ports: settled.** CR-018 writes a rewritten on bare and off `"0"`; docx4j
  and jsonix write `"true"`/`"false"`. Word reads them alike, and python's answer is that the ports should agree on
  readings, not bytes. So the spelling is free.

## 8. Tests

- **The probes:** docx4j's copy, `docx4j-core-tests/src/test/resources/runprops/word365/` at `5163c2986` (rounds 1 to
  7: 33 probes, each with Word's save, and its README of provenance). It is copied into `test/runprops-word365/` with a
  `manifest.json` of SHA-1s, as `test/theme-colors/` was. It sits outside `fixtures/`, so that no fixture-wide test or
  harness counts it.
- **docx4j's tests**, ported as `theme-colors.test.mjs` ported CR-034's: `RunPropertiesWord365Test`,
  `ColourNormalFormTest`, `ColourNormalFormOptOutTest` and `OnOffAttributesTest`.
- **docx4j's oracle:** each marker reads in the probe as it reads in Word's save (`RunPropertiesWord365Test`), here
  through `Font`, the effective resolver and the text model, compared as what is drawn: off the same as absent,
  colours as upper-case hex. H7 and U4 are stated from Word's PDFs, since Word's own saves change their meaning
  (section 2, rule 5), as docx4j's test states them.
- **What a test holds:** for each probe it holds each marker to Word's reading:
  - the part unmarshals;
  - `Font`'s bold, italic, strikes, underline, superscript, highlight and colour, effective and direct;
  - the effective toggles for the rest; the text model's runs and hidden text;
  - and a re-marshalled save keeps the meaning (reloaded, every marker reads the same) in D2's form.
- **Word's saves:** these are loaded too, and must read as their sources do.
- **The round-trip contract** ("re-marshalled parts deep-equal after reload") gains, under D2's normaliser, one tolerated
  difference: a value written in its normal form (CR-018's `word-normal-form` category). The test applies the same
  normaliser to the source before comparing, so nothing else may move; under `normalise: false` the contract stands
  unchanged.
- **Dependency defects:** until the dependencies land, the failing values are asserted as failing, with the upstream
  commit in a comment (CLAUDE.md's rule), so that each fix announces itself.

## 9. Phases and estimate

- **Phase 0:** the probes copied, and the test asserting today's behaviour (section 3). Half a day.
- **Phase 1:** on/off and the enumerations through the dependencies (D1), the catalogue's unset rule (after D5),
  `highlight`. One day once the dependencies are released.
- **Phase 2:** `w:color` (D3), in `readRunOptions`'s release and here, with `w:shd/@w:fill` and border colours (D7).
  Half a day.
- **Phase 2b:** what a save writes (D2) and its record (D8), on the objects facade's marshal (CR-013.C, grown to D2
  its normaliser), the round-trip contract's tolerated category, a revision's former properties
  (D9). Half a day here once objects has it.
- **Phase 3:** D4's check, and the shim's answers held to it. Half a day plus Jason's run.

## 10. Review against docx4j-python CR-018 (2026-10-24)

Jason asked for CR-015 to be reviewed against CR-018's decisions (python `523afb4`), with leave to discuss with the
python and docx4j sessions. Both were asked the same day (questions below); their answers go into this section.

**Where the two agree** (the readings): on/off off only for exactly `false`, `off`, `0`; an enumeration value outside its
enumeration unset, in the merge as in the readers; `w:u`'s six on/off words read as none; one colour reader, the 140
names, only exactly `auto` automatic; an unusable colour absent for `w:color` and `w:u/@w:color`, automatic and
overriding for `w:shd/@w:fill` and borders; a character style's highlight never drawn; every on/off reader in scope,
outside `w:rPr` by extension (CR-018 Q4); probes copied from docx4j's `runprops/word365/` (CR-018 Q14); a padded
enumeration value unset (CR-018 Q15: jsonix 3.5.0 doesn't trim, so `" single"`, `"dot "` and `" superscript"` fail its
parse and reach the callback, and `" yellow"` loads raw, unset by membership; measured here 2026-10-24).

**What CR-015 changes on review:**
- D2's recommendation becomes a normaliser on by default with an opt-out, on Jason's direction to python, which
  came after CR-060's acceptance. docx4j has since decided the same (CR-061), so D2 is docx4j's rule value for value,
  and python's CR-018 section 4.7 agrees but for on/off's spelling.
- New D8 (a record of values read by Word's rule and of a save's rewrites, CR-018 Q16 and Q17), D9 (a revision's
  former properties normalised), D10 (`w14` and `w15` in the facade's scope by extension, revised on python's
  answer), D11 (the spelling
  of off).
- The round-trip contract gains CR-018's tolerated category under D2's normaliser (section 8).

**Not taken from CR-018, and why:** its `effective()` reader of schema defaults (its (d), over CR-059's corrected
`wml.xsd`) has no counterpart here. This package's readers state their own rule for an absent value, docx4j's
(`isTrue`: absent `w:val` means on), and consult no generated default (the portfolio's question of 2026-10-24). Its
python-docx facade (its Q9) is the Word shim's place here, whose Office JS answers D4's check measures.

**docx4j-25's answers** (2026-10-24, as docx4j stands at CR-060 phase 4, `1d1b824a9`, unpushed):
1. CR-060's D2 and D4 stood as accepted, and docx4j-25 put them to Jason. His answer, the same day:
   - D2 stays (`true`/`false`).
   - D4 is revisited by docx4j CR-061, "Colours written plainly on save" (`c6a2c6b47`). Jason decided CR-061's D1 to
     D6 as recommended the same day, and its implementation started.
   - D2 above is docx4j's rule as docx4j-25 stated it, so this port can match docx4j value for value.
2. Phase 4 applies rounds 5 and 6.
   - **Underline:** `UnderlineValueAdapter` reads exactly `true`, `false`, `on`, `off`, `1` and `0` as `NONE`, so a
     save writes `none`. Any other invalid `w:u` value is still written bare, which Word reads as absent.
   - **Fill and border colours:** an unusable one is *not* rewritten as `auto`; docx4j keeps the source value, which
     Word reads as automatic, so meaning is kept there too.
   - **Readings:** a fill or border name resolves and an invalid one is automatic; an invalid text or underline
     colour is absent; only exactly `auto` is automatic; a character style's highlight is not drawn.
   So D7's readings are docx4j's now, and this port's to follow.
3. **A revision's former properties:** in docx4j `w:rPrChange/w:rPr` is a list (`CTRPrChange.RPr`,
   `CT_ParaRPrOriginal`), so phase 3's drop of an invalid `highlight`, `vertAlign` or `em` doesn't reach it. The
   type-level adapters do: its on/off elements are written `true`/`false`, its `w:u` on/off words `none`, and its
   colours are kept as read. docx4j-25's view is that they should get the same treatment as live properties, since
   Reject All restores them. Whether docx4j extends the drop there, or adopts a normaliser at all, goes to Jason with
   point 1.

**The python session's answers** (2026-10-24; CR-018 decided, every question as recommended, `45154c0`, and in
implementation; its stale sections 3.2 and 4.10 fixed):
1. Off's spelling doesn't matter for agreement: the ports should agree on readings, not bytes (D11).
2. On is written bare only where the normaliser rewrites a value the parse kept as text (`" false"`, `"on"`,
   `"yes"`, `""`). That is everywhere the save re-marshals, `w:rPrChange`'s old properties and `styles.xml` included.
   A value that bound as a boolean is written `"1"`, as today.
3. An unusable `w:color` beside a `w:themeColor` keeps the theme, written `w:val="auto"` since `w:val` is required.
   Without a theme it is dropped (D2, python `531efcf`).
4. Its record and report names, which D8 takes.
5. Neither round measured `bCs`/`iCs`, the paragraph-level on/off properties (`keepNext`, `tblHeader`, `titlePg`,
   `trackRevisions`), a bare `w14:checked` or `w:fldLock` `" false"`/`"TRUE"`; python takes them by extension too.
   Its parse covers every `xs:boolean` in every namespace, so the `w14:checked` gap doesn't arise there (D10).

Python agrees with the jsonix trimming measurement (CR-018 Q15).

## 11. Implementation (2026-10-24)

Jason, 2026-10-24: "implement CR-015". Phases 0, 1, 2 and 2b are implemented; phase 3's check is written and waits for
his run. The code needs objects CR-013 phases A to C, verified here at objects `b2c9163` (section 11.8) and published the
same day as **0.5.0** (tag at `2c18fbb`, `b2c9163` with the version; a minor, Jason's choice, since phase C changes what a
save writes by default), over `@docx4j/jsonix` ^3.8.0. The range is `^0.5.0`.

### 11.1 The goldens at docx4j `5163c2986`, and `w:framePr`'s `anchorLock`

docx4j's `82c64bf14` (its CR-059 section 11, a regression docx4j-python found in `ac97b56c5`) reads
`w:framePr/@w:anchorLock` as stated: a `w:framePr` is empty only with `anchorLock` absent too, and a merge keeps the
style's where the more specific `w:framePr` omits it. Until then docx4j read it through `isAnchorLock()`, whose old
schema default made an absent attribute true, and wrote that `true` out in every merge. The catalogue followed the old
reading (`isEmptyFramePr`, `mergeFramePr`); it now follows the new one, with docx4j's `StyleUtilFramePrAnchorLockTest`
ported to `properties.test.mjs`.

All 502 committed goldens were regenerated from docx4j `VERSION_17_3_5` at `5163c2986` (pushed; a `git archive` built
into a scratch Maven repository; harness versions 4, 1, 2 and 1). Eight answers moved, each a paragraph's
`effectivePPr` losing `w:anchorLock="true"`: five in `loadAndSave.json`, two in `NumberingImplicitNumId.json`, one in
`NumberingIndents.json`. Nothing else moved: CR-060 and CR-061 change no fixture's answer, as section 5 predicted, and no
XPath hit, revision listing or text block moved. Parity is zero differences, under objects 0.4.3 and under `b2c9163`.
`test/golden/README.md` and CLAUDE.md record the commit; `test/java/pom.xml` still resolves the 17.3.4 release for a
hand-run, which gives the eight `framePr` answers the old way.

### 11.2 Phase 0: the probes

docx4j's `runprops/word365/` at `5163c2986` (rounds 1 to 7: 33 probes, Word's 33 saves and the README) is copied
unedited into `test/runprops-word365/`, with a `manifest.json` of SHA-1s that `runprops-word365.test.mjs` checks.
docx4j has since added round 8's probes there (its CR-062, `6f00d217f`, unpushed); they are rendering probes
(underline styles, `w:em` marks) with nothing for this package's readers, and are not copied.

Under objects 0.4.3 thirteen of the 66 files fail to unmarshal (the part refused at its first invalid on/off or
enumeration value). Phase 0's assertions of that behaviour were never committed: the dependency arrived the same day,
so the test was written against it.

### 11.3 Phase 1: on/off and the enumerations

- **The parse** is the objects facade's (D1 (a)): `wordInvalidValue` as jsonix-CR-010's `onInvalidValue`, scoped to
  `wml`, `w14`, `w15`, `w15symex`, `w16`, `w16cex` and `w16cid` (D10). Every reader here was already right once the
  parse gives Word's boolean; nothing on/off was changed.
- **The catalogue** follows docx4j CR-060's `StyleUtil`:
  - `isEmptyHighlight` is membership of `ST_HighlightColor`, case-exact (`isHighlightColor`, docx4j's
    `WordRepairs`), so an invalid highlight leaves a style's beneath;
  - `isEmptyColor` and `isEmptyU` count a colour value only where Word states one (`WordColors.stated`: one Word
    accepts, or exactly `auto`), and `mergeColor`/`mergeU` carry it only then (`statesWordColour`); a fill's and a
    border's keep `statesColour`, any value, since Word reads one it rejects there as automatic, overriding.
- **The resolver:** a character style's chain applies as a level without its `w:highlight` (`characterLevel`, docx4j
  `1d1b824a9`), for a run and for a paragraph mark; Word draws no highlight from a character style (round 6).
- **D8, the record of a read:** each value the facade reads by Word's rule is a `ValueKeptWarning` on the part's
  `warnings` (`value.kept.exact` or `value.kept.invalid`, CR-018's names; the element and attribute as written, the
  text, the value held, the line and column where the DOM gives them), from objects' `onWordReading`, and
  `OpcPackage.warnings` gathers them. `PartWarning` is now `McBranchesDifferWarning | ValueKeptWarning`. A private read
  (`readContents`) records them too, and leaves the part untouched.

### 11.4 Phase 2: colours

- `WordColors.normalise` and `stated` are docx4j's (`WordColorValues`), over the objects package's `wordColor` (O2),
  which its smoke test holds to docx4j's 140-name table.
- `WordColors.of` (a run's colour) and `ofUnderline` resolve the normalised value, as docx4j's do.
- `WordColors.fill`, `pattern` and `ofBorder` normalise too, as docx4j's do since `a9e0303ff` (local, unpushed when
  written; its CR-060 section 14, `WordColorsNamesTest`). At `5163c2986` docx4j's `WordColors` normalised only in
  `of(Color)` and `of(U)`, so its `fill(CTShd)` read `aqua` as automatic. Its `Property.fillColour` and `borderColour`,
  which its HTML and PDF draw with, normalised, and so does Word (round 5: a name in a fill or a border is drawn). This
  port normalised all five, and asked docx4j-25, who agreed and made the resolver follow. The parity goldens record no
  resolved colour, so the commit moves none of them.
- The same answer: docx4j's `RunPropertiesWord365Test` round trip still checked that `props-enum-other`'s save kept
  `w:color w:val="red"` (CR-060's D4, before CR-061). It passed only because the probe's label text holds those words.
  `a9e0303ff` checks the `N6_COLOR_NAME` run for `FF0000`, as `runprops-word365.test.mjs` does.
- `Font.color` needed nothing: the objects package's `readRunOptions` reads `w:color` by `wordColor`, and a theme
  colour goes through `WordColors.of`.

### 11.5 Phase 2b: what a save writes

- The objects facade's normaliser (its CR-013 phase C, docx4j CR-061's rule value for value) is on by default in every
  marshal. `OpcPackage.normaliseOnSave` (undefined, false or true) is passed to each part's marshal
  (`XmlPart.marshalToNode`, so `getBytes`, `getXml` and a save); undefined defers to the objects package's
  `setNormaliseOnSave`, process-wide, which this package re-exports. `clone()` carries the setting.
- **D8, the save report:** `OpcPackage.lastSave.normalised` lists each value the last save normalised, a
  `ValueNormalised` (`value.normalised`; the part, the path from its root, before and after, `after` null where the
  attribute or the element was left out), from objects' `onNormalised`. A part nobody unmarshalled is copied and
  reports nothing.
- **The round-trip contract** gains the tolerated category (section 8): `onoff/pandoc-code.docx`'s styles part, the one
  fixture part holding a value outside the normal form, reloads deep-equal but for its 20 lower-case hex colours, each
  written upper-case; with `normaliseOnSave: false` it reloads exactly.
- The editor's requirement (its ED-002 section 3.7, relayed by jsonix-d2): an unedited document saves byte for byte.
  It still does: only an unmarshalled part is marshalled, and so normalised.

### 11.6 Phase 3: check 60

`test/fixtures/check60/build.mjs` puts 25 of the probes in one file (`60-runprops.docx`, git-ignored;
the share's `__tmp/60/`), each probe's paragraphs after a `PROBE <name>` paragraph, the styles together (no id clashes).
The three `props-enum` probes are left out: Word repairs them, so its repair decides what they show. The snippet reads,
for every marker, Office JS's `font` properties one at a time (the eleven `Font` properties this package has, then
WordApiDesktop's on/off ones and `underlineColor`, which a host may lack). Through the Word shim it reads each marker as
the probe's own `Font` does.

**Run 2026-10-24** (Jason; Word 365, 16.0.20326.20232, PC; Script Lab; "no issues": Word opened the file with no repair).
The JSON is `fixtures/check60/60-runprops-word365.json`: 340 paragraphs, 337 markers read, the three hidden by `w:vanish`
recorded by their paragraph's text, as the shim records them.
- **Office JS reads as this package does,** marker for marker, in every property of the shim's `Font` (`bold`, `italic`,
  `underline`, `strikeThrough`, `doubleStrikeThrough`, `superscript`, `subscript`, `color`, `highlightColor`), but one: an
  invalid on/off value on, an invalid `w:u` no underline, a colour name its hex, a rejected colour automatic (`#000000`),
  an invalid highlight `null`, a character style's highlight `null`.
- **The one difference:** H7's direct green highlight. Office JS reports Word's `green` as `#00FF00`, and the shim
  `#008000`, the objects package's `highlightHexValue`, which is docx4j's `Highlight.getHexVal()` table (CSS's green; and
  CSS's values for the dark names, with `darkYellow` gold). Only `green` and `yellow` were in the file, and `yellow`
  agrees. The fix is the objects package's (its `readRunOptions` and `applyRunOptions` speak Office JS's vocabulary):
  told 2026-10-24, with the suggestion of a check of all sixteen names, read and written, before its table changes. The
  test asserts the difference as it stands, so that the fix announces itself.
- **WordApiDesktop's properties,** which the shim's `Font` lacks: `allCaps`, `smallCaps`, `hidden`, `emboss`, `engrave`,
  `outline` and `shadow` agree on every marker with the resolver's `caps`, `smallCaps`, `vanish`, `emboss`, `imprint`,
  `outline` and `shadow` (2,359 readings), and `underlineColor` with this rule: the `w:u`'s own colour where it states one,
  even where no underline is drawn (`U7_COLOR_ONLY`, and `w:u w:val="true" w:color="red"`, which reads none), else the
  text's colour where there is an underline, else `#000000`. A test holds the resolver to each, for the day the shim's
  `Font` gains them.

### 11.7 The tests

`runprops-word365.test.mjs`, 25 tests:
- **the probes and their manifest; every probe and save unmarshals;**
- **docx4j's `RunPropertiesWord365Test`, through this package's readers:** each of the 376 markers in the 28 files
  reads in the probe as in Word's save (with H7, M6 and U4 from the PDFs): through the effective resolver (on/off, `u`,
  highlight, `vertAlign`, `em`, colour, underline colour, fill, border), the text model (bold, italic, strike) and
  Office JS's `Font` (effective), and after a round trip here. The direct `Font` differs from Word's save in six
  markers only, each where the save rewrites direct markup (listed in the test). Its HTML and PDF tests have no
  counterpart here;
- **`anInvalidValueOverAStyleLeavesTheStyles`, `colourOnlyUnderlineIsNotDrawn`, `wordsOwnSaveChangesWhatItDrew`;**
- **D8's records** (a read's warnings, a private read's, the save report);
- **docx4j's `OnOffAttributesTest`:** the wml mapping's 48 `Boolean` attributes (docx4j's 46, and `checkStyle` and
  `actualPg`, which docx4j binds as primitives its adapter can't reach and jsonix as `Boolean` like the rest), each
  reached through a document from a global element by a search of the mapping, read for 15 values; `BooleanDefaultFalse`
  is declared but no element has the type, so no document reaches it;
- **docx4j's `ColourNormalFormTest` and `ColourNormalFormOptOutTest`,** the round-trip contract's category, and
  `WordColors.normalise`/`stated`;
- **check 60's file, snippet and shim run.**

Elsewhere:
- `properties.test.mjs` gains the `anchorLock` test, and a `{ val: 'yellow' }` sample, since a highlight and a colour
  are now stated only by a value Word reads;
- `mce.test.mjs`'s kept-branch save asserts objects 0.5.0's declaration of a source-only `Requires` prefix on its
  `mc:Choice` (the objects package's CR-012, `869b725`; Word declares it there), where 0.4.3 declared it on the root.

### 11.8 Verifying the objects release (the six steps)

On objects `7feac7f` (phases A and B) and then `b2c9163` (A to C), each packed from a `git archive` and installed with
`npm install --no-save`, with jsonix 3.8.0:

1. The swap confirmed (`wordColor`, then `setNormaliseOnSave`, in the installed `dist`), `package.json` and the lock
   untouched.
2. The whole suite: with 11.1 alone, 1,516 of 1,516 under 0.4.3 and 1,515 of 1,516 under `7feac7f`, the one being the
   `mce.test.mjs` assertion above; with all of this CR on `b2c9163`, 1,541 of 1,541; typecheck clean; parity zero.
3. Every part a load re-marshals, hashed under 0.4.3 and under `b2c9163`, over this package's fixtures and docx4j's
   chart corpus (507 files, 8,896 parts): two differ, each named. Check 50's `document.xml` declares its source-only
   prefixes on their `mc:Choice` (objects CR-012), and pandoc-code's `styles.xml` writes 20 colours upper-case (phase
   C's normaliser, listed by `onNormalised`). The 66 probe files all load under both of objects' commits; 13 fail under
   0.4.3.
4. The probes are the real data. Under jsonix in strict mode (a `--require` hook prepending `"use strict"` to
   `jsonix-factory.js`, as a browser bundle loads it), the probe, round-trip, kept-branch and parity tests pass, 301 of
   301.
5. Restored with `npm ci` afterwards (the tree then needed the release: `WordColors` imports `wordColor`, which 0.4.3
   lacks).
6. Reported to the objects session, which published 0.5.0. **Against the published tarball** (the range `^0.5.0`,
   `npm ci`, integrity `sha512-S75UZqQ0...motYkafQ==`, jsonix 3.8.0): typecheck clean, 1,541 of 1,541, and step 3's hash
   set identical to `b2c9163`'s, part for part (508 files with check 60's, 8,901 parts), so the tarball is what was
   verified.

## 12. Follow-up: Word's highlight palette (check 61)

Check 60's one difference (section 11.6) is in the objects package's highlight table. `readRunOptions` and
`applyRunOptions` speak Office JS's vocabulary, and the table they use is docx4j's `Highlight.getHexVal()`: CSS's
values, `green` `008000`, and `darkYellow` gold. Office JS read Word's `green` as `#00FF00`. Word's palette for the
other names is unmeasured. The objects session agreed it should be measured both ways before the table changes, and the
fix probably has two parts: an objects CR giving `readRunOptions`, `applyRunOptions` and `highlightNameForColor` Office
JS's values, and a docx4j counterpart if Word's palette differs from docx4j's. Jason, 2026-10-24: "write check 61".

### 12.1 Check 61

**Check 61** (`test/fixtures/check61/`; the share's `__tmp/61/`; written and run 2026-10-24):
- **Reads:** a run in each of `w:highlight`'s sixteen names, one in `none`, and one with a character shading
  (`w:shd w:fill="FFFF00"`), each read through `font.highlightColor`.
- **Writes:** `font.highlightColor` set to each name; to Word's palette as it is usually given (`WdColorIndex`: `000080`,
  `008080`, `008000`, `800080`, `800000`, `808000`, `808080`, `C0C0C0` and the eight bright ones); to the objects
  table's values that differ (`00008B`, `008B8B`, `006400`, `8B008B`, `8B0000`, `FFD700`, `A9A9A9`, `D3D3D3`); to
  `#00ff00` and `Yellow`, for case; to `#123456` and `#FF8000`, in neither; and to `null` on a highlighted run. Each is
  read back, and the run's `w:highlight` and `w:shd` taken from the range's `getOoxml()`, so it says which name Word
  stores for each value, and whether it stores a shading for a value outside the palette. The writes change the
  document, which is closed without saving.
- **Through the Word shim:** each name reads the objects table's value; a value of the table writes its name; the other
  12 values are refused (`Not a highlight colour`), as `applyRunOptions` refuses a colour it has no name for.

**Run 2026-10-24** (Jason; Word 365, 16.0.20326.20232, PC; Script Lab; no repair; closed without saving). The JSON is
`fixtures/check61/61-highlights-word365.json`.
- **Read:** Office JS reports Word's palette, `WdColorIndex`'s: `black` `#000000`, `blue` `#0000FF`, `cyan` `#00FFFF`,
  `green` `#00FF00`, `magenta` `#FF00FF`, `red` `#FF0000`, `yellow` `#FFFF00`, `white` `#FFFFFF`, `darkBlue` `#000080`,
  `darkCyan` `#008080`, `darkGreen` `#008000`, `darkMagenta` `#800080`, `darkRed` `#800000`, `darkYellow` `#808000`,
  `darkGray` `#808080`, `lightGray` `#C0C0C0`; `none` and a character shading read `null`. Nine names differ from the
  objects (docx4j) table: `green` and the eight dark and grey ones.
- **Written, a name is a CSS colour name, not a `w:highlight` name**, in any case, snapped to the palette:
  - `green` is CSS's `#008000`, so Word stores `darkGreen`;
  - `darkBlue`, `darkCyan`, `darkGreen`, `darkMagenta`, `darkRed`, `darkGray` and `lightGray` are CSS's values, which
    snap to the same-named highlights;
  - `Yellow` is `yellow`;
  - `darkYellow`, which is not a CSS name, is refused (`InvalidArgument`).
- **Written, a hex** in any case is stored as a highlight, never as a shading.
  - Word's 16 palette values store their names exactly.
  - Every other value snaps to some palette entry: the objects table's eight to the same names (`#FFD700` to
    `yellow`); `#123456` to `darkCyan`; `#FF8000` to `yellow`.
  - The snap is not the nearest in RGB: `#A9A9A9` goes to `darkGray` (`#808080`) though `lightGray` (`#C0C0C0`) is
    nearer, and `#123456` to `darkCyan` though `darkBlue` is nearer. With ten points its rule is not known.
- **`null`** removes the highlight.

So the shim differs in 39 of the 63 cases, each by the objects package's table or its write rule: the nine reads; the
nine names that aren't the bright ones, `darkYellow` included; the nine palette hex values and `#008000`, written; the
objects table's own eight values, which Office JS snaps to Word's palette; `#00ff00` and `Yellow`; and the two outside
both. A test holds the shim to Word's run with those 39 listed, so that the objects fix announces itself.

**For the objects CR** (the objects session's to write, with Jason; told 2026-10-24):
- `readRunOptions` should read Word's palette;
- `applyRunOptions` and `highlightNameForColor` should take Office JS's input: a CSS colour name, any case (the 140
  names of `WORD_COLOR_NAMES`, or CSS's full list: unmeasured beyond these), or a hex in any case;
- it should store a palette value's own name. Snapping any other value needs Word's rule, which check 61 doesn't give;
  refusing such a value, as now, is the safe interim.
- **docx4j:** its `Highlight.getHexVal()` had the same nine values, so its HTML and PDF drew a dark or grey highlight in
  CSS's colour, not Word's (`darkYellow` gold, where Word draws `#808000`). It followed the same day, in `a3cfc9789`
  (local, unpushed when written; its CR-060 section 14, `HighlightPaletteTest`):
  - `getHexVal()` gives Word's palette;
  - `setVal("#RRGGBB")` takes Word's values first, then the former CSS ones as aliases (`#D3D3D3` is still `lightGray`;
    `#008000` is now `darkGreen`), exact values only, with no nearest-colour snap;
  - `setVal("none")` keeps `none`.

  The parity goldens record `w:highlight` as written, not a colour, so this moves none of them.
- **objects CR-014** (`e277637`, proposed; H1 to H4 Jason's) gives `HIGHLIGHT_COLORS` Word's palette, takes Office JS's
  input (CSS's 147 names in any case, or a hex in any case), stores a palette value's own name, uses check 61's ten
  measured snaps as data, and refuses anything else, `darkYellow` included. By its count all 63 cases then agree. When
  it is released, this package raises the range and empties `CHECK61_DIFFERS`.

### 12.2 Check 62: Word's rule for the highlight a colour becomes

Jason, 2026-10-24, relayed by the objects session: "write check 62", Word's rule for snapping a hex to the highlight
palette, "so we can see whether there is a simple rule". It decides objects CR-014's H4: whether `applyRunOptions` can
take any hex, as Office JS does, and store the name Word would. The objects session found no rule that fits check 61's
ten points: nearest in RGB (euclidean, Manhattan, by the largest channel), luma-weighted RGB, HLS distance and
per-channel thresholds to `00`, `80` and `FF` each miss at least two, always `#A9A9A9` (Word: `darkGray`, nearest
`lightGray`) and `#123456` (Word: `darkCyan`, nearest `darkBlue`).

**Check 62** (`test/fixtures/check62/`; the share's `__tmp/62/`; written and run 2026-10-24). The design is the
objects session's suggestion:
- **The file:** a title, then one paragraph per colour, its text the colour's six hex digits: 7,225 colours.
  - the grid, 16 levels per channel (`00`, `11`, ..., `FF`): 4,096;
  - the grey ramp in steps of 1, for where `black`, `darkGray`, `lightGray` and `white` meet: 256;
  - each channel's ramp with the other two at `00`, and again at `FF`: 768 each;
  - every colour within 8 of check 61's two counterexamples, per channel, in steps of 2: 729 each.
- **The snippet** sets each paragraph's `font.highlightColor` to its colour, 400 to a sync, and reads each back. A batch
  that fails is done one by one. For every 97th colour it also takes the `w:highlight` Word stored from the paragraph's
  `getOoxml()`, since check 61 showed the sixteen palette values and the names map one to one. Its output is compact,
  `[written, read back]` per colour. The writes change the document, which is closed without saving.
- **Through the Word shim** a colour of the objects table is stored as its name, and every other is refused. A test
  holds the file and snippet to each other and pins that.
- **The rule** is the objects session's to find, testing candidates against the run (distance metrics, colour spaces,
  Word's own 16-colour quantisation). If none fits every point, H4 falls back to the measured colours as a lookup table.

**Run 2026-10-24** (Jason; Word 365, 16.0.20326.20232, PC; Script Lab; no repair; closed without saving). The JSON is
`fixtures/check62/62-highlight-snap-word365.json`. Every colour was taken, with no refusal, and read back as one of the
sixteen. The 75 colours whose stored `w:highlight` was read hold the name of the colour read back, and each palette
colour among the points is itself. What the data shows outright, which a test pins:
- **The grey ramp** (`R = G = B`): `00` to `3F` black, `40` to `9F` darkGray, `A0` to `DF` lightGray, `E0` to `FF`
  white. The one exception is `#A9A9A9`, CSS's `darkgray`, which is darkGray inside lightGray's band (check 61's
  counterexample, alone: `#A8A8A8` and `#AAAAAA` are lightGray). So Word appears to know some colours by name before any
  rule applies.
- **One channel, the other two `00`:** any value but `00` gives the dark colour up to `C0` and the bright one above it
  (`#010000` is darkRed, not black; `#C10000` red).
- **One channel, the other two `FF`:** anything below `FF` gives the colour (`#FEFFFF` is cyan, not white).
- **Not nearest in RGB:** Word's answer is among the nearest palette colours, ties counted, for 4,393 of the 7,225
  (61%).
- **Not hue alone:** a hue-sector rule (the nearest of the six hues, dark at a value of `C0` and below, exact greys by
  the ramp's bands) misses about 2,000. Low-saturation colours go to a grey (`#8888BB` lightGray, `#333355` darkGray),
  and some near a sector's edge go to its neighbour (`#CC0011` magenta).

Sent to the objects session for the fitting (2026-10-24). **Its fit** (objects CR-014 section 2.1, unpushed when
written) agrees on 7,076 of the 7,225, `#A9A9A9`'s exception included (7,075 without it), and misses the other 149:
- **the hue** is Windows' `ColorRGBToHLS` integer hue (`HLSMAX` 240): red below 20, yellow 20 to 60, green 61 to 100,
  cyan 101 to 140, blue 141 to 179, magenta from 180. Every hue decision fits, `#0A325C` among them (cyan at integer
  hue 140, where the float hue is 210.7°);
- **dark** where `max + min` is `C0` or less, else bright: every decision fits;
- **greys** by `(max + min) / 2` at `40`, `A0` and `E0`, with `#A9A9A9` alone darkGray, matched by name first;
- **grey or colour:** an HLS saturation below 120 of 240 is the best single test, and still misses 149. No measure of
  saturation or chroma tried separates them: the edge bends with hue and lightness.

Its H4 recommendation became that rule plus check 62's 149 grey-edge answers as data: exact on every measured colour
(the 4,096 CSS three-digit colours among them), approximate only near the grey edge elsewhere. A check 63 of that edge
would follow only if a consumer needs it.

**Decided 2026-10-24** (Jason: "CR-015 accept recs for decisions H1 to H4"): objects CR-014's H1 to H4 as recommended.
- H1: one table, Word's palette.
- H2: Office JS's input.
- H3: CSS's 147 names.
- H4: the fitted rule plus check 62's 149 grey-edge answers as data, with no check 63 unless a consumer needs it.

Passed to the objects session the same day. When CR-014 is released, this package raises the range, empties
`CHECK61_DIFFERS`, and holds the shim to checks 61 and 62.

Objects CR-014 implemented 2026-10-24 at `ea8a5ba` (unpushed, unreleased), with checks 61 and 62 copied from `dc7e37d`
and `8a0945c`:
- `highlightHexValue` and `readRunOptions` read Word's palette;
- `highlightNameForColor` and `applyRunOptions` take a CSS Color 3 name in any case or any `#RRGGBB`, stored as Word stores
  it, and refuse anything else (`darkYellow`, `none`, a hex without `#`, `#FFF`, `rgb()`);
- its tests hold it to all 7,225 of check 62's colours and to check 61's 18 reads and 45 writes.
Its count above was corrected the same day (it had said 7,075 with the exception).

