# The parity goldens

One JSON file per `.docx` fixture, recording what **docx4j** answers for it: the effective
paragraph, run and table properties from `PropertyResolver`, the list label and counters from
`Emulator`, and the document font of every character from `RunFontSelector`. They are the
contract `@docx4j/core-ts` meets in CR-001 Phase B — steps 2, 3 and 4 make the TypeScript
resolver, emulator and selector produce the same answers, and `test/parity.test.mjs` compares.

Produced by [`../java/`](../java/README.md), which documents the shape of a golden field by
field, the font environment they are made in, and the three things it reaches by reflection.
Read that first; this file is the review guide.

## Provenance

| | |
|---|---|
| docx4j | branch `VERSION_17_1_1`, commit `01d661547d26ca7b16c86860082e82fc8462dc3d` (2026-09-19; after the batch 49 merge `d5809a1d8` and the `w:numId` 0 fix) |
| docx4j-core jar | `docx4j-core-17.1.1-SNAPSHOT.jar`, SHA-256 `d0bd889c3a135e27…` |
| harness version | 2 (the public parity accessors; version 1 used reflection) |
| fixtures | the 8 `.docx` in `test/fixtures/` and the 37 in `test/fixtures/parity/` (see `test/README.md` for their provenance) |
| fonts | docx4j's symbol, croscore, crosextra and theme2023 jars alone; the machine's own fonts are not discovered |

## Regenerating them

```
cd ../docx4j
mvn -q -pl docx4j-core,docx4j-JAXB-ReferenceImpl,docx4j-export-fo-fonts-symbol,docx4j-export-fo-fonts-croscore,docx4j-export-fo-fonts-crosextra,docx4j-export-fo-fonts-theme2023 -am -DskipTests -Dgpg.skip install
cd ../docx4j-core-ts/test/java
mvn -q -o compile exec:java -Dfixtures=../fixtures -Dout=../golden -Ddocx4j.commit=$(git -C ../../../docx4j rev-parse HEAD)
```

About three seconds. Two runs differ only in the header's `date`; that was checked by running
twice into two directories and diffing, and again across two different builds of
`docx4j-core` — every answer identical, only `docx4jCoreJarSha256` moved.
`.github/workflows/parity.yml` does the same weekly against docx4j's head and opens a pull
request when an answer changes.

## What is here

| golden | stories | paragraphs | numbered | runs | font spans | styles | tables | fonts | KB |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| `HelloWordOnline` | 1 | 1 | 0 | 2 | 2 | 4 | 0 | 1 | 7 |
| `Mac_OSX_Fonts` | 1 | 8 | 0 | 62 | 62 | 4 | 0 | 5 | 56 |
| `NumberingImplicitNumId` | 7 | 19 | 6 | 30 | 24 | 81 | 0 | 4 | 119 |
| `NumberingIndents` | 9 | 30 | 8 | 49 | 27 | 12 | 0 | 2 | 101 |
| `article-section-NotIsLgl` | 1 | 8 | 6 | 6 | 6 | 23 | 0 | 6 | 49 |
| `article-section-isLgl` | 1 | 8 | 6 | 6 | 6 | 23 | 0 | 6 | 49 |
| `comments-two` | 2 | 4 | 0 | 9 | 5 | 11 | 0 | 1 | 18 |
| `fonts-author-had` | 1 | 3 | 0 | 6 | 6 | 22 | 0 | 3 | 26 |
| `fonts-cjk-linebox` | 1 | 36 | 0 | 36 | 40 | 22 | 0 | 9 | 109 |
| `fonts-cs-off` | 1 | 7 | 0 | 14 | 14 | 23 | 0 | 2 | 35 |
| `fonts-georgia` | 1 | 4 | 0 | 8 | 8 | 22 | 0 | 2 | 27 |
| `fonts-hebrew-no-cs` | 1 | 6 | 0 | 12 | 12 | 22 | 0 | 3 | 35 |
| `fonts-light-bold` | 1 | 4 | 0 | 8 | 8 | 22 | 0 | 3 | 27 |
| `fonts-missing-slots` | 1 | 5 | 0 | 10 | 39 | 22 | 0 | 3 | 35 |
| `fonts-modesOfApplication` | 1 | 9 | 4 | 10 | 10 | 10 | 0 | 7 | 36 |
| `fonts-segoe-ui` | 1 | 5 | 0 | 10 | 10 | 22 | 0 | 3 | 30 |
| `fonts-space-cjk` | 1 | 5 | 0 | 10 | 107 | 22 | 0 | 2 | 53 |
| `fonts-symbol-and-emoji` | 1 | 6 | 0 | 12 | 12 | 22 | 0 | 5 | 31 |
| `fonts-theme-lang` | 1 | 2 | 0 | 4 | 4 | 22 | 0 | 1 | 24 |
| `fonts-unresolvable` | 1 | 8 | 0 | 16 | 16 | 22 | 0 | 7 | 39 |
| `header-no-rels` | 6 | 13 | 0 | 11 | 6 | 10 | 0 | 1 | 26 |
| `hyperlink_dupe` | 1 | 3 | 0 | 10 | 5 | 14 | 0 | 2 | 21 |
| `invoice` | 2 | 14 | 0 | 15 | 11 | 13 | 1 | 1 | 33 |
| `invoice2013` | 1 | 27 | 0 | 41 | 40 | 26 | 2 | 2 | 72 |
| `loadAndSave` | 10 | 60 | 0 | 106 | 56 | 63 | 1 | 6 | 192 |
| `numbering-default-style-numbered` | 1 | 5 | 3 | 10 | 10 | 23 | 0 | 2 | 36 |
| `numbering-label-ilvl0` | 1 | 29 | 13 | 58 | 58 | 28 | 0 | 2 | 123 |
| `numbering-leader-kinds` | 1 | 7 | 5 | 14 | 14 | 22 | 0 | 2 | 47 |
| `numbering-level-pstyle-in-ppr` | 1 | 12 | 6 | 24 | 24 | 24 | 0 | 2 | 57 |
| `numbering-level-vs-style-indent` | 1 | 8 | 6 | 16 | 16 | 25 | 0 | 2 | 47 |
| `numbering-lvlrestart` | 1 | 18 | 16 | 36 | 36 | 22 | 0 | 2 | 89 |
| `numbering-numstylelink-separate` | 1 | 10 | 9 | 20 | 20 | 23 | 0 | 2 | 57 |
| `numbering-override-rpr` | 1 | 8 | 6 | 16 | 16 | 22 | 0 | 2 | 48 |
| `numbering-shared-abstract` | 1 | 7 | 6 | 14 | 14 | 22 | 0 | 2 | 44 |
| `numbering-stories-coretests` | 6 | 36 | 27 | 78 | 65 | 22 | 0 | 2 | 143 |
| `numbering-stories` | 6 | 36 | 27 | 78 | 65 | 22 | 0 | 2 | 143 |
| `numbering_indentation` | 1 | 12 | 4 | 59 | 59 | 5 | 0 | 2 | 68 |
| `numbering_indentation_firstline` | 1 | 13 | 5 | 71 | 71 | 7 | 0 | 3 | 93 |
| `startOverride` | 1 | 13 | 12 | 12 | 12 | 23 | 0 | 4 | 71 |
| `styles-default-pstyle` | 1 | 7 | 0 | 10 | 10 | 24 | 0 | 3 | 36 |
| `styles-linerule` | 1 | 3 | 0 | 3 | 3 | 23 | 0 | 1 | 26 |
| `styles-no-size-anywhere` | 1 | 4 | 0 | 4 | 4 | 22 | 0 | 1 | 26 |
| `styles-numpr-ilvl-only` | 1 | 6 | 5 | 7 | 7 | 24 | 0 | 2 | 37 |
| `styles-table-default` | 1 | 9 | 0 | 12 | 12 | 24 | 3 | 2 | 38 |
| `tracked-changes` | 4 | 80 | 5 | 68 | 60 | 18 | 1 | 7 | 140 |
| **total (45)** | **88** | **618** | **185** | **1113** | **1112** | **984** | **8** | | **2617** |

2.6 MB in all, the largest golden 192 KB (`loadAndSave`). Nothing is truncated: a golden
records every style in the styles part, every paragraph of every story, every run of every
paragraph and every font span of every run, because a parity failure is usually in the case
nobody thought to keep.

## Reading them the first time

The point of this review is to decide whether these are the right answers to hold the port
to. Things worth a look, and things found while producing them:

- **Nothing threw.** Every `notes` array is empty: no style resolution raised
  `CyclicStylesException`, no `Emulator` call failed, no font selection failed, on any of the
  45 fixtures. (The harness records an exception's message in place of the value and lists it
  in `header.notes`, so a cyclic-styles fixture would be a legitimate golden rather than a
  gap. There is no such fixture yet; CR-015's cyclic cases live in docx4j's own tests.)
- **The story rule.** `numbering-stories.json` is the one to read: the body numbers 1 2 3, the
  text box starts again at 1 2 3, and the body then continues 4 5 6 — the text box's
  paragraphs sit in the main story's list, in document order, but count in a story of their
  own. The header runs 1 2 3 and the footer 4 5 6 from the state they share, and the
  footnotes, endnotes and comments parts each count from 1. That is exactly docx4j's
  `NumberingStoriesTest`, which CR-014 probe P7 measured in Word.
- **`stateAfter`.** Every numbered paragraph carries the story's counters after it, keyed
  `<abstractNumId>/<ilvl>`, with `encounteredAlready` and `resetPending` — and the `w:num`
  start overrides spent so far. `startOverride.json` shows two overrides spent (`1/0`, `2/0`)
  and the deeper levels sitting at their start with a reset pending. This is what makes a
  counting difference in step 3 land on the paragraph that caused it.
- **Font spans are real.** `fonts-missing-slots` folds ten runs into thirty-nine spans:
  "café über façade" splits at each accented character, Liberation Sans for the ASCII and
  Times New Roman for the rest, which is the ascii/hAnsi slot dispatch of
  [MS-OI29500] 17.3.2.26. `fonts-space-cjk` gives 107 spans from 10 runs.
- **Tables.** Only 8 tables in the whole corpus, 3 of them in `styles-table-default`, which is
  CR-015 phase 4's probe: a table naming no style and one whose chain reaches the default
  table style both get Word's built-in Normal Table (`w:tblInd` 0, cell margins 108/0/108/0)
  and `reachesDefaultTableStyle: true`; a table whose chain does not reach it gets no cell
  margin at all and `false`. There is no docx4j API for that flag — the harness recomputes it
  from the style chain (see the harness README's reflection table).
- **Deleted text is kept.** `tracked-changes.json` records the run inside `w:del` with
  `deleted: true` and its text (" A deletion"), rather than dropping it: `w:delText` is its
  own class in the object model, not a `w:t`, and a port that forgot it would silently lose
  half a tracked change.
- **The default font, and the theme.** 28 of the 45 fixtures have **no theme part** — the
  probe documents are minimal — and `fonts.themePart` / `fonts.defaultFont` record that
  explicitly, because a theme reference that cannot be answered changes every font span in
  the document. Sixteen of those 28 resolve `minorHAnsi` to **Aptos**, which is what Word 365
  does with a themeless document and what docx4j answers since CR-001 batch 49
  (`docx4j.fonts.defaultTheme`, left at its default here). These goldens come from the merged
  branch, so that is the branch's own answer and not a preview of it.
- **Mapping is meaningful, not a row of nulls.** With only docx4j's font jars available the
  `IdentityPlusMapper` still decides: Aptos → Akasia and Aptos Display → Intos Display by
  metric clone (the theme2023 jar), Calibri → Carlito, Arial → Arimo, a
  `w:altName` chain (`fonts-unresolvable`'s Probe D → Arial → Arimo), Word's own default for a
  face it cannot find (Probe B → Calibri → Carlito), a face of the same class
  (Liberation Serif → Tinos), and `UNMAPPED` where nothing stands in (Malgun Gothic). That
  covers every pass of CR-016 phase 3's precedence template, which is what step 4 ports.
- **`mc:AlternateContent`.** Resolved to its first `mc:Choice`, so a text box appears once
  rather than twice (the wps choice and its VML fallback). docx4j's own consumers make the
  same choice; docx4j's mc-preprocessor does not, and would give the fallback. See the harness
  README.
