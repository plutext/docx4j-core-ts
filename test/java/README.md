# The Java parity harness

`@docx4j/core-ts` claims to be docx4j for TypeScript. This is how the claim is checked: a
small Maven project that loads every `.docx` fixture with **docx4j itself** and writes down
what docx4j answers — the effective paragraph, run and table properties, the list label and
counters, the document font of every character — as one JSON golden per fixture under
[`../golden/`](../golden/). The TypeScript port is then measured against those goldens
(`test/parity.test.mjs`), so the contract is docx4j's behaviour rather than anyone's reading
of it.

CR-001 section 14.3 specifies the harness and the shape of a golden; sections 6 and 14.2 say
what the port will compute. docx4j's own CR-014 (list numbering), CR-015 (property
resolution) and CR-016 (font selection and mapping) are what that behaviour was settled
against, Word-verified.

It is not part of the npm package and `npm test` does not run it. It is run by hand when the
fixtures or docx4j change, and weekly by
[`.github/workflows/parity.yml`](../../.github/workflows/parity.yml).

## Build docx4j

The harness resolves `org.docx4j:docx4j-core:17.1.1-SNAPSHOT` (and the JAXB implementation
and three font jars) from your local Maven repository, so build them first from the docx4j
checkout you want to measure — the sibling `../docx4j` on branch `VERSION_17_1_1`:

```
cd ../docx4j          # ../../../docx4j from here
git rev-parse HEAD    # the hash you will pass to the harness
mvn -q -pl docx4j-core,docx4j-JAXB-ReferenceImpl,docx4j-export-fo-fonts-symbol,docx4j-export-fo-fonts-croscore,docx4j-export-fo-fonts-crosextra,docx4j-export-fo-fonts-theme2023 -am -DskipTests -Dgpg.skip install
```

`docx4j-core` is the engine, `docx4j-JAXB-ReferenceImpl` the JAXB implementation it needs at
run time, and the four font jars are the harness's font environment (below).

**One local repository, more than one docx4j.** `17.1.1-SNAPSHOT` in `~/.m2` is whatever was
installed there last, which is not necessarily the commit you are about to name. Every
golden's header therefore carries `docx4jCoreJarSha256`, the first sixteen hex digits of the
SHA-256 of the `docx4j-core` jar that actually ran, beside the `docx4jCommit` the harness was
told about. If the two disagree with what you expect, rebuild before believing a golden.

## Run it

Java 21 and Maven 3.9.

```
cd test/java
mvn -q -o compile exec:java \
  -Dfixtures=../fixtures \
  -Dout=../golden \
  -Ddocx4j.commit=$(git -C ../../../docx4j rev-parse HEAD)
```

- `-Dfixtures` the fixtures directory: every `.docx` directly in it and every `.docx` under
  its `parity/` subdirectory (`.pptx`, `.xlsx` and loose `.xml` are skipped).
- `-Dout` where the goldens go; `../golden` to refresh the committed ones, a temporary
  directory to compare without overwriting.
- `-Ddocx4j.commit` recorded in every header. The harness never runs git itself.

About three seconds for 45 fixtures. It prints a line per fixture and exits non-zero if any
fixture failed.

## Reading a golden

One file per fixture, pretty printed, keys in a fixed order (insertion order, not
alphabetical, so a paragraph reads top to bottom). Five sections:

| section | what it holds |
|---|---|
| `header` | harness version, docx4j commit, docx4j version and jar hash, the date (the only value that changes between two runs), the fixture's name and byte size, the font environment, and `notes` — anything the harness had to catch while producing this golden |
| `styles` | `defaultParagraphStyleId`, the document defaults (`pPr`, `rPr`), the `w:docDefaults` element as the part states it, and `byId`: for every style id, its type, `basedOn`, whether it is a default, and `effectivePPr(styleId)` / `effectiveRPr(styleId)` |
| `stories` | one entry per story — `main`, `header:<relId>`, `footer:<relId>`, `footnotes`, `endnotes`, `comments` — each with its part name and every `w:p` in document order |
| `tables` | every `w:tbl` of the main story: `tblStyle`, `effectiveTableStyle`, `reachesDefaultTableStyle` |
| `fonts` | the theme part and the default font it yields, `fontsInUse()`, `stylesInUse()`, and the `IdentityPlusMapper` decision per document font |

A paragraph carries `index` (0-based within the story), `paraId` (`w14:paraId` or null),
`pStyle` (the direct one or null), `text` (its first 80 characters, as a human check on the
address), `effectivePPr`, `paragraphMarkRPr`, `numbering` and `runs`. A run carries `text`,
`rStyle`, `deleted` (only when it is inside a `w:del`), `effectiveRPr` and `fontSpans`. A
`numbering` is null where docx4j does not number the paragraph, and otherwise carries
`numString`, `isBullet`, `numFont`, `ind`, `ilvl`, `numId`, `labelRPr`, `lvl`, `indResolved`
(`NumberingDefinitionsPart.getInd`, which follows a level's linked style), `numRef` (docx4j's
resolution: `numId`, `ilvl`, `direct`, `notNumbered`, `reason`) and `stateAfter` (the story's
counters once this paragraph has taken its number).

Every property value is **XML**, marshalled by docx4j with its own namespace prefixes and no
declaration. The TypeScript side unmarshals it through the objects facade and compares object
trees, so neither formatting nor prefixes are part of the contract — only content. The
namespace declarations no element or attribute uses are stripped before the string is
recorded: docx4j's prefix mapper pre-declares all ninety-odd Office namespaces on whatever it
marshals, which was 3 kB of `xmlns:` on each of tens of thousands of fragments and ten times
the size of the goldens.

## What makes a golden the same everywhere

A golden is committed and diffed weekly, so two runs on two machines must agree.

- **Fonts.** `PhysicalFonts` discovery scans the machine, and the `IdentityPlusMapper`'s
  answer depends on what it finds. The harness turns system discovery off
  (`docx4j.fonts.discoverPhysicalFonts.enabled=false`) before anything constructs a `Mapper`,
  leaves jar discovery on, and puts exactly the symbol, croscore, crosextra and theme2023
  font jars on the classpath: CR-016 phase 0c's `-Dfidelity.fonts=jars` environment, the one
  it measured a headless container in, with the theme2023 jar (Akasia, Intos Display) that
  CR-001 batch 49 added for Word 365's Aptos and Aptos Display. The font cache goes to a
  fresh temporary directory each run so that a cache left by something else cannot reach a
  golden.
- **Markup compatibility.** An `mc:AlternateContent` is walked as its first `mc:Choice`, or
  its `mc:Fallback` where it has no choice. `TraversalUtil` walks the choices *and* the
  fallback, which would put a text box's paragraphs in a golden twice; docx4j's own
  mc-preprocessor is no help, since it keeps an `mc:AlternateContent` whose parent is a `w:r`
  — which is where Word puts a text box — and prefers the fallback elsewhere. The TypeScript
  side resolves every `mc:AlternateContent` on the DOM before unmarshalling, taking the first
  choice whose `Requires` namespaces it knows (`wps` among them), so this is the rule that
  makes the two walk the same document.
- **Dates.** Only the header carries one, and the weekly workflow ignores it (and the jar
  hash) when it diffs.

Verified by running the harness twice into different directories and diffing.

## Stories, and text boxes

One `NumberingState` per story, from `NumberingStates.forPart(part)` as CR-014 phase 4
defines it: the body is one story, a section's header and footer share one, the footnotes,
endnotes and comments parts have their own. A text box gets a fresh state
(`NumberingStates.newStory()`), which is what `AbstractWmlConversionContext.enterTextBox`
does. Its paragraphs stay in the containing part's paragraph list, in document order — they
are in that part — but they number on their own, and the containing story's count runs past
them untouched. `numbering-stories.docx` is the probe: body 1 2 3, text box 1 2 3, body 4 5 6,
header 1 2 3, footer 4 5 6, each note part from 1.

## What the harness asks docx4j for

Four things the port has to reproduce have no other way in, and docx4j 17.1.1 exposes each
of them (CR-001 batch 49, `ParityAccessorsTest`):

| asked for | what it gives |
|---|---|
| `Emulator.numRefFor(pkg, pPr)` | `NumRef` — which `numId` and `ilvl` a paragraph resolves to, whether the `w:numPr` was direct, and why a paragraph is not numbered. Never null, and it takes no number, so it can be asked before `getNumber` increments |
| `NumberingState.counters()` and `.startOverridesApplied()`, with `ListLevel.Counter.getCurrentValue()` / `isEncounteredAlready()` / `isResetPending()` | `stateAfter`: the counters a story stands at, which is how a counting bug is localised to the paragraph that caused it rather than the one that showed it |
| `PropertyResolver.reachesDefaultTableStyle(tblPr)` | whether Word's built-in Normal Table underlies a table (CR-015 phase 4). Its only other sign is the cell margins in the resulting style |
| `FontsAnalysis.NO_OP_VISITOR` | the visitor `RunFontSelector`'s constructor requires and `documentFontFor` never calls |

Harness version 1 reached the first three by reflection (`Emulator.resolve`,
`NumberingState`'s private maps, `PropertyResolver.ancestry`) and carried a no-op visitor of
its own; version 2 calls the accessors, and produced goldens identical to version 1's on all
45 fixtures.

## Files

- `pom.xml` — Java 21, `docx4j-core` and friends from the local repository, `exec-maven-plugin`.
- `src/main/java/org/docx4j/parity/Harness.java` — the whole harness.
- `src/main/java/org/docx4j/parity/Json.java` — a small JSON writer, so the only dependency
  is docx4j.
