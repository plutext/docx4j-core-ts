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

The harness resolves `org.docx4j:docx4j-core:17.2.0` (docx4j's `<revision>` on `VERSION_17_2_0`;
`pom.xml`'s `docx4j.version`, overridable with `-Ddocx4j.version=...` when the branch you build
says otherwise) and the JAXB implementation and four font jars from your local Maven repository,
so build them first from the docx4j checkout you want to measure — the sibling `../docx4j` on the
release branch, `VERSION_17_2_0` (the old `VERSION_17_1_1` was renamed to it on 2026-09-21).
A released version number is a trap a SNAPSHOT was not: if nothing was installed locally, Maven
takes Central's 17.2.0 release and the harness measures that, not the head; the workflow builds
first, and so must you:

```
cd ../docx4j          # ../../../docx4j from here
git rev-parse HEAD    # the hash you will pass to the harness
mvn -q -pl docx4j-xjc-copy,docx4j-core,docx4j-JAXB-ReferenceImpl,docx4j-export-fo-fonts-symbol,docx4j-export-fo-fonts-croscore,docx4j-export-fo-fonts-crosextra,docx4j-export-fo-fonts-theme2023 -am -DskipTests -Dgpg.skip install
```

`docx4j-core` is the engine, `docx4j-JAXB-ReferenceImpl` the JAXB implementation it needs at
run time, and the four font jars are the harness's font environment (below). `docx4j-xjc-copy`
is a plugin dependency of `docx4j-generated-objects`' schema compilation, which `-am` does not
reach (it follows project dependencies only); a fresh local repository fails without it, a
well-used one hides the omission.

**One local repository, more than one docx4j.** `17.2.0` in `~/.m2` is whatever was
installed there last, which is not necessarily the commit you are about to name. Every
golden's header therefore carries `docx4jCoreJarSha256`, the first sixteen hex digits of the
SHA-256 of the `docx4j-core` jar that actually ran, beside the `docx4jCommit` the harness was
told about. If the two disagree with what you expect, rebuild before believing a golden.

**Use a separate local repository whenever the docx4j checkout is shared.** If anyone else may
be building docx4j at the same time — another session in `../docx4j`, another worktree, a
colleague on the same machine — an `install` into `~/.m2` overwrites their `17.2.0`
and theirs overwrites yours, and a golden can end up made from a jar neither of you meant. Give
the build and the harness run a repository of their own, and do not touch `~/.m2` at all:

```
mvn -Dmaven.repo.local=/tmp/parity-m2 -q -pl docx4j-core,... -am -DskipTests -Dgpg.skip install
mvn -Dmaven.repo.local=/tmp/parity-m2 -q compile exec:java -Dfixtures=../fixtures -Dout=... -Ddocx4j.commit=...
```

The same directory for both, and not `-o` on the first run there: it has to fetch the rest of
the dependencies from the network once. Measuring a commit the checkout has moved past is the
same rule the other way round — export it read-only (`git archive <hash> | tar -x -C <dir>`)
and build from that, rather than checking it out in a repository someone else is using. The
scheduled workflow always has a repository of its own, so this does not arise there.

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
| `header` | harness version, docx4j commit, docx4j version and jar hash, the date (the only value that changes between two runs), the fixture's name and byte size, the font environment, the `mc:Choice` prefixes (`mcPreferChoice`), and `notes` — anything the harness had to catch while producing this golden |
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
- **Markup compatibility.** Since harness version 3 the branch is docx4j's own choice, not the
  harness's: `TraversalUtil` in its default `McMode.READ` gives up the one branch
  `org.docx4j.jaxb.McSelection` selects (docx4j CR-021 phase 1) — the first `mc:Choice` whose
  `Requires` prefixes are all named in `docx4j.jaxb.mc.preferChoice`, else the `mc:Fallback`.
  The harness sets that property, before any walk, to the prefixes of this package's
  `UNDERSTOOD_NAMESPACES` (`src/opc/mce/understood.mts`) through the objects package's
  `NAMESPACE_PREFIXES` table, which is `Harness.MC_PREFER_CHOICE` and is recorded in every
  golden's header as `mcPreferChoice`:

  ```
  a a13cmd a14 a15 a16 a1611 a16svg a18hc adec am3d an18 anam3d b c c14 c15 c16 c16ac c173
  cdr cdr14 comp cp cppr cs cx dc dcterms dgm dgm14 dgm1612 ds dsp iact ink16 lc m mc msink
  o p p13cmd p14 p15 p1510 p159 p16 p166 p1710 p173 p184 pic pic14 pkg prop properties psez
  pslz psuz pvml r rel sl thm15 v vt w w10 w14 w15 w16cid w16se we wetp wne wp wp14 wp15
  wpc wpg wps xdr xdr14 xvml
  ```

  Eighty-four prefixes. Nine understood namespaces have no entry in that table (MathML,
  InkML, the two Excel mains, the three encryption ones) and SpreadsheetML's entry there is
  the default namespace, so none of those is nameable; a `Requires` naming one would diverge,
  and no fixture has one. The TypeScript side resolves every `mc:AlternateContent` on the DOM
  before unmarshalling, taking the first choice whose `Requires` namespaces are in that set
  (`wps` among them), so naming the same prefixes here is what makes the two walk the same
  document — a text box once, from its `wps` Choice, which is what Word draws. The property
  goes through `Docx4jProperties`, which wins over anything the environment supplies, so the
  weekly workflow makes the same choice without setting anything of its own. Before CR-021
  phase 1 a `TraversalUtil` walk saw the choices *and* the fallback, which put a text box's
  paragraphs in a golden twice, and harness versions 1 and 2 took the first Choice themselves;
  docx4j's own mc-preprocessor was no help then either, since it keeps an
  `mc:AlternateContent` whose parent is a `w:r` — which is where Word puts a text box — and
  prefers the fallback elsewhere.
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
45 fixtures. Version 3 dropped the harness's own `mc:AlternateContent` branch selection for
docx4j's `McSelection` (above), and produced goldens identical to version 2's on all 45
fixtures too — so the harness now holds no rule of its own about what docx4j answers.

## Files

- `pom.xml` — Java 21, `docx4j-core` and friends from the local repository, `exec-maven-plugin`.
- `src/main/java/org/docx4j/parity/Harness.java` — the whole harness.
- `src/main/java/org/docx4j/parity/Json.java` — a small JSON writer, so the only dependency
  is docx4j.
