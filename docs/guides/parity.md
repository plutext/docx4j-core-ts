# Parity with docx4j

This package is called docx4j for TypeScript, and the claim is checked rather than asserted.

## What the goldens are

A small Java harness ([`test/java/`](../../test/java)) loads every `.docx` fixture with **docx4j
itself** and writes down what docx4j answers, as one JSON file per fixture under
[`test/golden/`](../../test/golden):

- for every style in the styles part: the effective `w:pPr` and `w:rPr`, and the document defaults;
- for every paragraph of every story (the body, each header and footer, the footnotes, the
  endnotes, the comments), in document order: the effective `w:pPr`, the paragraph mark's `w:rPr`,
  and the numbering — the label, whether it is a bullet, the label font, the indent, the level,
  and the story's counter state after that paragraph;
- for every run: the effective `w:rPr`, and the document font of every character span, with the
  `IdentityPlusMapper`'s physical font for it;
- for every table: the effective table style, and whether its chain reaches the default.

The property values are XML strings marshalled by docx4j. `test/parity.test.mjs` unmarshals them
through the object model and compares **object trees**, never text, against what
`PropertyResolver`, the numbering `Emulator` and `RunFontSelector` answer here for the same
document. 45 fixtures, zero differences, against docx4j `VERSION_17_1_1` commit `7fba7a150`.

Each golden's header records the docx4j commit, the SHA-256 prefix of the `docx4j-core` jar that
actually ran, the date, the fixture and its size, and the font environment the mapping was made
in (docx4j's four font jars, so the answer does not depend on what the machine has installed).

## Running it yourself against a newer docx4j

Java 21 and Maven 3.9, and a docx4j checkout. Build docx4j and the jars the harness needs into a
Maven repository of its own, then run the harness into a temporary directory and compare:

```
cd ../docx4j
git rev-parse HEAD
mvn -Dmaven.repo.local=/tmp/parity-m2 -q -DskipTests -Dgpg.skip -am install \
  -pl docx4j-core,docx4j-JAXB-ReferenceImpl,docx4j-export-fo-fonts-symbol,\
docx4j-export-fo-fonts-croscore,docx4j-export-fo-fonts-crosextra,docx4j-export-fo-fonts-theme2023

cd ../docx4j-core-ts/test/java
mvn -Dmaven.repo.local=/tmp/parity-m2 -q compile exec:java \
  -Dfixtures=../fixtures -Dout=/tmp/golden-new \
  -Ddocx4j.commit=$(git -C ../../../docx4j rev-parse HEAD)

diff -r ../golden /tmp/golden-new      # header lines aside, a difference is docx4j's answer moving
```

Point `-Dout` at `../golden` instead to refresh the committed goldens, then run `npm test`: any
difference is either docx4j having changed its answer, or this port having to follow it.

[`test/java/README.md`](../../test/java/README.md) has the detail — why the harness uses a
separate Maven repository, how to read a golden, and what each field means.

## The weekly check

[`.github/workflows/parity.yml`](../../.github/workflows/parity.yml) does the above on a schedule
against `plutext/docx4j` `VERSION_17_2_0` (the release branch; `VERSION_17_1_1` was renamed to it
on 2026-09-21), and opens a pull request when a golden differs, with
docx4j's commits since the recorded hash in its body. So a change in docx4j's behaviour arrives
here as a reviewable diff rather than as a surprise.

## What parity does not cover

The goldens are WordprocessingML resolution: effective properties, list labels, font selection.
Packaging is held to a different contract — untouched parts byte-identical after a round trip,
re-marshalled parts deep-equal after reload — by the rest of `test/`, and rendering fidelity is
not in this package at all. `test/README.md`'s Word acceptance checklist is the manual part:
saved output opening in Word, PowerPoint and Excel without a repair prompt.
