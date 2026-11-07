# CR-030: A prefix in `mc:Ignorable` with no declaration, read as naming nothing (WD REQ-323)

**Status:** Implemented 2026-11-08, at Jason's "REQ-318 to REQ-320 and REQ-323 as recommended", the recommendation being the portfolio's that all three engines build REQ-323. Measured here the same day on the portfolio's `undeclared-ignorable` round (section 1), and committed at his "commit" the same day. D1 is Jason's.

## 0. Summary

A prefix in `mc:Ignorable` that no declaration in scope on the element carrying the attribute binds names no namespace. WD 0.103's REQ-323 (portfolio `eceabd6`) reads it as though the list didn't hold it, at every value of `repair`, raises `mce.undeclared_prefix` once a part, and on a save has a part written anew hold no such prefix, dropped or declared with a namespace the engine knows, where a part copied from its source keeps it. This package already read and saved such files as REQ-323 has it, and said nothing. It now says so, and its `part.skipped` no longer calls an element ignorable by a prefix declared only below the list.

## 1. Measured (2026-11-08, `0b4be29`, objects 0.8.0)

The round, 13 files, each opened twice, the runs agreeing, Word 16.0.20326.20232 (`test/undeclared-ignorable/`). Word repairs a file whose list holds such a prefix, one prompt an open, and loses nothing.

| Files | Here before this CR |
|---|---|
| `ui-01` to `ui-04`, `ui-07`, `ui-10` to `ui-14` | Each read with Word's text by both load routes, nothing said |
| Every file, a save untouched | Each part as written, the prefix kept |
| A save after `unmarshalAll` | An unknown prefix (`zz9`) dropped from the list; a known one (`w14`, `w15`, `wp14`) kept and declared, from the objects package's prefix table. Both meet REQ-323's save rule as WD 0.103 words it, by its outcome |
| `ui-05` (Word refuses: the element declares the prefix for itself) | Read, the element a `part.skipped` with `ignorable: true` |
| `ui-06` (silent: the prefix declared on the root) | Word's text; a save of the read tree refused by CR-025's gate for the element dropped |

So the gap was the warning, and `ui-05`'s flag. `isIgnorable` counted a prefix that didn't resolve where the list stands by its name, for a branch `mcePreprocess` takes out of the `mc:Choice` declaring it, and so also counted one declared on the element itself.

## 2. What landed

- **The warning.** `McUndeclaredPrefixWarning` (`code` `mce.undeclared_prefix`, `partName`, `prefixes`, `message`) on the part and on `pkg.warnings`, from the markup compatibility pass, which already meets every element a reader can reach, in both load modes and in every part. Each prefix once, in document order. A list is read where its element stood in the source: a resolving load moves a branch's content out of the `mc:Choice` that may declare its prefixes, and Excel writes just that, `a14` declared on the Choice and listed on an `a:srgbClr` inside (`cr022-checkbox-linked.xlsx`, whose kept and resolved loads `mce.test.mjs` holds to the same warnings). So before a branch is moved, each list in it is read in place.
- **The text model** raises `TextWarning.MCE_UNDECLARED_PREFIX` once a part it took, no address, the part's name the selector and `detail` the prefixes separated by spaces: a story part, or the styles, numbering or settings part the reading consults.
- **`part.skipped`'s `ignorable`.** A prefix that doesn't resolve where the list stands counts by its name only where the item's own prefix doesn't resolve either, which is a branch taken out, and not where it was declared below the list.
- **Not changed:** the reading, which `mc:Ignorable` doesn't affect here, and the saves, which already meet the rule.

## 3. Tests and oracles

`test/undeclared-ignorable.test.mjs`: the round's sources and Word's saves, copied unedited with a SHA-1 manifest. The warning per part held to a table read from the files by hand, by both load routes, on the part and in the text model; each file Word opened held to Word's text by both routes; `ui-05` and `ui-06`'s flags; every part of an untouched save as written, and every part written anew holding no undeclared prefix, an unknown one dropped and a known one declared; and the pass on a list below the root, a leading space and a declaration on a descendant, a list in a branch under both load modes, and one in nested branches. Seven mutants of the pass, `isIgnorable` and the text model each failed a test, one only after the nested case was written; `mce.test.mjs`'s comparison of kept and resolved loads found the branch case on a real Excel drawing.

Across 2,233 packages (this repository's `test/` and docx4j's test resources) the warning comes only from the `ui-*` files, docx4j holding a copy, and `ignorable` changes only on `ui-05`. The 1,224 saves compared are the same byte for byte.

## 4. Decisions

| | Question | Recommendation |
|---|---|---|
| D1 | Which release? | 0.7.0. It adds a warning and changes no reading or save, and docx4j and python build it now |

## 5. Registry

`core-ts/CR-030`, no dependency. Sent with its commit.

## 6. Sources

- `docx4j-portfolio/docs/text_extraction_specification_wd.md`, WD 0.103 at `eceabd6`: REQ-323 and its NOTE, REQ-801's row; `repair_policy_across_ports.md` section 5.
- The share's `bridge/staging/undeclared-ignorable/` and `bridge/portfolio/ui-*`.
- `src/opc/mce/preprocessor.mts`, `src/opc/skipped.mts`, `src/model/text/Extractor.mts`.
