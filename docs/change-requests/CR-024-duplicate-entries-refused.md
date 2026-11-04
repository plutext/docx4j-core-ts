# CR-024: A package holding two entries of one name is refused (WD REQ-206), with `duplicateEntries: 'first'`

**Status:** Proposed 2026-11-04, at Jason's "Duplicate zip entries (REQ-206): go", on the WD's REQ-206 (0.74, Jason's
ruling "refuse the package, with an optional config option to read the first copy and warn", made in this session; its five
details ruled as recommended in 0.77, portfolio `b0d333d`, from this session's suggestions). Measured here the same day
(section 1). Nothing implemented. D1 to D7 decided 2026-11-04, all as recommended, subject to Word check 75 on a non-ASCII
case pair, which decides the key's fold (section 5). docx4j-python's sibling is its CR-032 (proposed,
`778b3bf`), whose names this CR follows where the two packages' shapes allow; docx4j has not built it.

## 0. Summary

A zip can hold two entries with one name, and OPC forbids it: two part names are one where they are equivalent, compared
case-insensitively. Readers split on such a file. Word 365 reads the first entry, but only after its unreadable-content
prompt and a repair, and its repaired save keeps one copy. POI, Tika, the Open XML SDK and LibreOffice refuse it. docx4j,
python-docx and this package read the last copy silently. A reader that silently reads a different entry from the one
Word shows reads a different document, which is what such a duplicate is made for (the WD's NOTE: Android's 2013 "Master
Key" bug is the precedent).

REQ-206 refuses such a package by default, naming the entry. Under `duplicate_entries: first` an implementation reads the
first copy of each name in the zip's central directory, raises `package.duplicate_entry` once per name, and a save writes
only the copy read. The rule covers every entry, `[Content_Types].xml` and the relationships parts included.

## 1. Here today (measured 2026-11-04, `f769515`)

Each zip was built with entries in the order shown, stored, through a scratch writer (fflate's own refuses a repeated
name), and loaded with `WordprocessingMLPackage.load`, read with `Texts.of` and saved untouched.

| The zip | Read | Saved | Said |
|---|---|---|---|
| `word/document.xml` twice (FIRST, SECOND) | SECOND | one `word/document.xml`, SECOND | nothing |
| three copies (FIRST, SECOND, THIRD) | THIRD | one, THIRD | nothing |
| `word/document.xml`, then `Word/Document.xml` (DECOY) | FIRST, by its exact name | the decoy dropped | nothing |
| `Word/Document.xml` (DECOY), then `word/document.xml` | SECOND, by its exact name | the decoy dropped | nothing |
| `[Content_Types].xml` twice, the second naming the main part a template | the second's types: the main part typed `template.main+xml` | the second's types | nothing |
| `_rels/.rels` twice, the second naming `word/other.xml` | `word/other.xml` | `word/document.xml` gone | nothing |
| an entry no relationship names, twice | (not read) | the orphan dropped | nothing |

- **The store.** `ZipPartStore.readCentralDirectory` keys its entries by exact name in a `Map`, so the last of a name
  wins. Its `partNames()` lists every entry, duplicates included, while `has` and `load` reach only the last.
- **The case variant.** It is a second part name here: lookups are exact, so whichever entry the relationship's target
  spells is read, and the other is an orphan, dropped on save.
- **Flat OPC.** `FlatOpcPartStore` keys its `pkg:part`s by `pkg:name` in a `Map` the same way: the last of a name wins,
  silently.
- **The loss on save.** A duplicated relationships part loses a part on an untouched save. That is the WD NOTE's
  `two-rels/` finding, here and in docx4j.

## 2. Proposal

### 2.1 One name, in the store (D1, D3)

Each store this package ships keys its entries by the OPC key, `PartName`'s own (`key`, the name lower-cased): any two
names the package would treat as one part are one entry.
- **The first wins.** The first entry of a key keeps its spelling, its position and its bytes. The store counts the copies
  per key, and `duplicates()` returns `{ first spelling: count }` in the container's order. `PartStore` gains it as an
  optional method.
- **Lookups fold case** (D3). `has`, `load` and `size` look the key up, and `partNames()` lists each key's first spelling
  once. After the check every key is unique, so the fold is never ambiguous. It also lets a relationship whose target
  differs from its entry only in case find its part, where today the load refuses "cannot find part".
- **A named departure from OPC's letter.** OPC [M1.12] compares part names as case-insensitive ASCII strings, but
  `toLowerCase()` folds non-ASCII letters too (`É` and `é`). The key is this package's equality (`PartName.equals`), so a
  pair it folds is one name here, and refusing it is the safe reading, as docx4j-python's CR-032 has it. Word's reading of
  a non-ASCII case pair is unmeasured: check 75 measures it (section 5).
- **Not folded.** Backslashes and percent-encoding are not in the key, as today: two spellings that differ only that way
  never confuse the store, and they are not refused.

### 2.2 Every store (D2)

- **Zip:** the central directory's order, not the local headers'.
- **Flat OPC:** the `pkg:part` order. A flat OPC string to `insertOoxml` is a `FlatOpcPartStore`, so a fragment with two
  `pkg:part` of one name is refused too.
- **Directory** (`./node`): the walk's order. Only case variants can collide, on a case-sensitive file system.
- **Memory:** the constructor's entries, in their order. `put`, a sink's write, still replaces.
- **A caller's own `PartStore`** without `duplicates()` is not checked, and the interface's doc says so.

### 2.3 The refusal (D4)

`loadPackage` asks the store for its duplicates before it reads `[Content_Types].xml`. Under the default it throws
`DuplicateEntryException`, a new subclass of `InvalidFormatException`:

> Duplicate entry word/document.xml in this package: 2 entries (REQ-206). Load with { duplicateEntries: 'first' } to read
> the first.

Every duplicated name is listed, in the container's order. The exception carries `code` (`package.duplicate_entry`) and
`entries` (the `{ name: count }` map), so a host can act on it without parsing the message. Nothing has been read at that
point.

### 2.4 The option and the warning (D5, D6)

- **The option.** `LoadOptions.duplicateEntries`: `'refuse'`, the default, or `'first'`, the WD's `duplicate_entries`.
  Any other value is refused with `Docx4JException` naming REQ-206. `clone()` carries it with the other load options.
- **The record.** Under `'first'` each duplicated name is a `DuplicateEntryWarning` on the package:
  - `code`: `package.duplicate_entry`;
  - `entry`: the first copy's spelling (`word/document.xml`, `[Content_Types].xml`), docx4j-python's field;
  - `count`: the number of entries, python's field;
  - `message`.
- **Where it shows.** `pkg.warnings` lists these first, in the container's order, ahead of the parts' warnings, so its
  element type widens from `PartWarning` to `PackageWarning`, the union of both.
- **The text model.** Every reading of the package carries each record into `TextResult.warnings` as
  `TextWarning('package.duplicate_entry')`, the selector the entry. `TextWarning` gains `detail`, the count, which
  `toJSON` includes only where set. That field is what REQ-801's relationships rule names too ("the number kept in
  `detail`").
- **`duplicate_entries` in `TextOptions`.** The name is accepted with its two values, as REQ-1001 requires of the WD's
  options table (0.74's row). It is documented as the load's option: a reading of a package already loaded doesn't decide
  again, and the load's setting holds. The suite's adapter (`test/text-extraction/adapter.mjs`) passes a scenario's value
  to the load.

### 2.5 The save

Under `'first'`, a save writes only the copy read, with no change to `Save.mts`:
- an untouched part's bytes come from `store.load`, which is now the first copy;
- `[Content_Types].xml` and the relationships parts are written from what was parsed, which is the first copy;
- an entry no relationship names is not carried, as today.

### 2.6 Packages inside packages

An embedded document, or an `altChunk` that is a package, holding a duplicate is refused by its own load under the
default. The text model already reads one that will not load as `content.not_read` or `altchunk.not_expanded`.

### 2.7 Not here

- **Several main parts.** A package whose package relationships hold two officeDocument relationships, where ECMA-376
  wants one, is a neighbouring case. The WD's NOTE has Word, through a repair, read the target whose `main+xml` Override
  is last in `[Content_Types].xml`, and recommends a statement (refuse, and under the same switch read that target, with
  `package.multiple_main_parts`) still open with Jason. Here `getMainDocumentPart()` and the readers take the last
  relationship's target, and `getMainPart()` and the package-type check take the first: a defect either way (measured
  2026-11-04, recorded under REQ-206's NOTE). If Jason rules the statement, it fits this CR's option as a phase 2.
- **The decoy first, under `'first'`.** With `Word/Document.xml` first and `word/document.xml` second, `'first'` reads
  the decoy's bytes for the part the relationship names, where an exact-name reading takes the real one. The default
  refuses both orders. Word's reading of that order is unmeasured: the identifier's `extraction-probe-dupcase` has the
  decoy second, where both rules agree, and Word read the probe and dropped the decoy (portfolio `c70c9c4`). The WD rules
  first-in-the-central-directory (detail (b)), so this CR follows it, and the question goes with the identifier's
  follow-up (docx4j CR-075 D9).
- **The repair policy's switch.** If the WD later folds `duplicate_entries: first` into a general `repair` switch (the
  portfolio's repair policy, P4), this option stays as its alias.

### 2.8 The CHANGELOG lines

- **Changed:** a package holding two entries of one name, part names compared without case, is refused with
  `DuplicateEntryException` naming each entry. Previously the last copy was read silently, and an untouched save could
  drop a part. A relationship whose target differs from its entry only in case now finds its part.
- **Added:** `LoadOptions.duplicateEntries: 'first'`; `DuplicateEntryWarning` (`package.duplicate_entry`) on
  `pkg.warnings` and in `TextResult.warnings`; `TextWarning.detail`; `PartStore.duplicates()`.

## 3. Tests and oracles

`test/duplicate-entries.test.mjs` builds each zip in the test with a small stored-zip writer in `test/helpers.mjs` (no
committed file). The first copy holds `FIRST` and the second `SECOND`.
- **The main part, the same name twice:** refused, the message naming it, nothing read. Under `'first'`: reads `FIRST`,
  one warning with count 2, and an untouched save holds one entry, `FIRST`.
- **The case variant** in both orders: refused. Under `'first'`, the first copy's spelling named and its bytes read.
- **Three copies:** count 3.
- **`[Content_Types].xml` twice:** refused. Under `'first'`, the first's types hold (the second gives the main part a
  template's type, to prove it).
- **`_rels/.rels` twice, and `word/_rels/document.xml.rels` twice:** the same, the first's relationships, and an untouched
  save keeping the first's main part.
- **A media part twice, and an orphan twice:** refused, since the rule covers every entry. Under `'first'`, the media's
  first copy is saved and the orphan is not carried.
- **A relationship's target in another case than its entry:** loads, and reads the part.
- **Flat OPC, directory and memory stores:** one duplicate each, refused, and read first under the option.
- **`insertOoxml`:** a flat OPC fragment with two `pkg:part` of one name is refused.
- **The text model:** `package.duplicate_entry` in `TextResult.warnings`, its selector and `detail`; `duplicate_entries` by
  name in `TextOptions`, its values and refusal.
- **A clean package:** every committed fixture loads as before. The gate is the full suite, whose fixtures hold no
  duplicate.

The oracles are the WD's NOTE (Word: first copy after a repair) and docx4j-python's CR-032, whose cases these mirror. Word
has not been run on these exact files: the refusal is the default, and the `'first'` reading is the WD's statement of
Word's repair.

## 4. Size

About half a day: the stores' index (four classes, one shared helper), the check in `loadPackage`, the exception, the
option, the warning through `pkg.warnings` and the text model, and the tests.

## 5. Decisions (taken 2026-11-04, all as recommended, subject to check 75)

Jason, 2026-11-04: "CR-024 decisions as recommended, subject to: Word's reading of a non-ASCII case pair is unmeasured -
let's measure it". Check 75 (`test/README.md` item 75, `test/fixtures/check75/`, staged in the share's `__tmp/75/`) measures
it, and with it section 2.7's decoy-first point. Its answer settles section 2.1's key: Word's reading of `É` and `é` as
one name or two decides whether the key folds every letter, as proposed, or ASCII letters only, OPC's letter.

| # | Decision | Recommendation |
|---|---|---|
| D1 | Where the check sits | In each store's index, which keeps the first copy per key and counts the rest (`duplicates()`); `loadPackage` refuses before reading anything. The alternative, a scan of `partNames()` in `loadPackage`, would still leave the stores reading the last copy under `'first'` |
| D2 | Which stores | Every store this package ships: zip, flat OPC, directory and memory (its constructor), as docx4j-python's CR-032. The WD names the zip; the others can hold the same ambiguity (`pkg:part` twice; `Word/` beside `word/` on Linux) |
| D3 | Lookups | By the OPC key, case folded, as python's store and as OPC's equivalence: unique after the check, and a relationship whose target differs only in case finds its part. docx4j's store looks up exact names |
| D4 | The refusal | `DuplicateEntryException extends InvalidFormatException`, its `code` and `entries`, every duplicated name in one message, as `MisplacedContentException` carries its findings |
| D5 | The option's name | `LoadOptions.duplicateEntries: 'refuse' \| 'first'` (camelCase, as `mcePreprocess`); `duplicate_entries` accepted by `TextOptions` as the WD's name, documented as the load's |
| D6 | The warning | `DuplicateEntryWarning` (`code`, `entry`, `count`, python's fields), first in `pkg.warnings`; in the text model as `package.duplicate_entry` with `TextWarning.detail` |
| D7 | Release | With 0.7.0: it changes what loads, and 0.7.0 already carries this repair policy's other behaviour changes (CR-016's refusing save, CR-012 section 28). The alternative is the release after |

**Check 75, round 1** (Jason, 2026-11-04; `test/README.md` item 75). Every file reported unreadable content and was
repaired.
- **The ASCII pair (`75a`) is one name to Word, and Word reads the first entry.** It shows `headerA.xml`'s header though
  the relationship names `headera.xml`, and saves that one as `word/header1.xml`. That answers section 2.7's decoy-first
  point: `'first'` reads as Word does, by central-directory order. An exact-name reading, this package's today, shows
  another part than Word.
- **The non-ASCII files (`75b` to `75e`) answer nothing about the fold.** Word finds any zip entry named in raw UTF-8
  unreadable, a single one too, by either form of target, and its repair drops the part. OPC maps a part name to a zip
  item name as ASCII, a non-ASCII character percent-encoded, so for names in that form lower-casing folds ASCII letters
  only: [M1.12]'s letter, and the key as proposed.
- **Round 2** (Jason, 2026-11-04): `75f`, `75g` and `75h` each reported unreadable content and were repaired with no
  header, the control `75f` too (one entry named `word/header%C3%A9.xml`, as OPC maps it). Word 365 reads no part whose
  name holds a non-ASCII character, in either form, so a non-ASCII case pair is never read by Word, and the fold is not Word's
  to decide. **The key stands as proposed** (section 2.1): for a percent-encoded name, which is ASCII, lower-casing folds
  the ASCII letters only, [M1.12]'s letter; a raw UTF-8 pair, which it folds, is a package Word repairs anyway. 
- **Round 3** (Jason, 2026-11-04) clears the writer:
  - `75j`, one ASCII header through the check's writer, opened with no message, so `75a`'s prompt was the duplicate's;
  - `75i`, one file from that writer, kept its ASCII header and dropped its percent-encoded one;
  - `75k`, the name through fflate's writer, was dropped as well.

  Round 2's reading stands, and with it the key (`test/README.md` item 75).
- **Separately, Word repairs a part whose name holds a non-ASCII character by dropping it,** in raw UTF-8 or
  percent-encoded alike, where this package reads it (`75b` to `75d`, `75f` to `75h`). That is a tier 2 repair of its own,
  not REQ-206's, recorded for the portfolio (its repair policy, section 4.2); nothing changes in this CR.

## 6. Registry

`core-ts/CR-024`, proposed, no dependency on another repository; docx4j-python's CR-032 is its sibling, and the WD's
`package.multiple_main_parts`, if ruled, a later phase.

## 7. Sources

- The WD's REQ-206 and its NOTE (portfolio `docs/text_extraction_specification_wd.md`, 0.74 and 0.77; Jason, 4 November
  2026).
- docx4j-python `docs/change-requests/CR-032-a-package-holding-two-entries-of-one-name-is-refused.md` (`778b3bf`).
- The docx-tool-identifier's `extraction-probe-dupmain` (81 configurations) and `extraction-probe-dupcase`, through the
  portfolio.
- Measured here: section 1's seven zips, built by a scratch writer, not committed.
