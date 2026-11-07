# CR-024: A package holding two entries of one name is refused (WD REQ-206), with `duplicateEntries: 'first'`

**Status:** Proposed 2026-11-04, at Jason's "Duplicate zip entries (REQ-206): go", on the WD's REQ-206 (0.74, Jason's
ruling "refuse the package, with an optional config option to read the first copy and warn", made in this session; its five
details ruled as recommended in 0.77, portfolio `b0d333d`, from this session's suggestions). Measured here the same day
(section 1). D1 to D7 decided 2026-11-04, all as recommended, subject to Word check 75 on a non-ASCII case pair, which left
the key as proposed (section 5). Implemented 2026-11-04 (section 8), unreleased (D7: with 0.7.0). docx4j-python's sibling is its CR-032 (proposed at
`778b3bf`, done at `3a884a5`), whose names this CR follows where the two packages' shapes allow; docx4j built it as its CR-075 (phase 1 at `27d94d04b`). Phase 2, more than one officeDocument relationship, WD 0.100's REQ-206 statement, implemented 2026-11-07 (Jason: "build CR-024 phase 2"; section 9), unreleased, with 0.7.0 as D7 has it.

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
  wants one, is a neighbouring case, left to a phase 2. Here `getMainDocumentPart()` and the readers took the last
  relationship's target, and `getMainPart()` and the package-type check the first: a defect either way (measured
  2026-11-04, recorded under REQ-206's NOTE). The NOTE then had Word, through a repair, read the target whose `main+xml`
  Override is last. **That reading was a coincidence of opening order** (the `open-order` round, WD 0.100): Word's choice
  alternates with each repaired open in a session, and a fresh Word reads the last relationship's target. Jason ruled the
  statement on 2026-11-07, and section 9 builds it.
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
- **Changed (phase 2):** a package with more than one officeDocument relationship is refused with
  `MultipleMainPartsException` naming the targets. Previously it loaded silently, `getMainDocumentPart()` and `getMainPart()`
  answering with different parts, and a save kept both.
- **Added (phase 2):** under `duplicateEntries: 'first'`, such a package reads the last relationship's target and a save
  writes one main part; `MultipleMainPartsWarning` (`package.multiple_main_parts`) on `pkg.warnings`, as
  `pkg.multipleMainParts`, and in `TextResult.warnings`.

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

`core-ts/CR-024`, no dependency on another repository; docx4j-python's CR-032 is its sibling, and docx4j's CR-075. The WD's
`package.multiple_main_parts`, ruled 2026-11-07, is phase 2 (section 9).

## 7. Sources

- The WD's REQ-206 and its NOTE (portfolio `docs/text_extraction_specification_wd.md`, 0.74 and 0.77; Jason, 4 November
  2026).
- docx4j-python `docs/change-requests/CR-032-a-package-holding-two-entries-of-one-name-is-refused.md` (`778b3bf`).
- The docx-tool-identifier's `extraction-probe-dupmain` (81 configurations) and `extraction-probe-dupcase`, through the
  portfolio.
- Measured here: section 1's seven zips, built by a scratch writer, not committed.
- Phase 2: the WD's REQ-206 statement and NOTE (0.100, portfolio `9034ad0`); docx4j `54d037c63` (`Load3.oneMainPart`,
  `PackageWiringTest`) and its `package-wiring` family at `27d94d04b`; docx4j-python `d734634` (its CR-032 section 11).

## 8. Implemented (2026-11-04)

Jason, 2026-11-04: "commit then Build CR-024". Built as section 2 proposes, with D1 to D7 as decided and the key as check 75
left it.

- **The index** (`PartStore.mts`): `EntryIndex<T>` keys each entry by `entryKey` (the name without a leading `/`,
  lower-cased, as `PartName`'s equality compares). It keeps the first entry of a key, with its spelling and value, and counts
  the rest. `PartStore` gains the optional `duplicates()`.
- **Every store** reads its container into one:
  - `ZipPartStore`: the central directory's order, its inflate cache keyed by the key;
  - `FlatOpcPartStore`: the `pkg:part` order;
  - `DirectoryPartStore`: the walk's order;
  - `MemoryPartStore`: the constructor's entries; `put` replaces, a sink's write.

  `has`, `load` and `size` look up without case, and `partNames()` lists each key's first spelling once.
- **The refusal** (`Load.mts`, `duplicateEntriesOf`): `loadPackage` asks the store before it reads `[Content_Types].xml`.
  It throws `DuplicateEntryException` (`InvalidFormatException`'s subclass, `code` and `entries`) unless
  `LoadOptions.duplicateEntries` is `'first'`, and refuses any other value of the option, naming REQ-206.
- **The warning.** Under `'first'`, `pkg.duplicateEntries` holds one `DuplicateEntryWarning` (`code`, `entry`, `count`,
  `message`) per name. `pkg.warnings` lists them first, its type widened to `PackageWarning`. `Texts.of(pkg)` carries each
  into `TextResult.warnings` as `package.duplicate_entry`, the selector the entry, with the new `TextWarning.detail` holding
  the count; `toJSON` writes `detail` only where set.
- **`TextOptions`** accepts `duplicate_entries` (`refuse`, `first`) between `warnings` and `profile`, the WD's order. The
  suite's adapter passes it to the load, a package per fixture and setting.
- **The save** is unchanged. Under `'first'` an untouched part's bytes are the first copy's, and an orphan is not
  carried. `insertOoxml`'s flat OPC route and an embedded package refuse through their own loads, with no code of their own.
- **The behaviour change for every package:** a relationship whose target differs from its entry only in case now finds
  its part, where the load refused "cannot find part". No committed document holds one; the full suite is unchanged.

**Tests** (`test/duplicate-entries.test.mjs`, nine; `storedZip` moved to `test/helpers.mjs`, check 75's builder now importing
it, its files unchanged):
- the main part twice: refused with the message, the code and the entries, the store's `partNames()` and `duplicates()`;
  under `'first'`, the first read, the warning, the save and its reload;
- three copies;
- the case variant in both orders, the decoy first read as Word reads it (check 75a);
- `[Content_Types].xml`, `_rels/.rels` and `word/_rels/document.xml.rels` twice;
- a media part and an orphan twice, refused together, the media's first copy saved and the orphan dropped;
- a target in another case;
- flat OPC, `insertOoxml`, memory and a directory (on a case-sensitive file system);
- the text model's warning, its `detail` and the option by name; the load's option refusing another value;
- check 75: Word 365's saves (`test/check75/`, eleven saves and Jason's `results.txt`, a SHA-1 manifest). `75a` is refused
  here, and read under `'first'` gives Word's header. `75j` loads clean. The non-ASCII files hold a pair or not, as the key
  says. Word's saves of the non-ASCII files hold no header.

`misplaced-content.test.mjs`'s gate leaves out check 75's built files, as it leaves out check 72's. npm run typecheck passes,
and npm test passes 1,795 of 1,795.

## 9. Phase 2: more than one officeDocument relationship (implemented 2026-11-07)

Jason, 2026-11-07: "build CR-024 phase 2". The statement is REQ-206's, ruled the same day in the docx4j session ("yes") and
written into WD 0.100 (portfolio `9034ad0`): a package whose package relationships hold more than one officeDocument
relationship is refused, naming the targets, as a duplicated entry is and for the same reason. Under the same switch an
implementation reads the last relationship's target, as a fresh Word does, and raises `package.multiple_main_parts` once,
the selector that part's name and `detail` the number of relationships. An unreferenced part is no case: an orphan, never
read, dropped on save.

### 9.1 What landed

- **The refusal.** `MultipleMainPartsException`, an `InvalidFormatException`, with `code` and `targets`, python's. `targets`
  gives each as a part name in the relationships' order.
- **The switch.** `oneMainPart` in `Load.mts` runs once the package relationships are read, before the package's class is
  chosen. Under `'first'` the other relationships go there, so the class, the walk, `getMainPart()` and
  `getMainDocumentPart()` all find one main part. Section 2.7's split needed nothing else.
- **The record.** `MultipleMainPartsWarning` (`code`, `partName`, `count`, `message`, python's fields) on
  `pkg.multipleMainParts`, and on `pkg.warnings` after any `DuplicateEntryWarning`, `PackageWarning` widened again.
- **The text model.** `TextWarning.PACKAGE_MULTIPLE_MAIN_PARTS`, after the entries' warnings: no address, the part read the
  selector, the relationships `detail`.
- **The save** is unchanged. A relationships part is always marshalled, so it writes the one relationship, and the other
  part was never loaded, so it isn't written. `clone()` goes through a save, so the copy holds one main part and says nothing.

### 9.2 What the build settled beyond the statement

- **The switch's name reads oddly.** `duplicateEntries: 'first'` reads the **last** relationship. It's the WD's one switch and
  its value, as in python, so the name stays, and the option's doc and the refusal's message say which.
- **What counts.** A transitional and a strict relationship are two, and two naming one target are two, as python's and
  docx4j's count (`e2d51d1e2`, which excludes an External one too). Word's `one-target` round fits that: two and three relationships to one part, and two spelling
  it differently, are each repaired, where the control opens silently (`ot-00` to `ot-03`, twice each). `graphicFrameDoc` and
  the diagram layout type, which `MAIN_PART_RELATIONSHIPS` holds, don't count, as they aren't officeDocument relationships.
  That's docx4j's `isMainPart`. Python counts all four types, so the two differ on a package holding an officeDocument
  relationship beside one of those, and no file measures that.
- **An External relationship isn't counted** (Jason, 2026-11-07: "stop counting External relationships"). It names no part. As
  first committed at `4003051` it counted, which python's check at `9077b95` showed was wrong: with one last, the switch kept
  the URL and dropped the only main part, and the default refused a package that had loaded. Now such a relationship is left
  alone: not named, not dropped, and a save keeps it. `loadPackage` and `getMainPart()` pass over it to a part, which also
  mends a package with one first, read as having no main part before this phase. Only External relationships read as they did,
  with no main part and nothing refused.
- **An Override goes after the walk, and only for a part not loaded.** A dropped target that a second relationship names, or
  the kept part reaches by a relationship of its own, stays and keeps its type. Python removes the Override before its walk,
  so it agrees on the first of those and not the second.
- **Both faults at once.** The entries are refused first, nothing having been read. Under the switch each is recorded, the
  entries first.
- **Every package type.** The rule is the load's, so it refuses a pptx or an xlsx too, as docx4j's does. Word alone was measured.
- **`insertOoxml`.** A flat OPC fragment is refused through its own load, and the message names an option its caller can't
  pass, as the duplicate entry's already does.

### 9.3 Left as it stands

- **A part only the dropped main part reached.** It isn't loaded or saved, but its Override stays in `[Content_Types].xml`, as
  an orphan's does here today (python's section 11.2 has the same for an orphan). Word's reading of such a save isn't
  measured.
- **A last target the package doesn't hold.** Under the switch the load refuses it, `cannot find part`, as it refuses one
  relationship naming an absent part, even where an earlier relationship's part is there. An absent earlier target goes with
  its relationship, and the last is read. No file section 9.4 rests on has either shape, so Word's reading of them isn't measured.
- **A package holding only an External relationship, its URL holding a `?`.** `PartName.resolve` refuses it, with a message about
  a segment's characters, where a plain URL loads with no main part. It did before this phase too.

### 9.4 Tests and oracles

`test/duplicate-entries.test.mjs`, eight tests after section 3's, zips built in the test as before:
- the refusal, its message and order, a strict and a transitional pair, two to one target, an xlsx;
- under the switch, the order deciding over entries and Overrides, three relationships, the record;
- an External relationship beside one part in either order, beside two, and alone;
- every route agreeing, the class coming from the part read, the other part not loaded;
- the save, what only the other part reached, the reload saying nothing, `clone()`, and the two cases where an Override stays;
- the text model once, both faults at once, one relationship saying nothing, `graphicFrameDoc` not counted;
- flat OPC and `insertOoxml`;
- docx4j's `package-wiring` probes.

**The oracle** is Word's saves in `test/package-wiring/` (seven probes and Word's save of each, copied unedited from docx4j
`27d94d04b`, with a SHA-1 manifest). The five `two-main-rels` saves are each a first open in a Word session. Refused here by
default, each reads under the switch as Word's save does, whatever the zip order, the Overrides or the names. The two
`one-main-rel` files read the related part with nothing said, by either option. A later repaired open in the same session
alternates, so those saves are what a fresh Word reads, which is the WD's statement.

**The corpus.** Of 1,880 packages, this repository's `test/` tree and docx4j's test resources, the only ones holding two
officeDocument relationships are docx4j's five probes, so none of the others that load today is refused.

Twelve deliberate breaks of the build were each caught by these tests: the first relationship kept, the Override left, the
Override always dropped, strict not counted, all four types counted, the relationships kept, the record missing from
`pkg.warnings`, the text model silent, an External relationship counted, the load or `getMainPart()` taking an External one
over a part, and one alone refused.

npm run typecheck passes, and npm test passes 1,907 of 1,907 on Node 22.
