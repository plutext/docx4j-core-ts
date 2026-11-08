# CR-031: The `repair` option: how a load treats a file Word reads only through a repair, or won't open (WD REQ-207 and REQ-208)

**Status:** Proposed 2026-11-08, at Jason's "draft the repair cr". Measured here the same day on `b2a3c04` (section 1). D9, the default, added 2026-11-09 with the engines' views (section 2.1.2), and D1, D2 and D6's recommendations revised the same day: `duplicateEntries` folded into `repair`, which leaves `checkLossOnSave` alone. D1 to D9 decided 2026-11-09 as recommended (Jason: "CR-031 decisions as now recommended. make the commits, then implement please"); being built.

## 0. Summary

WD 0.102 (portfolio `4c81f10`, REQ-207 and REQ-208, REQ-206's NOTE on `External`) and the repair policy's section 4.3 give every engine one load option, `repair`, with three values. `refuse`, the default, keeps every refusal. `word` follows Word's repair where the WD states one, and raises that repair's warning. `all` does that and also applies the engine's own repairs to shapes Word won't open, each one documented and warned. A text result says which happened in its `repaired` metadata: `none`, `word` or `own`. docx4j has built it as `RepairMode` (`880d7a3bb`), and python as `LoadOptions(repair=)` (`e05291b`).

Here today `duplicateEntries: 'first'` stands in for `word` on two of its three repairs, and nothing stands for `all`. A package holding an `External` `officeDocument` relationship beside its real one reads silently, where Word refuses it, so core-ts is now the one engine that reads it. This CR adds the option, folding `duplicateEntries` into it while that is still unreleased, the `External` rule at each value, the reading of a relationships part in a wrong namespace that CR-025 section 8 left to it, and the `repaired` metadata.

## 1. Here today (measured 2026-11-08, `b2a3c04`)

| Shape | Word | Here |
|---|---|---|
| Two entries of one name (REQ-206) | Repaired, the first read | Refused; `duplicateEntries: 'first'` reads the first, `package.duplicate_entry` (CR-024) |
| More than one `officeDocument` relationship | Repaired, a fresh Word reading the last | Refused; `duplicateEntries: 'first'` reads the last, `package.multiple_main_parts` (CR-024 section 9) |
| A relationships part in a wrong namespace (`rn-*`) | Repaired, every `Relationship` read by local name | A wrong root refused; wrong children load and a save loses parts. CR-025 section 8 proposes the refusal and leaves the reading to this CR (its D11) |
| `External` beside an internal relationship, first or last (`x-ext-*`, 8 jobs, the sources and python's saves) | Repair offered, then refused, in about 5 seconds, with no fetch | All 8 read silently by both settings of `duplicateEntries`, and a save keeps the `External` relationship, so it stays a file Word refuses (`9459d1e`, CR-024 section 9.2) |
| `External` alone | Not measured; REQ-321's case | Loads as a plain `OpcPackage` with no main part, and `Texts.of` refuses it as not a WordprocessingML package, without naming REQ-321 |

`pkg.checkLossOnSave` is this package's own way past CR-025's gate. CR-025 section 2.4 records that if this package builds the `repair` switch, turning it on turns that check off too, as docx4j's gate gives way under `WORD`. This CR no longer follows that record: `repair` leaves the check alone (D6).

## 2. Proposal

### 2.1 The option

`LoadOptions.repair?: 'refuse' | 'word' | 'all'`, `'refuse'` by default, acting at the load and so holding for every operation on the package (REQ-207). `TextOptions` accepts `repair` by the WD's name (REQ-1001) and passes it to the load, as it does `duplicate_entries`. Any other value is refused.

- **`word`** reads CR-024's two shapes as `duplicateEntries: 'first'` reads them today, the first entry of each name and the last `officeDocument` relationship's target, with their warnings, and reads a relationships part in a wrong namespace as Word's repair does (2.3).
- **`all`** is `word`, plus this package's own repairs (2.2). Today that is one.
- **`duplicateEntries` goes** (D1), from `LoadOptions` and, as `duplicate_entries`, from `TextOptions`, where a call naming it is refused with a message naming `repair: 'word'`. It is unreleased: CR-024 came after 0.6.0. The two exceptions' messages name `{ repair: 'word' }`. The warnings stay as they are: `pkg.duplicateEntries`, `pkg.multipleMainParts`, `package.duplicate_entry` and `package.multiple_main_parts`.

### 2.1.1 What `refuse` means (added 2026-11-09)

Jason asked whether `refuse` means a load fails on whatever needs a repair, or opens without any, and had the answer recorded here ("include all that as context in CR-031"). The answer is neither. REQ-207's words are "every refusal in this specification stands", so a load fails where the WD, or this package, refuses a shape, and nowhere else. A repair, in REQ-207's sense, is only what `word` and `all` switch on, so the shapes Word repairs fall into two groups.

**Refused at `refuse`.** The only reading these have is Word's repair, or Word won't open them, so the load fails:

- Two entries of one name (REQ-206): `DuplicateEntryException`, CR-024.
- More than one `officeDocument` relationship: `MultipleMainPartsException`, CR-024 section 9.
- A relationships part in a wrong namespace: `RelationshipsNamespaceException`, CR-025 section 8, built at `2fac740` after section 1 was measured. `word` reads it (2.3).
- `External` beside an internal relationship. It reads today, and D3 would refuse it.
- The roots REQ-318 names (CR-028 D11). Word won't open them, so they stay refused at `word` too (REQ-208), and at `all` until a repair is written for one (2.5).

Today `duplicateEntries: 'refuse'`, the default, plays this role for the first two, and `'first'` gives `word`'s reading of them. D1 folds that option into `repair`.

**Read at every value, `refuse` among them.** For each of these the WD states one reading, or none, and REQ-207's list for `word` doesn't name them, so `repair` doesn't change how they read, even where Word itself repairs the file:

- An undeclared prefix in `mc:Ignorable` (REQ-323, which says "at every value of `repair`" outright; CR-030). Word repairs such a file, one prompt an open, and loses nothing. Here it reads with Word's text and `mce.undeclared_prefix`.
- An element that selects nothing (REQ-305 and REQ-306, CR-029), and an element the model has no place for (REQ-308 to REQ-316, CR-026).
- The roots REQ-319 reads, and those REQ-320 reads as an empty document (CR-028 D11).
- A `w:br` outside a `w:r` (CR-021, docx4j CR-073's rules). The WD states nothing on it.
- A misplaced child (CR-016), which the WD doesn't name either. Check 72 measured 18 such files: Word won't open 11, and repairs the other 7 without a prompt. Two of the 7 are bare breaks, read here as the item above says. The other 16 are read with `target.misplaced`, the 11 under REQ-208's leave to read a shape the WD doesn't name, so that a caller can fix it. A save that writes such a part anew is refused, except for `72m`'s string in a run, which the read drops as Word's repair does.

So `refuse` doesn't fail a load on everything Word would repair. Nor does it open everything unrepaired: what can't be read without Word's repair, it refuses.

### 2.1.2 The default: the engines' views (2026-11-09, D9)

Jason: "CR-031 says 'refuse' is the default value for the Load Option. I wonder whether 'word' would be the better default? ask the other engines and portfolio for their opinion." The question went to docx4j, python and the portfolio the same day, with what a `word` default would change and without this package's view, so as not to anchor theirs. Each answered, and Jason had the summary recorded here ("add this summary to the CR").

**What a `word` default would change.** At the load, three shapes now refused would read, each with its warning and `repaired: word`: two entries of one name, the first read; more than one `officeDocument` relationship, the last read; and a relationships part in a wrong namespace, read by local name. Word opens each only after its repair prompt, so a `word` default amounts to answering Yes to that prompt on the caller's behalf. `External` beside an internal relationship, the REQ-318 roots and REQ-321 stay refused at `word` (REQ-206's NOTE, REQ-208). REQ-207's table names `refuse` as the default, so a `word` default is a WD change as well as each engine's.

**This package before 0.7.0.** All three refusals are new in 0.7.0 here, as in python. The last release, 0.6.0, read a duplicated entry's last copy and said nothing, and an untouched save of a doubled `_rels/.rels` lost a part (CR-024 section 1). It loaded a relationships part whose children were in a wrong namespace, and a save lost the parts they named (CR-025 section 8).

**The views.**

| | Default | Condition or reason |
|---|---|---|
| Portfolio | Keep `refuse`, and look again when its D10 count of real producers' files arrives | Loosening later breaks nobody, and tightening would break three engines. All three shapes hold two documents in one package, which readers disagree on (the identifier's `extraction-probe-dupmain` split 81 configurations four ways), and a warning a caller doesn't read leaves them unaware. Jason's REQ-206 ruling refuses a duplicate in every operation for this reason. A refusal names the producer's defect. If `word` all the same, the save gate must not follow `repair`. A possible middle course, if the count shows the shapes are common: `word` by default for operations that only read, `refuse` for anything that writes |
| Python | `word` | Its 0.2.0 read all three shapes silently and wrongly, so `refuse` would turn files that loaded into exceptions, where `word` reads them as Word does, with a warning, a log line and `repaired: word`. Its saves write Word's repaired form, which Word opened silently in its merge round (its CR-039 section 9.7). Conditions: `strict=True` keeps the refusals unless `repair` is given; `all` stays opt-in; its reading of a wrong-namespace relationships part is unbuilt (its CR-033 phase C), so its `word` isn't complete yet. About a quarter day there after the WD changes |
| Docx4j | Keep `refuse` while one value governs both the load and the save. It leans to its option (b): `word` at the load only, with the save refusing by default to make a repair permanent | 17.3.6 ships `RepairMode.NONE`, which is `refuse`. Its gate gives way under `WORD` however it was reached, so a `word` default would write a repaired package on a load and save with no consent, and its gate doesn't yet cover the duplicate-entry and two-main-part losses. (b) needs the WD to tell the load's default from the save's consent, and a small docx4j CR |

Docx4j measured its 31 `package-wiring` files the same day. 17.3.6 loads 8 at `refuse` and 29 at `word`, the two `External` files still refused. 17.3.5, before the refusals, loaded 28, but 5 of them read content other than Word's (`dup-entry-document`, `dup-central-reversed`, `case-variant-exact-last`, and both `two-rels-entries` files), two more failed with exceptions, and it loaded `x-ext-after`, which Word refuses, with nothing said. So the old leniency wasn't a correct reading of these files, which is python's argument too.

**Where they agree and differ.**

- **The save:** all three treat it as the real question. Docx4j and the portfolio want a separate yes before a save makes a repair permanent. Python is content for the save to write Word's repaired form by default.
- **The load:** python, and docx4j's (b), lean to `word`. The portfolio would wait for the count.
- **Python's CR-033 question**, which `repair` value opens the save gate for content a lenient load skipped: docx4j now recommends leaving tier-1 drops ungated (python's option C), and python leans C. This package gave its view as B, so it's the outlier now. With D6 revised the same day, `repair` opens this package's gate at no value: a save's consent is `checkLossOnSave` alone, the separation the portfolio and docx4j ask for below. Which items that gate passes by default is then CR-025's question, not this CR's.

**This package's view.** Keep `refuse` in 0.7.0, and decide on the portfolio's count. That matches the WD and docx4j 17.3.6, and moving to `word` later breaks nobody. If Jason then wants `word`, docx4j's (b) fits this package best: the save's consent is already a switch of its own here (`checkLossOnSave`, CR-025 D8). One gap: under `duplicateEntries: 'first'` today, and under `word` once D1 folds it in, a save writes the repaired package with no gate at all, so a `word` default would need that covered, as docx4j says of its own gate.

### 2.2 `External` beside an internal relationship (REQ-206's NOTE)

- **At `refuse` and `word`**, refused at the load whether the `External` relationship comes first or last: `ExternalMainPartException`, an `InvalidFormatException` with `code` `package.external_main_part`, `url` as written and `targets`, python's fields. The message names the URL and `repair: 'all'`. This undoes `9459d1e`'s silent reading (D3).
- **At `all`**, the internal target is read and every `External` `officeDocument` relationship dropped at the load, so a save writes none. The URL is never fetched, which nothing here does anyway. `ExternalMainPartWarning` (`code`, `partName` the part read, `url` the first URL as written, `message`) goes on `pkg.externalMainPart` and on `pkg.warnings`, after CR-024's warnings. In the text model it is `TextWarning.PACKAGE_EXTERNAL_MAIN_PART`, with the part as the selector and the URL as the `detail`. Where there are also two internal relationships, Word's last-relationship repair comes first, as docx4j has it.
- **`External` alone** is refused in every mode as a package with no main part (REQ-321), naming the URL, where today it loads as a bare `OpcPackage` (D4).

### 2.3 A relationships part in a wrong namespace, under `word`

This is CR-025 section 8's reading, which its D11 left here. Every `Relationship` is read by its local name, the walk goes on, and `package.relationships_namespace` (`partName`, `count`) goes on `pkg.warnings` and in the text model, as docx4j and the WD have it. A save writes the right namespace. CR-025 section 8 builds the refusal at `refuse`, and the two go together (D5).

### 2.4 What is said

- **`repaired` metadata** on every `TextResult` (REQ-207, REQ-801): `own` where a repair of this package's own was made, `word` where only Word's were followed, else `none`. It is read from the package's warnings, so it reports what the load did, whatever the option. `pkg.repaired` gives the same value.
- **The saves.** `repair` leaves `pkg.checkLossOnSave` alone at every value (D6), so a save that would lose what a read skipped still needs its own yes. A save of a package `word` read writes the repaired form, one entry of each name and one main relationship, as `duplicateEntries: 'first'` does today.
- **The documentation** REQ-207 requires: a "Repairs" section in `README.md`, listing what `word` follows and what `all` adds, each with its warning.

### 2.5 Not here

- Repairs of this package's own beyond `External`. They are added one case at a time, as Jason's words in the WD's NOTE have it.
- A refusal elsewhere in the WD that `all` might loosen (REQ-208): REQ-318's roots and REQ-321 stay refused at every value until a repair is written for one.
- CR-016's misplaced content and CR-026's unknown elements. They read the same at every value, since Word's reading of them doesn't depend on a repair switch.

## 3. Tests and oracles

- `x-ext-*`: the 8 sources and Word's results, copied unedited with a SHA-1 manifest, each refused at `refuse` and `word`, read at `all` with the warning, and a save under `all` holding no `External` relationship.
- `test/package-wiring/`'s and `check75/`'s files under `word` and `all`, each reading as under today's `duplicateEntries: 'first'`, measured before the option goes, and the `rn-*` files once CR-025 section 8 has them.
- The option on `Texts.of` and in the suite's adapter; `repaired` on each of the three outcomes; `duplicate_entries` refused on `Texts.of`, naming `repair: 'word'`; `checkLossOnSave` on under each value.
- A Word check of the saves `all` writes for `x-ext-*`, as docx4j's and python's were checked (silent), on Jason's word (D7).

## 4. Phases and size

| Phase | What | Size |
|---|---|---|
| 1 | The option, `word` over CR-024's two readings, `duplicateEntries` removed (D1), `repaired`, the text model's option | ½ day |
| 2 | `External`: the refusal, `all`'s repair and warning, REQ-321's refusal | ½ day |
| 3 | 2.3 with CR-025 section 8; the README's Repairs section; CLAUDE.md, the index | ¼ day |

No dependency on another repository.

## 5. Decisions

| | Question | Recommendation |
|---|---|---|
| D1 | Keep `duplicateEntries` beside `repair`? | **Decided** 2026-11-09 as recommended (Jason: "CR-031 decisions as now recommended"). **Recommendation revised** 2026-11-09, at Jason's question ("what is the argument for Keeping duplicateEntries beside repair? if we can simplify things without losing anything meaningful, why not do it now, while we can?") and his "yes" to revising it. Drop it, from `LoadOptions` and `TextOptions`, in the release that brings `repair`. It is unreleased, so dropping it now breaks nobody, where after 0.7.0 it would. The WD only says an implementation MAY offer `duplicate_entries`, and docx4j 17.3.6 has `RepairMode` alone. All it gives beyond `word` is reading a duplicated entry while refusing a relationships part in a wrong namespace, two packaging defects behind one Word prompt, and its `'first'` reads the last of two main relationships. It needs phase 1 in the same release, or nothing reads such files (D8), and D6 as revised. Was: keep it, as the WD's own option for REQ-206, as python does |
| D2 | `repair: 'word'` with an explicit `duplicateEntries: 'refuse'`? | **Decided** 2026-11-09 as recommended (Jason: "CR-031 decisions as now recommended"). **Moot** under D1 as revised (2026-11-09): there is no `duplicateEntries` to contradict it. If D1 keeps the option: refuse the call as contradicting itself, as CR-012 section 29 refuses `include` against `comments`, since silently letting one win hides a mistake |
| D3 | Refuse `External` beside an internal relationship at `refuse` and `word`? It reads today. | **Decided** 2026-11-09 as recommended (Jason: "CR-031 decisions as now recommended"). Yes, as the WD now states, and as docx4j and python do. Word refuses such a file even through its repair, so `word` has no reading to give. It's a `Changed` line: a package that loads today stops loading |
| D4 | Refuse `External` alone at the load (REQ-321)? | **Decided** 2026-11-09 as recommended (Jason: "CR-031 decisions as now recommended"). Yes, naming the URL, as docx4j does. Today it loads as a package with no main part, which every reader then refuses less clearly |
| D5 | Build 2.3 here or in CR-025 section 8? | **Decided** 2026-11-09 as recommended (Jason: "CR-031 decisions as now recommended"). Here, with section 8's refusal in the same build, so the shape is refused and readable from one release. Section 8's D11 already points here |
| D6 | Does `word` or `all` turn `checkLossOnSave` off? | **Decided** 2026-11-09 as recommended (Jason: "CR-031 decisions as now recommended"). **Recommendation revised** 2026-11-09, with D1: no. `repair` decides what a load reads, and a save that would lose what the read skipped still needs its own yes, `checkLossOnSave`, as CR-025 D8 made it. The engines' answers on the default agree that a save's consent shouldn't follow the load option (section 2.1.2: the portfolio's point 4, docx4j's (b)). With D1, `word` also carries the duplicate-entry reading, so a caller who wanted only that would otherwise lose the gate too. This departs from CR-025 section 2.4's record. Was: yes, by starting it false, as docx4j's gate gives way under `WORD` |
| D7 | A Word check of the saves `all` writes? | **Decided** 2026-11-09 as recommended (Jason: "CR-031 decisions as now recommended"). Yes, eight files, one round, as docx4j's and python's had. It's the only evidence that a repaired save opens |
| D8 | Which release? | **Decided** 2026-11-09 as recommended (Jason: "CR-031 decisions as now recommended"). 0.7.0, so the three engines agree on `External` when it ships. Otherwise core-ts 0.7.0 reads a file the other two refuse |
| D9 | The default: `refuse`, as WD 0.103 has it, or `word` (Jason's question, 2026-11-09; section 2.1.2)? | **Decided** 2026-11-09 as recommended (Jason: "CR-031 decisions as now recommended"). `refuse` in 0.7.0, and decide on the portfolio's count of real producers' files. If `word` then, at the load only, with a save that makes a repair permanent needing its own yes, as docx4j's (b), and a gate added for the save of `word`'s two readings (D1) |

## 6. Registry

`core-ts/CR-031`, `depends_on: [core-ts/CR-025]` for section 8's refusal. Sent once this is committed.

## 7. Sources

- `docx4j-portfolio/docs/text_extraction_specification_wd.md`: WD 0.102 at `4c81f10` (REQ-207, REQ-208, REQ-206's NOTE), WD 0.103 at `eceabd6`; REQ-801's metadata; REQ-1001's table.
- `docx4j-portfolio/docs/repair_policy_across_ports.md` section 4.3, and Jason's words on `all` quoted there.
- docx4j `880d7a3bb` (`RepairMode.ALL`, `followsWord()`, `makesOwnRepairs()`, `TextResult.REPAIRED`), `6025529d9`; its CR-075.
- docx4j-python `e05291b`, `30eab21`: `LoadOptions(repair=)`, `ExternalMainPartException`; its CR-032.
- The share's `bridge/portfolio/x-ext-*`.
- CR-024 sections 9 and 9.2, CR-025 sections 2.4 and 8.
