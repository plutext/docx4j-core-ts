# CR-031: The `repair` option: how a load treats a file Word reads only through a repair, or won't open (WD REQ-207 and REQ-208)

**Status:** Proposed 2026-11-08, at Jason's "draft the repair cr". Measured here the same day on `b2a3c04` (section 1). Nothing built. D1 to D8 for Jason.

## 0. Summary

WD 0.102 (portfolio `4c81f10`, REQ-207 and REQ-208, REQ-206's NOTE on `External`) and the repair policy's section 4.3 give every engine one load option, `repair`, with three values. `refuse`, the default, keeps every refusal. `word` follows Word's repair where the WD states one, and raises that repair's warning. `all` does that and also applies the engine's own repairs to shapes Word won't open, each one documented and warned. A text result says which happened in its `repaired` metadata: `none`, `word` or `own`. docx4j has built it as `RepairMode` (`880d7a3bb`), and python as `LoadOptions(repair=)` (`e05291b`).

Here today `duplicateEntries: 'first'` stands in for `word` on two of its three repairs, and nothing stands for `all`. A package holding an `External` `officeDocument` relationship beside its real one reads silently, where Word refuses it, so core-ts is now the one engine that reads it. This CR adds the option, the `External` rule at each value, the reading of a relationships part in a wrong namespace that CR-025 section 8 left to it, and the `repaired` metadata.

## 1. Here today (measured 2026-11-08, `b2a3c04`)

| Shape | Word | Here |
|---|---|---|
| Two entries of one name (REQ-206) | Repaired, the first read | Refused; `duplicateEntries: 'first'` reads the first, `package.duplicate_entry` (CR-024) |
| More than one `officeDocument` relationship | Repaired, a fresh Word reading the last | Refused; `duplicateEntries: 'first'` reads the last, `package.multiple_main_parts` (CR-024 section 9) |
| A relationships part in a wrong namespace (`rn-*`) | Repaired, every `Relationship` read by local name | A wrong root refused; wrong children load and a save loses parts. CR-025 section 8 proposes the refusal and leaves the reading to this CR (its D11) |
| `External` beside an internal relationship, first or last (`x-ext-*`, 8 jobs, the sources and python's saves) | Repair offered, then refused, in about 5 seconds, with no fetch | All 8 read silently by both settings of `duplicateEntries`, and a save keeps the `External` relationship, so it stays a file Word refuses (`9459d1e`, CR-024 section 9.2) |
| `External` alone | Not measured; REQ-321's case | Loads as a plain `OpcPackage` with no main part, and `Texts.of` refuses it as not a WordprocessingML package, without naming REQ-321 |

`pkg.checkLossOnSave` is this package's own way past CR-025's gate. CR-025 section 2.4 records that if this package builds the `repair` switch, turning it on turns that check off too, as docx4j's gate gives way under `WORD`.

## 2. Proposal

### 2.1 The option

`LoadOptions.repair?: 'refuse' | 'word' | 'all'`, `'refuse'` by default, acting at the load and so holding for every operation on the package (REQ-207). `TextOptions` accepts `repair` by the WD's name (REQ-1001) and passes it to the load, as it does `duplicate_entries`. Any other value is refused.

- **`word`** implies `duplicateEntries: 'first'`, with its two readings as CR-024 has them, and reads a relationships part in a wrong namespace as Word's repair does (2.3). An explicit `duplicateEntries: 'refuse'` beside it contradicts it and is refused (D2).
- **`all`** is `word`, plus this package's own repairs (2.2). Today that is one.
- **`duplicateEntries` stays** as the WD's own option for REQ-206 (D1).

### 2.2 `External` beside an internal relationship (REQ-206's NOTE)

- **At `refuse` and `word`**, refused at the load whether the `External` relationship comes first or last: `ExternalMainPartException`, an `InvalidFormatException` with `code` `package.external_main_part`, `url` as written and `targets`, python's fields. The message names the URL and `repair: 'all'`. This undoes `9459d1e`'s silent reading (D3).
- **At `all`**, the internal target is read and every `External` `officeDocument` relationship dropped at the load, so a save writes none. The URL is never fetched, which nothing here does anyway. `ExternalMainPartWarning` (`code`, `partName` the part read, `url` the first URL as written, `message`) goes on `pkg.externalMainPart` and on `pkg.warnings`, after CR-024's warnings. In the text model it is `TextWarning.PACKAGE_EXTERNAL_MAIN_PART`, with the part as the selector and the URL as the `detail`. Where there are also two internal relationships, Word's last-relationship repair comes first, as docx4j has it.
- **`External` alone** is refused in every mode as a package with no main part (REQ-321), naming the URL, where today it loads as a bare `OpcPackage` (D4).

### 2.3 A relationships part in a wrong namespace, under `word`

This is CR-025 section 8's reading, which its D11 left here. Every `Relationship` is read by its local name, the walk goes on, and `package.relationships_namespace` (`partName`, `count`) goes on `pkg.warnings` and in the text model, as docx4j and the WD have it. A save writes the right namespace. CR-025 section 8 builds the refusal at `refuse`, and the two go together (D5).

### 2.4 What is said

- **`repaired` metadata** on every `TextResult` (REQ-207, REQ-801): `own` where a repair of this package's own was made, `word` where only Word's were followed, else `none`. It is read from the package's warnings, so it reports what the load did, whatever the option. `pkg.repaired` gives the same value.
- **The saves.** Under `word` and `all`, `pkg.checkLossOnSave` starts false, as CR-025 section 2.4 has it (D6). A caller can still set it.
- **The documentation** REQ-207 requires: a "Repairs" section in `README.md`, listing what `word` follows and what `all` adds, each with its warning.

### 2.5 Not here

- Repairs of this package's own beyond `External`. They are added one case at a time, as Jason's words in the WD's NOTE have it.
- A refusal elsewhere in the WD that `all` might loosen (REQ-208): REQ-318's roots and REQ-321 stay refused at every value until a repair is written for one.
- CR-016's misplaced content and CR-026's unknown elements. They read the same at every value, since Word's reading of them doesn't depend on a repair switch.

## 3. Tests and oracles

- `x-ext-*`: the 8 sources and Word's results, copied unedited with a SHA-1 manifest, each refused at `refuse` and `word`, read at `all` with the warning, and a save under `all` holding no `External` relationship.
- `test/package-wiring/`'s and `check75/`'s files under `word` and `all`, each reading as under `duplicateEntries: 'first'`, and the `rn-*` files once CR-025 section 8 has them.
- The option on `Texts.of` and in the suite's adapter; `repaired` on each of the three outcomes; the contradiction (D2); `checkLossOnSave`'s start under each value.
- A Word check of the saves `all` writes for `x-ext-*`, as docx4j's and python's were checked (silent), on Jason's word (D7).

## 4. Phases and size

| Phase | What | Size |
|---|---|---|
| 1 | The option, `word` over CR-024's two readings, `repaired`, `checkLossOnSave`'s start, the text model's option | ½ day |
| 2 | `External`: the refusal, `all`'s repair and warning, REQ-321's refusal | ½ day |
| 3 | 2.3 with CR-025 section 8; the README's Repairs section; CLAUDE.md, the index | ¼ day |

No dependency on another repository.

## 5. Decisions

| | Question | Recommendation |
|---|---|---|
| D1 | Keep `duplicateEntries` beside `repair`? | Keep it. The WD keeps `duplicate_entries` as REQ-206's own option, and python keeps both. `repair: 'word'` implies it |
| D2 | `repair: 'word'` with an explicit `duplicateEntries: 'refuse'`? | Refuse the call as contradicting itself, as CR-012 section 29 refuses `include` against `comments`. Silently letting one win hides a mistake |
| D3 | Refuse `External` beside an internal relationship at `refuse` and `word`? It reads today. | Yes, as the WD now states, and as docx4j and python do. Word refuses such a file even through its repair, so `word` has no reading to give. It's a `Changed` line: a package that loads today stops loading |
| D4 | Refuse `External` alone at the load (REQ-321)? | Yes, naming the URL, as docx4j does. Today it loads as a package with no main part, which every reader then refuses less clearly |
| D5 | Build 2.3 here or in CR-025 section 8? | Here, with section 8's refusal in the same build, so the shape is refused and readable from one release. Section 8's D11 already points here |
| D6 | Does `word` or `all` turn `checkLossOnSave` off? | Yes, by starting it false, as CR-025 section 2.4 recorded, and docx4j's gate gives way under `WORD` too. A caller who asked for Word's repair has accepted Word's reading, which drops what Word drops |
| D7 | A Word check of the saves `all` writes? | Yes, eight files, one round, as docx4j's and python's had. It's the only evidence that a repaired save opens |
| D8 | Which release? | 0.7.0, so the three engines agree on `External` when it ships. Otherwise core-ts 0.7.0 reads a file the other two refuse |

## 6. Registry

`core-ts/CR-031`, `depends_on: [core-ts/CR-025]` for section 8's refusal. Sent once this is committed.

## 7. Sources

- `docx4j-portfolio/docs/text_extraction_specification_wd.md`: WD 0.102 at `4c81f10` (REQ-207, REQ-208, REQ-206's NOTE), WD 0.103 at `eceabd6`; REQ-801's metadata; REQ-1001's table.
- `docx4j-portfolio/docs/repair_policy_across_ports.md` section 4.3, and Jason's words on `all` quoted there.
- docx4j `880d7a3bb` (`RepairMode.ALL`, `followsWord()`, `makesOwnRepairs()`, `TextResult.REPAIRED`), `6025529d9`; its CR-075.
- docx4j-python `e05291b`, `30eab21`: `LoadOptions(repair=)`, `ExternalMainPartException`; its CR-032.
- The share's `bridge/portfolio/x-ext-*`.
- CR-024 sections 9 and 9.2, CR-025 sections 2.4 and 8.
