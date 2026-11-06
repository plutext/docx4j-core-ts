# CR-025: Content a read drops is reported, and a save doesn't lose it unasked

**Status:** Proposed 2026-11-05, at Jason's "we'd better draft those CRs please", after the portfolio session's question on the
repair policy's P7. Measured here the same day (section 1). D1 to D8 decided 2026-11-05, all as recommended (Jason: "CR-025 decisions as
recommended"; section 5), D1's code agreed with python the same day (its CR-033 Q9, `f96f463`); D2 is P7 for this package. D4, which waited for Word check 76 (run 2026-11-05), decided the same day as recommended (Jason: "CR-025 D4 as recommended"): what a read drops stays dropped. D9, raised afterwards, decided the same day as recommended (Jason: "CR-025 D9 as recommended"): CR-021's pass reports what it removes. Its wording decided 2026-11-06, python's narrower reading (Jason: "CR-025 D9's wording: take python's narrower version"): what goes with its content. Docs only, nothing implemented. Its sibling is docx4j-python's CR-033 (phase B, proposed), whose names this CR follows where the two packages'
shapes allow. D7 changed 2026-11-06 at Jason's word: this CR goes in 0.7.0. Implemented 2026-11-06, phases 1 to 3 (sections 2.8 to 2.10), unreleased. CR-026 depends on this CR.

## 0. Summary

When a part is unmarshalled, the runtime drops what the model has no place for: every attribute it doesn't map, and an unknown
element in 12 of the 32 hosts tried (in the other 20 the element fails the part, which is CR-026). Nothing is said. The tree no
longer holds them, so a save that re-marshals the part writes it without them. An untouched save is safe, but the loss needs
no edit: `getContents()` then `save()` is enough.

jsonix has been able to report all of this since its CR-006. This package has never asked. The proposal:

1. Ask, in `XmlPart.unmarshal`, and keep each dropped item as a `part.skipped` warning.
2. At save, copy the source bytes of a part that holds skipped content and was read but not changed, so that reading costs nothing.
3. Refuse the save of a part that was changed and would lose skipped content (`SaveLossException`, `save.content_lost`), unless
   the caller has said the loss doesn't matter (`pkg.checkLossOnSave = false`).

The case is rare in what we hold: 1 package of 720, one Word 2010 setting. Step 2 is what stops that file's readers being
refused. The model can lose more than that corpus shows, a password hash among it (section 1).

## 1. Here today (measured 2026-11-05, `b07134f`, objects 0.7.0, `@docx4j/jsonix` 3.11.0)

| Dropped by the read | Untouched save | `getContents()`, then save | Said |
|---|---|---|---|
| An attribute the model doesn't map | kept, byte for byte | **lost** | nothing |
| An unknown element in `w:document`, `w:pPr`, `w:rPr`, `w:tblPr`, `w:tblGrid`, `w:tcPr`, `w:sdt`, `w:sdtEndPr`, `w:sectPr`, `w:style`, `w:lvl`, `w:settings` | kept, byte for byte | **lost** | nothing |

- **The attributes tried:** on the first `w:p` and `w:r` of `loadAndSave.docx`, one each in the `wordml201011` namespace, in a
  private namespace (both listed in `mc:Ignorable`), and an unknown name in the `w:` namespace. All six went.
- **"Nothing" means:** `part.warnings` and `pkg.warnings` empty after the read and after the save, `pkg.lastSave`'s three lists
  empty, and no console output.
- **What survives:** `mc:Ignorable` and the `xmlns` declarations, so the saved part can declare prefixes it no longer uses.
- **A private read doesn't count:** `readContents()` leaves a part to be copied, and `getPropertyResolver()` alone left the
  attributes below in place. The loss follows `getContents()`.

**A real file.** `test/fixtures/parity/NumberingImplicitNumId.docx` (Word 2010, `AppVersion` 14.0000) has a
`w:stylePaneFormatFilter` holding `w:val="3F01"` and 15 more attributes (`w:allStyles`, `w:customStyles`, ...). The model maps
only `w:val`, so a re-marshal writes `<w:stylePaneFormatFilter w:val="3F01"/>`. Objects CR-019 finds the 15 repeat `w:val`'s bits,
so this is a loss of form rather than state.

**The corpus.** Every package committed under `test/`, through the real load path with the callbacks attached in memory:

| | |
|---|---|
| Packages | 724: 720 unmarshalled in full, 4 refused at the load (check 75's duplicate entries) |
| Attributes dropped | 15, all the file above |
| Elements dropped | 0 |
| Text dropped | 1: check 72's `72m`, already reported as `target.misplaced` (CR-016) |

This corpus was gathered for other reasons, so it says the case is rare in what we hold and no more. docx-platform-tests' four
content-control inputs, read as data (CR-023 section 1), hold both shapes. python reports its samples' `part.skipped` empty too.

**What the model can lose beyond the corpus.** Objects CR-019 compared `wml.xsd` with ISO 29500's transitional schema. Each of
these was then measured through the facade here. None is in the 724 packages, and how often they occur in the wild isn't known:

| Element | Dropped | What's left |
|---|---|---|
| `w:documentProtection` | `w:algorithmName`, `w:hashValue`, `w:saltValue`, `w:spinCount` | `w:edit` and `w:enforcement`: **the restriction stays and the password hash goes**, a loss of state |
| `w:writeProtection` | The same four | An empty element |
| `w:cnfStyle` | Its on/off attributes | `w:val` |
| `w:numFmt` | `w:format`, and `w:val="custom"` is refused too | `<w:numFmt/>` |
| `w:tc` | `w:id` | The cell |

The older `w:hash`, `w:salt` and `w:crypt*` attributes round trip.

**The other ports.**

| | The read | Reported | A re-marshal |
|---|---|---|---|
| docx4j 17.3.5 | drops (measured: 24 of 26 hosts, CR-026 section 1) | no (the repair policy's P7; `broken_producer_repairs_across_ports.md` section 4 item 4) | loses it |
| docx4j-python | drops | `part.skipped` (`SkippedNode`) | loses it, nothing on `last_save`, until its CR-033 phase B refuses |
| this package | drops | nothing | loses it |

**Word.** Check 36: Word 2010's saves drop what it doesn't know, declared or not (`w15:appearance`, `w16du:dateUtc`), and Word
365's saves keep those two, which it knows. Check 76 (CR-026 section 4) has Word 365 the same: its saves drop what it doesn't know,
attribute or element and all the element held, with no message, or it refuses the file. The exception is what an unknown `w:` element
holds in a content host, which it reads (CR-026 section 4.4, D7).

## 2. Proposal

### 2.1 The report

`XmlPart.unmarshal` passes `onUnexpectedAttribute`, `onUnexpectedElement` and `onUnexpectedText` beside its three existing
callbacks. They stay at that one call: objects CR-007 keeps them per call because a drop is sometimes the mechanism
(`recordPPrChange`'s `deepCopyAsSync` drops what `CT_PPrBase` doesn't declare, on purpose), so nothing changes for a deep copy.

Each dropped item is one warning on the part, gathered by `pkg.warnings`:

```ts
interface ContentSkippedWarning {
  code: 'part.skipped';
  partName: string;
  kind: 'element' | 'attribute' | 'text';
  name: string;         // 'ext:opaque'; for text, its first 80 characters
  parentName: string;   // 'w:sdtPr'
  parentClass?: string; // 'org_docx4j_wml.SdtPr'; absent for what CR-021's pass removes
  path: string;         // '/w:document/w:body/w:sdt/w:sdtPr/ext:opaque'; an attribute's last step '@ext:a'
  ignorable: boolean;   // its namespace is listed in the part's mc:Ignorable
  line?: number;        // where the parser gives one, as target.misplaced
  column?: number;
  message: string;
}
```

`kind`, `parentClass`, `path` and `line` are python's `SkippedNode`, with `qname` and `parent_qname` as `name` and `parentName`.
`ignorable`, `column` and the `text` kind are this package's. A string CR-016's walk already reports as `target.misplaced` isn't
reported twice. Parts are unmarshalled lazily, so a part never read reports nothing, and has lost nothing.

**One route doesn't reach the callbacks (D9).** CR-021's pass runs first, so an element in a `w:t`, `w:instrText`, `w:delText` or
`w:delInstrText` is gone before the runtime sees it: `<w:t>a<ext:opaque/>b</w:t>` becomes `<w:t>a</w:t><w:t>b</w:t>`. All that shows
is `br.outside_run`'s `textElements` count, which names nothing, and the gate can't see it. Phase 1 has the pass add a `part.skipped`
for each element that goes with all it holds, where `givesItsText` says no, so the report and the gate cover it. An element that gives its text
isn't listed, as in python's report, since `part.skipped` is shared: no words go with it. Check 76 raises the
stakes: Word drops such an element with its text, and since CR-026's D6 so does the pass, so words go here with only a count to show for them.
Reported so, such a part comes under 2.2: read and not changed, it's copied with the element still in it, where a save today writes it
repaired. If Word refuses that file (`76v`), it stays refused until the caller gives the part its tree back, as 2.4 says of `76e`.
An element in no namespace gives its text and isn't listed, so a read part holding one is written without it, as a save does today.

### 2.2 A part read and not changed is copied (D3)

The save re-marshals every unmarshalled part because it can't tell a part that was read from one that was changed. For a part
holding `part.skipped` warnings, and only such a part, it can afford to find out: unmarshal the source bytes again and compare
that tree's marshal with the current tree's. If they're equal, nothing changed, so the source bytes are copied as for an untouched
part, with nothing lost and nothing refused. Like an untouched part it then skips what a marshal does: the normalising (CR-015) and
the checks for values past Word's range and for misplaced content, so `lastSave` lists nothing for it.

Tried on the settings part of the real file above: equal after a read, unequal after setting `zoom`, about 15 ms including a
package load. That is one part. Phase 2 has to hold it over every fixture with skipped content injected.

### 2.3 The gate (D2, the repair policy's P7)

A part that holds skipped content and was changed would lose that content. Beside CR-016's check, before anything is written, the
save throws `SaveLossException` (`code: 'save.content_lost'`, `losses`: one `{ partName, kind: 'skipped', count }` per part), with
python's message and this package's way out:

> Saving would lose content: /word/document.xml holds 1 item the load skipped (part.skipped). Set pkg.checkLossOnSave = false to write the part without it.

What the gate never refuses: a part not re-marshalled, a part 2.2 copies, a part given new contents with `setContents` (the caller
replaced the tree).

The gate asks more than Word does. Where Word 365 opens such a file, its own save drops this content with no message (check 76). `checkLossOnSave` is
how a caller takes Word's behaviour.

### 2.4 Saving anyway (D8)

A caller who doesn't care about the loss sets `pkg.checkLossOnSave = false` (on by default). The part is written without the items,
and each goes to a new `pkg.lastSave.lost`, so the save still says what went. It is a package property like `normaliseOnSave`,
`checkContentOnSave` and `writeTimesOnSave`: copied by `clone()`, settable after a refusal, and reaching `save`, `saveFlatOpc` and
`saveTo` alike, where a per-call flag would have to go on all three. 2.2 still runs, so a part read and not changed is copied
whatever the setting: the switch only decides what happens to a part that really would lose something.

It is deliberately not the `repair` switch python's CR-033 uses here. In CR-033 section 4.4 that switch is the caller's Yes to Word's
prompt, given in advance, and decides only whether a lossy save may be written, the load reading the same either way. This package
has its own switch for three reasons:
- A caller here is accepting a loss, not a Word repair. Check 76 bears this out: Word repaired none of this content. It dropped it
  with no message, or refused the file and couldn't repair it (`76e`), so in no case measured is there a prompt for a Yes to stand for.
- The policy's switch is wider than this one question: R2 leaves `duplicate_entries: first` to fold into it, which would change what
  a load reads (CR-024 section 2.7).
- `...OnSave` is this package's form for what a save does.

So this CR builds no `repair`; if core-ts builds that switch, turning it on turns this check off too. It's a recorded difference
between the ports: python keeps `repair` (its CR-033 Q5, Jason's ruling of 2026-11-05), to be looked at again if P4 folds
`duplicate_entries` into that switch.

**A file Word refuses.** `76e` will read here under CR-026, and a read and a save leave it as it was: 2.2 copies the part, element and
all, so Word goes on refusing it. Nothing is made worse, and `part.skipped` has said what's there. For the save to be the fix, the
caller gives the part its tree back, `part.setContents(await part.getContents())`, and the part is written without the element. So 2.2
has to pass over a part given new contents, as 2.3 does.

### 2.5 Keeping instead of dropping (D4)

An attribute could be kept: recorded against its typed object on the read and written back on the marshal, which the facade's
source-binding records and CR-006's binder suggest is possible, though nothing here has tried it. The portfolio's
`benchmarking.md` has this as an open question for Jason (keep, or drop as Word drops unknown ignorable content on save). Check 76
answers the second half for Word 365: it drops the markup, ignorable or not, or refuses the file. Kept content would give the gate nothing to refuse, but it
would last only until the next Word save, so it stays dropped (D4, decided 2026-11-05; Jason: "CR-025 D4 as recommended").

### 2.6 Not here

- An element that fails its part today: CR-026.
- A single-valued child given twice in one parent: the runtime keeps the last with no callback, so the report can't see the first, and the gate can't refuse its loss. It isn't only a wildcard such as `a:ext`. Measured on the model 2026-11-06, `w:jc`, `w:b`, `w:sz`, `w:pPr`, `w:rPr`, `w:sectPr` and `w:tblPr` go the same way, and so do `w:sdtContent`, `w:txbxContent` and `w:body`, which take the first's words with them. jsonix traces it to each child being its own `unmarshalProperty` call, so its `Value already set.` never meets the second. It's outside the schema, and none of the 997 packages read here holds it: comparing each source with its save, the only children lost are CR-021's repairs. What Word does with it is unmeasured. A report would need the runtime, a later jsonix CR if Jason wants one.
- `w:stylePaneFormatFilter`'s attributes in the model: objects CR-019 (D6), docx4j first, with the other gaps in section 1 left to
  its D4.
- `mc:MustUnderstand` and `mc:ProcessContent`: nothing under `src/` reads them, and neither of this CR's measurements involves them.
- A `DefaultXmlPart`, which has no model to drop anything.

### 2.7 The CHANGELOG lines

- **Changed:** a save that would lose content the read dropped is refused with `SaveLossException` (`save.content_lost`) unless
  `pkg.checkLossOnSave` is false. Previously the content was lost with nothing said. A part read but not changed is now copied byte for byte
  where it holds such content.
- **Changed:** `clone()` is refused the same way, since it goes through a save. Under `mcePreprocess: true`, a part read and not changed that holds such content is saved in its source form, its `mc:AlternateContent` kept.
- **Added:** `ContentSkippedWarning` (`part.skipped`) on the part and on `pkg.warnings`; `pkg.checkLossOnSave`,
  `pkg.lastSave.lost`; `SaveLossException`, `ContentLoss`.
- **To expect:** a changed part holding attributes the model doesn't map is refused, so a Word 2010 file's `settings.xml` with `w:stylePaneFormatFilter`'s extra attributes refuses a save once `changeTrackingMode` is set, until objects CR-019 lands. `checkLossOnSave` is the way past it.

### 2.8 Phase 1 as built (2026-11-06)

Built at Jason's "after checks 78 and 79, you can start CR-025 phase 1": `src/opc/skipped.mts` holds the warning and its three makers, `XmlPart.unmarshal` passes the three callbacks, and CR-021's pass reports what goes with all it holds. A save is unchanged, since the gate is phase 2.

Five things the build settled:

- **`parentClass` is optional.** No class meets what CR-021's pass removes, so it's absent there, as in python's `skipped_node`. Every other warning has a `message`, so this one does too.
- **A text's path ends `text()`,** which 2.1 left open.
- **`mc:Ignorable` off a root** is itself reported, where its class has no place for it, because a re-marshal loses it. That includes a chartsheet's root (`p2-chartsheet.xlsx` in docx4j's `charts/`, with its `xr:uid`, both measured lost).
- **A prefix that no longer resolves counts by its name.** `mcePreprocess` takes content out of the `mc:Choice` declaring its prefix, and such content would otherwise read `ignorable: false`.
- **A Choice taken can report what a kept load doesn't.** `cr022-checkbox-linked.xlsx` under `mcePreprocess: true` reports `a14:legacySpreadsheetColorIndex` and `mc:Ignorable` on two `a:srgbClr`, which is CR-004 section 7's binding gap, silent until now. Phase 2's gate will refuse such a part if it's changed.

`test/content-skipped.test.mjs` holds it. Three tests that pinned a whole warning list were narrowed: `br-placement.test.mjs`'s `w:t` files, as D9 said they would be, which now name what went, and two in `mce.test.mjs`, one of whose files puts a `w:txbxContent` straight in a `w:drawing`.

### 2.9 Phase 2 as built (2026-11-06)

Built at Jason's "commit then do CR-025 phase 2". `XmlPart.isUnchangedSinceRead` is 2.2's test, and `bytesRead` the bytes it holds the tree to. `Saver.settleLosses` is the gate, `SaveLossException` its refusal, and `pkg.checkLossOnSave` the way past it.

Six things the build settled:

- **The gate runs before anything is written, over the parts the save reaches.** `sink.put` writes as it goes on a directory sink, so a refusal in the middle of the walk wouldn't have held to "nothing written". So the gate walks the relationships first, which is also how `losses` can name every part at once, and how a part no longer reached can't refuse a save.
- **A part given bytes is held to those bytes.** `getContents` drops `this.bytes`, so the part keeps them for 2.2 (`parsedBytes`). `setBytes` leaves the old warnings in place until the next read, so the gate reads `skippedByRead`, which is empty for a part not unmarshalled.
- **`clone()` is refused as a save is,** since it goes through one. That's CR-016's behaviour too.
- **What can't be compared counts as changed,** so the gate decides it. A tree that won't marshal is CR-016's anyway.
- **A part `mcePreprocess` resolved, read and not changed, is saved in its source form,** `mc:AlternateContent` and all, if it holds skipped content. `cr022-checkbox-linked.xlsx`'s drawing keeps its two `a14:legacySpreadsheetColorIndex` attributes through a resolved load, where they used to go with nothing said, and `sml-extensions.test.mjs` pinned that loss. Parts with nothing skipped are resolved as before.
- **A package setting on `settings.xml` is the case to expect.** `NumberingImplicitNumId.docx` with `changeTrackingMode` set is refused for its 15 `w:stylePaneFormatFilter` attributes, which repeat `w:val`'s bits, until objects CR-019 lands. The way past is `checkLossOnSave`. It's pinned in the test.

Over all the fixtures, each with an attribute injected into its first paragraph, unmarshalled, its paragraphs read and `Texts.of` run, the part is copied byte for byte, so a read changes nothing a marshal shows. Made to answer "unchanged" always, six tests fail. Made to answer "changed" always, six more do. npm test passes 1,884 of 1,884.

### 2.10 Phase 3 as built (2026-11-06)

The records, at Jason's "commit, then CR-025 phase 3: go": CLAUDE.md's status, key mechanics, layout and tests list, `test/README.md`, and the index. This repository keeps no CHANGELOG file, so section 2.7's lines are for the 0.7.0 release's notes, which go on the GitHub release.

## 3. Tests and oracles

`test/content-skipped.test.mjs`:
- section 1's six attributes and twelve hosts, each with its warning's fields, held to the source;
- `deepCopyAsSync` and `recordPPrChange` reporting nothing (objects CR-007's rule);
- an untouched save byte for byte; a part read and not changed copied byte for byte, attributes kept;
- a part changed: refused with nothing written, the message and `losses`; with `checkLossOnSave` off written, with
  `lastSave.lost`, through `save`, `saveFlatOpc` and `saveTo`; `clone()` carrying it; a part read and not changed copied with it off too;
- an element in a `w:t` that goes with all it holds reported by CR-021's pass, the gate counting it, and one that gives its text not reported (D9);
- `NumberingImplicitNumId.docx` reporting its 15 and no other fixture reporting any, a count guard that holds on a clean checkout,
  since git-ignored built fixtures aren't on CI. Objects CR-019 takes those 15 to 0 when it lands, so this is the only test that may
  rest on that file: the fields and the gate use injected attributes.

The gate re-runs CR-016's corpus: every file saves unrefused after unmarshalling, the real file through 2.2. `test/nodenext` names
the new types. The oracle for the fields is python's `SkippedNode` output on the same files, compared once by hand.

## 4. Phases and size

| Phase | What | Size |
|---|---|---|
| 1 | 2.1: the report | ½ day |
| 2 | 2.2 and 2.3: the unchanged test, the gate, `checkLossOnSave` | 1 day |
| 3 | The records: CHANGELOG, CLAUDE.md's key mechanics, `test/README.md`, the index | ¼ day |

No dependency on another repository. Phase 1 is additive and could ship alone.

## 5. Decisions (taken 2026-11-05, all as recommended; D7 changed 2026-11-06)

| | Question | Recommendation |
|---|---|---|
| D1 | The warning's code. `part.skipped` is python's member name, not yet a shared code. | `part.skipped`, settled with python first, as `target.misplaced` was. Agreed: python's CR-033 Q9 (`f96f463`) |
| D2 | P7 for this package: does a save refuse, or only report? | Refuse, as python's CR-033 does and CR-016 does for misplaced content. With 2.2 only a changed part can be refused. The cost is a `Changed` line: a save that worked today throws |
| D3 | Copy a part read but not changed (2.2)? python's gate had no such step. | Yes, and python took it too (its CR-033 Q8). Without it the real file's readers are refused for a style pane filter they never touched |
| D4 | Keep the content rather than drop it (2.5)? | Not now: decide after check 76. Check 76 is in and Word drops it, so: no, drop. Decided 2026-11-05 as recommended (Jason: "CR-025 D4 as recommended") |
| D5 | The `text` kind: python reports elements and attributes only. | Keep it. Only `72m`'s string is dropped in the corpus, and CR-016 already reports that, so it costs nothing and closes the gap for text elsewhere |
| D6 | An objects CR to map `w:stylePaneFormatFilter`'s attributes? It's the only real loss measured. | Yes, small, and separate from this CR. Written as objects CR-019 (2026-11-05, proposed): its route is docx4j first |
| D7 | Which release? | In 0.7.0, with CR-026 (Jason, 2026-11-06: "CR-025 D7 and CR-026 D5 both say 'after 0.7.0'; please correct to in 0.7.0"). Decided after 0.7.0 on 2026-11-05, because this changes what `save()` does. 0.7.0 holds CR-016's refusal already, so a caller meets both in one release |
| D8 | The way past the gate: this package's own `pkg.checkLossOnSave`, or python's `repair: 'word'`? | Its own (2.4), reported on `lastSave.lost`. A caller who doesn't care is accepting a loss, not a Word repair, and the policy's switch is wider than this one question |
| D9 | Raised after D1 to D8: should CR-021's pass report what it removes as `part.skipped` (2.1)? | Yes. Otherwise this CR's report and gate have a hole at `w:t`, the one place an element can go with only a count to show for it. Since CR-026's D6 its text goes too. Decided 2026-11-05 as recommended (Jason: "CR-025 D9 as recommended"). Which elements: only what goes with its content, python's reading (its CR-030 section 11), since `part.skipped` is shared. The first wording, every element but a `w:br` or a nested text element, also listed what gives its text, which would have left a file Word refuses (`76x`, `76y`) as it was and refused a changed part though no words went. Decided 2026-11-06 (Jason: "CR-025 D9's wording: take python's narrower version") |

## 6. Registry

`core-ts/CR-025`, no dependency on another repository. docx4j-python's CR-033 is its sibling; the repair policy's P7 and
P4 are the portfolio's. Done follows phase 3's commit.

## 7. Sources

- Measured with read-only probes outside the repository: the attributes through a re-marshal, the element hosts, the corpus (the
  runtime's callbacks attached in memory to the real load path), and 2.2's comparison.
- jsonix `nodejs/scripts/types/main.d.ts`: `UnmarshalCallbacks`, its CR-006 and CR-013.
- objects CR-007 (why the callbacks are per call), its `src/index.mts` `recordingOptions`.
- docx4j-python: CR-033 sections 1.2 and 4.3; `docx4j-xsdata` `parsers/skipped.py`.
- Portfolio: `repair_policy_across_ports.md` (R1, R4, P4, P7), `broken_producer_repairs_across_ports.md` section 2 row 3,
  `benchmarking.md`.
- `test/README.md` check 36; CR-016; CR-023 section 1.
