# CR-026: An element the model has no place for no longer fails its part

**Status:** Proposed 2026-11-05, at Jason's "we'd better draft those CRs please", from the finding of 2026-11-03 (CR-023 section
1: an unknown element in `w:sdtPr`). Measured here the same day (section 1), and three routes compared, two tried as prototypes in
memory (section 2). D1 to D5 decided 2026-11-05, all as recommended (Jason: "CR-026's decisions as recommended"; section 7). Docs
only, nothing implemented. It depends on CR-025 and on a jsonix CR not yet written (D2). Word check 76 (section 4) isn't written yet.

## 0. Summary

The finding was narrower than the defect. An element the model doesn't know fails the whole part, not only in `w:sdtPr` but in 20
of the 32 hosts tried, among them `w:body`, `w:p`, `w:r`, `w:tbl`, `w:tr`, `w:tc`, `w:hdr` and `w:comment`. In the main part that
means nothing can read the document. It makes no difference whether the element's namespace is in `mc:Ignorable`, the one case
markup compatibility says a consumer must get past. docx4j reads such a document, and so does python by default.

The decision (D2) is to skip the element and report it at the point the runtime throws today, a small jsonix change, which a
prototype shows reads all 52 cases with no change to any reader. The element is then dropped, as docx4j and python drop it, so
this has to land after CR-025: reading it without CR-025 would turn a loud refusal into a silent loss.

## 1. Here today (measured 2026-11-05, `b07134f`, objects 0.7.0, `@docx4j/jsonix` 3.11.0)

Two elements in each of 32 hosts: `ext:opaque`, its namespace in `mc:Ignorable`, and `w:madeUp`, an unknown name in the `w:`
namespace. Both behaved alike in every host.

| The read | Hosts |
|---|---|
| **Throws** (20) | `w:body`, `w:p`, `w:r`, `w:hyperlink`, `w:fldSimple`, `w:ins`, `w:smartTag`, `w:customXml`, `w:tbl`, `w:tr`, `w:trPr`, `w:tc`, `w:sdtPr` (block and run), `w:sdtContent` (block and run), `w:ffData`, `w:hdr`, `w:comment`, `w:footnote` |
| Drops it, silently (12; CR-025) | `w:document`, `w:pPr`, `w:rPr`, `w:tblPr`, `w:tblGrid`, `w:tcPr`, `w:sdt`, `w:sdtEndPr`, `w:sectPr`, `w:style`, `w:lvl`, `w:settings` |

Through the real load path, the 26 hosts under `w:document` as 52 files: 34 throw `PartUnmarshalException` and 18 read. The numbers
are the same under `mcePreprocess: true`, and an untouched save stays byte for byte.

**The cause.** A class whose content is one list of elements has a property the mapping marks `allowDom: false` (94 across 19
modules, 34 in `org_docx4j_wml`). For an element it can't type, `Jsonix.Binding.Unmarshalls.Element.unmarshalElement` throws
"could not be unmarshalled as is not known in this context and the property does not allow DOM content". A class of named
properties takes the other branch, which skips the element and calls `onUnexpectedElement`. So which of the two happens depends on
how a class is mapped, not on what the element is.

**`mc:Ignorable` plays no part.** Nothing on the load path reads it; `src/` only writes it (`declareIgnorable`).

**How often.** None of the 720 packages under `test/` holds such an element (CR-025 section 1), so nothing here has met this in a
file it held. docx-platform-tests' four content-control inputs fail on it, which is how it was found. That corpus says little about
files in the wild.

**The other ports.**

| | The same element |
|---|---|
| docx4j 17.3.5 | Measured here on the same 52 files: 48 read, the element dropped; 4 throw, all `w:sdtPr`. Its `86a203f16` (on `VERSION_17_3_6`, unreleased; not run here) fixes those, a false match on "entity" in the fallback's error test, and reports 2 of the suite's 4 scenarios passing, 2 failing on preservation |
| docx4j-python | Read by default, listed in `part.skipped`; strict mode raises, naming them; a re-marshal loses them (`benchmarking.md`) |

**Word.** Check 36: Word 2010 refused `w15:appearance` in a `w:sdtPr` when `w15` wasn't ignorable ("problems with the contents", no
repair offered), and opened the declared twin. Word 365 opened the undeclared file with no message, but it knows `w15`. So the only
measurement of Word against a namespace it doesn't know is one element in Word 2010. ECMA-376 Part 3 has a consumer treat such markup as an error unless its
namespace is ignorable. What Word 365 does, in any host, with an ignorable or undeclared namespace or an unknown `w:` name, is
unmeasured: section 4.

## 2. Three routes

| | A: keep it as DOM | B: skip and report at the throw | C: a DOM pass here |
|---|---|---|---|
| What | `allowDom: true`, so the tree holds the element | The runtime's refusing branch does what its skipping branch does | Before unmarshalling, remove elements whose namespace is in `mc:Ignorable` and not understood |
| Tried | In memory, over the 52 files | In memory, over the 52 files | Not tried |
| Reads | All 52 | All 52 | Ignorable elements only |
| Readers | **16 of the 34 crash** on a DOM node in `content`: 12 in `Texts.of`, 4 in the content API and `Texts.of` | Unaffected | Unaffected |
| Save | Kept in the 34; still lost in the other 18, whose classes have no such property | Lost, reported | Lost, reported if the pass reports |
| Changes | The objects package's declared types, and every walker | jsonix, one branch | One more walk here |
| `w:madeUp` | Read | Read | **Still fails** |
| A newer `w14:` element the model lacks | Read | Read | **Still fails**: the namespace is understood |
| Other modules | Per mapping | All, from one branch | Word only, unless extended |

**B, decided (D2).** It's the only route that reads every case with no change to any reader, and it makes the 20 hosts behave as
the 12 already do. The prototype wrapped each property's `unmarshalElement` in memory (4,187 properties, every module) and read
all 52 files, each reported, `Texts.of` unaffected, each element lost on save. Only the WordprocessingML hosts were tested.

A keeps the content, which is the better outcome if keeping is wanted, but it costs a reader audit and still leaves 18 of 52 cases
dropped, so it needs D1's answer first. C is the only route with no sibling dependency, and it fixes the wrong test: a namespace
list can't say what the model holds.

## 3. Proposal

### 3.1 The runtime (jsonix-CR-015, to be written)

In `Jsonix.Binding.Unmarshalls.Element.unmarshalElement`, where an element is neither typed nor allowed as DOM: if the caller set
`onUnexpectedElement`, call it and skip the element; otherwise throw as today. A caller who hasn't asked sees no change, so the
facade's `unmarshalString` still refuses. The real change has to get the input's position right for the loop that called it, which
the prototype's wrapper sidestepped. Then raise the floor here.

### 3.2 This package

With CR-025's callback in place, the read needs no code of its own: a skipped element is a `part.skipped` warning, and the save
copies, refuses or repairs as CR-025 says. What changes is that a main part which failed on this cause now reads. Every other cause
of `PartUnmarshalException` stays, and CR-012 section 25's rule that a main part which can't be read is refused stands.

`insertOoxml` needs tracing in phase 2: a caller's own fragment holding an unknown element is arguably better refused than read
with a warning on a package the caller never sees.

### 3.3 Reading without the element

A dropped element can hold text (`<ext:opaque>words</ext:opaque>`), which then doesn't read. For an ignorable namespace that is what
markup compatibility asks. For one that isn't ignorable, the text model would be silent about text a reader might want to know
about; the WD has no statement on either (D4).

### 3.4 Not here

- Keeping the element (D1).
- `mc:MustUnderstand` and `mc:ProcessContent`: nothing under `src/` reads them. An element with `mc:ProcessContent` should have its
  children read in its place, which skipping doesn't do. Check 76 doesn't cover them.
- Word-matching refusal of what Word refuses: the repair policy's tier 3 is read so that a caller can fix it.

### 3.5 CR-023

The suite's four content-control scenarios then load. Their two preservation rows go on failing unless D1 keeps the element, as
docx4j's do.

## 4. Word check 76 (phase 0, to be written)

Word 365 only. `test/fixtures/check76/build.mjs` builds one small file per case, each with marked paragraphs before and after:

| File | Case |
|---|---|
| 76a to 76d | `ext:opaque`, `ext` in `mc:Ignorable`, in `w:sdtPr`, `w:p`, `w:r`, `w:tc` |
| 76e | `ext:opaque` in `w:p`, `ext` declared, not ignorable |
| 76f, 76g | `w:madeUp` in `w:p`, in `w:pPr` |
| 76h, 76i | Attribute `ext:a` on `w:p`, ignorable, not ignorable |
| 76j | Attribute `w:madeUp` on `w:p` |
| 76k | `<ext:opaque>words</ext:opaque>` in `w:p`, ignorable: do the words show? |

Each open records any prompt (the policy's tier), what shows, and, Save As first, what Word's save keeps. That answers D1 and D3
and gives each shape its tier. It blocks neither D2 nor the build: B reads in every tier.

## 5. Tests and oracles

- `test/unknown-elements.test.mjs`: section 1's 32 hosts, both elements, each read, its warning's fields, `Paragraph.text`,
  `Texts.of` and a save; a module beyond WordprocessingML (a DrawingML and a SpreadsheetML host), since the prototype tested neither.
- docx4j's `UnexpectedElementInSdtPrTest` mirrored: inline and block, ignorable and not.
- The docx4j oracle, once `VERSION_17_3_6` is released: the 52 files, every host read, every element dropped.
- A caller with no callback still gets the runtime's refusal. No test here asserts today's message, so none has to change.
- Check 76's saves under `test/check76/`, with a SHA-1 manifest.

## 6. Phases and size

| Phase | What | Size | Waits on |
|---|---|---|---|
| 0 | Check 76: the files, Jason's run | ½ day here | Nothing |
| 1 | jsonix-CR-015 written, released; the floor raised | ½ day there | D2 |
| 2 | The read, `insertOoxml` traced, the tests | ½ day | CR-025 phases 1 and 2; phase 1 |
| 3 | The records, CR-023's line | ¼ day | |

## 7. Decisions (taken 2026-11-05, all as recommended)

| | Question | Recommendation |
|---|---|---|
| D1 | Keep the element, or drop it? `benchmarking.md` holds this open across the ports. | Drop, as docx4j and python do, and revisit with check 76. Keeping (route A) costs a reader audit and still can't cover 18 of the 52 cases |
| D2 | Route A, B or C? | B |
| D3 | Read an element whose namespace isn't ignorable, or an unknown `w:` name, or refuse it? | Read and report (`ignorable: false`): docx4j and python read it, and the policy's tier 3 reads so that a caller can fix it. Check 76 gives the tier |
| D4 | Should the text model warn of an element dropped from a paragraph? | A question for the WD, sent on commit; nothing here until it rules |
| D5 | Which release? | With CR-025, after 0.7.0 |

## 8. Registry

`core-ts/CR-026`, proposed, depending on `core-ts/CR-025` and on the jsonix CR once written (D2). CR-023's
line stands. Sent on commit.

## 9. Sources

- Measured with read-only probes outside the repository: the host matrix through the facade and the real load path, docx4j 17.3.5
  through `WordprocessingMLPackage.load` on the same files, and routes A and B patched in memory.
- jsonix `nodejs/scripts/jsonix-factory.js` (`Jsonix.Binding.Unmarshalls.Element`, the class's skipping branch).
- docx4j `86a203f16`, `UnexpectedElementInSdtPrTest`.
- `test/README.md` check 36; CR-023 section 1; CR-025.
- Portfolio: `benchmarking.md`, `repair_policy_across_ports.md` (tier 3, R4).
