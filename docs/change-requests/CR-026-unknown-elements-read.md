# CR-026: An element the model has no place for no longer fails its part

**Status:** Proposed 2026-11-05, at Jason's "we'd better draft those CRs please", from the finding of 2026-11-03 (CR-023 section
1: an unknown element in `w:sdtPr`). Measured here the same day (section 1), and three routes compared, two tried as prototypes in
memory (section 2). D1 to D5 decided 2026-11-05, all as recommended (Jason: "CR-026's decisions as recommended"; section 7). Docs
only, nothing implemented. It depends on CR-025 and on jsonix-CR-015 (D2; written 2026-11-05 at `34d2572`, its U1 to U3 Jason's). Word check 76 ran on 2026-11-05 and 2026-11-06, in five rounds and one further file (section 4). Word 365 dropped what it didn't know or refused the file, which supports D1 and D3, but for what an unknown `w:` element holds in a content host, which it reads (section 4.4). D7, decided 2026-11-06, reads it so in the fifteen hosts measured (Jason: "CR-026 D7: read as Word does"; section 3.7). Check 76 also showed CR-021's rule 2 reading words Word doesn't, fixed the same day (section 4.2, D6).

## 0. Summary

The finding was narrower than the defect. An element the model doesn't know fails the whole part, not only in `w:sdtPr` but in 20
of the 32 hosts tried, among them `w:body`, `w:p`, `w:r`, `w:tbl`, `w:tr`, `w:tc`, `w:hdr` and `w:comment`. In the main part that
means nothing can read the document. It makes no difference whether the element's namespace is in `mc:Ignorable`, the one case
markup compatibility says a consumer must get past. docx4j reads such a document, and so does python by default.

The decision (D2) is to skip the element and report it at the point the runtime throws today, a small jsonix change, which a
prototype shows reads all 52 cases with no change to any reader. The element is then dropped, as docx4j and python drop it, so
this has to land after CR-025: reading it without CR-025 would turn a loud refusal into a silent loss. The route doesn't reach every
host: section 3.4 says which, and this CR's title holds only if jsonix's U2 is yes.

D7 changes that for an unknown `w:` element in a content host. Word 365 reads what such an element holds in all fifteen hosts measured, so a DOM pass
takes away the tags and what they held reads as if it stood there (section 3.7). Everywhere else the element is dropped as above.

## 1. Here today (measured 2026-11-05, `b07134f`, objects 0.7.0, `@docx4j/jsonix` 3.11.0)

Two elements in each of 32 hosts: `ext:opaque`, its namespace in `mc:Ignorable`, and `w:madeUp`, an unknown name in the `w:`
namespace. Both behaved alike in every host.

| The read | Hosts |
|---|---|
| **Throws** (20) | `w:body`, `w:p`, `w:r`, `w:hyperlink`, `w:fldSimple`, `w:ins`, `w:smartTag`, `w:customXml`, `w:tbl`, `w:tr`, `w:trPr`, `w:tc`, `w:sdtPr` (block and run), `w:sdtContent` (block and run), `w:ffData`, `w:hdr`, `w:comment`, `w:footnote` |
| Drops it, silently (12; CR-025) | `w:document`, `w:pPr`, `w:rPr`, `w:tblPr`, `w:tblGrid`, `w:tcPr`, `w:sdt`, `w:sdtEndPr`, `w:sectPr`, `w:style`, `w:lvl`, `w:settings` |

Through the real load path, the 26 hosts under `w:document` as 52 files: 34 throw `PartUnmarshalException` and 18 read. The numbers
are the same under `mcePreprocess: true`, and an untouched save stays byte for byte.

**More hosts,** tried after jsonix-CR-015 raised them, the same two elements through the real load path:

| Host | Here | docx4j 17.3.5 |
|---|---|---|
| `w:b`, `w:br`, `w:jc`, `w:tab` | Throws, with a second message: "Illegal state: must be END_ELEMENT" | Reads, the element dropped |
| `w:t`, `w:instrText` | Reads: CR-021's pass gets there first (section 3.4) | Reads, the element dropped |

So 20 of 32 is a count of the hosts tried, not of what fails. jsonix counts 852 classes on that second path, 240 of wml's 385.

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
repair offered), and opened the declared twin. Word 365 opened the undeclared file with no message, but it knows `w15`. ECMA-376 Part 3 has a consumer treat such markup as an error unless its
namespace is ignorable. Check 76 measured Word 365 (sections 4.1 and 4.3): it refused an element in such a namespace in all seven hosts
tried, and one in no namespace, and read past an attribute in such a namespace and an unknown `w:` name.

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

### 3.1 The runtime (jsonix-CR-015, written 2026-11-05 at `34d2572`, its U1 to U3 Jason's)

In `Jsonix.Binding.Unmarshalls.Element.unmarshalElement`, where an element is neither typed nor allowed as DOM: if the caller set
`onUnexpectedElement`, call it and skip the element; otherwise throw as today, and always for a document's root. A caller with no
callback sees no change, so the facade's `unmarshalString` still refuses. A caller with one gets a report where it got a throw (its
U1), so raising the floor changes nothing here until CR-025 passes the callback, which is this CR's order.

### 3.2 This package

With CR-025's callback in place, the read needs no code of its own: a skipped element is a `part.skipped` warning, and the save
copies, refuses or repairs as CR-025 says. What changes is that a main part which failed on this cause now reads. Every other cause
of `PartUnmarshalException` stays, and CR-012 section 25's rule that a main part which can't be read is refused stands.

`insertOoxml` needs tracing in phase 2: a caller's own fragment holding an unknown element is arguably better refused than read
with a warning on a package the caller never sees. The same goes for CR-025's D9. The fragment route runs CR-021's pass and discards its
report (`ooxml.mts`), so what the pass removes from a fragment's `w:t` reaches no warning. Python refuses such a fragment (its CR-030 section 11).

### 3.3 Reading without the element

A dropped element can hold text (`<ext:opaque>words</ext:opaque>`), which then doesn't read. For an ignorable namespace that is what
markup compatibility asks, and what Word does: `76k` has a run in such an element, and Word neither shows nor saves it. For a foreign
namespace that isn't ignorable there's no Word reading to match, since Word refuses the file (`76e`). An unknown `w:` element is
different: Word reads what it holds in a content host (`76t`, section 4.4), and so will this package (section 3.7, D7). The WD has no statement on any of these (D4).

### 3.4 What route B doesn't reach

jsonix-CR-015 changes the branch that throws for the 110 classes holding lists. Two other paths remain, by its count:
- **A class with no element or value content** (`w:b`, `w:br`, `w:jc`, `w:tab`; 852 classes): a different throw. Its U2 asks whether
  that CR takes them too, and recommends yes, its prototype reporting them all. If the answer is no, these hosts go on failing their
  part, and this CR's title claims too much. Word reads past them: `76l` opened with no message, the element dropped, the text bold.
- **A value class** (23, with 183 simple-typed element properties): its U3 recommends leaving these. Word's reading is measured only
  at `w:t`, where nobody needs the runtime's skip, and elsewhere it's unmeasured, with a non-string value's `1<ext:o/>2` reading as 12.
  `w:t`, `w:instrText`, `w:delText` and `w:delInstrText` are among them, but here they never reach it, because CR-021's pass
  acts first and splits the text around the element. It reports `br.outside_run`, and under CR-025's D9, decided, it will report `part.skipped`
  too. CR-021's rule gave up an element's text, `<ext:opaque>WORDS</ext:opaque>` in a `w:t` reading as `WORDS`, until `76n` showed Word
  dropping it with the element (section 4.2, D6). `76m`, `76n` and `76u` also show Word skipping the element and joining the text around it, which is the reading its U3 (b) would give at `w:t`. The other value
  classes go on refusing, and Word's reading of them is unmeasured.

### 3.5 Not here

- Keeping the element (D1).
- `mc:MustUnderstand` and `mc:ProcessContent`: nothing under `src/` reads them. An element with `mc:ProcessContent` should have its
  children read in its place, which skipping doesn't do. Check 76 doesn't cover them.
- Word-matching refusal of what Word refuses: the repair policy's tier 3 is read so that a caller can fix it. `76e` is such a file,
  and CR-025 section 2.4 says how.

### 3.6 CR-023

The suite's four content-control scenarios then load. Their two preservation rows go on failing unless D1 keeps the element, as
docx4j's do.

### 3.7 Reading through (D7)

An unknown `w:` element in a content host loses its tags and keeps what it holds, as Word 365 reads it (section 4.4). A DOM pass does it, in
`XmlPart.preprocess` beside CR-021's, before the runtime sees the part, so it asks nothing of jsonix.

- **What it takes:** an element in the `w:` namespace whose name the model doesn't hold, in one of the fifteen measured hosts, at any depth (`76t7`). What it held takes its place.
- **The hosts:** `w:p`, `w:body`, `w:tc`, `w:r`, `w:tbl`, `w:tr`, `w:sdtContent`, `w:hyperlink`, `w:ins`, `w:hdr`, `w:comment`, `w:fldSimple`, `w:smartTag`, `w:customXml` and `w:footnote`.
- **The model's names:** they come from the mapping, and phase 2 has to derive them properly. Section 4.4's rough walk isn't complete.
- **The report:** one `part.skipped` for each element taken, so CR-025's 2.2 copies a part read and not changed, tags and all.
- **The gate:** a changed part holding such tags is refused until `checkLossOnSave` is off, though no words go. What a re-marshal loses is the tags, and Word's own save drops them with no message. This is wider than CR-025's D9, which lists only what goes with its content. D9's first reason doesn't apply here, since Word opens these files, and none of the 720 packages holds such an element, so the wider report stands as D7 says.
- **Everything else stays on route B:** `w:pPr`, `w:rPr` and `w:t` are measured, and drop as route B does. A host not measured (`w:ftr`, `w:endnote`, `w:del`, `w:sdtPr`, `w:trPr`, `w:tcPr`, `w:tblPr`, `w:sectPr`) drops the element with all it holds until a check says otherwise, and adding one to the list is then the whole change.
- **What its host doesn't admit:** unmeasured (section 4.4). The pass doesn't look, so a `w:p` left in a `w:p` is CR-016's misplaced child, reported on read and refused on save.
- **`insertOoxml`:** whether the pass runs on the fragment route goes with section 3.2's tracing.

This departs from docx4j and python, which drop the element with all it holds. Both are to be told on commit: the measurements are theirs to use.

## 4. Word check 76 (phase 0, run 2026-11-05 and 2026-11-06)

Word 365 only. `test/fixtures/check76/build.mjs` builds one small file per case, each with marked paragraphs before and after, on
the share in `__tmp/76/` and not under `fixtures/`, whose gate they'd fail until this CR is built (`test/README.md` item 76):

| File | Case |
|---|---|
| 76-control | Nothing unknown, the control for the writer |
| 76a to 76d | `ext:opaque`, `ext` in `mc:Ignorable`, in `w:sdtPr`, `w:p`, `w:r`, `w:tc` |
| 76e | `ext:opaque` in `w:p`, `ext` declared, not ignorable |
| 76f, 76g | `w:madeUp` in `w:p`, in `w:pPr` |
| 76h, 76i | Attribute `ext:note` on `w:p`, ignorable, not ignorable |
| 76j | Attribute `w:madeUp` on `w:p` |
| 76k | `ext:opaque` in `w:p`, ignorable, holding a run reading `HIDDEN WORDS`: do the words show? |
| 76l | `ext:opaque` in a `w:b`, ignorable: a class with no element content (section 3.4) |
| 76m | An empty `ext:opaque` inside a `w:t`, ignorable |
| 76n | `ext:opaque` inside a `w:t`, ignorable, holding the text `HIDDEN WORDS`: do the words show? Here they read as text |

Each open records any prompt (the policy's tier), what shows, and, Save As first, what Word's save keeps. That answers D1 and D3
and gives each shape its tier. It blocks neither D2 nor the build: B reads in every tier.

### 4.1 What Word 365 did

Jason's run of 2026-11-05: "check 76 all opened ok except 76e which word could not repair". Word's saves and PDFs are in
`test/check76/`, with a SHA-1 manifest. "Here" is the build they were compared with, before this CR and D6's fix.

| File | Word's open | Word's save | Here |
|---|---|---|---|
| 76-control | No message | All three paragraphs | Reads |
| 76a | No message | The control with its `w:alias` and `w:id`, no `ext:opaque` | Fails the part |
| 76b to 76d | No message | No `ext:opaque` | Fails the part |
| 76e | **Could not repair it** | None | Fails the part |
| 76f | No message | No `w:madeUp` | Fails the part |
| 76g | No message | No `w:madeUp` | Reads, dropped, nothing said |
| 76h to 76j | No message | No attribute | Reads, dropped, nothing said |
| 76k | No message | One run, `TARGET TEXT, then  and the rest.`; the element and its run gone | Fails the part |
| 76l | No message | `<w:b/>`, the text bold | Fails the part |
| 76m | No message | One `w:t`, `TARGET TEXT` | `TARGET TEXT`; `br.outside_run` |
| 76n | No message | One `w:t`, `TARGET TEXT`, the words gone | **`TARGET HIDDEN WORDSTEXT`**; `br.outside_run` |

The namespace's declaration and `mc:Ignorable` entry are gone from every save, and nothing of `ext` is left in any part. The PDFs
agree with the saves. All 14 saves read here with no warning, each middle paragraph Word's.

What it settles:
- **Word drops, so D1 stands, but for `76t` (section 4.4).** Everything Word opened lost what Word didn't know, with no message. Keeping it here would keep it only
  until the next Word save, so nothing in check 76 argues for route A. CR-023's preservation rows go on failing, as section 3.6 says.
- **Tiers, so D3 stands.** Everything is tier 1 but `76e`, tier 3, and round 2 adds to tier 3 (section 4.3). Nothing is tier 2: Word repaired nothing.
- **`ignorable` doesn't say what Word does.** `76f`, `76g`, `76i` and `76j` are all `ignorable: false` and opened with no message. What
  Word refused was an element, not an attribute, in a namespace it doesn't know, not `w:`. A caller can't read `ignorable: false` as
  "Word refuses this".
- **`76k` gives D4's question Word's answer** for an ignorable element: dropping it drops nothing Word shows.
- **`76l` answers U2's one objection** (section 3.4), and **`76m` and `76n` bear on U3.**

### 4.2 `76n` and CR-021's rule 2

Rule 2 gave every element in a `w:t` but a `w:br` its text. That was measured for a nested `w:t`, which Word knows. `76n` showed it
wrong for an element Word doesn't know: Word reads `TARGET TEXT`, and here it read `TARGET HIDDEN WORDSTEXT`, under a warning saying
"read as Word reads them". Before CR-021 the part failed, so the wrong reading was new with it and never released. The rule is
docx4j's (`mc-preprocessor.xslt`'s `*` template in `split-text`, `@since 17.3.6`, unreleased) and python's (`_split` in `breaks.py`),
so both should read the same way; that is from reading them, not running them.

Fixed 2026-11-05 (D6; Jason: "we should fix the bug (76n)"; CR-021 section 11). This section first proposed "an element outside `w:`".
The fix was narrower, an element in a namespace Word doesn't know, since that was all `76n` measured: a `br` in no namespace is the
shape of a real file, docx4j's survey sample `01af91808eac`. Round 2 then took it as far as the measurements go: `76u` showed the same bug
for an unknown `w:` name. `textElements` counts the element, and CR-025's D9 will name it.

### 4.3 Round 2 (run 2026-11-05)

Jason: "let's do the second round of files", then "word could not repair p q r s v x y. for others see -word365.docx and .pdf in the
dir" and "76o is a could not repair". The files are `76o` to `76z`, on the share beside round 1's, which rebuilt byte for byte, with Word's saves and PDFs in
`test/check76/`. "Here" is the build with D6's fix.

| File | Case | Word's open and save | Here |
|---|---|---|---|
| 76o | `76e`'s `ext:opaque`, not ignorable, in `w:sdtPr` | Could not repair it | Fails the part |
| 76p | The same in `w:pPr` | Could not repair it | Reads, dropped, nothing said |
| 76q to 76s | The same in `w:r`, `w:tc`, `w:b` | Could not repair it | Fails the part |
| 76t | `w:madeUp` in `w:p`, holding a run reading `HIDDEN WORDS` | No message; **the run read**: one run, `TARGET TEXT, then HIDDEN WORDS and the rest.`, the element gone | Fails the part |
| 76u | `w:madeUp` inside a `w:t`, holding `HIDDEN WORDS` | No message; `TARGET TEXT`, the words gone | `TARGET TEXT`, since section 4.2's fix |
| 76v | `ext:opaque`, not ignorable, inside a `w:t`, holding `HIDDEN WORDS` | Could not repair it | `TARGET TEXT` |
| 76w | `ext:opaque`, ignorable, inside a `w:t`, holding a `w:br` | No message; `TARGET TEXT`, no break | The same |
| 76x | `span` in no namespace inside a `w:t`, holding `HIDDEN WORDS` | Could not repair it | `TARGET HIDDEN WORDSTEXT` |
| 76y | `br` in no namespace inside a `w:t`, empty | Could not repair it | `TARGET TEXT`, no break |
| 76z | `w:tab` inside a `w:t` | No message; `TARGETTEXT`, no tab | The same |

What it settles:
- **The refusal is a rule.** A foreign element that isn't ignorable failed in all seven hosts tried (`76e`, `76o` to `76s`, `76v`), `w:sdtPr` among them as in Word 2010, and so did
  an element in no namespace (`76x`, `76y`). That's tier 3, read here so that a caller can fix it (D3). `76y` is the shape of docx4j's
  sample `01af91808eac`, so if that file is as its CR-073 describes, Word refuses it too.
- **`76p` reads here with nothing said,** like the other 11 hosts that drop: a file Word refuses, which CR-025's `part.skipped` will report.
- **Rule 2 is finished.** Word honours a `w:br` and a nested `w:t` in a `w:t`, and drops any other element it reads there with all it
  holds. `76w` and `76z` already read so here, and `w:tab` isn't a second bug: Word draws no tab either. CR-021 section 11 has the table.
- **`76t` reopens D1** for an unknown `w:` element.

### 4.4 Word reads what an unknown `w:` element holds, in a content host

Round 1's rule was that Word drops what it doesn't know with all it holds. `76k` (ignorable) follows it. `76t` doesn't: `w:madeUp` in a
`w:p` went, and the run inside it stayed, read and saved as if the tags weren't there. A third round asked where else (Jason: "commit,
then let's do the third round of files"; "run of 76t1 to 76t7 done"). All seven opened with no message. A fourth took the hosts the third
left (Jason: "CR-026 D7 we need to measure w:tr, w:sdtContent, w:hyperlink, w:ins, w:hdr and w:comment before making any decision"; "anyway, round 4 now done :-)"), and all ten opened with no message. Each host is valid with or
without the element, so only it is in question. Word's saves and PDFs are in `test/check76/`.

| File | `w:madeUp` holding, in | Word | Here |
|---|---|---|---|
| 76t | A run, in a `w:p` | **Read:** `TARGET TEXT, then HIDDEN WORDS and the rest.`, one run | Fails the part |
| 76t1 | A `w:p`, in `w:body` | **Read:** a paragraph `HIDDEN WORDS` | Fails the part |
| 76t2 | A `w:p`, in a `w:tc` | **Read:** the cell's first paragraph | Fails the part |
| 76t3 | A `w:t`, in a `w:r` | **Read:** one run, as `76t` | Fails the part |
| 76t4 | A `w:tr`, in a `w:tbl` | **Read:** the table's second row | Fails the part |
| 76t7 | A `w:madeUp` holding a run, in a `w:p` | **Read:** as `76t`, at any depth | Fails the part |
| 76t8 | A run, in a `w:p`, no `mc:Ignorable` on the root | **Read:** the body as `76t`'s | Fails the part |
| 76t9 | A `w:tc`, in a `w:tr` | **Read:** the row's second cell | Fails the part |
| 76t10 | A `w:p`, in a block control's `w:sdtContent` | **Read:** the control's second paragraph | Fails the part |
| 76t11 | A run, in a run-level control's | **Read:** one run, in the control | Fails the part |
| 76t12 | A `w:tr`, in a row-level control's | **Read:** the second row, in the control | Fails the part |
| 76t13 | A `w:tc`, in a cell-level control's | **Read:** the second cell, in the control | Fails the part |
| 76t14 | A run, in a `w:hyperlink` | **Read:** one run, in the link | Fails the part |
| 76t15 | A run, in a `w:ins` | **Read:** one run, tracked as inserted | Fails the part |
| 76t16 | A `w:p`, in a `w:hdr` | **Read:** the header's second paragraph | Fails the header part |
| 76t17 | A `w:p`, in a `w:comment` | **Read:** the comment's second paragraph | Fails the comments part |
| 76t5 | A `w:jc`, centred, in a `w:pPr` | **Gone:** no `w:jc`, not centred | Reads, dropped, nothing said |
| 76t6 | A `w:b`, in a `w:rPr` | **Gone:** no `w:b`, not bold | Reads, dropped, nothing said |
| 76u | Text, in a `w:t` | **Gone** (section 4.3) | `TARGET TEXT` |

So the host decides. Where it holds content the tags go and what they held stays, at any depth. Where it holds properties or text the
element goes with all it holds. With round 5, below, that's fifteen hosts and three, and every save from round 3 on reads here with no warning and Word's text.

The last three read here as in Word. The others read wrong nowhere today, since the part fails, but route B as decided skips the
element with all it holds: `76t1` would lose a paragraph and `76t4` a row, with `part.skipped` naming the element and not saying that
Word shows what it held. None of the 720 packages holds an unknown element, so this wasn't urgent, but it had to be settled before phase 2
built the drop. D7 takes (b).

Other readers don't settle it. docx-tool-identifier ran the files through 89 local reader configurations on 2026-11-05 (its
`runs-local/check76/`; nothing entered its library). On `76t`, 25 keep the words, 55 lose them, 5 fail, core-ts among them, and 4 give no
main text. Every reader gives `76t` and `76k` the same result, so Word's split between an unknown `w:` element and an ignorable one is
Word's alone, and neither (a) nor (b) is an outlier.

- **(a) Drop and report, as first decided.** Simple, and the gate still protects a save. Every reader lacks what Word shows.
- **(b) Read it as Word does, decided (D7, section 3.7).** A DOM pass ahead of the runtime, beside CR-021's, takes away the tags of an unknown `w:` element in a
  content host and reports them as `part.skipped`, what they held then reading as if it stood there. Route B stays for everything else,
  and in `w:pPr` and `w:rPr` its skip already does what Word does. It asks nothing of jsonix. It needs two lists:
  - the model's `w:` names, which the mapping holds: a rough walk of `org_docx4j_wml` found `w:p`, `w:jc` and `w:tr` and not `w:madeUp`,
    though it isn't complete, and phase 2 has to derive the list properly;
  - the content hosts. Fifteen are measured, every content host in section 1's table. No file holds `w:ftr`, `w:endnote` or `w:del`. The property hosts `w:sdtPr`, `w:trPr`, `w:tcPr`,
    `w:tblPr` and `w:sectPr` aren't measured either. Section 1's table splits nearly the same way, the model refusing in the fifteen and
    dropping in `w:pPr` and `w:rPr`, but `w:sdtPr` and `w:trPr` refuse there, so it can't stand in for a measurement.

Not covered by any round: an element of a namespace Word knows (`m:`, `w14:`, `mc:`) in a `w:t`; a known `w:` name holding text there,
which CR-021 section 11 treats as unknown; `w:madeUp` holding what its host doesn't admit, a `w:p` in a `w:p` say, since every file above
holds what its host admits, and under (b) the pass would leave a misplaced child in the tree (CR-016).

Every other `w:madeUp` file has `mc:Ignorable="ext"` on its root, which doesn't name `w:`, and the foreign elements Word refused have none,
so `76t8` holds `76t`'s content with no `mc:Ignorable` (Jason: "do That one extra file"; "76t8 done"). It opened with no message, and
Word's save has `76t`'s body exactly, `HIDDEN WORDS` in one run. So the attribute plays no part in reading through: with none on the root,
Word still reads `76t8`'s `w:` element and refuses `76e`'s foreign one, though `76e`'s is empty, so the two differ in what they hold as well
as their namespace. That's measured in a `w:p`, and round 4's hosts have none either. Round 3's other hosts each had the attribute, and no file repeats them without it.

Round 4 had a control. `76t18` holds every host of `76t9` to `76t17` with nothing unknown and `HIDDEN WORDS` standing bare, because `76t16` and
`76t17` are the first check 76 files with a part of their own. It opened with no message, and each of the nine saves matches its host in `76t18`'s
save exactly, so the element changed nothing but itself. `w:sdtContent` has four files because block, run, row and cell are four types in the
model, and all four read through. The words stayed in their host: in the control, in the link, tracked as inserted. Word saved the row-level and
cell-level controls wrapped round the whole table, in `76t18` too, so that isn't the element's doing. `76t17`'s PDF shows no comment; its save
has both paragraphs.

**Round 5** (Jason: "measure w:fldSimple, w:smartTag, w:customXml and w:footnote"; "round 5 done"), the content hosts of section 1's table that round
4 left. Eight files, none with `mc:Ignorable`. Saves and PDFs are in `test/check76/`, and all eight saves read here with no warning and Word's text.

| File | `w:madeUp` holding, in | Word | Here |
|---|---|---|---|
| 76t19 | A run, in a `w:fldSimple` | **Read:** one run, the field's result, the field saved as a complex field | Fails the part |
| 76t20 | A run, in a `w:smartTag` | **Read:** one run, in the `w:smartTag`, which stays | Fails the part |
| 76t21 | A `w:p`, in a block `w:customXml` | **Read:** a second paragraph, the `w:customXml` gone | Fails the part |
| 76t22 | A run, in a run-level `w:customXml` | **Read:** one run, the `w:customXml` gone | Fails the part |
| 76t23 | A `w:tr`, in a row-level `w:customXml` | **Read:** a second row, the `w:customXml` gone | Fails the part |
| 76t24 | A `w:tc`, in a cell-level `w:customXml` | **Read:** a second cell, the `w:customXml` gone | Fails the part |
| 76t25 | A `w:p`, in a `w:footnote` | **Read:** the note's second paragraph | Fails the footnotes part |
| 76t26 | Nothing: the control, every host above with `HIDDEN WORDS` standing bare | **Opened,** run on 2026-11-06 (Jason: "76t26 is done"): the field saved as a complex field, the `w:smartTag` kept, all four `w:customXml` gone | Reads, the words in all seven |

`w:customXml` has four files for the reason `w:sdtContent` has. Word gave a message on `76t24` (`76t24-customxml-message.png`): "This file has custom XML
elements that are no longer supported in Word. Saving the file will remove these elements permanently." It's the only message in check 76, an OK box
and no repair, and it names `w:customXml`, not `w:madeUp`. Which other files showed it isn't recorded.

`76t26` shows the element changed nothing but itself: `76t19` to `76t24`'s saves each match their host in `76t26`'s save exactly, and `76t25`'s footnotes part
matches `76t26`'s. So `76t19`'s complex field is Word's way with a `w:fldSimple`, and a `w:customXml` with nothing unknown in it loses its tags too, as the
message's words suggest. Jason's report doesn't say whether the message showed on `76t26`.

## 5. Tests and oracles

- `test/unknown-elements.test.mjs`: section 1's hosts, both elements, each read, its warning's fields, `Paragraph.text`,
  `Texts.of` and a save; a module beyond WordprocessingML (a DrawingML and a SpreadsheetML host), since the prototype tested neither.
- jsonix-CR-015 section 8's run, on its unreleased build before its release: the 52 files and the DrawingML and SpreadsheetML hosts,
  the callback set and unset.
- docx4j's `UnexpectedElementInSdtPrTest` mirrored: inline and block, ignorable and not.
- The docx4j oracle, once `VERSION_17_3_6` is released: the 52 files, every host read, every element dropped, but for an unknown `w:` element in section 3.7's hosts, which this package reads through (D7). That difference is listed, not fixed.
- A caller with no callback still gets the runtime's refusal. No test here asserts today's message, so none has to change.
- Check 76's saves under `test/check76/` (copied 2026-11-05 and 2026-11-06, with a SHA-1 manifest, which `br-placement.test.mjs` checks, holding `76m`, `76n`, `76u`, `76w` and `76z` to Word's text). With the build: each save read with no warning and its text Word's, each source file from `check76Files` held to section 4.1, and section 4.4's read through to Word's text, each with its `part.skipped`.

## 6. Phases and size

| Phase | What | Size | Waits on |
|---|---|---|---|
| 0 | Check 76: the files, Jason's five rounds and `76t8` of 2026-11-05, and `76t26`, the last control, of 2026-11-06 | Done | |
| 1 | jsonix-CR-015 (written) decided and released; the floor raised | ½ day there | Its U1 to U3 |
| 2 | The read, section 3.7's pass and the model's `w:` names, `insertOoxml` traced, the tests | 1 day | CR-025 phases 1 and 2; phase 1 |
| 3 | The records, CR-023's line | ¼ day | |

## 7. Decisions (D1 to D5 taken 2026-11-05, all as recommended; D6 raised by check 76 and taken the same day; D7 raised by check 76 and taken 2026-11-06, as recommended)

| | Question | Recommendation |
|---|---|---|
| D1 | Keep the element, or drop it? `benchmarking.md` holds this open across the ports. | Drop, as docx4j and python do, and revisit with check 76. Keeping (route A) costs a reader audit and still can't cover 18 of the 52 cases. Check 76: Word drops it too, so this stands, but for what an unknown `w:` element holds (D7) |
| D2 | Route A, B or C? | B |
| D3 | Read an element whose namespace isn't ignorable, or an unknown `w:` name, or refuse it? | Read and report (`ignorable: false`): docx4j and python read it, and the policy's tier 3 reads so that a caller can fix it. Check 76 gives the tier: 3 for a foreign element that isn't ignorable and one in no namespace, in every host tried, and 1 for the rest. This stands |
| D4 | Should the text model warn of an element dropped from a paragraph? | A question for the WD, sent on commit; nothing here until it rules. It goes with `76k`'s answer: Word shows nothing an ignorable element holds |
| D5 | Which release? | With CR-025, after 0.7.0 |
| D6 | Raised by check 76's `76n`: change CR-021's rule 2 so that an element Word doesn't know goes with all it holds, as Word does (section 4.2)? In which release? | Yes, before 0.7.0, CR-021's first release (Jason: "we should fix the bug (76n)"). Fixed the same day, and for an unknown `w:` name after round 2's `76u` (CR-021 section 11). docx4j and python to be told, the rule being theirs too |
| D7 | Raised by `76t`: an unknown `w:` element in a content host, whose content Word shows. Drop and report as decided, or read it as Word does (section 4.4)? | Read it as Word does. Five rounds give a clear rule: fifteen content hosts read through, three of properties or text don't, and `mc:Ignorable` plays no part where tested. A DOM pass in phase 2 over the measured hosts, the rest left to route B until they're measured. Round 4 was Jason's condition for deciding ("CR-026 D7 we need to measure w:tr, w:sdtContent, w:hyperlink, w:ins, w:hdr and w:comment before making any decision") and is back. So is round 5, asked for after it (Jason: "measure w:fldSimple, w:smartTag, w:customXml and w:footnote"), its control run on 2026-11-06 (Jason: "76t26 is done"). Decided 2026-11-06 as recommended (Jason: "CR-026 D7: read as Word does"). Section 3.7 is the pass. A departure from docx4j and python, to be told on commit |

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
