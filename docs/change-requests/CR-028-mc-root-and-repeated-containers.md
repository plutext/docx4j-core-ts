# CR-028: `mc:AlternateContent` as the root, and a repeated container, read as Word 365 reads them

**Status:** Proposed 2026-11-06, at Jason's "Please do write the Follow-up CR!", on his "A follow-up CR after 0.7.0, covering repeated containers and the mc root which we may defer indefinitely". Measured here the same day on `26c2c4f` (section 2), and both passes tried in scratch through the load hook (section 3.4). Nothing built. docx4j's sibling is its CR-079, proposed and uncommitted when this was read, at `f960df59e`. This CR keeps its case names and shares its phase 0. Another story's root in the main part measured 2026-11-07, at Jason's "what does Word do if the root element of the main document part is that of some other story", and added at his "Add these results to CR-028" (sections 1.1 and 2.4, D11). The same day's rounds added at his "add the rounds to CR-028, then commit": three copies, nesting, and what decides an unknown root (sections 1.2 to 1.4 and 2.5, D12), with WD 0.97's rules. Word's merge rule added the same day from `sm-*`, the round at Jason's "order a Word round to identify the real merge rule", read by python and again here (sections 1.5 and 2.6), which corrects this CR's sentences on `w:sectPr`, section 3.2 and D2 and D6.

## 0. Summary

Word opens both shapes silently, and this package reads neither as Word does.

- **A container given more than once.** Word treats it as one: `w:body` and `w:sdtContent` join in order, `w:txbxContent` joins in reverse order, and `w:pPr`, `w:rPr`, `w:sectPr` and `w:tblPr` merge, attribute by attribute: what a later copy states wins, and what it doesn't state stays (section 1.5). Three copies follow the same rules as two. The runtime keeps the last whole and says nothing, so of the 13 cases first measured with two copies this package matches Word in 8 and reads less in 5. In 3 of those, words go, and a save after a read writes them out of the file with nothing said. Of the merge round's 40 files it matches Word in 6, nearly all the rest differing in properties, not words (section 2.6).
- **`mc:AlternateContent` as the root of `word/document.xml`.** Word reads the branch it takes, joins two `w:document` in it, and saves an ordinary `w:document`. Here the part holds an `AlternateContent`, every reader sees an empty document, and a save writes the source's words back.

The root is worse than an empty read in two ways, neither known when CR-025 section 2.6 was written. **A write through `Body` is lost with nothing said:** the view reads it back, and the save leaves it out. And `mcePreprocess: true` fails the load with a raw DOM error.

A DOM pass ahead of CR-026's reads the root as Word does, and a second pass read the first five containers. A scratch trial of the two matched Word's saves in words on all 20 files tried, and in properties on 18, and kept every write. Word's rule has since proved wider than that second pass, so D2 now recommends the objects package's CR-022 for it. Neither shape occurs in the packages read here, and Word's saves of these files hold neither, so nothing presses this. It waits for 0.7.0 and may wait indefinitely (D1). If it does, D10 is the part worth doing alone.

Another root in the main part was measured afterwards (sections 1.1 and 1.4). Word has no one answer. It opens `w:hdr` and `w:ftr` as a blank page, refuses `w:footnotes` and `w:endnotes`, repairs `w:comments`, and treats a `w:` root it doesn't know as if its tags were absent, so what that root holds decides. A nested `w:body` reads through and a nested `w:document` is refused (section 1.3). Here most of these roots read empty and lose a write, as the `mc` root does, and an unknown one fails at the read. D10 as worded covers the write, and D11 follows WD 0.97's rules.

## 1. What Word does

Word 16.0.20326.20232, through the bridge, each file twice, the two runs equal apart from generated ids, every file silent but `rw-22`. The objects session built the files and read the saves (`rc-*`, `rw-*` in the share's `bridge/portfolio/`; the portfolio's `text_extraction_behaviour_model.md` section 7.15). docx4j read the rows again from Word's saves, and so did this session.

| Given twice | Word |
|---|---|
| `w:body` (`rc-10`) | Joined in order. The closing `w:sectPr` copies make one section, by section 1.5's rule, with no section break |
| `w:sdtContent`, block level (`rc-11`) | Joined in order |
| `w:txbxContent`, in a VML box (`rc-12`) | Joined in reverse order, the second first (section 1.2) |
| `w:pPr`, `w:rPr` (`rc-20` to `rc-23`) | Merged, the later value winning: `w:jc` then `w:ind` keeps both, `w:b` then `w:i` keeps both |
| `w:jc`, `w:b`, `w:sz` (`rc-30` to `rc-33`) | The last |
| `w:sectPr`, `w:tblPr` (`rc-40`, `rc-41`) | The later value. Each pair stated the same properties, so these two couldn't show a merge. Section 1.5 does |

| `word/document.xml`'s root is `mc:AlternateContent` with | Word |
|---|---|
| A `Choice` it understands (`rw-20`) | The Choice. The save has an ordinary `w:document` |
| A `Choice` it doesn't, and a `Fallback` (`rw-21`) | The Fallback |
| Two `w:document` in the branch taken (`rw-23`, `rw-26`, `rw-27`) | Both, joined in order. The sections merge by section 1.5's rule (`sm-52`), so the second's values stand, margins included |
| A `Choice` it doesn't understand, and no `Fallback` (`rw-24`) | An empty document |
| A `Fallback` and no `Choice` (`rw-22`) | Repaired, keeping the Fallback's words. No in-body control, so the repair may come from the missing `Choice` rather than the root |

### 1.1 Another root in `word/document.xml`

Word 16.0.20326.20232, through the bridge, 2026-11-07, in the same three-part package, each file twice, the two runs agreeing (`or-*` in the share's `bridge/portfolio/`). This session built the files and read the saves, and the portfolio's section 7.15 holds the table too. `or-14-comments-a` was killed at a dialog the bridge didn't yet know, and `-a2` is its rerun, on bridge `973e595`. The control, `or-00`, opens silently with its words in the body.

| Root, under the main document's content type | Word | Word's save |
|---|---|---|
| `w:hdr`, `w:ftr` (`or-10`, `or-11`) | Opens silently, a blank page | An empty body. The words are the first paragraph of the `separator` footnote (id -1), followed by an empty `Footer`-styled paragraph and an empty paragraph, and nothing refers to a footnote. The file holds no footnote, and its text and PDF are empty. Whether the words draw once one exists isn't measured |
| `w:footnotes`, with its separators and without (`or-12`, `or-15`); `w:endnotes`, with them (`or-13`) | Refused: "Word experienced an error trying to open the file" | None |
| `w:comments` (`or-14`) | Repair offered. After Yes, Show Repairs opens with its list empty | One paragraph holding the comment's range and reference, and the comment in `word/comments.xml` |
| `w:styles` (`or-30`) | Opens silently, empty | An empty body |
| `w:body` with no `w:document` (`or-32`) | Opens silently | The words in the body |
| `w:madeUp` holding a `w:p` (`or-31`) | Opens silently | The words in the body |
| `mc:AlternateContent`, `w:hdr` in a `w14` Choice, `w:document` in the Fallback (`or-50`) | As `or-10` | As `or-10`, and the Fallback's words are nowhere in the save |

Under a header's content type, `w:hdr` (`or-40`) and `w:document` (`or-41`) alike are offered repair and, after Yes, refused. So there the root makes no difference.

### 1.2 Three copies

Jason asked whether anything different happens with triplicates. No earlier file held three in one place: `rw-23`, `rw-26` and `rw-27` count three `w:document` only across both branches. So `rt-*`, 2026-11-07, 20 files in `rc-*`'s package and frames, each twice, every file silent and the two runs agreeing. The rules are the same as with two copies, and the third settles two things.

| Three copies of | Word |
|---|---|
| `w:body` (`rt-10`), block `w:sdtContent` (`rt-11`) | Joined in order |
| `w:document` in the branch taken (`rt-50`, `rt-51`) | Joined in order, the sections merged by section 1.5's rule |
| `w:txbxContent`, in a VML box (`rt-12`, `rt-13`) | **Reverse order:** third to first, and with four copies fourth to first. So `rc-12`'s "second first" was this, still in the one shape |
| `w:pPr`, `w:rPr` (`rt-20` to `rt-23`) | Merged, the later value winning: `w:jc`, `w:ind` and `w:spacing` all kept; `w:b`, `w:i` and `w:u` all kept |
| `w:jc`, `w:b`, `w:sz` (`rt-30` to `rt-33`) | The last: `w:b` on, off, on is on, and off, on, off is off |
| `w:sectPr` (`rt-40`) | Merged: the third's page size, with the `w:orient` only the second states (section 1.5) |
| `w:tblPr` (`rt-41`) | The last `w:jc` |

**Which section stands.** The closing `w:sectPr` copies merge in order by section 1.5's rule, whichever body each is in: with the second body's missing the third's values stand (`rt-16`), and with the third's missing the second's do (`rt-17`). A `w:sectPr` in a `w:pPr` stays as a section break, and the closing ones make one section after it (`rt-15`).

### 1.3 Nesting

At Jason's "please test what word does with w:document in w:document, and w:body in w:body": `nd-*`, 2026-11-07, 11 files, each twice, the two runs agreeing.

| Shape | Word |
|---|---|
| `w:document` under `w:document`: the only child, two and three deep, after or before the outer `w:body` (`nd-10`, `nd-14`, `nd-11`, `nd-12`) | Refused: "Word experienced an error trying to open the file" |
| `w:document` in a `w:body`, between paragraphs (`nd-13`) | Refused, the same |
| `w:body` in a `w:body`: the only child, between paragraphs, last, three deep (`nd-20` to `nd-24`) | Silent. The tags go, and the words stay in place and in order |

The section is section 1.5's merge, the copies taken in document order. An inner body's stands whole only where nothing follows it (`nd-20`, `nd-23`). Where one does, the later copy's values win, and a `w:orient` only the earlier states stays (`nd-22`, `nd-24`).

### 1.4 What decides an unknown root

`w:foo` holding a `w:document` (`rw-12`) is refused where `w:madeUp` holding a `w:p` opens, and both the name and the content differ. So `mu-*`, at Jason's "do the round for w:madeUp holding a w:document", 4 files, and docx4j's `mr-*`, 20 files (its CR-079 section 1.1), both 2026-11-07, each twice, the two runs agreeing, read again here.

| Root | Holding | Word |
|---|---|---|
| Unknown, in `w:` (`w:madeUp`, `w:foo`) | One or two `w:p`, a `w:tbl`, a block `w:sdt`, a further unknown element holding a `w:p` (`or-31`, `mu-12`, `mr-01`, `mr-10`, `mr-12`, `mr-13`, `mr-17`) | Silent, read as if the tags weren't there, with properties and a `w:pStyle` kept (`mr-18`, `mr-50`) |
| | A `w:body` (`mr-15`) | That body, its section read |
| | A `w:p` and a `w:sectPr` (`mr-11`) | The words. The `w:sectPr` has no effect |
| | Nothing (`mr-16`) | Silent, empty |
| | A bare `w:r` (`mr-14`) | Refused |
| | One or two `w:document` (`mu-10`, `mu-11`, `mr-03`, `rw-12`) | Refused |
| In another namespace, or in none | A `w:p` (`mr-20`, `mr-21`) | Refused |
| | A `w:document` (`rw-11`, `rw-13`) | Refused |
| In a namespace its `mc:Ignorable` names | A `w:p` (`mr-22`) | Silent, empty, the words nowhere in the save |
| `w:p`, `w:sdtContent` (`mr-32`, `mr-30`) | | The words shown |
| `w:tc` (`mr-31`) | | Refused |
| `mc:AlternateContent`, `w:madeUp` in the Choice, `w:document` in the Fallback (`mr-40`) | | The Choice read, the Fallback's words nowhere in the save |

So in the `w:` namespace the content decides, not the name, and outside it a `w:p` is refused too. Known names follow no such rule. WD 0.97's REQ-318 to REQ-320 are drawn from these.

### 1.5 The merge rule

The copies in sections 1 and 1.2 mostly state the same properties, so they couldn't show this, but in six files an earlier copy's `w:orient` survived with the last copy's size, leaving the width below the height (`rt-10`, `rt-40`, `rt-50`, `rt-51`, `nd-22`, `nd-24`). So `sm-*`, at Jason's "order a Word round to identify the real merge rule", 2026-11-07: python's 40 files, each twice, every file silent and the two runs equal. Python read them (its CR-039 section 9.5, at `02506c0`), and this session read all 40 again from Word's saves and agrees, but for the `w:tblInd` below.

**Word applies each copy of a repeated element in order to one object. What a copy states wins, and what it doesn't state stays.**

| Shape | Word | Files |
|---|---|---|
| An attribute only an earlier copy states | Stays, even where the schema has a default: `w:orient`, `w:code`, `w:gutter`, `w:h`, `w:w`, `w:num`, and `w:pgMar`'s others | `sm-01`, `03` to `06`, `10`, `11`, `21` |
| An attribute stated again | The later wins. A later `w:orient="portrait"` beats landscape, so this is a merge and landscape isn't sticky | `sm-02`, `07`, `10` |
| A child only an earlier copy states | Stays: `w:cols`, `w:titlePg`, `w:pgBorders`, `w:lnNumType`, `w:pgNumType`, `w:vAlign`, `w:type`, `w:docGrid`, `w:bidi`, `w:pgSz`, `w:pgMar`. An empty last copy changes nothing | `sm-20` to `sm-32` |
| Joined bodies, a branch's two `w:document`, two `w:sectPr` in one `w:pPr` | The same | `sm-50` to `sm-53` |
| `w:tblPr` | The same for children: a `w:jc` only the first states stays | `sm-60` |
| `w:ind`, `w:spacing`, `w:rFonts` given twice, in one holder or two | The same, one level in: their attributes merge | `sm-70` to `sm-74` |
| `w:vanish` | The last element decides, as `w:b`. A missing `w:val` is a statement here: `<w:vanish w:val="0"/>` then `<w:vanish/>` is on | `sm-80`, `sm-81` |

Two things stand outside the rule:

- **`w:headerReference` and `w:footerReference`.** Only the first copy's are read. A reference on a later copy is ignored, whether or not an earlier one has one, and its part goes from the save. The `footer-refs` round (`ft-*`, 12 files twice, all silent, the runs agreeing) measures it for both elements:
  - `sm-41` saves no header, and `sm-42` keeps the first's. `sm-40`, on the first copy only, keeps it, as `ft-01` keeps a footer.
  - A first copy with no reference leaves none, however many later copies state one (`ft-02`, `ft-05`).
  - Of two references the first's stands, and of three (`ft-03`, `ft-04`).
  - It doesn't go by kind. A later copy's header goes where the first holds a header and a footer (`ft-06`), and a later footer goes where the first holds only a header (`ft-07`).
  - It doesn't go by `w:type` either: a later `first` goes after a `default`, for a footer (`ft-08`) and a header (`ft-09`).
  - It holds across joined bodies: the first body's stands (`ft-10`, `ft-11`).
- **`w:tblW`.** `w:w="4000" w:type="dxa"` then `w:type="auto"` saves as `w:w="0" w:type="dxa"` (`sm-61`), which is neither the merge nor the last copy. It's one file, and nothing tests a lone `w:tblW w:type="auto"`, so this may be how Word reads that element and nothing to do with the repeat. Python's record counts a `w:tblInd` of 10 in it. That isn't the repeat: `te-00`, with one `w:tblW` and no `w:jc`, saves the same `w:tblInd`.

Nothing here goes past one level. Two copies of `w:pgBorders` naming different sides, `w:tabs`, `w:numPr` and the other holders in phase 0's list aren't measured.

## 2. Here today (measured 2026-11-06, `26c2c4f`, jsonix 3.11.0)

Each bridge source loaded, `unmarshalAll`, read through `Texts.of` and `Body.text`, and saved.

### 2.1 A repeated child

| File | Word | Read and saved here |
|---|---|---|
| `rc-10` | `FIRST BODY WORDS`, `SECOND BODY WORDS` | `SECOND BODY WORDS` |
| `rc-11` | `FIRST CONTROL WORDS`, `SECOND CONTROL WORDS` | `SECOND CONTROL WORDS` |
| `rc-12` | `SECOND BOX WORDS`, `FIRST BOX WORDS` | `SECOND BOX WORDS` |
| `rc-21` | `w:ind` and `w:jc` | `w:ind` |
| `rc-23` | `w:b` and `w:i` | `w:i` |

The other 8 match Word, `rc-31` to the same effect: `w:b w:val="false"` stays where Word writes nothing. `pkg.warnings` is empty on all 13 and `lastSave.lost` is empty. CR-025 section 2.6 says why neither its report nor its gate can see this. A part nobody unmarshalled is copied, so only a save after a read loses anything.

### 2.2 The root

| File | Word | `Texts.of`, `Body.text` | `pkg.warnings` | Saved |
|---|---|---|---|---|
| `rw-20`, `rw-21` | One branch | Nothing | `mce.branches_differ` | Both branches |
| `rw-22` | The Fallback, repaired | Nothing | **Nothing** | The Fallback |
| `rw-23`, `rw-26`, `rw-27` | Both documents | Nothing | `mce.branches_differ` | Every branch |
| `rw-24` | Nothing | Nothing | Nothing | The Choice |

- **A write is lost.** `MainDocumentPart.body` runs `doc.body ??= ...` on the object it finds, so here it adds a `body` to the `AlternateContent`, which the marshal ignores. On all seven files `pkg.body.insertParagraph('ADDED', 'End')` reads back as `ADDED`, the save leaves it out, `lastSave.lost` is empty, and neither CR-016's check nor CR-025's gate has anything to see.
- **`mcePreprocess: true` fails the load** with `HierarchyRequestError: Only one element can be added and only after doctype` on six of the seven: `preprocessAlternateContent` puts the branch's content before the element, and a `Document` can't hold two. `rw-24` loads, since nothing is selected.
- **What is said.** The text model's own warnings are empty, so a caller of `Texts.of` sees an empty document and no warning. `mce.branches_differ` names Word's branch and its words, but only on `pkg.warnings`, and on `rw-22` there is nothing at all.

### 2.3 How often

None of the 11,548 XML parts in the 1,045 packages read here has this root, and of the 997 packages compared source against save for CR-025, none holds a repeated child. Both shapes come from probes built to test them. The wider catalogue hasn't been counted (docx4j CR-079's D8). None of the 834 Word main parts holds more than one `w:body` or `w:document`, repeated or nested.

### 2.4 Another root (measured 2026-11-07, the build of `26c2c4f`, jsonix 3.11.0)

Each of section 1.1's files loaded, read through `Texts.of` and `Body.text`, written through `Body`, and saved.

| Root | Load | `Texts.of`, `Body.text` | A write through `Body` |
|---|---|---|---|
| `w:hdr`, `w:ftr`, `w:footnotes`, `w:endnotes`, `w:comments`, `w:styles`, `w:body` | A `WordprocessingMLPackage`, nothing said | Nothing | Lost, as in section 2.2 |
| `mc:AlternateContent`, `w:hdr` in the Choice | The same, with `mce.branches_differ` | Nothing | Lost, as in section 2.2 |
| `w:madeUp` | Loads, and reading the part throws `PartUnmarshalException` | Refused | Not reached |
| Either root under a header's content type | A plain `OpcPackage` with no main part | `Texts.of` refuses | `pkg.body` is undefined |

A save writes the source's words back, as in section 2.2. So on `w:footnotes`, `w:endnotes` and `w:comments` this package says nothing where Word refuses or repairs. It reads less than Word on `w:body` and `w:madeUp`, and on `w:comments`, whose comment Word keeps. On `w:hdr`, `w:ftr` and `w:styles` both show an empty document.

All 834 Word packages of the 1,045 in section 2.3 have `w:document` as the root of the part their `officeDocument` relationship names, under a main document content type, 2 of them macro-enabled. So this shape too comes only from probes built to test it.

### 2.5 Three copies, nesting and an unknown root (measured 2026-11-07, the same build)

- **Three copies.** The last stands, and `pkg.warnings` is empty. Of `rt-*`'s 19 files, 7 match Word: `rt-20`, `rt-22`, `rt-30` to `rt-33` and `rt-41`, `rt-32` to the same effect, keeping `w:b w:val="false"`. Ten read less: `rt-10`, `rt-11`, `rt-12`, `rt-13`, `rt-15`, `rt-16` and `rt-17` keep one copy's words, `rt-21` and `rt-23` one copy's property, and `rt-40`, counted as a match here until section 1.5, loses the second copy's `w:orient`. `rt-50` and `rt-51` read as section 2.2's files do.
- **Nesting.** `w:body` and `w:document` are schema names, so CR-026's pass leaves them alone, and the runtime keeps one in a `w:body` as an item of `body.content`. The text model walks in, so `nd-20` to `nd-24` give Word's words, and `nd-13` gives the inner words where Word refuses, with nothing said. A save writes the nesting back. The outer body of `nd-20` and `nd-23` has no `w:sectPr` where Word uses the inner one. A `w:document` under `w:document` goes to `part.skipped`: `nd-10` and `nd-14` read empty, `nd-11` and `nd-12` read the outer body, and a save after a write is refused.
- **An unknown root.** Every root the runtime doesn't bind throws `PartUnmarshalException` at the read, whichever way Word goes: all of section 1.4's unknown roots, and `w:sdtContent`. `w:p` and `w:tc` bind, and read empty with nothing said, as section 2.4's roots do. `mr-40` reads empty with `mce.branches_differ`.

### 2.6 The merge round (measured 2026-11-07, `1cae4ef`'s build, jsonix 3.11.0)

Each `sm-*` source loaded, `unmarshalAll`, saved, and the save held to Word's in its words and in the properties section 1.5 names. `pkg.warnings` is empty on all 40.

- **Six read as Word does:** `sm-00`, `sm-02`, `sm-33` and `sm-81`, and to the same effect `sm-07`, which writes `w:orient="portrait"` where Word writes nothing, and `sm-80`, which keeps `w:vanish w:val="false"`.
- **Thirty-four don't.** The last copy stands whole, so whatever an earlier copy alone states goes: `sm-01`, `03` to `06`, `10`, `11`, `20` to `32`, `53`, `60`, `71` and `74`. `sm-70`, `72` and `73` are plain repeats in one holder, which keep the last with no second container involved. `sm-61` gives `w:type="auto"` alone.
- **Headers.** `sm-40` has no header where Word keeps one, `sm-41` one where Word has none, and `sm-42` the second where Word keeps the first.
- **Words.** `sm-50` and `sm-51` lose the first body's words, as `rc-10` does, and `sm-52` is section 2.2's root. No other file differs in its words.

The `ft-*` round (measured 2026-11-07, `4003051`'s build, the same runtime, each source held to Word's save in the references it keeps and their parts' words) goes the same way. Only the control `ft-00` reads as Word does. In the other eleven the last copy's references stand and the first's go, and `pkg.warnings` is empty on all twelve. So `ft-01` has no footer where Word keeps one, `ft-02` and `ft-05` one where Word has none, and `ft-06` the second header with no footer where Word keeps the first of each. `ft-10` and `ft-11` also lose the first body's words.

## 3. Proposal

A pass in `src/opc/repair/`, in `XmlPart.preprocessed`, ahead of `readThroughUnknownElements`, reads the root. `isUnchangedSinceRead` goes through the same method and needs nothing new. Section 3.2 says what a repeated element then reads as, and D2 says by which route.

### 3.1 The root

Where the main document part's root is `mc:AlternateContent`, in either load mode:

1. The branch is selected by the rule `preprocessAlternateContent` uses, `READ_UNDERSTOOD_NAMESPACES`, with `Requires` resolved while the element is in place.
2. One `w:document` in the branch stands. Of several the last stands, the earlier ones' bodies put ahead of its own as repeated `w:body`, which section 3.2 then reads by either route. No branch gives an empty `w:document` (D4).
3. The root is removed first and the `w:document` appended after, since a `Document` holds one element.
4. Namespace declarations on the root go onto the `w:document`, so an `mc:Ignorable` prefix only the root bound still resolves. The trial didn't do this, and its files didn't need it.
5. `mce.branches_differ` is reported as now, from the pass, because the hook no longer meets the element. Each branch not taken is one `part.skipped`, so CR-025 holds it with nothing added (D3).
6. A branch taken that holds no `w:document` is read by D11's rules, as if what it holds were the root (REQ-317). No other branch stands in, since Word doesn't read the Fallback's `w:document` there (`or-50`, `mr-40`).

That also ends section 2.2's write loss and `HierarchyRequestError`: after this pass the part holds a `Document`, and the hook sees no such root. Step 6's root goes by D11.

### 3.2 The containers

What was measured, as section 1.5 has it. Anything else stays as it reads today (D6).

| Host, child | The read |
|---|---|
| `w:document`, `w:body` | Every body's content in order, in the last. The closing `w:sectPr` copies as the fourth row. A `w:sectPr` in a `w:pPr` stays |
| `w:sdt`, `w:sdtContent` | Every copy's content in order, in the last |
| `v:textbox`, `w:txbxContent` | Every copy's content, the last first (D7) |
| `w:pPr`, `w:rPr`, `w:sectPr`, `w:tblPr`, and in them `w:ind`, `w:spacing`, `w:rFonts`, `w:pgSz`, `w:pgMar`, `w:cols` | Each copy applied in order: what it states wins, what it doesn't stays, attributes and children alike, in two holders or one |
| An on/off element, as `w:b`, `w:vanish` | The last element whole. A missing `w:val` states on |
| `w:headerReference`, `w:footerReference` | The first copy's stand, of either kind and any `w:type`, across joined bodies too. A first copy with none leaves none. A later copy's go as `part.skipped` |
| `w:tblW` | The last whole, as today, pinned as a known difference until `sm-61` is explained |

This replaces the pass first proposed here, which moved every earlier copy's children to the head of the last and left the runtime to keep the last of each. That gives one `w:ind` of two (`sm-71`, `sm-74`), and never meets a plain repeat in one holder (`sm-70`, `sm-72`, `sm-73`).

A DOM pass can still do this, but it would have to know which children are single-valued, and that is the schema's knowledge, which the runtime already holds. jsonix 3.13.0's `onOverwrittenElement` reports each overwrite: on a trial build it fired on 38 of `sm-*`'s 40 files, all but the control and `sm-52`, whose root section 3.1 turns into repeated bodies. That is the route of the objects package's CR-022, and its D6 takes it to this rule. Whichever route is taken, three things are easy to get wrong:

- "Carry what the lost copy holds and the kept one doesn't" gets `sm-81` wrong, since the `w:val="0"` would be carried over a `<w:vanish/>`.
- That test never fires on `sm-41` or `ft-02`, where the kept copy holds the reference and has to lose it.
- CR-022's D3 has a caller's callback replace the facade's. If `XmlPart`'s read at `XmlPart.mts:441` ever passes one, it has to call `wordOverwritten`, and the re-read at `XmlPart.mts:335` has to read the same way, or a part with the shape never compares unchanged.

### 3.3 What is said

One warning a part with a count, as `br.outside_run` is, since a file read as Word reads it has lost nothing and a warning per child would say otherwise. Its code is settled with docx4j and python first (D8). What goes whole is `part.skipped` beside it, so a changed part holding a later copy's `w:headerReference` or `w:footerReference`, or a branch not taken, is refused, as CR-026 D10 has it for a `w:sectPr` it drops. An earlier `w:sectPr` no longer goes, since it merges.

### 3.4 A scratch trial (2026-11-06, the tree untouched)

The first form of both passes as a `preprocessor` hook, on 20 of the bridge's sources (`rc-10` to `rc-41`, `rw-20` to `rw-24`, `rw-26`, `rw-27`). Each file was read, had a paragraph inserted through `Body`, and was saved, and the save held to Word's in its words, its `w:pPr`, `w:rPr` and `w:tblPr` (Word's own `w:tblCellMar` and `w:tblLook` aside), its page size and its margins.

- **Words:** 20 of 20, and all 20 kept the inserted paragraph.
- **Properties:** 18 of 20. `rc-31` keeps `w:b w:val="false"`, as it does today, to the same effect. `rw-24` has no `w:sectPr` where Word writes a default section (D4).
- **Control:** the same comparison with no passes and no insert, on seven of the files, matches none: `rc-10`, `rc-11`, `rc-12`, `rw-20` and `rw-23` differ in words, and `rc-21` and `rc-23` in properties.

Its limits: the hook runs after CR-026's and CR-021's passes, not ahead of them; it wrote no reports; and there is one small file per shape, a three-part package with no styles part. It ran before section 1.2 and held two copies, and it wasn't run again on `rt-*` or `sm-*`. Its `w:sectPr` rule was the last whole, which its files couldn't tell from the merge, so section 3.2 as it now stands is untried.

### 3.5 `insertOoxml`

The flat OPC route reads through `getContents`, so it takes the root pass and section 3.2's read. The fragment route has to take section 3.2's read too, by calling a DOM pass as it calls CR-021's, or, on D2's (c), through the facade. CR-026 D12's refusal doesn't apply: nothing is lost by reading a repeated container as Word does, where an unknown element has no place to go.

## 4. Not here

- **`rw-25`, a `Choice` nothing understands with no `Fallback`, in a body.** Word shows neither branch and this package the Choice. WD 0.97's REQ-305 now selects nothing, as ISO and Word do, in a body and at the root, for all three engines, and WD 0.98 has that raise `mce.nothing_selected`. This package's change is its own step, with the raise of the objects package, whose CR-021 selects by URI over the understood set docx4j has since published (114 URIs, `824565363`). So the list this waited on has grown, and what's left is that raise and Jason's word.
- **The faults Word refuses:** `wf-10` to `wf-17`. The roots are D11's.
- **A report of every repeated child.** jsonix 3.13.0's callback could now make one. D8 is whether anything is said at all.
- **Another root read as Word reads it** (section 1.1). A header's words in a footnote's separator aren't a reading to copy. D11 says what happens instead.

## 5. Decisions

| | Question | Recommendation |
|---|---|---|
| D1 | When? | **Decided** (Jason, 2026-11-06): after 0.7.0, and it may be deferred indefinitely |
| D2 | Which route for section 3.2? (a) A DOM pass here, with its own list of what is single-valued. (b) Find the pairs and report, reading as today: the gate then refuses a changed part, but the read stays short of Word's. (c) The objects package's CR-022, reading through jsonix 3.13.0's callback, with only section 3.1's root pass here | (c). This was (a) until section 1.5 made the rule general: the runtime knows what is single-valued, and one route serves every caller of the facade. It waits on the objects release holding CR-022 with its D6. CR-022's D5 asks the same question from its side. (a) needs no release, and is the fallback if CR-022 doesn't come |
| D3 | A branch not taken, at a save? docx4j's D4 writes an ordinary `w:document` and the branch goes, as Word's save does | Through CR-025's gate: a part unchanged is copied with its root, one changed is refused unless `pkg.checkLossOnSave` is false, when it is written as Word's save writes it. Keeping every branch, as CR-010 does in a body, would mean every reader starting inside a branch |
| D4 | Nothing selected at the root (`rw-24`)? Word shows an empty document, and so does WD 0.97's REQ-317 | As Word, an empty document, each branch reported, as docx4j's D2, with `mce.nothing_selected` as WD 0.98's REQ-305 has it at a root. Whether it gets Word's default `w:sectPr` waits on phase 0. The body's case is section 4's |
| D5 | Other parts with this root? A header, notes, comments, styles, numbering, settings. Word is unmeasured | The main part only until phase 0, then each part where Word reads the branch, as docx4j's D1 |
| D6 | Which repeated children? | Section 3.2's table, as WD 0.100's REQ-322 and the objects package's CR-022 D6, and those phase 0 adds. The alternative is `w:body` and `w:sdtContent` alone, the two losing words in plain text. `w:tblW` stays as it reads until a lone `w:tblW w:type="auto"` is measured |
| D7 | `w:txbxContent` in reverse order? Measured with two, three and four copies, in a VML box only | Follow it once phase 0 has a second shape, and take `wps:txbx` by what it shows, as docx4j's D6 |
| D8 | The warning's code? | One a part with a count. The code is agreed with docx4j and python before any is built, as `value.out_of_range` and `target.misplaced` were. Python's CR-039 proposes `part.read_as_word`, and docx4j's view is that a file read as Word reads it has lost nothing and needs none. WD 0.97 leaves it open |
| D9 | Who goes first? | Neither waits. Phase 0 is one round for both, and no golden holds either shape, so the goldens don't order them. Here the root is a DOM pass where docx4j's needs a binder rebuild, so it's the cheaper to try first, if Jason wants either |
| D10 | If this stays deferred, does section 2.2's lost write? | No. Refuse it alone: `MainDocumentPart.body` throws where the part's root isn't a `w:document`, in place of `doc.body ??= ...`, and `preprocessAlternateContent` leaves a root alone. About a quarter day, and it can go in any release. `Body.text` then refuses where it reads empty today, which is the honest answer. It needn't hold 0.7.0, since no file here has the shape. As worded it also refuses section 2.4's roots, among them a bare `w:body`, which Word reads (D11) |
| D11 | Another root in the main part? | As WD 0.97's REQ-318 to REQ-320 (section 1.4), which python's CR-039 builds (Jason's Q9 there) and docx4j's D9 recommends. Refused: `w:footnotes`, `w:endnotes`, `w:tc`, `pkg:package`, Word 2003's `wordDocument`, an unknown `w:` root holding a `w:document` or a bare `w:r`, and a root outside the understood namespaces holding a `w:document` or a `w:p`. Read: `w:body`, `w:p`, `w:sdtContent`, and an unknown `w:` root holding `w:p`, `w:tbl`, a block `w:sdt` or a further unknown element. Every other root gives no text, with `part.unexpected_root` naming it on the part and in the text model's warnings (REQ-801), and D10 refuses its write. A header's words in a footnote's separator aren't followed. If only D10 is built, the roots REQ-319 reads stay unread, which is safe and short of the WD |
| D12 | Nesting (section 1.3)? WD 0.97 requires nothing of either shape | Leave both as they read. A nested `w:body` gives Word's words here, and no file read here holds either (section 2.3). A `w:document` in a `w:body`, which Word refuses and this package saves, would be a rule in the objects package's `onMisplaced` if a file ever shows it |

## 6. Phases

- **Phase 0, Word check 81, one round with CR-079's phase 0.** Its list: this root on the other parts (D5); an empty `w:body` among others; a `w:background` and an `mc:Ignorable` on one of two `w:document`; `w:sdtContent` at run, row and cell level; `w:txbxContent` in `wps:txbx` and a second VML shape; `w:pPr` and `w:rPr` holding a nested `w:rPr`, `w:numPr` or `w:pPrChange`; an unknown element in the branch taken and in one not taken. Added here:
  - **Answered by `sm-*` and `ft-*` (section 1.5):** a repeated property with different attributes, `w:sectPr` and `w:tblPr` with different properties, a `w:sectPr` in a `w:pPr` given twice, `w:vanish`, `w:footerReference`, and both references across joined bodies.
  - **Left by them:** a lone `w:tblW w:type="auto"`, and `w:tblW` given twice with `w:w` on the later; a reference on a later body where the first holds no `w:sectPr`; a `w:type` of `even`, or `first` before `default`; a child of a child, as two `w:pgBorders` naming different sides, and `w:tabs` in two `w:pPr`.
  - A `w:pPr` or `w:rPr` given twice and not adjacent, with a run or a `w:t` between.
  - The other single-valued containers: `w:tcPr`, `w:trPr`, `w:tblGrid`, `w:sdtPr`, `w:sdtEndPr`, `w:numPr`.
  - A repeated container in a header, a note and a comment, and `w:pPr` and `w:rPr` in a `w:style` and a `w:lvl`.
  - `mc:AlternateContent` under `w:document` holding a `w:body`, and one nested at the root.
  - A further unknown element holding a bare `w:r` or a `w:document`, which REQ-319 reads and the rule behind it may refuse (WD 0.97's NOTE).
  - A bare `w:body` in a branch taken, and an empty branch taken with a `w:document` in the Fallback (docx4j's, its CR-079).
  - `w:hdr` as the root with a `w:footnoteReference` and a footnotes part, to see whether the words in the `separator` footnote ever draw.
- **Phase 1, the root.** Section 3.1, and D11 if taken. Two `w:document` need section 3.2's first row, so on D2's (c) they join only once phase 2's release is in.
- **Phase 2, the repeated elements,** by D2, D6 and D7, and the warning (D8).
- **Phase 3, the records:** CLAUDE.md, the index, CR-025 section 2.6, `test/README.md`.

About a day and a half with phase 0 by D2's (a), and less here by (c), which needs the objects release holding CR-022 and its D6. Phase 1 needs no release of anything.

## 7. Tests and oracles

- `test/repeated-containers/`: the `rc-*`, `rw-*`, `rt-*`, `nd-*`, `mu-*`, `mr-*`, `sm-*` and `ft-*` sources and Word's saves, and those of the `or-*` files section 1.1 names, copied unedited with a SHA-1 manifest, outside `fixtures/`, as `test/br-placement/` is, with phase 0's.
- `test/repeated-containers.test.mjs`: every file held to Word's save as section 3.4 holds it, by `Texts.of`, `Body.text` and a save after `unmarshalAll`, in both load modes. A write on each root file kept. The gate on a changed part, `checkLossOnSave`, and an unchanged part copied with its root. `rw-25` pinned as it stands until section 4's step. Section 2.4's roots as D11 has them.
- The passes change nothing on a file without the shapes: every fixture's tree compared with and without them, and each pass's cost on the four largest main parts, as CR-021's was (0.4% of a load).
- Parity: nothing moves, as no golden's fixture has either shape. When docx4j's CR-079 lands, a fixture of each joins the goldens.

## 8. Risks

- **`w:tblW`.** One file fits neither reading. If a lone `w:tblW w:type="auto"` doesn't explain it, other leaves may not merge as `w:ind` does either, and the rule's reach is only the elements section 1.5 names.
- **The callback,** on D2's (c): the read rests on what jsonix 3.13.0 reports, and a test pins it, so a later jsonix that changes this announces itself.
- **Order against CR-021 and CR-026,** on D2's (c): their DOM passes run first and see each copy, not one container. A bare `w:br` at the end of a first `w:body` looks for its next paragraph there. No file tests it.
- **A walk more on every part** for the root, or two on D2's (a). They may fold into CR-026's, which already meets the same hosts.
- **REQ-321's third sentence,** on an embedded package with no main document part, is new, and no engine has been tested against it. `Extractor.mts`'s header has embedded packages refused, which wasn't tested here.
- **A part read and saved no longer matches its source's root,** once `checkLossOnSave` is false. That's D3, and Word's save does the same.

## 9. Registry

`core-ts/CR-028`, proposed, `depends_on: [core-ts/CR-026]`, for the pass order. Not `docx4j/CR-079` (D9). Sent once this is committed.

## 10. Sources

- The share's `bridge/portfolio/`: `rc-*`, `rw-*`, `rt-*`, `nd-*`, `mu-*`, `mr-*`, `sm-*`, `ft-*`, `te-00` and the `or-*` files section 1.1 names, Word's saves read here. `footer-refs` was run for the objects package, and WD 0.101's REQ-322 (portfolio `7150868`) has it.
- `docx4j-portfolio/docs/text_extraction_specification_wd.md`, WD 0.97 at `2fd79dc`: REQ-305 and section 3.2; WD 0.100 at `9034ad0`: REQ-322, with `w:tblW` unexplained.
- docx4j-python's CR-039, decided at `89ed0eb`, and its section 9.5 at `02506c0`.
- The objects package's CR-022, its sections 3.1 and 5 and D5 and D6, uncommitted when read.
- `docx4j-portfolio/docs/text_extraction_behaviour_model.md` section 7.15; the registry's `objects-ts/repeated-child-word-check` and `root-element-word-check`.
- docx4j's `docs/developer/change-requests/CR-079-mc-root-and-repeated-containers.md`, its sections 1, 4, 5 and 6.
- `docx-tool-identifier`'s `runs-local/two-bodies/`.
- CR-025 section 2.6, CR-026 sections 3.7 and 5.2, CR-021.
- `src/parts/XmlPart.mts`, `src/parts/wml/index.mts`, `src/opc/mce/preprocessor.mts`, `src/opc/repair/`.
