# CR-016: Misplaced content refused before it reaches a saved file

**Status:** Proposed 2026-10-26, at Jason's request ("yes, open a CR for refusing misplaced content"), on docx4j
CR-064's hand-off (its section 7: whether the ports refuse misplaced children is theirs to decide, with its table of
mistakes as the list worth catching). Today's behaviour measured the same day (section 2). D1 to D6 decided 2026-10-27 as recommended (section 4); check 72 (D5) written and run the same day (sections 7 and 7.1: 11 of the 18 files don't open in Word 365, no repair offered, 7 are repaired silently, one losing its content). Phase 1 is the objects package's CR-017; phase 2 waits on it.

## 1. Why

A run straight in the body, a table inside a paragraph or a cell straight in a table is never legal OOXML. Word meets
one on opening, reports unreadable content, and repairs the file by dropping what it can't place, far from the line of
code that put it there (docx4j CR-064 section 0). docx4j's content lists hold `Object`, so its compiler can't catch
the mistake. Since CR-064 phase 1 (docx4j `befc57371`) its `ArrayListDocx4j` refuses each such child as it is added,
with the fix in the message, for example "You can't add Text directly to a hyperlink; ...".

This package has no such list type: a content list is a plain array, so nothing can see an element being added. What
the mistakes are, where they can be caught here, and what to do about them is this CR's question.

## 2. Today (measured 2026-10-26, at `fd898b9`)

**At compile time,** most of CR-064's table is already caught. The objects package types each content list as a union
of the elements the schema allows there. `Body.content`, for example, is a union of `TypedNamedValue<P>`,
`TypedNamedValue<Tbl>`, `TypedNamedValue<SdtBlock>`, and so on, with no run and no cell. So in TypeScript
`body.content.push(el.r(...))` is a type error, which is more than Java can check. Three things get past it:
- JavaScript callers;
- a value cast to `any`, or built by hand without its type;
- a `{ name, value }` pair whose name and value disagree.

**The content API's inserts already refuse** a misplaced block-level element before anything changes.
`Body.insertXml`, `Body.insertOoxml`, `Body.insertElement` and `Range.insertOoxml` given a run or a cell at body
level each throw `Docx4JException`, for example "A body cannot hold r (org_docx4j_wml.R); block-level elements only
(w:p, w:tbl, w:sdt, ...)". `Paragraph.insertOoxml` given a table places it after the paragraph, at block level.

**A direct edit of the tree is unchecked,** and saved as written:

| Edit | What the save writes |
|---|---|
| a run pushed into `body.content` | `<w:r>` as a child of `<w:body>` |
| a table pushed into a paragraph's content | `<w:p>...<w:tbl/></w:p>` |
| a paragraph pushed into a paragraph's content | `<w:p>...<w:p/></w:p>` |
| a cell pushed straight into a table's content | `<w:tbl>...<w:tc>` beside the rows |
| a string pushed into a run's content | the save fails: "Problem saving part /word/document.xml", with nothing about the cause |

What Word does with each of these files isn't measured here. CR-064 section 2 has it from docx4j's forum threads:
unreadable content, then a repair that drops content.

## 3. Where a check can run

- **At save, over the parts being marshalled.** Only a part that was unmarshalled, or given contents, is
  re-marshalled; every other part is copied byte for byte. So the check walks only what the caller could have changed,
  and a document read and saved untouched is never refused. There is a precedent: `XmlPart.getBytes` already refuses
  a part holding a character XML 1.0 excludes, naming it (CR-002 section 60, D3).
- **On request,** as a check over a part, a story or the package: it returns each finding, with its path, the rule
  and the fix, and changes nothing. It is for tests, for the editor before it saves, and for a caller who wants the
  answer without a save.
- **In the objects facade's marshal,** for every caller of the facade, whether or not they use this package. By the
  dividing rule, a check that needs only an object tree belongs there (CLAUDE.md, "The dividing rule").
- **As each element is added,** through proxy-wrapped arrays. This is rejected: every content list of every part
  would become a proxy. That costs on every read, breaks array identity, which the views rely on (`indexOf` on the
  container), and Jsonix builds the arrays itself.

## 4. Decisions

**Decided 2026-10-27 as recommended** (Jason: "CR-016 D1 to D6 as recommended, but D5 do check 72"): D5 (a) is not optional; check 72 is written (section 7) and run the same day (section 7.1).

| # | Question | Options | Recommendation |
|---|---|---|---|
| D1 | Where the rules and the walk live | (a) the objects package (its own CR, since proposed as objects CR-017, `f2be032`: its builders refuse a misplaced child, and `onMisplaced` reports each one on read and on save, changing nothing), called here at save and by this package's check; (b) here | (a), by the dividing rule: the check needs only an object tree, and the editor, docx4j-python's users of the facade and the marshal itself can call it there. This package wires it in. |
| D2 | What a save does with a finding | (a) refuse the save, naming the part, the path, the rule and the fix, with an option that turns it off (`pkg.checkContentOnSave = false`) and reports instead; (b) report on `pkg.lastSave`, and write; (c) nothing at save, the check API only | (a). Word's repair drops content, silently, at the reader's end, so a file holding a misplaced child loses data either way; a refused save loses nothing and names the line to fix. The opt-out is for a document another producer wrote that way, whose part the caller then touched (D6). |
| D3 | The rules | (a) CR-064 section 8's five sets, by `TYPE_NAME`: each refuses only elements the schema never allows under that parent (run-level containers: a table, a paragraph, bare text, a drawing, a break, a `w:pPr`, a string; a run: a run, a paragraph, a table, a `w:rPr`, a string; block-level containers: a run, bare text, a drawing, a break, a cell, a row, a `w:sectPr` in the body, a `w:tcPr` in a cell, a string; a table: a cell, a paragraph, a run, a table, a `w:tblPr`, a string; a row: a paragraph, a run, a table, a `w:trPr`, a string), and a name and value that disagree; (b) an allow-list per parent, derived from the schema | (a). Never an allow-list: a content list legally holds range markup, `mc:AlternateContent`, `w:customXml` and `w:sdt` wrappers, revisions and DOM for what nothing binds, and a wrong refusal of a legal document is the worst outcome (CR-064 D8). `mc:AlternateContent`'s branches are checked by their host's set. |
| D4 | The messages | (a) docx4j's ("You can't add a table (Tbl) directly to a paragraph; it should be a sibling of the paragraph"), with the part and the path | (a): one wording across the ports, the kind a person or an LLM fixes code from. |
| D5 | Word's answer for each row | (a) a check 72: a file per row of section 2's table, opened in Word 365 and saved, recording the repair and what it drops; (b) take CR-064's reading | (a), cheaply. It turns "Word repairs it" into a measurement, and settles D2's premise. |
| D6 | A document that already holds a misplaced child | (a) the check runs at save on any re-marshalled part; a finding in content the caller didn't add still refuses, and the message says the opt-out; (b) the load records each misplaced child it meets, and the save refuses only new ones | (a), with the fixture gate (CR-064 D9): if the corpus of fixtures and Office-made files trips a rule, that rule reports instead of refusing. (b) needs the load to walk every part, which it doesn't do today, and identity to survive edits. |
| D7 (2026-10-28) | A misplaced child in a part this package reads | (a) a part warning per misplaced child the objects read report gives (`onMisplaced` on unmarshal, its CR-017 and CR-017.text), code `content.misplaced`, on `XmlPart.warnings` beside `value.out_of_range`, with the container's type, the child, its index, the rule, docx4j's message and, for a string, the text; the part's tree unchanged; (b) D6 (a) alone, the save the only check | (a). D6 (b) was set aside because the load would have had to walk every part; the facade's read report now costs nothing on a part this package unmarshals anyway, as the out-of-range warning does. It is the only place a reader learns of a string in a run (72m), which the runtime drops before the tree holds it and the save can never see. Only a part read is warned on, as for every part warning. |

**D7 decided 2026-10-28** (Jason: "add the part warning per misplaced child read"), as recommended; in phase 2.

**D7's code and fields are docx4j-python's** (its CR-020 section 11.12 and `docx4j_py/validator.py`'s `MisplacedContent`,
85c3261; one code across the ports, as `value.out_of_range`): code `target.misplaced`, not `content.misplaced`, for the load
warning, the save report's rows and the refused save's error alike. Field for field, a field with no value left out:
`partName` (`/word/document.xml`); `path`, on save, the child's own element the last step (`/w:document/w:body[1]/w:p[1]/w:r[1]/w:tbl[1]`),
the container's for a string (python's departure 5, a single-valued field's item, has no case here: the facade tests lists);
`address` absent (python's content API address; agreed absent for `out_of_range` too); `holder`, the container's class as
docx4j's Java name (`R`, `P`, `Tbl`, `CTSdtContentRun`): the objects `TYPE_NAME`'s last segment, which is that name;
`field` absent (the container's list property; the objects report doesn't name it; python's is snake_case);
`index`; `item`, the child's class the same way (`Tbl`, `R`), absent for a string (python writes its type name `str`; this
package's `text` carries the string instead); `element`, the child's prefixed name, absent for a string; `message`,
docx4j's wording here where python writes its own sentence in `checkParent`'s style; `hint` absent (python's "where the item
goes"; docx4j's message holds the fix in its second clause). This package's own, with no counterpart there: `rule`
(`misplaced` or `name-value`), `text` (a string's content) and, on load, `line` where the parser gives one. Python's
rule is an allow rule from its model (its D2), a superset of CR-064's refuse sets; every CR-064 case is one of its.

## 5. Phases

| Phase | Work | Status |
|---|---|---|
| 0 | This CR; today's behaviour measured (section 2) | done 2026-10-26 |
| 1 | The objects package's CR (D1 (a)): the rules (D3) and the walk over a tree, with the messages (D4); its tests per rule and per container, legal children accepted | done there as objects CR-017 (proposed `f2be032` 2026-10-26; M1 to M7 decided and implemented 2026-10-28, bb18596, the `'#text'` read report 9ac19ae and 2b87bdf; unreleased, 0.7.0) |
| 2 | Here: the check at save, through `onMisplaced`'s save-side report (D2, refusing where a report arrives unless turned off), and on request (D6); the part warning per misplaced child read (D7, `content.misplaced`); the fixture gate over `test/fixtures/` and the Office-made corpus; tests per row of section 2's table and check 72's files | after phase 1: the objects package's CR-017 implemented 2b87bdf (2026-10-28), unreleased; starts at the 0.7.0 raise |
| 3 | Check 72 (D5), Word's repairs | done: written and run 2026-10-27 (section 7.1): 11 of the 18 don't open ("Word experienced an error trying to open the file", no repair offered), 7 open with no message and are repaired silently, one losing its content |

## 6. Risks

- **Refusing a save of a legal document.** Each rule refuses only what the schema never allows under that parent, and
  the gate (D6) runs the rules over every fixture first.
- **Cost at save.** The walk visits each element of each re-marshalled part once, with a set lookup per child. The
  marshal itself visits each element already, so the walk adds a fraction to it; this is measured in phase 2.
- **TypeScript users see nothing new** for what their compiler already refused. The gain is for JavaScript callers,
  for values typed loosely, and for the string in a run, whose save fails today with no cause given.

## 7. Check 72 written, and what this package does with its files today (2026-10-27)

Jason: "CR-016 D1 to D6 as recommended, but D5 do check 72". D1 to D6 are decided as section 4 recommends. Check 72
(`test/README.md` check 72) is 18 files on the share at `__tmp/72/`, as `fixtures/check72/build.mjs` writes them, one
misplaced child each between a paragraph BEFORE and a paragraph AFTER, so what Word drops or keeps shows against them, and
one form per file, since Word refuses a whole document for some faults (docx4j CR-065 round 1) and a refusal of a file
holding several would say nothing about each. The forms cover each rule of CR-064 section 8's five sets at least once, and
the string in a run: in the body a run (72a), bare text (72b), a cell (72c), a row (72d) and a break (72e); a run beside a
cell's paragraph (72f); in a paragraph a table (72g), a paragraph (72h), bare text (72i) and a bare break (72j); in a run a
run (72k), a paragraph (72l) and a string (72m); in a table beside its row a cell (72n), a paragraph (72o) and a table (72p);
in a row beside its cells a paragraph (72q) and a run (72r). Jason opens each, says whether it opens clean, offers a repair
or is refused, and saves each that opens as `<name>-word365.docx`; the saves then say what Word keeps.

**What this package does with each file today** (`test/check72.test.mjs`, measured 2026-10-27, before phase 2). Every file
loads with no warning: the facade's unmarshal does not refuse a child its content model has no place for. The readers see
the misplaced content in two files only, a cell and a row straight in the body (72c, 72d), whose paragraphs `Body` walks
as it walks any cell's; in the sixteen others `Body.text` holds BEFORE and AFTER and nothing of the child. Yet a save writes
the child back as it was in seventeen of the eighteen: the tree keeps it (a bare break in the body or in a paragraph, 72e
and 72j, read typed as `Br`; the rest as content the model does not type), invisible to the readers and faithful on save.
One is lost on the way: a string straight in a run (72m). (Until 2026-10-28 this section counted the two bare breaks as
lost too; the test's check looked for a MISPLACED marker the break files don't carry. The objects package's session saw it
while revising its CR-017; measured again, both breaks are written back where they were.) So a document holding a misplaced
child round-trips here unseen, and reaches Word with the
fault intact, which is D6's case: the check at save (phase 2) is what will report it, and the opt-out what lets a caller
write such a file knowingly. None of this is what Word does; that is the check.

Not every misplaced element reads at all. The objects session reports (2026-10-28, its CR-017 M7) that a child the objects
context declares only locally fails the part's unmarshal outright, with the runtime's "could not be unmarshalled as is not
known in this context and the property does not allow DOM content": a run's `w:instrText`, `w:delInstrText`, `w:object`,
`w:ruby`, `w:footnoteReference` or `w:endnoteReference` straight in a paragraph or the body, and `w:dir`, `w:bdo` or `w:subDoc`
in the body: nine elements its CR-009 mapping declares globally beyond docx4j's roots. Pushed from code, such a child saves
today with `xsi:type="w:CT_Text"`, which the facade then can't read back. Its M7 first left that as it is; later the same day
it recommends, deliberately diverging from docx4j's roots (docx4j-12's suggestion, docx4j calling it harmless since its own
inspector reads marshalled XML and neither the MCP server nor docx4j-python consults roots), making the nine roots, so a
misplaced one reads typed, is reported by `onMisplaced` where CR-064's rules refuse it (the two field-instruction elements,
as Text), saves as held, and is written from code without the `xsi:type`. Measured there on a prototype: the 342 fidelity
parts marshal byte-identical. For this package, if Jason takes it: such a part loads instead of failing (the bytes were kept
either way), D6 sees it on load, and phase 2's refused save applies to it as to any other. Fifteen elements still fail the
read when misplaced inside a `w:r` only: `w:customXml`, `w:moveFrom`, `w:moveTo` and twelve range markers. The choice is the
objects package's, for Jason; this package takes whichever lands, through the raise. Check 72 has no file of that kind: its
eighteen all read.

**The objects package's CR-017 gate, run here (2026-10-28, at bb18596 through a `--no-save` install, the locked 0.6.0
restored after).** Every XML part of the corpus the objects 0.6.0 raise was byte-checked over (`test/fixtures` and docx4j's
charts corpus: 542 files on disk, 9,075 XML parts, 8,454 with a typed root) read through the facade's `unmarshalString` and
written back through `marshalString`, `onMisplaced` on both: no read or save failed, and outside check 72's own files (on
disk, git-ignored) nothing is reported. Check 72's files give 34 reports: each of 17 children once on read (rule
`misplaced`, docx4j's message, the container's type and the child's) and once on save with its path
(`/w:document/w:body[1]/w:tbl[1]/w:tr[1]/w:r[1]` for 72r). The one not reported is 72m, the string in a run: the runtime
drops character data in a non-mixed element on read, before the rule sees it, and so the tree holds nothing to report on
save; only a string pushed from code reaches the save-side report. Raised with the objects session: unreported by design
for now (its CR-017 section 4.2; the gate run in its section 10). The runtime has no hook for it: `@docx4j/jsonix`'s
`ClassInfo` unmarshal skips character data in a class with no mixed property silently, and its only callbacks are
`onUnexpectedAttribute` and `onUnexpectedElement`. Word's repair drops the string too (72m), and so does docx4j's JAXB.
Closing it would take a jsonix CR, an `onUnexpectedText` for non-whitespace character data in content that isn't mixed,
which the facade would turn into a `'#text'` read report and D6 into a load warning. Not proposed; Jason's call. This CR's
view: worth proposing, at low priority, since 72m is the one of the eighteen forms in which content is lost with no trace,
here and in Word alike, and a load warning is the only place a reader could learn of it; phase 2 does not wait on it. Jason asked for
it the same day ("do you recommend the jsonix CR?"; yes); proposed to the jsonix session, which designs it; the facade's
`'#text'` read report and D6's warning follow through the ranges.
Proposed there the same day as jsonix CR-013 (`docs/change-requests/jsonix-CR-013-report-dropped-character-data.md`,
e859805 on its master, unpushed; planned as 3.10.0; registered as `jsonix/CR-013`, proposed): a per-call unmarshaller option
`onUnexpectedText(text, node, classInfo)`, called once per text or CDATA node `ClassInfo.unmarshal` drops that holds a
character outside XML whitespace, in document order, before the container's `onElement`, the data still dropped and the
return value ignored. Its measurement found two more silent drops beside the non-mixed class: text in a class with no
content at all (`<w:b>x</w:b>`, 852 classes), which the hook covers, and text the runtime's own `isEmpty` counts as space,
26 characters including the no-break space, so `<w:r>&#xA0;<w:t/></w:r>` is dropped as whitespace today. Its T1 to T3 are
for Jason (XML's four whitespace characters rather than the runtime's 26; reporting in mixed classes too; the name and the
node argument). The objects facade would collect the calls by `parentNode` and report them in its composed `onElement`.
Jason decided T1 to T3 as recommended and the hook is implemented, jsonix master c259464 (local, unreleased; 3.10.0 on his
word after the objects package runs it over its fidelity parts), and released as 3.10.0 the same day (tag 5f9db0b; 0 reports
over the objects package's 359 corpus checks, and with the option unset every part re-marshals byte-identical to 3.9.0's).
The objects facade's `'#text'` read report landed the same day (its CR-017.text, 9ac19ae, over a jsonix ^3.10.0 floor,
unreleased, in 0.7.0): a file's string straight in a CR-017 container reported on read as child `'#text'`, rule `misplaced`,
with `text` the string and `index` where it stood among the list's items; nothing on save, the string being gone from the
tree; text inside a class that is no container (`<w:b>x</w:b>`) and whitespace not reported. The gate rerun here at 9ac19ae
(the same 542 files and 9,075 parts, 0 failures, the 34 element reports as before) gave no text report, 72m's included.
Cause, found in the facade and sent to the objects session: a string that is its container's only child goes unreported,
the read-side reporter returning where the container's list is undefined, which the runtime leaves it when no element
child was read; a string between two `w:t` (the facade's own test) is reported, at index 1, and one first or last beside a
`w:t` at 0 or 1. 72m's form is the only-child one. Fixed there the same day (2b87bdf: an unset list taken as empty for the
dropped-text pass; its smoke test now 72m's own form and the five probes; its CR-017 section 10 records the miss and the
reruns). Rerun here at 2b87bdf: 35 reports, the 34 as before and 72m's one on read, `'#text'` at index 0 in the run with
the string message, nothing on save; no real file in the corpus holds such a string; 0 failures. The locked 0.6.0 and jsonix
3.9.0 restored after each run. D6's load warning for it follows at the 0.7.0 raise.

Later the same day the objects package mirrored docx4j CR-064's addendum (ce87ec6d7, Jason's yes) at 202b58d: `w:delText`
refused where text is and `w:tab` where a break is, in both sets; a content control at the wrong level a class rule
(`misplaced`, docx4j's message) in every set rather than the name rule (run content refuses block, row and cell controls,
block content run, row and cell controls, a run any, a table any but a row control, a row any but a cell control); the
`w:tcPr` message naming its container. The same commit fixed a false positive on a legal document, present since bb18596: a
kept `mc:Choice` types its children by their global declarations, and the global `w:sdt` is `SdtBlock`, so a legal run-level
control in a paragraph's Choice was reported, and its runs with it; now a branch's children are tested leniently, legal if
the host's list names them, and sdt content and custom XML below a branch at the level the branch's host gives. That is
the kind of wrong refusal D6's fixture gate exists for; phase 2's gate runs over the kept-branch fixtures too. Rerun here
at 202b58d: the same 35 reports, report for report, 0 failures. Nothing in this package matched on `name-value`, phase 2
not having started.

**Codes shared with docx4j-python** (its CR-020 section 11.12, 85c3261, 2026-10-28, for phase 2's warnings): its validator
reports a misplaced child as `target.misplaced` beside `value.out_of_range` (this package's CR-017 fields and storage strings),
and a table cell that doesn't end with a paragraph as Word reads it as `cell.no_end_paragraph`, with `part_name`, `path`
(`/w:document/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]`), `address` (`body/0/0/0`), `ends_with` (`w:tbl`), `message` and `hint`;
the rule is CR-002 section 55's `TableRules.endsWithParagraph`, ported there, looking past range markup, through block-level
controls and custom XML, and through the branch of a kept `mc:AlternateContent` that Word reads. Its checked save (off by
default) refuses a broken cell only when the session made it; one the source held is saved and listed. If this package
reports either condition it uses the same codes: D7's warning is `content.misplaced` here only until phase 2 settles the
name with docx4j-python (one code across the ports, as `value.out_of_range` is). Its out-of-range check flags 413 of the 660
integer cases to this package's 432: the 19 are `w:gridCol`, `w:wBefore` and `w:wAfter`, which its limits table lists as
unmeasured and so doesn't report (Jason, the same day); this package reports them from the objects table, held to Word's
saves (CR-017 E5).
Phase 2 can build on bb18596 through a file dependency; whether it waits for the 0.7.0 release is Jason's call.

### 7.1 Check 72's run (2026-10-27)

Jason opened the 18 files in Word 365 (16.0.20326.20232, PC, each in its own session; "for the cases which did not give
an error on opening, i have saved a screenshot and -word365.docx"). Eleven gave an error on opening and have no save: a
run, bare text, a cell or a row straight in the body (72a to 72d); a run beside a cell's paragraph (72f); a table, a
paragraph or bare text inside a paragraph (72g, 72h, 72i); a run in a run (72k); a cell beside a table's row (72n); a run
beside a row's cells (72r). For each of the eleven Word said "Word experienced an error trying to open the file" and
offered no repair (Jason, the same day), so the file does not open at all. Seven opened with no message at all, and Word's save of each (`test/check72/`, with its screenshot and a
SHA-1 manifest) shows a silent repair:

| File | The child | What Word's save holds |
|---|---|---|
| 72e | `w:br` in the body | the break is the first run of the next paragraph: `<w:p><w:r><w:br/></w:r><w:r><w:t>AFTER</w:t></w:r></w:p>` |
| 72j | `w:br` in a paragraph, outside any run | the break goes into the run that follows: `<w:r><w:br/><w:t>then a run</w:t></w:r>` |
| 72l | `w:p` inside `w:r` | the inner paragraph's run joins the outer paragraph (`outer text MISPLACED paragraph in a run`), and an empty paragraph follows it |
| 72m | character data in `w:r`, outside `w:t` | **dropped**: the paragraph is `<w:p/>`, the text gone, with no message |
| 72o | `w:p` beside a table's row | the paragraph stands before the table, in the body |
| 72p | `w:tbl` beside a table's row | the nested table is the outer table's first row: its one cell in the first column, `w:gridAfter` 1 |
| 72q | `w:p` beside a row's cells | the paragraph is the first in the row's first cell, before `CELL A1` |

So Word's reading of a misplaced child is, where it reads it at all, "the nearest place it fits": a break into the next
run, a paragraph into the next cell or before the table, a table as a row. Content is lost in one of the seven, and
silently (72m), and the other eleven don't open at all, with no repair offered, which is more than the forum reports
("unreadable content", then a repair) suggested. D2's premise
holds twice over: a file this package writes with a misplaced child either fails to open in Word or is changed on
opening with no word to the user, and in one form loses its content. Phase 2's refusal at save stands as decided.

This package reads each of Word's saves as any file: BEFORE, AFTER and the moved text where Word put it, the string gone,
no warning (`check72.test.mjs`, under the manifest). What this package does with the original files is unchanged from
section 7: the eleven Word can't open are among the seventeen it saves back as they were.
