# CR-016: Misplaced content refused before it reaches a saved file

**Status:** Proposed 2026-10-26, at Jason's request ("yes, open a CR for refusing misplaced content"), on docx4j
CR-064's hand-off (its section 7: whether the ports refuse misplaced children is theirs to decide, with its table of
mistakes as the list worth catching). Today's behaviour measured the same day (section 2). D1 to D6 open (section 4).

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

| # | Question | Options | Recommendation |
|---|---|---|---|
| D1 | Where the rules and the walk live | (a) the objects package (its own CR, since proposed as objects CR-017, `f2be032`: its builders refuse a misplaced child, and `onMisplaced` reports each one on read and on save, changing nothing), called here at save and by this package's check; (b) here | (a), by the dividing rule: the check needs only an object tree, and the editor, docx4j-python's users of the facade and the marshal itself can call it there. This package wires it in. |
| D2 | What a save does with a finding | (a) refuse the save, naming the part, the path, the rule and the fix, with an option that turns it off (`pkg.checkContentOnSave = false`) and reports instead; (b) report on `pkg.lastSave`, and write; (c) nothing at save, the check API only | (a). Word's repair drops content, silently, at the reader's end, so a file holding a misplaced child loses data either way; a refused save loses nothing and names the line to fix. The opt-out is for a document another producer wrote that way, whose part the caller then touched (D6). |
| D3 | The rules | (a) CR-064 section 8's five sets, by `TYPE_NAME`: each refuses only elements the schema never allows under that parent (run-level containers: a table, a paragraph, bare text, a drawing, a break, a `w:pPr`, a string; a run: a run, a paragraph, a table, a `w:rPr`, a string; block-level containers: a run, bare text, a drawing, a break, a cell, a row, a `w:sectPr` in the body, a `w:tcPr` in a cell, a string; a table: a cell, a paragraph, a run, a table, a `w:tblPr`, a string; a row: a paragraph, a run, a table, a `w:trPr`, a string), and a name and value that disagree; (b) an allow-list per parent, derived from the schema | (a). Never an allow-list: a content list legally holds range markup, `mc:AlternateContent`, `w:customXml` and `w:sdt` wrappers, revisions and DOM for what nothing binds, and a wrong refusal of a legal document is the worst outcome (CR-064 D8). `mc:AlternateContent`'s branches are checked by their host's set. |
| D4 | The messages | (a) docx4j's ("You can't add a table (Tbl) directly to a paragraph; it should be a sibling of the paragraph"), with the part and the path | (a): one wording across the ports, the kind a person or an LLM fixes code from. |
| D5 | Word's answer for each row | (a) a check 72: a file per row of section 2's table, opened in Word 365 and saved, recording the repair and what it drops; (b) take CR-064's reading | (a), cheaply. It turns "Word repairs it" into a measurement, and settles D2's premise. |
| D6 | A document that already holds a misplaced child | (a) the check runs at save on any re-marshalled part; a finding in content the caller didn't add still refuses, and the message says the opt-out; (b) the load records each misplaced child it meets, and the save refuses only new ones | (a), with the fixture gate (CR-064 D9): if the corpus of fixtures and Office-made files trips a rule, that rule reports instead of refusing. (b) needs the load to walk every part, which it doesn't do today, and identity to survive edits. |

## 5. Phases

| Phase | Work | Status |
|---|---|---|
| 0 | This CR; today's behaviour measured (section 2) | done 2026-10-26 |
| 1 | The objects package's CR (D1 (a)): the rules (D3) and the walk over a tree, with the messages (D4); its tests per rule and per container, legal children accepted | proposed there as objects CR-017 (`f2be032`, 2026-10-26; its M1 to M6 for Jason) |
| 2 | Here: the check at save, through `onMisplaced`'s save-side report (D2, refusing where a report arrives unless turned off), and on request (D6); the fixture gate over `test/fixtures/` and the Office-made corpus; tests per row of section 2's table | after phase 1 |
| 3 | Check 72 (D5), Word's repairs, if wanted | optional |

## 6. Risks

- **Refusing a save of a legal document.** Each rule refuses only what the schema never allows under that parent, and
  the gate (D6) runs the rules over every fixture first.
- **Cost at save.** The walk visits each element of each re-marshalled part once, with a set lookup per child. The
  marshal itself visits each element already, so the walk adds a fraction to it; this is measured in phase 2.
- **TypeScript users see nothing new** for what their compiler already refused. The gain is for JavaScript callers,
  for values typed loosely, and for the string in a run, whose save fails today with no cause given.
