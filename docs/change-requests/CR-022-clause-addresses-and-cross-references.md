# CR-022: Clause addresses, and typed mentions made Word cross-references (the port of docx4j's S2; S5's engine half)

**Status:** Proposed 2026-11-03, from Jason's rulings of the same day, relayed by the portfolio session (docx4j-portfolio-10;
portfolio `8e6a8fc`). The rulings are in `mcp_strategy.md` section 12.3 (S1 to S6, as recommended, with three refinements)
and in the address grammar, `mcp_tool_vocabulary.md` section 3 (the clause form). docx4j's S2 is its CR-074 (`51d81176a`, the
17.3.6 line; registry `1078d27`). Its phase 0 is a Word probe staged for Jason: 655 `REF` fields with `\w`, `\n`, `\r`
and `\t` over 79 numbered paragraphs in twelve list shapes, plus Word's own Insert > Cross-reference. Phase 1 here, the
port of S2, waits on docx4j's CR-074.2 (the clause and the lookup); phase 2, S5's engine half, waits on its CR-074.3 (the
cross-reference writer, held to Word's insertions). D1 to D6 for Jason (section 5). docx4j-python ports S2
under its own entry; ts-editor (S4, and S5's function) and docx4j-mcp (S3) consume what this CR builds.

## 0. Summary

Lawyers cite clauses by number. Word computes those numbers when it draws the page, so they appear nowhere in the file,
and an agent that cannot see them cannot find "2.01(a)". Jason ruled that a clause becomes an address
(`clause:11.1`, `clause:2.01(a)`, `clause:"Article III"`), and that **resolving it belongs in the engines**: docx4j
first (S2), then this package and docx4j-python, each to docx4j's goldens, so one document gives the same clause in Java,
TypeScript and Python. ts-editor and docx4j-mcp consume the result.

This CR has two phases:

1. **The port of S2.** For every numbered paragraph, a full-context label from the numbering model: the labels of the
   levels above joined as Word's "paragraph number (full context)" joins them, so `2.01(a)`, not `(a)`.
   - It is read-only, on `Paragraph`, and in every address the content API returns.
   - `clause:` joins `Address`: it resolves to the paragraph carrying the label, and is an error naming the candidates
     when none or several carry it.
   - As a scope it covers that paragraph and those after it until the next label at its level or above.
2. **S5's engine half.** A typed mention ("see clause 4.2", "Section 11.1(b)") is resolved to its paragraph, and a real
   Word cross-reference is put there:
   - a bookmark on the numbered paragraph;
   - a `REF` field showing its paragraph number, in the context matching how the text cited it (full or relative).
   This gives ts-editor's "manual to real cross-references" function an engine to call. docx4j is the oracle for the
   field's switches and the bookmark's shape; Word's own insertion is the golden. Drift warnings after renumbering stay
   later.

## 1. What this package has today (2026-11-03)

- **Labels, as Office JS gives them.** These come from the numbering `Emulator`, held to docx4j's parity goldens (list
  labels and counters, 62 fixtures):
  - `ListItem.listString`: the label Word paints, which is the level's own `w:lvlText` with its counters filled in,
    so `(a)` for a level whose text is `(%2)`;
  - `ListItem.level`, `getAncestor`, `getDescendants` and `siblingIndex`;
  - `List.getLevelString`;
  - `Body.listLabels()`, every label of a story in one walk;
  - the text model's `list_labels` and per-block `numbering` (REQ-510: `numId`, `level`, `format`, `levelText`,
    `start`, `value`).
  None of these is the full-context form: a paragraph labelled `(a)` under `2.01` reads `(a)`.
- **Addresses** (`Address`, `Body.mts`): the ordinal (`body/3`, `header:rId5/0`, ...), `{ paraId }` and
  `{ contains }`. They are resolved by `pkg.paragraphAt` and `Body.paragraphAt`, which return `undefined` where the
  address names nothing. `outline()`'s entries carry `address` and `paraId`.
- **Scopes.** CR-014's `Range` across paragraphs is the natural scope: `Body.getRange` and `expandTo` make one, and
  `search` runs within it.
- **Bookmarks and fields.**
  - Bookmarks are kept and moved with their content, but there is no API to make or find one. Office JS's
    `Range.insertBookmark`, `Range.getBookmarks` and `Document.getBookmarkRange` (WordApi 1.4) are not in the subset.
  - Fields are read: their stored results (CR-012), and Office JS's text (CR-002 section 62). Nothing inserts one;
    Office JS's `Range.insertField` (WordApi 1.5, `Word.FieldType.ref` among its types) is not in the subset.
  - There is no field evaluator: docx4j's CR-067 evaluates `REF`, and this package reads the stored result.
- **The label is never text here.** `Paragraph.text`, `search` and the writers see no list label, as Office JS's
  don't. So the vocabulary's rule ("an edit that would change the text of a computed label is refused") holds without a
  refusal of its own: no write can reach a label.

### 1.1 docx4j's probe, read here (2026-11-03)

docx4j CR-074's phase 0 is done (`af8c991ba`; portfolio `9dda5ea`). Jason ran the probe, and every one of its 655 `REF`
fields was updated. Word's saves and the run's log (`clause-numbers-log.txt`: each paragraph's drawn label, every
field's code and result, the three views and Word's insertions) are docx4j's goldens in
`docx4j-core-tests/src/test/resources/clause-numbers/`. Its `README.md` states the rules: full, relative and no
context, `\t`, `\p`, `\d`, the accepted view, and Word's insertion.

This package's labels (`Body.listLabels()`, so `ListItem.listString`), read on `cn-1-labels.docx`, Word's save of it and
`cn-2-insert.docx`, against Word's `ListFormat.ListString` in the log (all markup shown):

- **77 of the 79 equal Word's**, every shape included: the legal `%1.%2`, `w:isLgl`, style-linked levels, a list
  style, skipped levels, a restarted list, a level of words, `Step 1:`, a bullet, `NOTE` and an empty level.
  Phase 1's `ParagraphNumber` therefore starts from labels that are Word's.
- **T2 and T4 differ.** T2 is a numbered paragraph whose content and mark are tracked as deleted.
  - Word counts as if the deletion were accepted, even with all markup shown. It draws T2 `1.`, the counter as it
    stands, not advanced, and T4 `2.`.
  - Here `listString` counts T2: T2 `2.`, T4 `3.`.
  - The text model is the same in its `markup` and `original` views (T4 `3.`). Its `accepted` view gives Word's `2.`.

  So phase 1 counts in the accepted view, as docx4j's D6 now says (Word's `REF` results are the accepted view's in all
  three views). The Office JS reader's difference is a finding for CR-002 (section 8). The views' labels are the text
  goldens', docx4j's to decide (reported to it through the portfolio).

## 2. Phase 1: the port of S2 (waits on docx4j's CR-074.2 and its goldens)

docx4j's CR-074 (sections 4.1 to 4.3, phases 1 and 2) gives the shapes. Its D1 to D10 are open with Jason, and its rules
come from its phase 0 probe; this port follows what it decides.

- **`ParagraphNumber`** (docx4j section 4.1, in `src/model/listnumbering/`). This is a snapshot taken when the
  `Emulator` counts a paragraph: the list (`w:abstractNum`), the `w:num`, the level, the counters of levels 0 to that
  level, and the levels' definitions.
  - It answers `label()` (REQ-510's, no suffix), `noContext()` (`\n`), `fullContext()` (`\w`) and
    `relativeTo(from)` (`\r`, `from` undefined for a field outside any list), each also in the `\t` form.
  - The `Emulator`'s `NumberingResult` gains it, so every consumer counting through the `Emulator` has it.
  - It is held to Word's `REF` results in the probe's `cn-1-labels-word365.docx` (655 fields), at zero differences.
- **The clause.** It is the full-context label: exactly what Word's `\w` shows for the paragraph (docx4j's D1), with no
  suffix (its D10). Only a paragraph whose level shows a number carries one (its D5, after phase 0): Word's `\w` gives a bullet's
  character, `NOTE` for a level of format none, and `0` for an empty level or a paragraph that isn't numbered, none of
  them a clause.
  - It is on the text model's `numbering`, docx4j's `TextNumbering.getClause()` (its section 4.2).
  - It reaches the JSON output's `numbering` object as the WD writes it (its D9).
  - Plain text and Markdown don't change: `list_labels` keeps the level's own label.
- **`Clauses`** (docx4j section 4.3, in `src/model/text/`).
  - `Clauses.of(result)` lists every clause: its label, block, address, level and list.
  - `find(label)` returns one clause, or throws `ClauseLookupException`.
  - Matching is docx4j's D3: an exact match on the clause; failing that, on the `\t` form (`3.01` for `Section 3.01`)
    where exactly one paragraph matches. The query is trimmed and one trailing period dropped. There is no case
    folding.
  - The exception names the candidates and their addresses where several match, or the near misses where none does.
    Its message is written for an agent, as docx4j-mcp passes it through.
  - `Clause.scope()` is every block, tables included, from the clause's paragraph to the next paragraph of the same
    list at its level or above, or to the end of the story (its D4).
  - It covers the main document's blocks as the text model walks them: no text boxes, headers or notes (its D6).
- **The view (Jason's rule V1, 3 November 2026; portfolio `e314871`, `docs/clause_numbers.md` section 2).** "Semantics of
  'find clause 10.3' when there are tracked changes must depend on the specified view (as if accepted, rejected etc)."
  - A clause belongs to a paragraph under a revision view. Every route resolves under a stated view, accepted by
    default: the `clause:` address, `Paragraph.clause`, `Clauses.find` and the scope.
  - A hit carries the view it was resolved under.
  - `10.3` can name different paragraphs under `accepted` and `original` where a numbered paragraph's mark is tracked
    as deleted or inserted.
  - The `markup` view counts as accepted (WD 0.64, REQ-510; Word's `REF` results are the accepted view's in every view,
    case T). How the `original` view counts is open with Jason.
  - How the view is passed is docx4j CR-074's D11 (`b0e256ed1`; portfolio `2c5d947`), copied here:
    - the view is the `TextResult`'s: `Clauses.of(result)` counts in the view the result was extracted under
      (`TextOptions.view`, `accepted` by default), and `Clauses.of(pkg, view)` is the shorthand;
    - every `Clause` has its `view`;
    - `ClauseLookupException` has its `view`, and its message names it ("no paragraph carries 10.3 under the accepted
      view; candidates ...");
    - the text model's `numbering.clause` is the block's clause in its result's view, with no argument of its own;
    - one label can name different paragraphs in different views; no lookup crosses views, and nothing guesses.
  - The content API takes the view as the text model does:
    - `pkg.paragraphAt('clause:10.3', { view })` and the scope's verb take `{ view }`, `accepted` by default;
    - `Paragraph.clause` is the accepted view's;
    - `Paragraph.getClause(view)` gives another view's.
    These are names for D2.
- **The content API, over `Clauses`** (this package's part; docx4j has no content API). Every value here is read-only,
  and `undefined` where the paragraph has no clause.
  - `Paragraph.clause` and `ListItem.clause` are extensions, since Office JS has no such member.
  - `OutlineParagraph.clause` comes free, since `outline()` is the text model's (CR-012 D3).
  - `clause:<label>` is a string `Address` (the label quoted where it holds spaces or quotes), which `pkg.paragraphAt`
    and `Body.paragraphAt` resolve through `Clauses.find` in the main document. An error is thrown where the other forms
    give `undefined` (D3).
  - The scope becomes a CR-014 `Range` from the clause's paragraph to the end of the scope's last block, and `search`
    runs within it.
  - Where docx4j's scope ends inside a table, CR-014's rules for a range ending in a table apply; how they meet
    docx4j's D4 is checked when its goldens exist.
- **Every address returned carries the clause** where the paragraph has one: outline entries, and search hits
  (`Range.paragraph.clause`).
- **Oracle.** docx4j's probe files and Word's saves (`docx4j-core-tests/src/test/resources/clause-numbers/`) are
  copied unedited with a SHA-1 manifest. docx4j's answers for them (the clauses, the lookups, the scopes) are written by
  the parity harness or copied as goldens, at zero differences (CLAUDE.md's parity rule).

## 3. Phase 2: S5's engine half (waits on phase 1 and the cross-reference probe)

- **Office JS's members, into the subset** (D4). They are assignable to Office JS, and the shim gets them too:
  - `Range.insertBookmark(name)` (WordApi 1.4): a bookmark of the same name elsewhere goes first. The name is checked
    as Office JS documents it: letters, digits and underscores, a letter first unless hidden (`_`), at most 40
    characters.
  - `Range.getBookmarks(includeHidden?, includeAdjacent?)` and `Document.getBookmarkRange(name)` (`ItemNotFound` where
    there is none).
  - `Range.insertField(location, 'Ref', text)` (WordApi 1.5), returning a `Word.Field` with `code` and `result`. `Ref`
    comes first; other types are refused (`NotImplemented`) until they are wanted.
- **docx4j's writer, ported.** `CrossReferences.toParagraphNumber(target, where, context, hyperlink)` (docx4j CR-074
  section 4.5, its D8, phase 3). It adds the hidden `_Ref` bookmark to the target, or reuses one, and writes the `REF`
  field with Word's switches and its result, as `cn-2-insert-word365.docx` shows Word writing them. Here it writes
  through the content API, so a write under tracking is tracked.
- **The engine function** (an extension; the names are agreed with the editor; it calls the writer):
  - `resolveMention(range)`: finds the label a typed mention cites, as phase 1's addresses resolve it, with the
    candidates where it is ambiguous. stetdoc's states (resolved, ambiguous, relative, outside, unresolved; its
    decision 0015) are the reference, but the states are the editor's and docx4j's to settle.
  - `insertCrossReference(range, paragraph, { context, hyperlink })` replaces the mention's number with a cross-reference:
    - a hidden `_Ref` bookmark on the numbered paragraph, reused where one already marks it, as Word reuses its own;
    - a `REF` field naming it, with `\w` (full context) or `\r` (relative context), as the text cited it ("clause
      4.2(a)" in full, "(a)" within 4.2 relative), and `\h` for a hyperlink.
- **How the editor calls them** (ts-editor ED-012 section 2, docx4j-ts-editor-16, 2026-11-03; the names are agreed):
  - `resolveMention(range, context?)` takes the mention's range: its paragraph and the label token's offsets in
    `Paragraph.text`, as the editor's find reads them. It answers one of stetdoc's five outcomes: `resolved`,
    `ambiguous` (with its candidates), `relative`, `outside` or `unresolved`. Each candidate gives its paragraph's address
    and full-context label.
  - The editor runs it over a `readTree` copy for every mention in one pass, so the labels must not be recounted from the
    start on every call: 255 pages would be too slow. It therefore takes a context made once, as `TextContext` is: the
    stories' labels counted once and an index from label to paragraphs. It is reused across calls, and made again only
    after an edit renumbers.
  - `insertCrossReference(range, candidate, { context })` takes the same range, the chosen candidate and, where the typed
    form decides it, full or relative context. It inserts the bookmark and the `REF` field through the content API, the
    field's stored result the label as computed now.
  - The editor calls it inside `runOnTree`, in reverse document order, so earlier offsets stay valid. Under change
    tracking it inserts through the paragraph primitives, so the field's runs are tracked, as every write here is.
    Whether Word's insertion under tracking also marks the replaced mention deleted is a case for the probe.
  - The editor never calls the four D4 primitives itself; it gets them only as the subset the console's Word shim offers.
- **The field's stored result** is phase 1's `ParagraphNumber` in that context, so every reader shows the right text
  without Word updating the field (D5). Word's own `REF` results in docx4j's probe (`\w`, `\n`, `\r` and `\t`, which
  docx4j-python asked for) are its goldens, as is docx4j's evaluator (its section 4.4, CR-067's).
- **Goldens.**
  - **Word's own insertion:** Insert > Cross-reference > Numbered item, "Paragraph number (full context)" and
    "(relative context)", "Insert as hyperlink", saved, with Track Changes off and again with it on. This is the probe asked of docx4j and Jason. It gives the
    bookmark's name form, what it spans, the field code's switches and spacing, and the stored result. Phase 2's output
    is held to it, region for region, as checks 52 to 57 hold writes.
  - **Check 74** (D6): a Script Lab snippet doing the same through Office JS's `insertBookmark` and `insertField` on
    Word, reading back `field.code` and `field.result`. It is held to the shim's run of the same snippet, as checks 52
    to 73 are.
  - **Word's `REF` results**: docx4j-python's request, taken by docx4j CR-074 (its D7). The probe's
    `cn-1-labels-word365.docx` holds Word's result under `\w`, `\n`, `\r` and `\t` for each of 79 numbered
    paragraphs, so one oracle covers both the label and the field, and each stored result has a golden beside Word's
    insertion (`cn-2-insert-word365.docx`).
- **Out:** a drift warning when an edit renumbers what a cross-reference or a typed mention cites (S5's rest, later);
  `PAGEREF`; a reference to a heading or bookmark by text rather than number.

## 4. Tests

- **Phase 1:**
  - docx4j's S2 goldens at zero differences;
  - the address resolved, with none and several refused and the candidates named;
  - the scope's ends, including a restarted list and a list across a table;
  - every story;
  - the outline's and the hits' clauses;
  - `Paragraph.text` unchanged;
  - an untouched save byte for byte after the reads.
- **Phase 2:**
  - the probe's saves held to this package's insertion, region for region;
  - check 74's run held to the shim's;
  - the bookmark reused;
  - the stored result per switch held to docx4j's evaluator;
  - the inserted field read back by the text model (its stored result) and by `Paragraph.text` (Office JS's, check 57's
    rules).

## 5. Decisions (for Jason)

| # | Question | Options | Recommendation |
|---|---|---|---|
| D1 | S5's engine half: a phase of this CR, or a CR-002 section | (a) phase 2 here; (b) a CR-002 section | (a). It is built on phase 1's labels, and the portfolio left it to this session. The Office JS members it adds are recorded in CR-002's subset list when they land. |
| D2 | Names | (a) docx4j CR-074's names, camelCased (`ParagraphNumber`, `TextNumbering`'s `clause`, `Clauses`, `ClauseLookupException`, `Clause.scope`, `CrossReferences.toParagraphNumber`), with this package's own only for the content API's (`Paragraph.clause`, the `clause:` address); (b) this package's own throughout | (a), as CLAUDE.md's naming rule. |
| D3 | A `clause:` address naming no paragraph, or several | (a) an error naming the candidates, as the vocabulary rules; (b) `undefined` for none, as the other forms give, and an error for several | (a). The grammar says so, and the editor and docx4j-mcp would otherwise each tell the two cases apart themselves. |
| D4 | Office JS's bookmark and field members | (a) `insertBookmark`, `getBookmarks`, `getBookmarkRange` and `insertField` for `Ref` only, into the subset; (b) the extension function only | (a). An add-in's code then runs on both sides, CR-002's premise. |
| D5 | The inserted `REF` field's stored result | (a) computed here from phase 1's label; (b) left for Word to compute (`w:dirty`, or `updateFields` in the settings) | (a), held to Word's insertion. A document read before Word opens it then shows the number. |
| D6 | Check 74 (Office JS's own insertion on Word) | (a) yes, once the probe's file exists; (b) the probe's saves only | (a). It holds the shim's `insertField` and `insertBookmark` to Word's, which the probe's dialog run cannot. |

## 6. Phases

- **Phase 1**, the port of S2: `ParagraphNumber`, the clause, `Clauses`, and the content API's address and scope.
  Depends on docx4j's CR-074.2, phases 1 and 2 there, and their Word probe, CR-074.0.
- **Phase 2**, S5's engine half: the Office JS members, the writer, and the engine function. Depends on phase 1 and on
  docx4j's CR-074.3, the writer held to Word's insertion.

## 7. Registry

Sent to the portfolio session:
- `core-ts/CR-022`, proposed;
- `core-ts/CR-022.1`, planned, depending on `docx4j/CR-074.2`;
- `core-ts/CR-022.2`, planned, depending on `core-ts/CR-022.1` and `docx4j/CR-074.3`.

The portfolio applied these as `53674ec`, and the docx4j dependencies as `1078d27`.

## 8. Findings beyond this CR (recorded, nothing changed)

- **`ListItem.listString` counts a paragraph tracked as deleted** (section 1.1, case T of docx4j's probe). Word, with
  all markup shown:
  - doesn't advance the counter for a numbered paragraph whose mark is tracked as deleted;
  - draws that paragraph with the counter as it stands (T2 `1.`);
  - numbers the paragraphs after it as accepted (T4 `2.`).
  `listString` and `Body.listLabels()` give T2 `2.` and T4 `3.`. This is a small CR-002 change, for Jason: the
  `Emulator` counted in the accepted view for the Office JS reader.
  - Measured through VBA's `ListFormat.ListString`, not through Office JS's `listString`. A Script Lab read of
    `paragraphs.items[i].listItemOrNullObject.listString` on `cn-1-labels-word365.docx` would confirm the Office JS
    side.
  - None of the 62 parity fixtures (list labels and counters) has a paragraph numbered by its own `w:numPr` whose mark
    is tracked as deleted (scanned 2026-11-03), so the goldens don't say what docx4j's `Emulator` does there. That is to
    be checked before the change.
- **The text model's `markup` view, by WD 0.64** (portfolio `6b8ffd9`, REQ-510, from this finding). In every revision
  view labels count as accepted: a numbered paragraph whose mark is tracked as deleted doesn't advance the count and
  shows the label the counter stands at, and the next numbered paragraph takes the accepted view's number.
  - docx4j's text model (17.3.6-SNAPSHOT) reads case T as this package's does (markup T2 `2.`, T4 `3.`), so the WD's
    note records it as a finding for both engines.
  - This package's text model follows docx4j's text goldens, so the change lands here when docx4j's does, as a CR-012
    item.
  - The `original` view is marked unmeasured. Under Review > Original, docx4j read COM's `ListString` as not
    recomputed (T4 `1.`), and a look by hand is open with Jason.

## 9. Sources

- The portfolio:
  - `docs/mcp_strategy.md` section 12 (stetdoc; section 12.3, the rulings);
  - `docs/mcp_tool_vocabulary.md` section 3 (the clause form).
- stetdoc (`clauseo-ai/stetdoc` 0.9.2): `docs/concepts.md` ("Clause numbers"), decisions 0012 (several clauses as a union)
  and 0015 (typed mentions). Read by the portfolio session, not here.
- docx4j: `docs/developer/change-requests/CR-074-clause-numbers-full-context-labels-and-lookup.md` (`51d81176a`); its probe in
  `docx4j-core-tests/src/test/resources/clause-numbers/`.
- docx4j's phase 0: `af8c991ba`; `clause-numbers/README.md` (Results) and `clause-numbers-log.txt`.
- Office JS (`@types/office-js` 1.0.589): `Word.Range.insertBookmark` and `getBookmarks` (WordApi 1.4),
  `Word.Document.getBookmarkRange` (1.4), `Word.Range.insertField` (1.5), `Word.FieldType.ref`.
- Here:
  - `src/model/content/List.mts` (`listString`, `getAncestor`);
  - `src/model/content/Body.mts` (`Address`, `paragraphAt`, `listLabels`, `OutlineParagraph`);
  - `src/packages/WordprocessingMLPackage.mts` (`paragraphAt`, `outline`);
  - CR-014 (ranges across paragraphs), CR-012 (REQ-510's labels), CR-002 section 62 (fields read as Office JS reads them).
