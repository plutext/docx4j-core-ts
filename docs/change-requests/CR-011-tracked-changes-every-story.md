# CR-011: Tracked changes in every story: one listing and filters (docx4j CR-043's `TrackedChanges.in` and `TrackedChangeFilter`)

**Status:** Implemented 2026-10-20, all three phases (Jason: "do CR-011"; section 10): the parity on
docx4j 17.3.3 with 90 revisions goldens at zero differences, `getAllTrackedChanges`, `kind`, `part`
and `pieces`, and the filters; two resolution defects the goldens found are fixed (section 10.4). **Released in 0.6.0** (2026-10-25). Proposed 2026-10-18, at Jason's request ("write a CR which covers listing every story in
one call, filters, and a report; a package-level getTrackedChanges over every story, plus an
author/type/date filter on the resolve-all verbs"). Revised the same day, at Jason's word ("yes,
revise it that way please"): trimmed to the listing and the filters, with the report left until
someone asks for it (section 7). It came after the 0.5.0 release. It builds
on CR-002 sections 49 to 51 (`b6ef16a`, released in 0.5.0). Decisions D1 to D4 (section 6) were taken
by Jason on 2026-10-20, each as recommended: D3 with docx4j's three type names also taken by the
filter's `types`, and D4 as CR-002 section 53.

## 1. Why

docx4j 17.3.3 (tag `docx4j-17.3.3`, its CR-043) and docx4j-python (its CR-010) have three things
this package lacks:

1. **One listing of every story.** docx4j's `TrackedChanges.in(pkg)` and python's
   `revisions.tracked_changes(pkg)` list the changes in the main document, the notes, the comments,
   the headers and footers, and the text boxes in each.
2. **Filters**, on the listing and on accept and reject all: by author, type, kind, date range, part,
   and a range (a content control's content or a bookmark's extent).
3. **A report**: the listing counted by kind, author and part, with a JSON form that docx4j-mcp
   presents.

This CR takes the first two. Here:
- `Body`, `Paragraph` and `Range.getTrackedChanges()` list one story or part of one.
  `pkg.getTrackedChanges()` is the main body's, synchronous, as the shim's
  `context.document.getTrackedChanges()`.
- `await pkg.acceptAllRevisions()` and `rejectAllRevisions()` resolve every story (CR-002 section 49),
  but no call lists what they are about to resolve.
- Nothing takes a filter. The editor's `trackedChanges.acceptAll` and `rejectAll` take an `author`
  and do the filtering themselves (its `packages/editor-model/src/commands/tracking.mts`). Whether
  they move to this CR's filter is the editor's call.

**Why this much.**
- Nobody has asked for it. The case is parity with docx4j and python, which are meant to agree with
  this package, and the gap between what `acceptAllRevisions` resolves and what can be listed.
- **The listing** closes that gap. Phase 0's goldens are also worth having on their own: docx4j's
  listing as an oracle shows where the ports' grouping disagrees.
- **The filtered resolve** belongs in the engine for its loop, not its conditions. A caller can list,
  filter and accept each change. What is hard to get right is the rest:
  - re-listing after moves and broken moves (CR-002 section 29);
  - resolving only a group's matching pieces;
  - leaving out the sweep of what no change owns.
- **The report** is a few lines of counting over the listing, and nothing that uses this package
  consumes docx4j's JSON, so it waits for a consumer.
- **None of it touches Office JS's surface.** Everything new is on the package, or is an optional
  argument on `Body`'s resolve-all verbs, which are themselves an extension. Nothing goes on
  `context.document`, so code written to Office JS's shapes still runs in Word.

## 2. What docx4j and docx4j-python do (read 2026-10-18)

From docx4j's `docx4j-core/src/main/java/org/docx4j/model/revisions/` at `docx4j-17.3.3`, and
docx4j-python's `docx4j_py/model/revisions/` at `0e3c8c3`. Python follows docx4j in each point
below except where it says otherwise.

**The listing** (`TrackedChanges.in`, `Lister.stories`):
- Story by story: the main document; each footnote; each endnote; each comment; then the headers and
  footers, **in the order the sections reference them**, then any the main part relates to that no
  section references.
- A text box's changes come after the story the text box is in. Each story is in document order, and
  grouped as Office JS groups (the rules this package and docx4j share, CR-002 section 29; docx4j's
  CR-043 section 3.1 took them from here).
- A text box's changes are listed once, from its `mc:Choice`. Resolving one resolves its
  `mc:Fallback` twin too, paired by position, never by id.
- A view is valid until the next accept or reject on the package. After that, `accept()` and
  `reject()` throw `IllegalStateException` ("list again").

**The filter** (`TrackedChangeFilter`):
- Conditions: `authors(...)`, `types(...)` (`ADDED`, `DELETED`, `FORMATTED`), `kinds(...)`,
  `between(from, to)`, `parts(...)`, `within(SdtElement)` and `withinBookmark(name)`. The last two
  exclude each other, and a missing bookmark throws `IllegalArgumentException`.
- Every condition set must hold, and a condition not set holds for all.
- Authors are compared exactly. `from` is inclusive and `to` exclusive, either may be null, and a
  change with no date matches no date range.
- The conditions apply to **single changes**, the pieces of a group:
  - The listing includes a change when any of its pieces matches, and returns the whole change.
  - Resolving resolves only the matching pieces. CR-043 calls this an API choice (its U19): a kind
    filter on a deleted paragraph resolves its text and leaves its mark.
- A range is read once, from the package as it stands, before anything is resolved.
- `acceptAll(pkg, filter)` and `rejectAll(pkg, filter)`:
  - They return how many listed changes had a matching piece.
  - Each round relists and resolves last first, for up to 20 rounds; two rounds that do not shrink
    the matches end it.
  - With a filter, nothing is swept, because what the sweep takes belongs to no change (a
    `w:cellMerge`, a `w:numberingChange`, records that change nothing; this package's
    `resolveUnlisted`).
- Python takes the conditions as keywords or through docx4j's builder. Anywhere it takes a filter, a
  plain callable over a single change also works.

**The kinds** (`TrackedChange.Kind`, 21 values, and `getKind()`):

| docx4j `Kind` | This package's `target.kind` |
|---|---|
| `RUN_INSERTION`, `RUN_DELETION` | `run` with `ins`, `del`; also `mathRun` with `ins`, `del` |
| `MOVE` | `run` with `moveFrom` or `moveTo` |
| `PARAGRAPH_MARK_INSERTION`, `PARAGRAPH_MARK_DELETION` | `mark` |
| `RUN_PROPERTIES` | `runProperties` |
| `PARAGRAPH_MARK_PROPERTIES` | `markProperties` |
| `PARAGRAPH_PROPERTIES` | `paragraphProperties` |
| `ROW_INSERTION`, `ROW_DELETION` | `row` |
| `FIELD_INSERTION` | `field` |
| `CONTENT_CONTROL_INSERTION` | `holder` |
| `MATH_STRUCTURE_INSERTION`, `MATH_STRUCTURE_DELETION` | `mathStructure` |
| `TABLE_PROPERTIES` | `tableProperties` |
| `SECTION_PROPERTIES` | `sectionProperties` |
| `CELL_INSERTION`, `CELL_DELETION` | `cell` |
| a group: its pieces' kind when they are all one kind, else `INSERTION`, `DELETION` or `FORMATTING` | `group` |

Every target kind here maps to one docx4j kind, so the mapping needs no new reading of the markup.

## 3. Proposal

Names are D1's, conditions D2's, the kind vocabulary D3's.

### 3.1 Listing every story

`await pkg.getAllTrackedChanges(filter?)`, returning `Promise<TrackedChange[]>`:

- **Which stories, in what order:** docx4j's (section 2).
  - Each story's changes are its `Body`'s `getTrackedChanges()`, so the grouping is the one already
    held to Office JS.
  - A footnote, an endnote or a comment is a body of its own, as in `acceptAllRevisions`. Changes
    never group across two notes.
  - A text box's changes follow its story (CR-002 section 49's `resolveStory`).
- **The headers and footers in section order.** `acceptAllRevisions` takes all headers and then all
  footers. The resolution doesn't depend on the order, but a listing does, so the two share one
  order: docx4j's.
- **Parts read:**
  - Only story parts whose source holds revision markup, as `acceptAllRevisions` does; the others
    keep their bytes.
  - A part that holds markup is unmarshalled, because the views it returns must be resolvable. A save
    after a listing therefore re-marshals those parts. Their contents are deep-equal, but their bytes
    aren't kept.
- **`pkg.getTrackedChanges()` stays** the main body's and synchronous (D1).

On `TrackedChange`, two read-only extensions:

- `kind`: docx4j's kind (D3), by the mapping in section 2. `target.kind` stays as it is.
- `part`: the part the change's story is in: the main document part, a header or footer part, the
  footnotes, endnotes or comments part. A text box's change gives its host's part.

### 3.2 Filters

```ts
interface TrackedChangeFilter {
  authors?: readonly string[];
  types?: readonly ('Added' | 'Deleted' | 'Formatted' | 'ADDED' | 'DELETED' | 'FORMATTED')[];  // D3
  from?: Date;                              // at or after
  to?: Date;                                // before
  kinds?: readonly TrackedChangeKind[];     // D2
  parts?: readonly (Part | string)[];       // D2; a string is a part name (section 9)
}
type TrackedChangeSelection = TrackedChangeFilter | ((change: TrackedChange) => boolean);
```

- **Taken by** `getAllTrackedChanges`, `acceptAllRevisions` and `rejectAllRevisions` on the package,
  and by `Body.acceptAll` and `rejectAll`.
  - `Body`, `Paragraph` and `Range.getTrackedChanges()` take none: they return an array, and a caller
    filters it.
- **Both forms.**
  - The function, over a single change as python's callable is, covers author, type and date and any
    condition the object lacks: `acceptAllRevisions(c => c.author === 'Author B')`.
  - The object is docx4j's shape and travels as JSON, which an MCP tool's arguments need.
- **The semantics are docx4j's** (section 2):
  - every condition set holds;
  - the conditions apply to the single pieces;
  - a listed change is the whole change;
  - resolving takes only its matching pieces;
  - the count is of listed changes with a matching piece, so with a filter it is of the changes
    resolved, as `acceptAll` counts today (section 9);
  - rounds relist and resolve until no piece matches or two rounds stall;
  - nothing is swept with a filter. With none, `acceptAll` and `rejectAll` sweep as they do today
    (section 9).
- **Dates** are compared with `TrackedChange.date`. That is `w16du:dateUtc` where present, else
  `w:date` read as local wall-clock time (CR-002 section 29), the same rule as docx4j's `getDate()`.

## 4. Tests and the oracle

- **docx4j's answers as goldens**, under the parity rule in this repository's CLAUDE.md:
  - A harness class in `test/java` (`RevisionsHarness`, beside `XPathHarness`) writes the listing for
    each fixture with revision markup (90 of the 228 here). Each change is one line: type, kind,
    author, date, id, text, part name, number of pieces.
  - For the fixtures with two or more authors (checks 18, 20, 21 and 29, `revisions-word15.docx`, and
    the editor's review file, whose three authors are Author A, Author B and Claude), it also writes
    the listing left after:
    - `rejectAll(pkg, authors("Author B"))`;
    - `acceptAll(pkg, types(DELETED))`;
    - a `kinds` filter that splits a deleted paragraph (U19).
  - The goldens go in `test/golden/revisions/`, compared by `test/revisions-parity.test.mjs`.
  - The harness needs docx4j 17.3.3, where `TrackedChanges` first appears. That is beyond the pom's
    17.3.2, and the parity workflow still checks out `VERSION_17_3_3`, which is now frozen at its
    release. Moving the parity to the 17.3.3 release, as CR-001 section 22 moved it to 17.3.2, comes
    first (section 7).
  - A difference is investigated, and fixed in the port unless docx4j is plainly wrong; then it is
    reported to docx4j and the golden stays.
- **Unit tests** in `test/tracking.test.mjs`:
  - the date range on a package built in the test, with `withTracking` at three dates by three
    authors (every fixture's revisions carry one date);
  - the function form;
  - `types` in both spellings, Office JS's and docx4j's (D3);
  - a listing leaving parts without markup unread;
  - the headers and footers in section order.
- Each test fails on the code before.
- Over the 474 files of CR-002 section 49's run, a listing's count equals what `acceptAllRevisions`
  returns.

## 5. Phases and sizes

My estimates, about 2 days in all plus whatever the goldens turn up:

| Phase | What | Size |
|---|---|---|
| 0 | The parity moved to docx4j 17.3.3 (pom, workflow branch, CLAUDE.md), the harness's `MC_PREFER_CHOICE` refreshed to `UNDERSTOOD_NAMESPACES`' prefixes (CR-010 D6, section 11.5), and `RevisionsHarness` with its goldens | ½ to 1 day |
| 1 | `getAllTrackedChanges`, `kind`, `part`, the section order; the listing goldens at zero differences | ½ day, plus whatever the goldens find |
| 2 | The filters, both forms, on the listing and the resolve-all verbs; the goldens of what each filter leaves | 1 day |

## 6. Decisions (for Jason)

1. **D1, names.** Recommended: `pkg.getAllTrackedChanges(filter?)` for the listing, and an optional
   filter on `acceptAllRevisions`, `rejectAllRevisions`, `Body.acceptAll` and `rejectAll`.

   The alternative, making `pkg.getTrackedChanges()` itself cover every story, would change it from
   synchronous to a promise. The shim's `document.getTrackedChanges()` returns it, and the Office JS
   subset declares that method as `getTrackedChanges(): ArrayLike<TrackedChange>`
   (`test/office-js-subset.ts`), so the subset would no longer typecheck.

   **Decided 2026-10-20** as recommended (Jason: "D1 and D2 recommendations accepted").
2. **D2, the conditions.** You named author, type and date. docx4j and python also have kinds,
   parts and a range. Recommended:
   - **Kinds and parts too.** Each is a line once `kind` and `part` exist, and with them a filter
     written for docx4j carries over. A part may be given by its name, which JSON can carry (the
     editor's review, section 9).
   - **The range later.** A content control's content or a bookmark's extent spans paragraphs, and
     this package's `Range` lies within one paragraph.

   **Decided 2026-10-20** as recommended: author, type, date, kinds and parts; the range later.
3. **D3, the kind vocabulary.** Recommended: docx4j's names verbatim (`'RUN_INSERTION'`). Then an
   object filter written for docx4j, or passed as JSON by a tool that also drives docx4j, carries over
   unchanged. These are data values, not identifiers. The alternative follows this repository's
   camelCase rule (`'runInsertion'`), as python follows its own (`run_insertion`). With the report
   left out, the case for verbatim is weaker than it was, but it still holds.

   **Decided 2026-10-20** (Jason: "D3: ok"), with one addition. The camelCase rule is about
   identifiers; string values here are spelled as their source spells them. Four docx4j enums are
   already carried over verbatim (`AddPartBehaviour`, `PageSizePaper`, `FontDecision.Source` as
   `FontDecisionSource`, `FontFallback.FontClass`), Office JS's values keep Office JS's case
   (`'Added'`, `'TrackAll'`) and OOXML's the schema's; no docx4j enum was renamed. Verbatim is the
   fifth, not a departure. The claim above held for `kinds` but not for `types`, which follow Office
   JS's `TrackedChange.type`, while docx4j's are `ADDED`, `DELETED` and `FORMATTED`. So the filter's
   `types` also takes docx4j's three names, and a docx4j filter carries over. `TrackedChange.type`
   stays Office JS's.
4. **D4, stale views.** docx4j's views throw once the package has been resolved since they were
   listed. This package's views aren't checked, and a caller lists again. Recommended: leave it out
   of this CR. A check would need a generation on the package, bumped by every resolution, which
   touches every view's `accept` and `reject`.

   **Decided 2026-10-20** (Jason: "D4 recommendation accepted"), on a recommendation revised after
   measuring. Views that aren't checked are not harmless: for most kinds a view resolved a second
   time acts on what now stands in its place (a deleted mark accepted twice joins one paragraph too
   many; formatting rejected after `acceptAllRevisions` is restored; an accepted inserted row,
   rejected through the same view, goes). docx4j's counter is not taken either: the editor's
   `resolvePreservedOnTree` and `Body.acceptAll` resolve several views from one listing, which it
   would make throw. Instead each view acts only while the record it was listed for is still in
   place, and otherwise does nothing. That is a defect in what is released, so it is CR-002 section
   53, fixed on its own, not part of this CR.

## 7. Not in scope, and hand-offs

Not in scope:
- **The report**, until a consumer asks. For whoever picks it up, docx4j's `TrackedChangesReport`:
  - It holds `total`, `byKind` (in the `Kind` enum's order), `byAuthor` and `byPart` (both in the
    order first listed), over the unfiltered listing.
  - `byPart` is keyed by part name (`/word/document.xml`); a change without a part counts under `""`.
  - Its `toJson()` gives
    `{"total":3,"byKind":{"RUN_INSERTION":2,"TABLE_PROPERTIES":1},"byAuthor":{"Jason":3},"byPart":{"/word/document.xml":3}}`.
  - A change without an author counts under `"null"` in docx4j (`String.valueOf`) and `""` in python,
    whose `to_dict()` also has its own keys (`by_kind`).
  - Here it would be a reduce over `getAllTrackedChanges()`. It could read privately
    (`XmlPart.readContents()`), so that a save after it changes no byte.
- **The shim stays Office JS's surface.** These are extensions on the package, not on
  `context.document`.
- `acceptAllForConversion`, which serves docx4j's PDF and HTML output; this package has none.
- **Text boxes in kept `mc:AlternateContent` branches** (`mcePreprocess: false`):
  - docx4j lists the Choice's changes and resolves the Fallback twin.
  - This package's content API doesn't read kept branches yet (CR-010 phase 3).
  - What `acceptAllRevisions` does there: CR-010 section 11.2 measured that every branch's text box
    is resolved, and that the count returned counts a box's changes once per branch. Phase 1 lists
    them once, by the selected branch.

Hand-offs:
- **The registry:** `core-ts/CR-011`, proposed, through the session in `../docx4j-portfolio`.
- **The parity move to docx4j 17.3.3:** CLAUDE.md, `test/java/pom.xml` and
  `.github/workflows/parity.yml` (phase 0), with `MC_PREFER_CHOICE` refreshed in the same
  regeneration (CR-010 D6).
- **The editor**, once released: the filter exists. Its own author filter is its to keep or replace.

## 8. References

- docx4j at tag `docx4j-17.3.3`:
  - `docx4j-core/src/main/java/org/docx4j/model/revisions/`: `TrackedChanges.java`,
    `TrackedChange.java`, `TrackedChangeFilter.java`, `Lister.java` (`stories`,
    `headersAndFooters`), `Story.java`, and `TrackedChangesReport.java` for section 7;
  - `docs/developer/change-requests/CR-043-tracked-changes-accept-reject-api.md`: sections 3.1, 11
    (U19), 13 and 14.
- docx4j-python `docx4j_py/model/revisions/` at `0e3c8c3`: `__init__.py` (`tracked_changes`,
  `accept_all`, `reject_all`) and `filter.py`, with `report.py` for section 7; its
  `docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md`.
- This package:
  - CR-002 section 29 (the grouping and the dates) and sections 49 to 51 (every story, what goes with
    removed content, equations);
  - `src/packages/WordprocessingMLPackage.mts` (`revisedStories`, `holdsRevisions`);
  - `src/model/content/Body.mts` (`getTrackedChanges`, `resolveAll`, `resolveStory`).

## 9. The editor's review (2026-10-18)

The editor (`docx4j-ts-editor`) reviewed this CR with the 0.5.0 release, on main at `3d2ec61`.

**Where it would use the filter:**
- **Settling with an author.** Its tracked-changes Accept All and Reject All, when given an author,
  now list this package's changes, filter them by author, and resolve them last first
  (`preserved.mts`). `Body.acceptAll({ authors: [a] })` would replace that loop, provided it
  returns the count of changes resolved, as `acceptAll` does today: the review panel's status line
  reports that number.
- **The agent's resolve tool** (`tools-resolve.mts`) takes an author and filters the listing
  itself. A tool's arguments are JSON, so the object form is what it needs. The editor's
  `docs/agent.md` would adopt `authors`, `types`, `from` and `to` as the tool's filter when that
  contract next changes.

**On the decisions:**
- **D1: agreed.** `pkg.getTrackedChanges()` stays synchronous. The editor's console binds the shim's
  `document.getTrackedChanges()` against the Office JS subset, so a promise there would break the
  console's lint and typecheck too. `getAllTrackedChanges` with `part` is also what the editor needs
  when its review panel extends to notes, headers and comments, which it already projects as stories
  of their own but does not yet list revisions in.
- **D2: kinds and parts yes, the range later.** It adds one remark about shape (item 1 below).
- **D3: docx4j's names verbatim, as data values.** `target.kind` stays camelCase for programs. The
  editor maps the names to its own contract words (insertion, deletion, formatting, properties,
  move), as it does now.
- **D4: left out.** The editor lists again after every resolution in any case.

**The three remarks**, recorded in the proposal above:
1. **A part by its name.** `parts?: readonly Part[]` does not travel as JSON. So a part may also be
   given by its name, as `PartName.name` spells it (`'/word/footnotes.xml'`). Otherwise the tool
   layer has to map names to parts on every call. `TrackedChange.part` stays a `Part`, its name
   reachable as `part.partName`.
2. **The count** that `acceptAll` and `rejectAll` return with a filter is the number of changes
   resolved, as without one today. That agrees with section 3.2's "listed changes with a matching
   piece".
3. **No filter, a sweep.** `Body.acceptAll()` and `rejectAll()` with no filter go on sweeping as they
   do today. The editor's "all" settle depends on the sweep. The semantics, in which a listed change
   is the whole change and resolving takes only its matching pieces, suit it: it resolves its own
   listing by key, and the changes it preserves through this package.

## 10. Implemented (2026-10-20)

Jason: "do CR-011". All three phases, in one session. Uncommitted when this was written.

### 10.1 Phase 0: the parity on docx4j 17.3.3, and the revisions harness

- **The pom and the workflow.** `test/java/pom.xml` resolves the 17.3.3 release (Maven Central,
  tag `docx4j-17.3.3` at `c95b55be6`), and `.github/workflows/parity.yml` checks out
  `VERSION_17_3_4`, docx4j's development branch since 17.3.3 was released. The workflow runs the new
  harness after the XPath one, compares `test/golden/revisions/` at any depth, uploads it and takes it
  with the rest. CLAUDE.md, `test/java/README.md` and `test/golden/README.md` say so.
- **`MC_PREFER_CHOICE`** now names `READ_UNDERSTOOD_PREFIXES`, 104 prefixes, where D6 of CR-010
  (section 11.5) said `UNDERSTOOD_NAMESPACES`' 100. D6 was decided before CR-010 section 15, which gave
  the readers their own set. The parity test's walks go through `selectedBranchOf`, which reads that
  set, so docx4j is told the same set. The four namespaces section 15 adds are `w16sdtdh`,
  `w16sdtfl`, `w16du` and `x14ac`. `c173`, in the old list but not understood, is gone.
- **The goldens.** All 62 parity goldens and the 62 XPath goldens were regenerated from the 17.3.3
  release jars (a scratch Maven repository; the docx4j checkout is shared and was not built). Header
  fields aside, including `mcPreferChoice`, **no answer moved** and no XPath hit. Every header now
  records `c95b55be6`, `docx4j-core-17.3.3.jar` and its hash, `ec27210095a880f3`.
- **`RevisionsHarness`** (`test/java`, a third main class). For each of the 90 fixtures whose `word/`
  parts hold revision markup (any depth under `test/fixtures/`, `REVISION_MARKUP`'s names), it
  records:
  - the listing, each change as section 4 says, with the kinds of its pieces rather than only their
    number;
  - then, each on a fresh load, the count returned and what is left listed after five resolutions:
    `acceptAll` and `rejectAll` unfiltered, `rejectAll` by the second author listed (where the
    listing has two or more: 63 fixtures), `acceptAll(types(DELETED))`, and
    `acceptAll(kinds(RUN_DELETION))`, which splits a deleted paragraph (U19).

  It runs in UTC, which the header records, because docx4j reads a `w:date` without `w16du:dateUtc`
  in the JVM's zone. The test sets `TZ=UTC` for the same reason. A golden's path mirrors its
  fixture's, so the check folders' names never collide.
- **docx4j's head.** `VERSION_17_3_4` at `5c691fc1f` was exported with `git archive`, built into
  the scratch repository, and run through all three harnesses: the same answers on all 214 goldens.
  So the first weekly run on the new branch has nothing to report, as of that commit.
- **docx4j's commits since `2e381ef26`** up to the release touch nothing this package ports. `McSelection` gained
  `selectedBranch(ac, host)` for chartex output (`dd4a79cda`), which the READ walk does not use;
  `Mapper` changed only a comment (`1af3088b3`); `AutofitLayout` is layout (`30fe44c95`); the rest
  is CR-043's revisions package, the oracle itself.

### 10.2 Phase 1: the listing

- `await pkg.getAllTrackedChanges(filter?)`. Its stories are `revisedStories()`, the same ones
  `acceptAllRevisions` resolves. Each story's changes are its `Body.getTrackedChanges()`, then each
  selected-branch text box's in it (`trackedChangesOfStory`, beside `resolveStory`).
- **Headers and footers in section order** (`inSectionOrder`), docx4j's `Lister.headersAndFooters`:
  - the `w:sectPr`s in document order, not inside a run, the body's last;
  - each one's references in their order, a part once;
  - then any other header or footer, in the order of the main part's relationships.

  This is worked out only when two or more headers or footers hold markup, since otherwise there is
  nothing to order. The main part's tree is read with `readContents()`, so a main part with no
  revision keeps its bytes. `acceptAllRevisions` resolves in the same order.
- **`TrackedChange.kind`** is the table in section 2, a group's as docx4j's `getKind()`.
  **`TrackedChange.part`** is the part of the change's body. **`TrackedChange.pieces`** is docx4j's
  `getPieces()`: a group's pieces, else the change itself. `pieces` is one more than section 3.1
  named. The filter's semantics are stated over pieces, and the function form is given pieces, so a
  caller needs to see them.
- **Every change now knows its body.** A table's records, the body's last section and a block-level
  inserted control were made with no paragraph and no body, so they had no `part`. So was a group
  whose first piece is one of those. `tablePropertiesChangeOf`, `bodySectionChangeOf`,
  `holderInsertionAt` and `groupTouching` now pass the body. That also gives such a change, resolved on
  its own, the bookmark and relationship bookkeeping of `resolving`, which inside a resolve-all it
  already had from the outer resolution.
- **The listing goldens agreed at once: 0 differences on all 90.**

### 10.3 Phase 2: the filters

- `src/model/content/trackedChangeFilter.mts`: `TrackedChangeFilter`, `TrackedChangeSelection`, and
  `pieceMatcher`, which turns either form into a test over one piece. The conditions, both
  spellings of `types` and the semantics are section 3.2's. Two departures from the sketch there:
  - **`from` and `to` also take a string**, read by `Date`, since a tool's JSON arguments carry no
    `Date`.
  - **An object filter is checked before use.** An unknown key, a value not in the vocabulary, or a
    date that does not parse throws a `TypeError`, and nothing is read or resolved. docx4j's builder
    cannot be misspelt, but a JSON object can, and `{ author: 'Author B' }` silently ignored would
    accept every change. `within` and `withinBookmark` are refused the same way, since they come
    later (D2).
- **Taken by** `getAllTrackedChanges`, `acceptAllRevisions`, `rejectAllRevisions`, `Body.acceptAll`
  and `Body.rejectAll`. `resolveStory` passes the filter to each text box's story. The count is of
  the listed changes with a matching piece, the selected branch's in a kept `mc:AlternateContent`, as
  before.
- **The loop is docx4j's, filtered or not.** A change whose pieces all match is resolved whole, and
  one with only some matching has those resolved alone (`resolvePieces`, now exported). Passes repeat
  up to 20 times, and two passes in a row that leave as many matching changes as before end it. Until
  now one such pass ended it. The unfiltered sweep (`resolveUnlisted`) runs only without a filter.
  Measured on the corpus below, the loop changed no result.

### 10.4 What the goldens found, and fixed

1. **A deleted paragraph mark did not join across a range marker.** `joinWithNext` joined only when
   the very next item was a paragraph. A `w:moveFromRangeEnd`, a `w:bookmarkEnd` or a `w:permEnd`
   between the two made it only drop the mark, and the paragraph stayed, empty.

   docx4j's join skips such markers (`*RangeStart`, `*RangeEnd`, `bookmarkStart`, `bookmarkEnd`,
   `permStart`, `permEnd`, `proofErr`) and carries them into the joined paragraph where they stood.
   That is its CR-043 rule, from Word's saves C10 and C10b. It is ported.

   - **Why the unfiltered goldens missed it.** After an unfiltered accept nothing is left to list, so
     those goldens could not show it. `acceptAll(types(DELETED))` on check 20's 13-moveTo-no-ranges
     files left the section's `w:sectPrChange`, whose text, every paragraph of the section, carried the
     empty paragraph.
   - **What moved.** Over the corpus below, Accept All or Reject All changed on 11 files, all check
     20's broken moves (case 12's resave, case 13 three times, case 14 seven times), and nothing else.
   - **They now agree with docx4j** (section 10.5); before, all 11 differed. The single-change tests
     against Word's own saves of those cases (check 20) pass as before.
2. **A note whose reference a resolution took was still a story until save.**
   `acceptAll(kinds(RUN_DELETION))` on the CR-043 probe takes C12's endnote reference. This package
   settles the note itself on save (CR-002 section 44); docx4j settles it when the resolve-all ends.
   So this package went on listing the note's deleted mark.

   Now `revisedStories()` leaves out a footnote or endnote that `notesGoing` says will go, meaning
   pending and with no reference standing, and a comment pending removal. That covers the listing
   and the resolve-alls alike. When anything is removed is unchanged.

After both, **the revisions goldens are at 0 differences on all 90**: the listing, five
resolutions each, every count.

### 10.5 Checks beyond the goldens

- **The ports' agreement run, repeated** (one-off, scratchpad). docx4j's own `PortsAgreement` and
  `DocumentDescription` (its CR-043 section 14, copied from `docx4j-core-tests` at the 17.3.3 tag)
  compared docx4j's Accept All and Reject All of the 90 fixtures with this package's saved results.
  - With this build, **179 of 180 describe the same**. The one is the CR-043 probe's Reject All, its
    C23 table grid, which is CR-002 section 46.2, recorded for later.
  - The build before differed on those 11 broken-move files as well.
- **The corpus**: the 520 `.docx` of this repository's fixtures and docx4j-core-tests' resources.
  - On 519, the listing's count equals what `acceptAllRevisions` returns and what
    `rejectAllRevisions` returns. The 520th is the strict-conformance sample, which does not load.
  - 191 hold changes, 3,698 in all.
  - On the 328 with none, a listing unmarshals no part.
  - The all-verbs hash comparison of CR-002 section 53, against a baseline from before this CR: only
    the 11 files of item 1 changed. With the marker skip disabled, the rest of this CR changed none.

### 10.6 Tests

- `test/revisions-parity.test.mjs`: 91 tests, one per golden and one over their headers.
- `test/tracking.test.mjs`: 8 more tests. Each fails on the build before; the 77 already there pass
  on both.
  - date ranges (`from` inclusive, `to` exclusive, an ISO string, no date matching no range), over
    three authors' insertions made with `withTracking` at three dates;
  - the function form, `types` in both spellings, `kinds` and `parts` by name and by part;
  - a misread filter throwing before anything is resolved;
  - `kind`, `part` and `pieces`, and U19 on check 22;
  - no sweep with a filter;
  - the join across a bookmark end and a range end;
  - section order, with the main part and a footer without markup left unread;
  - a note pending removal not listed, and gone on save.
- The suite: 977 tests, all passing; the typecheck passes.

### 10.7 Not done

- The range conditions (D2: later) and the report (section 7).
- The editor's own author filter: whether to move to this one is the editor's call (section 9).
