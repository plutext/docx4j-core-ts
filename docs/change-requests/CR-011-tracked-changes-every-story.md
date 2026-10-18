# CR-011: Tracked changes in every story: one listing and filters (docx4j CR-043's `TrackedChanges.in` and `TrackedChangeFilter`)

**Status:** Proposed 2026-10-18, at Jason's request ("write a CR which covers listing every story in
one call, filters, and a report; a package-level getTrackedChanges over every story, plus an
author/type/date filter on the resolve-all verbs"). Revised the same day, at Jason's word ("yes,
revise it that way please"): trimmed to the listing and the filters, with the report left until
someone asks for it (section 7). It is to follow the 0.5.0 release, and is not scheduled. It builds
on CR-002 sections 49 to 51 (`b6ef16a`, not released). Decisions D1 to D4 (section 6) are Jason's.

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
  types?: readonly ('Added' | 'Deleted' | 'Formatted')[];
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
2. **D2, the conditions.** You named author, type and date. docx4j and python also have kinds,
   parts and a range. Recommended:
   - **Kinds and parts too.** Each is a line once `kind` and `part` exist, and with them a filter
     written for docx4j carries over. A part may be given by its name, which JSON can carry (the
     editor's review, section 9).
   - **The range later.** A content control's content or a bookmark's extent spans paragraphs, and
     this package's `Range` lies within one paragraph.
3. **D3, the kind vocabulary.** Recommended: docx4j's names verbatim (`'RUN_INSERTION'`). Then an
   object filter written for docx4j, or passed as JSON by a tool that also drives docx4j, carries over
   unchanged. These are data values, not identifiers. The alternative follows this repository's
   camelCase rule (`'runInsertion'`), as python follows its own (`run_insertion`). With the report
   left out, the case for verbatim is weaker than it was, but it still holds.
4. **D4, stale views.** docx4j's views throw once the package has been resolved since they were
   listed. This package's views aren't checked, and a caller lists again. Recommended: leave it out
   of this CR. A check would need a generation on the package, bumped by every resolution, which
   touches every view's `accept` and `reject`.

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
