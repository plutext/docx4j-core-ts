# CR-019: Manual redline to tracked changes, an explicit, opt-in conversion (the port of docx4j CR-068)

**Status:** Proposed 2026-10-31, from docx4j CR-068's phase 3 hand-off (docx4j-12, at Jason's "do CR-068 phase 3";
docx4j `VERSION_17_3_5` at `8a98c78e7`, unpushed). Nothing implemented; the port waits for Jason's go. Decisions D1 to
D8 for Jason in section 6, each with a recommendation. Registered with the portfolio as `core-ts/CR-019`, depending on
`docx4j/CR-068`.

## 0. Summary

A "manual redline" is a document whose author struck the old words through and underlined (or coloured) the new
ones by hand instead of turning Track Changes on. The file holds formatting, not revisions: Accept All and Reject All
leave the struck text, every text view shows it, and a plain-text reader prints the old number against the new
(`1.316 1.335`; in a cell with no space, `1.3161.335`). That is right for an extraction, and Jason ruled the handling is
not a view option (docx4j CR-068 section 1).

docx4j CR-068 built the operation a reviewer runs on purpose: `ManualRedline.toTrackedChanges(pkg, options)` and
`preview`, in `org.docx4j.model.revisions` beside CR-043's `TrackedChanges`. By default every run struck through by its
own `w:strike` or `w:dstrike` becomes a `w:del` by the author `manual redline`, the strike removed so that the deletion
alone carries the meaning; adjacent runs share one `w:del`; the space that separated the old words from the new goes
with them; a struck paragraph mark goes where Word's own deletion would take it; tracked content, fields, math and
style-only strike are left and reported. Underline and named colours as insertions, style strike, the comments story, a
range and an audit property are opt-ins. Word 365's Accept All and Reject All of the conversions are the oracle, three
runs, every prediction held (its sections 8.4, 11.6, 12.6).

This CR ports it here as CR-011 ported CR-043: the same operation, options and report, in this package's shapes, held
to the same Word saves by the same structural comparison. The WD takes no requirement (the conversion is not
extraction), so the text model is untouched.

## 1. What exists here

- **CR-011** (`pkg.getAllTrackedChanges`, `acceptAllRevisions`, `rejectAllRevisions`, `TrackedChangeFilter`):
  `revisedStories()` enumerates every story as docx4j's `Story` does (main, notes, comments, headers and footers, text
  boxes in every `mc` branch; `Listing` pairs twins); `trackedChangeFilter.mts` has `authors`, `types`, `from`, `to`,
  `kinds`, `parts`, but no `within` or `withinBookmark` (CR-011 D2 left the range filters out; section 7 there).
- **The change tracker** (`tracking.mts`, CR-002 sections 29 and 35): `ChangeTracker.del(items)` and `ins(items)` make a
  `w:del` or `w:ins` with `nextId()` above `highestAnnotationId` of the part and the stamp `markup()`; `toDeletedText`
  turns `w:t` into `w:delText`; `markParagraphDeleted`; `declareIgnorable(part, 'w16du')`; a stamp is `w:date` the wall
  clock in `pkg.effectiveDateTimeZone()` with a `Z` and `dateUtc` the instant, to the minute (CR-018), which is what
  CR-068's D4 as clarified writes (to the second there; D2 below).
- **The text model** (CR-012): each `TextBlock` carries `address` and `source`, the paragraph object it was read from,
  so a map from paragraph to address is one pass over `Texts.of(pkg)`; nothing public gives it yet (D3). Strike is
  resolved through the `PropertyResolver` (`getEffectiveRPr(rPr, pPr)`) for `TextRun.isStrike`.
- **Parts**: `DocPropsCustomPart` exists (typed `custom.Properties`), with no verb to set a property (D5).
- **Word's saves as oracles**: `test/s17/`, `test/check52/`, `test/check54/` and the rest are compared region by region
  through test-local describers (`regionsOf` in `ranges.test.mjs`); nothing here is yet docx4j's `DocumentDescription`
  (its `docx4j-core-tests` `org.docx4j.model.revisions.DocumentDescription`: each story's paragraphs with style, indent,
  list, the mark's revision, runs merged to one formatting and one revision, tables, notes, comments sorted, text boxes,
  equations, bookmarks, sections' effective headers and footers; ids, dates, rsids, field codes and Word's written
  defaults left out). D6.

## 2. Today (nothing to measure)

No operation here turns a strike into a deletion. A caller can do it by hand through the content API (`Font.strikeThrough`
reads, `Range.delete()` under `changeTrackingMode = 'TrackAll'` writes a `w:del` by the tracker's author), one range at a
time, without D8's grouping, D13's spaces or D9's marks, and with the strike kept on the deleted run (Word's own shape,
CR-068 section 8.4 item 6), so that Reject All gives struck text back. The three views, `Paragraph.text` and the text model
all show struck text as text, which stays so (Jason, CR-068 section 1).

## 3. Proposal

### 3.1 The API

In `src/model/content/manualRedline.mts`, exported from `./model` and the root (D1):

```ts
import { ManualRedline } from '@docx4j/core-ts';

const report = await ManualRedline.toTrackedChanges(pkg);                 // strike only, every story but comments
const report = await ManualRedline.toTrackedChanges(pkg, {
  author: 'manual redline',            // default
  date: new Date(),                    // default: now; null writes no date
  separatorSpace: true,                // default (D13)
  styleStrike: true,                   // strike inherited from a style converts too; default false
  underlineAsInsertion: true,          // own w:u (a pattern other than none) becomes w:ins; default false
  colourAsInsertion: ['FF0000'],       // runs of these colours (RRGGBB) become w:ins; default none
  comments: true,                      // the comments story too; default false
  within: control, withinBookmark: 'Schedule2', parts: [part, '/word/footnotes.xml'],   // a range
  auditProperty: true,                 // a custom document property recording the run; default false
});
const report = await ManualRedline.preview(pkg, options);                 // the same report; nothing changed
```

`ManualRedlineOptions` is a plain object (the TypeScript idiom; docx4j's is a builder, D1), `author` not empty,
`colourAsInsertion` each `[0-9A-Fa-f]{6}` (upper-cased; `auto` and a theme colour refused), `within` a `ContentControl`
view or a `w:sdt` element, `parts` `Part`s or part names as `TrackedChangeFilter.parts` takes them. Both verbs are
async (the story parts are unmarshalled on the way; a story part with no struck run is left as it is and keeps its
bytes, as CR-011's listing leaves parts without markup).

`ManualRedlineReport`: `preview`, `deletions`, `insertions`, `runs`, `strikeAndUnderline`, `byKind`, `byPart`;
`changes: { kind, type, own, part, address, text }[]` (kind `strike`, `double_strike`, `style_strike`, `underline`,
`colour`, `separator_space`, `paragraph_mark`; type `deletion`, `insertion`, `both`; `own` whether it is a `w:del` or
`w:ins` of its own); `left: { reason, part, address, text }[]` (reason `in_tracked_change`, `field`, `math`,
`style_strike`, `paragraph_mark`, `comment`); `toJson()` the same shape, docx4j's keys, so that docx4j-mcp's verb and
the paper read one report; `toString()`. Addresses are the text model's (`body/12`, `header:rId8/2`,
`body/11/textbox:0/0`), read before the conversion (D3).

### 3.2 The rules (docx4j CR-068 sections 11.2 and 12.2, as built and measured)

- **Stories** (D2 there): every story CR-011 lists: the main document, footnotes, endnotes, headers and footers, each
  text box in every `mc:AlternateContent` branch (the branch the text model reads reported, the others converted as
  twins with their own ids, as Word's are); block- and paragraph-level `mc:AlternateContent` likewise; the comments
  walked and their struck runs reported `comment` unless `comments: true`.
- **What converts** (D5, D6): a run whose own `w:rPr` holds `w:strike` or `w:dstrike` on, in a paragraph, a cell, a
  control's content, a hyperlink, a text box, a note: wrapped in a `w:del` (id, author, date), each `w:t` a `w:delText`,
  `w:tab`, `w:br`, `w:sym`, drawings and objects left in the run; the strike removed (an empty `w:rPr` with it). A struck
  run holding nothing to delete (a comment reference, a note's own mark, a separator) is neither converted nor
  reported.
- **Left and reported** (D10): a run in `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`, or whose strike is itself a
  `w:rPrChange` (`in_tracked_change`); a complex field's runs `begin` to `end`, the depth carried across the story's
  paragraphs, and a `w:fldSimple`'s (`field`); `m:r` (`math`); a run with no strike of its own that its styles strike,
  resolved by the `PropertyResolver` (the defaults, the paragraph style, the run style; a table style not consulted)
  (`style_strike`, unless `styleStrike`); consecutive runs of one paragraph and one reason are one entry, so a field is
  one. A run whose own `w:strike w:val="0"` turns an inherited strike off: nothing to do.
- **Grouping** (D8): adjacent converted runs of one content list in one `w:del` (one `w:ins`); a `w:proofErr` between
  them does not part them; anything else does, a bookmark or a comment range among them (CR-043 section 4.4, CR-002
  section 46: a bookmark with an end in removed content goes whole on accept).
- **The separating space** (D13): where the struck text neither starts nor ends with whitespace: (a) with whitespace or
  the paragraph's start before it and a space after it, that space; (b) otherwise, where it ends the paragraph or meets
  punctuation (`.,;:!?)]}` and closing quotes) with a space before it, that space; taken only from a plain run (not
  tracked, not a field's), split off with the run's formatting, into the struck text's `w:del` where the run is beside
  it, else a `w:del` of its own (the space after a struck hyperlink). In every story alike (Word's own deletion took it
  in the main story and not in a footnote or text box; the recommendation applied everywhere, section 8.4). An insertion
  takes its space by the rejected text (D16); the one space between a deletion and the insertion replacing it is `w:ins`
  holding `w:del`, gone in either outcome (D15). `separatorSpace: false` turns all of it off.
- **The paragraph mark** (D9 revisited): a struck mark (its own `w:pPr/w:rPr`), no text left shown in the paragraph after
  the conversion, the next item of its container a `w:p`, no `w:sectPr` in the paragraph: the mark takes `w:del`
  (`markParagraphDeleted`) and loses its strike; otherwise, where the text converted whole, left and reported
  `paragraph_mark` (a cell's last paragraph, one before a table or a control, the last of its story, a section's; the
  table rule of CR-002 section 55 agrees).
- **The opt-ins** (sections 3.3 and 12.2 there): style strike converts with its own `w:strike w:val="0"` (or `w:dstrike`)
  written; underline (`w:u` with a pattern other than `none`; a `w:u` with no pattern inherits and is not taken) becomes
  `w:ins` with the `w:u` removed (D7), strike winning where both (the underline kept on the deleted run, counted
  `strikeAndUnderline`); a named colour with no strike, underline, highlight or shading of its own becomes `w:ins` with
  the colour removed; in a comment a converted run keeps its marks (D14: Word's comment card shows no tracked change,
  and does not draw a formatted one, with a "Missing content" notice; section 12.7); the range is
  `TrackedChangeFilter`'s semantics read before anything converts (a run partly inside a bookmark counts as inside; a
  mark only if its paragraph is); `parts` limits the story parts; the audit property `docx4j:manualRedline`, its value
  the instant and the options, in `docProps/custom.xml`, none from a preview.
- **Dates and ids** (D4 clarified): `w:date` the instant's wall clock in `pkg.effectiveDateTimeZone()` with a `Z` (the
  one Word shows: Word 365 ignores `dateUtc`, section 8.4), `w16du:dateUtc` the instant, `w16du` made ignorable on each
  changed part; ids from one above the largest annotation id in the story parts (tracked changes, comments, bookmarks,
  ranges). `date: null` writes neither.
- **Idempotent**; `w:trackRevisions` untouched (D12); `preview` a dry run of the same walk, every decision made, nothing
  written (not a clone, so a `within` control of the package is found).

### 3.3 What changes for a reader of the converted file

After the default conversion `accepted` drops the struck text, `original` shows it plain, `markup` marks it deleted;
`Paragraph.text` reads Office JS's default view; `getAllTrackedChanges` lists the deletions by `manual redline`, and
`rejectAllRevisions({ authors: ['manual redline'] })` undoes the conversion to plain text (D6's choice: Word's own
deletion would give the struck text back). Nothing in the readers changes.

## 4. Tests and the oracle

- **The oracle copied**: docx4j's `docx4j-core-tests/src/test/resources/revisions/cr068/` (the probes built by
  `ManualRedlineProbeBuilder`, every text ASCII; docx4j's conversions `manual-redline-phase1.docx`, `-phase2.docx`,
  `manual-redline-2-phase2.docx`, `-d9.docx`; Word 365's `word-<key>-asis.txt`, `-accept-all.docx/.txt`,
  `-reject-all.docx/.txt` for `base`, `default`, `optins`, `d9`, `own`, `phase1`, `phase2`, `phase2b`; `cr068-log.txt`,
  the Review pane listings; `manual-redline-dates.docx`) into `test/cr068/`, unedited, with a SHA-1 manifest naming
  the docx4j commit, outside `fixtures/` as `test/s17/` and `test/mc-fun/` are.
- **The comparison**: `test/describe.mjs`, a port of `DocumentDescription` (test code only, D6): our conversion of each
  base, accepted and rejected by `acceptAllRevisions` and `rejectAllRevisions`, describes exactly as Word's
  `word-<key>-accept-all.docx` and `-reject-all.docx` do, for the default (`phase1`), `separatorSpace: false` (`d9`,
  Word's resolution of the phase 0 prediction), and every opt-in on both bases (`phase2`, `phase2b`, with the three
  normalisations of Word's save of the second base taken as Word saves them: N06's `w:u none` dropped, the unreferenced
  footnote and the empty header removed). The files Word was given (`manual-redline-phase1.docx`, the two phase 2 files)
  held to what the conversion writes, story part by story part as trees (`plain()`); the bytes depend on the marshal.
- **Unit tests** (`test/manual-redline.test.mjs`), docx4j's 27 mirrored: the report (counts, kinds, reasons, addresses);
  idempotence; `preview` changing nothing and reporting what the conversion does; ids unique and above the base's (100);
  every `w:del` in a paragraph, hyperlink or inline control's content; the dates in Sydney, UTC and none, the author,
  `w16du` declared on a document with no `mc:Ignorable`; the separating space's cases (punctuation, the paragraph's start
  and end, two struck words apart, digits meeting, a space inside the strike, off, a space in an insertion not taken);
  left alone (a tracked strike, a struck comment reference); the paragraph marks (joined; before a table, a cell's last,
  left and reported; not whole, kept unreported); the opt-ins (N07's own `w:dstrike w:val="0"`, N02's deletion keeping
  its underline, D15's `w:ins` holding `w:del`, D14 in the comment against the main story, colours and `auto` refused,
  a bookmark's range, a control's range (M17), `parts`, the audit property and none from a preview); the three views
  before and after through `Texts.of`; `Paragraph.text` after.
- **The revisions goldens** (CR-011) and the text goldens (CR-012 phase 6) unchanged: nothing read changes.
- **The gate**: the suite in full; the conversion run over every fixture with a struck run (count to be measured in
  phase 0) saves under CR-016's check unrefused and reopens.

## 5. Phases and sizes

| Phase | What | Size |
|---|---|---|
| 0 | The oracle copied with its manifest; `describe.mjs` ported and shown to describe Word's `base` saves as docx4j's does (a describe of `manual-redline.docx` equal under both, through the text saves); the fixtures with struck runs counted | ½ day |
| 1 | `ManualRedline.toTrackedChanges` and `preview`, the default scope and rules (D5, D6, D8, D9, D10, D13), the report, ids and dates, idempotence; `Texts.addresses` (D3); held to `phase1` and `d9` | 1 to 2 days |
| 2 | The opt-ins (D7, D14 to D16, the range, `parts`, the audit property); held to `phase2` and `phase2b` | 1 day |
| 3 | Docs (CLAUDE.md, README, this section), the registry, the editor and docx4j-python told; release with the next minor | ½ day |

## 6. Decisions (for Jason)

| # | Decision | Recommendation |
|---|---|---|
| D1 | Name and place | `ManualRedline.toTrackedChanges(pkg, options?)` and `ManualRedline.preview`, docx4j's names, a class with static methods in `src/model/content/manualRedline.mts`, exported from `./model` and the root; `ManualRedlineOptions` a plain object interface (not a builder), `ManualRedlineReport` a class with `toJson()`; no `pkg.` verb (one way to call it, docx4j's) |
| D2 | The default date's precision | **To the minute**, as every stamp this package writes (CR-018, CR-002 section 29) and as Word stamps; docx4j writes to the second. A caller's `date` as given. Word's Review pane shows minutes, so the Word saves cannot tell; `cr068-log.txt` lists 11:00. Noted as a departure in section 7 |
| D3 | The report's addresses | A public `Texts.addresses(pkg, options?)`: `Map<object, string>` from each block's `source` to its `address`, one `Texts.of` pass (docx4j's `Texts.addresses`, public since CR-068); a twin branch's paragraph takes its outer paragraph's address |
| D4 | The range | `within` (a `ContentControl` view or a `w:sdt` element) and `withinBookmark` as `TrackedChangeFilter` semantics, implemented for this CR in `trackedChangeFilter.mts` and so available to CR-011's listing and resolve-all too (CR-011 section 7 left them out until asked; this asks) |
| D5 | The audit property | `DocPropsCustomPart` gains `set(name, value)` (string, docx4j's `DocPropsCustomPart.setProperty` shape), the part created when absent; used only under `auditProperty: true` |
| D6 | The comparison | `DocumentDescription` ported as test code, `test/describe.mjs`, reusable for every Word-save oracle after this; `regionsOf` in `ranges.test.mjs` stays |
| D7 | The comments story | D14 as docx4j built it: a converted run in a comment keeps its marks (Word's card, section 12.7); no Word check here, docx4j's screenshots are the measurement |
| D8 | Release | With the next minor (new public API), after the items pending since 0.6.0; the editor told at phase 3 (it may want the verb) |

## 7. Departures and notes

- Options are a plain object, not a builder (D1).
- The default stamp is to the minute (D2), where docx4j's is to the second.
- Nothing else departs: the rules are docx4j's as measured; every difference found in phase 1 or 2 against Word's saves
  is investigated and fixed here unless docx4j's own test disagrees with Word, in which case it is reported to docx4j.
- Not in scope, as there: detecting a redline, deleted-row and deleted-cell marks, moves, property changes, turning
  Track Changes on, any change to a view or a reader.

## 8. Sources

- docx4j CR-068 `docs/developer/change-requests/CR-068-manual-redline-to-tracked-changes.md` (`8a98c78e7`), sections 6,
  8.5, 12.5 (D1 to D16), 11.2 and 12.2 (the rules as built), 8.4, 11.6 and 12.6 (Word's three runs), 12.7 (the comment
  card), 13 (the hand-off); `org.docx4j.model.revisions.ManualRedline`, `ManualRedlineOptions`, `ManualRedlineReport`;
  `ManualRedlineTest` (15), `ManualRedlineOptInsTest` (12), `DocumentDescription`, `ManualRedlineProbeBuilder`;
  `docx4j-core-tests/src/test/resources/revisions/cr068/`.
- Here: CR-011 (the CR-043 port and its oracle rule), CR-002 sections 29, 35, 46, 55, CR-018 (the stamp), CR-012 (the
  addresses), CR-016 (the save gate).
- The paper, `../docx4j-portfolio/docs/text_extraction_behaviour_model.md` section 7.13; the portfolio's
  `docs/mcp_tool_vocabulary.md` section 5 (`manual_redline_to_tracked_changes`).
