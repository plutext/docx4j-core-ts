# CR-018: Dates in the package's zone

**Status:** Proposed and implemented 2026-10-28 (`d9c0cef`; D6's shim reading the same day, after docx4j's probe, `4275c36`), at Jason's word through the portfolio ("commit, then have the
ports do their work"): the portfolio's date rule (`../docx4j-portfolio/docs/date_handling_across_ports.md` section 3.6,
`4ac305b`), ruled the same day, applied here. Mirrors docx4j CR-067 D12 (`c3e9920e5`: `OpcPackage.setDateTimeZone`,
`setBoundDatesConvertedAsWord`, its `PackageDateTimeZoneTest`) and docx4j's `docx4j.dcterms.write`. Unreleased.

## 1. Why

Word stores two kinds of date. A revision's or a comment's `w:date` and a date field's source are **instants**, which Word
shows on the reader's wall clock: `w:date` is the author's local time labelled `Z`, with the true UTC beside it in
`w16du:dateUtc` or `w16cex:dateUtc` where Word 2019+ writes it (CR-002 section 29, check 23). A bound date is a
**nominal day**, which the picker stores as local midnight labelled `Z` and Word then shows converted to the machine's
zone, a day early west of UTC, a wart (CR-002 section 56; docx4j CR-058 D6's probe).

Until this CR the zone every instant was rendered in here was the runtime's, with no way to say otherwise. Under Office
JS that is the client's zone and right. Under Node on a server it is the server's, and a document stamped there for a
client in another zone carries the server's wall clock in `w:date`, which Word then shows as the author's local time.
Jason (the review's section 3.4): "docx4j running on a server shouldn't necessarily use server time. For a given OPC
package, it is desirable to be able to specify time zone (i.e. to use the client's time zone) ... might be desirable for
core-ts since it could be used server side."

The review also found this package stamping a new comment to the second where Word, and a revision here, stamp to the
minute; and it set, for every port, one shape for writing the save's time into the core properties, and the rule that
Word's bound-date conversion is an opt-in named for what it does, never a side effect of the zone.

## 2. Today (measured 2026-10-28, at `b7c2680`)

Three rules, each from a measurement, and the gaps the ruling closes:

| Place | Read | Written | Zone | Precision |
|---|---|---|---|---|
| A revision's `w:date` / `w16du:dateUtc` (every revision element, `tracking.mts`) | `dateUtc` as an instant; else `w:date` as local wall-clock time, its `Z` notwithstanding; a real offset as it says | `w:date` the local wall clock with a `Z`, and `dateUtc` in UTC from a part-bound tracker (every `Body` edit), `w:date` alone from the package's own | the runtime's; no option | the default stamp now to the minute; a caller's date as given |
| A comment's `w:date` / `w16cex:dateUtc` | as above | both, one instant | the runtime's; no option | **to the second** |
| A bound date (`bindings.mts`, `kinds.mts`) | nominal, any `Z` or offset ignored | `w:fullDate` nominal with a `Z` | none | as stored |
| Date fields | none: `fields: update` refused (CR-012 phase 5) | | | |
| `dcterms:created` / `modified`, `cp:lastPrinted` | the Word shim only, by `new Date(text)` | the shim only, on the caller's set, UTC `Z` to the second | | a save never touches them; `createPackage` writes an **empty** core properties part (no dates), and an `app.xml` |

## 3. Decisions (the portfolio's rule, Jason 2026-10-28; nothing open)

| # | Decision | Here |
|---|---|---|
| D1 | **The zone is a property of the package**, the client's for the package the caller is working on; unset, the process's zone. | `pkg.dateTimeZone`, an IANA id or undefined; `pkg.effectiveDateTimeZone()` the one in force. Copied by `clone()`. Never saved: it decides which wall-clock fields are written for an instant. |
| D2 | **An instant is rendered, and a `w:date` without `dateUtc` read, in that zone.** | A new revision's and comment's `w:date` is the zone's wall clock; `TrackedChange.date` and `Comment.creationDate` read a `w:date` without `dateUtc` on the zone's wall clock; `dateUtc` is UTC whatever the zone; a `w:date` with a real offset as it says. |
| D3 | **Stamps to the minute**, as Word writes them. | A new comment is stamped `nowToTheMinute()`, as a revision is; a caller's `date` as given. |
| D4 | **Word's wart is a separate opt-in**, named for what it does, off by default; setting the zone does not turn it on. | `pkg.boundDatesConvertedAsWord` (docx4j's `setBoundDatesConvertedAsWord`): on, a bound value carrying a `Z` or an offset is the instant it names, shown on the package zone's wall clock, as Word reads the `Z` as UTC and shows the machine's time; a value with no marker, or a date alone, stays nominal (unmeasured in Word, point 4). Off, CR-002 section 56 unchanged. |
| D5 | **Core properties on save: untouched by default; docx4j's opt-in is the shape**, one name per port. | `pkg.writeTimesOnSave` (docx4j `docx4j.dcterms.write`, `DocPropsCorePart.writeTimes`): a save sets `dcterms:modified` to now, UTC to the second in Office's W3CDTF form, and `dcterms:created` too for a package created here that has none; a loaded package with no core properties part is left without one. |
| D6 | **Other producers' forms** (date-only, offset-free) were unmeasured in Word; no port changed its reading until a Word probe had run (docx4j's to build). **Measured the same day** (docx4j CR-067 section 12.2, Jason, Word 365 2608/20326.20232, Sydney): in a core property a date alone and a time with no offset are **UTC**, as a `Z` value is (`2024-02-03` showed 11:00 at UTC+11; `2025-06-07T08:09:10` showed 18:09 at UTC+10; `2025-01-02T03:04:05Z` in `cp:lastPrinted` 14:04), shown to the minute; in a revision's `w:date` a date alone is local midnight, an offset-free time local, a `Z` ignored (section 29's rule, seen again); a date control's `w:fullDate` is nominal in all three forms (section 56 stands). Word's save rewrites every `w:date` and `w:fullDate` as the nominal clock to the minute with a `Z`, keeps `lastPrinted` to the minute, and sets created and modified to the save's moment. | The shim's `creationDate` and `lastSaveTime` read a date-only or offset-free value as UTC fields (`document.mts` `dateOf`; `new Date(text)` read the offset-free time as local), a `Z` or offset as before, seconds kept; `lastPrintDate` already read a missing offset as UTC. The revision and bound-date readers already agree with the probe and are unchanged. |
| D7 | **Date fields** take the zone when they exist. | CR-012 phase 5: a field's instant shown in `effectiveDateTimeZone()`, to the minute, docx4j CR-067 section 11.3's measured rule; a per-call override as docx4j's `FieldContext.zone`. Not built here. |

## 4. Implementation (2026-10-28)

- `src/model/dateZone.mts` (new): `runtimeTimeZone()`; `assertTimeZone(id)`, a `RangeError` for a zone `Intl` does not
  know (a wrong zone shifts every stamp, so a package refuses it at the set rather than falling back, where docx4j's
  property is ignored with a warning); `wallClockIn(date, zone)`, the fields `Intl.DateTimeFormat` gives for the zone
  (`hour12: false`, the 24 some engines give at midnight folded); `instantOf(clock, zone)`, the instant at which the
  zone's wall clock shows the fields, found by offset in two rounds, which settles across a daylight-saving step; a clock
  time that does not exist on a spring-forward day gives the instant an hour on, as `Date` does for the runtime's zone.
  `zone` undefined is the runtime's zone through `Date`'s local getters, as before.
- `tracking.mts`: `TrackingHost.dateTimeZone`; `wordCalendarOf(date, zone?)` and `revisionDateOf(value, zone?)`;
  `ChangeTracker.markup()` stamps in the host's zone. `WordprocessingMLPackage` is the host and inherits the property.
- `OpcPackage.mts`: `dateTimeZone` (a getter and a validating setter), `effectiveDateTimeZone()`,
  `boundDatesConvertedAsWord = false`, `writeTimesOnSave = false`; `clone()` copies the three.
- `Comment.mts`: `creationDate` and the new comment's stamp pass the package's zone (duck-typed, as the package module
  cannot be imported there); the stamp `nowToTheMinute()`. `TrackedChange.date` passes its package's zone.
- `bindings.mts`: `applyBindingTo` reads the package's opt-in: `dateAsWordShows(value, zone)` over the same parser as
  `dateOf` (`parsedDateOf`, now keeping the marker's offset); the converted date is nominal in UTC fields, as every bound
  date here is, so `formatDate`, `w:fullDate` and the reverse take it unchanged.
- `Save.mts`: `savePackage` calls `writeTimes(pkg)` first under `writeTimesOnSave`: the part unmarshalled (so this save
  marshals it), `modified` set, `created` set where the package has no source container and none; a loaded package
  without the part gets none.

The content API's own reads and writes are the only users; the text model gives every stored date as written (the
rule's point 1), and the Word shim's document properties are untouched (D6).

## 5. Tests

`test/dates.test.mjs` (7): the option's default, the runtime zone in force, a bad zone refused at the set with nothing
changed, the opt-ins off by default and the zone not turning the conversion on; a revision stamped in Tokyo, New York,
UTC, Sydney (daylight time: +11 in February) and the runtime's zone, `w:date` the zone's wall clock and `dateUtc` the
instant; a `w:date` without `dateUtc` read in three zones, `dateUtc` winning in any, a real offset as it says (docx4j's
`aRevisionsLocalClockIsReadInThePackagesZone`); a new comment to the minute, a given date to the second, its `w:date` in
the zone, read back in another zone; the bound-date opt-in over five value forms, off and on, in Sydney and New York
(the picker's midnight `Z` the day before west of UTC, docx4j CR-058 D6's probe); `writeTimesOnSave` off and on, a created
package (`created` and `modified`, W3CDTF `Z` to the second, `xsi:type`), a loaded document (`created` kept, `modified`
moved; off, byte for byte) and a loaded package stripped of its core properties part (none made); `clone()`.
`tracking.test.mjs`, `comments.test.mjs` and `customxml.test.mjs` (130) pass unchanged: the runtime's zone is still the
default, and a comment's given date is still written to the second.

## 6. Departures from docx4j, and notes

- `dateTimeZone` is a property, and `boundDatesConvertedAsWord` and `writeTimesOnSave` plain booleans, in this package's
  style (`normaliseOnSave`, `trackedChangeDate`); docx4j has setters, a JVM-wide property between the package's zone and the
  JVM default, and a `Boolean` opt-in with a property fallback. There is no process-wide setting here: the fallback is the
  runtime's zone, which is the client's under Office JS.
- A zone the runtime does not know is refused at the set (`RangeError`); docx4j logs and ignores a bad property.
- A converted bound date is written to `w:fullDate` too (the shown date, as the picker would store what it shows); docx4j
  converts only the displayed run. Unmeasured either way.
- `createPackage` writes an empty core properties part (the review's table said none here; corrected 2026-10-28), so a
  created package under `writeTimesOnSave` gets both dates in that part rather than a new one.
- `cp:lastPrinted` is never written by a save, as docx4j never writes it.
