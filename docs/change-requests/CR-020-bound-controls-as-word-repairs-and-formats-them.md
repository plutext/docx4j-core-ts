# CR-020: Bound controls as Word repairs and formats them (the port of docx4j CR-058)

**Status:** Proposed 2026-11-02, from the portfolio's note (docx4j-portfolio-10) of docx-tool-identifier's phase 0, which
measured `@docx4j/core-ts` 0.6.0 at `f426a31` against docx4j 17.3.5 and found `bind-plain.docx` B7 read as its stale
surface. Measured here the same day (section 1), with a second gap the identifier could not see (a bound date's AM/PM).
D1 to D6 decided 2026-11-02, all as recommended (Jason: "CR-020 as recommended; for D3 we'll do the Script Lab read";
section 5), D3 to be held to check 73, that read (`test/README.md`). Check 73 run the same day (section 5.1): Word
binds two-parts' B7 to part 2 where its October save had part 1; check 73b the same day (section 5.2): in four fresh
sessions part 1, as docx4j's rule, and Office JS's `customXmlPart` the repaired part (D3 stands). D7 (keep the
part-name order) and D8 (`isMapped` false where the XPath selects nothing) decided the same day, as recommended (Jason:
"CR-020's D7 and D8 as recommended"). Nothing implemented. docx4j's
CR-058 is complete (phase 1 `FormattingSwitchHelper`, phase 2 `StoreItemResolver`, `35cdeaee6`), released in 17.3.5;
docx4j-python's counterpart is `python/CR-014.1c`.

## 0. Summary

1. **A `w:storeItemID` that names no part.** A template whose custom XML part was replaced by a tool that minted a new
   `ds:itemID` binds its controls to an ID no part has. Word repairs it silently on open: it binds the control to the
   custom XML part the XPath fits and its save writes that part's ID into `w:storeItemID` (Jason's Word 365 runs of
   22 and 23 October, docx4j CR-058 section 8). Here every reader and the bind treat the binding as unresolved: the
   text model reads the stored surface (`TextContext.boundValue`), `XmlMapping.customXmlPart` is undefined, and
   `applyBindings` skips the control. Nothing is lost or refused (this package never threw, as docx4j's NonXSLT
   traverser did, nor emptied the control, as its XSLT one did), but 3 of the probes' 5 stale controls show the stale
   text where Word shows the value. Fix: docx4j's `StoreItemResolver`, shared by the readers and the bind; the bind
   writes the resolved ID back as Word's save does.
2. **A bound date's `am/pm` and `AM/PM`.** `formatDate` knows `tt` but not Word's `am/pm` and `AM/PM`, so their `m`
   and `M` are read as minutes and months: `h:mm am/pm` at 15:05 gives `3:05 a5/p5` and `h:mm AM/PM` at 12:05
   `12:05 A10/P10`, where docx4j gives `3:05 PM` and `12:05 PM` (Word shows the upper-case names for both forms,
   docx4j's `wordAmPm`; section 1 has Word's own run, in its zone). Every other format in
   docx4j's two date probes agrees with Word here, in Word's zone under `boundDatesConvertedAsWord` (CR-018) and
   nominally otherwise.
3. **A date control with no `w:lid`.** Here `en-US`; docx4j (its D2 revised) the document's default language
   (`w:docDefaults/w:rPrDefault/w:rPr/w:lang/@w:val`), else the platform's, the nearest a server has to the editing
   language Word measurably used. No probe separates the two here (DF7 reads "22 October 2026" both ways).

## 1. Measured here (2026-11-02, `f426a31`'s build)

docx4j's probes (`docx4j-core-tests/src/test/resources/OpenDoPE/`, byte-identical to the share's `bind-fun/` copies
that Jason ran in Word 365), read with `Texts.of` and bound with `pkg.customXmlParts.applyBindings()`; Word's column
is its save (`bind-fun/*-word365.docx`) and docx4j CR-058 section 8's table.

| Probe, control | Binding | Here (text model, and after `applyBindings`) | Word 365 | docx4j 17.3.5 |
|---|---|---|---|---|
| `bind-plain` B7 | `{...00FF}`, one part `{...0001}` with `p:root/p:b7` | `B7_STALE_SURFACE`, skipped | `B7_XML_VALUE`; save `{...0001}` | as Word |
| `bind-plain-two-parts` B7 | `{...00FF}`, both parts hold `b7` | `B7_STALE_SURFACE`, skipped | `B7_XML_VALUE_PART1`; `{...0001}` | as Word |
| `bind-plain-two-parts` B11 | `{...00FF}`, only part 2 holds `b11` | `B11_STALE_SURFACE`, skipped | `B11_XML_VALUE_PART2`; `{...0002}` | as Word |
| `bind-plain-no-part` B7 | `{...00FF}`, XPath in `urn:probe:other` | `B7_STALE_SURFACE`, skipped | the surface, ID kept | as Word |
| `bind-plain-no-part` B12 | `{...00FF}`, XPath from `p:nothing` | `B12_STALE_SURFACE`, skipped | the surface, ID kept | as Word |

The dates (`typed-controls-dates.docx` = `bind-typed-formats.docx`, `typed-controls-twelve.docx` =
`bind-typed-twelve.docx`), after `applyBindings` with `pkg.dateTimeZone = 'Australia/Sydney'` and
`boundDatesConvertedAsWord` (Jason's machine was at UTC+11):

| Control | Format | Here | Word 365 |
|---|---|---|---|
| DF2 | `h:mm am/pm` | `2:05 a5/p5` | `2:05 AM` |
| DF3 | `h:mm AM/PM` | `11:05 A10/P10` | `11:05 PM` |
| TW1, TW3, TW4 | `h:mm AM/PM`, `h:mm AM/PM d MMM` | `12:05 A10/P10`, `... 22 Oct` | `12:05 PM`, `... 22 Oct` |
| TW2 | `hh:mm am/pm` | `12:05 a5/p5` | `12:05 AM` |
| TW5 | `h:mm AM/PM d MMM` | `4:05 A10/P10 23 Oct` | `4:05 AM 23 Oct` |
| DF1, DF4 to DF13, TW6 | | as Word | |

The hour, the day, the zone, the names in `w:lid`'s language, a value that is not a date (shown as stored) and the
empty value (the placeholder, CR-002 section 63) already agree. Without the zone the times are nominal (CR-002
section 56, docx4j CR-058 D6), as docx4j's are by default.

## 2. Proposal

### 2.1 The store item resolver (`src/model/customxml/storeItems.mts`)

docx4j's `StoreItemResolver.resolve` rule (its D4), over the package's custom XML data storage parts:
- the part the ID names (the comparison as `CustomXmlPartCollection.getItem`'s: case and braces ignored); the core and
  extended properties' well-known IDs as today, never repaired;
- else the parts whose root element's namespace and local name match the XPath's first step, its prefix read with the
  binding's `w:prefixMappings`, else the objects package's `NAMESPACE_PREFIXES` (docx4j falls back to its own table);
  an XPath with no named first step (`//`, `*`, an axis, an attribute) fits every part. In part-name order: one
  fitting part is taken; of several, the first whose XPath selects a node, else the first; none, unresolved;
- remembered per package and (ID, XPath), since one stale ID fits different parts for different XPaths (the
  two-parts probe: B7 part 1, B11 part 2); a package's parts changing (`customXmlParts.add`, a delete) forgets it.

`XmlMapping.isMapped` reads through it too (D8, check 73b): `true` only where the binding names a node, its part's by
ID or the resolved one's; `false` where the XPath selects nothing there, or a stale ID nothing fits, with `customXmlNode`
undefined, as Office JS reports B4.

Routed through it, as docx4j routes its lookups: `TextContext.boundValue` (the text model, REQ-509),
`XmlMapping.customXmlPart` and so `customXmlNode`, `applyBindingTo` (text, date, list, checkbox, picture),
`updateFromControl` (the surface written to the part Word would write it to), and the `od:xpath` passes'
part lookup (`opendope.mts`, `xhtml.mts`).

**Write-back (docx4j D5).** `applyBindings` writes the resolved part's ID into the control's `w:dataBinding/@w:storeItemID`
(and `w15:dataBinding`'s) where it binds the control, as Word's save does, in memory; saving is the caller's. No reader
writes: a document read and saved untouched stays byte for byte.

**Reporting.** `BindingResult.notes` names each repair (the control, the stale ID, the part taken, and where several
fitted) and each stale ID nothing fits, as docx4j logs them at INFO and WARN. docx4j's `strictStoreItemIds` restores
an exception this package never threw, so it has no counterpart (D4).

### 2.2 The date formatter (`formatDate` in `bindings.mts`)

- `am/pm` and `AM/PM`: the locale's AM and PM names (`Intl.DateTimeFormat`'s `dayPeriod` part under `hour12`)
  upper-cased in the locale, matched before `M` and `m` are. `A/P` and `a/p` are not in docx4j's
  table and were not measured; they stay literal.
- No `w:lid`: the document's default language, else `en-US` (D5).

## 3. Tests and the oracle

`test/store-items/` (outside `fixtures/`, a SHA-1 manifest): docx4j's five probes as above, copied unedited, and
Word 365's five saves from the share. `test/store-items.test.mjs`:
- docx4j's `StoreItemRepairTest` mirrored: B7 bound and its ID written back; the text model's value; two parts, B7 to
  part 1 and B11 to part 2 with the IDs; no part, B7 and B12 left with their surfaces and stale IDs while B1 binds;
- each probe's text after `applyBindings`, and each saved `w:storeItemID`, held to Word's save, control by control;
  the text model's plain text of each probe held to the text model's of Word's save for the bound controls;
- `XmlMapping.customXmlPart` on B7, and an untouched round trip of each probe byte for byte after every reader;
- docx4j's `BoundDateFormatTest` rows DF1 to DF13 and TW1 to TW6 against Word's saves, in Word's zone and nominally,
  and `withoutALidTheDocumentDefaultLanguageIsUsed` (a German `w:docDefaults` added in the test).
- check 73b's snippet (`test/fixtures/check73b/`) through the Word shim on the three probes and the swapped copy, held
  to Word's runs beside it: each control's `customXmlPart` and `customXmlNode`, `isMapped` (B4 `false`), the text after
  the open; the lookups' one difference listed (Word's `body.contentControls` gives none of these controls; the shim's
  gives them all), and check 73's runs the same where they read;
- `isMapped` and `customXmlNode` on B4 and on the no-part probe's B7 and B12 (D8).
The `customxml.test.mjs` cases already pinning a missing part's behaviour move with it.

## 4. Phases and sizes

One phase, small: the resolver and its routing (half a day), the formatter (an hour), the tests and the copy. Section
2.1's change to `XmlMapping.customXmlPart` is the content API's (CR-002), and is held to Office JS only by D3.

## 5. Decisions (taken 2026-11-02, all as recommended)

Jason, 2026-11-02: "CR-020 as recommended; for D3 we'll do the Script Lab read". D3's read is check 73 (`test/README.md`,
`test/fixtures/check73/`). Office JS's `XmlMapping` has no `storeItemId` (D3's row supposed one), so the check reads
`xmlMapping.customXmlPart` and `customXmlNode` on each of docx4j's three stale-ID probes, and the ID the body's OOXML
holds once Word has opened it. If Office JS's `customXmlPart` is not the part the resolver takes, D3 comes back to
Jason before the port is held to it.

| # | Decision | Recommendation |
|---|---|---|
| D1 | Where the resolver applies | Every reader and the bind, as docx4j: the text model, `XmlMapping`, `applyBindings`, `updateFromContentControls` and the `od:xpath` passes, so they all name the part Word binds |
| D2 | Write-back | `applyBindings` only, as docx4j's generator and Word's save; never a reader; `updateFromContentControls` writes to the resolved part and leaves the ID (docx4j's `UpdateXmlFromDocumentSurface`) |
| D3 | `XmlMapping.storeItemID` | The stored ID (the XML), `customXmlPart` the resolved part. Word repairs on open, so Office JS probably reads the part's ID; unmeasured: a Script Lab read of B7's `xmlMapping` on `bind-plain.docx` settles it, optional (check 73, taken) |
| D4 | Reporting, and docx4j's strict property | `notes` on `BindingResult` for each repair and each unresolved stale ID; no strict option, since nothing here ever threw on a stale ID |
| D5 | The date items | Take them in this CR: `am/pm` and `AM/PM` as Word shows them (measured), and the document's default language where `w:lid` is absent (docx4j's rule; unmeasured here) |
| D6 | Release | With the next release; the visible change is a stale control showing its value and a fixed `am/pm`, both Word's |

### 5.1 Check 73's run (2026-11-02, Word 365 16.0.20326.20232)

Jason ran check 73 on docx4j's three probes (`test/README.md` check 73; the runs in `test/fixtures/check73/`):
- **What Word shows:** `bind-plain` B7 `B7_XML_VALUE`; `bind-plain-no-part` B7 and B12 their stale surfaces, as on
  23 October; but `bind-plain-two-parts` B7 `B7_XML_VALUE_PART2` (and B11 `B11_XML_VALUE_PART2`), where Word's save of
  23 October shows `B7_XML_VALUE_PART1` and writes `{...0001}`. docx4j CR-058's D4 took part-name order from that save,
  and section 2.1 here took docx4j's rule.
- **Word's custom XML parts, in its order:** its own three, then in two-parts `{...0002}` before `{...0001}`, the
  reverse of the package's (part names, relationships and zip entries all list part 1 first). B7 went to the first
  fitting part in Word's order today.
- **D3, half settled:** the body's OOXML right after the open already holds the ID Word bound (`{...0001}` in
  `bind-plain`, `{...0002}` twice in two-parts, the stale `{...00FF}` kept in no-part). So an add-in reading OOXML sees
  the repaired ID before any save, where `XmlMapping.storeItemID` here gives the stored one until `applyBindings` writes it
  back (D2). `xmlMapping.customXmlPart` was not read: `body.contentControls` gave no control in Word, in any file.
- **Check 73b** (`test/README.md`) looks for the controls five ways and reads `customXmlPart` through the first that finds
  them; opens two-parts three times in fresh sessions, to see whether Word's order and its choice hold; and opens a copy
  with the two relationships listed the other way round, to see whether they follow the relationships.

D7 (of several parts the XPath fits, which: section 2.1, docx4j CR-058 D4) was put then, to wait for check 73b: Word's
order taken if it held and followed the file, docx4j's part-name order kept if it changed between opens. Section 5.2
has it with 73b's result.

### 5.2 Check 73b's run (2026-11-02)

Jason ran check 73b (`test/README.md` check 73b; the runs in `test/fixtures/check73b/`), each file in a fresh session:
- **Two-parts, three opens, and a copy with its two relationships listed the other way round:** B7 shows
  `B7_XML_VALUE_PART1` and B11 `B11_XML_VALUE_PART2` every time; Word's parts are in the package's order, `{...0001}`
  first, whichever way the relationships run; B7's `customXmlPart` is `{...0001}` and B11's `{...0002}`, as the OOXML
  says. That is Word's save of 23 October and docx4j's rule: of the fitting parts, the first in part-name order (here
  also the zip's and the item properties' order; the relationships are ruled out). Check 73's open, part 2 first, came
  straight after `bind-plain.docx` had been open; it is recorded and not explained.
- **D3 stands:** Office JS's `xmlMapping.customXmlPart` for a stale control is the part Word bound it to, the part the
  resolver takes (`bind-plain` B7 and the two-parts controls), so `customXmlPart` here is the resolved part, as D3 has
  it. `XmlMapping.storeItemID` (an extension; Office JS has none) stays the ID as stored until `applyBindings` writes the
  part's back (D2); Word's own OOXML holds the repaired ID from the open, a difference the shim's `getOoxml` keeps.
- **`isMapped`:** Word reports `false` for B4, bound to its part by ID with an XPath that selects nothing, and
  `customXmlNode` a `GeneralException`; here `isMapped` is `true` wherever a `w:dataBinding` is (`XmlMapping.isMapped`).
  Every other control, the stale ones repaired, `true` in both.
- **Only `body.getContentControls()` finds these probes' controls in Word**; `body.contentControls`,
  `document.contentControls` and the paragraphs' collections give none, in every file. The shim gives them through all.
  Recorded only: what in these files Word's older collections pass over is unmeasured, and no decision depends on it.

Jason, 2026-11-02: "CR-020's D7 and D8 as recommended".

| # | Decision | Recommendation (taken) |
|---|---|---|
| D7 | Of several parts the XPath fits, which | Keep docx4j's rule as section 2.1 has it, the first in part-name order that selects a node: Word's choice in every fresh session and in its save; check 73's reversed open recorded as unexplained. Nothing for docx4j to change; it hears the result |
| D8 | `XmlMapping.isMapped` where the XPath selects nothing | `false`, as Office JS (check 73b's B4): mapped is "bound to a node", read through the resolver, so a stale ID nothing fits is `false` too (unmeasured; it follows); `customXmlNode` stays undefined there. In this CR, since `XmlMapping` goes through the resolver here; a small change in the content API (CR-002), noted there |

## 6. Registry

`core-ts/CR-020`, proposed, depending on `docx4j/CR-058` (done). The identifier's other two findings are already
entries: CR-069's case mapping is `core-ts/CR-012.24` (proposed; CR-012 section 24), and `fields: update` is CR-012
phase 5a.

## 7. Sources

- docx4j CR-058 (`docs/developer/change-requests/CR-058-bound-date-format-and-store-item-repair.md`), sections 0, 5,
  5a, 8 to 10; `StoreItemResolver.java`; `FormattingSwitchHelper.formatDate` and `wordAmPm`; `StoreItemRepairTest`,
  `BoundDateFormatTest`.
- Word 365's saves, `~/vbShares/Office 2016/bind-fun/*-word365.docx` and `README.txt` (Jason, 22 and 23 October 2026).
- The scratch probes of 2026-11-02 (not committed): the store-item and date runs over the five probes.
