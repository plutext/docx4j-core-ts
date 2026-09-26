# CR-004: SpreadsheetML extension parts (Excel 2010 and 2013: slicers, timelines, control properties, custom data, survey, data model)

**Status:** Phase A implemented 2026-09-24 (section 5), phase B 2026-09-25 (section 6). **Closed
2026-09-26**: acceptance check 14 passed on both halves once it was measured on a control that
worked - Excel 365 drives the check box of a re-marshalled worksheet whose `mc:AlternateContent`
wrapper is gone, so resolving markup compatibility on load has no counterexample here (section 7,
and docx4j CR-021 section 8.13). Two lessons survive from getting there: `mc:Choice/@Requires` is a
prefix reference, and an after-save check needs a before-save fixture that works. Phase B was unblocked 2026-09-20 by objects 0.1.6 (the docx4j CR-022 regeneration, the
`x14` and `x15` Excel modules; 0.1.5 carried it but could not load a pptx or xlsx whose text body
holds an equation, CR-001 section 17.5), which this package now depends on.
**Depends on:** CR-001 Phase A (parts, registry, `DefaultXmlPart`, `BinaryPart`); for Phase B, the
objects package's CR-022 regeneration (done 2026-09-20 as objects `141f6bd` from docx4j `16844ff03`,
unreleased: six new modules including `x14` and `x15`, the seven roots unmarshalling typed,
`mc:Ignorable` on the sml roots, `x14ac` and `xr:revisionPtr` typed; released as 0.1.5 or later).
**Counterpart:** docx4j CR-022 phases 1 and 2 (`docs/developer/change-requests/CR-022-excel-2010-2013-extensions.md`,
sections 17 and 18), `org.docx4j.openpackaging.parts.SpreadsheetML.{SlicerCachePart, SlicersPart,
TimelineCachePart, TimelinesPart, ControlPropertiesPart, CustomDataPropertiesPart, CustomDataPart,
SurveyPart, DataModelPart}`, `docx4j-core-tests org.xlsx4j.ExcelExtensionPartsTest` and
`ExcelExtensionsTest` on `cr022-slicers-timelines.xlsx`.

## 1. Why

docx4j CR-022 (2026-09-20) typed the Excel 2010 and 2013 extension parts that Excel writes for
slicers, timelines, form controls, custom data, surveys and the Power Pivot data model. Here they
load as `DefaultXmlPart` or `BinaryPart` and round-trip untouched byte for byte, which is correct
but anonymous: `pkg.getPart(name) instanceof SlicersPart` is false, `WorkbookPart` has no shortcut
to them, and once the objects package types their roots nothing here would unmarshal them. Parts
and relationships are this package's territory (the dividing rule in `CLAUDE.md`), so the nine
part classes belong here.

## 2. The parts

Content type, relationship type, root element, source part; the strings are Excel 365's as docx4j
measured them (the timeline pair is lowercase and 2011, not the specification pages' forms):

| Part | Content type | Relationship type | Root | From |
|---|---|---|---|---|
| `SlicerCachePart` | `application/vnd.ms-excel.slicerCache+xml` | `http://schemas.microsoft.com/office/2007/relationships/slicerCache` | `x14:slicerCacheDefinition` | workbook |
| `SlicersPart` | `application/vnd.ms-excel.slicer+xml` | `.../office/2007/relationships/slicer` | `x14:slicers` | worksheet |
| `TimelineCachePart` | `application/vnd.ms-excel.timelineCache+xml` | `.../office/2011/relationships/timelineCache` | `x15:timelineCacheDefinition` | workbook |
| `TimelinesPart` | `application/vnd.ms-excel.timeline+xml` | `.../office/2011/relationships/timeline` | `x15:timelines` | worksheet |
| `ControlPropertiesPart` | `application/vnd.ms-excel.controlproperties+xml` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp` | `x14:formControlPr` | worksheet (the `x:control` inside its `mc:AlternateContent` carries the `r:id`) |
| `CustomDataPropertiesPart` | `application/vnd.openxmlformats-officedocument.customDataProperties+xml` | `.../officeDocument/2006/relationships/customDataProps` | `x14:datastoreItem` | workbook |
| `CustomDataPart` (binary) | `application/binary` | `.../officeDocument/2006/relationships/customData` | | its properties part |
| `SurveyPart` | `application/vnd.ms-excel.Survey+xml` | `http://schemas.microsoft.com/office/2010/relationships/Survey` | `x15:survey` | workbook |
| `DataModelPart` (binary) | `application/vnd.openxmlformats-officedocument.model+data` | `.../officeDocument/2006/relationships/powerPivotData` | ([MS-XLDM]) `xl/model/item.data` | workbook |

`CustomDataPart`'s content type is generic, so it is chosen by relationship type, as the registry
already does for altChunk, embedded and custom XML parts (CR-001 section 5.4).

## 3. Phases

**Phase A (no objects dependency):** the nine classes in `src/parts/sml/index.mts`, the seven XML
ones over `DefaultXmlPart` for now (their DOM round-trips byte for byte, and `mc:Ignorable` is
untouched because nothing re-marshals them); `ContentTypes.SPREADSHEETML_SLICER_CACHE`, `_SLICERS`,
`_TIMELINE_CACHE`, `_TIMELINES`, `_CONTROL_PROPERTIES`, `_CUSTOM_DATA_PROPERTIES`, `_CUSTOM_DATA`,
`_SURVEY`, `_DATA_MODEL` and the same names in `Namespaces`; registry entries beside the pivot cache
parts (content type for eight, relationship type for `CustomDataPart`); shortcuts `WorkbookPart.
slicerCacheParts`, `timelineCacheParts`, `customDataPropertiesParts`, `surveyPart`, `dataModelPart`
and `WorksheetPart.slicersParts`, `timelinesParts`, `controlPropertiesParts`. Fixture: docx4j's
`cr022-slicers-timelines.xlsx` copied with provenance; tests: parts by relationship type and by
part name, a save keeping the content-type overrides and relationships, byte-identical round trip.
Effort: half a day.

**Phase B (after the objects regeneration):** the seven XML parts become `XmlPart<T>` over the
`x14` / `x15` roots; a re-marshalled slicers root is identical in shape to Excel's (default
namespace `x14`, `mc:Ignorable="x xr10"` with both declared: CR-001 section 17.3's re-declaration
covers it); `ExcelExtensionPartsTest` and `ExcelExtensionsTest` mirrored. Effort: half a day.

**Scope re-checked 2026-09-25, before starting it.** Two halves, and the second is no longer
blocked:

1. *The typed roots.* Unblocked since objects 0.1.5 and confirmed against 0.2.0: all seven roots
   unmarshal typed - `x14` `CTSlicerCacheDefinition`, `CTSlicers`, `CTFormControlPr`,
   `CTDatastoreItem`; `x15` `CTTimelineCacheDefinition`, `CTTimelines`, `CTSurvey` - the four with
   a fixture from `cr022-slicers-timelines.xlsx` and the other three from minimal documents. Excel's
   `mc:Ignorable="x xr10"` survives a round trip of the slicer and slicer cache parts with both
   prefixes declared (objects CR-006), which is what re-marshalling them needs.
2. *`UNDERSTOOD_NAMESPACES` and the slicer drawings.* This CR and CR-001 section 17.4 both said
   this waits on docx4j binding the slicer schemas (`sle`, `sle15`, `tsle`), which it still has
   not: those prefixes are in the objects package's `NAMESPACE_PREFIXES` but **no module carries
   them** (checked). It is no longer a blocker, because objects 0.2.0 made `a:graphicData`'s
   wildcard **lax** (CR-004 section 5, docx4j `8e8f6ea83`): an unbound graphic now stays DOM
   instead of being fatal, so a Choice naming those namespaces can be *taken* without a module for
   its content. Measured on `cr022-slicers-timelines.xlsx` with a preprocessor whose understood set
   adds the three:

   | | `drawing1.xml` | `drawing2.xml` |
   |---|---|---|
   | as shipped | 1 slicer kept | **0 - its only slicer lost to the Fallback** |
   | the three understood | 2 slicers kept | 1 slicer kept |

   and no part throws either way. So phase B can close the loss CR-001 section 17.4 recorded ("a
   re-marshalled spreadsheet drawing loses its slicers and timelines") by adding the three
   namespaces, with the slicer content kept as DOM - no objects modules needed, and none in
   prospect. What to weigh when doing it: a kept Choice is lossless only if everything in it
   round-trips, which is the rule docx4j learned in its CR-021 section 8.9 and which the lax
   wildcard now satisfies for this content; the Fallback that would be given up is the "Excel 2010
   or higher" placeholder box, which is worth less than the slicer.

## 4. Not in scope

The slicer, timeline and control content itself (a SpreadsheetML content API is not planned); the
DrawingML `mc:AlternateContent` in drawings (`a14`, `tsle`, `sle15`), which docx4j still resolves to
the Fallback and lists as a DrawingML CR.

## 5. Phase A implementation notes (2026-09-24)

As specified in section 3, with three things worth recording.

**The nine classes** are in `src/parts/sml/index.mts`: `SlicerCachePart`, `SlicersPart`,
`TimelineCachePart`, `TimelinesPart`, `ControlPropertiesPart`, `CustomDataPropertiesPart` and
`SurveyPart` over `DefaultXmlPart`, `CustomDataPart` and `DataModelPart` over `BinaryPart`, each
with docx4j's default part name, content type and relationship type (asserted class by class, the
strings taken from docx4j's `ContentTypes` and `Namespaces` rather than from the specification
pages). The registry takes eight by content type and `CustomDataPart` by relationship type, its
content type being the generic `application/binary`. Shortcuts: `WorkbookPart.slicerCacheParts`,
`timelineCacheParts`, `customDataPropertiesParts`, `surveyPart` and `dataModelPart`;
`WorksheetPart.slicersParts`, `timelinesParts` and `controlPropertiesParts`. The six collection
getters share a new `Part.partsByRelationshipType(relationshipType, class)`, which is what
`WorksheetPart.tableParts` already did by hand; that one getter now uses it too.

**Tests** (`test/sml-extensions.test.mjs`, 4 tests): over docx4j's `cr022-slicers-timelines.xlsx`,
the parts load with their classes and the shortcuts answer per sheet (sheet2 a slicer and a
timeline, sheet3 a slicer, sheet1 neither); the workbook round-trips byte for byte, with the four
content-type overrides and the workbook and sheet relationships intact, and a reload of the saved
bytes finds them again. The fixture has no control properties, custom data, survey or data model
part - Excel writes those only for a form control, an add-in, a survey or Power Pivot - so those
five are built, added, saved and reloaded, which is what exercises `CustomDataPart`'s
relationship-type entry and every new content-type override.

**A gap this turned up, upstream of here.** `xl/drawings/drawing1.xml` of that fixture cannot be
unmarshalled: its `mc:AlternateContent` Choice requires `a14`, which **is** understood, so the
preprocessor takes it, and inside it `a:graphicData` holds the slicer's `sle:slicer`, for which
the model has no module - and `CT_GraphicalObjectData`'s wildcard is `allowDom: false`, so the
unknown graphic is fatal instead of staying DOM. JAXB's `@XmlAnyElement(lax=true)` keeps it, which
is why docx4j does not see this. It is the same remedy as docx4j CR-021 applied to the `mce`
wildcards (CR-001 section 15.4): the schema's `processContents` wants to be lax. Reported to the
objects session for docx4j. Consequences here, none of which phase A can fix: a workbook whose
drawing frames a slicer or a timeline through an understood Choice cannot have that drawing part
unmarshalled (the part still round-trips from its bytes, and every other part of the workbook
reads); `drawing2.xml` of the same fixture frames a slicer too but its Choice requires `sle15`,
which is not understood, so its Fallback picture is taken and the part reads - what saves it is
the branch being given up. `test/ignorable.test.mjs` names `drawing1.xml` as unreadable and
asserts the rejection rather than skipping it, so a second such part cannot hide behind it.

**Fixed upstream the same day** (relayed 2026-09-24): docx4j `8e8f6ea83` on `VERSION_17_2_1`
makes the `a:graphicData` wildcard lax - in `xsd/dml/dml-graphicalObject.xsd`, not `dml-main.xsd`
as guessed above - and the `vmlDrawing` root's `##any` wildcard with it (`xsd/vml/vml__ROOT.xsd`,
CR-026's other half); docx4j records it in its CR-024 section 10. The objects session regenerates
once that commit is pushed, together with `dae2dfc8b` (the `CT_Settings` order, CR-001 section
17.7). When the release lands here, `drawing1.xml` should unmarshal with its `a14` branch taken
and the slicer kept as DOM: the entry for it in `test/ignorable.test.mjs`'s `UNREADABLE` is then
removed, and its removal is the check that the fix arrived.

**It did, and that is how this session found out** (2026-09-25): objects 0.2.0 (npm, tag
`f90a2c6`) carries the regeneration, and the first run against it failed `ignorable.test.mjs`
with `Missing expected rejection` for `drawing1.xml`. The part now unmarshals with the `a14`
branch taken and `sle:slicer` kept as DOM, and re-marshals with the slicer intact;
`sml-extensions.test.mjs` asserts both, and that Excel's `mc:Ignorable="x xr10"` survives a round
trip of the slicer and slicer cache parts with both prefixes declared (objects CR-006), which is
what phase B needs to re-marshal them. The device - recording a dependency's defect as an
assertion of the broken behaviour rather than as a skip - is written up in CR-001 section 19.

**Not affected:** the parity goldens (no fixture of theirs has an extension part), and
`createPackage()`, which writes none of these parts. docx4j 17.2.1 has since added two more
SpreadsheetML content types, threaded comments and persons, which it leaves as `DefaultXmlPart`s
and which are therefore already right here.

## 6. Phase B implementation notes (2026-09-25)

Both halves of the 2026-09-25 re-check in section 3, and nothing else; 512 tests pass, the parity
goldens do not move (no parity fixture has an extension part), and `npm run typecheck` is clean.

**The typed roots.** The seven XML classes in `src/parts/sml/index.mts` are `XmlPart<T>` over the
`x14` and `x15` roots instead of `DefaultXmlPart`: `SlicerCachePart` `x14.CTSlicerCacheDefinition`,
`SlicersPart` `x14.CTSlicers`, `ControlPropertiesPart` `x14.CTFormControlPr`,
`CustomDataPropertiesPart` `x14.CTDatastoreItem`, `TimelineCachePart` `x15.CTTimelineCacheDefinition`,
`TimelinesPart` `x15.CTTimelines`, `SurveyPart` `x15.CTSurvey`, each with its root name, so
`setContents()` on a new part needs no name. `CustomDataPart` and `DataModelPart` stay
`BinaryPart`s. Nothing else changed: the registry entries, the shortcuts and the content and
relationship types are phase A's. Excel's `mc:Ignorable` is a property of each of these roots
(objects CR-022), and `XmlPart.declareIgnorablePrefixes` re-declares what the model does not bind,
so a re-marshalled `xl/slicers/slicer1.xml` is `<x14:slicers ... mc:Ignorable="x xr10">` with both
prefixes declared, which is what `ExcelExtensionPartsTest.partsSurviveASave` asserts of docx4j.

One correction to phase A: `x14:datastoreItem`'s attribute is `id`, not `itemID` (docx4j's
`xsd/xlsx/office_spreadsheetml_2009_9_main.xsd` line 794). Phase A's built fixture wrote `itemID`,
which nothing read because the part was a DOM; the typed read found it.

**`UNDERSTOOD_NAMESPACES`.** The set was a copy of the objects package's module namespaces at
0.1.0 and had gone stale by nine modules. It is now 0.2.0's, in three marked groups: the module
defaults; four namespaces the model binds without any module defaulting to them (`r`, `prop`,
`cs`, `ink16`); and the three slicer namespaces, which nothing binds at all.
`mce.test.mjs` asserts the first group module by module, so the next regeneration announces an
addition instead of quietly leaving it out of the Choice rule.

Two of the nine catch a real loss:

- `x14` (`.../spreadsheetml/2009/9/main`). Excel wraps a worksheet's form controls in an
  `mc:AlternateContent` whose only Choice requires `x14` and which has **no Fallback**, so the
  preprocessor was dropping the lot - the `x:controls`, and with it the `r:id` naming each
  `ControlPropertiesPart`. Measured on `cr022-checkbox.xlsx`: `worksheet.controls` was undefined
  and a re-marshalled sheet had no controls; it now holds the control, its `r:id` resolves to the
  part, and the sheet writes them back. This is docx4j CR-022's
  `ExcelExtensionsTest.formControlAlternateContentKept` reached by the other route: docx4j keeps
  the `mc:AlternateContent` element itself, where resolving on load cannot rebuild the wrapper, so
  the saved sheet holds a bare `x:controls`. That is schema-valid (`CT_Worksheet` has `controls`,
  and `CT_Control` has `controlPr`, in the 4th-edition transitional schema) but it is not Excel's
  own markup, and it went to Excel as acceptance check 14 (`test/README.md`). **It failed**, and
  section 7 records what is known and what is still to be isolated.
- `sle`, `sle15` and `tsle` (`.../drawing/2010/slicer`, `/2012/slicer`, `/2012/timeslicer`), which
  closes the loss CR-001 section 17.4 recorded. The Choice these gate is an `xdr:graphicFrame`
  whose `a:graphicData` holds the one unbound element, and that wildcard is lax since docx4j
  `8e8f6ea83`, so the whole Choice round-trips with the slicer kept as DOM. Measured on
  `cr022-slicers-timelines.xlsx`, matching the table in section 3: `drawing1.xml` keeps its slicer
  *and* its timeline (it kept only the slicer before, the timeline's `tsle` Choice losing to the
  Fallback), `drawing2.xml` keeps the slicer it lost entirely. What is given up is the "Slicers are
  supported in Excel 2010 or later" placeholder box, which is worth less than the slicer. The rule
  the Choice has to pass is docx4j CR-021 section 8.9's - a kept Choice is lossless only if
  everything in it round-trips - and the lax wildcard is what makes these pass it.

The other seven additions (`x15`, `x15ac`, `x12ac`, `x16`, `xr`, and Word's `w16`, `w16cex`,
`cei`) moved no fixture; they are in the set because their content is bound, which is the rule.

**Tests.** `test/sml-extensions.test.mjs` is 9 tests (was 4), mirroring
`ExcelExtensionPartsTest` in full and the part of `ExcelExtensionsTest` that exercises this
package rather than the binding: the typed contents of all four slicer and timeline parts and
their `mc:Ignorable`, a save that is Excel's shape and reads back the same, the control properties
part and the kept `x:controls`, the two drawings, and the `x14`/`x15` `extLst` content of the
workbook, sheet 2 and the styles part. `cr022-checkbox.xlsx` is copied from docx4j with
provenance (11 KB); `ignorable.test.mjs` picks it up with the rest.

**Deliberately not mirrored**, with the reason: `ExcelExtensionPartsTest.dataModelPartTyped` and
`ExcelExtensionsTest.dataModelWorkbookTypedAndSavesItsConnections` would want
`cr022-data-model.xlsx`, 245 KB for a binary part whose size is asserted, where phase A's built
package already covers `DataModelPart`; and `sparklinesTyped`, `conditionalFormattingsTyped`,
`dataValidationsTyped` and `x12acAndX16MarshalWithExcelsPrefixes` test the objects package's
binding, which has its own tests there, over three more fixtures. The slicers fixture already
carries sparklines, and the extLst test reads its `x14`/`x15` content.

## 7. The inert check box (acceptance check 14, 2026-09-26)

Jason's Excel 365 verdict on the two files staged at `fidelity/cr004b-core-ts-controls/`:
`cr022-slicers-timelines-coretsSaved.xlsx` opens and looks right, so the slicer and timeline half
of phase B is accepted. `cr022-checkbox-coretsSaved.xlsx` opens with **no repair prompt** and the
check box is drawn, but **clicking it does nothing** - and a form-control check box toggles its own
mark whether or not a cell is linked, and this one links none, so that is inert rather than merely
unwired.

**The cause is not yet isolated, and it is not safely attributable to the bare `x:controls`**: a
re-marshal changes more than one thing at a time. Measured part by part against the fixture (only
the differences that could bear on behaviour; `[Content_Types].xml`, the four relationships parts
and the untouched `docProps`, `sharedStrings`, `styles`, `theme` and `workbook` parts aside):

| | fixture | core-ts saved |
|---|---|---|
| `sheet1.xml` controls | `mc:AlternateContent > mc:Choice Requires="x14" > controls` | bare `controls` |
| booleans (`sheetView@tabSelected`, `controlPr@autoFill`/`autoLine`/`autoPict`/`defaultSize`, `anchor@moveWithCells`) | `1` / `0` | `true` / `false` |
| `ctrlProps/ctrlProp1.xml` | `formControlPr` in the default namespace, `lockText="1" noThreeD="1"` | `x14:formControlPr`, `lockText="true" noThreeD="true"` |
| `drawings/vmlDrawing1.vml` | | **byte-identical** (it is a `BinaryPart`, never re-marshalled) |
| `drawings/drawing1.xml` | one `mc:AlternateContent`, Choice `Requires="a14"`, **empty** `mc:Fallback` | the Choice, resolved |
| `a14:compatExt/@spid`, `a14:hiddenFill`, `a14:hiddenLine`, `a16:creationId` | | all present |
| `a:srgbClr/@a14:legacySpreadsheetColorIndex` (and its attribute-level `mc:Ignorable="a14"`) | on both hidden colours | **dropped** |

Four of those rule themselves out. The VML is byte-identical, and it is the VML's
`x:ClientData ObjectType="Checkbox"` that makes the shape a live control. `a14:compatExt/@spid`
`_x0000_s1025` survives, so the DrawingML shape is still tied to that VML shape and to
`control/@shapeId`. The drawing's own `mc:AlternateContent` had an **empty** `mc:Fallback`, so
resolving it to the `a14` Choice is what Excel does too and loses nothing.
`a14:legacySpreadsheetColorIndex` is a real loss but it is the hidden fill and line colours of a
shape Excel draws from the VML. `CT_SRgbColor` has no `xsd:anyAttribute` in docx4j's
`xsd/dml/dml-baseTypes.xsd` line 357, so nothing binds that attribute in either port - but **the
cause is not the binding gap alone, it is the MCE preprocessing, and that is worth knowing generally**
(measured 2026-09-27, after the objects session reported that the attribute round-trips through its
facade and does not reproduce on its corpus - both measurements were right):

| `xl/drawings/drawing1.xml` of `cr022-checkbox.xlsx`, unmarshalled and re-marshalled | occurrences of the attribute |
|---|---|
| through the objects facade alone | 2 in, **2 out** |
| with `resolveAlternateContent` run first, as every part loaded here is | 2 in, **0 out** |

Unresolved, the whole `xdr:twoCellAnchor` sits inside `mc:AlternateContent`, whose branches the model
holds as DOM (the `mce` wildcards being lax since docx4j CR-021), so nothing in the shape is typed
and nothing can be dropped. Resolving the Choice hands the same content to the typed `xdr` model,
`a14:hiddenFill` and its `a:srgbClr` included, and the unbound attribute goes. **So resolving markup
compatibility converts DOM into typed content, and can therefore expose a binding gap that the
unresolved form hid.** That is a property of the preprocessor rather than of this fixture: any
`mc:Choice` whose content the model binds *incompletely* loses the unbound part on a re-marshal,
where leaving the Choice alone would have kept all of it. It is the cost side of the judgment
CR-001 section 17.4 and section 3 of this CR make when they take a Choice, and the reason the
question there is "does everything in it round-trip" rather than "can the model read it".

Nothing here changes: the attribute is two hidden colours on a shape Excel redraws from the VML, and
the alternative - not resolving - would cost the slicers and the form controls that phase B recovered.
Recorded so that the next reader of a "docx4j drops this too" note knows it is the weaker half of the
explanation.

**There is a second, sharper form of the same cost**, found the same day and written up as CR-001
section 21: resolving a Choice can move *bound* content to a position its parent does not accept, and
then all of it goes. `xl/workbook.xml`'s `x15ac:absPath` is the case - bound as `CTAbsolutePath`, and
kept when the Choice is left alone, gone when it is resolved, because `CT_Workbook` accepts it only
inside the wrapper. Where the model binds `mc:AlternateContent` itself, as it does on `CT_Workbook`
and `CT_Worksheet`, resolving is not free.

That leaves two candidates, and the docx4j session put a one-change variant of each on the share
for Jason - `cr022-checkbox-A-wrapper-restored.xlsx` and `cr022-checkbox-B-booleans-01.xlsx` - plus
the discriminator that matters: **docx4j's own re-save**, which keeps the wrapper *and* writes
`true`/`false`, because JAXB serialises `xsd:boolean` the way Jsonix does. So:

- docx4j's re-save inert too -> the booleans are the cause, the defect is shared by both ports,
  and the fix belongs in the marshalling layer (a lexical-form choice for `xsd:boolean`), not in
  CR-004. Both lexical forms are valid XML Schema, so this would be Excel reading `@autoFill` and
  friends with a `== "1"` test.
- docx4j's re-save fine -> the bare `x:controls` is the cause, and the remedy here is for the
  worksheet to write the `x14` wrapper back on save: a narrow exception to resolving markup
  compatibility on load, which nothing else in the port has needed.

**Round 1 (2026-09-26): neither, and a lesson about writing a wrapper.** Variant B (booleans
`1`/`0`, bare `controls`) opens and is inert; docx4j's own re-save (wrapper kept, JAXB booleans
`true`/`false`) is inert too. Variant A errored on open, and the reason is a constraint on any
future fix here: resolving the Choice leaves nothing using the `x14` prefix, so the facade drops
`xmlns:x14` from the worksheet root - correctly, an unused declaration - and a wrapper put back
then carries `Requires="x14"` naming a prefix nothing declares. Excel treats that as an XML error,
not a repairable one. It is exactly `mc:Ignorable`'s rule (CR-001 section 17.3, docx4j CR-023) one
attribute over: **`Requires` is a prefix reference too, so writing the wrapper means declaring the
prefix on the root in the same breath.** `XmlPart.declareIgnorablePrefixes` is where that would
live. Note this does not bite `mc:Ignorable` here - the sheet's is `x14ac xr xr2 xr3`, all still
declared - which is why nothing caught it before.

**What round 1 does not settle.** A (wrapper only) and B (booleans only) each change one thing, so
between them they never test the **conjunction**: if Excel needs the wrapper *and* `1`/`0`, both
are inert and neither says so. The docx4j session's round 2 adds the right control -
`cr022-checkbox-ORIGINAL-untouched.xlsx`, to find out whether the fixture's check box ever toggled
(it is docx4j-generated and Excel-re-saved, so that is a real possibility) - and this session added
`cr022-checkbox-C-everything-restored.xlsx`: the wrapper in Excel's nesting, `xmlns:x14` declared,
the six booleans `1`/`0`, and `ctrlProp1.xml` and `drawing1.xml` restored to Excel's bytes, with
everything else still core-ts-marshalled. Every variant was checked namespace-well-formed before
staging, after A.

**Round 2 (2026-09-26): the fixture was never live, and every verdict so far was worthless.** The
untouched `cr022-checkbox.xlsx` does not toggle either, nor does A2. So `cr022-checkbox.xlsx` is
not a witness for this question at all: it is docx4j-generated (`Excel2010ExtensionsSamples`) and
only ever *re-saved* by Excel, its `x14:formControlPr` has no `fmlaLink` and no `checked`, and its
VML `x:ClientData` no `x:FmlaLink` - a check box nobody had clicked before this week. B, docx4j's
re-save, A2 and C all inherit that, so none of them can discriminate.

The lesson is about the evidence, not the code: **an acceptance check needs a fixture whose feature
is known to work before the save.** This one asked "does the control still work?" of a control that
never worked, and three rounds of bisecting measured nothing. Two of the four findings survive
anyway, because they do not depend on the verdict: the `Requires` prefix rule above, and the
`a14:legacySpreadsheetColorIndex` loss, both real and both independently checkable.

**Round 3 (2026-09-26): accepted, and resolve-on-load stands.** Round 3 moved to `cr022-checkbox-linked.xlsx`, the one check box Excel has been seen
to toggle (docx4j CR-026 section 11's bisect: `fmlaLink="$D$4"` on the properties part and
`x:FmlaLink` in the VML, so a click writes `TRUE`/`FALSE` into the empty D4). Two files on the
share, same base, and between them they separate the wrapper from everything else: the docx4j
session's `cr022-checkbox-linked-BARE-controls.xlsx` (byte-identical but for the two wrappers
removed, `xmlns:x14` still declared) and this session's `cr022-checkbox-linked-coretsSaved.xlsx`
(the full pipeline, every part unmarshalled and saved - which drops the wrappers *and* `xmlns:x14`,
writes the booleans `true`/`false` and prefixes the properties root, while keeping `fmlaLink`, the
VML byte-identical and `control/@shapeId` 1025 tied to `a14:compatExt/@spid _x0000_s1025`). Both
toggle and drive D4: resolve-on-load stands and phase B needs nothing. Neither: the wrapper is the
cause and it gets written back, `xmlns:x14` with it. Only the bare one: the wrapper is exonerated
and the cause is elsewhere in what core-ts re-marshals, the booleans first.

**Both toggled and both drove D4.** So Excel accepts a worksheet's controls with the wrapper gone,
`xmlns:x14` gone from the root, the booleans `true`/`false` and the properties root prefixed `x14:`.
The `mc:AlternateContent` Excel writes there is forward-compatibility markup for readers that do
not know `x14`, not a requirement of the ones that do, and **nothing in phase B changes**: the
worksheet keeps writing the resolved `x:controls`, and there is no exception to resolving markup
compatibility on load. docx4j records the same in its CR-022 section 20 (closed) and CR-021
section 8.13.

`cr022-checkbox-linked.xlsx` is copied in as a fixture even so, and
`sml-extensions.test.mjs` has a tenth test over it: `fmlaLink="$D$4"` survives a whole-package
re-marshal, the VML keeps its `x:FmlaLink` half of the link by never being re-marshalled at all,
and `control/@shapeId` 1025 still resolves to the properties part and matches
`a14:compatExt/@spid`. That is the machine-checkable half of what Excel confirmed, on the one file
whose Excel behaviour is known - the lesson of section 7 applied to our own suite rather than only
to the manual checklist.

Nothing is changed here until those verdicts arrive. What phase B keeps either way is the read:
`worksheet.controls` holds the control and its `r:id` resolves to the `ControlPropertiesPart`,
where before phase B the whole `mc:AlternateContent` was dropped and neither existed. The docx4j
session records the same in its CR-022 section 20.
