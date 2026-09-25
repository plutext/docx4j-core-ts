# CR-004: SpreadsheetML extension parts (Excel 2010 and 2013: slicers, timelines, control properties, custom data, survey, data model)

**Status:** Phase A implemented 2026-09-24 (section 5). Phase B unblocked 2026-09-20 by
objects 0.1.6 (the docx4j CR-022 regeneration, the `x14` and `x15` Excel modules; 0.1.5 carried
it but could not load a pptx or xlsx whose text body holds an equation, CR-001 section 17.5),
which this package now depends on.
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
