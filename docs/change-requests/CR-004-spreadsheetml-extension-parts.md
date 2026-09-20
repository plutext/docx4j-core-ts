# CR-004: SpreadsheetML extension parts (Excel 2010 and 2013: slicers, timelines, control properties, custom data, survey, data model)

**Status:** Proposed 2026-09-20. Phase A ready to implement; Phase B unblocked the same day by
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

## 4. Not in scope

The slicer, timeline and control content itself (a SpreadsheetML content API is not planned); the
DrawingML `mc:AlternateContent` in drawings (`a14`, `tsle`, `sle15`), which docx4j still resolves to
the Fallback and lists as a DrawingML CR.
