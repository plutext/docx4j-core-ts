// CR-004: the Excel 2010 and 2013 extension parts (slicers, timelines, form control properties,
// custom data, survey, data model), the counterpart of docx4j's
// org.docx4j.openpackaging.parts.SpreadsheetML classes and of its ExcelExtensionPartsTest and
// ExcelExtensionsTest.  Phase A gave them their nine classes, phase B their typed x14 and x15
// roots and the three slicer namespaces the preprocessor now understands.
// `cr022-slicers-timelines.xlsx` covers the four Excel writes for slicers and timelines and
// `cr022-checkbox.xlsx` the form control; the other four are built here, saved and reloaded,
// which is what exercises their registry entries.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  OpcPackage, SpreadsheetMLPackage, ZipPartStore, ContentTypes, Namespaces, PartName,
  SlicerCachePart, SlicersPart, TimelineCachePart, TimelinesPart, ControlPropertiesPart,
  CustomDataPropertiesPart, CustomDataPart, SurveyPart, DataModelPart,
} from '../dist/index.mjs';
import { fixture } from './helpers.mjs';

const FIXTURE = 'cr022-slicers-timelines.xlsx';
const CHECKBOX = 'cr022-checkbox.xlsx';
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';

/** The ext content of a given element name, unwrapped (docx4j's ExcelExtensionsTest.ext). */
function ext(extLst, localPart) {
  assert.ok(extLst, 'no extLst');
  const found = extLst.ext.find((e) => e.any?.name?.localPart === localPart);
  assert.ok(found, `no ${localPart} in extLst; found ${extLst.ext.map((e) => e.any?.name?.localPart).join(', ')}`);
  return found.any.value;
}

function entry(store, name) {
  return new TextDecoder().decode(store.loadSync(name));
}

/** Every prefix an mc:Ignorable names is declared on the root that names it. */
function assertDeclares(store, name, ignorable) {
  const xml = entry(store, name);
  const root = xml.match(/<[^?!][^>]*>/)[0];
  assert.equal(/mc:Ignorable="([^"]*)"/.exec(root)?.[1], ignorable, name);
  for (const prefix of ignorable.split(' ')) assert.ok(root.includes(`xmlns:${prefix}="`), `${name}: xmlns:${prefix}`);
}

test('slicer and timeline parts load with their classes and shortcuts', async () => {
  const pkg = await OpcPackage.load(await fixture(FIXTURE));
  assert.ok(pkg instanceof SpreadsheetMLPackage);
  const workbook = pkg.getWorkbookPart();

  const caches = workbook.slicerCacheParts;
  assert.equal(caches.length, 2);
  for (const part of caches) {
    assert.ok(part instanceof SlicerCachePart);
    assert.equal(part.contentType, ContentTypes.SPREADSHEETML_SLICER_CACHE);
    assert.equal(part.relationshipType, Namespaces.SPREADSHEETML_SLICER_CACHE);
  }
  assert.deepEqual(caches.map((p) => p.partName.name).sort(),
    ['/xl/slicerCaches/slicerCache1.xml', '/xl/slicerCaches/slicerCache2.xml']);

  const timelineCaches = workbook.timelineCacheParts;
  assert.equal(timelineCaches.length, 1);
  assert.ok(timelineCaches[0] instanceof TimelineCachePart);
  assert.equal(timelineCaches[0].partName.name, '/xl/timelineCaches/timelineCache1.xml');

  // The slicers and timelines hang off the sheets that show them, not off the workbook.
  const bySheet = new Map(workbook.worksheetParts.map((s) => [s.partName.name, s]));
  assert.deepEqual(bySheet.get('/xl/worksheets/sheet2.xml').slicersParts.map((p) => p.partName.name),
    ['/xl/slicers/slicer1.xml']);
  assert.deepEqual(bySheet.get('/xl/worksheets/sheet3.xml').slicersParts.map((p) => p.partName.name),
    ['/xl/slicers/slicer2.xml']);
  assert.deepEqual(bySheet.get('/xl/worksheets/sheet2.xml').timelinesParts.map((p) => p.partName.name),
    ['/xl/timelines/timeline1.xml']);
  assert.equal(bySheet.get('/xl/worksheets/sheet1.xml').slicersParts.length, 0);
  assert.equal(bySheet.get('/xl/worksheets/sheet1.xml').timelinesParts.length, 0);
  for (const part of [...bySheet.values()].flatMap((s) => s.slicersParts)) {
    assert.ok(part instanceof SlicersPart);
    assert.equal(part.contentType, ContentTypes.SPREADSHEETML_SLICERS);
  }
  assert.ok(bySheet.get('/xl/worksheets/sheet2.xml').timelinesParts[0] instanceof TimelinesPart);
});

// docx4j ExcelExtensionPartsTest.slicerAndTimelinePartsTyped: phase B's half of the CR. The
// contents are the x14 and x15 roots, not a DOM, and Excel's mc:Ignorable is a property of each.
test('the slicer and timeline parts unmarshal typed', async () => {
  const pkg = await OpcPackage.load(await fixture(FIXTURE));
  await assertTypedSlicersAndTimelines(pkg);
});

async function assertTypedSlicersAndTimelines(pkg) {
  const workbook = pkg.getWorkbookPart();
  for (const part of workbook.slicerCacheParts) {
    const def = await part.getContents();
    assert.equal(def.TYPE_NAME, 'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main.CTSlicerCacheDefinition');
    assert.ok(def.name);
    assert.ok(def.sourceName);
    assert.equal(def.ignorable, 'x xr10');
    assert.equal(part.typedElement.name.namespaceURI, X14);
    assert.equal(part.typedElement.name.localPart, 'slicerCacheDefinition');
  }
  assert.deepEqual((await Promise.all(workbook.slicerCacheParts.map((p) => p.getContents()))).map((d) => d.name).sort(),
    ['Slicer_Region', 'Slicer_Region1']);

  const timelineCache = await workbook.timelineCacheParts[0].getContents();
  assert.equal(timelineCache.TYPE_NAME, 'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_main.CTTimelineCacheDefinition');
  assert.equal(timelineCache.name, 'NativeTimeline_Date');
  assert.equal(timelineCache.sourceName, 'Date');
  assert.equal(timelineCache.ignorable, 'xr10');
  assert.ok(timelineCache.pivotTables.pivotTable.length >= 1);

  const sheet2 = (await workbook.getWorksheetParts())[1];
  const slicers = await sheet2.slicersParts[0].getContents();
  assert.equal(slicers.slicer.length, 1);
  assert.equal(slicers.slicer[0].cache, 'Slicer_Region1');
  assert.equal(slicers.slicer[0].caption, 'Region');
  assert.equal(slicers.ignorable, 'x xr10');

  const timelines = await sheet2.timelinesParts[0].getContents();
  assert.equal(timelines.timeline.length, 1);
  assert.equal(timelines.timeline[0].cache, 'NativeTimeline_Date');

  // and by part name, as docx4j asserts
  assert.ok(pkg.parts.get('/xl/slicers/slicer1.xml') instanceof SlicersPart);
  assert.ok(pkg.parts.get('/xl/slicerCaches/slicerCache1.xml') instanceof SlicerCachePart);
  assert.ok(pkg.parts.get('/xl/timelines/timeline1.xml') instanceof TimelinesPart);
  assert.ok(pkg.parts.get('/xl/timelineCaches/timelineCache1.xml') instanceof TimelineCachePart);
}

// docx4j ExcelExtensionPartsTest.partsSurviveASave. Excel writes mc:Ignorable="x xr10" on these
// roots, naming a prefix (x) whose namespace it declares as the default; objects CR-006 declares
// such a prefix rather than dropping it, and XmlPart re-declares the ones no module binds.
test('a re-marshalled slicer part is Excel\'s shape, with its mc:Ignorable declared', async () => {
  const pkg = await OpcPackage.load(await fixture(FIXTURE));
  await assertTypedSlicersAndTimelines(pkg);  // unmarshals all four, so all four are re-marshalled
  const saved = new ZipPartStore(await pkg.save());

  assertDeclares(saved, 'xl/slicers/slicer1.xml', 'x xr10');
  assertDeclares(saved, 'xl/slicerCaches/slicerCache1.xml', 'x xr10');
  assertDeclares(saved, 'xl/timelines/timeline1.xml', 'x xr10');
  assertDeclares(saved, 'xl/timelineCaches/timelineCache1.xml', 'xr10');
  assert.ok(entry(saved, 'xl/slicers/slicer1.xml').includes('<x14:slicers '), 'the x14 root');
  assert.ok(entry(saved, 'xl/timelines/timeline1.xml').includes('<x15:timelines '), 'the x15 root');

  // and the saved package reads back the same way
  await assertTypedSlicersAndTimelines(await OpcPackage.load(await pkg.save()));
});

test('a workbook with slicers and timelines round-trips byte for byte', async () => {
  const bytes = await fixture(FIXTURE);
  const pkg = await OpcPackage.load(bytes);
  const source = new ZipPartStore(bytes);
  const saved = new ZipPartStore(await pkg.save());
  for (const part of pkg.parts.values()) {
    const name = part.partName.name.slice(1);
    assert.deepEqual(saved.loadSync(name), source.loadSync(name), name);
  }
  // and the content-type overrides and relationships that name them survive
  const types = entry(saved, '[Content_Types].xml');
  for (const contentType of [ContentTypes.SPREADSHEETML_SLICER_CACHE, ContentTypes.SPREADSHEETML_SLICERS,
    ContentTypes.SPREADSHEETML_TIMELINE_CACHE, ContentTypes.SPREADSHEETML_TIMELINES]) {
    assert.ok(types.includes(`ContentType="${contentType}"`), contentType);
  }
  const workbookRels = entry(saved, 'xl/_rels/workbook.xml.rels');
  assert.equal(workbookRels.split(Namespaces.SPREADSHEETML_SLICER_CACHE).length - 1, 2);
  assert.ok(workbookRels.includes(Namespaces.SPREADSHEETML_TIMELINE_CACHE));
  const sheetRels = entry(saved, 'xl/worksheets/_rels/sheet2.xml.rels');
  assert.ok(sheetRels.includes(Namespaces.SPREADSHEETML_SLICERS));
  assert.ok(sheetRels.includes(Namespaces.SPREADSHEETML_TIMELINES));

  // reloading the saved package finds the same parts
  const back = await OpcPackage.load(await pkg.save());
  assert.equal(back.getWorkbookPart().slicerCacheParts.length, 2);
  assert.equal(back.getWorkbookPart().timelineCacheParts.length, 1);
});

// docx4j ExcelExtensionPartsTest.controlPropertiesPartTyped, and the half of
// ExcelExtensionsTest.formControlAlternateContentKept this port can answer: Excel wraps a
// worksheet's controls in an mc:AlternateContent whose only Choice requires x14 and which has no
// Fallback, so before phase B made x14 understood the whole thing - the control, and the r:id
// naming its ControlPropertiesPart - was dropped from a re-marshalled sheet. The preprocessor
// resolves the branch on load, so the saved sheet holds a bare x:controls (schema-valid; the
// wrapper cannot be rebuilt from the resolved tree) rather than Excel's mc:AlternateContent.
test('a form control\'s properties part is typed, and the sheet keeps its x14 controls', async () => {
  const pkg = await OpcPackage.load(await fixture(CHECKBOX));
  const sheet = pkg.getWorkbookPart().getWorksheet(0);
  const part = sheet.controlPropertiesParts[0];
  assert.ok(part instanceof ControlPropertiesPart);
  assert.equal(part.partName.name, '/xl/ctrlProps/ctrlProp1.xml');
  const props = await part.getContents();
  assert.equal(props.TYPE_NAME, 'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main.CTFormControlPr');
  assert.equal(props.objectType, 'CheckBox');

  const controls = (await sheet.getContents()).controls;
  assert.equal(controls.control.length, 1);
  assert.equal(controls.control[0].name, 'Check Box 1');
  // the r:id is what ties the control to the properties part
  assert.equal(sheet.relationshipsPart.getPart(controls.control[0].id), part);

  const saved = new ZipPartStore(await pkg.save());
  const xml = entry(saved, 'xl/worksheets/sheet1.xml');
  assert.ok(xml.includes('<controls>') || xml.includes('<controls '), 'the controls survived the round trip');
  assert.ok(!xml.includes('<mc:AlternateContent'), 'its branch was chosen on load');
  assert.ok(entry(saved, '[Content_Types].xml').includes(ContentTypes.SPREADSHEETML_CONTROL_PROPERTIES));

  const back = await OpcPackage.load(await pkg.save());
  const backSheet = back.getWorkbookPart().getWorksheet(0);
  assert.ok(back.parts.get('/xl/ctrlProps/ctrlProp1.xml') instanceof ControlPropertiesPart);
  assert.equal((await backSheet.getContents()).controls.control.length, 1);
  assert.equal((await backSheet.controlPropertiesParts[0].getContents()).objectType, 'CheckBox');
});

test('the four parts the fixtures have no example of: added, saved, and found again by class', async () => {
  const pkg = await SpreadsheetMLPackage.createPackage();
  const workbook = pkg.getWorkbookPart();

  const survey = new SurveyPart();
  survey.setXml(`<survey xmlns="${X15}" guid="{00000000-0000-0000-0000-000000000001}" id="1" title="Survey 1">`
    + '<questions><question binding="1" type="singleLineOfText" text="Name"/></questions></survey>');
  workbook.addTargetPart(survey);

  const dataModel = new DataModelPart();
  dataModel.setBytes(new Uint8Array([1, 2, 3, 4]));
  workbook.addTargetPart(dataModel);

  const customDataProps = new CustomDataPropertiesPart();
  customDataProps.setXml(`<datastoreItem xmlns="${X14}" id="{00000000-0000-0000-0000-000000000002}"/>`);
  workbook.addTargetPart(customDataProps);

  const customData = new CustomDataPart();
  customData.setBytes(new Uint8Array([5, 6, 7]));
  customDataProps.addTargetPart(customData);

  const back = await OpcPackage.load(await pkg.save());
  const backWorkbook = back.getWorkbookPart();
  assert.ok(backWorkbook.surveyPart instanceof SurveyPart);
  const surveyContents = await backWorkbook.surveyPart.getContents();
  assert.equal(surveyContents.guid, '{00000000-0000-0000-0000-000000000001}');
  assert.equal(surveyContents.questions.question[0].text, 'Name');

  assert.ok(backWorkbook.dataModelPart instanceof DataModelPart);
  assert.deepEqual(await backWorkbook.dataModelPart.getBytes(), new Uint8Array([1, 2, 3, 4]));

  const props = backWorkbook.customDataPropertiesParts;
  assert.equal(props.length, 1);
  assert.ok(props[0] instanceof CustomDataPropertiesPart);
  assert.equal((await props[0].getContents()).id, '{00000000-0000-0000-0000-000000000002}');
  // chosen by relationship type: its content type is the generic application/binary
  const data = back.parts.get('/xl/customData/customData1.dat');
  assert.ok(data instanceof CustomDataPart, `expected CustomDataPart, got ${data?.constructor.name}`);
  assert.equal(data.contentType, ContentTypes.SPREADSHEETML_CUSTOM_DATA);
  assert.deepEqual(await data.getBytes(), new Uint8Array([5, 6, 7]));

  const types = entry(new ZipPartStore(await pkg.save()), '[Content_Types].xml');
  for (const contentType of [ContentTypes.SPREADSHEETML_SURVEY, ContentTypes.SPREADSHEETML_DATA_MODEL,
    ContentTypes.SPREADSHEETML_CUSTOM_DATA_PROPERTIES]) {
    assert.ok(types.includes(`ContentType="${contentType}"`), contentType);
  }
});

test('the nine classes carry docx4j\'s default part names, content types and relationship types', () => {
  const cases = [
    [new SlicerCachePart(), '/xl/slicerCaches/slicerCache1.xml', ContentTypes.SPREADSHEETML_SLICER_CACHE, Namespaces.SPREADSHEETML_SLICER_CACHE],
    [new SlicersPart(), '/xl/slicers/slicer1.xml', ContentTypes.SPREADSHEETML_SLICERS, Namespaces.SPREADSHEETML_SLICERS],
    [new TimelineCachePart(), '/xl/timelineCaches/timelineCache1.xml', ContentTypes.SPREADSHEETML_TIMELINE_CACHE, Namespaces.SPREADSHEETML_TIMELINE_CACHE],
    [new TimelinesPart(), '/xl/timelines/timeline1.xml', ContentTypes.SPREADSHEETML_TIMELINES, Namespaces.SPREADSHEETML_TIMELINES],
    [new ControlPropertiesPart(), '/xl/ctrlProps/ctrlProp1.xml', ContentTypes.SPREADSHEETML_CONTROL_PROPERTIES, Namespaces.SPREADSHEETML_CONTROL_PROPERTIES],
    [new CustomDataPropertiesPart(), '/xl/customDataProps/customDataProps1.xml', ContentTypes.SPREADSHEETML_CUSTOM_DATA_PROPERTIES, Namespaces.SPREADSHEETML_CUSTOM_DATA_PROPERTIES],
    [new CustomDataPart(), '/xl/customData/customData1.dat', ContentTypes.SPREADSHEETML_CUSTOM_DATA, Namespaces.SPREADSHEETML_CUSTOM_DATA],
    [new SurveyPart(), '/xl/surveys/survey1.xml', ContentTypes.SPREADSHEETML_SURVEY, Namespaces.SPREADSHEETML_SURVEY],
    [new DataModelPart(), '/xl/model/item.data', ContentTypes.SPREADSHEETML_DATA_MODEL, Namespaces.SPREADSHEETML_DATA_MODEL],
  ];
  for (const [part, name, contentType, relationshipType] of cases) {
    assert.equal(part.partName.name, name);
    assert.equal(part.contentType, contentType);
    assert.equal(part.relationshipType, relationshipType);
  }
  // a name may still be given, as every other part class allows
  assert.equal(new SlicersPart(new PartName('/xl/slicers/slicer7.xml')).partName.name, '/xl/slicers/slicer7.xml');
  // and each XML one knows its root element, so setContents() on a new part needs no name
  assert.deepEqual(new SlicersPart().rootName, { namespaceURI: X14, localPart: 'slicers' });
  assert.deepEqual(new TimelinesPart().rootName, { namespaceURI: X15, localPart: 'timelines' });
});

// The slicer and timeline drawings, CR-004 phase B's second half (and the loss CR-001 section
// 17.4 recorded). Every Choice here holds an xdr:graphicFrame whose a:graphicData carries the one
// element no module binds; that wildcard is lax since docx4j 8e8f6ea83, so the Choice round-trips
// as DOM and is worth taking over the Fallback, which is only an "Excel 2010 or later" placeholder
// box. drawing1 frames a slicer (Requires a14, understood before phase B) and a timeline (tsle);
// drawing2 frames a slicer through sle15.
test('the slicer and timeline drawings keep their graphics, as DOM', async () => {
  const pkg = await OpcPackage.load(await fixture(FIXTURE));
  const domNodesOf = (contents) => {
    const names = [];
    const walk = (v) => {
      if (!v || typeof v !== 'object') return;
      if (v.nodeType) { names.push(v.nodeName); return; }
      for (const k of Object.keys(v)) if (k !== 'PARENT') walk(v[k]);
    };
    walk(contents);
    return names;
  };
  assert.deepEqual(domNodesOf(await pkg.parts.get('/xl/drawings/drawing1.xml').getContents()),
    ['sle:slicer', 'tsle:timeslicer']);
  assert.deepEqual(domNodesOf(await pkg.parts.get('/xl/drawings/drawing2.xml').getContents()),
    ['sle:slicer']);

  const saved = new ZipPartStore(await pkg.save());
  const one = entry(saved, 'xl/drawings/drawing1.xml');
  assert.ok(one.includes('<sle:slicer'), 're-marshalled with its slicer');
  assert.ok(one.includes('<tsle:timeslicer'), 're-marshalled with its timeline');
  assert.ok(!one.includes('<mc:Choice'), 'the branches were chosen on load');
  assert.ok(!one.includes('This shape represents a slicer'), 'the fallback placeholder was given up');
  assert.ok(entry(saved, 'xl/drawings/drawing2.xml').includes('<sle:slicer'));
});

// docx4j ExcelExtensionsTest.assertWorkbook / assertWorksheets / assertStylesAndTable: the x14 and
// x15 content Excel writes in an extLst is typed, and survives a round trip here too. The binding
// is the objects package's; what this asserts is that the parts read it and write it back.
test('the x14 and x15 extension content of the workbook, sheets and styles reads typed', async () => {
  const pkg = await OpcPackage.load(await fixture(FIXTURE));
  const workbook = pkg.getWorkbookPart();
  const wb = await workbook.getContents();
  assert.equal(wb.ignorable, 'x15 xr xr6 xr10 xr2');
  assert.ok(wb.revisionPtr?.documentId, 'xr:revisionPtr');
  assert.equal(wb.extLst.ext.filter((e) => e.any?.name?.localPart === 'slicerCaches').length, 2, 'x14 and x15 slicerCaches');
  assert.equal(ext(wb.extLst, 'timelineCacheRefs').timelineCacheRef.length, 1);

  const sheet2 = (await workbook.getWorksheetParts())[1];
  const sheet = await sheet2.getContents();
  assert.equal(sheet.ignorable, 'x14ac xr xr2 xr3');
  assert.ok(sheet.sheetFormatPr.dyDescent !== undefined, 'x14ac:dyDescent');
  assert.ok(ext(sheet.extLst, 'slicerList').slicer.length >= 1);
  assert.equal(ext(sheet.extLst, 'timelineRefs').timelineRef.length, 1);

  const styles = await workbook.stylesPart.getContents();
  assert.equal(styles.ignorable, 'x14ac x16r2 xr');
  assert.equal(styles.fonts.knownFonts, true);
  assert.ok(ext(styles.extLst, 'slicerStyles').defaultSlicerStyle);
  assert.ok(ext(styles.extLst, 'timelineStyles').defaultTimelineStyle);

  // and a save writes them with Excel's prefixes (ignorable.test.mjs checks every declaration)
  const saved = new ZipPartStore(await pkg.save());
  const sheetXml = entry(saved, 'xl/worksheets/sheet2.xml');
  assert.ok(sheetXml.includes('x14ac:dyDescent='));
  assert.ok(sheetXml.includes('<x14:slicerList'));
  assert.ok(sheetXml.includes('<x15:timelineRefs'));
  assert.ok(entry(saved, 'xl/workbook.xml').includes('<xr:revisionPtr '));
});
