// CR-004 phase A: the Excel 2010 and 2013 extension parts (slicers, timelines, form control
// properties, custom data, survey, data model), the counterpart of docx4j's
// org.docx4j.openpackaging.parts.SpreadsheetML classes and its ExcelExtensionPartsTest.
// `cr022-slicers-timelines.xlsx` covers the four Excel writes for slicers and timelines; the
// other five are built here, saved and reloaded, which is what exercises their registry entries.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  OpcPackage, SpreadsheetMLPackage, ZipPartStore, ContentTypes, Namespaces, PartName,
  SlicerCachePart, SlicersPart, TimelineCachePart, TimelinesPart, ControlPropertiesPart,
  CustomDataPropertiesPart, CustomDataPart, SurveyPart, DataModelPart,
} from '../dist/index.mjs';
import { fixture } from './helpers.mjs';

const FIXTURE = 'cr022-slicers-timelines.xlsx';

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
  const types = new TextDecoder().decode(saved.loadSync('[Content_Types].xml'));
  for (const contentType of [ContentTypes.SPREADSHEETML_SLICER_CACHE, ContentTypes.SPREADSHEETML_SLICERS,
    ContentTypes.SPREADSHEETML_TIMELINE_CACHE, ContentTypes.SPREADSHEETML_TIMELINES]) {
    assert.ok(types.includes(`ContentType="${contentType}"`), contentType);
  }
  const workbookRels = new TextDecoder().decode(saved.loadSync('xl/_rels/workbook.xml.rels'));
  assert.equal(workbookRels.split(Namespaces.SPREADSHEETML_SLICER_CACHE).length - 1, 2);
  assert.ok(workbookRels.includes(Namespaces.SPREADSHEETML_TIMELINE_CACHE));
  const sheetRels = new TextDecoder().decode(saved.loadSync('xl/worksheets/_rels/sheet2.xml.rels'));
  assert.ok(sheetRels.includes(Namespaces.SPREADSHEETML_SLICERS));
  assert.ok(sheetRels.includes(Namespaces.SPREADSHEETML_TIMELINES));

  // reloading the saved package finds the same parts
  const back = await OpcPackage.load(await pkg.save());
  assert.equal(back.getWorkbookPart().slicerCacheParts.length, 2);
  assert.equal(back.getWorkbookPart().timelineCacheParts.length, 1);
});

test('the five parts the fixture has no example of: added, saved, and found again by class', async () => {
  const pkg = await SpreadsheetMLPackage.createPackage();
  const workbook = pkg.getWorkbookPart();
  const sheet = pkg.createWorksheetPart('Sheet1');

  const survey = new SurveyPart();
  survey.setXml('<survey xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" guid="{00000000-0000-0000-0000-000000000001}"/>');
  workbook.addTargetPart(survey);

  const dataModel = new DataModelPart();
  dataModel.setBytes(new Uint8Array([1, 2, 3, 4]));
  workbook.addTargetPart(dataModel);

  const customDataProps = new CustomDataPropertiesPart();
  customDataProps.setXml('<datastoreItem xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" itemID="{00000000-0000-0000-0000-000000000002}"/>');
  workbook.addTargetPart(customDataProps);

  const customData = new CustomDataPart();
  customData.setBytes(new Uint8Array([5, 6, 7]));
  customDataProps.addTargetPart(customData);

  const controlProps = new ControlPropertiesPart();
  controlProps.setXml('<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" checked="Checked"/>');
  sheet.addTargetPart(controlProps);

  const back = await OpcPackage.load(await pkg.save());
  const backWorkbook = back.getWorkbookPart();
  assert.ok(backWorkbook.surveyPart instanceof SurveyPart);
  assert.ok(backWorkbook.dataModelPart instanceof DataModelPart);
  assert.deepEqual(await backWorkbook.dataModelPart.getBytes(), new Uint8Array([1, 2, 3, 4]));

  const props = backWorkbook.customDataPropertiesParts;
  assert.equal(props.length, 1);
  assert.ok(props[0] instanceof CustomDataPropertiesPart);
  // chosen by relationship type: its content type is the generic application/binary
  const data = back.parts.get('/xl/customData/customData1.dat');
  assert.ok(data instanceof CustomDataPart, `expected CustomDataPart, got ${data?.constructor.name}`);
  assert.equal(data.contentType, ContentTypes.SPREADSHEETML_CUSTOM_DATA);
  assert.deepEqual(await data.getBytes(), new Uint8Array([5, 6, 7]));

  const backSheet = (await backWorkbook.getWorksheetParts())[0];
  assert.deepEqual(backSheet.controlPropertiesParts.map((p) => p.partName.name), ['/xl/ctrlProps/ctrlProp1.xml']);
  assert.ok(backSheet.controlPropertiesParts[0] instanceof ControlPropertiesPart);

  const types = new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync('[Content_Types].xml'));
  for (const contentType of [ContentTypes.SPREADSHEETML_SURVEY, ContentTypes.SPREADSHEETML_DATA_MODEL,
    ContentTypes.SPREADSHEETML_CUSTOM_DATA_PROPERTIES, ContentTypes.SPREADSHEETML_CONTROL_PROPERTIES]) {
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
});

// The gap CR-004 phase A found and docx4j 8e8f6ea83 closed (CR-004 section 5): a:graphicData's
// wildcard was strict, so a graphic no module binds was fatal rather than DOM. drawing1.xml
// frames a slicer through a Choice requiring a14, which the preprocessor takes; drawing2.xml
// frames one through a Choice requiring sle15, which it does not understand, so that one keeps
// its Fallback picture.
test('a drawing that frames a slicer unmarshals, with the slicer kept as DOM', async () => {
  const pkg = await OpcPackage.load(await fixture(FIXTURE));
  const drawing = pkg.parts.get('/xl/drawings/drawing1.xml');
  const contents = await drawing.getContents();
  const domNodes = [];
  const walk = (v) => {
    if (!v || typeof v !== 'object') return;
    if (v.nodeType) { domNodes.push(v.nodeName); return; }
    for (const k of Object.keys(v)) if (k !== 'PARENT') walk(v[k]);
  };
  walk(contents);
  assert.deepEqual(domNodes, ['sle:slicer']);

  const out = new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync('xl/drawings/drawing1.xml'));
  assert.ok(out.includes('<sle:slicer'), 're-marshalled with its slicer');
  assert.ok(!out.includes('<mc:Choice'), 'the understood a14 branch was taken on load');
});

// Excel writes mc:Ignorable="x xr10" on these parts, naming a prefix whose namespace it declares
// as the default; objects CR-006 declares such a prefix rather than dropping it.
test('a re-marshalled slicer part keeps Excel\'s mc:Ignorable', async () => {
  const pkg = await OpcPackage.load(await fixture(FIXTURE));
  for (const part of [...pkg.parts.values()].filter((p) => p instanceof SlicersPart || p instanceof SlicerCachePart)) {
    await part.getDocument();
  }
  const saved = new ZipPartStore(await pkg.save());
  for (const name of ['xl/slicers/slicer1.xml', 'xl/slicerCaches/slicerCache1.xml']) {
    const xml = new TextDecoder().decode(saved.loadSync(name));
    const root = xml.match(/<[^?!][^>]*>/)[0];
    const ignorable = /mc:Ignorable="([^"]*)"/.exec(root)?.[1];
    assert.equal(ignorable, 'x xr10', name);
    for (const prefix of ignorable.split(' ')) assert.ok(root.includes(`xmlns:${prefix}="`), `${name}: xmlns:${prefix}`);
  }
});
