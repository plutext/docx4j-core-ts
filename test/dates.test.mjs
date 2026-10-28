// CR-018: dates in the package's zone (the portfolio's date rule, Jason 2026-10-28; docx4j CR-067 D12, its
// PackageDateTimeZoneTest mirrored). An instant the package holds is shown, and a w:date without dateUtc read, in
// pkg.dateTimeZone, else the runtime's zone; a new comment is stamped to the minute; Word's conversion of a bound date
// is a separate opt-in; a save writes dcterms times only under writeTimesOnSave.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { WordprocessingMLPackage, ZipPartStore } from '../dist/index.mjs';
import { Word } from '../dist/office-js/index.mjs';
import { fixture } from './helpers.mjs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const W16DU = 'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"';
const partOf = async (pkg, name) => {
  const store = new ZipPartStore(await pkg.save());
  return store.has(name) ? new TextDecoder().decode(store.loadSync(name)) : undefined;
};
const runtimeZone = () => new Intl.DateTimeFormat().resolvedOptions().timeZone;
/** What a wall clock in `zone` shows for `date`, as `YYYY-MM-DDTHH:MM:SS` (what w:date carries before its Z). */
const clockIn = (date, zone) => {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-US', { timeZone: zone, hour12: false, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' })
    .formatToParts(date).map((x) => [x.type, x.value]));
  return `${p.year}-${p.month}-${p.day}T${String(Number(p.hour) % 24).padStart(2, '0')}:${p.minute}:${p.second}`;
};

test('CR-018: pkg.dateTimeZone is undefined by default, the runtime zone in force; a zone the runtime does not know is refused at the set', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  assert.equal(pkg.dateTimeZone, undefined);
  assert.equal(pkg.effectiveDateTimeZone(), runtimeZone());
  pkg.dateTimeZone = 'Asia/Tokyo';
  assert.equal(pkg.effectiveDateTimeZone(), 'Asia/Tokyo');
  assert.throws(() => { pkg.dateTimeZone = 'Not/AZone'; }, { name: 'RangeError', message: /Not a time zone this runtime knows: "Not\/AZone"/ });
  assert.equal(pkg.dateTimeZone, 'Asia/Tokyo', 'a refused set changes nothing');
  pkg.dateTimeZone = undefined;
  assert.equal(pkg.effectiveDateTimeZone(), runtimeZone());
  assert.equal(pkg.boundDatesConvertedAsWord, false, 'off by default');
  assert.equal(pkg.writeTimesOnSave, false, 'off by default');
  pkg.dateTimeZone = 'Asia/Tokyo';
  assert.equal(pkg.boundDatesConvertedAsWord, false, 'a zone alone does not turn the conversion on');
});

test("CR-018: a revision's w:date is the wall clock of the package's zone, its dateUtc the instant; the runtime's zone when none is set", async () => {
  const when = new Date('2024-02-03T04:05:06Z');
  const stamp = async (zone) => {
    const pkg = await WordprocessingMLPackage.createPackage();
    pkg.dateTimeZone = zone;
    pkg.trackedChangeDate = when;
    pkg.changeTrackingMode = 'TrackAll';
    const body = await pkg.getBody();
    body.insertParagraph('tracked', 'End');
    const xml = await partOf(pkg, 'word/document.xml');
    return { date: /w:date="([^"]*)"/.exec(xml)[1], utc: /w16du:dateUtc="([^"]*)"/.exec(xml)[1] };
  };
  assert.deepEqual(await stamp('Asia/Tokyo'), { date: '2024-02-03T13:05:06Z', utc: '2024-02-03T04:05:06Z' }, 'UTC+9, no daylight time');
  assert.deepEqual(await stamp('America/New_York'), { date: '2024-02-02T23:05:06Z', utc: '2024-02-03T04:05:06Z' }, 'UTC-5 in February: the day before');
  assert.deepEqual(await stamp('UTC'), { date: '2024-02-03T04:05:06Z', utc: '2024-02-03T04:05:06Z' });
  assert.deepEqual(await stamp(undefined), { date: `${clockIn(when, runtimeZone())}Z`, utc: '2024-02-03T04:05:06Z' }, "the runtime's zone");
  // daylight saving: Sydney is UTC+11 in February, UTC+10 in July
  assert.equal((await stamp('Australia/Sydney')).date, '2024-02-03T15:05:06Z');
});

test("CR-018: a w:date without dateUtc is read as the package zone's wall clock; dateUtc is UTC whatever the zone", async () => {
  const read = async (zone, markup) => {
    const pkg = await WordprocessingMLPackage.createPackage();
    pkg.dateTimeZone = zone;
    const body = await pkg.getBody();
    await body.insertXml(`<w:p ${W} ${W16DU}>${markup}</w:p>`, 'End');
    const changes = await pkg.getAllTrackedChanges();
    assert.equal(changes.length, 1);
    return changes[0].date.toISOString();
  };
  const local = '<w:ins w:id="1" w:author="A" w:date="2024-02-03T04:05:06Z"><w:r><w:t>x</w:t></w:r></w:ins>';
  assert.equal(await read('Asia/Tokyo', local), '2024-02-02T19:05:06.000Z', "w:date is Word's local clock, here Tokyo's: 04:05 there is 19:05Z the day before");
  assert.equal(await read('UTC', local), '2024-02-03T04:05:06.000Z');
  assert.equal(await read('America/New_York', local), '2024-02-03T09:05:06.000Z');
  const withUtc = '<w:ins w:id="1" w:author="A" w:date="2024-02-03T04:05:06Z" w16du:dateUtc="2024-02-02T19:05:06Z"><w:r><w:t>x</w:t></w:r></w:ins>';
  assert.equal(await read('Asia/Tokyo', withUtc), '2024-02-02T19:05:06.000Z', 'dateUtc, where present, is UTC');
  assert.equal(await read('America/New_York', withUtc), '2024-02-02T19:05:06.000Z');
  // a w:date with a real offset, which Word never writes, is read as it says
  const offset = '<w:ins w:id="1" w:author="A" w:date="2024-02-03T04:05:06+09:00"><w:r><w:t>x</w:t></w:r></w:ins>';
  assert.equal(await read('America/New_York', offset), '2024-02-02T19:05:06.000Z');
});

test("CR-018: a new comment is stamped to the minute, its w:date in the package's zone; creationDate reads it back in that zone", async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.dateTimeZone = 'Asia/Tokyo';
  const body = await pkg.getBody();
  const before = Date.now();
  const comment = await body.insertParagraph('a comment here', 'End').insertComment('Needs work');
  const made = comment.creationDate;
  assert.equal(made.getUTCSeconds(), 0, 'to the minute');
  assert.equal(made.getUTCMilliseconds(), 0);
  assert.ok(made.getTime() <= before && before - made.getTime() < 61_000, 'now, truncated');
  const given = await body.insertParagraph('another', 'End').insertComment('Given a date', { date: new Date('2024-02-03T04:05:06Z') });
  assert.equal(given.creationDate.toISOString(), '2024-02-03T04:05:06.000Z', "a caller's date as given, to the second");
  const xml = await partOf(pkg, 'word/comments.xml');
  assert.match(xml, /w:date="2024-02-03T13:05:06Z"/, "w:date is Tokyo's wall clock");
  // read back with no dateUtc, in another zone, the same w:date is another instant
  const reread = await WordprocessingMLPackage.load(await pkg.save());
  reread.dateTimeZone = 'UTC';
  const back = (await (await reread.getBody()).getComments()).find((c) => c.content === 'Given a date');
  assert.ok(back);
  const utcPart = await partOf(reread, 'word/commentsExtensible.xml');
  if (utcPart === undefined || !utcPart.includes('dateUtc')) {
    assert.equal(back.creationDate.toISOString(), '2024-02-03T13:05:06.000Z', 'w:date read as a UTC wall clock');
  } else {
    assert.equal(back.creationDate.toISOString(), '2024-02-03T04:05:06.000Z', 'w16cex:dateUtc wins');
  }
});

test("CR-018: boundDatesConvertedAsWord shows a bound date as Word displays it in the package's zone; off, the nominal day (section 56)", async () => {
  const bind = async (values, configure) => {
    const pkg = await WordprocessingMLPackage.createPackage();
    configure(pkg);
    const body = await pkg.getBody();
    const controls = [];
    for (const value of values) {
      const control = body.insertParagraph('date', 'End').insertContentControl('DatePicker');
      control.datePickerContentControl.dateFormat = 'd MMMM yyyy HH:mm';
      const data = pkg.customXmlParts.add(`<root><d>${value}</d></root>`);
      await pkg.customXmlParts.load();
      assert.ok(control.xmlMapping.setMapping('/root[1]/d[1]', '', data));
      controls.push(control);
    }
    await pkg.customXmlParts.applyBindings();
    const xml = await partOf(pkg, 'word/document.xml');
    return { shown: controls.map((c) => c.text), fullDates: [...xml.matchAll(/w:fullDate="([^"]*)"/g)].map((m) => m[1]) };
  };
  const values = ['2026-10-22T15:05:00Z', '2026-10-22T00:00:00Z', '2026-10-22T00:00:00+05:00', '2026-10-22T00:00:00', '2026-10-22'];
  // off (the default): nominal, in every zone (CR-002 section 56)
  assert.deepEqual(await bind(values, (pkg) => { pkg.dateTimeZone = 'Australia/Sydney'; }), {
    shown: ['22 October 2026 15:05', '22 October 2026 00:00', '22 October 2026 00:00', '22 October 2026 00:00', '22 October 2026 00:00'],
    fullDates: ['2026-10-22T15:05:00Z', '2026-10-22T00:00:00Z', '2026-10-22T00:00:00Z', '2026-10-22T00:00:00Z', '2026-10-22T00:00:00Z'],
  }, 'a zone alone converts nothing');
  // on: a value with a Z or an offset is the instant, on the zone's wall clock (Sydney is UTC+11 in October:
  // 2026-10-22T15:05:00Z shows 2:05 AM the next day, docx4j CR-058 D6's probe); no marker, or a date alone, stays nominal
  assert.deepEqual(await bind(values, (pkg) => { pkg.dateTimeZone = 'Australia/Sydney'; pkg.boundDatesConvertedAsWord = true; }), {
    shown: ['23 October 2026 02:05', '22 October 2026 11:00', '22 October 2026 06:00', '22 October 2026 00:00', '22 October 2026 00:00'],
    fullDates: ['2026-10-23T02:05:00Z', '2026-10-22T11:00:00Z', '2026-10-22T06:00:00Z', '2026-10-22T00:00:00Z', '2026-10-22T00:00:00Z'],
  });
  assert.deepEqual((await bind(values.slice(0, 1), (pkg) => { pkg.dateTimeZone = 'America/New_York'; pkg.boundDatesConvertedAsWord = true; })).shown,
    ['22 October 2026 11:05'], 'west of UTC');
  assert.deepEqual((await bind(['2026-10-22T00:00:00Z'], (pkg) => { pkg.dateTimeZone = 'America/New_York'; pkg.boundDatesConvertedAsWord = true; })).shown,
    ['21 October 2026 20:00'], "Word's wart: the picker's midnight Z is the day before west of UTC");
});

test('CR-018: writeTimesOnSave: off, a save leaves dcterms alone; on, modified is now and created only for a package made here without one', async () => {
  const created = await WordprocessingMLPackage.createPackage();
  // createPackage writes an empty core properties part, no dates in it
  assert.equal(new TextDecoder().decode(new ZipPartStore(await created.save()).loadSync('docProps/core.xml')).includes('dcterms'), false, 'no dates by default');
  created.writeTimesOnSave = true;
  const before = Math.floor(Date.now() / 1000) * 1000;
  const xml = new TextDecoder().decode(new ZipPartStore(await created.save()).loadSync('docProps/core.xml'));
  const dates = [...xml.matchAll(/<dcterms:(created|modified)[^>]*>([^<]*)</g)].map((m) => [m[1], m[2]]);
  assert.deepEqual(dates.map((d) => d[0]), ['created', 'modified']);
  for (const [, value] of dates) {
    assert.match(value, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/, 'W3CDTF, UTC to the second');
    const t = new Date(value).getTime();
    assert.ok(t >= before && t <= Date.now(), 'now');
  }
  assert.match(xml, /xsi:type="dcterms:W3CDTF"/);
  assert.equal(created.lastSave.normalised.length, 0);

  // a loaded document: modified moves, created stays
  const bytes = await fixture('loadAndSave.docx');
  const loaded = await WordprocessingMLPackage.load(bytes);
  const original = new TextDecoder().decode(new ZipPartStore(bytes).loadSync('docProps/core.xml'));
  const createdBefore = /<dcterms:created[^>]*>([^<]*)</.exec(original)[1];
  const untouched = new TextDecoder().decode(new ZipPartStore(await loaded.save()).loadSync('docProps/core.xml'));
  assert.equal(untouched, original, 'off: byte for byte');
  loaded.writeTimesOnSave = true;
  const written = new TextDecoder().decode(new ZipPartStore(await loaded.save()).loadSync('docProps/core.xml'));
  assert.equal(/<dcterms:created[^>]*>([^<]*)</.exec(written)[1], createdBefore, 'created kept');
  assert.notEqual(/<dcterms:modified[^>]*>([^<]*)</.exec(written)[1], /<dcterms:modified[^>]*>([^<]*)</.exec(original)[1], 'modified moved');

  // a loaded package without the part is left without one, as docx4j leaves it
  const entries = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  delete entries['docProps/core.xml'];
  entries['_rels/.rels'] = strToU8(strFromU8(entries['_rels/.rels']).replace(/<Relationship [^>]*core-properties[^>]*\/>/, ''));
  entries['[Content_Types].xml'] = strToU8(strFromU8(entries['[Content_Types].xml']).replace(/<Override [^>]*core\.xml[^>]*\/>/, ''));
  const bare = await WordprocessingMLPackage.load(zipSync(entries));
  assert.equal(bare.docPropsCorePart === undefined, true, 'the fixture has no core properties part');
  bare.writeTimesOnSave = true;
  assert.equal(new ZipPartStore(await bare.save()).has('docProps/core.xml'), false, 'no part made for a loaded package');
});

test('CR-018: clone() carries the zone and the two opt-ins', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.dateTimeZone = 'Asia/Tokyo';
  pkg.boundDatesConvertedAsWord = true;
  pkg.writeTimesOnSave = true;
  const copy = await pkg.clone();
  assert.equal(copy.dateTimeZone, 'Asia/Tokyo');
  assert.equal(copy.boundDatesConvertedAsWord, true);
  assert.equal(copy.writeTimesOnSave, true);
});

// D6, measured by docx4j CR-067 section 12.2 (Jason, Word 365, Sydney, 2026-10-28): in a core property a date alone and a time
// with no offset are UTC, as a Z value is (2024-02-03 showed 11:00 at UTC+11; 2025-06-07T08:09:10 showed 18:09 at UTC+10).
test("CR-018 D6: the shim reads a date-only or offset-free dcterms value as UTC, as Word does; a Z or an offset as it says", async () => {
  const W3CDTF = (text) => ({ TYPE_NAME: 'org_docx4j_docProps_core_dc_terms.W3CDTF', content: [text] });
  const read = async (created, modified, lastPrinted) => {
    const pkg = await WordprocessingMLPackage.createPackage();
    const props = pkg.addDocPropsCorePart().contents;
    props.created = W3CDTF(created);
    props.modified = W3CDTF(modified);
    if (lastPrinted) props.lastPrinted = lastPrinted;
    let out;
    await Word.run(pkg, async (context) => {
      const p = context.document.properties;
      out = [p.creationDate?.toISOString(), p.lastSaveTime?.toISOString(), p.lastPrintDate?.toISOString()];
    });
    return out;
  };
  assert.deepEqual(await read('2024-02-03', '2025-06-07T08:09:10', { year: 2025, month: 1, day: 2, hour: 3, minute: 4, second: 5 }),
    ['2024-02-03T00:00:00.000Z', '2025-06-07T08:09:10.000Z', '2025-01-02T03:04:05.000Z'], 'UTC, whatever the runtime zone');
  assert.deepEqual(await read('2024-02-03T04:05:06Z', '2025-06-07T08:09:10+10:00', { year: 2025, month: 1, day: 2, hour: 3, minute: 4, second: 5, timezone: 600 }),
    ['2024-02-03T04:05:06.000Z', '2025-06-06T22:09:10.000Z', '2025-01-01T17:04:05.000Z'], 'a Z or an offset as it says');
  assert.deepEqual(await read('2024-02-03T04:05:06.250Z', '2025-06-07T08:09:10.5', undefined),
    ['2024-02-03T04:05:06.250Z', '2025-06-07T08:09:10.500Z', undefined], 'fractions kept');
  assert.deepEqual(await read('not a date', '', undefined), [undefined, undefined, undefined]);
});
