// CR-002 phase F: change tracking and replaceText.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { WordprocessingMLPackage, ZipPartStore, ChangeTracker, highestAnnotationId } from '../dist/index.mjs';
import { fixture, bytesEqual } from './helpers.mjs';

const DATE = new Date(Date.UTC(2026, 8, 16, 10, 30, 0));
/** A date as Word writes `w:date`: local wall-clock time, with a `Z` (CR-002 section 29, check 23). */
const wordDate = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
  + `T${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}:${String(d.getSeconds()).padStart(2, '0')}Z`;
/**
 * A pattern with `@DATE` standing for DATE as `w:date` carries it, whatever this process's time zone,
 * and `@UTC` for the `w16du:dateUtc` beside it (CR-002 section 29).
 */
const at = (pattern) => new RegExp(pattern.source.replaceAll('@DATE', wordDate(DATE)).replaceAll('@UTC', ' w16du:dateUtc="2026-09-16T10:30:00Z"'));

/** A package of the given paragraphs, tracking on, author Ada, with a fixed revision date. */
async function tracked(texts = ['The quick brown fox jumps', 'Second paragraph here', 'Third one']) {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.author = { name: 'Ada' };
  pkg.trackedChangeDate = DATE;
  for (const text of texts) pkg.body.insertParagraph(text, 'End');
  pkg.changeTrackingMode = 'TrackAll';
  return pkg;
}

/** The same package with tracking off, for comparing the accepted text with an untracked edit. */
async function untracked(texts) {
  const pkg = await WordprocessingMLPackage.createPackage();
  for (const text of texts) pkg.body.insertParagraph(text, 'End');
  return pkg;
}

const xmlOf = (pkg) => pkg.getMainDocumentPart().getXml();

test('changeTrackingMode: off by default, written to w:trackRevisions, read back after a round trip', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  assert.equal(pkg.changeTrackingMode, 'Off');
  assert.equal(pkg.changeTracker, undefined);
  pkg.changeTrackingMode = 'TrackAll';
  assert.equal(pkg.changeTrackingMode, 'TrackAll');
  assert.ok(pkg.changeTracker, 'a tracker while the mode is on');
  const settings = pkg.getMainDocumentPart().documentSettingsPart;
  assert.deepEqual(settings.contents.trackRevisions, {});
  const back = await WordprocessingMLPackage.load(await pkg.save());
  assert.equal(back.changeTrackingMode, 'Off', 'sync read before the settings part is unmarshalled');
  assert.equal(await back.getChangeTrackingMode(), 'TrackAll');
  assert.equal(back.changeTrackingMode, 'TrackAll', 'cached after the async read');
  await back.setChangeTrackingMode('Off');
  assert.equal(back.getMainDocumentPart().documentSettingsPart.contents.trackRevisions, undefined);
});

test('changeTrackingMode set on a loaded package is written through on save', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('tracked-changes.docx'));
  assert.equal(pkg.getMainDocumentPart().documentSettingsPart.isUnmarshalled, false);
  pkg.changeTrackingMode = 'TrackAll';
  const back = await WordprocessingMLPackage.load(await pkg.save());
  assert.equal(await back.getChangeTrackingMode(), 'TrackAll');
});

test('insertText writes w:ins; the accepted text is what an untracked edit gives', async () => {
  const pkg = await tracked();
  const p = pkg.body.paragraphs[0];
  p.insertText('Hello ', 'Start');
  assert.equal(p.text, 'Hello The quick brown fox jumps');
  assert.equal(p.getText({ view: 'original' }), 'The quick brown fox jumps');
  const xml = await xmlOf(pkg);
  assert.match(xml, at(/<w:ins w:id="1"@UTC w:author="Ada" w:date="@DATE"><w:r><w:t xml:space="preserve">Hello <\/w:t><\/w:r><\/w:ins>/));
  // the tree: one w:ins holding one run
  const ins = p.p.content[0];
  assert.equal(ins.name.localPart, 'ins');
  assert.equal(ins.value.TYPE_NAME, 'org_docx4j_wml.RunIns');
  assert.equal(ins.value.customXmlOrSmartTagOrSdt.length, 1, 'runs live under customXmlOrSmartTagOrSdt');
  assert.equal(ins.value.customXmlOrSmartTagOrSdt[0].value.PARENT, ins.value, 'PARENT linked');

  const plainPkg = await untracked(['The quick brown fox jumps']);
  plainPkg.body.paragraphs[0].insertText('Hello ', 'Start');
  assert.equal(p.text, plainPkg.body.paragraphs[0].text);
});

test('a second insertion by the same author extends the w:ins rather than nesting one', async () => {
  const pkg = await tracked(['abc']);
  const p = pkg.body.paragraphs[0];
  p.insertText('X', 'Start');
  p.insertText('Y', 'Start');
  assert.equal(p.text, 'YXabc');
  const inserts = p.p.content.filter((el) => el.name.localPart === 'ins');
  assert.equal(inserts.length, 1, 'one w:ins');
  assert.equal((await xmlOf(pkg)).includes('<w:ins'), true);
  assert.match(await xmlOf(pkg), /<w:ins w:id="1"[^>]*><w:r><w:t>YX<\/w:t><\/w:r><\/w:ins>/);
});

test('a deletion moves the runs into w:del with w:delText; the original view keeps the text', async () => {
  const pkg = await tracked(['The quick brown fox jumps']);
  const p = pkg.body.paragraphs[0];
  p.search('brown ')[0].delete();
  assert.equal(p.text, 'The quick fox jumps');
  assert.equal(p.getText({ view: 'original' }), 'The quick brown fox jumps');
  const xml = await xmlOf(pkg);
  assert.match(xml, at(/<w:del w:id="1"@UTC w:author="Ada" w:date="@DATE"><w:r><w:delText xml:space="preserve">brown <\/w:delText><\/w:r><\/w:del>/));
  assert.equal(xml.includes('<w:t xml:space="preserve">brown'), false);

  const plainPkg = await untracked(['The quick brown fox jumps']);
  plainPkg.body.paragraphs[0].search('brown ')[0].delete();
  assert.equal(p.text, plainPkg.body.paragraphs[0].text);
});

test('a replacement is the w:del first and the w:ins after it', async () => {
  const pkg = await tracked(['The quick brown fox']);
  const p = pkg.body.paragraphs[0];
  p.search('brown')[0].insertText('red', 'Replace');
  assert.equal(p.text, 'The quick red fox');
  assert.equal(p.getText({ view: 'original' }), 'The quick brown fox');
  const xml = await xmlOf(pkg);
  assert.match(xml, /<w:del w:id="1"[^>]*><w:r><w:delText>brown<\/w:delText><\/w:r><\/w:del><w:ins w:id="2"[^>]*><w:r><w:t>red<\/w:t><\/w:r><\/w:ins>/);
});

test('deleting text this author inserted takes it back; editing another author\'s w:del is refused', async () => {
  const pkg = await tracked(['abc']);
  const p = pkg.body.paragraphs[0];
  p.insertText('XY', 'Start');
  assert.equal(p.text, 'XYabc');
  p.search('XY')[0].delete();
  assert.equal(p.text, 'abc');
  assert.equal(p.p.content.some((el) => el.name.localPart === 'ins'), false, 'the emptied w:ins went too');

  // a segment inside a w:del never reaches the accepted view, so the refusal is on the primitive
  const tracker = pkg.changeTracker;
  assert.throws(() => tracker.assertEditable({ kind: 'del', element: {}, value: { author: 'Bob', id: 9 }, items: [], owner: [] }),
    /Cannot edit text inside a w:del \(deleted by Bob\)/);
});

test('Font and paragraph properties keep the old ones in w:rPrChange and w:pPrChange', async () => {
  const pkg = await tracked(['Second paragraph here']);
  const p = pkg.body.paragraphs[0];
  p.search('Second')[0].font.bold = true;
  p.style = 'Heading1';
  p.alignment = 'Centered';
  const xml = await xmlOf(pkg);
  assert.match(xml, at(/<w:rPr><w:b\/><w:bCs\/><w:rPrChange w:id="1"@UTC w:author="Ada" w:date="@DATE"><w:rPr\/><\/w:rPrChange><\/w:rPr>/));
  assert.match(xml, at(/<w:pPrChange w:id="2"@UTC w:author="Ada" w:date="@DATE"><w:pPr\/><\/w:pPrChange>/));
  assert.equal(xml.includes('xsi:type'), false, 'w:pPrChange/w:pPr is a CT_PPrBase, no xsi:type');
  // one w:pPrChange only: the first write keeps the original
  assert.equal(xml.match(/<w:pPrChange/g).length, 1);

  const kinds = pkg.body.getTrackedChanges().map((c) => [c.type, c.target.kind]);
  assert.deepEqual(kinds, [['Formatted', 'paragraphProperties'], ['Formatted', 'runProperties']]);
});

test('the original in w:pPrChange is a CT_PPrBase: no w:rPr, no w:sectPr, no xsi:type', async () => {
  // objects 0.1.4's deepCopyAsSync types the copy as the base and drops what it does not declare
  // (CR-002 section 15); before it, the three properties were deleted by hand.
  const pkg = await tracked(['A paragraph']);
  const inserted = pkg.body.insertParagraph('New paragraph', 'End');
  assert.ok(inserted.p.pPr.rPr.ins, 'the mark of a tracked new paragraph is an insertion');
  inserted.p.pPr.pStyle = { val: 'Heading1' };   // on the tree: the next tracked write records it
  inserted.alignment = 'Centered';
  const xml = await xmlOf(pkg);
  const recorded = xml.match(/<w:pPrChange [^>]*>.*?<\/w:pPrChange>/g);
  assert.equal(recorded.length, 1);
  assert.match(recorded[0], /<w:pStyle w:val="Heading1"\/>/, 'the style it had is in the original');
  assert.equal(recorded[0].includes('<w:rPr'), false, 'CT_PPrBase declares no w:rPr');
  assert.equal(recorded[0].includes('<w:sectPr'), false, 'nor a w:sectPr');
  assert.equal(xml.includes('xsi:type'), false, 'no xsi:type on w:pPrChange/w:pPr');
});

test('a formatting change records the properties that were there, and reject puts them back', async () => {
  const pkg = await tracked(['abc']);
  const p = pkg.body.paragraphs[0];
  p.font.italic = true;                    // untracked baseline written through the tracker
  const run = p.runs[0].value;
  assert.ok(run.rPr.rPrChange, 'w:rPrChange recorded');
  assert.deepEqual(run.rPr.rPrChange.rPr.egrPrBase, [], 'nothing was set before');
  p.font.bold = true;
  assert.equal(p.p.content[0].value.rPr.rPrChange.id, 1, 'still the first original');
  const change = p.getTrackedChanges().find((c) => c.target.kind === 'runProperties');
  change.reject();
  assert.equal(p.runs[0].value.rPr, undefined, 'back to no direct formatting');
});

test('insertParagraph marks the new mark inserted; delete() marks the mark deleted', async () => {
  const pkg = await tracked(['one', 'two']);
  const body = pkg.body;
  const fresh = body.insertParagraph('three', 'End');
  let xml = await xmlOf(pkg);
  assert.match(xml, /<w:p><w:pPr><w:rPr><w:ins w:id="1"[^>]*\/><\/w:rPr><\/w:pPr><w:ins w:id="2"[^>]*><w:r><w:t>three<\/w:t><\/w:r><\/w:ins><\/w:p>/);
  assert.equal(fresh.text, 'three');

  body.paragraphs[1].delete();
  xml = await xmlOf(pkg);
  assert.match(xml, /<w:pPr><w:rPr><w:del w:id="4"[^>]*\/><\/w:rPr><\/w:pPr><w:del w:id="3"[^>]*><w:r><w:delText>two<\/w:delText><\/w:r><\/w:del>/);
  assert.equal(body.paragraphs.length, 3, 'the paragraph is still there while the change is pending');
  assert.equal(body.getText({ view: 'original' }), 'one\ntwo\n');

  // a paragraph this author inserted is taken back outright
  fresh.delete();
  assert.deepEqual(body.paragraphs.map((q) => q.text), ['one', '']);
  assert.throws(() => body.paragraphs[1].delete(), /already marked deleted/);
});

test('table rows: w:trPr/w:ins and w:trPr/w:del, and an inserted table marks every row', async () => {
  const pkg = await tracked([]);
  const body = pkg.body;
  pkg.changeTrackingMode = 'Off';
  await body.insertXml('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>r1</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>r2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>', 'End');
  pkg.changeTrackingMode = 'TrackAll';
  const rows = body.tables[0].element.value.content.filter((el) => el.name.localPart === 'tr');
  pkg.changeTracker.markRowDeleted(rows[0].value);
  pkg.changeTracker.markRowInserted(rows[1].value);
  const xml = await xmlOf(pkg);
  assert.match(xml, /<w:tr><w:trPr><w:del w:id="1"[^>]*\/><\/w:trPr>/);
  assert.match(xml, /<w:tr><w:trPr><w:ins w:id="2"[^>]*\/><\/w:trPr>/);
  assert.deepEqual(body.getTrackedChanges().map((c) => [c.type, c.target.kind, c.text]), [['Deleted', 'row', 'r1'], ['Added', 'row', 'r2']]);
  assert.equal(body.acceptAll(), 2);
  assert.equal(body.text, 'r2');

  // a table inserted while tracking marks each row and each paragraph
  const [added] = await body.insertXml('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>new</w:t></w:r></w:p></w:tc></w:tr></w:tbl>', 'End');
  const tr = added.element.value.content[0].value;
  assert.ok(tr.trPr.ins, 'the row is an insertion');
  assert.ok(tr.content[0].value.content[0].value.pPr.rPr.ins, 'and so is the paragraph mark');
});

test('revision ids start above the highest annotation id in the parts already unmarshalled', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('tracked-changes.docx'));
  pkg.author = { name: 'Ada' };
  pkg.trackedChangeDate = DATE;
  const body = await pkg.getBody();
  await pkg.setChangeTrackingMode('TrackAll');
  body.paragraphs[0].insertText('Q', 'Start');
  const change = body.getTrackedChanges().find((c) => c.author === 'Ada');
  assert.ok(change.id > 2, `above the fixture's w:ins 1 and w:del 2, got ${change.id}`);
});

test('getTrackedChanges over a Word fixture; accept and reject of each kind', async () => {
  const load = async () => {
    const pkg = await WordprocessingMLPackage.load(await fixture('tracked-changes.docx'));
    return { pkg, body: await pkg.getBody() };
  };
  const { body } = await load();
  const changes = body.getTrackedChanges();
  assert.deepEqual(changes.map((c) => [c.type, c.target.kind, c.author, c.text]), [
    ['Added', 'run', 'Jason Harrop', 'An insertion'],
    ['Deleted', 'run', 'Jason Harrop', ' A deletion'],
  ]);
  // w:date="2007-12-09T10:14:00Z" and no w16du:dateUtc: Word's local wall-clock time, as Office JS reads it
  assert.equal(changes[0].date.toISOString(), new Date(2007, 11, 9, 10, 14).toISOString());
  assert.equal(changes[0].id, 1);
  const p = changes[0].paragraph;
  assert.equal(p.text, 'Here is some change tracking. An insertion Followed by.');
  assert.equal(p.getText({ view: 'original' }), 'Here is some change tracking.  Followed by A deletion.');
  assert.deepEqual([changes[0].getRange().start, changes[0].getRange().end], [30, 42]);
  assert.deepEqual([changes[1].getRange().start, changes[1].getRange().end], [54, 54], 'a deletion is collapsed where it sits');
  assert.equal(p.getTrackedChanges().length, 2);
  assert.equal(p.getRange('Whole').getTrackedChanges().length, 2);
  assert.equal(new (Object.getPrototypeOf(p.getRange()).constructor)(p, 0, 5).getTrackedChanges().length, 0, 'a span before both');

  // accept all: the insertion is unwrapped, the deletion removed, nothing left
  const accepted = await load();
  assert.equal(accepted.body.acceptAll(), 2);
  assert.equal(accepted.body.getTrackedChanges().length, 0);
  const acceptedP = accepted.body.paragraphs.find((q) => q.text.startsWith('Here is some'));
  assert.equal(acceptedP.text, 'Here is some change tracking. An insertion Followed by.');
  const acceptedXml = await xmlOf(accepted.pkg);
  assert.equal(acceptedXml.includes('<w:ins '), false);
  assert.equal(acceptedXml.includes('<w:del '), false);
  assert.equal(acceptedXml.includes('<w:delText'), false);

  // reject all: the insertion goes, the deletion comes back as w:t
  const rejected = await load();
  assert.equal(rejected.body.rejectAll(), 2);
  assert.equal(rejected.body.getTrackedChanges().length, 0);
  const rejectedP = rejected.body.paragraphs.find((q) => q.text.startsWith('Here is some'));
  assert.equal(rejectedP.text, 'Here is some change tracking.  Followed by A deletion.');
  const rejectedXml = await xmlOf(rejected.pkg);
  assert.equal(rejectedXml.includes('<w:delText'), false);
  assert.match(rejectedXml, /<w:t xml:space="preserve"> A deletion<\/w:t>/);
});

test('accept and reject of a paragraph mark join the paragraphs, as docx4j AcceptTrackedChanges does', async () => {
  // a paragraph's text and mark are one change, as Office JS lists them (CR-002 section 29, check 22)
  const deleted = await tracked(['one', 'two']);
  deleted.body.paragraphs[0].delete();
  assert.equal(deleted.body.acceptAll(), 1);
  assert.deepEqual(deleted.body.paragraphs.map((q) => q.text), ['two'], 'the deleted mark joined the two');

  const inserted = await tracked(['one']);
  inserted.body.insertParagraph('two', 'End');
  assert.equal(inserted.body.rejectAll(), 1);
  assert.deepEqual(inserted.body.paragraphs.map((q) => q.text), ['one'], 'the inserted paragraph went');

  const kept = await tracked(['one']);
  kept.body.insertParagraph('two', 'End');
  assert.equal(kept.body.acceptAll(), 1);
  assert.deepEqual(kept.body.paragraphs.map((q) => q.text), ['one', 'two']);
  const xml = await xmlOf(kept);
  assert.equal(xml.includes('<w:ins'), false);
});

test('replaceText: counts, offsets with several matches per paragraph, tracked and not', async () => {
  const plainPkg = await untracked(['a b a b a', 'none here', 'and a here']);
  assert.equal(plainPkg.body.replaceText('a', 'X'), 5);
  assert.deepEqual(plainPkg.body.paragraphs.map((q) => q.text), ['X b X b X', 'none here', 'Xnd X here']);
  assert.equal(plainPkg.body.paragraphs[0].replaceText('X', 'zz'), 3);
  assert.equal(plainPkg.body.paragraphs[0].text, 'zz b zz b zz');

  const pkg = await tracked(['a b a b a']);
  const p = pkg.body.paragraphs[0];
  assert.equal(p.replaceText('a', 'XY'), 3);
  assert.equal(p.text, 'XY b XY b XY');
  assert.equal(p.getText({ view: 'original' }), 'a b a b a');
  const xml = await xmlOf(pkg);
  assert.equal(xml.match(/<w:del /g).length, 3);
  assert.equal(xml.match(/<w:ins /g).length, 3);
  assert.equal(pkg.body.acceptAll(), 6);
  assert.equal(p.text, 'XY b XY b XY');

  // a range restricts the replacement to its span
  const inRange = await tracked(['a b a b a']);
  const q = inRange.body.paragraphs[0];
  const range = q.search('a b a')[0];
  assert.equal(range.replaceText('a', 'Z'), 2);
  assert.equal(q.text, 'Z b Z b a');
});

test('Table: addRows, insertRows, deleteRows, row delete and table delete are tracked', async () => {
  const pkg = await tracked([]);
  const body = pkg.body;
  pkg.changeTrackingMode = 'Off';
  const table = body.insertTable(2, 2, 'End', [['a', 'b'], ['c', 'd']]);
  pkg.changeTrackingMode = 'TrackAll';

  const [added] = table.addRows('End', 1, [['e', 'f']]);
  assert.ok(added.tr.trPr.ins, 'w:trPr/w:ins on the added row');
  assert.ok(added.cells[0].body.paragraphs[0].p.pPr.rPr.ins, 'and its paragraph marks are insertions');
  assert.match(await xmlOf(pkg), at(/<w:trPr><w:ins w:id="1"@UTC w:author="Ada" w:date="@DATE"\/><\/w:trPr>/));

  const [between] = table.rows[0].insertRows('After', 1, [['x', 'y']]);
  assert.ok(between.tr.trPr.ins, 'insertRows marks the row inserted too');

  table.rows[0].delete();
  assert.ok(table.rows[0].tr.trPr.del, 'a deleted row stays and takes w:trPr/w:del');
  assert.equal(table.rowCount, 4, 'nothing is removed while the change is pending');

  table.deleteRows(2, 1);
  assert.ok(table.rows[2].tr.trPr.del);

  const kinds = body.getTrackedChanges().filter((c) => c.target.kind === 'row').map((c) => c.type);
  assert.deepEqual(kinds, ['Deleted', 'Added', 'Deleted', 'Added']);

  // accepting: the deleted rows go, the added ones stay
  assert.ok(body.acceptAll() > 0);
  assert.deepEqual(table.values, [['x', 'y'], ['e', 'f']]);
  assert.equal((await xmlOf(pkg)).includes('<w:trPr>'), false, 'no revision left on any row');

  // a whole table deleted is every row marked deleted
  const second = await tracked([]);
  second.changeTrackingMode = 'Off';
  const t2 = second.body.insertTable(2, 1, 'End', [['p'], ['q']]);
  second.changeTrackingMode = 'TrackAll';
  t2.delete();
  assert.equal(second.body.tables.length, 1, 'the table is still there');
  assert.deepEqual(t2.rows.map((r) => r.tr.trPr.del !== undefined), [true, true]);
  assert.equal(second.body.acceptAll(), 2);
  assert.equal(second.body.tables[0].rowCount, 0);

  // and a table inserted while tracking marks every row and paragraph
  const third = await tracked([]);
  const t3 = third.body.insertTable(1, 1, 'End', [['z']]);
  assert.ok(t3.rows[0].tr.trPr.ins);
  assert.ok(t3.rows[0].cells[0].body.paragraphs[0].p.pPr.rPr.ins);
});

test('a deleted row marks its cells: w:delText, w:del on every paragraph mark, one change, accept and reject', async () => {
  const pkg = await tracked([]);
  const body = pkg.body;
  pkg.changeTrackingMode = 'Off';
  const table = body.insertTable(2, 2, 'End', [['a', 'b'], ['c', 'd']]);
  pkg.changeTrackingMode = 'TrackAll';

  table.deleteRows(0, 1);
  const row = table.rows[0];
  assert.ok(row.tr.trPr.del, 'w:trPr/w:del on the row');
  for (const cell of row.cells) {
    assert.ok(cell.body.paragraphs[0].p.pPr.rPr.del, 'and w:pPr/w:rPr/w:del on every paragraph mark');
  }
  const xml = await xmlOf(pkg);
  assert.match(xml, /<w:delText>a<\/w:delText>/, 'the cell text is w:delText, so Word strikes it out');
  assert.match(xml, /<w:delText>b<\/w:delText>/);
  assert.equal(xml.includes('<w:t>a</w:t>'), false, 'and no w:t is left in the deleted row');
  assert.match(xml, /<w:del w:id="\d+"[^>]*><w:r><w:delText>a<\/w:delText><\/w:r><\/w:del>/);

  // Office JS reports the row, not the markup inside it: one change (CR-002 section 13)
  const changes = body.getTrackedChanges();
  assert.deepEqual(changes.map((c) => [c.type, c.target.kind]), [['Deleted', 'row']]);
  assert.equal(changes[0].text, 'a\nb');

  // rejecting brings the row back, content readable
  body.rejectAll();
  assert.deepEqual(table.values, [['a', 'b'], ['c', 'd']]);
  assert.equal(body.getTrackedChanges().length, 0);
  assert.equal((await xmlOf(pkg)).includes('w:delText'), false, 'and no deleted text is left');

  // accepting takes the row away
  table.deleteRows(0, 1);
  assert.equal(body.acceptAll(), 1);
  assert.deepEqual(table.values, [['c', 'd']]);

  // the same for a whole table deleted, and for an inserted row deleted again
  const second = await tracked([]);
  second.changeTrackingMode = 'Off';
  const t2 = second.body.insertTable(1, 1, 'End', [['p']]);
  second.changeTrackingMode = 'TrackAll';
  t2.delete();
  assert.match(await second.getMainDocumentPart().getXml(), /<w:delText>p<\/w:delText>/);
  assert.equal(second.body.getTrackedChanges().length, 1, 'one change for the one row');
  second.body.rejectAll();
  assert.deepEqual(second.body.tables[0].values, [['p']]);

  const third = await tracked([]);
  const t3 = third.body.insertTable(1, 1, 'End', [['z']]);      // inserted while tracking
  assert.equal(third.body.getTrackedChanges().length, 1, 'an inserted row reports once too');
  t3.rows[0].delete();
  const rowChanges = third.body.getTrackedChanges();
  assert.deepEqual(rowChanges.map((c) => c.type), ['Added', 'Deleted']);
  assert.equal(t3.rows[0].cells[0].body.paragraphs.length, 1, 'the cell keeps its paragraph');
  assert.equal(third.body.acceptAll(), 2);
  assert.equal(third.body.tables[0].rowCount, 0);
});

test('a comment made while tracking is on is a comment, not an insertion', async () => {
  const pkg = await tracked(['The quick brown fox jumps']);
  const body = pkg.body;
  const p = body.paragraphs[0];
  const comment = await p.getRange().search('brown')[0].insertComment('Which brown?');
  assert.equal(comment.content, 'Which brown?');
  const xml = await xmlOf(pkg);
  // no revision anywhere in the document part: the markers and the reference run are not tracked
  assert.equal(xml.includes('<w:ins'), false, 'no w:ins around the markers or the reference run');
  assert.equal(xml.includes('<w:del'), false);
  assert.equal(p.text, 'The quick brown fox jumps', 'the text is untouched');
  assert.equal(body.getTrackedChanges().length, 0);
  assert.equal((await body.getComments()).length, 1);

  // and the markers stay outside a w:ins when the commented text is itself an insertion
  const inserted = await tracked(['before after']);
  const q = inserted.body.paragraphs[0];
  q.search('before')[0].insertText('NEW', 'Replace');
  await q.search('NEW')[0].insertComment('on the insertion');
  const insXml = await xmlOf(inserted);
  assert.match(insXml, /<w:commentRangeStart w:id="0"\/><w:ins /, 'the range start is before the w:ins');
  assert.match(insXml, /<\/w:ins><w:commentRangeEnd w:id="0"\/><w:r>/, 'the end and the reference run follow it, outside');
  assert.equal(insXml.includes('<w:ins w:id="2"'), true, 'the insertion itself is still tracked');
  assert.equal(inserted.body.getTrackedChanges().filter((c) => c.target.kind === 'run').length, 2, 'still just the del and the ins');
  assert.equal((await inserted.body.getComments()).length, 1);
});

test('a content control inherits tracking from the paragraph primitives', async () => {
  const pkg = await tracked([]);
  const body = pkg.body;
  pkg.changeTrackingMode = 'Off';
  await body.insertXml('<w:sdt><w:sdtPr><w:tag w:val="t"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:sdtContent></w:sdt>', 'End');
  pkg.changeTrackingMode = 'TrackAll';
  const control = body.contentControls[0];
  control.insertText(' more', 'End');
  assert.equal(control.text, 'inside more');
  assert.match(await xmlOf(pkg), /<w:ins w:id="1"[^>]*><w:r><w:t xml:space="preserve"> more<\/w:t><\/w:r><\/w:ins>/);
  assert.deepEqual(body.getTrackedChanges().map((c) => c.type), ['Added']);
});

test('the Word shim: document.changeTrackingMode delegates and getTrackedChanges reports', async () => {
  const { Word } = await import('../dist/office-js/index.mjs');
  const pkg = await tracked(['The quick brown fox']);
  await Word.run(pkg, async (context) => {
    const doc = context.document;
    assert.equal(doc.changeTrackingMode, 'TrackAll');
    doc.changeTrackingMode = 'Off';
    assert.equal(pkg.changeTrackingMode, 'Off', 'the shim writes the package');
    doc.changeTrackingMode = 'TrackAll';
    doc.body.search('brown').items[0].insertText('red', 'Replace');
    await context.sync();
    const collection = doc.getTrackedChanges();
    assert.ok(Array.isArray(collection.items), 'the shim hands back a collection with items, as Office JS does');
    const changes = collection.items;
    assert.deepEqual(changes.map((c) => [c.type, c.author, c.text]), [['Deleted', 'Ada', 'brown'], ['Added', 'Ada', 'red']]);
  });
  assert.equal(pkg.body.text, 'The quick red fox');
});

test('an untouched tracked-changes fixture round trips byte for byte', async () => {
  const original = await fixture('tracked-changes.docx');
  const pkg = await WordprocessingMLPackage.load(original);
  const saved = await pkg.save();
  const before = new ZipPartStore(original);
  const after = new ZipPartStore(saved);
  assert.deepEqual(new Set(after.partNames()), new Set(before.partNames()));
  for (const name of after.partNames()) {
    if (name === '[Content_Types].xml' || name.endsWith('.rels')) continue;
    assert.ok(bytesEqual(before.loadSync(name), after.loadSync(name)), `${name} byte-identical`);
  }
});

// CR-002 section 28.2: the annotation id counter must be above every id in the DOCUMENT, not only
// in the parts that happen to be unmarshalled. A revision in an unread footnote, endnote, header or
// comment part did not raise it, so the next id collided with one already in use - ECMA-376
// 17.13.5.4 makes that one id space. `seedAnnotationIds()` reads the untouched parts privately
// (CR-001 Phase D's device), so they keep their byte-for-byte round trip.
test('a revision id is above every id in the document, not only in the parts read', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  // a document whose footnotes part carries a high revision id, and which nothing has unmarshalled
  const source = await WordprocessingMLPackage.createPackage();
  const main = source.getMainDocumentPart();
  const { FootnotesPart } = await import('../dist/index.mjs');
  const footnotes = new FootnotesPart();
  footnotes.setXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="${W}">`
    + '<w:footnote w:id="1"><w:p><w:ins w:id="9100" w:author="Bob" w:date="2026-01-01T00:00:00Z">'
    + '<w:r><w:t>a tracked footnote</w:t></w:r></w:ins></w:p></w:footnote></w:footnotes>');
  main.addTargetPart(footnotes);
  source.body.insertParagraph('Body text', 'End');

  const pkg = await WordprocessingMLPackage.load(await source.save());
  assert.equal(pkg.getMainDocumentPart().footnotesPart?.isUnmarshalled, false, 'the footnotes part is untouched');

  // the async setter seeds the floor, so the id cannot collide
  await pkg.getMainDocumentPart().getContents();
  pkg.author = { name: 'Ada' };
  await pkg.setChangeTrackingMode('TrackAll');
  assert.ok(pkg.annotationIdFloor >= 9100, `floor ${pkg.annotationIdFloor} is above the footnote's id`);
  const p = pkg.body.paragraphs[0];
  p.getRange('End').insertText(' and more', 'End');
  const ins = [...(await pkg.getMainDocumentPart().getContents()).body.content]
    .flatMap((e) => e.value?.content ?? []).map((e) => e.value).filter((v) => v?.TYPE_NAME === 'org_docx4j_wml.RunIns');
  assert.ok(ins.length >= 1, 'the edit was tracked');
  assert.ok(ins.every((v) => v.id > 9100), `ids ${ins.map((v) => v.id)} are above the footnote's 9100`);

  // and reading it did not cost the footnotes part its round trip
  assert.equal(pkg.getMainDocumentPart().footnotesPart.isUnmarshalled, false, 'still not unmarshalled');
  const saved = new ZipPartStore(await pkg.save());
  const original = new ZipPartStore(await source.save());
  assert.deepEqual(saved.loadSync('word/footnotes.xml'), original.loadSync('word/footnotes.xml'));
});

// CR-002 section 28.2: Word stamps a revision to the minute, and groups adjacent revisions by
// author and timestamp; a per-second stamp makes one change of every keystroke run.
test('a revision dated now is stamped to the minute, and an explicit date is untouched', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.author = { name: 'Ada' };
  pkg.body.insertParagraph('text', 'End');
  pkg.changeTrackingMode = 'TrackAll';
  pkg.body.paragraphs[0].getRange('End').insertText(' more', 'End');
  const changes = pkg.body.getTrackedChanges();
  assert.equal(changes.length, 1);
  assert.equal(changes[0].date.getUTCSeconds(), 0, 'seconds zero');
  assert.equal(changes[0].date.getUTCMilliseconds(), 0, 'milliseconds zero');
  const xml = await pkg.getMainDocumentPart().getXml();
  assert.match(xml, /w:date="\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:00Z"/, 'written with :00 seconds');

  // two edits a moment apart carry the same stamp, which is what lets Word group them
  pkg.body.paragraphs[0].getRange('End').insertText(' and more', 'End');
  const dates = pkg.body.getTrackedChanges().map((c) => c.date.toISOString());
  assert.equal(new Set(dates).size, 1, `one timestamp for both edits, got ${dates.join(', ')}`);

  // an explicit date is the caller's business and is not truncated
  const exact = new Date(Date.UTC(2026, 8, 16, 10, 30, 45, 250));
  const p2 = await WordprocessingMLPackage.createPackage();
  p2.author = { name: 'Ada' };
  p2.trackedChangeDate = exact;
  p2.body.insertParagraph('text', 'End');
  p2.changeTrackingMode = 'TrackAll';
  p2.body.paragraphs[0].getRange('End').insertText('!', 'End');
  assert.equal(p2.body.getTrackedChanges()[0].date.getUTCSeconds(), 45, 'the caller\'s seconds kept');
});

// CR-002 section 28.1: an editor's agent records its edits as itself, in a room whose
// w:trackRevisions belongs to the document. So withTracking must leave no trace of itself: not in
// the settings part, not in a settings part it created, not in trackingPending, and not in the
// cached mode a later read reports.
test('withTracking records revisions without touching the document\'s setting', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.author = { name: 'Ada' };
  pkg.body.insertParagraph('text', 'End');
  const settings = pkg.getMainDocumentPart().documentSettingsPart;
  const before = await settings.getXml();
  assert.equal(pkg.changeTrackingMode, 'Off');

  const date = new Date(Date.UTC(2026, 8, 27, 9, 15, 0));
  const returned = await pkg.withTracking({ author: { name: 'Agent', initials: 'AG' }, date }, () => {
    pkg.body.paragraphs[0].getRange('End').insertText(' from the agent', 'End');
    return 'result';
  });
  assert.equal(returned, 'result', 'the value of fn is returned');

  const changes = pkg.body.getTrackedChanges();
  assert.equal(changes.length, 1);
  assert.equal(changes[0].author, 'Agent');
  assert.equal(changes[0].date.toISOString(), date.toISOString());

  // and nothing of the call survives it
  assert.equal(pkg.changeTrackingMode, 'Off', 'the mode is as it was');
  assert.equal(pkg.author.name, 'Ada', 'the package author is restored');
  assert.equal(pkg.trackedChangeDate, undefined, 'the package date is restored');
  assert.equal(await settings.getXml(), before, 'the settings part is byte-identical');
  assert.ok(!(await settings.getXml()).includes('trackRevisions'), 'no w:trackRevisions was written');
  const reloaded = await WordprocessingMLPackage.load(await pkg.save());
  assert.equal(await reloaded.getChangeTrackingMode(), 'Off', 'and the saved document is not tracking');
});

test('withTracking restores after a throw, and nests over an on setting', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.author = { name: 'Ada' };
  pkg.body.insertParagraph('text', 'End');
  await assert.rejects(() => pkg.withTracking({ author: { name: 'Agent' } }, () => { throw new Error('boom'); }), /boom/);
  assert.equal(pkg.changeTrackingMode, 'Off', 'restored after a throw');
  assert.equal(pkg.author.name, 'Ada');

  // over a document that IS tracking: the call's author wins inside, the document's outside
  await pkg.setChangeTrackingMode('TrackAll');
  pkg.body.paragraphs[0].getRange('End').insertText(' by Ada', 'End');
  await pkg.withTracking({ author: { name: 'Agent' } }, () => {
    pkg.body.paragraphs[0].getRange('End').insertText(' by the agent', 'End');
  });
  pkg.body.paragraphs[0].getRange('End').insertText(' by Ada again', 'End');
  assert.deepEqual([...new Set(pkg.body.getTrackedChanges().map((c) => c.author))].sort(), ['Ada', 'Agent']);
  assert.equal(pkg.changeTrackingMode, 'TrackAll', 'the document is still tracking afterwards');

  // and mode Off suppresses tracking for the call only
  const untrackedText = 'plain';
  await pkg.withTracking({ author: { name: 'Agent' }, mode: 'Off' }, () => {
    pkg.body.insertParagraph(untrackedText, 'End');
  });
  const added = pkg.body.paragraphs.at(-1);
  assert.equal(added.text, untrackedText);
  assert.equal(added.getTrackedChanges().length, 0, 'the paragraph added with mode Off is not a revision');
  assert.equal(pkg.changeTrackingMode, 'TrackAll');
});

// CR-002 section 31, from the review of the release: the id counter lived on the tracker, and
// withTracking put the document's tracker back with the counter it had before the call - so the next
// revision after it reused the id the call's last revision was given.
test('withTracking: a revision after the call does not reuse an id given out inside it', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.author = { name: 'Ada' };
  pkg.body.insertParagraph('text', 'End');
  await pkg.setChangeTrackingMode('TrackAll');
  const end = () => pkg.body.paragraphs[0].getRange('End');
  end().insertText(' a', 'End');
  await pkg.withTracking({ author: { name: 'Agent' } }, () => end().insertText(' b', 'End'));
  end().insertText(' c', 'End');
  await pkg.withTracking({ author: { name: 'Agent' }, mode: 'TrackMineOnly' }, () => end().insertText(' d', 'End'));
  end().insertText(' e', 'End');

  const changes = pkg.body.getTrackedChanges();
  assert.deepEqual(changes.map((c) => c.author), ['Ada', 'Agent', 'Ada', 'Agent', 'Ada']);
  const ids = changes.map((c) => c.id);
  assert.equal(new Set(ids).size, ids.length, `every id is distinct: ${ids}`);
});

test('withTracking calls that overlap end in either order and leave the package as it was', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.author = { name: 'Ada' };
  pkg.body.insertParagraph('text', 'End');
  const gate = () => { let open; const shut = new Promise((resolve) => { open = resolve; }); return { shut, open }; };
  const a = gate();
  const b = gate();
  const entered = new Set();
  const first = pkg.withTracking({ author: { name: 'A' } }, async () => { entered.add('A'); await a.shut; });
  const second = pkg.withTracking({ author: { name: 'B' } }, async () => { entered.add('B'); await b.shut; });
  while (entered.size < 2) await new Promise((resolve) => setImmediate(resolve));

  // while both run, the later call's markup is the package's (nothing can tell whose an edit is)
  assert.equal(pkg.author.name, 'B');
  assert.equal(pkg.changeTrackingMode, 'TrackAll');
  // the earlier one ends first: the later one still applies
  a.open();
  await first;
  assert.equal(pkg.author.name, 'B');
  assert.equal(pkg.changeTrackingMode, 'TrackAll');
  // and when the last one ends the package is as it was before the first - it used to be left with
  // A's author and tracking on for good
  b.open();
  await second;
  assert.equal(pkg.author.name, 'Ada');
  assert.equal(pkg.changeTrackingMode, 'Off');
  assert.equal(pkg.trackedChangeDate, undefined);

  // a nested call hands back to its outer one, and that to the package
  await pkg.withTracking({ author: { name: 'Outer' } }, async () => {
    await pkg.withTracking({ author: { name: 'Inner' }, mode: 'Off' }, () => {
      assert.equal(pkg.changeTrackingMode, 'Off');
    });
    assert.equal(pkg.author.name, 'Outer');
    assert.equal(pkg.changeTrackingMode, 'TrackAll');
  });
  assert.equal(pkg.author.name, 'Ada');
  assert.equal(pkg.changeTrackingMode, 'Off');
});

test('seedAnnotationIds reads a part once for the bytes it holds', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('comments-two.docx'));
  let reads = 0;
  for (const part of pkg.parts.values()) {
    if (typeof part.readContents !== 'function') continue;
    const read = part.readContents.bind(part);
    part.readContents = () => { reads++; return read(); };
  }
  await pkg.withTracking({ author: { name: 'Agent' } }, () => {});
  assert.ok(reads > 0, 'the first call reads the parts nobody has unmarshalled');
  reads = 0;
  await pkg.withTracking({ author: { name: 'Agent' } }, () => {});
  await pkg.seedAnnotationIds();
  assert.equal(reads, 0, 'later calls read nothing: an untouched part cannot change');

  // bytes that are set are read again, and an id in them counts
  const styles = pkg.getMainDocumentPart().styleDefinitionsPart;
  styles.setXml('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:style w:type="character" w:styleId="X"><w:rPr><w:b/><w:rPrChange w:id="900" w:author="x">'
    + '<w:rPr/></w:rPrChange></w:rPr></w:style></w:styles>');
  assert.equal(await pkg.seedAnnotationIds(), 900);
  assert.equal(reads, 1, 'the part whose bytes were set, and only it');
  reads = 0;
  await pkg.seedAnnotationIds();
  assert.equal(reads, 0, 'and not again while they stay the same');
});

// CR-002 section 29: a move is one change. Word's Review tab accepts or rejects the whole move from
// either half, the four range markers and both paragraph marks with it; Office JS cannot even list
// a move in Word 16.0.20326.20158 (test/README.md checks 16 and 17), so the oracle is check 18: each
// action done by hand in Word on a fresh copy of fixtures/revisions/revisions-word15.docx and saved
// into fixtures/revisions/check18/. Section 3 of that fixture is the move.
const MOVED = 'second paragraph of the move';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function loaded(name) {
  const pkg = await WordprocessingMLPackage.load(await fixture(name));
  await pkg.getMainDocumentPart().getContents();
  return pkg;
}

/** Section 3's blocks: element, paragraph properties, text, and how much move markup is left in each. */
async function moveSection(pkg) {
  const { parseXml, serializeXml } = await import('../dist/index.mjs');
  const body = parseXml(await pkg.getMainDocumentPart().getXml()).getElementsByTagNameNS(W, 'body')[0];
  const blocks = Array.from(body.childNodes).filter((n) => n.nodeType === 1);
  const start = blocks.findIndex((n) => n.localName === 'p' && n.textContent.trim() === '3. A move');
  const end = blocks.findIndex((n, i) => i > start && n.localName === 'p' && n.textContent.startsWith('4. '));
  return blocks.slice(start + 1, end).filter((n) => !n.textContent.startsWith('[')).map((n) => {
    const pPr = Array.from(n.childNodes).find((c) => c.localName === 'pPr');
    const moves = Array.from(n.getElementsByTagNameNS(W, '*')).filter((x) => x.localName.startsWith('move')).length
      + (n.localName.startsWith('move') ? 1 : 0);
    // ids, dates and session ids are Word's to renumber when it saves; what is compared is the markup
    const properties = pPr ? serializeXml(pPr).replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w:id|w:date|w16du:dateUtc|w:rsid\w*)="[^"]*"/g, '') : '';
    return [n.localName, properties, n.textContent, moves];
  });
}

/** The text of every paragraph outside section 3, in both views. */
function outsideTheMove(pkg) {
  const paragraphs = pkg.body.paragraphs;
  const start = paragraphs.findIndex((p) => p.text === '3. A move');
  const end = paragraphs.findIndex((p) => p.text.startsWith('4. '));
  const kept = paragraphs.filter((_, i) => i <= start || i >= end);
  return [kept.map((p) => p.text), kept.map((p) => p.getText({ view: 'original' }))];
}

for (const [scenario, half, action] of [
  ['accept-destination', 0, 'accept'], ['accept-source', 1, 'accept'],
  ['reject-destination', 0, 'reject'], ['reject-source', 1, 'reject'],
]) {
  test(`a move is one change: ${scenario} does what Word's Review tab did (check 18)`, async () => {
    const pkg = await loaded('revisions/revisions-word15.docx');
    const before = outsideTheMove(pkg);
    const changes = pkg.body.getTrackedChanges();
    // the section's change (section 35) holds the whole section's text, the moved words too: the halves are the run-level changes
    const move = changes.filter((c) => c.target.kind === 'run' && c.text.includes(MOVED));
    assert.deepEqual(move.map((c) => c.type), ['Added', 'Deleted'], 'the move is listed as its two halves, destination first');
    move[half][action]();

    const word = await loaded(`revisions/check18/${scenario}.docx`);
    assert.deepEqual(await moveSection(pkg), await moveSection(word), 'section 3 as Word left it: the whole move resolved, no marker left');
    assert.equal(pkg.body.getTrackedChanges().length, changes.length - 2, 'both halves gone, and nothing else');
    assert.deepEqual(outsideTheMove(pkg), before, 'the rest of the document untouched, in both views');
    // the other half, already resolved with it, does nothing when asked
    move[1 - half][action]();
    assert.deepEqual(await moveSection(pkg), await moveSection(word));
  });
}

for (const scenario of ['accept-all', 'reject-all']) {
  test(`a move is one change: ${scenario} leaves the document Word's does (check 18)`, async () => {
    const pkg = await loaded('revisions/revisions-word15.docx');
    if (scenario === 'accept-all') pkg.body.acceptAll(); else pkg.body.rejectAll();
    const word = await loaded(`revisions/check18/${scenario}.docx`);
    assert.deepEqual(await moveSection(pkg), await moveSection(word));
    assert.equal(pkg.body.getTrackedChanges().length, 0);
    assert.ok(!/<w:move/.test(await pkg.getMainDocumentPart().getXml()), 'no move markup anywhere');
    assert.deepEqual(pkg.body.paragraphs.map((p) => p.text), word.body.paragraphs.map((p) => p.text),
      'every paragraph of the document reads as in Word\'s, its other 22 revisions included');
  });
}

// CR-002 section 34, asked for by the editor: its export takes every revision id from the package's
// counter, and seeding it re-read the main part the editor had already imported itself - 1.06 s on a
// 255-page document. noteAnnotationIds records a part's highest id instead of reading it.
test('noteAnnotationIds: a part the caller has read is not read again by the seed', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('revisions/revisions-word15.docx'));
  const main = pkg.getMainDocumentPart();
  let reads = 0;
  const read = main.readContents.bind(main);
  main.readContents = () => { reads++; return read(); };

  await pkg.noteAnnotationIds(main, 1000);
  assert.equal(pkg.annotationIdFloor, 1000, 'the floor rises at once');
  await pkg.seedAnnotationIds();
  assert.equal(reads, 0, 'the seed does not read the noted part');
  assert.equal(main.isUnmarshalled, false);
  assert.equal(new ChangeTracker(pkg, 'TrackAll').nextId(), 1001, 'and the next id is above the noted one');

  // bytes set again are not the bytes noted: the next seed reads the part afresh
  main.setXml(await main.getXml());
  await pkg.seedAnnotationIds();
  assert.equal(reads, 1);

  await assert.rejects(() => pkg.noteAnnotationIds(main, -1), /non-negative integer/);

  // the number to vouch with is highestAnnotationId over the tree the caller read, which the index
  // exports for this (the editor's import cannot see every id: comment ranges, a numbering w:ins,
  // content it keeps opaque) - and it is the number the seed itself would have found
  const fresh = await WordprocessingMLPackage.load(await fixture('revisions/revisions-word15.docx'));
  const seeded = await fresh.seedAnnotationIds();
  const noted = await WordprocessingMLPackage.load(await fixture('revisions/revisions-word15.docx'));
  const tree = await noted.getMainDocumentPart().readContents();
  await noted.noteAnnotationIds(noted.getMainDocumentPart(), highestAnnotationId(tree));
  assert.equal(await noted.seedAnnotationIds(), seeded, 'the same floor as reading every part');
  const other = await WordprocessingMLPackage.createPackage();
  await assert.rejects(() => pkg.noteAnnotationIds(other.getMainDocumentPart(), 5), /not a part of this package/);
});

// CR-002 section 29, from checks 20 and 21 (Word 16.0.20326.20158, by hand; the files are the oracle).
// Text typed where a revision is splits the revision: the typist's w:ins between its halves with
// tracking on, plain text with it off; typed at a revision's edge it goes beside it, and typed at the
// start of moved text it goes outside the move's range. Before, tracked text typed inside another's
// revision went to the revision's edge - the wrong place - and untracked text joined the revision.

/** A paragraph as the spans of text it shows: [kind, author, text], adjacent spans of one kind and author merged. */
function spansOf(pkg, startsWith) {
  const p = pkg.body.paragraphs.find((x) => x.getText({ view: 'original' }).startsWith(startsWith) || x.text.startsWith(startsWith));
  const out = [];
  const add = (kind, author, text) => {
    const last = out[out.length - 1];
    if (last && last[0] === kind && last[1] === author) last[2] += text; else out.push([kind, author, text]);
  };
  const textOfRun = (r) => (r.value.content ?? []).map((c) => c.value?.value ?? '').join('');
  for (const item of p.p.content ?? []) {
    const name = item.name?.localPart;
    if (name === 'r') add('plain', '', textOfRun(item));
    else if (['ins', 'del', 'moveTo', 'moveFrom'].includes(name)) {
      const runs = item.value.customXmlOrSmartTagOrSdt ?? item.value.accOrBarOrBox ?? [];
      add(name, item.value.author, runs.filter((r) => r.name?.localPart === 'r').map(textOfRun).join(''));
    }
  }
  return out;
}

async function typedAt(file, author, tracked, startsWith, at, text) {
  const { Range } = await import('../dist/index.mjs');
  const pkg = await loaded(file);
  pkg.author = { name: author };
  await pkg.setChangeTrackingMode(tracked ? 'TrackAll' : 'Off');
  const p = pkg.body.paragraphs.find((x) => x.text.startsWith(startsWith));
  const offset = at(p);
  new Range(p, offset, offset).insertText(text, 'Start');
  return pkg;
}

test('typing into moved text, tracked, writes what Word wrote (check 20, cases 01 to 04)', async () => {
  const goes = (p) => p.search('goes ')[0].end;
  for (const [name, author, at, text, find] of [
    ['01-type-inside-a', 'Author A', goes, 'quickly ', 'The second paragraph'],
    ['02-type-inside-b', 'Author B', goes, 'quickly ', 'The second paragraph'],
    ['03-type-at-start', 'Author B', () => 0, 'Now ', 'Now The second'],
    ['04-type-at-end', 'Author B', (p) => p.text.length, ' Really.', 'The second paragraph'],
  ]) {
    const pkg = await typedAt('revisions/revisions-word15.docx', author, true, 'The second paragraph', at, text);
    const word = await loaded(`revisions/check20/${name}.docx`);
    assert.deepEqual(spansOf(pkg, find), spansOf(word, find), name);
    // and where the markers are: case 03's insertion is outside the move's range, as Word put it
    const xml = await pkg.getMainDocumentPart().getXml();
    if (name === '03-type-at-start') assert.match(xml, /<w:ins [^>]*><w:r><w:t xml:space="preserve">Now <\/w:t><\/w:r><\/w:ins><w:moveToRangeStart /);
  }
});

test('typing into a revision with tracking off is plain and splits it (check 21, re-run)', async () => {
  const inVery = (p) => p.search('very')[0].start + 2;
  for (const [name, author, at, text] of [
    ['21a-in-insertion', 'Author B', inVery, 'X'],
    ['21b-at-insertion-end', 'Author B', (p) => p.search('very ')[0].end, 'Y'],
    ['21d-in-own-insertion', 'Author A', inVery, 'X'],
  ]) {
    const pkg = await typedAt('revisions/revisions-word15.docx', author, false, 'The quick brown', at, text);
    const word = await loaded(`revisions/check21/${name}.docx`);
    assert.deepEqual(spansOf(pkg, 'The quick brown'), spansOf(word, 'The quick brown'), name);
  }
  // with tracking on, the typist's own w:ins between the halves: what check 21's first run recorded
  // (its files were replaced by the re-run; CR-002 section 29 keeps the XML)
  const on = await typedAt('revisions/revisions-word15.docx', 'Author B', true, 'The quick brown', inVery, 'X');
  assert.deepEqual(spansOf(on, 'The quick brown').slice(3, 6), [['ins', 'Author A', 've'], ['ins', 'Author B', 'X'], ['ins', 'Author A', 'ry ']]);
  const ids = spansOf(on, 'The quick brown').length && (await on.getMainDocumentPart().getXml()).match(/<w:ins w:id="(\d+)"/g);
  assert.equal(new Set(ids).size, ids.length, 'the split half has an id of its own');

  // a replacement that starts inside a revision is a deletion and then typing: plain, beside it
  const replaced = await loaded('tracked-changes.docx');
  await replaced.setChangeTrackingMode('Off');
  assert.equal(replaced.body.replaceText('An', 'XY', { matchCase: true }), 1);
  assert.deepEqual(spansOf(replaced, 'Here is some').slice(0, 2),
    [['plain', '', 'Here is some change tracking. XY'], ['ins', 'Jason Harrop', ' insertion']]);
});

test('accepting and rejecting an edited, re-moved or broken move leaves what Word left (check 20)', async () => {
  for (const name of ['01-type-inside-a', '03-type-at-start', '04-type-at-end', '06-delete-part-a', '07-delete-part-b',
    '08-delete-all', '10-move-again', '12-partner-missing', '13-moveTo-no-ranges']) {
    for (const action of ['accept', 'reject']) {
      const pkg = await loaded(`revisions/check20/${name}.docx`);
      const destination = pkg.body.getTrackedChanges().find((c) => c.target.kind === 'run' && c.target.revision === 'moveTo');
      destination[action]();
      const word = await loaded(`revisions/check20/${name}-${action}.docx`);
      assert.deepEqual(await moveSection(pkg), await moveSection(word), `${name} ${action}`);
    }
  }
  // case 14, its source's range markers gone, taken from the input on either side (the re-run):
  // the side clicked resolved, paragraph mark and all, and the other side a plain w:ins or w:del
  for (const [side, suffix] of [['moveFrom', 'input'], ['moveTo', 'dest-input']]) {
    for (const action of ['accept', 'reject']) {
      const pkg = await loaded('revisions/check20/input/moveFrom-no-ranges.docx');
      pkg.body.getTrackedChanges().find((c) => c.target.kind === 'run' && c.target.revision === side)[action]();
      const word = await loaded(`revisions/check20/14-moveFrom-no-ranges-${action}-${suffix}.docx`);
      assert.deepEqual(await moveSection(pkg), await moveSection(word), `14 ${side} ${action}`);
    }
  }
});

// --- CR-002 section 29, fix F: touching insertions or deletions are one change (checks 22, 23) ---

/** Each paragraph: its accepted text, its original text, and the revision on its mark. */
const paragraphsOf = (pkg) => pkg.body.paragraphs.map((p) => {
  const rPr = p.p.pPr?.rPr;
  return [p.text, p.getText({ view: 'original' }), rPr?.ins ? 'ins' : rPr?.del ? 'del' : ''];
});
const listingOf = (changes) => changes.map((c) => [c.type, c.author, c.text]);

test('touching deletions, text and paragraph marks, are one change, as Office JS lists them (check 22)', async () => {
  const file = 'revisions/check22/22-deleted.docx';
  const pkg = await loaded(file);
  assert.deepEqual(listingOf(pkg.body.getTrackedChanges()), [
    ['Deleted', 'Author A', 'Two.\r'], ['Deleted', 'Author A', '\r'], ['Deleted', 'Author A', 'six.\rSeven '],
  ]);
  // a paragraph lists the changes that start in it, cut at its end
  assert.deepEqual(pkg.body.paragraphs.map((p) => p.getTrackedChanges().map((c) => c.text)),
    [[], ['Two.\r'], [], ['\r'], ['six.\r'], [], [], []]);

  const original = paragraphsOf(pkg);
  const lone = { accept: [...original], reject: [...original] };            // the JSON's, for the lone mark
  lone.accept.splice(3, 2, ['Four.Five ', 'Four.Five six.', 'del']);
  lone.reject.splice(3, 1, ['Four.', 'Four.', '']);
  for (const action of ['accept', 'reject']) {
    for (const [index, expected] of [[0, `22-${action}-whole`], [1, lone[action]], [2, `22-${action}-across`]]) {
      const edited = await loaded(file);
      edited.body.getTrackedChanges()[index][action]();
      const word = typeof expected === 'string' ? paragraphsOf(await loaded(`revisions/check22/${expected}.docx`)) : expected;
      assert.deepEqual(paragraphsOf(edited), word, `${action} ${index}`);
      if (typeof expected === 'string') {
        // and each paragraph keeps the w14:paraId Word kept: the second's, where a mark went
        const wordIds = (await loaded(`revisions/check22/${expected}.docx`)).body.paragraphs.map((q) => q.p.paraId);
        assert.deepEqual(edited.body.paragraphs.map((q) => q.p.paraId), wordIds, `${action} ${index}: paragraph ids`);
      }
      assert.equal(edited.body.getTrackedChanges().length, 2, `${action} ${index}: the other two left`);
    }
  }
});

test('what makes one change: kind and author, not date; insertions as deletions (check 23)', async () => {
  const D1 = '2026-09-28T01:00:00Z';
  const D2 = '2026-09-28T02:00:00Z';
  let nextId = 100;
  const who = (author, date) => `w:id="${nextId++}" w:author="Author ${author}" w:date="${date}"`;
  const p = (inner, mark) => `<w:p>${mark ? `<w:pPr><w:rPr>${mark}</w:rPr></w:pPr>` : ''}${inner}</w:p>`;
  const r = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const ins = (a, d, text) => `<w:ins ${who(a, d)}>${r(text)}</w:ins>`;
  const del = (a, d, text) => `<w:del ${who(a, d)}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
  const insMark = (a, d) => `<w:ins ${who(a, d)}/>`;
  const delMark = (a, d) => `<w:del ${who(a, d)}/>`;
  // [markup, Office JS's listing, the paragraphs after accepting the first change, after rejecting it]
  const CASES = {
    'ins-paragraph': [p(r('One.')) + p(ins('A', D1, 'Two.'), insMark('A', D1)) + p(r('Three.')),
      [['Added', 'Author A', 'Two.\r']], ['One.', 'Two.', 'Three.'], ['One.', 'Three.']],
    'ins-across': [p(r('Four ') + ins('A', D1, 'five.'), insMark('A', D1)) + p(ins('A', D1, 'Six ') + r('seven.')),
      [['Added', 'Author A', 'five.\rSix ']], ['Four five.', 'Six seven.'], ['Four seven.']],
    'ins-mark-alone': [p(r('Eight.'), insMark('A', D1)) + p(r('Nine.')),
      [['Added', 'Author A', '\r']], ['Eight.', 'Nine.'], ['Eight.Nine.']],
    'del-paragraph-dates': [p(r('One.')) + p(del('A', D1, 'Two.'), delMark('A', D2)) + p(r('Three.')),
      [['Deleted', 'Author A', 'Two.\r']], ['One.', 'Three.'], ['One.', 'Two.', 'Three.']],
    'del-paragraph-authors': [p(r('One.')) + p(del('A', D1, 'Two.'), delMark('B', D1)) + p(r('Three.')),
      [['Deleted', 'Author A', 'Two.'], ['Deleted', 'Author B', '\r']], ['One.', '', 'Three.'], ['One.', 'Two.', 'Three.']],
    'del-same': [p(r('Ten ') + del('A', D1, 'eleven ') + del('A', D1, 'twelve') + r('.')),
      [['Deleted', 'Author A', 'eleven twelve']], ['Ten .'], ['Ten eleven twelve.']],
    'del-dates': [p(r('Ten ') + del('A', D1, 'eleven ') + del('A', D2, 'twelve') + r('.')),
      [['Deleted', 'Author A', 'eleven twelve']], ['Ten .'], ['Ten eleven twelve.']],
    'del-authors': [p(r('Ten ') + del('A', D1, 'eleven ') + del('B', D1, 'twelve') + r('.')),
      [['Deleted', 'Author A', 'eleven '], ['Deleted', 'Author B', 'twelve']], ['Ten .'], ['Ten eleven .']],
    'del-then-ins': [p(r('Ten ') + del('A', D1, 'eleven') + ins('A', D1, 'twelve') + r('.')),
      [['Deleted', 'Author A', 'eleven'], ['Added', 'Author A', 'twelve']], ['Ten twelve.'], ['Ten eleventwelve.']],
  };
  const make = async (blocks) => {
    const pkg = await WordprocessingMLPackage.createPackage();
    pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}"><w:body>${blocks}</w:body></w:document>`);
    await pkg.getMainDocumentPart().getContents();
    return pkg;
  };
  for (const [name, [blocks, listing, accepted, rejected]] of Object.entries(CASES)) {
    const pkg = await make(blocks);
    const changes = pkg.body.getTrackedChanges();
    assert.deepEqual(listingOf(changes), listing, name);
    for (const [action, expected] of [['accept', accepted], ['reject', rejected]]) {
      const edited = await make(blocks);
      edited.body.getTrackedChanges()[0][action]();
      assert.deepEqual(edited.body.paragraphs.map((q) => q.text), expected, `${name} ${action}`);
      assert.deepEqual(listingOf(edited.body.getTrackedChanges()), listing.slice(1), `${name} ${action}: the rest left`);
    }
  }
  // a group's date is its first piece's (Office JS gave the run's, not the mark's an hour later),
  // read as local wall-clock time as Office JS read it
  const dated = (await make(CASES['del-paragraph-dates'][0])).body.getTrackedChanges()[0];
  assert.equal(dated.date.toISOString(), new Date(2026, 8, 28, 1, 0, 0).toISOString());
});

test('a revision date as Word means it: w16du:dateUtc when there, else w:date as local time; written both ways, w16du declared', async () => {
  // Word's file: w:date="2026-09-28T10:54:00Z" (Jason's wall clock, UTC+10), w16du:dateUtc="2026-09-28T00:54:00Z"
  const word = await loaded('revisions/check22/22-deleted.docx');
  assert.equal(word.body.getTrackedChanges()[0].date.toISOString(), '2026-09-28T00:54:00.000Z', 'the true UTC, in any time zone');

  // written into Word's document, whose root lists w16du in mc:Ignorable: both attributes
  const when = new Date(Date.UTC(2026, 8, 28, 1, 30, 0));
  word.author = { name: 'Ada' };
  word.trackedChangeDate = when;
  await word.setChangeTrackingMode('TrackAll');
  word.body.paragraphs[0].getRange('End').insertText(' more', 'End');
  const written = word.body.getTrackedChanges().find((c) => c.author === 'Ada');
  assert.equal(written.date.toISOString(), when.toISOString(), 'read back as written');
  const xml = await word.getMainDocumentPart().getXml();
  const ins = xml.match(/<w:ins [^>]*w:author="Ada"[^>]*>/)[0];
  assert.match(ins, new RegExp(`w:date="${wordDate(when)}"`), 'w:date on the local wall clock');
  assert.match(ins, /w16du:dateUtc="2026-09-28T01:30:00Z"/, 'w16du:dateUtc in UTC');
  assert.match(xml, /mc:Ignorable="[^"]*\bw16du\b/, 'the declaration kept');

  // written into a created document, which declares nothing: w16du made ignorable first, as docx4j
  // does for w14, and then both attributes
  const created = await tracked(['one']);
  created.body.paragraphs[0].getRange('End').insertText(' two', 'End');
  const createdXml = await xmlOf(created);
  assert.match(createdXml, at(/<w:ins w:id="\d+"@UTC w:author="Ada" w:date="@DATE">/));
  assert.match(createdXml, /<w:document [^>]*mc:Ignorable="w16du"/, 'w16du listed as ignorable');
  assert.match(createdXml, /<w:document [^>]*xmlns:w16du="http:\/\/schemas.microsoft.com\/office\/word\/2023\/wordml\/word16du"/);
  assert.equal(created.body.getTrackedChanges()[0].date.toISOString(), DATE.toISOString(), 'read back as written');
  const reloaded = await WordprocessingMLPackage.load(await created.save());
  await reloaded.getMainDocumentPart().getContents();
  assert.equal(reloaded.body.getTrackedChanges()[0].date.toISOString(), DATE.toISOString(), 'and after a save');
});

// CR-002 section 29: Word's check 18 files, accept all and reject all, over the whole document. Word
// restores a list paragraph without the indent its w:pPrChange recorded when the list level gives it
// (section 9); the engine drops such an indent when it can read the numbering part.
test('acceptAll and rejectAll over the check 18 document leave what Word left, list indents included', async () => {
  const { parseXml, serializeXml } = await import('../dist/index.mjs');
  const blocksOf = async (pkg) => {
    const body = parseXml(await pkg.getMainDocumentPart().getXml()).getElementsByTagNameNS(W, 'body')[0];
    return Array.from(body.childNodes)
      .filter((n) => n.nodeType === 1 && n.localName !== 'sectPr' && !/^(bookmarkStart|bookmarkEnd)$/.test(n.localName))
      .map((n) => {
        const pPr = Array.from(n.childNodes).find((c) => c.localName === 'pPr');
        // ids, dates and session ids are Word's to renumber
        const properties = pPr ? serializeXml(pPr).replace(/ xmlns:\w+="[^"]*"/g, '')
          .replace(/ (w:id|w:date|w16du:dateUtc|w:rsid\w*|w14:\w+)="[^"]*"/g, '') : '';
        // the paragraph's w14:paraId: when a mark goes, Word keeps the second paragraph's
        return [n.localName, n.textContent, properties, n.getAttributeNS('http://schemas.microsoft.com/office/word/2010/wordml', 'paraId')];
      });
  };
  const read = async (name) => {
    const pkg = await WordprocessingMLPackage.load(await fixture(name));
    await pkg.getBody();                                                 // the property resolver reads numbering
    return pkg;
  };
  for (const action of ['accept', 'reject']) {
    const pkg = await read('revisions/revisions-word15.docx');
    pkg.body[`${action}All`]();
    assert.deepEqual(await blocksOf(pkg), await blocksOf(await read(`revisions/check18/${action}-all.docx`)), `${action}All`);
    assert.equal((await pkg.getMainDocumentPart().getXml()).includes('<w:trPr/>'), false, `${action}All: a resolved row's empty w:trPr goes, as Word writes none`);
  }

  // without the numbering part read there is nothing to compare with, and the recorded indent comes back
  const unread = await loaded('revisions/revisions-word15.docx');
  unread.body.rejectAll();
  const item = unread.body.paragraphs.find((p) => p.text === 'The first numbered item.');
  assert.deepEqual([item.p.pPr.ind.left, item.p.pPr.ind.hanging], [720, 360]);
});

// Found by the editor's C5: accepting a move carried a deletion nested in its moved text out of it,
// and resolving a broken move dissolved the rest into a new w:ins, so "all" left changes listed.
test('acceptAll and rejectAll leave nothing listed, over every revisions fixture', async () => {
  const { readdir } = await import('node:fs/promises');
  const check20 = (await readdir(new URL('./fixtures/revisions/check20/', import.meta.url))).filter((f) => /^\d\d-[A-Za-z-]+\.docx$/.test(f) && !/-(accept|reject)/.test(f));
  const files = ['tracked-changes.docx', 'revisions/revisions-word15.docx', 'revisions/check22/22-deleted.docx',
    ...check20.map((f) => `revisions/check20/${f}`), ...['partner-missing', 'moveTo-no-ranges', 'moveFrom-no-ranges'].map((f) => `revisions/check20/input/${f}.docx`)];
  assert.ok(check20.length >= 14, `check 20's case files found: ${check20.length}`);
  for (const file of files) {
    for (const action of ['accept', 'reject']) {
      const pkg = await loaded(file);
      pkg.body[`${action}All`]();
      assert.deepEqual(pkg.body.getTrackedChanges().map((c) => `${c.type} ${c.author}: ${c.text}`), [], `${file} ${action}All`);
    }
  }
});

// CR-002 section 29, check 25: Word, rejecting a list turned to bullets, drops a recorded indent the
// level gives even in part (the level-0 items recorded only hanging=360 of the level's left=720
// hanging=360), and keeps a level-2 item's level. A list-definition edit (25c) is not tracked at all.
test('rejecting a list change writes the recorded level and drops a recorded indent the level gives, as Word does (check 25)', async () => {
  const read = async (name) => {
    const pkg = await WordprocessingMLPackage.load(await fixture(`revisions/check25/${name}.docx`));
    await pkg.getBody();
    return pkg;
  };
  // the level as written: Word writes the recorded level, w:ilvl 0 where the record had none, even when
  // the current level differs (25e: recorded numId 1 alone, current ilvl 1)
  const listProperties = (pkg) => pkg.body.paragraphs.filter((p) => p.p.pPr?.numPr).map((p) => {
    const { numPr, ind } = p.p.pPr;
    return [p.text, numPr.ilvl?.val, Number(numPr.numId?.val), ind ? JSON.stringify(ind, (key, value) => (key === 'PARENT' || key === 'TYPE_NAME' ? undefined : value)) : null];
  });
  for (const [input, count, expected] of [['25a-bullets', 3, '25b-bullets-reject'], ['25e-bullet-demote', 1, '25f-bullet-demote-reject']]) {
    const pkg = await read(input);
    assert.equal(pkg.body.getTrackedChanges().length, count, input);
    pkg.body.rejectAll();
    assert.deepEqual(listProperties(pkg), listProperties(await read(expected)), `${input} rejected`);
  }

  // an indent of its own - an attribute the level does not give - is restored
  const direct = await read('25a-bullets');
  const one = direct.body.paragraphs.find((p) => p.text === 'One.');
  one.p.pPr.pPrChange.pPr.ind.left = 999;
  direct.body.rejectAll();
  assert.equal(direct.body.paragraphs.find((p) => p.text === 'One.').p.pPr.ind?.left, 999);

  // bullets clicked with a caret and Tab at a list's first item edited the list definition, which Word
  // does not track: 25c lists no change, and Reject All left 25d the same
  assert.equal((await read('25c-demote')).body.getTrackedChanges().length, 0);
});

// CR-002 section 29, check 24: Word keeps no formatting change once the formatting is back to what it
// recorded, whoever changed it back (24b the same author, 24c another; 24e a character style cleared).
test('a formatting change undone is no change at all, whoever undoes it, as Word does (check 24)', async () => {
  const read = async (name) => {
    const pkg = await WordprocessingMLPackage.load(await fixture(`revisions/check24/${name}.docx`));
    await pkg.getBody();
    return pkg;
  };
  const bodyXml = async (pkg) => (await pkg.getMainDocumentPart().getXml()).match(/<w:body>(.*?)<w:sectPr/)[1]
    .replace(/ w:(rsid\w*|date)="[^"]*"| w16du:dateUtc="[^"]*"| w14:\w+="[^"]*"|<w:proofErr [^>]*\/>/g, '');
  for (const [author, expected] of [['Author A', '24b-unbold'], ['Author B', '24c-unbold-by-b']]) {
    const pkg = await read('24a-bold');                                  // "beta" bolded by Author A, tracked
    pkg.author = { name: author };
    await pkg.setChangeTrackingMode('TrackAll');
    pkg.body.paragraphs[0].search('beta', { matchCase: true })[0].font.bold = false;
    assert.equal(await bodyXml(pkg), await bodyXml(await read(expected)), `unbolded by ${author}`);
    assert.deepEqual(pkg.body.getTrackedChanges(), []);
  }

  // undone in part, it is still a change: bold and italic, then bold off
  const partial = await tracked(['alpha beta gamma']);
  const word = () => partial.body.paragraphs[0].search('beta', { matchCase: true })[0];
  word().font.bold = true;
  word().font.italic = true;
  word().font.bold = false;
  assert.deepEqual(partial.body.getTrackedChanges().map((c) => c.type), ['Formatted']);

  // a link made and removed under tracking restyles its runs and back: nothing is left to review
  const linked = await tracked(['alpha beta gamma']);
  const beta = () => linked.body.paragraphs[0].search('beta', { matchCase: true })[0];
  beta().hyperlink = 'https://example.com/';
  assert.ok(linked.body.getTrackedChanges().length > 0, 'the restyle is a formatting change');
  beta().hyperlink = '';
  assert.deepEqual(linked.body.getTrackedChanges(), []);
  assert.equal((await xmlOf(linked)).includes('rPrChange'), false);
});

// CR-002 section 29, check 26: Word keeps no paragraph formatting change once the paragraph is as it
// was, whoever changes it back, and writes no w:pPr (Ctrl+L, Decrease Indent remove the direct values).
test('a paragraph formatting change undone is no change at all, as Word does (check 26)', async () => {
  const read = async (name) => {
    const pkg = await WordprocessingMLPackage.load(await fixture(`revisions/check26/${name}.docx`));
    await pkg.getBody();                                                 // the property resolver compares
    return pkg;
  };
  const bodyXml = async (pkg) => (await pkg.getMainDocumentPart().getXml()).match(/<w:body>(.*?)<w:sectPr/)[1]
    .replace(/ w:(rsid\w*|date)="[^"]*"| w16du:dateUtc="[^"]*"| w14:\w+="[^"]*"|<w:proofErr [^>]*\/>/g, '');
  for (const [from, author, index, act, expected] of [
    ['26a-centred', 'Author A', 0, (p) => { p.alignment = 'Left'; }, '26b-left-again'],
    ['26a-centred', 'Author B', 0, (p) => { p.alignment = 'Left'; }, '26c-left-by-b'],
    ['26d-indented', 'Author A', 1, (p) => { p.leftIndent = 0; }, '26d-indented-back'],
  ]) {
    const pkg = await read(from);
    pkg.author = { name: author };
    await pkg.setChangeTrackingMode('TrackAll');
    act(pkg.body.paragraphs[index]);
    assert.equal(await bodyXml(pkg), await bodyXml(await read(expected)), `${from} undone by ${author}`);
    assert.deepEqual(pkg.body.getTrackedChanges(), []);
  }

  // undone in part, it is still a change: centred and indented, then left again
  const partial = await tracked(['one', 'two']);
  await partial.getPropertyResolver();
  const p = partial.body.paragraphs[0];
  p.alignment = 'Centered';
  p.leftIndent = 36;
  p.alignment = 'Left';
  assert.deepEqual(partial.body.getTrackedChanges().map((c) => c.type), ['Formatted']);

  // a list attached and detached again under tracking leaves nothing to review
  const listed = await tracked(['one', 'two']);
  await listed.getPropertyResolver();
  await listed.body.paragraphs[0].startNewList();
  assert.equal(listed.body.getTrackedChanges().filter((c) => c.target.kind === 'paragraphProperties').length, 1, 'attached: a formatting change');
  listed.body.paragraphs[0].detachFromList();
  assert.deepEqual(listed.body.getTrackedChanges().filter((c) => c.target.kind === 'paragraphProperties'), [], 'detached: none');
});

// CR-002 section 29 (fix F) and 35, check 28: bold put on an italic run and a plain one is one change,
// as Office JS lists it ("alpha beta"), whatever each run recorded.
test('touching formatting changes by one author are one change, as Office JS lists them (check 28)', async () => {
  const read = async (name) => {
    const pkg = await WordprocessingMLPackage.load(await fixture(`revisions/check28/${name}.docx`));
    await pkg.getBody();
    return pkg;
  };
  const alpha = (pkg) => pkg.body.paragraphs.find((p) => p.text.startsWith('alpha')).p.content
    .filter((el) => el.name?.localPart === 'r')
    .map((r) => [(r.value.content ?? []).map((c) => c.value?.value ?? '').join(''), Boolean(r.value.rPr?.b), Boolean(r.value.rPr?.i)]);
  const pkg = await read('28a-changes');
  // Office JS's listing, the two tables' property changes (section 35) with it
  assert.deepEqual(pkg.body.getTrackedChanges().map((c) => [c.type, c.author, c.text]), [
    ['Formatted', 'Author A', '1a\t1b\r\n2a\t2b\r\n3a\t3b\r\n'],
    ['Formatted', 'Author A', 'x1\ty1\r\n'],
    ['Formatted', 'Author A', 'alpha beta'],
  ]);
  for (const [action, expected] of [['accept', '28a-accept-all'], ['reject', '28a-reject-all']]) {
    const edited = await read('28a-changes');
    edited.body.getTrackedChanges()[2][action]();
    assert.deepEqual(alpha(edited), alpha(await read(expected)), action);
    assert.equal(edited.body.getTrackedChanges().length, 2, `${action}: the tables' changes left`);
  }
});

// --- CR-002 section 35: the tracked changes that were not listed (checks 18, 27 and 28) ---

/** A part's tables and body section, ids, dates and session ids aside, paragraphs as their text. */
async function tablesAndSection(pkg) {
  const clean = (s) => s.replace(/ xmlns:\w+="[^"]*"| w:(rsid\w*|date)="[^"]*"| w16du:dateUtc="[^"]*"| w14:\w+="[^"]*"|<w:proofErr [^>]*\/>/g, '')
    .replace(/<w:p>.*?<\/w:p>|<w:p [^>]*>.*?<\/w:p>/g, (p) => `[${[...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('')}]`);
  const xml = await pkg.getMainDocumentPart().getXml();
  return {
    tables: [...xml.matchAll(/<w:tbl>.*?<\/w:tbl>/gs)].map((m) => clean(m[0])),
    section: clean(xml.match(/<w:sectPr[ >](?:(?!<w:sectPr).)*?<\/w:sectPr>\s*<\/w:body>/s)?.[0] ?? ''),
  };
}

test('accept all and reject all resolve table and section property changes as Word did (checks 18 and 28)', async () => {
  const read = async (name) => {
    const pkg = await WordprocessingMLPackage.load(await fixture(`revisions/${name}.docx`));
    await pkg.getBody();
    return pkg;
  };
  for (const [source, accepted, rejected] of [['revisions-word15', 'check18/accept-all', 'check18/reject-all'], ['check28/28a-changes', 'check28/28a-accept-all', 'check28/28a-reject-all']]) {
    for (const [action, word] of [['accept', accepted], ['reject', rejected]]) {
      const pkg = await read(source);
      pkg.body[`${action}All`]();
      assert.deepEqual(await tablesAndSection(pkg), await tablesAndSection(await read(word)), `${source} ${action}All`);
    }
  }
  // listed where they stand: the table's change once, the body's section last
  const listing = (await read('revisions-word15')).body.getTrackedChanges().filter((c) => ['tableProperties', 'sectionProperties'].includes(c.target.kind));
  assert.deepEqual(listing.map((c) => [c.target.kind, c.type, c.author]), [['tableProperties', 'Formatted', 'Author A'], ['sectionProperties', 'Formatted', 'Author A']]);
  assert.equal(listing[0].text, 'Top left\tTop right\r\nBottom left\tBottom right\r\n');
});

test('table, cell, mark and section changes one at a time, as Office JS listed and resolved them (check 27)', async () => {
  const A = 'w:author="Author A" w:date="2026-09-28T01:00:00Z"';
  let nextId = 500;
  const id = () => `w:id="${nextId++}"`;
  const para = (text, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const tc = (text, tcPr) => `<w:tc><w:tcPr>${tcPr}</w:tcPr>${para(text)}</w:tc>`;
  const make = async (blocks) => {
    const pkg = await WordprocessingMLPackage.createPackage();
    pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}"><w:body>${blocks}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
    await pkg.getBody();
    return pkg;
  };
  const listing = (pkg) => pkg.body.getTrackedChanges().map((c) => [c.target.kind, c.type, c.text]);

  // a width change: Word records the old cells over a grid of the old and new edges together, and
  // rejecting collapses it again (check 27, table-width: 2000/1152/853/1995/304/3152, spans 2, 3, 1)
  const W3152 = '<w:tcW w:w="3152" w:type="dxa"/>';
  const widthRow = (n) => `<w:tr>${tc(`a${n}`, `${W3152}<w:tcPrChange ${id()} ${A}><w:tcPr>${W3152}<w:gridSpan w:val="2"/></w:tcPr></w:tcPrChange>`)}`
    + `${tc(`b${n}`, `${W3152}<w:tcPrChange ${id()} ${A}><w:tcPr>${W3152}<w:gridSpan w:val="3"/></w:tcPr></w:tcPrChange>`)}`
    + `${tc(`c${n}`, `${W3152}<w:tcPrChange ${id()} ${A}><w:tcPr>${W3152}</w:tcPr></w:tcPrChange>`)}</w:tr>`;
  const widths = para('Before.') + `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblPrChange ${id()} ${A}><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange></w:tblPr>`
    + `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2005"/><w:gridCol w:w="1995"/><w:tblGridChange ${id()}><w:tblGrid>`
    + [2000, 1152, 853, 1995, 304, 3152].map((w) => `<w:gridCol w:w="${w}"/>`).join('') + '</w:tblGrid></w:tblGridChange></w:tblGrid>'
    + widthRow(1) + widthRow(2) + '</w:tbl>' + para('After.');
  const width = await make(widths);
  assert.deepEqual(listing(width), [['tableProperties', 'Formatted', 'a1\tb1\tc1\r\na2\tb2\tc2\r\n']]);
  width.body.getTrackedChanges()[0].reject();
  const rejected = (await tablesAndSection(width)).tables[0];
  assert.match(rejected, /<w:tblGrid><w:gridCol w:w="3152"\/><w:gridCol w:w="3152"\/><w:gridCol w:w="3152"\/><\/w:tblGrid>/);
  assert.match(rejected, /<w:tblW w:type="auto" w:w="0"\/>/);
  assert.equal(/gridSpan|Change/.test(rejected), false, 'no spans left, no records');

  // a cell's shading: one row changed, the table's own records unchanged, listed with the row's text
  const same = `<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblPrChange ${id()} ${A}><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange></w:tblPr>`
    + `<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:tblGridChange ${id()}><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid></w:tblGridChange></w:tblGrid>`;
  const W3000 = '<w:tcW w:w="3000" w:type="dxa"/>';
  const shaded = await make(para('Before.') + `<w:tbl>${same}<w:tr>${tc('a1', `${W3000}<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:tcPrChange ${id()} ${A}><w:tcPr>${W3000}</w:tcPr></w:tcPrChange>`)}`
    + `${tc('b1', `${W3000}<w:tcPrChange ${id()} ${A}><w:tcPr>${W3000}</w:tcPr></w:tcPrChange>`)}</w:tr><w:tr>${tc('a2', W3000)}${tc('b2', W3000)}</w:tr></w:tbl>` + para('After.'));
  assert.deepEqual(listing(shaded), [['tableProperties', 'Formatted', 'a1\tb1\r\n']]);

  // an inserted and a deleted cell: listed with the row's text; the cell kept or taken, the one
  // before it spanning its column (check 27: w:gridSpan 2, w:tcW 6000)
  const cellTable = (marker) => para('Before.') + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>`
    + `<w:tr>${tc('a1', W3000)}${tc('b1', W3000)}${tc('c1', `${W3000}<w:${marker} ${id()} ${A}/>`)}</w:tr></w:tbl>` + para('After.');
  const cellsOfRow = async (pkg) => [...(await tablesAndSection(pkg)).tables[0].matchAll(/<w:tc>(.*?)<\/w:tc>/g)].map((m) => m[1]);
  for (const [marker, type, keep, take] of [['cellIns', 'Added', 'accept', 'reject'], ['cellDel', 'Deleted', 'reject', 'accept']]) {
    const kept = await make(cellTable(marker));
    assert.deepEqual(listing(kept), [['cell', type, 'a1\tb1\tc1\r\n']]);
    kept.body.getTrackedChanges()[0][keep]();
    assert.deepEqual(await cellsOfRow(kept), ['<w:tcPr><w:tcW w:type="dxa" w:w="3000"/></w:tcPr>[a1]', '<w:tcPr><w:tcW w:type="dxa" w:w="3000"/></w:tcPr>[b1]', '<w:tcPr><w:tcW w:type="dxa" w:w="3000"/></w:tcPr>[c1]'], `${marker} ${keep}`);
    const taken = await make(cellTable(marker));
    taken.body.getTrackedChanges()[0][take]();
    assert.deepEqual(await cellsOfRow(taken), ['<w:tcPr><w:tcW w:type="dxa" w:w="3000"/></w:tcPr>[a1]', '<w:tcPr><w:tcW w:type="dxa" w:w="6000"/><w:gridSpan w:val="2"/></w:tcPr>[b1]'], `${marker} ${take}`);
  }

  // a mark made bold with its run: one change, "Before.\r" (check 27, mark-bold); accept keeps both
  // bold, reject takes both off
  const bold = `<w:p><w:pPr><w:rPr><w:b/><w:rPrChange ${id()} ${A}><w:rPr/></w:rPrChange></w:rPr></w:pPr><w:r><w:rPr><w:b/><w:rPrChange ${id()} ${A}><w:rPr/></w:rPrChange></w:rPr><w:t>Before.</w:t></w:r></w:p>` + para('After.');
  const boldened = await make(bold);
  assert.deepEqual(listing(boldened), [['group', 'Formatted', 'Before.\r']]);
  for (const [action, bolded] of [['accept', true], ['reject', false]]) {
    const pkg = await make(bold);
    pkg.body.getTrackedChanges()[0][action]();
    const p = pkg.body.paragraphs[0].p;
    assert.deepEqual([Boolean(p.pPr?.rPr?.b), Boolean(p.content[0].value.rPr?.b), Boolean(p.pPr?.rPr?.rPrChange)], [bolded, bolded, false], action);
  }

  // a section break's properties: "The end of section one.\f"; reject lays the record over them
  const sectioned = para('The end of section one.', `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/><w:sectPrChange ${id()} ${A}><w:sectPr><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:docGrid w:linePitch="0"/></w:sectPr></w:sectPrChange></w:sectPr>`) + para('Section two.');
  const section = await make(sectioned);
  assert.deepEqual(listing(section), [['sectionProperties', 'Formatted', 'The end of section one.\f']]);
  section.body.getTrackedChanges()[0].reject();
  const sectPr = section.body.paragraphs[0].p.pPr.sectPr;
  assert.deepEqual([sectPr.pgSz.w, sectPr.pgMar.top, sectPr.docGrid, sectPr.sectPrChange], [11906, 1440, undefined, undefined]);

  // never listed, resolved by "all": a cell merge (Office JS throws listing one) and a numbering change
  const unlisted = para('Before.') + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>`
    + `<w:tr>${tc('a1', `${W3000}<w:vMerge w:val="restart"/>`)}</w:tr><w:tr>${tc('', `${W3000}<w:vMerge/><w:cellMerge ${id()} ${A} w:vMerge="cont"/>`)}</w:tr></w:tbl>`
    + para('Numbered.', `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:numberingChange ${id()} ${A} w:original="%1."/></w:numPr>`);
  for (const action of ['accept', 'reject']) {
    const pkg = await make(unlisted);
    assert.deepEqual(listing(pkg), []);
    pkg.body[`${action}All`]();
    const xml = await pkg.getMainDocumentPart().getXml();
    assert.equal(/cellMerge|numberingChange/.test(xml), false, `${action}All: both gone`);
    assert.equal(/<w:vMerge\/>/.test(xml), action === 'accept', `${action}All: the merge ${action === 'accept' ? 'kept' : 'undone'}`);
  }
});

// CR-002 section 37, the editor's finding on 0.3.0: a tracked replacement inside another author's
// insertion left the new w:ins nested in it; Word splits the insertion around it, as for anything
// typed in it (check 21), the replacement's w:del staying in the first half.
test('a tracked replacement inside another\'s insertion splits it around the new text, as Word does', async () => {
  const shape = (p) => p.p.content.map((el) => {
    const name = el.name.localPart;
    if (name === 'r') return `r "${el.value.content.map((c) => c.value.value ?? '').join('')}"`;
    const inner = (el.value.customXmlOrSmartTagOrSdt ?? []).map((i) => i.name.localPart === 'r'
      ? `"${i.value.content.map((c) => c.value.value ?? '').join('')}"`
      : `${i.name.localPart}(${i.value.author}) "${(i.value.customXmlOrSmartTagOrSdt ?? []).map((r) => r.value.content.map((c) => c.value.value ?? '').join('')).join('')}"`);
    return `${name}(${el.value.author}) [${inner.join(', ')}]`;
  });
  const replaced = async (find) => {
    const pkg = await loaded('tracked-changes.docx');
    pkg.author = { name: 'Author C' };
    await pkg.setChangeTrackingMode('TrackAll');
    const p = pkg.body.paragraphs.find((q) => q.text.includes('An insertion'));
    p.search(find, { matchCase: true })[0].insertText('W', 'Replace');
    const ids = (await pkg.getMainDocumentPart().getXml()).match(/<w:(ins|del) w:id="\d+"/g).map((m) => m.match(/\d+/)[0]);
    assert.equal(new Set(ids).size, ids.length, `${find}: every revision id unique`);
    return shape(p).slice(1, 4);
  };
  // at its start: the editor's case
  assert.deepEqual(await replaced('An'), [
    'ins(Jason Harrop) [del(Author C) "An"]', 'ins(Author C) ["W"]', 'ins(Jason Harrop) [" insertion"]',
  ]);
  // in its middle
  assert.deepEqual(await replaced('ser'), [
    'ins(Jason Harrop) ["An in", del(Author C) "ser"]', 'ins(Author C) ["W"]', 'ins(Jason Harrop) ["tion"]',
  ]);
  // at its end the new text goes after it (check 21b)
  assert.deepEqual((await replaced('insertion')).slice(0, 2), [
    'ins(Jason Harrop) ["An ", del(Author C) "insertion"]', 'ins(Author C) ["W"]',
  ]);
});
