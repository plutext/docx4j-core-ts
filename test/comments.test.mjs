// CR-002 phase G: comments. The four comment parts, the Comment views, threads and round trips.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { WordprocessingMLPackage, Comment, Range, Paragraph } from '../dist/index.mjs';
import { fixture, bytesEqual } from './helpers.mjs';

const CEX = 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible';

/** The bytes of a part of a saved package, by part name. */
async function partBytes(bytes, name) {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return pkg.getPart(name).getBytes();
}

test('reads a comment, its author, date and range from a document Word wrote', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('loadAndSave.docx'));
  const body = await pkg.getBody();
  const comments = await body.getComments();
  assert.equal(comments.length, 1);
  const [comment] = comments;
  assert.ok(comment instanceof Comment);
  assert.equal(comment.id, 4);
  assert.equal(comment.authorName, 'Buck Cronk');
  assert.equal(comment.initials, 'BC');
  assert.equal(comment.authorEmail, 'buck.cronk@oracle.com', 'from w:people, unwrapped from the AD form');
  assert.equal(comment.content, 'Comment');
  assert.equal(comment.paraId, '36593EBD');
  assert.equal(comment.creationDate.toISOString(), '2026-05-18T18:21:00.000Z');
  assert.equal(comment.resolved, false);
  assert.deepEqual(comment.replies, []);
  const ranges = comment.getRange();
  assert.equal(ranges.length, 1);
  assert.ok(ranges[0] instanceof Range);
  assert.equal(ranges[0].text, 'Lorem ipsum…');
  // the same comment through the paragraph and the range it marks
  const paragraph = ranges[0].paragraph;
  assert.deepEqual((await paragraph.getComments()).map((c) => c.id), [4]);
  assert.deepEqual((await ranges[0].getComments()).map((c) => c.id), [4]);
  assert.deepEqual((await paragraph.getRange('Start').getComments()).map((c) => c.id), [4]);
  // a paragraph with no markers has none
  const other = body.paragraphs.find((p) => p.element !== paragraph.element);
  assert.deepEqual(await other.getComments(), []);
  // reading comments does not unmarshal the bookkeeping parts
  const main = pkg.getMainDocumentPart();
  assert.equal(main.commentsIdsPart.isUnmarshalled, false, 'w16cid is only needed when comments change');
});

test('reads the two comments of a document with no w15 parts', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('comments-two.docx'));
  const body = await pkg.getBody();
  const comments = await body.getComments();
  assert.deepEqual(comments.map((c) => [c.id, c.authorName, c.content]), [[0, 'jharrop2', 'One comment'], [1, 'jharrop2', 'blagh']]);
  assert.equal(comments[0].authorEmail, '', 'no people part');
  assert.equal(comments[0].resolved, false, 'no commentsExtended part');
  assert.equal(comments[0].paragraphs.length, 1);
  assert.equal(comments[0].paragraphs[0].styleId, 'CommentText');
  assert.equal(comments[0].paragraphs[0].style, 'Comment Text');
});

test('a document whose comments are not touched round-trips its comment parts byte for byte', async () => {
  const source = await fixture('loadAndSave.docx');
  const pkg = await WordprocessingMLPackage.load(source);
  await pkg.getBody();                       // the main part is re-marshalled, the comment parts are not
  const saved = await pkg.save();
  for (const name of ['/word/comments.xml', '/word/commentsExtended.xml', '/word/commentsIds.xml', '/word/people.xml', '/word/commentsExtensible.xml']) {
    const before = await (await WordprocessingMLPackage.load(source)).getPart(name).getBytes();
    const after = await partBytes(saved, name);
    assert.ok(bytesEqual(before, after), `${name} changed`);
  }
});

test('insertComment into a new package: markers, the comment, the side parts and the styles', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.author = { name: 'Ada Lovelace', initials: 'AL', email: 'ada@example.com' };
  const body = pkg.body;
  const paragraph = body.insertParagraph('Hello comment world', 'End');
  const [range] = paragraph.search('comment');
  const comment = await range.insertComment('Is this right?');
  assert.equal(comment.id, 0);
  assert.equal(comment.authorName, 'Ada Lovelace');
  assert.equal(comment.initials, 'AL');
  assert.equal(comment.authorEmail, 'ada@example.com');
  assert.equal(comment.content, 'Is this right?');
  assert.equal(comment.resolved, false);
  assert.ok(comment.paraId, 'a w14:paraId keys the side entries');
  assert.deepEqual(comment.getRange().map((r) => r.text), ['comment']);

  const main = pkg.getMainDocumentPart();
  // the four parts, with their relationships and content types
  assert.ok(main.commentsPart && main.commentsExtendedPart && main.commentsIdsPart && main.peoplePart);
  assert.equal(main.commentsPart.partName.name, '/word/comments.xml');
  assert.ok(main.relationshipsPart.getRelationshipByType('http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'));
  // the comment: CommentText style, opened by the annotation reference run
  const [written] = main.commentsPart.contents.comment;
  assert.equal(written.content[0].value.pPr.pStyle.val, 'CommentText');
  assert.equal(written.content[0].value.content[0].value.content[0].name.localPart, 'annotationRef');
  assert.equal(written.content[0].value.content[0].value.rPr.rStyle.val, 'CommentReference');
  // the markers in the body, around the commented run
  const names = paragraph.p.content.map((e) => e.name.localPart);
  assert.deepEqual(names, ['r', 'commentRangeStart', 'r', 'commentRangeEnd', 'r', 'r'], 'the run was split at the boundaries');
  assert.equal(paragraph.p.content[1].value.id, 0);
  assert.equal(paragraph.p.content[3].value.id, 0);
  assert.equal(paragraph.p.content[4].value.content[0].name.localPart, 'commentReference');
  assert.equal(paragraph.p.content[4].value.rPr.rStyle.val, 'CommentReference');
  assert.equal(paragraph.text, 'Hello comment world', 'the text is unchanged');
  // the side-part entries, keyed by the paraId
  assert.deepEqual(main.commentsExtendedPart.contents.commentEx.map((e) => [e.paraId, e.done]), [[comment.paraId, '0']]);
  assert.equal(main.commentsIdsPart.contents.commentId[0].paraId, comment.paraId);
  assert.match(main.commentsIdsPart.contents.commentId[0].durableId, /^[0-9A-F]{8}$/);
  assert.deepEqual(main.peoplePart.contents.person.map((p) => [p.author, p.presenceInfo.userId]), [['Ada Lovelace', 'ada@example.com']]);
  // the styles Word needs
  const styles = await main.styleDefinitionsPart.getContents();
  assert.ok(styles.style.some((s) => s.styleId === 'CommentText' && s.type === 'paragraph'));
  assert.ok(styles.style.some((s) => s.styleId === 'CommentReference' && s.type === 'character'));

  // saved, reloaded and found again
  const back = await WordprocessingMLPackage.load(await pkg.save());
  const [reloaded] = await (await back.getBody()).getComments();
  assert.equal(reloaded.content, 'Is this right?');
  assert.equal(reloaded.authorName, 'Ada Lovelace');
  assert.deepEqual(reloaded.getRange().map((r) => r.text), ['comment']);
});

test('reply, resolve and delete: the thread, w15:done and everything removed', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.author = { name: 'Grace Hopper' };
  const body = pkg.body;
  const paragraph = body.insertParagraph('A first draft', 'End');
  const comment = await paragraph.insertComment('Needs work');
  assert.equal(comment.initials, 'GH', 'initials default to the name');
  const reply = await comment.reply('Agreed');
  assert.equal(reply.parent.id, comment.id);
  assert.deepEqual(comment.replies.map((r) => r.content), ['Agreed']);
  const main = pkg.getMainDocumentPart();
  const entries = main.commentsExtendedPart.contents.commentEx;
  assert.equal(entries.length, 2);
  assert.equal(entries[1].paraIdParent, comment.paraId, 'the reply is linked to its parent');
  // the reply's markers sit next to its parent's, as Word nests them
  assert.deepEqual(paragraph.p.content.map((e) => e.name.localPart), ['commentRangeStart', 'commentRangeStart', 'r', 'commentRangeEnd', 'commentRangeEnd', 'r', 'r']);

  comment.resolved = true;
  assert.equal(comment.resolved, true);
  assert.equal(entries[0].done, '1');
  comment.resolved = false;
  assert.equal(entries[0].done, '0');
  comment.resolved = true;

  // through a save: the thread is rebuilt from w15:paraIdParent
  const back = await WordprocessingMLPackage.load(await pkg.save());
  const backBody = await back.getBody();
  const top = await backBody.getComments();
  assert.equal(top.length, 1, 'a reply is not a top-level comment');
  assert.equal(top[0].resolved, true);
  assert.deepEqual(top[0].replies.map((r) => r.content), ['Agreed']);

  // content can be replaced
  top[0].content = 'Needs more work\nsecond paragraph';
  assert.equal(top[0].content, 'Needs more work\nsecond paragraph');
  assert.equal(top[0].paragraphs.length, 2);
  assert.equal(top[0].paraId, top[0].paragraphs[0].paraId, 'the paraId is kept, so the side entries still match');
  assert.equal(top[0].resolved, true);

  // delete takes the replies, the markers and the side entries with it
  await top[0].delete();
  assert.deepEqual(await backBody.getComments(), []);
  const backMain = back.getMainDocumentPart();
  assert.deepEqual(backMain.commentsPart.contents.comment, []);
  assert.deepEqual(backMain.commentsExtendedPart.contents.commentEx, []);
  assert.deepEqual(backMain.commentsIdsPart.contents.commentId, []);
  assert.deepEqual(backBody.paragraphs[0].p.content.map((e) => e.name.localPart), ['r'], 'the markers and the reference runs are gone');
  assert.equal(backBody.paragraphs[0].text, 'A first draft');
});

test('insertComment into a loaded document that has no comment parts', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('HelloWordOnline.docx'));
  const body = await pkg.getBody();
  assert.deepEqual(await body.getComments(), []);
  const paragraph = body.paragraphs[0];
  const comment = await paragraph.insertComment('First comment');
  assert.equal(comment.id, 0);
  assert.equal(comment.authorName, 'docx4j', 'the default author');
  const main = pkg.getMainDocumentPart();
  assert.ok(main.commentsPart && main.commentsExtendedPart && main.commentsIdsPart && main.peoplePart);

  const bytes = await pkg.save();
  const back = await WordprocessingMLPackage.load(bytes);
  const comments = await (await back.getBody()).getComments();
  assert.deepEqual(comments.map((c) => c.content), ['First comment']);
  // the new parts came back through [Content_Types].xml and the relationships (they loaded as their classes)
  assert.equal(back.getMainDocumentPart().commentsIdsPart.contentType, 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml');
  assert.equal(back.getMainDocumentPart().peoplePart.partName.name, '/word/people.xml');
});

test('a comment in a document that has a w16cex part keeps it in step', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('loadAndSave.docx'));
  const body = await pkg.getBody();
  const paragraph = body.paragraphs[0];
  const comment = await paragraph.insertComment('Added');
  const main = pkg.getMainDocumentPart();
  const rel = main.relationshipsPart.getRelationshipByType(CEX);
  const cex = await main.relationshipsPart.getPart(rel).getXml();
  const durableId = main.commentsIdsPart.contents.commentId.find((e) => e.paraId === comment.paraId).durableId;
  assert.ok(cex.includes(`w16cex:durableId="${durableId}"`), 'a w16cex entry was added');
  assert.match(cex, new RegExp(`w16cex:durableId="${durableId}" w16cex:dateUtc="\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}Z"`));
  await comment.delete();
  const after = await main.relationshipsPart.getPart(rel).getXml();
  assert.ok(!after.includes(`w16cex:durableId="${durableId}"`), 'and removed with the comment');
  assert.ok(after.includes('w16cex:durableId="05546856"'), "Word's own entry is untouched");
});

test('comments are returned in document order with replies nested; a range sees only what it touches', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const first = body.insertParagraph('alpha beta', 'End');
  const second = body.insertParagraph('gamma', 'End');
  const onBeta = await first.search('beta')[0].insertComment('about beta');
  const onAlpha = await first.search('alpha')[0].insertComment('about alpha');
  const onGamma = await second.insertComment('about gamma');
  await onAlpha.reply('a reply to alpha');
  const all = await body.getComments();
  assert.deepEqual(all.map((c) => c.content), ['about alpha', 'about beta', 'about gamma'], 'document order, by the first marker');
  assert.deepEqual(all[0].replies.map((c) => c.content), ['a reply to alpha']);
  assert.deepEqual((await second.getComments()).map((c) => c.id), [onGamma.id]);
  assert.deepEqual((await first.search('beta')[0].getComments()).map((c) => c.content), ['about beta']);
  assert.deepEqual((await first.getRange().getComments()).map((c) => c.content), ['about alpha', 'about beta']);
  assert.ok(onBeta instanceof Comment);
  assert.ok(first instanceof Paragraph);
});

test('anchors: an empty range gets a reference run only; cells and hyperlinks work too', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const paragraph = body.insertParagraph('alpha beta', 'End');
  const caret = paragraph.getRange();
  caret.start = 5;
  caret.end = 5;
  const atCaret = await caret.insertComment('at the caret');
  assert.deepEqual(paragraph.p.content.map((e) => e.name.localPart), ['r', 'r', 'r'], 'no range markers, the run split at the offset');
  assert.equal(paragraph.p.content[1].value.content[0].name.localPart, 'commentReference');
  assert.deepEqual(atCaret.getRange().map((r) => [r.start, r.end]), [[5, 5]]);
  assert.equal(paragraph.text, 'alpha beta');

  // in a table cell
  await body.insertXml('<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4513"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>', 'End');
  const cell = body.paragraphs.find((p) => p.text === 'cell text');
  const inCell = await cell.insertComment('on a cell');
  assert.deepEqual(inCell.getRange().map((r) => r.text), ['cell text']);

  // inside a hyperlink: the markers go in the hyperlink's own run list
  const [link] = await body.insertXml('<w:p><w:hyperlink r:id="rIdX"><w:r><w:t>linked text</w:t></w:r></w:hyperlink></w:p>', 'End');
  const inLink = await link.search('linked')[0].insertComment('on a hyperlink');
  assert.deepEqual(link.p.content[0].value.content.map((e) => e.name.localPart), ['commentRangeStart', 'r', 'commentRangeEnd', 'r', 'r']);
  assert.deepEqual(inLink.getRange().map((r) => r.text), ['linked']);
  assert.deepEqual((await body.getComments()).map((c) => c.content), ['at the caret', 'on a cell', 'on a hyperlink']);
  await inLink.delete();
  assert.deepEqual(link.p.content[0].value.content.map((e) => e.name.localPart), ['r', 'r']);
  assert.equal(link.text, 'linked text');
});
