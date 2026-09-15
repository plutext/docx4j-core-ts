// CR-002 phase I: the Word shim (@docx4j/core-ts/office-js) and toApiScript.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { WordprocessingMLPackage } from '../dist/index.mjs';
import { Word, toApiScript, NotSupportedError, ItemNotFoundError, ValueNotLoadedError, ClientResult, unwrap } from '../dist/office-js/index.mjs';
import { fixture } from './helpers.mjs';

/** The emitted script is a sequence of statements over `body`, with `await` on the insertXml lines. */
function asyncScript(script) {
  return new Function('body', `return (async () => {${script}})()`);
}

test('Word.run: an add-in callback against a created package, saved and reloaded', async () => {
  const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
  const seen = await Word.run(pkg, async (context) => {
    const body = context.document.body;
    body.insertParagraph('Hello World', 'End');
    const heading = body.insertParagraph('Title', 'Start');
    heading.styleBuiltIn = Word.Style.heading1;
    heading.alignment = Word.Alignment.centered;

    const paragraphs = body.paragraphs;
    // Office JS's load/sync lines are no-ops here, and must stay legal
    assert.equal(paragraphs.load('text'), paragraphs);
    assert.equal(paragraphs.items[0].load(), paragraphs.items[0]);
    await context.sync();
    assert.deepEqual(paragraphs.items.map((p) => p.text), ['Title', 'Hello World']);
    assert.equal(paragraphs.length, 2);
    assert.equal(paragraphs.getCount().value, 2);

    for (const range of body.search('World')) range.font.bold = true;

    context.document.properties.title = 'Phase I';
    context.document.properties.author = 'Claude';
    context.document.properties.keywords = 'office-js, docx4j';
    context.document.properties.creationDate = new Date(Date.UTC(2026, 8, 15, 10, 30, 0));
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    assert.equal(context.document.changeTrackingMode, 'TrackAll');

    const ooxml = body.getOoxml();
    assert.ok(ooxml instanceof ClientResult);
    assert.throws(() => ooxml.value, ValueNotLoadedError);
    await context.sync();
    assert.ok(ooxml.value.includes('pkg:package'), 'getOoxml is the flat OPC package');
    return body.text;
  });
  assert.equal(seen, 'Title\nHello World');

  const back = await WordprocessingMLPackage.load(await pkg.save());
  const body = await back.getBody();
  assert.deepEqual(body.paragraphs.map((p) => [p.text, p.style, p.alignment]), [['Title', 'Heading 1', 'Centered'], ['Hello World', 'Normal', 'Unknown']]);
  assert.equal(body.paragraphs[1].runs.filter((r) => r.value.rPr?.b).length, 1, 'the searched word is bold');
  const core = await back.docPropsCorePart.getContents();
  assert.equal(core.title.value.content[0], 'Phase I');
  assert.equal(core.creator.content[0], 'Claude');
  assert.equal(core.keywords, 'office-js, docx4j');
  assert.equal(core.created.content[0], '2026-09-15T10:30:00Z');
  const settings = await back.getMainDocumentPart().documentSettingsPart.getContents();
  assert.ok(settings.trackRevisions, 'w:trackRevisions was written');
});

test('Word.run: a loaded fixture, edited through the shim and saved', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('HelloWordOnline.docx'));
  await Word.run(pkg, async (context) => {
    const body = context.document.body;
    assert.equal(body.text, 'Hello World from Word Online');
    const first = body.paragraphs.getFirst();
    first.insertParagraph('Added by an add-in', 'After');
    const matches = body.search('Word Online', { matchCase: true });
    assert.equal(matches.length, 1);
    matches.getFirst().font.italic = true;
    const xml = body.getXml();
    await context.sync();
    assert.match(xml.value, /<w:t[^>]*>Word Online<\/w:t>/, 'the part was re-marshalled with the run split at the match');
  });
  const back = await WordprocessingMLPackage.load(await pkg.save());
  const body = await back.getBody();
  assert.deepEqual(body.paragraphs.map((p) => p.text), ['Hello World from Word Online', 'Added by an add-in']);
  assert.ok(body.paragraphs[0].runs.some((r) => r.value.rPr?.i), 'the match is italic');
});

test('collections: items, getFirst, getFirstOrNullObject, getLast and the null object', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  await Word.run(pkg, async (context) => {
    const body = context.document.body;
    assert.equal(body.paragraphs.length, 0);
    assert.throws(() => body.paragraphs.getFirst(), ItemNotFoundError);
    const none = body.paragraphs.getFirstOrNullObject();
    assert.equal(none.isNullObject, true);
    assert.equal(none.load(), none, 'load on a null object is still a no-op');
    assert.throws(() => none.text, (error) => error instanceof ItemNotFoundError && /null object/.test(error.message));

    body.insertParagraph('one', 'End');
    body.insertParagraph('two', 'End');
    const paragraphs = body.paragraphs;
    assert.equal(paragraphs.getFirst().text, 'one');
    assert.equal(paragraphs.getLast().text, 'two');
    assert.equal(paragraphs.getFirstOrNullObject().isNullObject, false);
    assert.equal(paragraphs.getLastOrNullObject().text, 'two');
    assert.deepEqual(paragraphs.items.map((p) => p.text), ['one', 'two']);
    assert.deepEqual([...paragraphs].map((p) => p.text), ['one', 'two'], 'a collection is still an array');
    assert.throws(() => paragraphs.getItemOrNullObject(0), NotSupportedError);
  });
});

test('proxies: an unsupported member names the class and the member', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  await Word.run(pkg, async (context) => {
    const paragraph = context.document.body.insertParagraph('x', 'End');
    assert.throws(() => paragraph.getNextOrNullObject(), (error) => {
      assert.ok(error instanceof NotSupportedError);
      assert.equal(error.message, 'Word.Paragraph.getNextOrNullObject is not supported by @docx4j/core-ts');
      assert.equal(error.code, 'NotImplemented');
      assert.equal(error.className, 'Paragraph');
      assert.equal(error.member, 'getNextOrNullObject');
      return true;
    });
    assert.throws(() => context.document.body.insertHtml('<p>x</p>', 'End'), /Word\.Body\.insertHtml is not supported/);
    assert.throws(() => { paragraph.listItem = 1; }, /Word\.Paragraph\.listItem is not supported/);
    assert.throws(() => context.application, /Word\.RequestContext\.application is not supported/);
    // getSelection has nothing to return in Node unless the caller supplies one
    assert.throws(() => context.document.getSelection(), /Word\.Document\.getSelection is not supported[\s\S]*selection/);
    // what is implemented passes through, and the tree is reachable through the extensions
    assert.equal(paragraph.text, 'x');
    assert.equal(unwrap(paragraph).constructor.name, 'Paragraph');
  });
});

test('Word.run: a supplied selection is what getSelection returns', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const paragraph = body.insertParagraph('select me', 'End');
  const selection = paragraph.search('me')[0];
  await Word.run(pkg, async (context) => {
    const range = context.document.getSelection();
    assert.equal(range.text, 'me');
    range.font.bold = true;
    await context.sync();
  }, { selection });
  assert.equal(paragraph.runs.filter((r) => r.value.rPr?.b).length, 1);
});

test('Word.supported lists what the subset declares, and nothing else', () => {
  assert.ok(Word.supported.has('Body.insertParagraph'));
  assert.ok(Word.supported.has('Paragraph.styleBuiltIn'));
  assert.ok(Word.supported.has('Font.highlightColor'));
  assert.ok(Word.supported.has('Range.insertText'));
  assert.ok(Word.supported.has('RequestContext.sync'));
  assert.ok(Word.supported.has('Document.getComments'));
  assert.ok(Word.supported.has('ParagraphCollection.getFirstOrNullObject'));
  assert.ok(Word.supported.has('Body.load'), 'load is a no-op on every object');
  assert.equal(Word.supported.has('Body.insertHtml'), false);
  assert.equal(Word.supported.has('Paragraph.getNextOrNullObject'), false);
  assert.equal(Word.supported.has('Body.insertContentControl'), false, 'CR-002 phase C/E');
});

test('document.getComments(): empty until CR-002 phase G adds it to Body', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  await Word.run(pkg, async (context) => {
    const comments = context.document.getComments();
    const items = comments instanceof ClientResult ? (await context.sync(), comments.value) : comments;
    assert.equal(items.length, 0);
  });
});

test('toApiScript: paragraphs, formatting and the insertXml fallback', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  body.insertParagraph('Hello World', 'End');
  const heading = body.insertParagraph('Chapter 1', 'End');
  heading.style = 'Heading1';
  heading.alignment = 'Centered';
  heading.spaceAfter = 6;
  heading.leftIndent = 18;
  await body.insertXml(
    '<w:p><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve"> plain </w:t></w:r>'
    + '<w:r><w:rPr><w:i/><w:iCs/><w:color w:val="FF0000"/></w:rPr><w:t>red</w:t></w:r></w:p>',
    'End',
  );

  assert.equal(await toApiScript(body.paragraphs[0]), "body.insertParagraph('Hello World', 'End');");
  assert.equal(await toApiScript(body.paragraphs[1]), [
    "const p1 = body.insertParagraph('Chapter 1', 'End');",
    "p1.styleBuiltIn = 'Heading1';",
    "p1.alignment = 'Centered';",
    'p1.leftIndent = 18;',
    'p1.spaceAfter = 6;',
  ].join('\n'));
  assert.equal(await toApiScript(body.paragraphs[2]), [
    "const p1 = body.insertParagraph('', 'End');",
    "const r1 = p1.insertText('bold', 'End');",
    'r1.font.bold = true;',
    "const r2 = p1.insertText(' plain ', 'End');",
    'r2.font.bold = false;',
    "const r3 = p1.insertText('red', 'End');",
    'r3.font.italic = true;',
    "r3.font.color = '#FF0000';",
  ].join('\n'));

  // a range emits the run formatting of the span only
  assert.equal(await toApiScript(body.paragraphs[2].search('red')[0]), [
    "const p1 = body.insertParagraph('', 'End');",
    "const r1 = p1.insertText('red', 'End');",
    'r1.font.italic = true;',
    "r1.font.color = '#FF0000';",
  ].join('\n'));

  // what no verb expresses falls back to insertXml with the marshalled fragment
  await body.insertXml('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>list item</w:t></w:r></w:p>', 'End');
  const fallback = await toApiScript(body.paragraphs[3]);
  assert.match(fallback, /^\/\/ paragraph properties \(numPr\): as XML\nawait body\.insertXml\(`<w:p /);
  assert.match(fallback, /w:numId w:val="1"/);

  // the variable is the caller's
  assert.equal(await toApiScript(body.paragraphs[0], { variable: 'cell' }), "cell.insertParagraph('Hello World', 'End');");
});

test('toApiScript: the emitted script reproduces the content when run through Word.run', async () => {
  const source = await WordprocessingMLPackage.createPackage();
  const body = await source.getBody();
  const heading = body.insertParagraph('Chapter 1', 'End');
  heading.style = 'Heading1';
  heading.alignment = 'Centered';
  await body.insertXml(
    '<w:p><w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>big bold</w:t></w:r>'
    + '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve"> underlined</w:t></w:r></w:p>',
    'End',
  );
  body.insertParagraph('plain tail', 'End');
  const script = await toApiScript(body);

  const target = await WordprocessingMLPackage.createPackage();
  await Word.run(target, async (context) => {
    await asyncScript(script)(context.document.body);
    await context.sync();
  });
  const back = await WordprocessingMLPackage.load(await target.save());
  const reloaded = await back.getBody();
  assert.equal(reloaded.text, body.text);
  assert.deepEqual(reloaded.paragraphs.map((p) => [p.style, p.alignment]), body.paragraphs.map((p) => [p.style, p.alignment]));
  const formatting = (b) => b.paragraphs[1].runs.map((r) => {
    const font = r.value.rPr ?? {};
    return [Boolean(font.b), font.sz?.val ?? 0, font.u?.val ?? 'none'];
  });
  assert.deepEqual(formatting(reloaded), formatting(body));
});

test("Word.Style agrees with the content API's built-in list", async () => {
  const { BUILT_IN_STYLES } = await import('../dist/index.mjs');
  const enumValues = Object.values(Word.Style).filter((v) => v !== 'Other').sort();
  assert.deepEqual(enumValues, [...BUILT_IN_STYLES].sort());
});
