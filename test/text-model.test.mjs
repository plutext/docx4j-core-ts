// CR-012 phase 2: the text model (src/model/text). docx4j's TextsTest, ported where this package
// implements the statement (REQ-203's strictAsTransitional is not: CR-001 section 17.6), and this
// package's own: the package saved byte for byte after a reading, TextRun.officeOffset against
// Paragraph.text, the part overloads under one TextContext, and the block model's JSON.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import {
  WordprocessingMLPackage, Texts, PlainText, TextOptions, TextOptionError, TextResult, TextContext, TEXT_OPTION_NAMES,
  ImageDescription, Paragraph, XmlPart,
} from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const phase0 = join(here, 'text-extraction', 'suite', 'phase0');
const PROBE = readFileSync(join(phase0, 'extraction-probe.docx'));
const probe = await WordprocessingMLPackage.load(PROBE);

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS = `xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`
  + ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
  + ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
  + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  + ' xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'
  + ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const DEL = ' w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** docx4j's MiniDocx: a docx made from raw XML, loaded through load. */
function miniDocx(body) {
  const parts = { 'word/document.xml': `${XML}<w:document ${NS}><w:body>${body}</w:body></w:document>` };
  let types = '';
  let rels = '';
  let packageRels = '';
  const api = {
    relationship(rId, relType, target) {
      rels += `<Relationship Id="${rId}" Type="${relType}" Target="${target}"/>`;
      return api;
    },
    entry(path, contentType, xml) {
      parts[path] = xml;
      if (contentType) types += `<Override PartName="/${path}" ContentType="${contentType}"/>`;
      return api;
    },
    packagePart(path, contentType, relType, rId, xml) {
      parts[path] = xml;
      types += `<Override PartName="/${path}" ContentType="${contentType}"/>`;
      packageRels += `<Relationship Id="${rId}" Type="${relType}" Target="${path}"/>`;
      return api;
    },
    part(name, contentType, relType, rId, xml) {
      parts[`word/${name}`] = xml;
      types += `<Override PartName="/word/${name}" ContentType="${contentType}"/>`;
      rels += `<Relationship Id="${rId}" Type="${relType}" Target="${name}"/>`;
      return api;
    },
    footnotes(notes) {
      return api.part('footnotes.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes', 'rIdFn', `${XML}<w:footnotes ${NS}>${notes}</w:footnotes>`);
    },
    bytes() {
      const entries = {
        '[Content_Types].xml': strToU8(`${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
          + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          + `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${types}</Types>`),
        '_rels/.rels': strToU8(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
          + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>${packageRels}</Relationships>`),
        'word/_rels/document.xml.rels': strToU8(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`),
      };
      for (const [k, v] of Object.entries(parts)) entries[k] = strToU8(v);
      return zipSync(entries);
    },
    load: () => WordprocessingMLPackage.load(api.bytes()),
  };
  return api;
}

const text = async (pkg, options) => PlainText.of(await Texts.of(pkg, options));
const blockWith = (result, token) => {
  const b = result.blocks.find((block) => block.text.includes(token));
  assert.ok(b, `no block holds ${token}`);
  return b;
};

// ---------------------------------------------------------------- docx4j's TextsTest

test('REQ-401: three views, accepted the default; current a synonym', () => {
  assert.equal(new TextOptions().view, 'accepted');
  assert.equal(TextOptions.of({ view: 'current' }).view, 'accepted');
  assert.equal(TextOptions.of({ view: 'markup' }).view, 'markup');
  assert.throws(() => TextOptions.of({ view: 'final' }), (e) => e instanceof TextOptionError && e.message.includes('REQ-401'));
});

test('REQ-204: no view or scope changes the package, and a save afterwards is the save before, byte for byte', async () => {
  const before = unzipSync(await (await WordprocessingMLPackage.load(PROBE)).save());
  const pkg = await WordprocessingMLPackage.load(PROBE);
  for (const view of ['accepted', 'original', 'markup']) {
    await Texts.of(pkg, { view, include: ['headers', 'footers', 'glossary'], comments: true, hidden: true, field_codes: true, list_labels: true });
  }
  // nothing was unmarshalled: every part was read privately
  for (const part of pkg.parts) if (part instanceof XmlPart && !part.partName.name.endsWith('.rels')) assert.equal(part.isUnmarshalled, false, part.partName.name);
  const after = unzipSync(await pkg.save());
  assert.deepEqual(Object.keys(after).sort(), Object.keys(before).sort());
  for (const name of Object.keys(before)) assert.deepEqual(after[name], before[name], name);
});

test('REQ-204: a package with no styles part is not given one', async () => {
  const pkg = await miniDocx('<w:p><w:r><w:t>TEXT</w:t></w:r></w:p>').load();
  const parts = [...pkg.parts].length;
  for (const view of ['accepted', 'original', 'markup']) assert.equal(await text(pkg, { view, list_labels: true }), 'TEXT\n');
  assert.equal(pkg.getMainDocumentPart().styleDefinitionsPart, undefined);
  assert.equal([...pkg.parts].length, parts);
});

test('REQ-402, REQ-405: a deleted mark joins in accepted, under the first paragraph\'s address and paraId; two blocks otherwise', async () => {
  const pkg = await miniDocx(`<w:p w14:paraId="0000000A"><w:pPr><w:rPr><w:del${DEL}/></w:rPr></w:pPr><w:r><w:t>FIRST</w:t></w:r></w:p>`
    + '<w:p w14:paraId="0000000B"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>SECOND</w:t></w:r></w:p>').load();
  const accepted = await Texts.of(pkg);
  assert.equal(accepted.blocks.length, 1);
  const joined = accepted.blocks[0];
  assert.equal(joined.text, 'FIRSTSECOND');
  assert.equal(joined.address, 'body/0');
  assert.equal(joined.paraId, '0000000A');
  assert.equal(joined.style, 'Heading1');                             // the joined paragraph keeps the second one's properties
  assert.equal(joined.revisions, 1);
  assert.equal(joined.source.paraId, '0000000A');                    // the first paragraph's w:p, as its paraId (CR-012 section 4.4)
  assert.equal(await text(pkg, { view: 'original' }), 'FIRST\nSECOND\n');
  assert.equal(await text(pkg, { view: 'markup' }), 'FIRST\nSECOND\n'); // REQ-403: no structure resolved, no marker
});

test('REQ-601, REQ-602, REQ-603: run content mapped, nothing normalised or trimmed; a line per paragraph and break', async () => {
  const pkg = await miniDocx('<w:p><w:r><w:t xml:space="preserve">  lead</w:t><w:tab/><w:t>A</w:t><w:br/><w:t>B</w:t>'
    + '<w:br w:type="page"/><w:t>C</w:t><w:cr/><w:t>D</w:t><w:noBreakHyphen/><w:softHyphen/><w:sym w:font="Wingdings" w:char="F0FC"/>'
    + '<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/><w:t xml:space="preserve">é  trail  </w:t></w:r></w:p>'
    + '<w:p/><w:p><w:r><w:t>last</w:t></w:r></w:p>').load();
  assert.equal(await text(pkg), '  lead\tA\nB\nC\nD‑­\té  trail  \n\nlast\n');
});

test('REQ-505: codes excluded unless asked for; a field with no result contributes nothing, is a field with no result, and warns', async () => {
  const xml = '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    + '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    + '<w:r><w:t>RESULT</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    + '<w:p><w:fldSimple w:instr=" AUTHOR "/></w:p>';
  const pkg = await miniDocx(xml).load();
  const result = await Texts.of(pkg);
  assert.equal(PlainText.of(result), 'Page \nRESULT\n\n');
  const page = result.blocks[0].fields[0];
  assert.equal(page.instr, ' PAGE ');
  assert.equal(page.result, undefined);
  assert.equal(result.blocks[1].fields[0].result, 'RESULT');
  assert.equal(result.blocks[2].fields[0].result, undefined);
  assert.deepEqual(result.warnings.map(String), ['field.no_result at body/0', 'field.no_result at body/2']);
  assert.equal(await text(pkg, { field_codes: true }), 'Page  PAGE \n QUOTE x RESULT\n AUTHOR \n');
});

const NOTES_BODY = '<w:p><w:r><w:t>A</w:t></w:r><w:r><w:footnoteReference w:id="7"/></w:r>'
  + '<w:r><w:t xml:space="preserve"> B</w:t></w:r><w:r><w:footnoteReference w:customMarkFollows="1" w:id="8"/></w:r><w:r><w:t>*</w:t></w:r>'
  + '<w:r><w:t xml:space="preserve"> C</w:t></w:r><w:r><w:footnoteReference w:id="9"/></w:r></w:p>';
const NOTES = '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
  + '<w:footnote w:id="9"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> NOTE_NINE</w:t></w:r></w:p></w:footnote>'
  + '<w:footnote w:id="7"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> NOTE_SEVEN</w:t></w:r></w:p></w:footnote>'
  + '<w:footnote w:id="8"><w:p><w:r><w:t>* NOTE_EIGHT</w:t></w:r></w:p></w:footnote>';

test('REQ-503, REQ-707: [n] in reference order; a custom mark gets no number but its note is read; note_references: false drops the inline marker only', async () => {
  const pkg = await miniDocx(NOTES_BODY).footnotes(NOTES).load();
  assert.equal(await text(pkg), 'A[1] B* C[2]\n[1] NOTE_SEVEN\n* NOTE_EIGHT\n[2] NOTE_NINE\n');
  assert.equal(await text(pkg, { note_references: false }), 'A B* C\n[1] NOTE_SEVEN\n* NOTE_EIGHT\n[2] NOTE_NINE\n');
  assert.equal(await text(pkg, { notes: false }), 'A B* C\n');
  const note = blockWith(await Texts.of(pkg), 'NOTE_NINE');
  assert.equal(note.story, 'FOOTNOTE');
  assert.equal(note.noteNumber, 2);
  assert.equal(note.address, 'footnote:9/0');
  assert.ok(!note.text.includes('['));                                // the number is a field, not text (REQ-707)
});

test('REQ-501, REQ-606: the main story, then footnotes, then endnotes; headers, footers and comments only when included', async () => {
  const t = await text(probe);
  const lastBody = t.indexOf('Case 54:');
  const footnote = t.indexOf('[1] TC17_NOTE_BASE');
  const endnote = t.indexOf('[1] EN27_NOTE_TEXT');
  assert.ok(lastBody >= 0 && footnote > lastBody && endnote > footnote);
  assert.ok(!t.includes('HDR26_') && !t.includes('FTR26_') && !t.includes('{>>'));
});

test('REQ-502, REQ-519: per section, the headers Word displays before its body and the footers after; each once; no orphan', async () => {
  const options = { include: ['headers', 'footers'] };
  const t = await text(probe, options);
  const first = t.indexOf('HDR26_FIRST'), dflt = t.indexOf('HDR26_DEFAULT'), even = t.indexOf('HDR26_EVEN');
  const body = t.indexOf('Text extraction probe'), endOfSection1 = t.indexOf('The last paragraph of section 1.');
  const footer = t.indexOf('FTR26_DEFAULT'), section2 = t.indexOf('HDR26_SECTION2'), body2 = t.indexOf('The first paragraph of section 2.');
  assert.ok(first >= 0 && first < dflt && dflt < even && even < body);
  assert.ok(endOfSection1 < footer && footer < section2 && section2 < body2);
  assert.equal(t.lastIndexOf('HDR26_EVEN'), even);
  assert.ok(!t.includes('HDR26_ORPHAN'));
  const stories = (await Texts.of(probe, options)).blocks.filter((b) => b.story === 'HEADER' || b.story === 'FOOTER');
  assert.ok(stories.length >= 5);
  for (const b of stories) assert.match(b.address, /^(header|footer):rId\d+\/\d+$/);
});

test('REQ-504: comments in markup at their reference, replies after them, resolved flagged; otherwise only under comments: true', async () => {
  const markup = await Texts.of(probe, { view: 'markup' });
  assert.ok(PlainText.of(markup).includes('CMT28_ANCHOR{>>CMT28_COMMENT_TEXT<<}{>>CMT28_REPLY<<} and CMT28_RESOLVED_ANCHOR{>>CMT28_RESOLVED_COMMENT<<}'));
  const comments = [...markup.comments.values()];
  assert.equal(comments.find((c) => c.text.includes('CMT28_RESOLVED_COMMENT')).resolved, true);
  assert.equal(comments.find((c) => c.text.includes('CMT28_REPLY')).parentId, '0');
  assert.ok(!(await text(probe)).includes('{>>'));
  assert.ok((await text(probe, { comments: true })).includes('{>>CMT28_COMMENT_TEXT<<}'));
});

test('REQ-507: hidden text and the hidden paragraph included and flagged under hidden: true', async () => {
  const result = await Texts.of(probe, { hidden: true });
  assert.equal(blockWith(result, 'HID18_HIDDEN_PARA').hidden, true);
  assert.equal(blockWith(result, 'HID18_VANISH').runs.find((r) => r.text.includes('HID18_VANISH')).hidden, true);
  assert.ok(!(await text(probe)).includes('HID18_VANISH'));
});

test('REQ-508: placeholder text included and flagged; placeholders: false leaves it out', async () => {
  assert.equal(blockWith(await Texts.of(probe), 'PLC22_PLACEHOLDER').placeholder, true);
  assert.ok(!(await text(probe, { placeholders: false })).includes('PLC22_PLACEHOLDER'));
});

test('REQ-510: no list labels by default; the numbering model\'s under list_labels: true', async () => {
  assert.ok((await text(probe)).includes('\nLST29_ITEM_A\n'));
  assert.ok((await text(probe, { list_labels: true })).includes('\n1.\tLST29_ITEM_A\n2.\tLST29_ITEM_B\n3.\tLST29_ITEM_C\n'));
});

test('REQ-518, REQ-519: a box\'s blocks after its anchor, addressed through it; after the table for a box in a cell', async () => {
  const result = await Texts.of(probe);
  const box = blockWith(result, 'TBX25_FLOAT');
  assert.equal(box.story, 'TEXT_BOX');
  assert.equal(box.textbox, true);
  assert.match(box.address, /^body\/\d+\/textbox:0\/0$/);
  const at = result.blocks.indexOf(box);
  assert.equal(result.blocks[at - 1].address, box.address.replace(/\/textbox:.*/, ''));
  const inCell = blockWith(result, 'TBX25_IN_CELL');
  const before = result.blocks[result.blocks.indexOf(inCell) - 1];
  assert.equal(before.kind, 'TABLE');
  assert.ok(inCell.address.startsWith(`${before.address}/`));
});

test('REQ-702, REQ-706: CriticMarkup, the detail in the marker\'s opening; runs carry their revision', async () => {
  const type = await text(probe, { view: 'markup' });
  assert.ok(type.includes('TC9_BASE {++TC9_INSERTED ++}{--TC9_DELETED--}'));
  assert.ok(!type.includes('{~~'));
  assert.ok((await text(probe, { view: 'markup', markup_detail: 'author' })).includes('{++Probe: TC9_INSERTED ++}{--Probe: TC9_DELETED--}'));
  assert.ok((await text(probe, { view: 'markup', markup_detail: 'author_date' })).includes('{++Probe, 2026-10-17T00:00:00Z: TC9_INSERTED ++}'));
  const block = blockWith(await Texts.of(probe, { view: 'markup' }), 'TC9_INSERTED');
  const inserted = block.runs.find((r) => r.text.includes('TC9_INSERTED')).revision;
  assert.equal(inserted.type, 'INSERTION');
  assert.equal(inserted.author, 'Probe');
  assert.equal(inserted.date, '2026-10-17T00:00:00Z');
  assert.equal(block.revisions, 2);
  for (const r of blockWith(await Texts.of(probe), 'TC9_INSERTED').runs) assert.equal(r.revision, undefined);
});

test('REQ-801, REQ-802: warnings are metadata, with no document text', async () => {
  const result = await Texts.of(probe, { view: 'markup' });
  assert.ok(result.warnings.length > 0);
  for (const w of result.warnings) assert.doesNotMatch(w.toString(), /[A-Z]+[0-9]+_[A-Z_0-9]+/);
  assert.equal(result.metadata.fields_update_on_open, false);
});

test('REQ-1001: every WD option name and value accepted; anything else refused; what is not implemented refused, naming its statement', async () => {
  const defaults = new TextOptions();
  assert.deepEqual(Object.keys(defaults.toMap()), [...TEXT_OPTION_NAMES]);
  assert.deepEqual(TextOptions.of(defaults.toMap()).toMap(), defaults.toMap());
  const o = TextOptions.of({ view: 'current', include: 'headers, footers comments', hidden: 'true', markup_detail: 'author_date', tables: 'tsv' });
  assert.equal(o.view, 'accepted');
  assert.ok(o.comments && o.hidden);
  assert.equal(o.markupDetail, 'author_date');
  for (const [name, value] of [['colour', 'red'], ['hidden', 'maybe'], ['fields', 'recalculate'], ['include', 'sidebars']]) {
    assert.throws(() => new TextOptions().set(name, value), (e) => e instanceof TextOptionError && e.message.includes('REQ-1001'), name);
  }
  // implemented in phase 2
  for (const [name, value] of [['metadata', 'true'], ['alt_text', 'true'], ['include', 'glossary'], ['tables', 'grid'], ['page_markers', 'true'], ['layout', 'indent']]) {
    await Texts.of(probe, new TextOptions().set(name, value));
  }
  // not implemented: never (docx4j's two), and not yet (CR-012 phase 5)
  for (const [name, value, statement, reason] of [['math', 'unicodemath', 'REQ-604', 'never'], ['symbols', 'unicode', 'REQ-601', 'never'],
    ['fields', 'update', 'REQ-506', 'phase 5a'], ['embedded', 'true', 'REQ-512', 'phase 5d']]) {
    await assert.rejects(Texts.of(probe, new TextOptions().set(name, value)), (e) => e instanceof TextOptionError && e.reason === 'not_implemented'
      && e.statement === statement && e.message.includes('REQ-1001') && e.message.includes(`${name}: ${value}`), `${name} (${reason})`);
  }
});

// ---------------------------------------------------------------- this package's own

/** The characters Office JS's text may hold where the model writes `c` (check 48; CR-012 phase 4). */
function officeCharacterOf(c) {
  switch (c) {
    case '\n': return '\u000B\u000C\u000E\r';
    case '\u2011': return '\u001E';
    case '\u00AD': return '\u001F';
    default: return '(';                                               // a symbol, whatever its code
  }
}

test('CR-012 section 4.4: a run\'s officeOffset is where its text lies in its paragraph\'s Paragraph.text', async () => {
  const pkg = await WordprocessingMLPackage.load(PROBE);
  await pkg.getBody();
  let mapped = 0;
  let unmapped = 0;
  for (const view of ['accepted', 'markup']) {
    for (const block of (await Texts.of(pkg, { view, hidden: true, list_labels: true, field_codes: true })).blocks) {
      for (const run of block.runs) {
        if (run.officeOffset === undefined) { unmapped++; continue; }
        const source = run.source;
        assert.ok(source, `${block.address}: a run with an officeOffset has a source`);
        const office = new Paragraph({ name: { namespaceURI: W, localPart: 'p' }, value: source.paragraph }, [], pkg.body).text;
        // character for character, in Office JS's characters where the model writes its own (check 48: a
        // break \n there, U+000B to U+000E here; the hyphens; any symbol `(`)
        const span = office.slice(run.officeOffset, run.officeOffset + run.text.length);
        assert.equal(span.length, run.text.length, `${view} ${block.address}: '${run.text}'`);
        for (let i = 0; i < span.length; i++) {
          assert.ok(span[i] === run.text[i] || officeCharacterOf(run.text[i]).includes(span[i]), `${view} ${block.address}: '${run.text}' at ${i}: ${JSON.stringify(span[i])}`);
        }
        mapped++;
      }
    }
  }
  assert.ok(mapped > 300, `${mapped} runs mapped`);
  assert.ok(unmapped > 0);
  // what only the model writes has none: a note reference, a field code, a list label
  const model = await Texts.of(pkg, { field_codes: true, list_labels: true });
  for (const block of model.blocks) for (const run of block.runs) if (run.kind !== 'TEXT') assert.equal(run.officeOffset, undefined, `${run.kind} at ${block.address}`);
});

test('CR-012 section 4.4 and phase 4: hidden text, which Paragraph.text leaves out as Office JS does, has no officeOffset; a note reference is U+0002 there', async () => {
  const pkg = await miniDocx('<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN</w:t></w:r>'
    + '<w:r><w:t xml:space="preserve"> after</w:t></w:r><w:r><w:footnoteReference w:id="7"/></w:r><w:r><w:t>!</w:t></w:r></w:p>').footnotes(NOTES).load();
  const block = (await Texts.of(pkg)).blocks[0];
  assert.deepEqual(block.runs.map((r) => [r.kind, r.text, r.officeOffset]), [
    ['TEXT', 'before ', 0], ['TEXT', ' after', 7], ['NOTE_REFERENCE', '', undefined], ['TEXT', '!', 14],
  ]);
  const shown = (await Texts.of(pkg, { hidden: true })).blocks[0];
  assert.deepEqual(shown.runs.map((r) => [r.text, r.officeOffset]), [['before ', 0], ['HIDDEN', undefined], [' after', 7], ['', undefined], ['!', 14]]);
  await pkg.getBody();
  assert.equal(pkg.body.paragraphs[0].text, 'before  after\u0002!');
});

test('CR-012 section 4.4: the part overloads under one TextContext, concatenated, are the package\'s text', async () => {
  const pkg = await WordprocessingMLPackage.load(PROBE);
  const main = pkg.getMainDocumentPart();
  for (const view of ['accepted', 'original', 'markup']) {
    const whole = await Texts.of(pkg, { view });
    const context = await Texts.contextOf(pkg);
    assert.ok(context instanceof TextContext);
    const parts = [
      await Texts.of(main, { view }, context),
      await Texts.of(main.footnotesPart, { view }, context),
      await Texts.of(main.endnotesPart, { view }, context),
      ...(view === 'markup' ? [await Texts.of(main.commentsPart, { view }, context)] : []),
    ];
    assert.equal(PlainText.of(TextResult.concat(...parts)), PlainText.of(whole), view);
  }
  // a context is one view's: another is refused
  const context = await Texts.contextOf(pkg);
  await Texts.of(main, { view: 'accepted' }, context);
  await assert.rejects(Texts.of(main, { view: 'markup' }, context), (e) => e instanceof TextOptionError && e.message.includes('accepted view'));
  await assert.rejects(Texts.of(await WordprocessingMLPackage.load(PROBE), {}, context), /another package/);
  // a notes part read first, with no context, numbers its notes by the main story's references
  const alone = await Texts.of(main.footnotesPart);
  assert.deepEqual(alone.blocks.map((b) => b.text), (await Texts.of(pkg)).blocks.filter((b) => b.story === 'FOOTNOTE').map((b) => b.text));
  // a header, as its own story
  const header = main.headerParts[0];
  const own = await Texts.of(header);
  assert.ok(own.blocks.length > 0);
  for (const b of own.blocks) assert.match(b.address, /^header:rId\d+\/\d+(\/textbox:\d+\/\d+)?$/);
});

test('CR-012 section 4.4: content in a package reads in the package\'s view, addressed from content', async () => {
  const pkg = await miniDocx(`<w:p><w:pPr><w:rPr><w:del${DEL}/></w:rPr></w:pPr><w:r><w:t>FIRST</w:t></w:r></w:p>`
    + '<w:p><w:r><w:t>SECOND</w:t></w:r></w:p>').load();
  // content in the package's tree: the part unmarshalled, so that its tree is the one every reading walks
  const document = await pkg.getMainDocumentPart().getContents();
  const first = document.body.content[0];
  const result = await Texts.of(pkg, first);
  assert.equal(result.blocks.length, 1);
  assert.equal(result.blocks[0].address, 'content/0');
  assert.equal(result.blocks[0].text, 'FIRST');                      // one paragraph given: no next to join
  const body = await Texts.of(pkg, document.body);
  assert.deepEqual(body.blocks.map((b) => [b.address, b.text]), [['content/0', 'FIRSTSECOND']]);
  // with no package, run-level revisions only
  const planless = await Texts.of(document.body, { view: 'original' });
  assert.deepEqual(planless.blocks.map((b) => b.text), ['FIRST', 'SECOND']);
});

test('REQ-706: the block model\'s JSON holds the blocks\' fields, text and flags, and no tree object', async () => {
  const result = await Texts.of(probe, { view: 'markup', hidden: true });
  const json = JSON.parse(JSON.stringify(result));
  assert.equal(json.blocks.length, result.blocks.length);
  const first = json.blocks.find((b) => b.kind === 'PARAGRAPH' && b.text !== '');
  for (const key of ['kind', 'story', 'address', 'text', 'runs', 'placeholder', 'textbox', 'decorative', 'revisions']) assert.ok(key in first, key);
  const flat = JSON.stringify(result);
  assert.ok(!flat.includes('TYPE_NAME'), 'no tree object');
  assert.ok(!flat.includes('"source"'));
  assert.ok(blockWith(result, 'TC9_INSERTED').source !== undefined, 'the source is there, out of the JSON');
});

test('REQ-514: a describer\'s description, its provenance written; one that fails raises image.not_described', async () => {
  const generated = await Texts.of(probe, { alt_text: true, imageDescriber: (image) => (image.decorative ? undefined : ImageDescription.generated('A PICTURE')) });
  assert.ok(PlainText.of(generated).includes('[image, generated: A PICTURE]'));
  const failing = await Texts.of(probe, { imageDescriber: () => { throw new Error('no model'); } });
  assert.ok(failing.warnings.some((w) => w.kind === 'image.not_described' && w.partName?.startsWith('/word/media/')));
});

// ---------------------------------------------------------------- docx4j's TextsTocTest (REQ-520)

const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const field = (instr, result) => '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
  + `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}`
  + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const entry = (bookmark, title, page) => `<w:hyperlink w:anchor="${bookmark}">${run(title)}<w:r><w:tab/></w:r>${field(`PAGEREF ${bookmark} \\h`, run(page))}</w:hyperlink>`;
/** Word's shape: a heading, the field begun in the first entry's paragraph, ended in a paragraph of its own. */
const TOC = '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>'
  + `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>${run('Contents')}</w:p>`
  + '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
  + `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${entry('_Toc1', 'Introduction', '1')}</w:p>`
  + `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>${entry('_Toc2', 'Method', '2')}</w:p>`
  + '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt>';
const TOC_BODY = `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="1" w:name="_Toc1"/>${run('Introduction')}<w:bookmarkEnd w:id="1"/></w:p>`
  + `<w:p>${run('Dated ')}${field('DATE', run('today'))}</w:p>`;

test('REQ-520: a table of contents is left out by default, its heading kept, no warning and no field recorded', async () => {
  const pkg = await miniDocx(TOC + TOC_BODY).load();
  assert.equal(new TextOptions().toc, false);
  const r = await Texts.of(pkg);
  assert.equal(PlainText.of(r), 'Contents\nIntroduction\nDated today\n');
  assert.deepEqual(r.warnings, []);
  for (const b of r.blocks) for (const f of b.fields) assert.ok(!f.instr.includes('TOC') && !f.instr.includes('PAGEREF'), f.instr);
});

test('REQ-520: under toc: true it is read as any field', async () => {
  const pkg = await miniDocx(TOC + TOC_BODY).load();
  assert.equal(await text(pkg, { toc: true }), 'Contents\nIntroduction\t1\nMethod\t2\n\nIntroduction\nDated today\n');
  assert.equal(await text(pkg, { toc: 'true' }), await text(pkg, { toc: true }));
  assert.equal(new TextOptions().toMap().toc, false);
});

test('REQ-520: under field_codes nothing of the TOC field is read, its code included; other fields\' codes are', async () => {
  const pkg = await miniDocx(TOC + TOC_BODY).load();
  assert.equal(await text(pkg, { field_codes: true }), 'Contents\nIntroduction\nDated  DATE today\n');
  const kept = await text(pkg, { field_codes: true, toc: true });
  assert.ok(kept.includes(' TOC \\o "1-3" \\h \\z \\u Introduction\t PAGEREF _Toc1 \\h 1'), kept);
});

test('REQ-520: a table of figures is a TOC field too; a simple field\'s is left out as a complex one\'s', async () => {
  const pkg = await miniDocx(`<w:p>${run('Figures')}</w:p>`
    + `<w:p><w:fldSimple w:instr=" TOC \\h \\z \\c &quot;Figure&quot; ">${run('Figure 1: A')}</w:fldSimple></w:p>`
    + `<w:p>${field('TOC \\c "Table"', run('Table 1: B'))}</w:p><w:p>${run('After')}</w:p>`).load();
  assert.equal(await text(pkg), 'Figures\nAfter\n');
  assert.equal(await text(pkg, { toc: true }), 'Figures\nFigure 1: A\nTable 1: B\nAfter\n');
});

test('REQ-520: a table of contents whose field was unlinked is the document\'s text', async () => {
  const pkg = await miniDocx(`<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>${run('Introduction\t1')}</w:p>`).load();
  assert.equal(await text(pkg), 'Introduction\t1\n');
});

test('REQ-520: in every view, text beside the field stays, and so does a revision outside it; two runs either side, as Paragraph.text holds the field between them', async () => {
  const pkg = await miniDocx(`<w:p>${run('See: ')}${field('TOC \\o', run('Entry\t3'))}${run(' (end)')}</w:p>`
    + `<w:p>${field('TOC \\o', '<w:ins w:id="1" w:author="A"><w:r><w:t>Added entry</w:t></w:r></w:ins>')}</w:p><w:p>${run('Body')}</w:p>`).load();
  for (const view of ['accepted', 'original', 'markup']) assert.equal(await text(pkg, { view }), 'See:  (end)\nBody\n', view);
  // docx4j makes them one run; here a run's text is one stretch of Paragraph.text, which holds the table of
  // contents' result between them (CR-012 section 15), so they are two, with the same text
  for (const options of [{}, { field_codes: true }]) {
    const first = (await Texts.of(pkg, options)).blocks[0];
    assert.deepEqual(first.runs.map((r) => [r.text, r.officeOffset]), [['See: ', 0], [' (end)', 12]], JSON.stringify(options));
    assert.equal(first.text, 'See:  (end)');
  }
});

// ---------------------------------------------------------------- docx4j's TextsPhase2Test, where it applies

const EXTERNAL = await WordprocessingMLPackage.load(readFileSync(join(phase0, 'extraction-probe-external.docm')));

/** The text from `from` up to `to`, `from` included. */
function between(t, from, to) {
  const i = t.indexOf(from);
  assert.ok(i >= 0, `no '${from}'`);
  const j = t.indexOf(to, i + from.length);
  assert.ok(j >= 0, `no '${to}'`);
  return t.slice(i, j);
}

const p = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const tc = (props, ...paragraphs) => `<w:tc><w:tcPr>${props}</w:tcPr>${paragraphs.map((s) => (s === '' ? '<w:p/>' : p(s))).join('')}</w:tc>`;
const grid = (...widths) => `<w:tbl><w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
const SPAN2 = '<w:gridSpan w:val="2"/>';
const SPAN3 = '<w:gridSpan w:val="3"/>';
const HEADER = '<w:trPr><w:tblHeader/></w:trPr>';

test('REQ-605: case 31 in the listed format, as TSV under tables: tsv, as a grid under tables: grid; a simple table tab-separated', async () => {
  const listed = await text(probe);
  assert.equal(between(listed, '[table: listed format]\nTBL31', 'After the table.\nCase 32'),
    '[table: listed format]\nTBL31_HEADER\nTBL31_CELL_P1; TBL31_CELL_P2\tTBL31_NESTED\nTBL31_MERGED\tTBL31_LAST_CELL\n');
  // a block-level control in a cell, with Word's empty paragraph after it, is a simple table's cell
  assert.equal(between(listed, 'SDT24_IN_CELL', 'After the table.\nCase 25'), 'SDT24_IN_CELL\n');
  assert.equal(between(await text(probe, { tables: 'tsv' }), 'TBL31_HEADER', 'After the table.\nCase 32'),
    'TBL31_HEADER\nTBL31_CELL_P1 TBL31_CELL_P2\tTBL31_NESTED\t\nTBL31_MERGED\tTBL31_LAST_CELL\n');
  assert.equal(between(await text(probe, { tables: 'grid' }), 'Case 31: a table', 'After the table.\nCase 32'),
    'Case 31: a table: header row, two-paragraph cell, nested table, merged cell, empty cell\n'
    + '+---------------+------------------+-----------------+\n'
    + '| TBL31_HEADER                                       |\n'
    + '+===============+==================+=================+\n'
    + '| TBL31_CELL_P1 | +--------------+ |                 |\n'
    + '|               | | TBL31_NESTED | |                 |\n'
    + '| TBL31_CELL_P2 | +--------------+ |                 |\n'
    + '+---------------+------------------+-----------------+\n'
    + '| TBL31_MERGED                     | TBL31_LAST_CELL |\n'
    + '+---------------+------------------+-----------------+\n');
});

test('REQ-605: headers from the first row with no w:tblHeader; a vertical merge\'s value on each row; an empty cell left out', async () => {
  const merged = `${grid(2000, 2000)}<w:tr>${tc('', 'Name')}${tc('', 'Value')}</w:tr>`
    + `<w:tr>${tc('<w:vMerge w:val="restart"/>', 'a')}${tc('', '1')}</w:tr><w:tr>${tc('<w:vMerge/>', '')}${tc('', '2')}</w:tr>`
    + `<w:tr>${tc('', '')}${tc('', '3')}</w:tr></w:tbl>${p('After.')}`;
  const pkg = await miniDocx(merged).load();
  assert.equal(await text(pkg), '[table: listed format]\nName: a; Value: 1\nName: a; Value: 2\nValue: 3\nAfter.\n');
  assert.equal(await text(pkg, { tables: 'grid' }), '+------+-------+\n| Name | Value |\n+======+=======+\n| a    | 1     |\n'
    + '|      +-------+\n|      | 2     |\n+------+-------+\n|      | 3     |\n+------+-------+\nAfter.\n');
  const headed = await miniDocx(`${grid(2000, 2000)}<w:tr>${HEADER}${tc('', 'H1')}${tc('', 'H2')}</w:tr><w:tr>${tc(SPAN2, 'wide')}</w:tr></w:tbl>`
    + `${grid(2000, 2000)}<w:tr>${tc(SPAN2, 'only')}</w:tr><w:tr>${tc('', 'x')}${tc('', 'y')}</w:tr></w:tbl>`).load();
  assert.equal(await text(headed), '[table: listed format]\nwide\n[table: listed format]\nonly\nx\ty\n');
});

test('REQ-605: the listed format\'s headers, as ruled 18 October 2026', async () => {
  const rows = `<w:tr>${tc('', 'P1', 'P2')}${tc('', 'N')}${tc('', '')}</w:tr><w:tr>${tc(SPAN2, 'M')}${tc('', 'L')}</w:tr></w:tbl>`;
  const t = async (body) => text(await miniDocx(body).load());
  assert.equal(await t(`${grid(3000, 3000, 3000)}<w:tr>${HEADER}${tc('', 'A')}${tc('', 'B')}${tc('', 'C')}</w:tr>${rows}`),
    '[table: listed format]\nA: P1; P2; B: N\nA: M; C: L\n');
  assert.equal(await t(`${grid(3000, 3000, 3000)}<w:tr>${HEADER}${tc(SPAN2, 'AB')}${tc('', 'C')}</w:tr>${rows}`),
    '[table: listed format]\nAB: P1; P2; AB: N\nAB: M; C: L\n');
  assert.equal(await t(`${grid(3000, 3000, 3000)}<w:tr>${HEADER}${tc('<w:vMerge w:val="restart"/>', 'Region')}${tc(SPAN2, 'Sales')}</w:tr>`
    + `<w:tr>${HEADER}${tc('<w:vMerge/>', '')}${tc('', '2024')}${tc('', '2025')}</w:tr><w:tr>${tc('', 'North')}${tc('', '1')}${tc('', '2')}</w:tr></w:tbl>`),
  '[table: listed format]\nRegion: North; Sales / 2024: 1; Sales / 2025: 2\n');
  assert.equal(await t(`${grid(3000, 3000, 3000)}<w:tr>${HEADER}${tc('', '')}${tc('', 'B')}${tc('', 'C')}</w:tr>`
    + `<w:tr>${tc(SPAN2, 'ab')}${tc('', 'c')}</w:tr><w:tr>${tc('', 'x')}${tc('', 'y')}${tc('', 'z')}</w:tr></w:tbl>`),
  '[table: listed format]\nB: ab; C: c\nx; B: y; C: z\n');
  assert.equal(await t(`${grid(3000, 3000)}<w:tr>${HEADER}${tc(SPAN2, 'Title')}</w:tr><w:tr>${HEADER}${tc('', 'H1')}${tc('', 'H2')}</w:tr>`
    + `<w:tr>${tc('', 'a')}${tc('', 'b')}</w:tr><w:tr>${tc(SPAN2, 'Part B')}</w:tr><w:tr>${tc('', 'c')}${tc('', 'd')}</w:tr></w:tbl>`),
  '[table: listed format]\nTitle\nH1: a; H2: b\nPart B\nH1: c; H2: d\n');
  assert.equal(await t(`${grid(3000, 3000, 3000)}<w:tr>${HEADER}${tc(SPAN3, 'Title')}</w:tr><w:tr>${tc('', 'a', 'a2')}${tc('', '')}${tc('', 'b')}</w:tr></w:tbl>`),
    '[table: listed format]\nTitle\na; a2\tb\n');
});

test('REQ-604, not yet LaTeX (CR-012 phase 5b): an equation is its m:t text inside the delimiters, the fallback REQ-604 names', async () => {
  assert.equal(between(await text(probe), '$$MTH32', 'Case 33'), '$$MTH32_A+MTH32_B2$$\n');
  const inline = await miniDocx('<w:p><w:r><w:t xml:space="preserve">So </w:t></w:r><m:oMath><m:r><m:t>x</m:t></m:r>'
    + '<m:sSup><m:e><m:r><m:t>y</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></w:p>').load();
  assert.equal(await text(inline), 'So $xy2$\n');
});

test('REQ-515, REQ-801: the properties always in the metadata; in the text only under metadata: true, before the body; no statistics', async () => {
  const r = await Texts.of(EXTERNAL);
  assert.equal(r.metadata.fields_update_on_open, true);
  assert.equal(r.metadata.app.Company, 'APP52_COMPANY');
  assert.equal(r.metadata.custom.ProbeCustom, 'CPR52_CUSTOM');
  assert.ok(!PlainText.of(r).includes('APP52_COMPANY'));
  const t = await text(EXTERNAL, { metadata: true });
  assert.ok(t.startsWith('[application properties: Application: Microsoft Office Word; Company: APP52_COMPANY]\n'
    + '[custom properties: ProbeCustom: CPR52_CUSTOM]\nCase 50'), t.slice(0, 200));
  assert.ok((await text(probe, { metadata: true })).startsWith('[core properties: title: PRP38_TITLE; subject: PRP38_SUBJECT]\n'));
  const stats = await miniDocx(p('Body.')).packagePart('docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'rIdApp',
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    + '<Template>Normal.dotm</Template><Pages>7</Pages><Words>1234</Words><Characters>5678</Characters></Properties>').load();
  assert.deepEqual((await Texts.of(stats)).metadata.app, { Template: 'Normal.dotm' });
});

test('REQ-501, REQ-516: the glossary only under include: glossary, last, after the endnotes', async () => {
  assert.ok(!(await text(probe)).includes('GLS54'));
  const r = await Texts.of(probe, { include: ['glossary'] });
  const glossary = r.blocks.filter((b) => b.story === 'GLOSSARY');
  assert.equal(glossary.length, 2);
  assert.equal(glossary[0].address, 'glossary:0/0');
  assert.equal(glossary[1].text, 'GLS54_ENTRY_TWO');
  assert.deepEqual(r.blocks.slice(-2), glossary);
  assert.ok(PlainText.of(r).endsWith('[1] EN27_NOTE_TEXT\nGLS54_ENTRY_ONE\nGLS54_ENTRY_TWO\n'));
});

test('REQ-514, REQ-707: the default describer\'s alt text, else the name; nothing for a decorative picture; the title never; provenance carried', async () => {
  const t = await text(probe, { alt_text: true });
  assert.equal(between(t, 'Alt text and title:', 'Case 34'),
    'Alt text and title: [image: IMG33_ALT_TEXT]\nDecorative: \nSVG, its PNG fallback in a:blip: [image: IMG33_SVG_ALT]\n');
  assert.ok(!t.includes('IMG33_TITLE'));
  assert.ok(!(await text(probe)).includes('[image'));
  const images = (await Texts.of(probe)).blocks.flatMap((b) => b.runs.filter((r) => r.kind === 'IMAGE'));
  const decorative = images.find((r) => r.image.decorative);
  assert.ok(decorative);
  assert.equal(decorative.description, undefined);
  assert.equal(decorative.image.descr, 'IMG33_DECORATIVE_DESCR');
  for (const r of images) if (r.description) assert.equal(r.description.provenance, 'AUTHOR');
  const captions = (image) => (image.decorative ? undefined : ImageDescription.generated(`a picture named ${image.name}`));
  assert.ok((await text(probe, { alt_text: true, imageDescriber: captions })).includes('Alt text and title: [image, generated: a picture named Picture 14]'));
  const vml = await miniDocx('<w:p><w:r><w:pict><v:shape style="width:10pt;height:10pt" alt="VML alt"><v:imagedata r:id="rIdNone"/></v:shape></w:pict></w:r></w:p>').load();
  assert.equal(await text(vml, { alt_text: true }), '[image: VML alt]\n');
  for (const b of (await Texts.of(probe)).blocks) assert.equal(b.decorative, false, b.text);
});

test('REQ-511, legacy WordArt: a one-paragraph text box after its anchor (charts and SmartArt are CR-012 phase 5c)', async () => {
  const wordArt = (await Texts.of(probe)).blocks.find((b) => b.story === 'TEXT_BOX' && b.text === 'WA25_TEXTPATH');
  assert.ok(wordArt);
  assert.match(wordArt.address, /\/textbox:0\/0$/);
});

const STORE = '{11111111-2222-3333-4444-555555555555}';
const control = (type, binding, placeholder, surface) => `<w:p><w:sdt><w:sdtPr>${placeholder ? '<w:showingPlcHdr/>' : ''}${binding}${type}</w:sdtPr><w:sdtContent>`
  + `<w:r><w:t xml:space="preserve">${surface}</w:t></w:r></w:sdtContent></w:sdt></w:p>`;
const binding = (xpath, store) => `<w:dataBinding w:prefixMappings="xmlns:t='urn:t'" w:xpath="${xpath}" w:storeItemID="${store}"/>`;

test('REQ-509: the bound value where the binding resolves; the surface where it does not, for other kinds of control and under bindings: surface; sdt.binding_differs', async () => {
  const r = await Texts.of(probe);
  assert.ok(PlainText.of(r).includes('The control: BND23_XML_VALUE\n'));
  const address = blockWith(r, 'BND23_XML_VALUE').address;
  assert.ok(r.warnings.some((w) => w.kind === 'sdt.binding_differs' && w.address === address), r.warnings.map(String).join(', '));
  const surface = await Texts.of(probe, { bindings: 'surface' });
  assert.ok(PlainText.of(surface).includes('The control: BND23_SURFACE_STALE\n'));
  assert.ok(surface.warnings.some((w) => w.kind === 'sdt.binding_differs'));

  const pkg = await miniDocx(
    control('<w:text/>', binding('/t:d/t:a', STORE), false, 'Bound A')                    // agrees: no warning
    + control('<w:text/>', binding('/t:d/t:empty', STORE), true, 'Click here')          // empty, a placeholder shown
    + control('<w:text/>', binding('/t:d/t:missing', STORE), false, 'Surface C')        // doesn't resolve
    + control('', binding('/t:d/t:a', STORE), false, 'Rich D')                          // rich text: the surface
    + control('<w:text/>', binding('/ns0:coreProperties[1]/ns1:title[1]', '{6C3C8BC8-F283-45AE-878A-BAB7291924A1}')
      .replace("xmlns:t='urn:t'", "xmlns:ns0='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:ns1='http://purl.org/dc/elements/1.1/'"),
    false, 'Old title'))                                                                  // the core properties
    .relationship('rIdXml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml', '../customXml/item1.xml')
    .entry('customXml/item1.xml', null, '<d xmlns="urn:t"><a>Bound A</a><empty/></d>')
    .entry('customXml/_rels/item1.xml.rels', null, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>')
    .entry('customXml/itemProps1.xml', 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
      `<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="${STORE}"/>`)
    .packagePart('docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml',
      'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'rIdCore',
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Core title</dc:title></cp:coreProperties>')
    .load();
  const b = await Texts.of(pkg);
  assert.equal(PlainText.of(b), 'Bound A\nClick here\nSurface C\nRich D\nCore title\n');
  assert.deepEqual(b.warnings.map(String), ['sdt.binding_differs at body/4']);
});

test('REQ-509 in a branch: a bound control\'s surface is its selected branch, not every branch (CR-013)', async () => {
  const ac = (choice, fallback) => '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
    + `<mc:Choice Requires="w14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
  const run = (text) => `<w:r><w:t>${text}</w:t></w:r>`;
  const bound = (content) => `<w:p><w:sdt><w:sdtPr>${binding('/t:d/t:a', STORE)}<w:text/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt></w:p>`;
  const pkg = await miniDocx(
    bound(ac(run('Bound A'), run('Old A')))                                                // run-level: agrees
    + bound(`<w:r>${ac('<w:t>Bound A</w:t>', '<w:t>Old A</w:t>')}</w:r>`)                  // in a w:r: agrees
    + bound(ac(run('Other'), run('Bound A'))))                                             // the selected branch differs
    .relationship('rIdXml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml', '../customXml/item1.xml')
    .entry('customXml/item1.xml', null, '<d xmlns="urn:t"><a>Bound A</a></d>')
    .entry('customXml/_rels/item1.xml.rels', null, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>')
    .entry('customXml/itemProps1.xml', 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
      `<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="${STORE}"/>`)
    .load();
  const b = await Texts.of(pkg);
  assert.equal(PlainText.of(b), 'Bound A\nBound A\nBound A\n');
  // every branch's text made the first two surfaces "Bound AOld A", and warned (CR-013's audit)
  assert.deepEqual(b.warnings.filter((w) => w.kind === 'sdt.binding_differs').map(String), ['sdt.binding_differs at body/2']);
});

test('REQ-708: [page n] at each w:lastRenderedPageBreak of the main story, under page_markers: true only', async () => {
  const resaved = await WordprocessingMLPackage.load(readFileSync(join(here, 'text-extraction', 'suite', 'extraction-probe.word-resave.docx')));
  assert.ok(!(await text(resaved)).includes('[page '));
  const t = await text(resaved, { page_markers: true });
  const markers = (await Texts.of(resaved)).blocks.flatMap((b) => b.runs.filter((r) => r.kind === 'PAGE_MARKER'));
  assert.ok(markers.length > 0);
  markers.forEach((m, i) => {
    assert.equal(m.pageNumber, i + 2);
    assert.ok(t.includes(`[page ${i + 2}]`));
  });
  assert.ok(!t.includes(`[page ${markers.length + 2}]`));
});

test('REQ-709: under layout: indent, four spaces a 720-twip tab stop of a paragraph\'s first-line indent; tabs kept; none in a table', async () => {
  const pkg = await miniDocx('<w:p><w:pPr><w:ind w:left="720"/></w:pPr><w:r><w:t>one stop</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:ind w:left="1440" w:firstLine="720"/></w:pPr><w:r><w:t>three</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:ind w:left="1440" w:hanging="720"/></w:pPr><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>'
    + `${p('none')}<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:pPr><w:ind w:left="720"/></w:pPr><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`)
    .part('styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'rIdStyles', `<w:styles ${NS}/>`)
    .load();
  assert.equal(await text(pkg), 'one stop\nthree\na\tb\nnone\ncell\n');
  assert.equal(await text(pkg, { layout: 'indent' }), '    one stop\n            three\n    a\tb\nnone\ncell\n');
});

test('REQ-204: with every option phase 2 takes on, in every view, each probe saves byte for byte as before', async () => {
  for (const name of ['extraction-probe.docx', 'extraction-probe-structure.docx', 'extraction-probe-external.docm']) {
    const bytes = readFileSync(join(phase0, name));
    const before = unzipSync(await (await WordprocessingMLPackage.load(bytes)).save());
    const pkg = await WordprocessingMLPackage.load(bytes);
    for (const view of ['accepted', 'original', 'markup']) {
      await Texts.of(pkg, { view, metadata: true, alt_text: true, include: ['headers', 'footers', 'comments', 'glossary'], tables: 'grid',
        page_markers: true, layout: 'indent', hidden: true, list_labels: true, field_codes: true });
    }
    const after = unzipSync(await pkg.save());
    for (const entryName of Object.keys(before)) assert.deepEqual(after[entryName], before[entryName], `${name} ${entryName}`);
  }
});

test('REQ-307: a branch\'s text is every w:t in it, whatever the view and hidden or not (Jason, 2026-10-21, docx4j\'s rule)', async () => {
  const branches = (choice, fallback) => `<w:p><mc:AlternateContent><mc:Choice Requires="w14">${choice}</mc:Choice>`
    + `<mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent></w:p>`;
  const pkg = await miniDocx(
    branches('<w:r><w:t>SAME</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN</w:t></w:r>', '<w:r><w:t>SAME</w:t></w:r>')
    + branches(`<w:r><w:t>SAME</w:t></w:r><w:ins${DEL}><w:r><w:t>ADDED</w:t></w:r></w:ins>`, '<w:r><w:t>SAME</w:t></w:r>')
    + branches('<w:r><w:t>SAME</w:t></w:r>', '<w:r><w:t xml:space="preserve">S A M E</w:t></w:r>')).load();
  for (const view of ['accepted', 'original', 'markup']) {
    for (const hidden of [false, true]) {
      const r = await Texts.of(pkg, { view, hidden });
      assert.deepEqual(r.warnings.map(String), ['mce.branches_differ at body/0 (mc:Choice Requires="w14")',
        'mce.branches_differ at body/1 (mc:Choice Requires="w14")'], `${view}, hidden: ${hidden}`);
    }
  }
  assert.equal(await text(pkg), 'SAME\nSAMEADDED\nSAME\n');
  assert.deepEqual((await Texts.of(pkg, { warnings: false })).warnings, []);
});

// docx4j CR-043 section 17 (29dbf5964), its ViewAgreementTest over core-ts's check 54 saves: the accepted view reads
// as this package's Accept All copy reads, and as Word's own Accept All save, block by block, Word's three rules where no
// paragraph follows a deleted mark among them (cases 9, 13 and 19), and a table whose every row goes (case 6)
test("check 54's save, its accepted view, its Accept All copy and Word's Accept All save read alike, block by block (docx4j CR-043 section 17)", async () => {
  const saves = join(here, 'check54');
  const blocksOf = (result) => {
    const view = (blocks) => blocks.map((b) => (b.rows.length > 0
      ? { rows: b.rows.map((r) => r.cells.map((c) => view(c.blocks))) }
      : { kind: b.kind, style: b.style, text: b.runs.map((r) => r.text).join('') }));
    return view(result.blocks);
  };
  const run = await WordprocessingMLPackage.load(readFileSync(join(saves, '54-writes-word365.docx')));
  const viewed = blocksOf(await Texts.of(run, { view: 'accepted' }));
  const copy = await WordprocessingMLPackage.load(readFileSync(join(saves, '54-writes-word365.docx')));
  await copy.acceptAllRevisions();
  assert.deepEqual(viewed, blocksOf(await Texts.ofResolved(copy, { view: 'accepted' })));
  // and Word's: a region's blocks after its BEGIN paragraph, for the cases this rule set decides
  const word = blocksOf(await Texts.ofResolved(await WordprocessingMLPackage.load(readFileSync(join(saves, '54-writes-word365-accepted.docx'))), { view: 'accepted' }));
  const region = (blocks, n) => {
    const begin = blocks.findIndex((b) => b.text?.startsWith(`BEGIN${n} `));
    const end = blocks.findIndex((b, k) => k > begin && b.text === `END${n}`);
    return blocks.slice(begin + 1, end);
  };
  for (const n of [6, 9, 13, 15, 17, 19]) assert.deepEqual(region(viewed, n), region(word, n), `case ${n}`);
});
