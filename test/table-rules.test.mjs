// CR-002 section 55: Word's table rule, a cell ends with a paragraph as Word reads it. docx4j's TableRulesTest
// (18a3a31f4) ported, and the edits of section 55.2's table held to section 55.6's answers, which are Word 365's in
// check 52 (cases 86 to 95; ranges.test.mjs holds the shim's run of those cases to Word's run and save).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { WordprocessingMLPackage, contentOf, endsWithParagraph, endWithParagraph, marshalString } from '../dist/index.mjs';

// --- docx4j's TableRulesTest -------------------------------------------------------------------------------------

const NS = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w99x="urn:docx4j:probe:unknown"';
const P = '<w:p><w:r><w:t>p</w:t></w:r></w:p>';
const TBL = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>${P}</w:tc></w:tr></w:tbl>`;
const sdt = (content) => `<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const mc = (requires, choice, fallback) => `<mc:AlternateContent><mc:Choice Requires="${requires}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;

/** A `w:tc` holding `content`, unmarshalled as it is: `contentOf` repairs nothing. */
async function tc(content) {
  const [tbl] = await contentOf(`<w:tbl ${NS}><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>${content}</w:tc></w:tr></w:tbl>`);
  return tbl.value.content[0].value.content[0].value;
}
const opens = async (content) => assert.equal(endsWithParagraph((await tc(content)).content), true, content);
const unreadable = async (content) => assert.equal(endsWithParagraph((await tc(content)).content), false, content);

test("docx4j's TableRulesTest: the probes Word 365 opens and the ones it repairs, through controls and the branch Word reads", async () => {
  await opens(TBL + P); // 01
  await opens(TBL + sdt(P)); // 03
  await opens(TBL + sdt(sdt(P))); // 04
  await opens(TBL + mc('w14', P, P)); // 06
  await opens(TBL + mc('w99x', P, P)); // 07, the Fallback read
  await opens(TBL + mc('w14', P, TBL)); // 08, the unread Fallback not checked
  await opens(sdt(P)); // 10
  await opens(mc('w14', P, P)); // 11
  await unreadable(TBL); // 02
  await unreadable(TBL + sdt(TBL)); // 05
  await unreadable(TBL + mc('w14', TBL, P)); // 09, the read Choice ends in a table
  await unreadable(sdt(TBL)); // 12
  await unreadable(''); // no block at all
});

test("docx4j's TableRulesTest: range markup after the last block doesn't count; the repair appends once, after the wrapper", async () => {
  await opens(`${P}<w:bookmarkEnd w:id="0"/>`);
  await unreadable(`${TBL}<w:bookmarkEnd w:id="0"/>`);
  const cell = await tc(TBL + sdt(TBL));
  assert.ok(endWithParagraph(cell));
  assert.match(await marshalString({ name: { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localPart: 'tc' }, value: cell }), /<\/w:sdt><w:p\/><\/w:tc>$/);
  assert.equal(endWithParagraph(cell), undefined, 'already ends with a paragraph');
  assert.equal(cell.content.length, 3);
  assert.equal(cell.content[2].value.PARENT, cell);
  const wrapped = await tc(TBL + sdt(P));
  assert.equal(endWithParagraph(wrapped), undefined);
  assert.equal(wrapped.content.length, 2);
});

// --- the edits ---------------------------------------------------------------------------------------------------

/** The cell's paragraph: a heading, centred, its text underlined (check 52's was a list item, underlined). */
const CELL_P = '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>CELL</w:t></w:r></w:p>';
const NESTED = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>NESTED</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
const NEW_TABLE = NESTED.replace('NESTED', 'NEWTABLE');

/** A new document whose one table's first cell holds `content` (p: the paragraph; tp: a nested table, then it). */
async function cellHolding(content) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  await body.insertXml(`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${content}</w:tc></w:tr></w:tbl>`, 'End');
  return { pkg, cell: () => pkg.body.tables[0].getCell(0, 0) };
}

/** A run list's text, a revision holder's as `ins:` or `del:`. */
const runs = (items) => (items ?? []).map((el) => {
  const name = el.name?.localPart;
  if (name === 'ins' || name === 'del') return `${name}:${runs(el.value.customXmlOrSmartTagOrSdt)}`;
  if (name === 'r') return (el.value.content ?? []).map((c) => (typeof c.value?.value === 'string' ? c.value.value : '')).join('');
  return '';
}).join('');

/** A cell's blocks: a paragraph as its style, alignment and mark (`u`, `ins`, `del`), then its text; a table; a control. */
const shape = (items) => (items ?? []).map((el) => {
  const v = el.value;
  switch (el.name?.localPart) {
    case 'p': {
      const pPr = v.pPr ?? {};
      const mark = ['u', 'ins', 'del'].filter((k) => pPr.rPr?.[k]).join(',');
      return `P[${[pPr.pStyle?.val, pPr.jc?.val].filter(Boolean).join(',')}|${mark}] ${runs(v.content)}`.trimEnd();
    }
    case 'tbl': return `TBL${v.content?.some((tr) => tr.value.trPr?.ins) ? '+ins' : ''}${v.content?.some((tr) => tr.value.trPr?.del) ? '+del' : ''}`;
    case 'sdt': return `SDT{${shape(v.sdtContent?.content).join(' ')}}`;
    default: return `<${el.name?.localPart}>`;
  }
});
const cellShape = (cell) => shape(cell().body.content);

test('an edit that would take the paragraph a cell ends with keeps it, emptied, its properties kept (check 52, cases 88, 89, 92, 94 and 95)', async () => {
  for (const [content, before] of [[CELL_P, []], [NESTED + CELL_P, ['TBL']]]) {
    let { cell } = await cellHolding(content);
    cell().body.paragraphs.at(-1).delete();
    assert.deepEqual(cellShape(cell), [...before, 'P[Heading1,center|]'], 'Paragraph.delete');

    ({ cell } = await cellHolding(content));
    cell().body.clear();
    assert.deepEqual(cellShape(cell), ['P[Heading1,center|]'], 'Body.clear');

    ({ cell } = await cellHolding(content));
    await cell().body.insertOoxml(NEW_TABLE, 'Replace');
    assert.deepEqual(cellShape(cell), ['TBL', 'P[Heading1,center|]'], "Body.insertOoxml(table, 'Replace')");

    ({ cell } = await cellHolding(content));
    await cell().body.paragraphs.at(-1).insertOoxml(NEW_TABLE, 'Replace');
    assert.deepEqual(cellShape(cell), [...before, 'TBL', 'P[Heading1,center|]'], "Paragraph.insertOoxml(table, 'Replace')");
  }
  // a control around the cell's paragraph, deleted with its content: the paragraph stays, outside it (case 95)
  let { cell } = await cellHolding(CELL_P);
  cell().body.insertContentControl();
  assert.deepEqual(cellShape(cell), ['SDT{P[Heading1,center|] CELL}']);
  cell().body.contentControls[0].delete(false);
  assert.deepEqual(cellShape(cell), ['P[Heading1,center|]']);
  ({ cell } = await cellHolding(NESTED + CELL_P));
  cell().body.paragraphs.at(-1).insertContentControl();
  cell().body.contentControls[0].delete(false);
  assert.deepEqual(cellShape(cell), ['TBL', 'P[Heading1,center|]']);
  // where the paragraph before keeps the rule, the last one goes, as before
  ({ cell } = await cellHolding(CELL_P.replace('CELL', 'FIRST') + CELL_P));
  cell().body.paragraphs.at(-1).delete();
  assert.deepEqual(cellShape(cell), ['P[Heading1,center|] FIRST']);
  // a Replace whose content ends with a paragraph replaces the cell's paragraph, as before
  ({ cell } = await cellHolding(CELL_P));
  await cell().body.insertOoxml('<w:p><w:r><w:t>NEW</w:t></w:r></w:p>', 'Replace');
  assert.deepEqual(cellShape(cell), ['P[|] NEW']);
});

test("tracked, a cell's last paragraph has its text deleted and its mark kept; Accept All leaves it emptied, Reject All as it was (check 52, cases 90, 91 and 93)", async () => {
  for (const [content, before] of [[CELL_P, []], [NESTED + CELL_P, ['TBL']]]) {
    for (const edit of ['delete', 'clear']) {
      const { pkg, cell } = await cellHolding(content);
      pkg.changeTrackingMode = 'TrackAll';
      if (edit === 'delete') cell().body.paragraphs.at(-1).delete(); else cell().body.clear();
      pkg.changeTrackingMode = 'Off';
      // a tracked clear deletes the nested table's rows too
      const nested = edit === 'clear' ? before.map(() => 'TBL+del') : before;
      assert.deepEqual(cellShape(cell), [...nested, 'P[Heading1,center|] del:CELL'], `${edit} ${content === CELL_P ? 'p' : 'tp'}`);
      const saved = await pkg.save();
      const accepted = await WordprocessingMLPackage.load(saved);
      await accepted.getBody();
      await accepted.acceptAllRevisions();
      assert.deepEqual(shape(accepted.body.tables[0].getCell(0, 0).body.content), [...(edit === 'clear' ? [] : before), 'P[Heading1,center|]'], `${edit}, accepted`);
      const rejected = await WordprocessingMLPackage.load(saved);
      await rejected.getBody();
      await rejected.rejectAllRevisions();
      assert.deepEqual(shape(rejected.body.tables[0].getCell(0, 0).body.content), [...before, 'P[Heading1,center|] CELL'], `${edit}, rejected`);
    }
  }
});

test("a table at a cell's end: the cell's paragraph is split at its end and its mark stays last, after the table (check 52, cases 86 and 87)", async () => {
  const insertions = {
    insertTable: (cell) => cell().body.insertTable(1, 1, 'End', [['NEWTABLE']]),
    insertXml: (cell) => cell().body.insertXml(NEW_TABLE, 'End'),
    'insertOoxml After': (cell) => cell().body.paragraphs.at(-1).insertOoxml(NEW_TABLE, 'After'),
  };
  for (const [name, insert] of Object.entries(insertions)) {
    let { pkg, cell } = await cellHolding(CELL_P);
    await insert(cell);
    // the first half's mark in its text's formatting (case 86's w:u), the cell's mark with its properties after the table
    assert.deepEqual(cellShape(cell), ['P[Heading1,center|u] CELL', 'TBL', 'P[Heading1,center|]'], name);
    ({ pkg, cell } = await cellHolding(CELL_P));
    pkg.changeTrackingMode = 'TrackAll';
    await insert(cell);
    pkg.changeTrackingMode = 'Off';
    // tracked, the inserted mark is the first half's, with the rows; the cell's own mark is not marked (case 87)
    assert.deepEqual(cellShape(cell), ['P[Heading1,center|u,ins] CELL', 'TBL+ins', 'P[Heading1,center|]'], `${name}, tracked`);
    const saved = await pkg.save();
    const accepted = await WordprocessingMLPackage.load(saved);
    await accepted.getBody();
    await accepted.acceptAllRevisions();
    assert.deepEqual(shape(accepted.body.tables[0].getCell(0, 0).body.content), ['P[Heading1,center|u] CELL', 'TBL', 'P[Heading1,center|]'], `${name}, accepted`);
    const rejected = await WordprocessingMLPackage.load(saved);
    await rejected.getBody();
    await rejected.rejectAllRevisions();
    assert.deepEqual(shape(rejected.body.tables[0].getCell(0, 0).body.content), ['P[Heading1,center|] CELL'], `${name}, rejected`);
  }
  // at the start the rule holds, and nothing is added
  let { cell } = await cellHolding(CELL_P);
  cell().body.insertTable(1, 1, 'Start');
  assert.deepEqual(cellShape(cell), ['TBL', 'P[Heading1,center|] CELL']);
  // the paragraph keeps its text, so a range over it still holds it: Range.insertOoxml(table, 'Replace'), unmeasured
  // in Word, puts the table after it, then takes its text
  for (const where of ['Content', 'Whole']) {
    ({ cell } = await cellHolding(CELL_P));
    await cell().body.paragraphs[0].getRange(where).insertOoxml(NEW_TABLE, 'Replace');
    assert.deepEqual(cellShape(cell), ['P[Heading1,center|u]', 'TBL', 'P[Heading1,center|]'], where);
  }
});

test("elsewhere docx4j's repair: a table at a control's end in a cell, the cells of inserted content (D3); what kept the rule still keeps it", async () => {
  // a table put at the end of a control's content, the control the cell's last block: an empty paragraph after the control
  let { pkg, cell } = await cellHolding(CELL_P);
  const control = cell().body.insertContentControl();
  await control.paragraphs[0].insertOoxml(NEW_TABLE, 'After');
  assert.deepEqual(cellShape(cell), ['SDT{P[Heading1,center|] CELL TBL}', 'P[|]']);
  // D3: a table whose cell ends with a table, inserted, has that cell repaired; the outer cell keeps its paragraph
  ({ pkg, cell } = await cellHolding(CELL_P));
  const inner = `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${NESTED}</w:tc></w:tr></w:tbl>`;
  await pkg.body.insertXml(inner, 'End');
  assert.deepEqual(shape(pkg.body.tables[1].getCell(0, 0).body.content), ['TBL', 'P[|]']);
  // what kept the rule: a cell's value, Range.delete of the whole range, insertText Replace, a tracked value
  ({ pkg, cell } = await cellHolding(NESTED + CELL_P));
  cell().value = 'VALUE';
  assert.deepEqual(cellShape(cell), ['P[|] VALUE']);
  ({ pkg, cell } = await cellHolding(CELL_P));
  cell().body.paragraphs[0].getRange('Whole').delete();
  assert.deepEqual(cellShape(cell), ['P[Heading1,center|]']);
  ({ pkg, cell } = await cellHolding(CELL_P));
  cell().body.insertText('TEXT', 'Replace');
  assert.deepEqual(cellShape(cell), ['P[|] TEXT']);
  ({ pkg, cell } = await cellHolding(CELL_P));
  pkg.changeTrackingMode = 'TrackAll';
  cell().value = 'VALUE';
  pkg.changeTrackingMode = 'Off';
  await pkg.acceptAllRevisions();
  assert.deepEqual(cellShape(cell), ['P[|] VALUE']);
});

test('D4: a file whose cell ends with a table is left as it is until an edit touches that cell', async () => {
  const { pkg, cell } = await cellHolding(CELL_P);
  // made broken behind the API's back, as a file from elsewhere is
  const [nested] = await contentOf(NESTED);
  cell().body.content.push(nested);
  const again = await WordprocessingMLPackage.load(await pkg.save());
  await again.getBody();
  assert.deepEqual(shape(again.body.tables[0].getCell(0, 0).body.content), ['P[Heading1,center|] CELL', 'TBL']);
  // an edit in that cell repairs it, docx4j's way
  again.body.tables[0].getCell(0, 0).body.insertParagraph('FIRST', 'Start');
  assert.deepEqual(shape(again.body.tables[0].getCell(0, 0).body.content), ['P[|] FIRST', 'P[Heading1,center|] CELL', 'TBL', 'P[|]']);
});
