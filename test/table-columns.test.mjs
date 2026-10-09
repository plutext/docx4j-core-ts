// CR-002 section 42: columns added and removed in Word's three table layout modes, held to Word 365's
// save of the editor's probe (test/README.md check 39, test/fixtures/check39/).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { WordprocessingMLPackage, Table, TableCell } from '../dist/index.mjs';
import { fixture } from './helpers.mjs';

const DATE = new Date(Date.UTC(2026, 9, 9, 10, 30, 0));

/** The widths that matter of a table: w:tblW, the grid, and each row's w:tcW. */
function shape(table) {
  const tblW = table.tbl.tblPr?.tblW;
  return {
    tblW: tblW ? `${tblW.w} ${tblW.type}` : undefined,
    grid: table.tbl.tblGrid?.gridCol?.map((c) => c.w),
    rows: table.rows.map((row) => row.cells.map((cell) => { const w = cell.tc.tcPr?.tcW; return w ? `${w.w} ${w.type}` : undefined; })),
  };
}

async function probe(name = 'table-autofit-probe.docx') {
  const pkg = await WordprocessingMLPackage.load(await fixture(`check39/${name}`));
  const body = await pkg.getBody();
  return { pkg, body, tables: body.tables };
}

test('check 39: Insert Right in the first cell of each mode, held to Word 365\'s save', async () => {
  const { tables } = await probe();
  const saved = (await probe('table-autofit-probe-word365.docx')).tables;
  assert.deepEqual(tables.map((t) => t.layoutMode), ['fixed', 'window', 'contents']);
  for (const table of tables) table.getCell(0, 0).insertColumns('After', 1);

  // fixed: the neighbour's width copied, the grid grown, w:tblW grown past the margin; the save exactly
  assert.deepEqual(shape(tables[0]), shape(saved[0]));
  assert.deepEqual(shape(tables[0]), { tblW: '13539 dxa', grid: [4513, 4513, 4513], rows: [['4513 dxa', '4513 dxa', '4513 dxa'], ['4513 dxa', '4513 dxa', '4513 dxa']] });

  // window: the table keeps 100%, the cells' shares are the cumulative floored percentages; the grid is
  // rescaled to the width it had (Word's own grid, 3005/3006/3006, is its layout's measurement, 9017)
  const window = shape(tables[1]);
  assert.equal(window.tblW, saved[1].tbl.tblPr.tblW.w + ' ' + saved[1].tbl.tblPr.tblW.type);
  assert.deepEqual(window.rows, shape(saved[1]).rows);
  assert.deepEqual(window.rows[0], ['1666 pct', '1667 pct', '1667 pct']);
  assert.equal(window.grid.length, 3);
  assert.equal(window.grid.reduce((a, b) => a + b, 0), 9026);
  assert.deepEqual(window.grid, [3009, 3009, 3008]);

  // contents: auto throughout; the grid of the right shape (Word's 593/222/597 is its measurement of the text)
  const contents = shape(tables[2]);
  assert.equal(contents.tblW, '0 auto');
  assert.deepEqual(contents.rows, shape(saved[2]).rows);
  assert.deepEqual(contents.rows[0], ['0 auto', '0 auto', '0 auto']);
  assert.equal(contents.grid.length, 3);
});

test('addColumns at the start and the end, insertColumns before and after, with values; the new cells per row come back; saved and reloaded', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const table = body.insertTable(2, 2, 'End', [['a', 'b'], ['c', 'd']]);
  assert.equal(table.layoutMode, 'contents');
  const atEnd = table.addColumns('End', 1, [['e'], ['f']]);
  assert.equal(atEnd.length, 2);
  assert.ok(atEnd[0][0] instanceof TableCell);
  assert.deepEqual(table.values, [['a', 'b', 'e'], ['c', 'd', 'f']]);
  const atStart = table.addColumns('Start', 2, [['x', 'y'], ['z']]);
  assert.deepEqual(atStart.map((cells) => cells.map((c) => c.text)), [['x', 'y'], ['z', '']]);
  assert.deepEqual(table.values, [['x', 'y', 'a', 'b', 'e'], ['z', '', 'c', 'd', 'f']]);
  table.getCell(1, 2).insertColumns('Before', 1, [['p'], ['q']]);
  table.getCell(0, 0).insertColumns('After', 1);
  assert.deepEqual(table.values, [['x', '', 'y', 'p', 'a', 'b', 'e'], ['z', '', '', 'q', 'c', 'd', 'f']]);
  assert.equal(table.tbl.tblGrid.gridCol.length, 7);
  assert.equal(table.columnWidths().length, 7);
  assert.throws(() => table.addColumns('End', 0), /positive integer/);
  const again = await WordprocessingMLPackage.load(await pkg.save());
  const back = (await again.getBody()).tables[0];
  assert.deepEqual(back.values, table.values);
  assert.equal(back.rows[0].cellCount, 7);
});

test('the new cell copies the column to the right of the boundary, else the left (check 40): its width in its form, its shading and its paragraph\'s properties, not its borders, alignment, span or merges (check 41); fixed mode grows the table by it', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  await body.insertXml(`<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid>`
    + `<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcBorders><w:left w:val="single" w:sz="24" w:space="0" w:color="FF0000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>`
    + `<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`, 'End');
  const table = body.tables[0];
  assert.equal(table.layoutMode, 'fixed');
  // after the shaded cell: the column to the right (4,000, the merged one) is copied, less its vMerge, and not the shading
  const [[afterShaded], [belowIt]] = table.getCell(0, 0).insertColumns('After', 1);
  assert.equal(afterShaded.tc.tcPr.shd, undefined, 'the right-hand neighbour has no shading');
  assert.equal(afterShaded.tc.tcPr.vMerge, undefined);
  assert.equal(belowIt.tc.tcPr.vMerge, undefined);
  assert.equal(afterShaded.paragraphs[0].p.pPr, undefined, 'the neighbour\'s paragraph has no properties');
  assert.deepEqual(shape(table), { tblW: '10000 dxa', grid: [2000, 4000, 4000], rows: [['2000 dxa', '4000 dxa', '4000 dxa'], ['2000 dxa', '4000 dxa', '4000 dxa']] });
  // before the shaded cell: it is the column to the right, so its shading and its paragraph's centring and bold continue; its border and vertical alignment do not
  const [[yellow]] = table.getCell(0, 0).insertColumns('Before', 1);
  assert.equal(yellow.tc.tcPr.shd.fill, 'FFFF00');
  assert.equal(yellow.tc.tcPr.vAlign, undefined, 'the vertical alignment is not copied (check 41)');
  assert.equal(yellow.tc.tcPr.tcBorders, undefined, 'the borders are not copied (check 41)');
  assert.equal(yellow.paragraphs[0].p.pPr.jc.val, 'center');
  assert.ok(yellow.paragraphs[0].p.pPr.rPr.b, 'the mark\'s run properties come too');
  assert.deepEqual(shape(table).grid, [2000, 2000, 4000, 4000]);
  assert.equal(shape(table).tblW, '12000 dxa');
  // at the end there is no column to the right: the last is copied
  table.addColumns('End', 1);
  assert.deepEqual(shape(table).grid, [2000, 2000, 4000, 4000, 4000]);
});

test('a boundary inside a spanning cell puts the new cell before it, the span kept (check 40b1); at the edge of a row\'s skipped columns the cell goes in; strictly inside them they grow', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  await body.insertXml(`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>`
    + `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr>`
    + `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr><w:tc><w:p><w:r><w:t>d</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>e</w:t></w:r></w:p></w:tc></w:tr>`
    + `<w:tr><w:trPr><w:gridBefore w:val="2"/></w:trPr><w:tc><w:p><w:r><w:t>f</w:t></w:r></w:p></w:tc></w:tr>`
    + `<w:tr><w:tc><w:p><w:r><w:t>g</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>h</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>i</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`, 'End');
  const table = body.tables[0];
  // after column 0: inside the first row's span; at the second row's skipped edge; inside the third row's skipped columns; before h in the fourth
  const added = table.getCell(3, 0).insertColumns('After', 1, [['x'], ['y'], ['-'], ['new']]);
  assert.deepEqual(added.map((cells) => cells.map((c) => c.text)), [['x'], ['y'], [], ['new']]);
  assert.equal(table.getCell(0, 1).tc.tcPr.gridSpan.val, 2, 'the span is kept, the new cell before it');
  assert.deepEqual(table.values, [['x', 'wide', 'c'], ['y', 'd', 'e'], ['f'], ['g', 'new', 'h', 'i']]);
  const skipped = (row) => table.rows[row].tr.trPr.cnfStyleOrDivIdOrGridBefore.find((i) => i.name.localPart === 'gridBefore').value.val;
  assert.equal(skipped(1), 1, 'at the edge of the skipped columns the new cell is in the row');
  assert.equal(skipped(2), 3, 'strictly inside them, the skipped columns grow');
  assert.deepEqual(shape(table).grid, [3000, 3000, 3000, 3000]);
});

test('Word check 40 (2026-10-09): the four hand insertions, held to Word 365\'s saves', async () => {
  const cases = [
    ['40a1-window-unequal.docx', '40a1-word365.docx', (t) => t.getCell(0, 0).insertColumns('After', 1)],
    ['40a2-window-unequal.docx', '40a2-word365.docx', (t) => t.getCell(0, 1).insertColumns('Before', 1)],
    ['40b1-span.docx', '40b1-word365.docx', (t) => t.getCell(1, 0).insertColumns('After', 1)],
    ['40b2-span.docx', '40b2-word365.docx', (t) => t.getCell(0, 0).insertColumns('After', 1)],
  ];
  for (const [original, saved, act] of cases) {
    const table = (await (await WordprocessingMLPackage.load(await fixture(`check40/${original}`))).getBody()).tables[0];
    act(table);
    const word = shape((await (await WordprocessingMLPackage.load(await fixture(`check40/${saved}`))).getBody()).tables[0]);
    const engine = shape(table);
    if (original.startsWith('40a')) {
      // window mode: the table's width, every cell's share and the values as Word wrote them; the grid is Word's measurement (9,017 against 9,026)
      assert.deepEqual({ ...engine, grid: engine.grid.length }, { ...word, grid: word.grid.length }, original);
      assert.deepEqual(word.rows[0], ['996 pct', '2002 pct', '2002 pct'], `${original}: the wide column copied (the one to the right of the boundary), the rescale proportional`);
    } else {
      assert.deepEqual(engine, word, original);
    }
  }
});

test('a table without a w:tblGrid gets one from its first row', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  await body.insertXml(`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`, 'End');
  const table = body.tables[0];
  table.addColumns('End', 1, [['c']]);
  assert.deepEqual(table.values, [['a', 'b', 'c']]);
  assert.deepEqual(shape(table).grid, [3000, 4513, 4513]);
});

test('deleteColumns and TableCell.deleteColumn: fixed shrinks the table, window writes the shares again, a spanning cell shrinks, the last column deletes the table', async () => {
  const { tables } = await probe();
  const [fixed, window, contents] = tables;
  for (const table of tables) table.getCell(0, 0).insertColumns('After', 1);
  fixed.deleteColumns(1);
  assert.deepEqual(shape(fixed), { tblW: '9026 dxa', grid: [4513, 4513], rows: [['4513 dxa', '4513 dxa'], ['4513 dxa', '4513 dxa']] });
  window.getCell(0, 1).deleteColumn();
  // the grid, rescaled twice in integer twips, is two twips off symmetric; the cells' shares, what Word reads in this mode, are back to halves
  assert.deepEqual(shape(window), { tblW: '5000 pct', grid: [4514, 4512], rows: [['2500 pct', '2500 pct'], ['2500 pct', '2500 pct']] });
  contents.deleteColumns(0, 2);
  assert.deepEqual(contents.values, [['A1B'], ['A2B']]);
  assert.equal(shape(contents).grid.length, 1);
  assert.throws(() => contents.deleteColumns(1), /No columns/);

  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  await body.insertXml(`<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>`
    + `<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`, 'End');
  const table = body.tables[0];
  table.deleteColumns(1);
  assert.deepEqual(table.values, [['wide', 'c']]);
  assert.equal(table.getCell(0, 0).tc.tcPr.gridSpan, undefined, 'the span shrank to one');
  assert.deepEqual(shape(table), { tblW: '6000 dxa', grid: [3000, 3000], rows: [['3000 dxa', '3000 dxa']] });
  table.deleteColumns(0, 2);
  assert.equal(body.tables.length, 0, 'the last columns gone, the table goes');
});

test('TableCell.width is written in the form the cell or its table states: percent on a window table, twips otherwise', async () => {
  const { tables } = await probe();
  const [fixed, window, contents] = tables;
  window.getCell(0, 0).width = 4513 / 20 / 2;
  assert.deepEqual(shape(window).rows[0], ['1250 pct', '2500 pct']);
  fixed.getCell(0, 0).width = 100;
  assert.equal(shape(fixed).rows[0][0], '2000 dxa');
  contents.getCell(0, 0).width = 100;
  assert.equal(shape(contents).rows[0][0], '2000 dxa', 'a stated width is what was asked for');
  assert.equal(contents.getCell(0, 0).width, 100);
});

test('tracked: a column added is the table change Office JS records (check 40c): the old properties, every cell\'s properties the new ones included, the grid as after, the text inserted with no mark insertion; accepted the records go, rejected the recorded table comes back and the column stays, empty', async () => {
  const make = async () => {
    const pkg = await WordprocessingMLPackage.createPackage();
    pkg.author = { name: 'Ada' };
    pkg.trackedChangeDate = DATE;
    const body = pkg.body;
    const table = body.insertTable(2, 2, 'End', [['a', 'b'], ['c', 'd']]);
    table.tbl.tblPr.tblLayout = { type: 'fixed' };
    table.tbl.tblPr.tblW = { w: 9027, type: 'dxa' };
    await body.insertXml('<w:p><w:r><w:t>after</w:t></w:r></w:p>', 'End');
    pkg.changeTrackingMode = 'TrackAll';
    return { pkg, body, table };
  };
  const { pkg, body, table } = await make();
  table.getCell(0, 0).insertColumns('After', 1, [['new'], ['']]);
  assert.deepEqual(table.values, [['a', 'new', 'b'], ['c', '', 'd']]);
  const tbl = table.tbl;
  assert.equal(tbl.tblPr.tblPrChange.author, 'Ada');
  assert.equal(tbl.tblPr.tblPrChange.tblPr.tblW.w, 9027, 'the old width recorded');
  assert.equal(tbl.tblPr.tblW.w, 9027 + 4514, 'fixed: the table grows by the copied column');
  assert.deepEqual(tbl.tblGrid.gridCol.map((c) => c.w), [4513, 4514, 4514], 'the column to the right of the boundary copied');
  assert.deepEqual(tbl.tblGrid.tblGridChange.tblGrid.gridCol.map((c) => c.w), [4513, 4514, 4514], 'the record holds the grid as it is after, as Word writes it');
  const cells = table.rows.flatMap((r) => r.cells);
  assert.deepEqual(cells.map((c) => c.tc.tcPr?.tcPrChange !== undefined), [true, true, true, true, true, true], 'every cell records its properties, the new ones included');
  assert.equal(table.getCell(0, 1).paragraphs[0].p.pPr, undefined, 'no mark insertion on the new cell\'s paragraph');
  const xml = await pkg.getMainDocumentPart().getXml();
  assert.match(xml, /<w:tblPrChange w:id="\d+"[^>]*w:author="Ada"[^>]*><w:tblPr>/);
  assert.match(xml, /<w:tblGridChange w:id="\d+"><w:tblGrid><w:gridCol w:w="4513"\/><w:gridCol w:w="4514"\/><w:gridCol w:w="4514"\/><\/w:tblGrid><\/w:tblGridChange>/);
  assert.match(xml, /<w:p><w:ins [^>]*><w:r><w:t>new<\/w:t><\/w:r><\/w:ins><\/w:p>/);
  assert.match(xml, /<w:tc><w:tcPr><w:tcW [^>]*\/><w:tcPrChange [^>]*><w:tcPr><w:tcW [^>]*\/><\/w:tcPr><\/w:tcPrChange><\/w:tcPr><w:p\/><\/w:tc>/, 'an empty new cell: an empty paragraph, its properties recorded');
  assert.doesNotMatch(xml, /w:cellIns/);
  const changes = body.getTrackedChanges();
  assert.ok(changes.some((c) => c.target.kind === 'tableProperties' && c.type === 'Formatted'), changes.map((c) => c.target.kind).join(','));
  assert.ok(changes.some((c) => c.type === 'Added' && c.text === 'new'), 'the new cell\'s text is an insertion');

  // accepted: the records go, the column and its text stay
  assert.ok(body.acceptAll() > 0);
  assert.deepEqual(table.values, [['a', 'new', 'b'], ['c', '', 'd']]);
  assert.doesNotMatch(await pkg.getMainDocumentPart().getXml(), /PrChange|GridChange|<w:ins /);

  // rejected: the recorded properties and cells come back, the grid as recorded; the column stays, empty (Word's Reject All, section 29)
  const fresh = await make();
  fresh.table.getCell(0, 0).insertColumns('After', 1, [['new'], ['']]);
  assert.ok(fresh.body.rejectAll() > 0);
  assert.deepEqual(fresh.table.values, [['a', '', 'b'], ['c', '', 'd']]);
  assert.deepEqual(fresh.table.tbl.tblGrid.gridCol.map((c) => c.w), [4513, 4514, 4514]);
  assert.equal(fresh.table.tbl.tblPr.tblW.w, 9027, 'the recorded width back');
  assert.doesNotMatch(await fresh.pkg.getMainDocumentPart().getXml(), /PrChange|GridChange|<w:ins /);
});

test('tracked: a column deleted loses its cells untracked and records the table change, as Office JS does (check 40c); rejected, the recorded width comes back', async () => {
  const D = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9026" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3009"/><w:gridCol w:w="3009"/><w:gridCol w:w="3008"/></w:tblGrid>`
    + ['1', '2'].map((n) => `<w:tr>${['A', 'B', 'C'].map((c, i) => `<w:tc><w:tcPr><w:tcW w:w="${i === 2 ? 3008 : 3009}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>D${n}${c}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`).join('') + '</w:tbl>';
  const make = async () => {
    const pkg = await WordprocessingMLPackage.createPackage();
    pkg.author = { name: 'Author A' };
    pkg.trackedChangeDate = DATE;
    const body = pkg.body;
    await body.insertXml(D, 'End');
    pkg.changeTrackingMode = 'TrackAll';
    return { pkg, body, table: body.tables[0] };
  };
  const { pkg, body, table } = await make();
  table.deleteColumns(1);
  assert.deepEqual(table.values, [['D1A', 'D1C'], ['D2A', 'D2C']], 'the cells are gone');
  assert.deepEqual(shape(table), { tblW: '6017 dxa', grid: [3009, 3008], rows: [['3009 dxa', '3008 dxa'], ['3009 dxa', '3008 dxa']] });
  assert.equal(table.tbl.tblPr.tblPrChange.tblPr.tblW.w, 9026, 'the old width recorded');
  assert.deepEqual(table.tbl.tblGrid.tblGridChange.tblGrid.gridCol.map((c) => c.w), [3009, 3008], 'the grid as after, as Word records it');
  assert.ok(table.rows.flatMap((r) => r.cells).every((c) => c.tc.tcPr.tcPrChange), 'every remaining cell records its properties');
  const xml = await pkg.getMainDocumentPart().getXml();
  assert.doesNotMatch(xml, /<w:del |cellDel/);
  assert.deepEqual(body.getTrackedChanges().map((c) => [c.target.kind, c.type, c.text]), [['tableProperties', 'Formatted', 'D1A\tD1C\r\nD2A\tD2C\r\n']], 'listed as Office JS lists it');
  assert.equal(body.acceptAll(), 1);
  assert.doesNotMatch(await pkg.getMainDocumentPart().getXml(), /PrChange|GridChange/);
  const fresh = await make();
  fresh.table.deleteColumns(1);
  assert.equal(fresh.body.rejectAll(), 1);
  assert.equal(fresh.table.tbl.tblPr.tblW.w, 9026, 'the recorded width back; the cells stay gone');
  assert.deepEqual(fresh.table.values, [['D1A', 'D1C'], ['D2A', 'D2C']]);
});

test('Word check 40c (2026-10-09): the engine\'s tracked delete and add over the Office JS file, held to Word 365\'s save', async () => {
  const records = (t) => ({
    tblW: t.tbl.tblPr.tblW.w, oldTblW: t.tbl.tblPr.tblPrChange?.tblPr.tblW.w,
    grid: t.tbl.tblGrid.gridCol.map((c) => c.w), recordedGrid: t.tbl.tblGrid.tblGridChange?.tblGrid.gridCol.map((c) => c.w),
    rows: t.rows.map((r) => r.cells.map((c) => ({ w: c.tc.tcPr.tcW.w, recorded: c.tc.tcPr.tcPrChange?.tcPr.tcW.w, text: c.text, ins: JSON.stringify(c.tc.content ?? []).includes('"localPart":"ins"') }))),
  });
  const pkg = await WordprocessingMLPackage.load(await fixture('check40/40c-officejs.docx'));
  pkg.author = { name: 'Author A' };
  const body = await pkg.getBody();
  pkg.changeTrackingMode = 'TrackAll';
  const [d, e] = body.tables;
  d.deleteColumns(1, 1);
  e.addColumns('End', 1, [['E1D'], ['E2D']]);
  const word = (await (await WordprocessingMLPackage.load(await fixture('check40/40c-word365.docx'))).getBody()).tables;
  assert.deepEqual(records(d), records(word[0]), 'table D, the column deleted');
  assert.deepEqual(records(e), records(word[1]), 'table E, the column added');
  assert.ok(records(e).rows[0][3].ins && records(e).rows[0][3].recorded === 3008);
  const xml = await pkg.getMainDocumentPart().getXml();
  assert.doesNotMatch(xml, /<w:pPr><w:rPr><w:ins /, 'no mark insertion on the new cells, as Word');
});

test('Word check 41 (2026-10-09): the new column takes the right-hand cell\'s shading and paragraph properties, not its borders or alignment, from either side; held to the saves', async () => {
  const summary = (t) => t.rows.map((r) => r.cells.map((c) => ({
    text: c.text, w: `${c.tc.tcPr.tcW.w} ${c.tc.tcPr.tcW.type}`, fill: c.tc.tcPr.shd?.fill, vAlign: c.tc.tcPr.vAlign?.val, borders: Object.keys(c.tc.tcPr.tcBorders ?? {}).filter((k) => k !== 'TYPE_NAME' && k !== 'PARENT'),
    jc: c.paragraphs[0].p.pPr?.jc?.val, markBold: !!c.paragraphs[0].p.pPr?.rPr?.b, markItalic: !!c.paragraphs[0].p.pPr?.rPr?.i,
  })));
  for (const [original, saved, act] of [
    ['41a-properties.docx', '41a-word365.docx', (t) => t.getCell(0, 0).insertColumns('After', 1)],
    ['41b-properties.docx', '41b-word365.docx', (t) => t.getCell(0, 1).insertColumns('Before', 1)],
  ]) {
    const table = (await (await WordprocessingMLPackage.load(await fixture(`check41/${original}`))).getBody()).tables[0];
    act(table);
    const word = (await (await WordprocessingMLPackage.load(await fixture(`check41/${saved}`))).getBody()).tables[0];
    assert.deepEqual(summary(table), summary(word), original);
    assert.deepEqual(summary(word)[0][1], { text: '', w: '4513 dxa', fill: '99CCFF', vAlign: undefined, borders: [], jc: 'right', markBold: false, markItalic: true }, `${original}: the blue, right-aligned italic column, less its border and bottom alignment`);
  }
});

test('Word check 41c (2026-10-09): Office JS\'s tracked add records nothing on an AutoFit-to-contents table and the merged-edge grid on an AutoFit-to-window one; held to the save', async () => {
  const records = (t) => ({
    tblW: `${t.tbl.tblPr.tblW.w} ${t.tbl.tblPr.tblW.type}`, tblPrChange: t.tbl.tblPr.tblPrChange !== undefined, recordedColumns: t.tbl.tblGrid.tblGridChange?.tblGrid.gridCol.length,
    rows: t.rows.map((r) => r.cells.map((c) => ({ text: c.text, w: `${c.tc.tcPr.tcW.w} ${c.tc.tcPr.tcW.type}`, recorded: c.tc.tcPr.tcPrChange ? `${c.tc.tcPr.tcPrChange.tcPr.tcW?.w} ${c.tc.tcPr.tcPrChange.tcPr.tcW?.type} span${c.tc.tcPr.tcPrChange.tcPr.gridSpan?.val ?? 1}` : undefined, ins: JSON.stringify(c.tc.content ?? []).includes('"localPart":"ins"') }))),
  });
  const pkg = await WordprocessingMLPackage.load(await fixture('check41/41c-officejs.docx'));
  pkg.author = { name: 'Author A' };
  const body = await pkg.getBody();
  pkg.changeTrackingMode = 'TrackAll';
  const [f, g] = body.tables;
  f.addColumns('End', 1, [['F1D'], ['F2D']]);
  g.addColumns('End', 1, [['G1D'], ['G2D']]);
  const word = (await (await WordprocessingMLPackage.load(await fixture('check41/41c-word365.docx'))).getBody()).tables;
  assert.deepEqual(records(f), records(word[0]), 'contents: no record at all, the text inserted');
  assert.equal(records(f).tblPrChange, false);
  assert.deepEqual(records(g), records(word[1]), 'window: the records, the shares, the merged grid');
  assert.equal(records(g).recordedColumns, 7, 'the old grid\'s edges and the new grid\'s together');
  assert.deepEqual(records(g).rows[0].map((c) => c.recorded), ['1666 pct span2', '1667 pct span2', '1667 pct span2', '1 pct span1']);
  assert.deepEqual(records(g).rows[0].map((c) => c.w), ['1250 pct', '1250 pct', '1250 pct', '1250 pct']);
  const kinds = body.getTrackedChanges().map((c) => `${c.target.kind}:${c.type}`);
  assert.deepEqual(kinds, ['run:Added', 'run:Added', 'tableProperties:Formatted', 'run:Added', 'run:Added'], 'F lists its two insertions alone, as Office JS does; G its table change');
  // rejected, the window table's old widths come back and the grid folds to the old three columns plus the empty new one
  assert.ok(body.rejectAll() > 0);
  assert.deepEqual(g.rows[0].cells.map((c) => `${c.tc.tcPr.tcW.w} ${c.tc.tcPr.tcW.type}`), ['1666 pct', '1667 pct', '1667 pct', '1 pct']);
  assert.deepEqual(g.values, [['G1A', 'G1B', 'G1C', ''], ['G2A', 'G2B', 'G2C', '']]);
  assert.equal(g.tbl.tblGrid.gridCol.length, 4, 'the merged grid folded back to one column per cell');
  assert.deepEqual(f.values, [['F1A', 'F1B', 'F1C', ''], ['F2A', 'F2B', 'F2C', '']]);
});
