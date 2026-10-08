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

test('the new cell takes the neighbour\'s properties, less its span and merges, and its width in the neighbour\'s form; fixed mode copies the width and grows the table', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  await body.insertXml(`<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid>`
    + `<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>`
    + `<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`, 'End');
  const table = body.tables[0];
  assert.equal(table.layoutMode, 'fixed');
  const [[yellow]] = table.getCell(0, 0).insertColumns('After', 1);
  assert.equal(yellow.tc.tcPr.shd.fill, 'FFFF00', 'the shading continues');
  assert.equal(yellow.tc.tcPr.vAlign.val, 'center');
  assert.equal(yellow.tc.tcPr.vMerge, undefined);
  assert.deepEqual(shape(table), { tblW: '8000 dxa', grid: [2000, 2000, 4000], rows: [['2000 dxa', '2000 dxa', '4000 dxa'], ['2000 dxa', '2000 dxa', '4000 dxa']] });
  // after the merged column: the new cells take the merged cell's width, not its vMerge
  const [[afterMerge], [belowMerge]] = table.getCell(0, 2).insertColumns('After', 1);
  assert.equal(afterMerge.tc.tcPr.vMerge, undefined);
  assert.equal(belowMerge.tc.tcPr.vMerge, undefined);
  assert.deepEqual(shape(table).grid, [2000, 2000, 4000, 4000]);
  assert.equal(shape(table).tblW, '12000 dxa');
});

test('a cell spanning the boundary grows instead of a cell being inserted; a row\'s skipped columns grow', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  await body.insertXml(`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>`
    + `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr>`
    + `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr><w:tc><w:p><w:r><w:t>d</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>e</w:t></w:r></w:p></w:tc></w:tr>`
    + `<w:tr><w:tc><w:p><w:r><w:t>f</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>g</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>h</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`, 'End');
  const table = body.tables[0];
  // after column 0: inside the first row's span, at the second row's skipped column's edge, before g in the third
  const added = table.getCell(2, 0).insertColumns('After', 1, [['-'], ['-'], ['new']]);
  assert.deepEqual(added.map((cells) => cells.map((c) => c.text)), [[], [], ['new']]);
  assert.equal(table.getCell(0, 0).tc.tcPr.gridSpan.val, 3, 'the spanning cell grew');
  assert.deepEqual(table.values, [['wide', 'c'], ['d', 'e'], ['f', 'new', 'g', 'h']]);
  const skipped = table.rows[1].tr.trPr.cnfStyleOrDivIdOrGridBefore.find((i) => i.name.localPart === 'gridBefore');
  assert.equal(skipped.value.val, 2, 'the skipped columns grew');
  assert.deepEqual(shape(table).grid, [3000, 3000, 3000, 3000]);
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

test('tracked: a column added is the table change Word records plus the new cells\' text inserted; accepted the records go, rejected the recorded table comes back and the column stays, empty', async () => {
  const make = async () => {
    const pkg = await WordprocessingMLPackage.createPackage();
    pkg.author = { name: 'Ada' };
    pkg.trackedChangeDate = DATE;
    const body = pkg.body;
    const table = body.insertTable(2, 2, 'End', [['a', 'b'], ['c', 'd']]);
    await body.insertXml('<w:p><w:r><w:t>after</w:t></w:r></w:p>', 'End');
    pkg.changeTrackingMode = 'TrackAll';
    return { pkg, body, table };
  };
  const { pkg, body, table } = await make();
  const before = JSON.parse(JSON.stringify(shape(table)));
  table.getCell(0, 0).insertColumns('After', 1, [['new'], ['']]);
  assert.deepEqual(table.values, [['a', 'new', 'b'], ['c', '', 'd']]);
  const tbl = table.tbl;
  assert.equal(tbl.tblPr.tblPrChange.author, 'Ada');
  assert.deepEqual(tbl.tblGrid.tblGridChange.tblGrid.gridCol.map((c) => c.w), before.grid, 'the old grid recorded');
  assert.equal(tbl.tblGrid.gridCol.length, 3);
  const cells = table.rows.flatMap((r) => r.cells);
  assert.deepEqual(cells.map((c) => c.tc.tcPr?.tcPrChange !== undefined), [true, false, true, true, false, true], 'every cell that was there records its properties');
  assert.ok(table.getCell(0, 1).paragraphs[0].p.pPr.rPr.ins, 'the new cell\'s paragraph mark is an insertion');
  const xml = await pkg.getMainDocumentPart().getXml();
  assert.match(xml, /<w:tblPrChange w:id="\d+"[^>]*w:author="Ada"[^>]*><w:tblPr>/);
  assert.match(xml, /<w:tc><w:tcPr><w:tcW [^>]*\/><\/w:tcPr><w:p><w:pPr><w:rPr><w:ins [^>]*\/><\/w:rPr><\/w:pPr><\/w:p><\/w:tc>/, 'an empty new cell: an empty paragraph, its mark inserted');
  assert.match(xml, /<w:tblGridChange w:id="\d+"><w:tblGrid><w:gridCol w:w="4513"\/><w:gridCol w:w="4514"\/><\/w:tblGrid><\/w:tblGridChange>/);
  assert.match(xml, /<w:ins [^>]*><w:r><w:t>new<\/w:t><\/w:r><\/w:ins>/);
  assert.doesNotMatch(xml, /w:cellIns/);
  const changes = body.getTrackedChanges();
  assert.ok(changes.some((c) => c.target.kind === 'tableProperties' && c.type === 'Formatted'), changes.map((c) => c.target.kind).join(','));
  assert.ok(changes.some((c) => c.type === 'Added' && c.text.startsWith('new')), 'the new cell\'s text is an insertion (grouped with its mark)');

  // accepted: the records go, the column and its text stay
  assert.ok(body.acceptAll() > 0);
  assert.deepEqual(table.values, [['a', 'new', 'b'], ['c', '', 'd']]);
  assert.doesNotMatch(await pkg.getMainDocumentPart().getXml(), /PrChange|GridChange|<w:ins /);

  // rejected: the recorded properties, grid and cells come back; the column stays, empty (Word's Reject All, section 29)
  const fresh = await make();
  fresh.table.getCell(0, 0).insertColumns('After', 1, [['new'], ['']]);
  assert.ok(fresh.body.rejectAll() > 0);
  assert.deepEqual(fresh.table.values, [['a', '', 'b'], ['c', '', 'd']]);
  assert.deepEqual(fresh.table.tbl.tblGrid.gridCol.map((c) => c.w), before.grid);
  assert.doesNotMatch(await fresh.pkg.getMainDocumentPart().getXml(), /PrChange|GridChange|<w:ins /);
});
