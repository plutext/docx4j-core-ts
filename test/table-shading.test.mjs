// CR-002 section 67: Office JS's shadingColor on Table, TableRow and TableCell, held to Word 365's own setters and getters.
// test/table-shading/ holds the portfolio's bridge round core-ts-table-shading (its sets -r, -2a and -3 too; bridge 958f35f,
// Word 365 16.0.20326.20318, 2026-11-17), copied unedited from bridge/portfolio/<job>/: each job's result.json, and for the
// jobs Word answered its check-result.json (the getters before and after, each set's outcome) and its save,
// <job>.ts-template.word.docx. Each job's own script, <job>.check-script.js, runs here through the Word shim on the same
// template, so the getters, the sets and the save are held to Word's. The clears raised WD-004 in Word and never answered;
// they're held to section 67.3's E1.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { unzipSync, strFromU8 } from 'fflate';
import { OpcPackage, parseXml, serializeXml } from '../dist/index.mjs';
import { Word } from '../dist/office-js/index.mjs';

const dir = new URL('./table-shading/', import.meta.url);
const read = (name) => readFileSync(new URL(name, dir));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
/** The browser's two XML classes a script uses, over this package's own parser and serializer. */
class DOMParser { parseFromString(text) { return parseXml(text); } }
class XMLSerializer { serializeToString(node) { return serializeXml(node); } }

/** The jobs Word answered and saved, and those it never answered (WD-004). */
const ANSWERED = ['ts01-gets', 'ts02-row-hex', 'ts03-row-names', 'ts05-row-merged-continuation', 'ts06-row-merged-restart',
  'ts07-cell', 'ts08-table', 'ts09-tracked'];
const CLEARS = ['ts04-row-clear', 'ts04r-row-clear', 'ts04a-row-clear-styled-own', 'ts04b-row-clear-banded-own',
  'ts04c-row-clear-plain', 'ts04d-row-clear-styled-none', 'ts04e-row-null', 'ts04f-row-clear-mixed', 'ts05b-row-merged-clear',
  'ts07b-cell-clear-styled', 'ts07c-cell-clear-banded', 'ts08b-table-clear', 'ts09b-tracked-clear'];

/** A job's script run through the Word shim on the template: its report, and the package after it. */
async function runJob(job) {
  const pkg = await OpcPackage.load(new Uint8Array(read('ts-template.docx')));
  await pkg.getBody();
  let click;
  const out = { value: '' };
  const document = { getElementById: (id) => (id === 'out' ? out : { addEventListener: (_, f) => { click = f; } }) };
  const run = new Function('document', 'Word', 'Office', 'DOMParser', 'XMLSerializer', read(`${job}.check-script.js`).toString('utf8'));
  run(document, { run: (fn) => Word.run(pkg, fn) }, { context: { diagnostics: {} } }, DOMParser, XMLSerializer);
  click();
  for (let i = 0; i < 800 && !out.value; i++) await new Promise((r) => setTimeout(r, 10));
  return { report: JSON.parse(out.value), pkg };
}

/**
 * A shading as its attributes, the fill left out beside a theme fill (Word's open rewrites that cache to the theme's
 * colour, 156082 for accent1, where the template stores FF0000), or '-' for none.
 */
function shdOf(parent) {
  const shd = parent && [...parent.childNodes].find((n) => n.nodeType === 1 && n.namespaceURI === W && n.localName === 'shd');
  if (!shd) return '-';
  const attrs = [...shd.attributes].filter((a) => !(a.localName === 'fill' && shd.getAttributeNS(W, 'themeFill')));
  return attrs.map((a) => `${a.localName}=${a.value}`).sort().join(' ');
}
const child = (e, name) => e && [...e.childNodes].find((n) => n.nodeType === 1 && n.namespaceURI === W && n.localName === name);
const children = (e, name) => [...e.childNodes].filter((n) => n.nodeType === 1 && n.namespaceURI === W && n.localName === name);

/** Each top-level table of a saved document: its own shading and change, and each cell's shading, merge and change. */
function tablesOf(docx) {
  const doc = parseXml(strFromU8(unzipSync(docx)['word/document.xml']));
  const body = doc.getElementsByTagNameNS(W, 'body')[0];
  return children(body, 'tbl').map((tbl) => {
    const tblPr = child(tbl, 'tblPr');
    const tblChange = child(tblPr, 'tblPrChange');
    return {
      shd: shdOf(tblPr),
      change: tblChange ? shdOf(child(tblChange, 'tblPr')) : undefined,
      gridChange: child(child(tbl, 'tblGrid'), 'tblGridChange') !== undefined,
      rows: children(tbl, 'tr').map((tr) => ({
        exChange: child(child(tr, 'tblPrEx'), 'tblPrExChange') !== undefined,
        cells: children(tr, 'tc').map((tc) => {
          const tcPr = child(tc, 'tcPr');
          const change = child(tcPr, 'tcPrChange');
          const vMerge = child(tcPr, 'vMerge');
          return { shd: shdOf(tcPr), merge: vMerge ? (vMerge.getAttributeNS(W, 'val') || 'continue') : '', change: change ? shdOf(child(change, 'tcPr')) : undefined };
        }),
      })),
    };
  });
}

/** The getters of a report, a table, its rows and their cells. */
const gettersOf = (phase) => phase.map((t) => ({ table: t.shadingColor, rows: t.rows.map((r) => ({ row: r.shadingColor, cells: r.cells, count: r.cellCount })) }));

/**
 * This package's report as Office JS counts a row's cells: Word leaves a vertical merge's continuation out (T4's and T5's
 * second rows list one cell), where `cells` and `cellCount` here still hold it (section 67.4).
 */
function asOfficeCounts(getters) {
  return getters.map((t, i) => ({ ...t, rows: t.rows.map((r, j) => ((i === 3 || i === 4) && j === 1
    ? { ...r, cells: r.cells.slice(0, 1), count: r.count - 1 } : r)) }));
}

test('CR-002 section 67: the round\'s files are the bridge\'s, as the manifest records', async () => {
  const manifest = JSON.parse(read('manifest.json'));
  for (const [name, hash] of Object.entries(manifest.files)) {
    assert.equal(createHash('sha1').update(read(name)).digest('hex'), hash, name);
  }
  assert.equal(Object.keys(manifest.files).length, 61);
  for (const job of [...ANSWERED, ...CLEARS]) {
    const result = JSON.parse(read(`${job}.result.json`));
    assert.equal(result.word, '16.0.20326.20318', job);
    assert.equal(result.opened, 'silent', job);
    assert.equal(result.installed.version, '958f35f', job);
  }
});

test('CR-002 section 67.3: every clear raised WD-004\'s false dialog in Word, and every colour job answered and saved', () => {
  for (const job of CLEARS) {
    const result = JSON.parse(read(`${job}.result.json`));
    assert.equal(result.check.answered, false, job);
    assert.deepEqual(result.dialogs.map((d) => d.text), ['Microsoft Word\nThere is not enough memory or disk space to update the display.'], job);
  }
  for (const job of ANSWERED) {
    const result = JSON.parse(read(`${job}.result.json`));
    assert.equal(result.check.answered, true, job);
    assert.deepEqual(result.dialogs, [], job);
    assert.deepEqual(result.problems, [], job);
    assert.equal(result.stripped !== undefined, true, job);
    assert.deepEqual(JSON.parse(read(`${job}.check-result.json`)).errors, {}, job);
  }
});

for (const job of ANSWERED) {
  test(`CR-002 section 67.3: ${job}'s script through the shim reads and writes as Word 365 did`, async () => {
    const word = JSON.parse(read(`${job}.check-result.json`));
    const { report, pkg } = await runJob(job);
    assert.deepEqual(report.errors, {});
    // the getters, before and after the sets, table by table, row by row, cell by cell
    assert.deepEqual(asOfficeCounts(gettersOf(report.before)), gettersOf(word.before), 'before');
    assert.deepEqual(asOfficeCounts(gettersOf(report.after)), gettersOf(word.after), 'after');
    // each set: done and read back, or Word's error code
    const outcome = (s) => ({ done: s.done === true, readBack: s.readBack, error: s.error?.code });
    assert.deepEqual(report.sets.map(outcome), word.sets.map(outcome));
    if (job === 'ts09-tracked') assert.equal(report.trackingMode, word.trackingMode);
    // the save: the table's own shading and each cell's changed exactly where Word's did, and to what Word's did. Word's
    // open and save normalise on their own (ts01, which sets nothing: a `fill="auto"` in an unstyled table dropped, a theme
    // fill's cache rewritten), so Word's changes are against its ts01 save, and this package's against the template
    pkg.checkLossOnSave = false;
    const ours = tablesOf(await pkg.save());
    const theirs = tablesOf(read(`${job}.ts-template.word.docx`));
    const template = tablesOf(read('ts-template.docx'));
    const baseline = tablesOf(read('ts01-gets.ts-template.word.docx'));
    const shading = (tables) => tables.flatMap((t, i) => [[`T${i + 1}`, t.shd],
      ...t.rows.flatMap((r, j) => r.cells.map((c, k) => [`T${i + 1} r${j + 1} c${k + 1}`, `${c.merge}|${c.shd}`]))]);
    const changed = (now, then) => shading(now).filter(([at, value], n) => value !== shading(then)[n][1]).map(([at, value]) => [at, value]);
    assert.deepEqual(changed(ours, template), changed(theirs, baseline));
    if (job !== 'ts09-tracked') {
      assert.ok(ours.every((t) => t.change === undefined && t.rows.every((r) => r.cells.every((c) => c.change === undefined))));
      return;
    }
    // tracked (E4): a w:tcPrChange, holding Word's old shading, on each cell whose shading Word's record changes, and a
    // w:tblPrChange where the table's does; none of Word's records of no change
    const real = (tables) => tables.map((t) => ({
      table: t.change !== undefined && t.change !== t.shd ? t.change : undefined,
      cells: t.rows.map((r) => r.cells.map((c) => (c.change !== undefined && c.change !== c.shd ? c.change : undefined))),
    }));
    assert.deepEqual(real(ours), real(theirs));
    const recorded = (tables) => tables.flatMap((t) => [t.change, ...t.rows.flatMap((r) => r.cells.map((c) => c.change))]).filter((c) => c !== undefined).length;
    const changing = (tables) => real(tables).flatMap((t) => [t.table, ...t.cells.flat()]).filter((c) => c !== undefined).length;
    assert.equal(changing(ours), 12);   // T1's row (2) and cell (1), T8's table and its 4 cells, T6's row (2), T4's row with its continuation (2)
    assert.equal(recorded(ours), changing(ours), 'every record here records a change');
    assert.ok(recorded(theirs) > changing(theirs), 'Word writes records of no change too');
    assert.ok(ours.every((t) => !t.gridChange && t.rows.every((r) => !r.exChange)));
    assert.ok(theirs.some((t) => t.gridChange) && theirs.some((t) => t.rows.some((r) => r.exChange)));
  });
}

test('CR-002 section 67.3, E4: the tracked writes accepted keep the colours, rejected give the template back', async () => {
  const template = tablesOf(read('ts-template.docx'));
  for (const verb of ['acceptAllRevisions', 'rejectAllRevisions']) {
    const { pkg } = await runJob('ts09-tracked');
    const after = tablesOf(await pkg.save());
    const changes = await pkg.getAllTrackedChanges();
    assert.ok(changes.length > 0);
    await pkg[verb]();
    const resolved = tablesOf(await pkg.save());
    const shading = (tables) => tables.map((t) => ({ shd: t.shd, cells: t.rows.map((r) => r.cells.map((c) => c.shd)) }));
    assert.deepEqual(shading(resolved), shading(verb === 'acceptAllRevisions' ? after : template), verb);
    assert.ok(resolved.every((t) => t.change === undefined && t.rows.every((r) => r.cells.every((c) => c.change === undefined))), verb);
  }
});

/** Each clear job's one set, and what E1 leaves: the target's own shading gone, and what the getter then reads. */
const CLEARED = {
  'ts04-row-clear': undefined,   // its six sets as the split jobs, each below
  'ts04r-row-clear': undefined,
  'ts04a-row-clear-styled-own': { table: 2, row: 0, reads: '#FFC000' },
  'ts04b-row-clear-banded-own': { table: 3, row: 0, reads: '#FFC000' },
  'ts04c-row-clear-plain': { table: 1, row: 0, reads: '' },
  'ts04d-row-clear-styled-none': { table: 2, row: 2, reads: '#FFC000' },
  'ts04e-row-null': { table: 2, row: 1, reads: '#FFC000' },
  'ts04f-row-clear-mixed': { table: 6, row: 0, reads: '' },
  'ts05b-row-merged-clear': { table: 5, row: 1, reads: '' },
  'ts07b-cell-clear-styled': { table: 2, cell: [0, 0], reads: '#FFC000' },
  'ts07c-cell-clear-banded': { table: 3, cell: [0, 1], reads: '#FFC000' },
  'ts08b-table-clear': { table: 7, reads: '' },
  'ts09b-tracked-clear': { table: 2, row: 0, reads: '#FFC000', tracked: true },
};

for (const [job, expect] of Object.entries(CLEARED)) {
  if (!expect) continue;
  test(`CR-002 section 67.3, E1: ${job}'s clear, which Word failed on (WD-004), takes the own shading away`, async () => {
    const { report, pkg } = await runJob(job);
    assert.deepEqual(report.sets.map((s) => ({ done: s.done, readBack: s.readBack })), [{ done: true, readBack: expect.reads }]);
    const t = report.after[expect.table - 1];
    if (expect.row !== undefined) assert.deepEqual(t.rows[expect.row].shadingColor, { value: expect.reads });
    if (expect.cell) assert.deepEqual(t.rows[expect.cell[0]].cells[expect.cell[1]], { value: expect.reads });
    pkg.checkLossOnSave = false;
    const saved = tablesOf(await pkg.save())[expect.table - 1];
    const before = tablesOf(read('ts-template.docx'))[expect.table - 1];
    const cells = expect.row !== undefined ? saved.rows[expect.row].cells
      : expect.cell ? [saved.rows[expect.cell[0]].cells[expect.cell[1]]] : saved.rows.flatMap((r) => r.cells);
    assert.ok(cells.every((c) => c.shd === '-'), 'every target cell\'s own shading gone');
    if (expect.table && expect.row === undefined && !expect.cell) assert.equal(saved.shd, '-', 'the table\'s own too');
    // the rest of the table as it was
    const others = (t) => t.rows.flatMap((r, i) => r.cells.map((c, j) => ({ i, j, shd: c.shd })))
      .filter(({ i, j }) => expect.row !== undefined ? i !== expect.row : expect.cell ? !(i === expect.cell[0] && j === expect.cell[1]) : false);
    assert.deepEqual(others(saved), others(before));
    if (expect.tracked) {
      const own = before.rows[expect.row].cells.map((c) => c.shd);
      assert.deepEqual(saved.rows[expect.row].cells.map((c) => c.change), own, 'each cell\'s old shading recorded');
    }
  });
}

test('CR-002 section 67.3: the values Word takes, and the one it refuses before anything changes', async () => {
  const pkg = await OpcPackage.load(new Uint8Array(read('ts-template.docx')));
  const body = await pkg.getBody();
  const row = body.tables[0].rows[0];
  for (const [value, reads] of [['#ffff00', '#FFFF00'], ['00b0f0', '#00B0F0'], ['LightGreen', '#90EE90'], ['grey', '#808080'],
    ['gray', '#808080'], ['DarkSlateGrey', '#2F4F4F'], ['rebeccapurple', undefined], ['#FFF', undefined], [' #FFFF00', undefined],
    [42, undefined]]) {
    if (reads === undefined) {
      const before = row.shadingColor;
      assert.throws(() => { row.shadingColor = value; }, (e) => e.code === 'InvalidArgument' && e.name === 'InvalidArgument', String(value));
      assert.equal(row.shadingColor, before, `${value}: nothing changed`);
    } else {
      row.shadingColor = value;
      assert.equal(row.shadingColor, reads, String(value));
      for (const cell of row.cells) {
        const { val, color, fill, ...rest } = cell.tc.tcPr.shd;
        assert.deepEqual({ val, color, fill }, { val: 'clear', color: reads.slice(1), fill: reads.slice(1) }, String(value));
        assert.deepEqual(Object.keys(rest).filter((k) => k !== 'TYPE_NAME' && k !== 'PARENT'), [], 'no theme attributes');
      }
    }
  }
});

test('CR-002 section 67.3: getCell at a merge\'s continuation throws ItemNotFound; the table getter reads null where cells differ', async () => {
  const pkg = await OpcPackage.load(new Uint8Array(read('ts-template.docx')));
  const body = await pkg.getBody();
  const t4 = body.tables[3];
  assert.throws(() => t4.getCell(1, 1), (e) => e.code === 'ItemNotFound');
  assert.equal(t4.getCell(0, 1).shadingColor, '');
  assert.equal(t4.rows[1].cellCount, 2, 'cellCount still counts the w:tc (section 67.4)');
  assert.equal(body.tables[5].shadingColor, null);
  assert.equal(body.tables[5].getCell(1, 0).shadingColor, '#156082', 'a theme fill as drawn, not the stored FF0000');
  assert.equal(body.tables[6].getCell(0, 0).shadingColor, '', 'the table\'s own w:tblPr/w:shd isn\'t read');
  // the row getter reads the cells Office JS counts, so a continuation's own shading doesn't make its row mixed. Inferred
  // from Word's count (ts01: one cell), unmeasured: in every case Word ran, a continuation agreed with its row
  const t5 = body.tables[4];
  t5.rows[1].cells[1].tc.tcPr.shd = { val: 'clear', color: 'auto', fill: '00B0F0' };
  assert.equal(t5.rows[1].cells[1].shadingColor, '#00B0F0');
  assert.equal(t5.rows[1].shadingColor, '');
});

test('CR-002 section 67.3, E4: a tracked cell write records once, and goes where the cell is set back', async () => {
  const pkg = await OpcPackage.load(new Uint8Array(read('ts-template.docx')));
  const body = await pkg.getBody();
  pkg.changeTrackingMode = 'TrackAll';
  const cell = body.tables[1].getCell(0, 0);   // its own yellow
  cell.shadingColor = '#00B0F0';
  const change = cell.tc.tcPr.tcPrChange;
  assert.equal(change.tcPr.shd.fill, 'FFFF00');
  cell.shadingColor = 'red';
  assert.equal(cell.tc.tcPr.tcPrChange, change, 'the first original kept');
  cell.shadingColor = '#00B0F0';
  cell.shadingColor = '#00B0F0';   // the same colour again records nothing new
  assert.equal(cell.tc.tcPr.tcPrChange, change);
  // a cell with no shading of its own, shaded and then cleared, is back as it was: the record goes
  const plain = body.tables[0].getCell(0, 0);
  plain.shadingColor = '#FFFF00';
  assert.ok(plain.tc.tcPr.tcPrChange);
  plain.shadingColor = '';
  assert.equal(plain.tc.tcPr.tcPrChange, undefined);
  // the table setter records the table's properties too
  const t8 = body.tables[7];
  t8.shadingColor = 'yellow';
  assert.equal(t8.tbl.tblPr.tblPrChange.tblPr.shd, undefined);
  assert.equal(t8.tbl.tblPr.shd.fill, 'FFFF00');
});

test('CR-002 section 67.3: Word\'s open and save, setting nothing (ts01), normalise two shadings of their own', () => {
  const template = tablesOf(read('ts-template.docx'));
  const word = tablesOf(read('ts01-gets.ts-template.word.docx'));
  const cells = (tables) => tables.map((t) => t.rows.map((r) => r.cells.map((c) => c.shd)));
  const differ = [];
  cells(template).forEach((t, i) => t.forEach((r, j) => r.forEach((shd, k) => {
    const saved = cells(word)[i][j][k];
    if (saved !== shd) differ.push([`T${i + 1} r${j + 1} c${k + 1}`, shd, saved]);
  })));
  // a fill="auto" with nothing beneath it to hide, in an unstyled table, dropped (the editor's round kept the same form
  // over a table style, section 67.1); a theme fill's stored value rewritten, which shdOf leaves out
  assert.deepEqual(differ, [['T6 r3 c1', 'color=auto val=clear fill=auto'.split(' ').sort().join(' '), '-']]);
  const raw = strFromU8(unzipSync(read('ts01-gets.ts-template.word.docx'))['word/document.xml']);
  assert.equal((raw.match(/w:fill="156082" w:themeFill="accent1"/g) ?? []).length, 2);
  assert.equal((raw.match(/w:fill="FF0000"/g) ?? []).length, 0);
});
