// CR-002 section 67's Office JS round, core-ts-table-shading: Word 365's own shadingColor setters and getters on TableRow,
// TableCell and Table, through the portfolio's Word bridge (the `check` verb, with `formats: ["docx"]` so Word saves the copy
// after the script's answer). Jason, 2026-11-17: "yes, shading on cells and tables as well as rows", and "yes" to running it.
//   node test/table-shading/build.mjs test/table-shading/base/shd1-style-forms-word365.docx <an empty directory>
// The base is the editor's shd1-style-forms-word365.docx (its packages/editor-model/test/fixtures/row-shading/ at a99c6d1, Word
// 365's own save, compatibility mode 15, the styles ShadedTable and BandedTable, both orange FFC000; SHA-256 6f3dc4ce...). Its
// document.xml is replaced by the ten tables below; every other part is kept as Word wrote it. Writes ts-template.docx and the
// staging set: core-ts-table-shading/<job>/ (job.json, check.html, check-script.js, the template), the order and sha256 files.
import { readFileSync, writeFileSync, mkdirSync, existsSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { createHash } from 'node:crypto';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

const [base, out] = process.argv.slice(2);
if (!base || !out || !existsSync(out) || readdirSync(out).length > 0) {
  console.error('usage: node test/table-shading/build.mjs <base.docx> <an empty directory>');
  process.exit(1);
}
const SET = 'core-ts-table-shading';
const TEMPLATE = 'ts-template.docx';

// ------------------------------------------------------------------ the template

const Y = '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>';
const B = '<w:shd w:val="clear" w:color="auto" w:fill="00B0F0"/>';
// accent1 is 156082 in the base's theme; the stored fill says red, so a getter reading the stored value reads #FF0000
const THEME = '<w:shd w:val="clear" w:color="auto" w:fill="FF0000" w:themeFill="accent1"/>';
const AUTO = '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>';
const BORDERS = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  .map((e) => `<w:${e} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
const LOOK = '<w:tblLook w:val="0400" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>';

const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
/** A cell: its text, its own w:shd (or none) and its w:vMerge (restart, continue or none). */
const cell = (text, shd = '', merge = '') => {
  const vMerge = merge === 'restart' ? '<w:vMerge w:val="restart"/>' : merge === 'continue' ? '<w:vMerge/>' : '';
  // a merge's continuation holds an empty paragraph, as Word writes one
  return `<w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/>${vMerge}${shd}</w:tcPr>${merge === 'continue' ? '<w:p/>' : para(text)}</w:tc>`;
};
/** A table of two columns: `rows` is one array per row of [shd, merge] for its two cells; `style` a table style or none. */
const table = (n, rows, { style, shd = '' } = {}) => {
  const tblPr = style
    ? `<w:tblPr><w:tblStyle w:val="${style}"/><w:tblW w:w="0" w:type="auto"/>${shd}${LOOK}</w:tblPr>`
    : `<w:tblPr><w:tblW w:w="0" w:type="auto"/>${BORDERS}${shd}${LOOK}</w:tblPr>`;
  const trs = rows.map((cells, r) => `<w:tr>${cells.map(([s, m], c) => cell(`T${n} r${r + 1} c${c + 1}`, s, m)).join('')}</w:tr>`);
  return `<w:tbl>${tblPr}<w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid>${trs.join('')}</w:tbl>`;
};
const none = ['', ''];
const TABLES = [
  ['T1 plain: no style, no shading.', table(1, [[none, none], [none, none], [none, none], [none, none]])],
  ['T2 styled: ShadedTable (orange); rows 1 and 2 hold their own yellow, row 3 none.',
    table(2, [[[Y], [Y]], [[Y], [Y]], [none, none]], { style: 'ShadedTable' })],
  ['T3 banded: BandedTable, odd rows orange; row 1 holds its own yellow.',
    table(3, [[[Y], [Y]], [none, none], [none, none], [none, none]], { style: 'BandedTable' })],
  ['T4 merged: column 2 merged down rows 1 and 2; no shading.',
    table(4, [[none, ['', 'restart']], [none, ['', 'continue']], [none, none]])],
  ['T5 merged and shaded: the merge\'s restart yellow, its continuation none.',
    table(5, [[none, [Y, 'restart']], [none, ['', 'continue']], [none, none]])],
  ['T6 mixed: row 1 yellow and blue; row 2 themeFill accent1 over a stored red; row 3 fill auto and none.',
    table(6, [[[Y], [B]], [[THEME], [THEME]], [[AUTO], none]])],
  ['T7 table shading: the table\'s own w:shd yellow; row 2 cell 2 its own blue.',
    table(7, [[none, none], [none, [B]]], { shd: Y })],
  ['T8 plain: for the table setter.', table(8, [[none, none], [none, none]])],
  ['T9 styled: ShadedTable; row 1 cell 1 its own yellow.', table(9, [[[Y], none], [none, none]], { style: 'ShadedTable' })],
  ['T10 table shading: as T7.', table(10, [[none, none], [none, [B]]], { shd: Y })],
];

const parts = unzipSync(readFileSync(base));
const documentXml = strFromU8(parts['word/document.xml']);
const root = documentXml.match(/^[\s\S]*?<w:document\b[^>]*>/)[0];
const sectPr = documentXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)[0];
const body = [para('core-ts table shading round (CR-002 section 67). Each table follows its label.'),
  ...TABLES.flatMap(([label, tbl]) => [para(label), tbl]), para('After the tables.')].join('');
parts['word/document.xml'] = strToU8(`${root}<w:body>${body}${sectPr}</w:body></w:document>`);
const template = zipSync(parts, { level: 6, mtime: new Date('2026-11-17T00:00:00Z') });
writeFileSync(join(out, TEMPLATE), template);

// ------------------------------------------------------------------ the jobs

// Each set names its target: `t` the table (1-based, in body order), and `row` (0-based), `cell` [row, cell] (0-based,
// Office JS's getCell) or `table`; `v` the value given to shadingColor.
const SETS = {};
SETS[SET] = [
  ['ts01-gets', 'every getter, nothing set', false, []],
  ['ts02-row-hex', 'TableRow.shadingColor given hex forms', false, [
    { t: 1, row: 0, v: '#FFFF00' }, { t: 1, row: 1, v: 'FFFF00' }, { t: 1, row: 2, v: '#00b0f0' }, { t: 1, row: 3, v: '#123456' },
    { t: 6, row: 1, v: '#FFFF00' }]],
  ['ts03-row-names', 'TableRow.shadingColor given names', false, [
    { t: 1, row: 0, v: 'yellow' }, { t: 1, row: 1, v: 'Red' }, { t: 1, row: 2, v: 'lightgreen' }, { t: 1, row: 3, v: 'nonsense' }]],
  ['ts04-row-clear', 'TableRow.shadingColor cleared', false, [
    { t: 2, row: 0, v: '' }, { t: 3, row: 0, v: '' }, { t: 1, row: 0, v: '' }, { t: 2, row: 2, v: '' }, { t: 2, row: 1, v: null },
    { t: 6, row: 0, v: '' }]],
  ['ts05-row-merged-continuation', 'TableRow.shadingColor on a row a merge reaches from above', false, [
    { t: 4, row: 1, v: '#FFFF00' }, { t: 5, row: 1, v: '' }]],
  ['ts06-row-merged-restart', 'TableRow.shadingColor on the row a merge starts in', false, [{ t: 4, row: 0, v: '#00B0F0' }]],
  ['ts07-cell', 'TableCell.shadingColor', false, [
    { t: 1, cell: [0, 0], v: '#FFFF00' }, { t: 1, cell: [1, 1], v: 'yellow' }, { t: 2, cell: [0, 0], v: '' },
    { t: 4, cell: [1, 1], v: '#00B0F0' }, { t: 6, cell: [1, 0], v: '#FFFF00' }, { t: 3, cell: [0, 1], v: '' }]],
  ['ts08-table', 'Table.shadingColor', false, [
    { t: 8, table: true, v: '#FFFF00' }, { t: 9, table: true, v: '#00B0F0' }, { t: 7, table: true, v: '' },
    { t: 10, table: true, v: '#00B0F0' }, { t: 6, table: true, v: '#FFFF00' }]],
  ['ts09-tracked', 'the setters with change tracking on', true, [
    { t: 1, row: 0, v: '#FFFF00' }, { t: 1, cell: [3, 1], v: '#00B0F0' }, { t: 2, row: 0, v: '' }, { t: 8, table: true, v: '#FFFF00' },
    { t: 6, row: 1, v: '#FFFF00' }, { t: 4, row: 1, v: '#FFFF00' }]],
];
// The rerun of ts04 to ts09 (2026-11-17). ts01 to ts03 ran; ts04 stopped about a second after its script began, Word showing
// "There is not enough memory or disk space to update the display." with nothing to press it, and the rest didn't run. So
// each clearing value ("" or null) is a job of its own here, and the other jobs hold none, so that such a dialog costs one job.
// Jason's plan (the VM session, VM 305): ts04 run again as it was first (the portfolio's set core-ts-table-shading-r); these six
// only if the dialog comes again (-2a); then ts05 to ts09 (-3). The two sets were one, -2, until then; the jobs are the same.
const SPLIT = [
  ['ts04a-row-clear-styled-own', 'TableRow.shadingColor "" on a styled row holding its own yellow', false, [{ t: 2, row: 0, v: '' }]],
  ['ts04b-row-clear-banded-own', 'TableRow.shadingColor "" on a banded row holding its own yellow', false, [{ t: 3, row: 0, v: '' }]],
  ['ts04c-row-clear-plain', 'TableRow.shadingColor "" on a plain row with no shading', false, [{ t: 1, row: 0, v: '' }]],
  ['ts04d-row-clear-styled-none', 'TableRow.shadingColor "" on a styled row with none of its own', false, [{ t: 2, row: 2, v: '' }]],
  ['ts04e-row-null', 'TableRow.shadingColor null on a styled row holding its own yellow', false, [{ t: 2, row: 1, v: null }]],
  ['ts04f-row-clear-mixed', 'TableRow.shadingColor "" on a row of yellow and blue', false, [{ t: 6, row: 0, v: '' }]],
  ['ts05-row-merged-continuation', 'TableRow.shadingColor on a row a merge reaches from above', false, [{ t: 4, row: 1, v: '#FFFF00' }]],
  ['ts05b-row-merged-clear', 'TableRow.shadingColor "" on a row a yellow merge reaches from above', false, [{ t: 5, row: 1, v: '' }]],
  ['ts06-row-merged-restart', 'TableRow.shadingColor on the row a merge starts in', false, [{ t: 4, row: 0, v: '#00B0F0' }]],
  ['ts07-cell', 'TableCell.shadingColor', false, [
    { t: 1, cell: [0, 0], v: '#FFFF00' }, { t: 1, cell: [1, 1], v: 'yellow' }, { t: 4, cell: [1, 1], v: '#00B0F0' },
    { t: 6, cell: [1, 0], v: '#FFFF00' }]],
  ['ts07b-cell-clear-styled', 'TableCell.shadingColor "" on a styled cell holding its own yellow', false, [{ t: 2, cell: [0, 0], v: '' }]],
  ['ts07c-cell-clear-banded', 'TableCell.shadingColor "" on a banded cell holding its own yellow', false, [{ t: 3, cell: [0, 1], v: '' }]],
  ['ts08-table', 'Table.shadingColor', false, [
    { t: 8, table: true, v: '#FFFF00' }, { t: 9, table: true, v: '#00B0F0' }, { t: 10, table: true, v: '#00B0F0' },
    { t: 6, table: true, v: '#FFFF00' }]],
  ['ts08b-table-clear', 'Table.shadingColor "" on a table holding its own yellow', false, [{ t: 7, table: true, v: '' }]],
  ['ts09-tracked', 'the setters with change tracking on', true, [
    { t: 1, row: 0, v: '#FFFF00' }, { t: 1, cell: [3, 1], v: '#00B0F0' }, { t: 8, table: true, v: '#FFFF00' },
    { t: 6, row: 1, v: '#FFFF00' }, { t: 4, row: 1, v: '#FFFF00' }]],
  ['ts09b-tracked-clear', 'TableRow.shadingColor "" with change tracking on', true, [{ t: 2, row: 0, v: '' }]],
];
SETS[`${SET}-2a`] = SPLIT.filter(([name]) => name.startsWith('ts04'));
SETS[`${SET}-3`] = SPLIT.filter(([name]) => !name.startsWith('ts04'));

const CHECK_HTML = '<button id="run">Run the check</button>\n<p>Copy this back:</p>\n'
  + '<textarea id="out" rows="30" style="width: 100%; font-family: monospace;"></textarea>\n';

const COMMON = String.raw`
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const tidy = (s) => s.replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:paraId|w14:textId|w:rsid\w*|w16du:dateUtc)="[^"]*"/g, '');
const err = (e) => ({ name: e && e.name, code: e && e.code, message: String((e && e.message) || e),
  debugInfo: e && e.debugInfo ? { code: e.debugInfo.code, message: e.debugInfo.message, errorLocation: e.debugInfo.errorLocation } : undefined });

async function tablesOf(context) {
  const tables = context.document.body.tables;
  tables.load('items');
  await context.sync();
  if (tables.items.length !== 10) throw new Error('tables: ' + tables.items.length);
  return tables.items;
}

// One property read in a sync of its own, so that one refusal doesn't hide the rest.
async function read(context, object, property) {
  try { object.load(property); await context.sync(); return { value: object[property] }; } catch (e) { return { error: err(e) }; }
}

// Every getter: each table's, each row's and each cell's, one at a time.
async function getters(context) {
  const result = [];
  const tables = await tablesOf(context);
  for (let i = 0; i < tables.length; i++) {
    const t = tables[i];
    const r = { table: i + 1 };
    const values = await read(context, t, 'values');
    r.first = values.error ? values : values.value[0][0];
    r.shadingColor = await read(context, t, 'shadingColor');
    try {
      t.rows.load('items/cellCount');
      await context.sync();
      r.rows = [];
      for (const row of t.rows.items) {
        const rr = { cellCount: row.cellCount, shadingColor: await read(context, row, 'shadingColor'), cells: [] };
        try {
          row.cells.load('items');
          await context.sync();
          for (const c of row.cells.items) rr.cells.push(await read(context, c, 'shadingColor'));
        } catch (e) { rr.cellsError = err(e); }
        r.rows.push(rr);
      }
    } catch (e) { r.rowsError = err(e); }
    result.push(r);
  }
  return result;
}

// Each top-level table's XML as Word serialises the body, tidied.
async function tableXml(context) {
  const flat = context.document.body.getOoxml();
  await context.sync();
  const doc = new DOMParser().parseFromString(flat.value, 'application/xml');
  const main = Array.from(doc.getElementsByTagNameNS(PKG, 'part')).find((p) => p.getAttribute('pkg:name') === '/word/document.xml');
  const body = main && main.getElementsByTagNameNS(W, 'body')[0];
  if (!body) return null;
  return Array.from(body.childNodes).filter((n) => n.nodeType === 1 && n.localName === 'tbl')
    .map((n) => tidy(new XMLSerializer().serializeToString(n)));
}

async function targetOf(context, set) {
  const t = (await tablesOf(context))[set.t - 1];
  if (set.table) return t;
  if (set.row !== undefined) { t.rows.load('items'); await context.sync(); return t.rows.items[set.row]; }
  return t.getCell(set.cell[0], set.cell[1]);
}

async function check() {
  const report = { job: JOB.name, about: JOB.about, host: Office.context.diagnostics, track: JOB.track, before: null, sets: [],
    after: null, xml: null, errors: {} };
  await Word.run(async (context) => {
    try { report.before = await getters(context); } catch (e) { report.errors.before = err(e); }
    if (JOB.track) {
      try {
        context.document.changeTrackingMode = 'TrackAll';
        context.document.load('changeTrackingMode');
        await context.sync();
        report.trackingMode = context.document.changeTrackingMode;
      } catch (e) { report.errors.track = err(e); }
    }
    for (const set of JOB.sets) {
      const s = { set };
      try {
        const target = await targetOf(context, set);
        target.shadingColor = set.v;
        await context.sync();
        s.done = true;
        try { target.load('shadingColor'); await context.sync(); s.readBack = target.shadingColor; } catch (e) { s.readBackError = err(e); }
      } catch (e) { s.error = err(e); }
      report.sets.push(s);
    }
    try { report.after = await getters(context); } catch (e) { report.errors.after = err(e); }
    try { report.xml = await tableXml(context); } catch (e) { report.errors.xml = err(e); }
  });
  out.value = JSON.stringify(report, null, 1);
}
`;

const scriptOf = ([name, about, track, sets]) => `// core-ts table shading check ${name} (CR-002 section 67): ${about}. Office JS's shadingColor setters and getters on
// ts-template.docx's ten tables, the getters read before and after, each table's XML as Word serialises the body; Word then saves
// the copy (the job's formats). Built by docx4j-core-ts's test/table-shading/build.mjs, 2026-11-17.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const JOB = ${JSON.stringify({ name, about, track, sets })};
${COMMON}`;

console.log(`${TEMPLATE} ${createHash('sha256').update(template).digest('hex')}`);
for (const [set, jobs] of Object.entries(SETS)) {
  const sums = [];
  const put = (rel, bytes) => {
    const path = join(out, set, rel);
    mkdirSync(join(path, '..'), { recursive: true });
    writeFileSync(path, bytes);
    sums.push(`${createHash('sha256').update(bytes).digest('hex')}  ${set}/${rel}`);
  };
  for (const job of jobs) {
    const [name] = job;
    put(`${name}/check.html`, strToU8(CHECK_HTML));
    put(`${name}/check-script.js`, strToU8(scriptOf(job)));
    put(`${name}/${TEMPLATE}`, template);
    put(`${name}/job.json`, strToU8(JSON.stringify({ verb: 'check', input: TEMPLATE, checkHtml: 'check.html', checkScript: 'check-script.js',
      formats: ['docx'], timeoutSeconds: 120 })));
  }
  writeFileSync(join(out, `${set}.order.txt`), jobs.map(([name]) => name).join('\n') + '\n');
  writeFileSync(join(out, `${set}.sha256`), sums.join('\n') + '\n');
  console.log(`${jobs.length} jobs in ${set}`);
}
