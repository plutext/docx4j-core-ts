// core-ts table shading check ts06-row-merged-restart (CR-002 section 67): TableRow.shadingColor on the row a merge starts in. Office JS's shadingColor setters and getters on
// ts-template.docx's ten tables, the getters read before and after, each table's XML as Word serialises the body; Word then saves
// the copy (the job's formats). Built by docx4j-core-ts's test/table-shading/build.mjs, 2026-11-17.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const JOB = {"name":"ts06-row-merged-restart","about":"TableRow.shadingColor on the row a merge starts in","track":false,"sets":[{"t":4,"row":0,"v":"#00B0F0"}]};

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
