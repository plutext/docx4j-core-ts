// Word check 41c (CR-002 section 42.6): what Office JS writes and lists for a column added, tracking on, to an
// AutoFit-to-contents table and an AutoFit-to-window one. Script Lab, with 41c-officejs.docx open: paste this
// as the script and check41.html as the HTML, press "Run the check", copy the JSON back, then Save As 41c-word365.docx.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml = (node) => new XMLSerializer().serializeToString(node)
  .replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*|w16du:dateUtc)="[^"]*"/g, '');

async function tablesXml(context) {
  const flat = context.document.body.getOoxml();
  await context.sync();
  const doc = new DOMParser().parseFromString(flat.value, 'application/xml');
  const main = Array.from(doc.getElementsByTagNameNS(PKG, 'part')).find((p) => p.getAttribute('pkg:name') === '/word/document.xml');
  return Array.from(main.getElementsByTagNameNS(W, 'tbl')).map(xml);
}

async function listed(context) {
  const changes = context.document.body.getTrackedChanges();
  changes.load('items/type,items/author,items/text');
  await context.sync();
  return changes.items.map((c) => ({ type: c.type, author: c.author, text: c.text }));
}

async function check() {
  const report = { host: Office.context.diagnostics, before: null, after: null, listed: null, errors: {} };
  await Word.run(async (context) => {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();
    report.before = await tablesXml(context);
    const tables = context.document.body.tables;
    tables.load('items');
    await context.sync();
    const [f, g] = tables.items;
    try { f.addColumns('End', 1, [['F1D'], ['F2D']]); await context.sync(); } catch (err) { report.errors.contents = String(err && err.message || err); }
    try { g.addColumns('End', 1, [['G1D'], ['G2D']]); await context.sync(); } catch (err) { report.errors.window = String(err && err.message || err); }
    report.after = await tablesXml(context);
    try { report.listed = await listed(context); } catch (err) { report.errors.getTrackedChanges = String(err && err.message || err); }
  });
  out.value = JSON.stringify(report, null, 2);
}
