// Word check 40c (CR-002 section 42.5): what Office JS does to a column deleted and a column added while
// changes are tracked. Script Lab, with 40c-officejs.docx open: paste this as the script and check40.html
// as the HTML, press "Run the check", copy the JSON back, then File > Save As 40c-word365.docx.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml = (node) => new XMLSerializer().serializeToString(node)
  .replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:\w+|w:rsid\w*|w16du:dateUtc)="[^"]*"/g, '');

/** The w:tbl elements of the body, as Word's flat OPC has them now. */
async function tablesXml(context) {
  const flat = context.document.body.getOoxml();
  await context.sync();
  const doc = new DOMParser().parseFromString(flat.value, 'application/xml');
  const parts = Array.from(doc.getElementsByTagNameNS(PKG, 'part'));
  const main = parts.find((p) => p.getAttribute('pkg:name') === '/word/document.xml');
  return Array.from(main.getElementsByTagNameNS(W, 'tbl')).map(xml);
}

async function listed(context) {
  const changes = context.document.body.getTrackedChanges();
  changes.load('items/type,items/author,items/text');
  await context.sync();
  return changes.items.map((c) => ({ type: c.type, author: c.author, text: c.text }));
}

async function check() {
  const report = { host: Office.context.diagnostics, before: null, deleted: null, added: null, after: null, listed: null, errors: {} };
  await Word.run(async (context) => {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();
    report.before = await tablesXml(context);
    const tables = context.document.body.tables;
    tables.load('items');
    await context.sync();
    const [d, e] = tables.items;
    // Table 1 (D): the middle column deleted through Office JS, tracking on
    try {
      d.deleteColumns(1, 1);
      await context.sync();
      report.deleted = (await tablesXml(context))[0];
    } catch (err) { report.errors.deleteColumns = String(err && err.message || err); }
    // Table 2 (E): a column added at the end through Office JS, tracking on
    try {
      e.addColumns('End', 1, [['E1D'], ['E2D']]);
      await context.sync();
      report.added = (await tablesXml(context))[1];
    } catch (err) { report.errors.addColumns = String(err && err.message || err); }
    report.after = await tablesXml(context);
    try { report.listed = await listed(context); } catch (err) { report.errors.getTrackedChanges = String(err && err.message || err); }
  });
  out.value = JSON.stringify(report, null, 2);
}
