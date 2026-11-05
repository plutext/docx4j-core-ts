// Word check 80 (CR-002 section 66): does Word 365 keep the ids of a file this package has edited? Each input is Word's own save
// with rows, a table or columns added here, the new paragraphs and rows holding a w14:paraId and no w14:textId. This reads the ids
// Word holds at open, for comparing with the file's. Nothing is written. #out is set once, last.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => {
  out.value = JSON.stringify({ check: 80, failed: String((e && e.stack) || e) });
}));

const idsOf = (flat, tag) => Array.from(flat.matchAll(new RegExp(`<w:${tag}\\b[^>]*?\\bw14:paraId="([0-9A-Fa-f]+)"`, 'g')), (m) => m[1]);
const bodyOf = (flat) => { const m = /<w:body[\s\S]*<\/w:body>/.exec(flat); return m ? m[0].replace(/ w:rsid\w*="[^"]*"/g, '') : null; };

async function check() {
  const report = { check: 80, ranAt: new Date().toISOString(), host: Office.context.diagnostics };
  await Word.run(async (context) => {
    const flat = context.document.body.getOoxml();
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text');
    await context.sync();
    report.paragraphIds = idsOf(flat.value, 'p');
    report.rowIds = idsOf(flat.value, 'tr');
    report.texts = paragraphs.items.map((p) => p.text);
    report.body = bodyOf(flat.value);
  });
  out.value = JSON.stringify(report);
}
