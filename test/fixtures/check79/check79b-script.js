// Word check 79b: how long insertOoxml takes on one paragraph against the whole body, by document size, and where the
// person's selection is left (Word-LLM-interaction-models.docx section 14; the view itself can't be judged unattended). The
// file is built for the purpose, and its first paragraph gives its size and this run's budget in seconds. Past the budget
// #out is set with what's finished, so a slow body replace still answers. #out is set once.
const out = document.getElementById('out');
const report = { check: '79b', ranAt: new Date().toISOString(), times: {}, selection: {}, stage: null };
const answer = () => { if (!out.value) out.value = JSON.stringify(report); };
document.getElementById('run').addEventListener('click', () => check().then(answer, (e) => { report.failed = String((e && e.stack) || e); answer(); }));

const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const AGENT = 'Agent Seven';
const WHEN = '2026-01-02T03:04:00Z';
const packageOf = (blocks) => `<pkg:package xmlns:pkg="${PKG}">`
  + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
  + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  + '</pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + `<pkg:xmlData><w:document xmlns:w="${W}"><w:body>${blocks}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
let nextId = 9000;
const redline = (label) => `<w:p><w:r><w:t xml:space="preserve">${label} before </w:t></w:r><w:del w:id="${nextId++}" w:author="${AGENT}" w:date="${WHEN}"><w:r><w:delText>TARGET</w:delText></w:r></w:del>`
  + `<w:ins w:id="${nextId++}" w:author="${AGENT}" w:date="${WHEN}"><w:r><w:t>REPLACED</w:t></w:r></w:ins></w:p>`;
const errorOf = (e) => ({ message: String((e && e.message) || e), code: e && e.code });

async function timed(name, fn) {
  report.stage = name;
  const t0 = performance.now();
  try { const v = await fn(); report.times[name] = Math.round(performance.now() - t0); return v; } catch (e) {
    report.times[name] = { error: errorOf(e), ms: Math.round(performance.now() - t0) };
    return undefined;
  }
}

/** Where the selection is: its text and the paragraph it starts in. */
async function selection(context, name) {
  try {
    const sel = context.document.getSelection();
    sel.load('text,isEmpty');
    const first = sel.paragraphs.getFirst();
    first.load('text');
    await context.sync();
    report.selection[name] = { isEmpty: sel.isEmpty, text: sel.text.slice(0, 40), paragraph: first.text.slice(0, 40) };
  } catch (e) { report.selection[name] = { error: errorOf(e) }; }
}

async function measure(context) {
  context.document.changeTrackingMode = 'Off';
  await context.sync();
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  await timed('load-paragraphs', async () => { paragraphs.load('items/text'); await context.sync(); });
  const items = paragraphs.items;
  report.paragraphs = items.length;
  const flat = await timed('body-getOoxml', async () => { const f = body.getOoxml(); await context.sync(); return f.value; });
  report.bodyChars = flat ? flat.length : null;
  const ids = await timed('scan-ids', async () => Array.from(flat.matchAll(/<w:p\b[^>]*?\bw14:paraId="([0-9A-Fa-f]+)"/g), (m) => m[1]));
  report.idsFound = ids ? ids.length : null;
  const mid = Math.floor(items.length / 2);
  report.idAtMiddle = ids ? ids[mid] : null;
  report.textAtMiddle = items[mid].text.slice(0, 16);
  await timed('paragraph-getOoxml', async () => { const f = items[mid].getOoxml(); await context.sync(); return f.value.length; });
  await timed('select-middle', async () => { items[mid].select(); await context.sync(); });
  await selection(context, 'before');
  report.single = [];
  for (const share of [0.1, 0.3, 0.7, 0.9]) {
    const at = Math.floor(items.length * share);
    const t0 = performance.now();
    report.stage = `paragraph-insertOoxml-${share}`;
    try {
      items[at].insertOoxml(packageOf(redline(`AT ${share}`)), 'Replace');
      await context.sync();
      report.single.push({ share, ms: Math.round(performance.now() - t0) });
    } catch (e) { report.single.push({ share, error: errorOf(e) }); }
  }
  await selection(context, 'after-elsewhere');
  await timed('paragraph-insertOoxml-selected', async () => { items[mid].insertOoxml(packageOf(redline('AT THE SELECTION')), 'Replace'); await context.sync(); });
  await selection(context, 'after-selected');
  await timed('body-insertOoxml', async () => { body.insertOoxml(flat, 'Replace'); await context.sync(); });
  await selection(context, 'after-body');
  await timed('count-after', async () => { const p = body.paragraphs; p.load('items'); await context.sync(); report.paragraphsAfter = p.items.length; });
  report.stage = 'finished';
}

async function check() {
  report.host = Office.context.diagnostics;
  let budget = 120;
  await Word.run(async (context) => {
    const first = context.document.body.paragraphs.getFirst();
    first.load('text');
    await context.sync();
    const m = /budget (\d+)/.exec(first.text);
    if (m) budget = Number(m[1]);
    report.heading = first.text;
    report.budget = budget;
    await Promise.race([measure(context), new Promise((resolve) => { setTimeout(() => { report.overBudget = true; answer(); resolve(); }, budget * 1000); })]);
  });
}
