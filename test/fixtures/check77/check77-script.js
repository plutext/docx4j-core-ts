// Word check 77: does Word keep w:ins and w:del written inside a package given to insertOoxml, with the package's author and
// date, or does it mark the replacement itself? Section 14 of the portfolio's Word-LLM-interaction-models.docx: this decides
// whether "always tracked, under the agent's name" can hold in Word. Run through the Word bridge's check verb, or in Script Lab
// with 77-tracked-insertooxml.docx open. Each case is the file's "BEGIN Cnn mode target payload" line. #out is set once, last.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => {
  out.value = JSON.stringify({ check: 77, failed: String((e && e.stack) || e) });
}));

const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const AGENT = 'Agent Seven';
const WHEN = '2026-01-02T03:04:05Z';
const MODES = { off: 'Off', on: 'TrackAll', mine: 'TrackMineOnly' };

const packageOf = (blocks) => `<pkg:package xmlns:pkg="${PKG}">`
  + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
  + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  + '</pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + `<pkg:xmlData><w:document xmlns:w="${W}"><w:body>${blocks}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;

let nextId = 9000;
const stamp = () => `w:id="${nextId++}" w:author="${AGENT}" w:date="${WHEN}"`;
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const ins = (t) => `<w:ins ${stamp()}>${run(t)}</w:ins>`;
const del = (t) => `<w:del ${stamp()}><w:r><w:delText xml:space="preserve">${t}</w:delText></w:r></w:del>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const insertedMark = (inner) => `<w:p><w:pPr><w:rPr><w:ins ${stamp()}/></w:rPr></w:pPr>${inner}</w:p>`;

/** What a case puts in: for the word, the replacement alone; for a paragraph, its whole text, the edit in its place. */
function blocksOf(c) {
  const head = c.target === 'word' ? '' : run(`${c.id} before `);
  const tail = c.target === 'word' ? '' : run(' after');
  switch (c.payload) {
    case 'plain': return para(head + run('PLAIN') + tail);
    case 'ins': return para(head + (c.target === 'word' ? '' : run('TARGET')) + ins('ADDED') + tail);
    case 'del': return para(head + del('TARGET') + tail);
    case 'redline': return para(head + del('TARGET') + ins('REPLACED') + tail);
    case 'twoparas': return insertedMark(head + ins('FIRST')) + para(ins('SECOND') + tail);
    default: throw new Error(`no payload ${c.payload}`);
  }
}

function parseCase(text) {
  const m = /^BEGIN (C\d+) (off|on|mine) (word|content|whole|para) (\w+)$/.exec(text.trim());
  return m ? { id: m[1], mode: m[2], target: m[3], payload: m[4] } : null;
}

/** The body's XML, rsids and paragraph ids left out; authors and dates kept. */
function bodyOf(flat) {
  const m = /<w:body[\s\S]*<\/w:body>/.exec(flat);
  return m ? m[0].replace(/ w:rsid\w*="[^"]*"/g, '').replace(/ w14:(paraId|textId)="[^"]*"/g, '') : null;
}

const errorOf = (e) => ({ message: String((e && e.message) || e), code: e && e.code, debug: e && e.debugInfo ? e.debugInfo.errorLocation : undefined });
const dated = (d) => (d instanceof Date ? d.toISOString() : String(d));

async function setMode(context, mode) {
  context.document.changeTrackingMode = MODES[mode];
  await context.sync();
  context.document.load('changeTrackingMode');
  await context.sync();
  return context.document.changeTrackingMode;
}

async function listed(context, holder) {
  const changes = holder.getTrackedChanges();
  changes.load('items/type,items/author,items/date,items/text');
  await context.sync();
  return changes.items.map((c) => ({ type: c.type, author: c.author, date: dated(c.date), text: c.text }));
}

async function readBody(context) {
  const flat = context.document.body.getOoxml();
  await context.sync();
  return bodyOf(flat.value);
}

async function runCase(context, c) {
  const r = { ...c, modeRead: null, error: null };
  try {
    r.modeRead = await setMode(context, c.mode);
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text');
    await context.sync();
    const target = paragraphs.items.find((p) => p.text.startsWith(`${c.id} before`));
    if (!target) throw new Error(`no target paragraph for ${c.id}`);
    let word = null;
    if (c.target === 'word') {
      const hits = target.search('TARGET', { matchCase: true });
      hits.load('items');
      await context.sync();
      word = hits.items[0];
    }
    if (c.payload === 'self') word.insertText('WORD WROTE THIS', 'Replace');
    else if (c.payload === 'selfdelete') word.delete();
    else {
      const xml = packageOf(blocksOf(c));
      if (c.target === 'word') word.insertOoxml(xml, 'Replace');
      else if (c.target === 'content') target.getRange('Content').insertOoxml(xml, 'Replace');
      else if (c.target === 'whole') target.getRange('Whole').insertOoxml(xml, 'Replace');
      else target.insertOoxml(xml, 'Replace');
    }
    await context.sync();
  } catch (e) { r.error = errorOf(e); }
  return r;
}

/** What lies between a case's BEGIN and END lines, read once every case has run. */
async function readRegion(context, r) {
  try {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text');
    await context.sync();
    const begin = paragraphs.items.findIndex((p) => p.text.startsWith(`BEGIN ${r.id} `));
    const end = paragraphs.items.findIndex((p) => p.text.trim() === `END ${r.id}`);
    r.paragraphs = begin >= 0 && end > begin ? paragraphs.items.slice(begin + 1, end).map((p) => p.text) : null;
    r.markers = { begin, end };
    if (begin < 0 || end <= begin) return;
    const region = paragraphs.items[begin].getRange('Start').expandTo(paragraphs.items[end].getRange('End'));
    r.listed = await listed(context, region);
    try {
      const original = region.getReviewedText('Original');
      const current = region.getReviewedText('Current');
      await context.sync();
      r.original = original.value;
      r.current = current.value;
    } catch (e) { r.reviewedError = errorOf(e); }
  } catch (e) { r.readError = errorOf(e); }
}

/** The whole body replaced, tracking off and then on: last, since it takes everything else with it. */
async function bodyCases(context, report) {
  const blocks = () => para(run('BODY kept')) + para(run('BODY before ') + del('TARGET') + ins('REPLACED') + run(' after'));
  const step = async (name, act) => {
    const s = { name, error: null };
    try { await act(s); s.body = await readBody(context); s.listed = await listed(context, context.document.body); } catch (e) { s.error = errorOf(e); }
    report.bodyCases.push(s);
  };
  await step('body-off', async (s) => { s.modeRead = await setMode(context, 'off'); context.document.body.insertOoxml(packageOf(blocks()), 'Replace'); await context.sync(); });
  await step('accepted', async (s) => { s.modeRead = await setMode(context, 'off'); context.document.body.getTrackedChanges().acceptAll(); await context.sync(); });
  await step('body-on', async (s) => { s.modeRead = await setMode(context, 'on'); context.document.body.insertOoxml(packageOf(blocks()), 'Replace'); await context.sync(); });
}

async function check() {
  const report = { check: 77, agent: AGENT, when: WHEN, ranAt: new Date().toISOString(), host: Office.context.diagnostics, cases: [], bodyCases: [] };
  report.wordApi = ['1.4', '1.6'].map((v) => `${v}: ${Office.context.requirements.isSetSupported('WordApi', v)}`);
  await Word.run(async (context) => {
    context.document.load('changeTrackingMode');
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text');
    await context.sync();
    report.startMode = context.document.changeTrackingMode;
    const cases = paragraphs.items.map((p) => parseCase(p.text)).filter(Boolean);
    for (const c of cases) report.cases.push(await runCase(context, c));
    report.endOfCasesMode = await setMode(context, 'off');
    for (const r of report.cases) await readRegion(context, r);
    report.body = await readBody(context);
    report.listed = await listed(context, context.document.body);
    await bodyCases(context, report);
  });
  out.value = JSON.stringify(report);
}
