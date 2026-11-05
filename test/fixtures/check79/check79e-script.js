// Word check 79e, after 79c: its run gave getByParaId AccessDenied 50 times, in every form, on plain paragraphs, then Word
// answered only "Wait until the previous call completes", so getById and getByIdOrNullObject were never reached. Here those two
// go first, few calls in all, each in a Word.run of its own, and a read after each shows whether Word still answers. A call that
// reaches a paragraph then writes through it, to show a range comes back. #out is set once, last.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => {
  out.value = JSON.stringify({ check: '79e', failed: String((e && e.stack) || e) });
}));

const errorOf = (e) => ({ message: String((e && e.message) || e), code: e && e.code });
const idsOf = (flat) => Array.from(flat.matchAll(/<w:p\b[^>]*?\bw14:paraId="([0-9A-Fa-f]+)"/g), (m) => m[1]);
const pause = (ms) => new Promise((resolve) => { setTimeout(resolve, ms); });

/** Whether Word still answers a plain read. */
async function answering() {
  try {
    await Word.run(async (context) => {
      const p = context.document.body.paragraphs.getFirst();
      p.load('text');
      await context.sync();
    });
    return true;
  } catch (e) { return errorOf(e); }
}

async function attempt(report, method, kind, value) {
  const row = { method, kind, value };
  try {
    await Word.run(async (context) => {
      const p = context.document.body.paragraphs[method](value);
      p.load('text,isNullObject');
      await context.sync();
      row.isNullObject = p.isNullObject;
      if (!p.isNullObject) {
        row.text = p.text;
        const range = p.getRange('Content');
        range.insertText(' REACHED', 'End');
        await context.sync();
        p.load('text');
        await context.sync();
        row.written = p.text;
      }
    });
  } catch (e) { row.error = errorOf(e); }
  await pause(300);
  row.stillAnswering = await answering();
  report.attempts.push(row);
}

async function check() {
  const report = { check: '79e', ranAt: new Date().toISOString(), host: Office.context.diagnostics, attempts: [] };
  await Word.run(async (context) => {
    const flat = context.document.body.getOoxml();
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text');
    await context.sync();
    report.ids = idsOf(flat.value);
    report.texts = paragraphs.items.map((p) => p.text);
  });
  try {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/uniqueLocalId');
      await context.sync();
      report.localIds = paragraphs.items.map((p) => p.uniqueLocalId);
    });
  } catch (e) { report.localIdError = errorOf(e); }
  const at = (prefix) => report.texts.findIndex((t) => t.startsWith(prefix));
  const plain = Math.max(at('P02'), 0);
  const local = report.localIds ? report.localIds[plain] : null;
  const paraId = report.ids[plain];
  if (local) {
    await attempt(report, 'getById', 'uniqueLocalId', local);
    await attempt(report, 'getByIdOrNullObject', 'uniqueLocalId', local);
  }
  await attempt(report, 'getById', 'paraId', paraId);
  await attempt(report, 'getByIdOrNullObject', 'paraId', paraId);
  await attempt(report, 'getByParaId', 'paraId', paraId);
  const cell = at('P12');
  if (cell >= 0) await attempt(report, 'getByParaId', 'paraId in a cell', report.ids[cell]);
  out.value = JSON.stringify(report);
}
