// Word check 79a: can a paragraph be reached from its w14:paraId, and do the ids last? Section 14 of the portfolio's
// Word-LLM-interaction-models.docx: the engine's block addresses would need mapping to Word ranges through that id. What
// Office JS offers is listed from its own prototypes, so a member this script doesn't know of still shows. Run on 79a (ids
// in the source) and 79n (none). Each step reads everything back; nothing is saved. #out is set once, last.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => {
  out.value = JSON.stringify({ check: '79a', failed: String((e && e.stack) || e) });
}));

const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

const packageOf = (blocks) => `<pkg:package xmlns:pkg="${PKG}">`
  + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
  + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  + '</pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + `<pkg:xmlData><w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}" mc:Ignorable="w14"><w:body>${blocks}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;

const para = (id, text) => `<w:p${id ? ` w14:paraId="${id}" w14:textId="77777777"` : ''}><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const errorOf = (e) => ({ message: String((e && e.message) || e), code: e && e.code });
const idsOf = (flat) => Array.from(flat.matchAll(/<w:p\b[^>]*?\bw14:paraId="([0-9A-Fa-f]+)"/g), (m) => m[1]);
const bodyOf = (flat) => { const m = /<w:body[\s\S]*<\/w:body>/.exec(flat); return m ? m[0].replace(/ w:rsid\w*="[^"]*"/g, '') : null; };

/** What Office JS itself has for an id: its members with id, unique or para in the name. */
function members() {
  const of = (c) => (c && c.prototype ? Object.getOwnPropertyNames(c.prototype).filter((n) => /id|unique|para/i.test(n)) : null);
  const names = {};
  for (const k of ['Paragraph', 'ParagraphCollection', 'Document', 'Range', 'Body', 'Bookmark']) names[k] = of(Word[k]);
  return names;
}

function requirementSets() {
  const has = (set, v) => { try { return Office.context.requirements.isSetSupported(set, v); } catch (e) { return String(e); } };
  const sets = {};
  for (const v of ['1.4', '1.5', '1.6', '1.7', '1.8', '1.9']) sets[`WordApi ${v}`] = has('WordApi', v);
  for (const v of ['1.1', '1.2', '1.3']) sets[`WordApiDesktop ${v}`] = has('WordApiDesktop', v);
  return sets;
}

async function paragraphsOf(context) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load('items/text');
  await context.sync();
  return paragraphs.items;
}

/** The body's ids in document order, the paragraphs Office JS lists, and the session's ids where it has them. */
async function read(context, label) {
  const s = { label };
  try {
    const flat = context.document.body.getOoxml();
    await context.sync();
    s.ids = idsOf(flat.value);
    s.body = bodyOf(flat.value);
    const items = await paragraphsOf(context);
    s.texts = items.map((p) => p.text);
    try {
      const again = context.document.body.paragraphs;
      again.load('items/uniqueLocalId');
      await context.sync();
      s.localIds = again.items.map((p) => p.uniqueLocalId);
    } catch (e) { s.localIdError = errorOf(e); }
  } catch (e) { s.error = errorOf(e); }
  return s;
}

function getCompressed() {
  return new Promise((resolve) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) { resolve({ error: res.error && res.error.message }); return; }
      const file = res.value; const parts = [];
      const finish = (error) => file.closeAsync(() => {
        if (error) { resolve({ error }); return; }
        let s = '';
        for (const bytes of parts) for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, Array.prototype.slice.call(bytes, i, i + 0x8000));
        resolve({ size: file.size, base64: btoa(s) });
      });
      const next = (i) => {
        if (i === file.sliceCount) { finish(); return; }
        file.getSliceAsync(i, (r) => { if (r.status !== Office.AsyncResultStatus.Succeeded) { finish(r.error && r.error.message); return; } parts.push(r.value.data); next(i + 1); });
      };
      next(0);
    });
  });
}

const find = (items, label) => {
  const p = items.find((x) => x.text.startsWith(label));
  if (!p) throw new Error(`no paragraph ${label}`);
  return p;
};

/** Each edit is on its own paragraph, so a later read shows what became of its id and of its neighbours'. */
const EDITS = [
  ['text-at-end', async (context, items) => { find(items, 'P02').insertText(' more', 'End'); }],
  ['package-own-id', async (context, items) => { find(items, 'P03').insertOoxml(packageOf(para('10000003', 'P03 replaced, own id')), 'Replace'); }],
  ['package-no-id', async (context, items) => { find(items, 'P04').insertOoxml(packageOf(para(null, 'P04 replaced, no id')), 'Replace'); }],
  ['content-own-id', async (context, items) => { find(items, 'P05').getRange('Content').insertOoxml(packageOf(para('10000005', 'P05 replaced, own id')), 'Replace'); }],
  ['read-and-put-back', async (context, items) => {
    const p = find(items, 'P06');
    const flat = p.getOoxml();
    await context.sync();
    p.insertOoxml(flat.value, 'Replace');
  }],
  ['paragraph-after', async (context, items) => { find(items, 'P07').insertParagraph('NEW after P07', 'After'); }],
  ['split', async (context, items) => {
    const hits = find(items, 'P08').search('SPLIT', { matchCase: true });
    hits.load('items');
    await context.sync();
    hits.items[0].insertText('\n', 'Start');
  }],
  ['deleted', async (context, items) => { find(items, 'P09').delete(); }],
  ['bold', async (context, items) => { find(items, 'P10').font.bold = true; }],
  ['tracked-text', async (context, items) => {
    context.document.changeTrackingMode = 'TrackAll';
    await context.sync();
    find(items, 'P11').insertText(' tracked', 'End');
    await context.sync();
    context.document.changeTrackingMode = 'Off';
  }],
];

async function reach(context, report) {
  const items = await paragraphsOf(context);
  report.perParagraph = [];
  for (let i = 0; i < items.length; i++) {
    const r = { index: i, text: items[i].text };
    try {
      const flat = items[i].getOoxml();
      await context.sync();
      r.ids = idsOf(flat.value);
    } catch (e) { r.error = errorOf(e); }
    report.perParagraph.push(r);
  }
  try {
    const again = context.document.body.paragraphs;
    again.load('items/uniqueLocalId');
    await context.sync();
    const local = again.items[Math.min(3, again.items.length - 1)].uniqueLocalId;
    const found = context.document.getParagraphByUniqueLocalId(local);
    found.load('text');
    await context.sync();
    report.byLocalId = { local, text: found.text, expected: items[Math.min(3, items.length - 1)].text };
  } catch (e) { report.byLocalId = { error: errorOf(e) }; }
}

async function check() {
  const report = { check: '79a', ranAt: new Date().toISOString(), host: Office.context.diagnostics, members: members(), sets: requirementSets(), steps: [] };
  report.fileAtOpen = await getCompressed();
  await Word.run(async (context) => {
    report.steps.push(await read(context, 'open'));
    report.steps.push(await read(context, 'read-again'));
    await reach(context, report);
    for (const [label, act] of EDITS) {
      const step = { label };
      try {
        const items = await paragraphsOf(context);
        await act(context, items);
        await context.sync();
      } catch (e) { step.error = errorOf(e); }
      report.steps.push({ ...(await read(context, label)), editError: step.error });
    }
  });
  report.fileEdited = await getCompressed();
  await Word.run(async (context) => {
    const step = {};
    try {
      const flat = context.document.body.getOoxml();
      await context.sync();
      context.document.body.insertOoxml(flat.value, 'Replace');
      await context.sync();
    } catch (e) { step.error = errorOf(e); }
    report.steps.push({ ...(await read(context, 'body-put-back')), editError: step.error });
  });
  out.value = JSON.stringify(report);
}
