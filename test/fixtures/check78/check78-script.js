// Word check 78: what getFileAsync gives against what SaveAs writes (the portfolio's office_output_methods.md section 6.1).
// Compressed twice (does a second export differ?), Text and Pdf, then, after edits, Compressed and Text again: an unsaved
// document's state, and a file-level reading of what check 77 could only read through getOoxml. Compared afterwards with a
// `save` job on the same input, allowing for the pane's parts and for DocSecurity. #out is set once, last.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => {
  out.value = JSON.stringify({ check: 78, failed: String((e && e.stack) || e) });
}));

const SLICE = 4194304;
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const AGENT = 'Agent Seven';

const packageOf = (blocks) => `<pkg:package xmlns:pkg="${PKG}">`
  + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
  + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  + '</pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + `<pkg:xmlData><w:document xmlns:w="${W}" xmlns:w16du="${W16DU}" xmlns:mc="${MC}" mc:Ignorable="w16du"><w:body>${blocks}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;

let nextId = 9000;
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
/** A redline by the agent, its date as the package gives it: `when` is the attributes after the author. */
const redline = (label, when) => `<w:p>${run(`${label} before `)}<w:del w:id="${nextId++}" w:author="${AGENT}" ${when}><w:r><w:delText>TARGET</w:delText></w:r></w:del>`
  + `<w:ins w:id="${nextId++}" w:author="${AGENT}" ${when}>${run('REPLACED')}</w:ins>${run(' after')}</w:p>`;

const errorOf = (e) => ({ message: String((e && e.message) || e), code: e && e.code });
const asyncError = (r) => ({ message: r.error && r.error.message, code: r.error && r.error.code, name: r.error && r.error.name });

function base64(parts) {
  let s = '';
  for (const bytes of parts) for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, Array.prototype.slice.call(bytes, i, i + 0x8000));
  return btoa(s);
}

/** One export, every slice read and the file closed. Text gives strings, the others bytes. */
function getFile(type) {
  return new Promise((resolve) => {
    const t0 = performance.now();
    Office.context.document.getFileAsync(type, { sliceSize: SLICE }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) { resolve({ type, error: asyncError(res) }); return; }
      const file = res.value;
      const parts = [];
      const done = (extra) => file.closeAsync(() => {
        const r = { type, size: file.size, sliceCount: file.sliceCount, ms: Math.round(performance.now() - t0), ...extra };
        if (!r.error) {
          if (parts.every((p) => typeof p === 'string')) r.text = parts.join('');
          else r.base64 = base64(parts);
        }
        resolve(r);
      });
      const next = (i) => {
        if (i === file.sliceCount) { done({}); return; }
        file.getSliceAsync(i, (s) => {
          if (s.status !== Office.AsyncResultStatus.Succeeded) { done({ error: asyncError(s), failedSlice: i }); return; }
          parts.push(s.value.data);
          next(i + 1);
        });
      };
      next(0);
    });
  });
}

/** Three new paragraphs at the end, tracking off: plain, a redline dated as check 77's, and one dated as Word's own marks are. */
async function edit(report) {
  await Word.run(async (context) => {
    context.document.changeTrackingMode = 'Off';
    await context.sync();
    const body = context.document.body;
    body.insertParagraph('EDITED PLAIN', 'End');
    const a = body.insertParagraph('EDITED A', 'End');
    const b = body.insertParagraph('EDITED B', 'End');
    await context.sync();
    a.getRange('Content').insertOoxml(packageOf(redline('NODATEUTC', 'w:date="2026-01-02T03:04:00Z"')), 'Replace');
    b.getRange('Content').insertOoxml(packageOf(redline('DATEUTC', 'w:date="2026-01-02T14:04:00Z" w16du:dateUtc="2026-01-02T03:04:00Z"')), 'Replace');
    await context.sync();
    const flat = body.getOoxml();
    const changes = body.getTrackedChanges();
    changes.load('items/type,items/author,items/date,items/text');
    await context.sync();
    const at = flat.value.indexOf('EDITED PLAIN');
    report.editedXml = at < 0 ? null : flat.value.slice(at - 200, flat.value.indexOf('</w:body>', at));
    report.listed = changes.items.map((c) => ({ type: c.type, author: c.author, date: c.date instanceof Date ? c.date.toISOString() : String(c.date), text: c.text }));
  });
}

async function check() {
  const report = { check: 78, ranAt: new Date().toISOString(), host: Office.context.diagnostics, mode: Office.context.document.mode, files: {} };
  report.files.compressed1 = await getFile(Office.FileType.Compressed);
  report.files.compressed2 = await getFile(Office.FileType.Compressed);
  report.files.text = await getFile(Office.FileType.Text);
  report.files.pdf = await getFile(Office.FileType.Pdf);
  try { await edit(report); } catch (e) { report.editError = errorOf(e); }
  report.files.compressedEdited = await getFile(Office.FileType.Compressed);
  report.files.textEdited = await getFile(Office.FileType.Text);
  out.value = JSON.stringify(report);
}
