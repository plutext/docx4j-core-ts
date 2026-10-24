// Check 61 (CR-015 section 12; test/README.md check 61), on 61-highlights.docx (test/fixtures/check61/build.mjs). WordApi
// 1.3. What Office JS's font.highlightColor says for each of w:highlight's sixteen names, and what Word stores for each
// value written to it: the objects package's highlight table is docx4j's, and check 60 found Word's green elsewhere.
// - Each `READ <name>: R_<name> (end)` paragraph's marker run carries w:highlight w:val="<name>" (R_SHD a character
//   shading instead); the snippet reads its font.highlightColor.
// - Each `WRITE <n> <value>: W<nn> (end)` paragraph's marker is plain (the one for a null write starts highlighted yellow);
//   the snippet sets font.highlightColor to WRITES[n] there, then reads it back, and reads the run's w:highlight and
//   w:shd from the range's getOoxml().
// The writes change the document: close it without saving. Nothing here depends on Script Lab but the last few lines,
// so a test can load CHECK61 with Word bound to another implementation.
const CHECK61 = (() => {
  const failed = (e) => 'error: ' + ((e && (e.code || e.message)) || String(e));

  /** w:highlight's values (ST_HighlightColor), as the schema spells them. */
  const NAMES = ['black', 'blue', 'cyan', 'green', 'magenta', 'red', 'yellow', 'white', 'darkBlue', 'darkCyan', 'darkGreen',
    'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray'];

  /**
   * What is written, in order: each name; Word's own palette as it is usually given (WdColorIndex); the objects
   * package's (docx4j's) values that differ from it; lower case; two colours in neither; and null on a highlighted run.
   */
  const WRITES = [
    ...NAMES,
    '#000000', '#0000FF', '#00FFFF', '#00FF00', '#FF00FF', '#FF0000', '#FFFF00', '#FFFFFF',
    '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#C0C0C0',
    '#00008B', '#008B8B', '#006400', '#8B008B', '#8B0000', '#FFD700', '#A9A9A9', '#D3D3D3',
    '#00ff00', 'Yellow', '#123456', '#FF8000', null,
  ];

  /** A paragraph's text as a read case, `{ read, marker }`, or a write case, `{ write, marker }`; else null. */
  function parse(text) {
    const r = /^READ (\S+): (\S+) \(end\)$/.exec(text);
    if (r) return { read: r[1], marker: r[2] };
    const w = /^WRITE (\d+) [\s\S]*: (\S+) \(end\)$/.exec(text);
    return w ? { write: Number(w[1]), marker: w[2] } : null;
  }

  /** The marker's range: the last case-sensitive hit in its paragraph. */
  async function markerRange(context, index, marker) {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();
    const hits = paragraphs.items[index].search(marker, { matchCase: true });
    hits.load('items');
    await context.sync();
    if (hits.items.length === 0) throw new Error('not found: ' + marker);
    return hits.items[hits.items.length - 1];
  }

  /** What getOoxml() says the run stores: its w:highlight's value and its w:shd, as written. */
  function stored(ooxml, marker) {
    const at = ooxml.indexOf('>' + marker + '<');
    const run = at < 0 ? ooxml : ooxml.substring(Math.max(ooxml.lastIndexOf('<w:r>', at), ooxml.lastIndexOf('<w:r ', at)), at);
    const highlight = /<w:highlight\b[^>]*?w:val="([^"]*)"/.exec(run);
    const shd = /<w:shd\b[^>]*\/>/.exec(run);
    return { highlight: highlight ? highlight[1] : null, shd: shd ? shd[0] : null };
  }

  async function read(context, index, marker) {
    const range = await markerRange(context, index, marker);
    range.font.load('highlightColor');
    await context.sync();
    return { highlightColor: range.font.highlightColor };
  }

  async function write(context, index, marker, value) {
    const range = await markerRange(context, index, marker);
    const out = {};
    try {
      range.font.highlightColor = value;
      await context.sync();
    } catch (e) {
      out.set = failed(e);
    }
    try {
      range.font.load('highlightColor');
      await context.sync();
      out.highlightColor = range.font.highlightColor;
    } catch (e) {
      out.highlightColor = failed(e);
    }
    try {
      const ooxml = range.getOoxml();
      await context.sync();
      out.stored = stored(ooxml.value, marker);
    } catch (e) {
      out.stored = failed(e);
    }
    return out;
  }

  /** The check: `run(fn)` is Word.run; `label` names the run; `progress(text)` is told each case. */
  async function runCheck(run, label, progress = () => {}, host = null) {
    const result = { check: 61, label, host, reads: [], writes: [] };
    const texts = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => p.text);
    });
    for (let i = 0; i < texts.length; i++) {
      const c = parse(texts[i]);
      if (c === null) continue;
      if (c.read !== undefined) {
        progress('read ' + c.read);
        result.reads.push({ name: c.read, marker: c.marker, ...(await run((context) => read(context, i, c.marker)).catch((e) => ({ error: failed(e) }))) });
      } else {
        const value = WRITES[c.write];
        progress('write ' + c.write + ' of ' + WRITES.length + ': ' + value);
        result.writes.push({ n: c.write, value, marker: c.marker, ...(await run((context) => write(context, i, c.marker, value)).catch((e) => ({ error: failed(e) }))) });
      }
    }
    return result;
  }

  return { NAMES, WRITES, parse, stored, runCheck };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const button = document.getElementById('run');
  button.addEventListener('click', () => {
    button.disabled = true;
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK61.runCheck((f) => Word.run(f), document.getElementById('label').value, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.reads.length + ' reads, ' + result.writes.length + ' writes. Copy the box below back, then close the document without saving.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => { button.disabled = false; });
  });
}
