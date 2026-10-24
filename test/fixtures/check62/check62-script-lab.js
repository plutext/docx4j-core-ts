// Check 62 (CR-015 section 12.2; test/README.md check 62), on 62-highlight-snap.docx (test/fixtures/check62/build.mjs).
// WordApi 1.3. Word's rule for the highlight a written colour becomes: check 61 found Office JS storing any #RRGGBB as one
// of w:highlight's sixteen colours, not always the nearest in RGB. Every paragraph after the title holds one colour as
// six hex digits; the snippet sets that paragraph's font.highlightColor to it, in batches, then reads each back, and for
// every SPOT-th paragraph also takes the w:highlight Word stored from its getOoxml(). The writes change the document:
// close it without saving. Nothing here depends on Script Lab but the last few lines, so a test can load CHECK62 with
// Word bound to another implementation.
const CHECK62 = (() => {
  const failed = (e) => 'error: ' + ((e && (e.code || e.message)) || String(e));
  const BATCH = 400;
  const SPOT = 97;

  /** A paragraph's text as the colour it holds, `#RRGGBB`, or null. */
  function parse(text) {
    return /^[0-9A-F]{6}$/.test(text) ? '#' + text : null;
  }

  /** The w:highlight value a paragraph's getOoxml() holds, or null. */
  function stored(ooxml) {
    const m = /<w:highlight\b[^>]*?w:val="([^"]*)"/.exec(ooxml);
    return m ? m[1] : null;
  }

  /** One batch: set every colour, sync once; read every one back, sync once. A batch that fails is done one by one. */
  async function batch(context, items, cases, out) {
    try {
      for (const c of cases) items[c.index].font.highlightColor = c.value;
      await context.sync();
    } catch (e) {
      for (const c of cases) {
        try {
          items[c.index].font.highlightColor = c.value;
          await context.sync();
        } catch (e2) {
          out[c.n].set = failed(e2);
        }
      }
    }
    for (const c of cases) items[c.index].font.load('highlightColor');
    await context.sync();
    for (const c of cases) out[c.n].read = items[c.index].font.highlightColor;
    const spots = cases.filter((c) => c.n % SPOT === 0);
    const ooxml = spots.map((c) => items[c.index].getOoxml());
    await context.sync();
    spots.forEach((c, i) => { out[c.n].stored = stored(ooxml[i].value); });
  }

  /** The check: `run(fn)` is Word.run; `label` names the run; `progress(text)` is told each batch. */
  async function runCheck(run, label, progress = () => {}, host = null) {
    const result = { check: 62, label, host, points: [] };
    const texts = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => p.text);
    });
    const cases = [];
    texts.forEach((text, index) => {
      const value = parse(text);
      if (value !== null) cases.push({ n: cases.length, index, value });
    });
    const out = cases.map((c) => ({ value: c.value }));
    for (let start = 0; start < cases.length; start += BATCH) {
      progress('colours ' + (start + 1) + ' to ' + Math.min(start + BATCH, cases.length) + ' of ' + cases.length);
      const slice = cases.slice(start, start + BATCH);
      await run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        await batch(context, paragraphs.items, slice, out);
      }).catch((e) => { for (const c of slice) if (out[c.n].read === undefined) out[c.n].error = failed(e); });
    }
    // compact: [written, read back] per point, and the set error or stored name where there is one
    result.points = out.map((p) => {
      const row = [p.value.substring(1), typeof p.read === 'string' ? p.read.replace(/^#/, '') : p.read === undefined ? null : p.read];
      if (p.set !== undefined || p.error !== undefined) row.push(p.set ?? p.error);
      else if (p.stored !== undefined) row.push({ stored: p.stored });
      return row;
    });
    return result;
  }

  return { parse, stored, runCheck, BATCH, SPOT };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const button = document.getElementById('run');
  button.addEventListener('click', () => {
    button.disabled = true;
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK62.runCheck((f) => Word.run(f), document.getElementById('label').value, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result); status.textContent = 'Done: ' + result.points.length + ' colours. Copy the box below back, then close the document without saving.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => { button.disabled = false; });
  });
}
