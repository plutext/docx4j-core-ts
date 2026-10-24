// Check 60 (CR-015 D4; test/README.md check 60), on 60-runprops.docx (test/fixtures/check60/build.mjs). WordApi
// 1.3. Read only: every paragraph after a `PROBE <name>` one is `label: MARKER (end)`, and for each the snippet finds
// the marker (a case-sensitive search in its paragraph, the last hit) and reads its font, each property synced alone,
// so one Office JS lacks reads as an error and the rest still come back. Nothing here depends on Script Lab but the
// last few lines, so a test can load CHECK60 with Word bound to another implementation.
const CHECK60 = (() => {
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && (e.code || e.message)) || String(e));

  /** The font properties read: WordApi 1.1's, then WordApiDesktop's on/off ones, which a host may lack. */
  const FONT = ['bold', 'italic', 'underline', 'strikeThrough', 'doubleStrikeThrough', 'superscript', 'subscript', 'color', 'highlightColor',
    'allCaps', 'smallCaps', 'hidden', 'emboss', 'engrave', 'outline', 'shadow', 'underlineColor'];

  /** A paragraph's text as a probe heading, `{ probe }`, or a marker paragraph, `{ label, marker }`; else null. */
  function parse(text) {
    const probe = /^PROBE (\S+)$/.exec(text);
    if (probe) return { probe: probe[1] };
    const m = /^([\s\S]*): ([\s\S]*) \(end\)$/.exec(text);
    return m ? { label: m[1], marker: m[2] } : null;
  }

  async function attempt(f) {
    try {
      return await f();
    } catch (e) {
      return failed(e);
    }
  }

  /** The marker's font, property by property; `hits` how many times the search found it in its paragraph. */
  async function markerReads(context, index, marker) {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();
    const hits = paragraphs.items[index].search(marker, { matchCase: true });
    hits.load('items/text');
    await context.sync();
    const out = { hits: hits.items.length };
    if (hits.items.length === 0) return out;
    const font = hits.items[hits.items.length - 1].font;
    out.font = {};
    for (const p of FONT) {
      out.font[p] = await attempt(async () => {
        font.load(p);
        await context.sync();
        return font[p];
      });
    }
    return out;
  }

  /** The check: `run(fn)` is Word.run; `label` names the run; `progress(text)` is told each probe. */
  async function runCheck(run, label, progress = () => {}, host = null) {
    const result = { check: 60, label, host, markers: [] };
    const texts = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => p.text);
    });
    let probe = null;
    for (let i = 0; i < texts.length; i++) {
      const parsed = parse(texts[i]);
      if (parsed !== null && parsed.probe) {
        probe = parsed.probe;
        progress('probe ' + probe);
        continue;
      }
      const entry = { probe, paragraph: i, text: vis(texts[i]) };
      if (parsed !== null && parsed.marker !== '') {
        entry.marker = vis(parsed.marker);
        entry.reads = await run((context) => markerReads(context, i, parsed.marker)).catch(failed);
      }
      result.markers.push(entry);
    }
    return result;
  }

  return { parse, runCheck, vis, FONT };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const button = document.getElementById('run');
  button.addEventListener('click', () => {
    button.disabled = true;
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK60.runCheck((f) => Word.run(f), document.getElementById('label').value, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.markers.length + ' paragraphs. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => { button.disabled = false; });
  });
}
