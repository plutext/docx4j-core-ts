// Check 68 (CR-002 section 63.19; test/README.md check 68): the run properties a write into a placeholder takes, and how
// Word combines a glossary entry with a control's own run properties, on 68-placeholder-run-properties.docx
// (test/fixtures/check68/build.mjs). Check 66's snippet with check 68's names and ops: WordApi 1.5 for getContentControls,
// 1.4 for changeTrackingMode, 1.3 for the rest. Every case is a region, BEGIN<n> <op> <shape> <form> run-<r> control-<c> ...
// END<n>, around one control tagged c68-<n> (build.mjs lists them). "Report only" reports every region and changes
// nothing; "Run the check" gives each case its write and reports the region after it. Developer > Design Mode stays off.
// Nothing here depends on Script Lab but the last few lines, so a test can load CHECK68 with Word bound to another
// implementation.
const CHECK68 = (() => {
  // every character outside printable ASCII as \uXXXX, so a control character shows
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && e.code) || 'no code') + (e && e.message ? ': ' + e.message : '');
  const TOKEN = { matchCase: true, matchWholeWord: true };
  const tagOf = (n) => 'c68-' + n;

  /** A BEGIN paragraph's text as a case, or null. */
  function parseRegion(text) {
    const m = /^BEGIN(\d+) ([\w-]+) (block|run) (bare|word|value) run-(\w+) control-(\w+)$/.exec(text.trim());
    return m ? { n: Number(m[1]), op: m[2], shape: m[3], form: m[4], run: m[5], control: m[6] } : null;
  }

  async function attempt(f) {
    try {
      const v = await f();
      return v === undefined ? 'undefined' : v;
    } catch (e) {
      return failed(e);
    }
  }
  async function find(context, token) {
    const hit = context.document.body.search(token, TOKEN).getFirstOrNullObject();
    hit.load('text');
    await context.sync();
    if (hit.isNullObject) throw new Error('not found: ' + token);
    return hit;
  }
  const paragraphOf = async (context, token) => (await find(context, token)).paragraphs.getFirst();
  /** The region from its BEGIN paragraph to its END paragraph, both whole. */
  const regionOf = async (context, n) => (await paragraphOf(context, 'BEGIN' + n)).getRange('Whole')
    .expandTo((await paragraphOf(context, 'END' + n)).getRange('Whole'));
  /** The region's paragraphs between its BEGIN and END paragraphs. */
  async function contentParagraphs(context, n) {
    const ps = (await regionOf(context, n)).paragraphs;
    ps.load('items/text');
    await context.sync();
    return ps.items.slice(1, -1);
  }

  /** The ways the case's control is looked for, in order (check 56: Word's contentControls lists no bound control). */
  const LOOKUPS = [
    ['body.getContentControls()', (context) => context.document.body.getContentControls()],
    ['body.contentControls', (context) => context.document.body.contentControls],
  ];
  /** The case's control through the first way that finds it: { control, via }, or null. */
  async function controlOf(context, n) {
    for (const [via, way] of LOOKUPS) {
      try {
        const controls = way(context);
        controls.load('items/tag');
        await context.sync();
        const mine = controls.items.filter((c) => c.tag === tagOf(n));
        if (mine.length > 0) return { control: mine[0], via };
      } catch (e) {
        // that way isn't there, or refuses: the next
      }
    }
    return null;
  }

  /** The case's write; what its result reads. */
  async function apply(context, region) {
    const found = await controlOf(context, region.n);
    if (!found) throw new Error('not found: ' + tagOf(region.n));
    const c = found.control;
    const text = 'NEW' + region.n;
    const read = async (r) => { r.load('text'); await context.sync(); return { via: found.via, returned: vis(r.text) }; };
    switch (region.op) {
      case 'control-text-replace': return read(c.insertText(text, 'Replace'));
      case 'control-text-end': return read(c.insertText(text, 'End'));
      case 'content-delete': c.getRange('Content').delete(); await context.sync(); return { via: found.via };
      default: throw new Error('no operation ' + region.op);
    }
  }

  const tidy = (xml) => xml.replace(/ w:rsid\w*="[^"]*"| w14:(?:paraId|textId)="[^"]*"| xmlns:\w+="[^"]*"/g, '');
  /** A range's OOXML: the body's content without its section properties. */
  function bodyOf(xml) {
    const at = xml.indexOf('pkg:name="/word/document.xml"');
    const start = xml.indexOf('<pkg:xmlData>', at);
    const end = xml.indexOf('</pkg:xmlData>', start);
    const body = /<w:body>([\s\S]*?)<\/w:body>/.exec(at < 0 || start < 0 || end < 0 ? xml : xml.slice(start, end));
    return body ? vis(tidy(body[1].replace(/<w:sectPr\b[\s\S]*<\/w:sectPr>/, ''))) : vis(tidy(xml));
  }
  /** The region as it stands: its paragraphs' text, its OOXML, and the control's reads. */
  async function report(context, region) {
    const out = {};
    out.paragraphs = await attempt(async () => (await contentParagraphs(context, region.n)).map((p) => vis(p.text)));
    out.ooxml = await attempt(async () => { const x = (await regionOf(context, region.n)).getOoxml(); await context.sync(); return bodyOf(x.value); });
    const found = await controlOf(context, region.n).catch(() => null);
    out.control = found === null ? null : await attempt(async () => {
      const c = found.control;
      c.load('text,type,placeholderText');
      const ps = c.paragraphs;
      ps.load('items/text');
      await context.sync();
      return { via: found.via, type: c.type, text: vis(c.text), placeholderText: vis(c.placeholderText), paragraphs: ps.items.map((p) => vis(p.text)) };
    });
    if (found !== null) {
      out.showingPlaceholderText = await attempt(async () => { found.control.load('showingPlaceholderText'); await context.sync(); return found.control.showingPlaceholderText; });
    }
    return out;
  }

  /** Change tracking on or off, as a case's mode says. */
  async function track(context, on) {
    context.document.changeTrackingMode = on ? 'TrackAll' : 'Off';
    await context.sync();
  }

  /**
   * The check: `mode` 'run' or 'report'. `run(fn)` is Word.run; `progress(text)` is told each case. Each step is a
   * Word.run of its own, so that a case Office JS refuses fails alone. Change tracking is off at the end.
   */
  async function runCheck(run, mode, progress = () => {}, host = null) {
    const result = { check: 68, mode, host, title: null, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      result.title = paragraphs.items.length > 0 ? paragraphs.items[0].text : null;
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    result.caseCount = regions.length;
    try {
      for (const region of regions) {
        progress('case ' + region.n + ' of ' + regions.length + ': ' + region.op + ' ' + region.shape + ' ' + region.form);
        const entry = { ...region };
        if (mode === 'run') {
          try {
            entry.result = await run(async (context) => {
              await track(context, false);
              return apply(context, region);
            });
          } catch (e) {
            entry.error = failed(e);
          }
          await run((context) => track(context, false)).catch(() => {});
        }
        entry.after = await run((context) => report(context, region)).catch(failed);
        result.cases.push(entry);
      }
    } finally {
      await run((context) => track(context, false)).catch(() => {});
    }
    return result;
  }

  return { parseRegion, runCheck, vis, bodyOf };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const buttons = ['report', 'run'].map((id) => document.getElementById(id));
  const start = (mode) => {
    buttons.forEach((b) => { b.disabled = true; });
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK68.runCheck((f) => Word.run(f), mode, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.caseCount + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => buttons.forEach((b) => { b.disabled = false; }));
  };
  buttons[0].addEventListener('click', () => start('report'));
  buttons[1].addEventListener('click', () => start('run'));
}
