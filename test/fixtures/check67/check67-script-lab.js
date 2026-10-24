// Check 67 (CR-002 section 63.13; test/README.md check 67): why checks 58 and 63 disagree on the bare form's look, on
// 67a-plain.docx, 67b-custom-xml.docx and 67c-revisions.docx (test/fixtures/check67/build.mjs), one file at a time.
// WordApi 1.5 for getContentControls, 1.4 for changeTrackingMode, 1.3 for the rest. Every case is a region,
// BEGIN<n> <op> <shape> <phase> <lookup> ... END<n>, around one control tagged c67-<n> in the bare form. "Report only"
// reads every region and changes nothing. Each phase button gives its phase's cases check 58's write and reads each region
// after it: in an `-untouched` phase changeTrackingMode is never touched, as check 58's fills left it; in an `-off` phase it
// is set 'Off' before each write and after it, as check 63's run set it. The control is found as the case's lookup says:
// `tag` as check 58 found it, `list` as check 63 did. The page keeps every result in one box. Nothing here depends on
// Script Lab but the last lines, so a test can load CHECK67 with Word bound to another implementation.
const CHECK67 = (() => {
  // every character outside printable ASCII as \uXXXX, so a control character shows
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && e.code) || 'no code') + (e && e.message ? ': ' + e.message : '');
  const TOKEN = { matchCase: true, matchWholeWord: true };
  const tagOf = (n) => 'c67-' + n;
  /** The phases in the order they run, and the view each wants. */
  const PHASES = [['simple-untouched', 'Simple Markup'], ['all-untouched', 'All Markup'], ['all-off', 'All Markup'], ['simple-off', 'Simple Markup']];

  /** A BEGIN paragraph's text as a case, or null. */
  function parseRegion(text) {
    const m = /^BEGIN(\d+) ([\w-]+) ([\w-]+) ([\w-]+) (tag|list)$/.exec(text.trim());
    return m ? { n: Number(m[1]), op: m[2], shape: m[3], phase: m[4], lookup: m[5] } : null;
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

  /**
   * The case's control, as its lookup finds it: `tag`, check 58's body.contentControls.getByTag(tag).getFirst(), unloaded;
   * `list`, check 63's body.getContentControls() loaded and filtered by tag.
   */
  async function controlOf(context, region) {
    if (region.lookup === 'tag') return context.document.body.contentControls.getByTag(tagOf(region.n)).getFirst();
    const controls = context.document.body.getContentControls();
    controls.load('items/tag');
    await context.sync();
    const mine = controls.items.filter((c) => c.tag === tagOf(region.n));
    if (mine.length === 0) throw new Error('not found: ' + tagOf(region.n));
    return mine[0];
  }

  /** The case's write, check 58's (FILL<n>, where check 63 wrote NEW<n>); what its result reads. */
  async function apply(context, region) {
    const c = await controlOf(context, region);
    const text = 'FILL' + region.n;
    const r = region.op === 'control-text-replace' ? c.insertText(text, 'Replace') : c.insertText(text, 'End');
    r.load('text');
    await context.sync();
    return { returned: vis(r.text) };
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

  /** The region as it stands: its paragraphs' text, its OOXML, and the control's reads, found by the case's lookup. */
  async function report(context, region) {
    const out = {};
    out.paragraphs = await attempt(async () => (await contentParagraphs(context, region.n)).map((p) => vis(p.text)));
    out.ooxml = await attempt(async () => { const x = (await regionOf(context, region.n)).getOoxml(); await context.sync(); return bodyOf(x.value); });
    out.control = await attempt(async () => {
      const c = await controlOf(context, region);
      c.load('text,type,placeholderText');
      await context.sync();
      return { type: c.type, text: vis(c.text), placeholderText: vis(c.placeholderText) };
    });
    return out;
  }

  /** changeTrackingMode set 'Off', as check 63 set it around each write. */
  async function trackingOff(context) {
    context.document.changeTrackingMode = 'Off';
    await context.sync();
  }

  /**
   * One step of the check: `step` 'report', or a phase's name. `run(fn)` is Word.run; `progress(text)` is told each case.
   * A report reads every region; a phase writes and reads its own cases only. Each write is a Word.run of its own, so
   * that a case Office JS refuses fails alone. An `-untouched` phase never reads or sets changeTrackingMode.
   */
  async function runCheck(run, step, progress = () => {}, host = null) {
    const result = { check: 67, step, host, title: null, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      result.title = paragraphs.items.length > 0 ? paragraphs.items[0].text : null;
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    result.caseCount = regions.length;
    const mine = step === 'report' ? regions : regions.filter((r) => r.phase === step);
    const off = step.endsWith('-off');
    for (const region of mine) {
      progress(step + ': case ' + region.n + ' (' + region.op + ' ' + region.shape + ', found by ' + region.lookup + ')');
      const entry = { ...region };
      if (step !== 'report') {
        try {
          entry.result = await run(async (context) => {
            if (off) await trackingOff(context);
            return apply(context, region);
          });
        } catch (e) {
          entry.error = failed(e);
        }
        if (off) await run((context) => trackingOff(context)).catch(() => {});
      }
      entry.after = await run((context) => report(context, region)).catch(failed);
      result.cases.push(entry);
    }
    return result;
  }

  return { parseRegion, runCheck, vis, bodyOf, PHASES };
})();

if (typeof document !== 'undefined' && document.getElementById('report')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const steps = ['report', ...CHECK67.PHASES.map(([name]) => name)];
  const buttons = steps.map((id) => document.getElementById(id));
  const all = { check: 67, host: null, title: null, steps: {} };
  let done = 0;   // the steps run so far, in order: each button is enabled once the one before it has run
  const enable = () => buttons.forEach((b, i) => { b.disabled = i !== done; });
  const next = () => (done < steps.length
    ? 'Next: ' + (done === 0 ? 'Report only, in Simple Markup' : 'set Review > Display for Review to ' + CHECK67.PHASES[done - 1][1] + ', then click ' + buttons[done].textContent)
    : 'Done: every step has run. Copy the box below back.');
  enable();
  status.textContent = next();
  buttons.forEach((button, i) => button.addEventListener('click', () => {
    buttons.forEach((b) => { b.disabled = true; });
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK67.runCheck((f) => Word.run(f), steps[i], (text) => { status.textContent = text; }, host)
      .then((result) => {
        all.host = result.host;
        all.title = result.title;
        all.steps[steps[i]] = result;
        out.value = JSON.stringify(all, null, 1);
        done = i + 1;
        status.textContent = next();
      })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(enable);
  }));
}
