// Check 57 (CR-002 section 62; test/README.md check 57): a paragraph break inside a field's result, and a tracked
// delete of a paragraph holding an equation, on 57-fields-equations.docx (test/fixtures/check57/build.mjs). WordApi 1.6
// for getTrackedChanges, 1.4 for fields and changeTrackingMode, 1.3 for the rest. Every case is a region,
// BEGIN<n> <op> <shape> ... END<n>, around one paragraph. "Run the check" gives each case its write: in a field's result,
// insertText('x\ny', 'After') or insertParagraph('NEWPARA<n>', 'After') on the word FA<n>, untracked; on an equation's
// paragraph, delete() of the paragraph or of a range, under TrackAll. An `open` case is only reported. Each region is
// then reported: its text, paragraphs, OOXML, tracked changes and fields. "Report only" reports and changes nothing.
// Nothing here depends on Script Lab but the last few lines, so a test can load CHECK57 with Word bound to another
// implementation.
const CHECK57 = (() => {
  // every character outside printable ASCII as \uXXXX (a surrogate pair as two), so a control character shows
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && e.code) || 'no code') + (e && e.message ? ': ' + e.message : '');
  const TOKEN = { matchCase: true, matchWholeWord: true };

  /** A BEGIN paragraph's text as a case, or null. */
  function parseRegion(text) {
    const m = /^BEGIN(\d+) ([\w-]+) ([\w-]+)$/.exec(text.trim());
    return m ? { n: Number(m[1]), op: m[2], shape: m[3] } : null;
  }

  async function find(context, token) {
    const hit = context.document.body.search(token, TOKEN).getFirstOrNullObject();
    hit.load('text');
    await context.sync();
    if (hit.isNullObject) throw new Error('not found: ' + token);
    return hit;
  }
  /** The paragraph holding a token. */
  const paragraphOf = async (context, token) => (await find(context, token)).paragraphs.getFirst();
  /** The region from its BEGIN paragraph to its END paragraph, both whole. */
  const regionOf = async (context, n) => (await paragraphOf(context, 'BEGIN' + n)).getRange('Whole')
    .expandTo((await paragraphOf(context, 'END' + n)).getRange('Whole'));
  /** The case's paragraph: the region's second. */
  async function targetOf(context, n) {
    const paragraphs = (await regionOf(context, n)).paragraphs;
    paragraphs.load('items');
    await context.sync();
    return paragraphs.items[1];
  }

  /** One read, synced alone, so that what Office JS refuses fails alone. */
  async function attempt(f) {
    try {
      return await f();
    } catch (e) {
      return failed(e);
    }
  }

  /** The case's write, the equations' under TrackAll; what it returns. */
  async function apply(context, region) {
    const tracked = region.shape.endsWith('-equation');
    context.document.changeTrackingMode = tracked ? 'TrackAll' : 'Off';
    context.document.load('changeTrackingMode');
    await context.sync();
    const out = { changeTrackingMode: context.document.changeTrackingMode };
    const returned = async (r) => { r.load('text'); await context.sync(); out.returned = vis(r.text); return out; };
    try {
      switch (region.op) {
        case 'insertText-after': return await returned((await find(context, 'FA' + region.n)).insertText('x\ny', 'After'));
        case 'insertParagraph-after': return await returned((await find(context, 'FA' + region.n)).insertParagraph('NEWPARA' + region.n, 'After'));
        case 'paragraph-delete': (await targetOf(context, region.n)).delete(); break;
        case 'range-delete':
          if (region.shape === 'inline-equation') (await find(context, 'L' + region.n)).expandTo(await find(context, 'R' + region.n)).delete();
          else (await targetOf(context, region.n)).getRange('Content').delete();
          break;
        default: throw new Error('no operation ' + region.op);
      }
      await context.sync();
      return out;
    } finally {
      context.document.changeTrackingMode = 'Off';
      await context.sync();
    }
  }

  const tidy = (xml) => xml.replace(/ w:rsid\w*="[^"]*"| w14:(?:paraId|textId)="[^"]*"| xmlns:\w+="[^"]*"/g, '');
  /** A range's OOXML: the body's content without its section properties (or the range's own XML, as given). */
  function bodyOf(xml) {
    if (xml.indexOf('pkg:package') < 0) return vis(tidy(xml));
    const at = xml.indexOf('pkg:name="/word/document.xml"');
    const start = xml.indexOf('<pkg:xmlData>', at);
    const end = xml.indexOf('</pkg:xmlData>', start);
    const body = /<w:body>([\s\S]*?)<\/w:body>/.exec(at < 0 || start < 0 || end < 0 ? '' : xml.slice(start, end));
    return body ? vis(tidy(body[1].replace(/<w:sectPr\b[\s\S]*<\/w:sectPr>/, ''))) : null;
  }

  /** The region as it stands: its text, paragraphs, OOXML, tracked changes and fields. */
  async function report(context, region) {
    const out = {};
    const whole = await regionOf(context, region.n);
    out.text = await attempt(async () => { whole.load('text'); await context.sync(); return vis(whole.text); });
    out.paragraphs = await attempt(async () => {
      const paragraphs = whole.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => vis(p.text));
    });
    out.ooxml = await attempt(async () => { const x = whole.getOoxml(); await context.sync(); return bodyOf(x.value); });
    out.trackedChanges = await attempt(async () => {
      const changes = whole.getTrackedChanges();
      changes.load('items/type,items/text,items/author');
      await context.sync();
      return changes.items.map((c) => ({ type: c.type, text: vis(c.text), author: c.author }));
    });
    out.fields = await attempt(async () => {
      const fields = whole.fields;
      fields.load('items/code');
      await context.sync();
      const results = fields.items.map((f) => { const r = f.result; r.load('text'); return r; });
      await context.sync();
      return fields.items.map((f, i) => ({ code: vis(f.code), result: vis(results[i].text) }));
    });
    return out;
  }

  /**
   * The check: `mode` 'run' or 'report'. `run(fn)` is Word.run; `progress(text)` is told each case. Each step is a
   * Word.run of its own, so that a case Office JS refuses fails alone.
   */
  async function runCheck(run, mode, progress = () => {}, host = null) {
    const result = { check: 57, mode, host, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    result.caseCount = regions.length;
    for (const region of regions) {
      progress('case ' + region.n + ' of ' + regions.length + ': ' + region.op + ' ' + region.shape);
      const entry = { n: region.n, op: region.op, shape: region.shape };
      try {
        if (mode === 'run' && region.op !== 'open') entry.result = await run((context) => apply(context, region));
      } catch (e) {
        entry.error = failed(e);
      }
      entry.region = await run((context) => report(context, region)).catch(failed);
      result.cases.push(entry);
    }
    return result;
  }

  return { parseRegion, runCheck, vis };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const buttons = ['run', 'report'].map((id) => document.getElementById(id));
  const start = (mode) => {
    buttons.forEach((b) => { b.disabled = true; });
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK57.runCheck((f) => Word.run(f), mode, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.caseCount + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => buttons.forEach((b) => { b.disabled = false; }));
  };
  buttons[0].addEventListener('click', () => start('run'));
  buttons[1].addEventListener('click', () => start('report'));
}
