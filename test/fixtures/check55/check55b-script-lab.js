// Check 55b (CR-002 section 60.5; test/README.md check 55b): the bound plain-text controls check 55 couldn't find, on
// 55b-bound.docx (test/fixtures/check55/build.mjs). WordApi 1.5 for getContentControls, 1.4 for the rest. As check
// 55's: every case is a region, BEGIN<n> <op> <tracking> <characters> ... END<n>. Each case's control (tagged c55b-<n>)
// is looked for five ways, each recorded; "Run the check" writes 'x', the case's characters and 'y' into it with
// insertText(..., 'Replace') through the first way that finds it, and reports the region, the control and its custom
// XML node; the bound-read case writes nothing. "Report only" reports and changes nothing. Nothing here depends on
// Script Lab but the last few lines, so a test can load CHECK55B with Word bound to another implementation.
const CHECK55B = (() => {
  // every character outside printable ASCII as \uXXXX (a surrogate pair as two), so a control character shows
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && e.code) || 'no code') + (e && e.message ? ': ' + e.message : '');
  const TOKEN = { matchCase: true, matchWholeWord: true };
  const DATA_NS = 'urn:check55b';
  const tagOf = (n) => 'c55b-' + n;

  /** A BEGIN paragraph's text as a case, or null: its characters, and the text the case writes. */
  function parseRegion(text) {
    const m = /^BEGIN(\d+) ([\w-]+) (untracked|tracked) (U\+[0-9A-F]{4}(?:,U\+[0-9A-F]{4})*)$/.exec(text.trim());
    if (!m) return null;
    const chars = m[4].split(',').map((u) => parseInt(u.slice(2), 16));
    return { n: Number(m[1]), op: m[2], tracked: m[3] === 'tracked', spec: m[4], chars, text: 'x' + String.fromCodePoint(...chars) + 'y' };
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

  /** One read, synced alone, so that what Office JS refuses fails alone. */
  async function attempt(f) {
    try {
      return await f();
    } catch (e) {
      return failed(e);
    }
  }

  /** The ways a case's control is looked for, given the paragraph holding it; the write goes through the first that finds it. */
  const LOOKUPS = [
    ['range.getContentControls', (context, p) => p.getRange('Whole').getContentControls({ types: ['PlainText'] })],
    ['body.getContentControls', (context) => context.document.body.getContentControls({ types: ['PlainText'] })],
    ['paragraph.contentControls', (context, p) => p.contentControls],
    ['range.contentControls', (context, p) => p.getRange('Whole').contentControls],
    ['body.contentControls', (context) => context.document.body.contentControls],
  ];

  /** The case's control through one way, or null where it doesn't find it; and what the way lists. */
  async function lookup(context, n, way) {
    const p = await paragraphOf(context, 'BOUND' + n);
    const controls = way(context, p);
    controls.load('items/tag,items/type');
    await context.sync();
    const mine = controls.items.filter((c) => c.tag === tagOf(n));
    return { control: mine[0] || null, listed: controls.items.length, type: mine[0] ? mine[0].type : null };
  }

  /** Every way, recorded: how many controls it lists, whether it finds the case's, and its type there. */
  async function lookups(context, region) {
    const out = {};
    for (const [name, way] of LOOKUPS) {
      out[name] = await attempt(async () => {
        const r = await lookup(context, region.n, way);
        return { listed: r.listed, found: r.control !== null, type: r.type };
      });
    }
    return out;
  }

  /** The case's control through the first way that finds it: { control, via }, or null. */
  async function controlOf(context, n) {
    for (const [name, way] of LOOKUPS) {
      try {
        const r = await lookup(context, n, way);
        if (r.control) return { control: r.control, via: name };
      } catch (e) {
        // that way isn't there, or refuses: the next
      }
    }
    return null;
  }

  /** The case's operation; what it returns. */
  async function apply(context, region) {
    if (region.op === 'bound-read') return {};
    const found = await controlOf(context, region.n);
    if (!found) throw new Error('no way found ' + tagOf(region.n));
    const r = found.control.insertText(region.text, 'Replace');
    r.load('text');
    await context.sync();
    return { via: found.via, returned: vis(r.text) };
  }

  const tidy = (xml) => xml.replace(/ w:rsid\w*="[^"]*"| w14:(?:paraId|textId)="[^"]*"| xmlns:\w+="[^"]*"/g, '');
  /** A paragraph's OOXML: the body's content without its section properties (or the paragraph's own XML, as given). */
  function bodyOf(xml) {
    if (xml.indexOf('pkg:package') < 0) return vis(tidy(xml));
    const at = xml.indexOf('pkg:name="/word/document.xml"');
    const start = xml.indexOf('<pkg:xmlData>', at);
    const end = xml.indexOf('</pkg:xmlData>', start);
    const body = /<w:body>([\s\S]*?)<\/w:body>/.exec(at < 0 || start < 0 || end < 0 ? '' : xml.slice(start, end));
    return body ? vis(tidy(body[1].replace(/<w:sectPr\b[\s\S]*<\/w:sectPr>/, ''))) : null;
  }
  /** value<n>'s content in the custom XML part's XML, as it stands there (not parsed). */
  function nodeOf(xml, n) {
    const name = '(?:[\\w.-]+:)?value' + n;
    const m = new RegExp('<' + name + '(?:\\s[^>]*)?(?:/>|>([\\s\\S]*?)</' + name + '>)').exec(xml);
    return m ? vis(m[1] === undefined ? '' : m[1]) : null;
  }
  function wellFormed(xml) {
    if (typeof DOMParser === 'undefined') return 'not checked';
    return new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagName('parsererror').length === 0;
  }

  /** The region as it stands: its paragraphs' text, the bound paragraph's OOXML, the control's text, and its node. */
  async function report(context, region) {
    const out = {};
    const whole = await regionOf(context, region.n);
    out.paragraphs = await attempt(async () => {
      const paragraphs = whole.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => vis(p.text));
    });
    out.ooxml = await attempt(async () => { const x = (await paragraphOf(context, 'BOUND' + region.n)).getOoxml(); await context.sync(); return bodyOf(x.value); });
    out.control = await attempt(async () => {
      const found = await controlOf(context, region.n);
      if (!found) return null;
      found.control.load('text,type');
      await context.sync();
      return { via: found.via, type: found.control.type, text: vis(found.control.text) };
    });
    out.node = await attempt(async () => {
      const parts = context.document.customXmlParts.getByNamespace(DATA_NS);
      parts.load('items');
      await context.sync();
      const xml = parts.items[0].getXml();
      await context.sync();
      return { value: nodeOf(xml.value, region.n), wellFormed: wellFormed(xml.value) };
    });
    return out;
  }

  /**
   * The check: `mode` 'run' or 'report'. `run(fn)` is Word.run; `progress(text)` is told each case. Each step is a
   * Word.run of its own, so that a case Office JS refuses fails alone.
   */
  async function runCheck(run, mode, progress = () => {}, host = null) {
    const result = { check: '55b', mode, host, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    result.caseCount = regions.length;
    for (const region of regions) {
      progress('case ' + region.n + ' of ' + regions.length + ': ' + region.op + ' ' + region.spec);
      const entry = { n: region.n, op: region.op, tracked: region.tracked, spec: region.spec, written: region.op === 'bound-read' ? null : vis(region.text) };
      entry.lookups = await run((context) => lookups(context, region)).catch(failed);
      try {
        if (mode === 'run') entry.result = await run((context) => apply(context, region));
      } catch (e) {
        entry.error = failed(e);
      }
      entry.region = await run((context) => report(context, region)).catch(failed);
      result.cases.push(entry);
    }
    return result;
  }

  return { parseRegion, runCheck, vis, nodeOf };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const buttons = ['run', 'report'].map((id) => document.getElementById(id));
  const start = (mode) => {
    buttons.forEach((b) => { b.disabled = true; });
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK55B.runCheck((f) => Word.run(f), mode, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.caseCount + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => buttons.forEach((b) => { b.disabled = false; }));
  };
  buttons[0].addEventListener('click', () => start('run'));
  buttons[1].addEventListener('click', () => start('report'));
}
