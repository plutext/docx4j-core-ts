// Check 55 (CR-002 section 60; test/README.md check 55): what Office JS's writers do with the characters its readers
// give, on 55-characters.docx (test/fixtures/check55/build.mjs). WordApi 1.6 (getTrackedChanges; the rest is 1.4).
// As check 54's: every case is a region, BEGIN<n> <op> <tracking> <characters> ... END<n>, so one run measures every
// case on content no other case has touched. "Run the check" writes 'x', the case's characters and 'y' with the case's
// operation, with change tracking as the case says, then reports the region; "Report only" reports every region and
// changes nothing (for after the saved file is opened again). Nothing here depends on Script Lab but the last few
// lines, so a test can load CHECK55 with Word bound to another implementation.
const CHECK55 = (() => {
  // every character outside printable ASCII as \uXXXX (a surrogate pair as two), so a control character shows
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && e.code) || 'no code') + (e && e.message ? ': ' + e.message : '');
  const TOKEN = { matchCase: true, matchWholeWord: true };
  const DATA_NS = 'urn:check55';

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
  const controlOf = (context, n) => context.document.body.contentControls.getByTag('c55-' + n).getFirst();

  /** One read, synced alone, so that what Office JS refuses fails alone. */
  async function attempt(f) {
    try {
      return await f();
    } catch (e) {
      return failed(e);
    }
  }

  /** The case's operation; what it returns. */
  async function apply(context, region) {
    const s = region.text;
    const returned = async (r) => { r.load('text'); await context.sync(); return { returned: vis(r.text) }; };
    switch (region.op) {
      case 'text-end': return returned((await paragraphOf(context, 'TARGET' + region.n)).insertText(s, 'End'));
      case 'paragraph-after': return returned((await paragraphOf(context, 'TARGET' + region.n)).insertParagraph(s, 'After'));
      case 'comment': {
        (await find(context, 'TARGET' + region.n)).insertComment(s);
        await context.sync();
        return {};
      }
      case 'comment-content': {
        const hit = await find(context, 'TARGET' + region.n);
        hit.insertComment('COMMENT' + region.n);
        await context.sync();
        const comments = hit.getComments();
        comments.load('items');
        await context.sync();
        comments.items[0].content = s;
        await context.sync();
        return {};
      }
      case 'bound-multi': case 'bound-single': return returned(controlOf(context, region.n).insertText(s, 'Replace'));
      default: throw new Error('no operation ' + region.op);
    }
  }

  /** A part's XML in a flat OPC package, or null. */
  function partOf(flat, name) {
    const at = flat.indexOf('pkg:name="' + name + '"');
    if (at < 0) return null;
    const start = flat.indexOf('<pkg:xmlData>', at);
    const end = flat.indexOf('</pkg:xmlData>', start);
    return start < 0 || end < 0 ? null : flat.slice(start + '<pkg:xmlData>'.length, end);
  }
  const tidy = (xml) => xml.replace(/ w:rsid\w*="[^"]*"| w14:(?:paraId|textId)="[^"]*"| xmlns:\w+="[^"]*"/g, '');
  /**
   * What a paragraph's OOXML holds: the body's content without its section properties, and the comments. Office JS
   * gives a flat OPC package; an implementation that gives the paragraph's own XML is read as the body's content.
   */
  function ooxmlOf(xml) {
    if (xml.indexOf('pkg:package') < 0) return { body: vis(tidy(xml)), comments: [] };
    const document = partOf(xml, '/word/document.xml') || '';
    const body = /<w:body>([\s\S]*?)<\/w:body>/.exec(document);
    const comments = (partOf(xml, '/word/comments.xml') || '').match(/<w:comment\b[\s\S]*?<\/w:comment>/g) || [];
    return {
      body: body ? vis(tidy(body[1].replace(/<w:sectPr\b[\s\S]*<\/w:sectPr>/, ''))) : null,
      comments: comments.map((c) => vis(tidy(c))),
    };
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

  /** The region as it stands: its text, paragraphs, OOXML, comments, tracked changes; a bound control and its node. */
  async function report(context, region) {
    const out = {};
    const whole = await regionOf(context, region.n);
    out.text = await attempt(async () => { whole.load('text'); await context.sync(); return vis(whole.text); });
    const paragraphs = whole.paragraphs;
    out.paragraphs = await attempt(async () => {
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => vis(p.text));
    });
    // each paragraph's OOXML but the BEGIN and END paragraphs', one at a time
    out.ooxml = await attempt(async () => {
      const inner = paragraphs.items.slice(1, -1).map((p) => p.getOoxml());
      await context.sync();
      return inner.map((x) => ooxmlOf(x.value));
    });
    out.comments = await attempt(async () => {
      const comments = whole.getComments();
      comments.load('items/content');
      await context.sync();
      return comments.items.map((c) => vis(c.content));
    });
    out.trackedChanges = await attempt(async () => {
      const changes = whole.getTrackedChanges();
      changes.load('items/type,items/text');
      await context.sync();
      return changes.items.map((c) => ({ type: c.type, text: vis(c.text) }));
    });
    if (region.op.startsWith('bound-')) {
      out.control = await attempt(async () => { const c = controlOf(context, region.n); c.load('text'); await context.sync(); return vis(c.text); });
      out.node = await attempt(async () => {
        const parts = context.document.customXmlParts.getByNamespace(DATA_NS);
        parts.load('items');
        await context.sync();
        const xml = parts.items[0].getXml();
        await context.sync();
        return { value: nodeOf(xml.value, region.n), wellFormed: wellFormed(xml.value) };
      });
    }
    return out;
  }

  /**
   * The check: `mode` 'run' or 'report'. `run(fn)` is Word.run; `progress(text)` is told each case. Each step is a
   * Word.run of its own, so that a case Office JS refuses fails alone.
   */
  async function runCheck(run, mode, progress = () => {}, host = null) {
    const result = { check: 55, mode, host, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    result.caseCount = regions.length;
    for (const region of regions) {
      progress('case ' + region.n + ' of ' + regions.length + ': ' + region.op + ' ' + region.spec + (region.tracked ? ' tracked' : ''));
      const entry = { n: region.n, op: region.op, tracked: region.tracked, spec: region.spec, written: vis(region.text) };
      try {
        if (mode === 'run') {
          await run(async (context) => {
            context.document.changeTrackingMode = region.tracked ? 'TrackAll' : 'Off';
            await context.sync();
          });
          entry.result = await run((context) => apply(context, region));
        }
      } catch (e) {
        entry.error = failed(e);
      }
      entry.region = await run((context) => report(context, region)).catch(failed);
      result.cases.push(entry);
    }
    if (mode === 'run') await run(async (context) => { context.document.changeTrackingMode = 'Off'; await context.sync(); });
    return result;
  }

  return { parseRegion, runCheck, vis, nodeOf, ooxmlOf };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const buttons = ['run', 'report'].map((id) => document.getElementById(id));
  const start = (mode) => {
    buttons.forEach((b) => { b.disabled = true; });
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK55.runCheck((f) => Word.run(f), mode, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.caseCount + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => buttons.forEach((b) => { b.disabled = false; }));
  };
  buttons[0].addEventListener('click', () => start('run'));
  buttons[1].addEventListener('click', () => start('report'));
}
