// Check 70 (CR-002 section 63.25; test/README.md check 70): where a paragraph's End is, when the paragraph ends in a
// bookmark, a comment's markers, a tracked deletion or insertion, a field or a text box; and which run's properties the
// text takes after a control (section 63.24's D13), on 70-paragraph-endings.docx (test/fixtures/check70/build.mjs).
// Check 69's snippet with check 70's names, each case's paragraph found by its TEXT<n> (a text box holds a paragraph of
// its own), and no document body route. WordApi 1.5 for getContentControls, 1.4 for changeTrackingMode, 1.3 for the rest.
// Every case is a region, BEGIN<n> <op> <shape>, its paragraph, END<n>. "Report only" reports every region and changes
// nothing; "Run the check" writes NEW<n> at the paragraph's End by the case's route and reports the region after it.
// Developer > Design Mode stays off; Review > Display for Review is All Markup. Nothing here depends on Script Lab but the
// last few lines, so a test can load CHECK70 with Word bound to another implementation.
const CHECK70 = (() => {
  // every character outside printable ASCII as \uXXXX, so a control character shows
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && e.code) || 'no code') + (e && e.message ? ': ' + e.message : '');
  const TOKEN = { matchCase: true, matchWholeWord: true };
  const tagOf = (n) => 'c70-' + n;

  /** A BEGIN paragraph's text as a case, or null. */
  function parseRegion(text) {
    const m = /^BEGIN(\d+) ([\w-]+) ([\w-]+)$/.exec(text.trim());
    return m ? { n: Number(m[1]), op: m[2], shape: m[3] } : null;
  }

  async function attempt(f) {
    try {
      const v = await f();
      return v === undefined ? 'undefined' : v;
    } catch (e) {
      return failed(e);
    }
  }
  /** The first hit for a token, or null. */
  async function findOrNull(context, token) {
    const hit = context.document.body.search(token, TOKEN).getFirstOrNullObject();
    hit.load('text');
    await context.sync();
    return hit.isNullObject ? null : hit;
  }
  async function paragraphOf(context, token) {
    const hit = await findOrNull(context, token);
    if (hit === null) throw new Error('not found: ' + token);
    return hit.paragraphs.getFirst();
  }
  /** The region from its BEGIN paragraph to its END paragraph, both whole. */
  const regionOf = async (context, n) => (await paragraphOf(context, 'BEGIN' + n)).getRange('Whole')
    .expandTo((await paragraphOf(context, 'END' + n)).getRange('Whole'));
  /** The case's paragraph, the one holding TEXT<n>; and for cell-body-end the table's cell. */
  async function targetOf(context, region) {
    const paragraph = await paragraphOf(context, 'TEXT' + region.n);
    paragraph.load('text');
    await context.sync();
    if (region.op !== 'cell-body-end') return { paragraph };
    const tables = (await regionOf(context, region.n)).tables;
    tables.load('items');
    await context.sync();
    return { paragraph, cell: tables.items[0].getCell(0, 0) };
  }

  /** The ways the case's control is looked for, in order (check 56: Word's contentControls lists no bound control). */
  const LOOKUPS = [
    ['body.getContentControls()', (context) => context.document.body.getContentControls()],
    ['body.contentControls', (context) => context.document.body.contentControls],
  ];
  /** The case's control through the first way that finds it: { control, via }, or null (only the D13 shapes have one). */
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
    const text = 'NEW' + region.n;
    const read = async (r) => { r.load('text'); await context.sync(); return { returned: vis(r.text) }; };
    const { paragraph: p, cell } = await targetOf(context, region);
    switch (region.op) {
      case 'paragraph-end': return read(p.insertText(text, 'End'));
      case 'end-before': return read(p.getRange('End').insertText(text, 'Before'));
      case 'end-after': return read(p.getRange('End').insertText(text, 'After'));
      case 'end-start': return read(p.getRange('End').insertText(text, 'Start'));
      case 'end-end': return read(p.getRange('End').insertText(text, 'End'));
      case 'end-replace': return read(p.getRange('End').insertText(text, 'Replace'));
      case 'content-end': return read(p.getRange('Content').insertText(text, 'End'));
      case 'cell-body-end': return read(cell.body.insertText(text, 'End'));
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
  /** The region as it stands: its paragraph's text, the region's OOXML, and the control's reads. */
  async function report(context, region) {
    const out = {};
    out.paragraph = await attempt(async () => vis((await targetOf(context, region)).paragraph.text));
    out.ooxml = await attempt(async () => { const x = (await regionOf(context, region.n)).getOoxml(); await context.sync(); return bodyOf(x.value); });
    const found = await controlOf(context, region.n).catch(() => null);
    out.control = found === null ? null : await attempt(async () => {
      const c = found.control;
      c.load('text,type');
      await context.sync();
      return { via: found.via, type: c.type, text: vis(c.text) };
    });
    if (found !== null) {
      out.showingPlaceholderText = await attempt(async () => { found.control.load('showingPlaceholderText'); await context.sync(); return found.control.showingPlaceholderText; });
    }
    return out;
  }

  /** Change tracking off: every case is untracked. */
  async function untracked(context) {
    context.document.changeTrackingMode = 'Off';
    await context.sync();
  }

  /**
   * The check: `mode` 'run' or 'report'. `run(fn)` is Word.run; `progress(text)` is told each case. Each step is a
   * Word.run of its own, so that a case Office JS refuses fails alone.
   */
  async function runCheck(run, mode, progress = () => {}, host = null) {
    const result = { check: 70, mode, host, title: null, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      result.title = paragraphs.items.length > 0 ? paragraphs.items[0].text : null;
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    result.caseCount = regions.length;
    for (const region of regions) {
      progress('case ' + region.n + ' of ' + regions.length + ': ' + region.op + ' ' + region.shape);
      const entry = { ...region };
      if (mode === 'run') {
        try {
          entry.result = await run(async (context) => {
            await untracked(context);
            return apply(context, region);
          });
        } catch (e) {
          entry.error = failed(e);
        }
      }
      entry.after = await run((context) => report(context, region)).catch(failed);
      result.cases.push(entry);
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
    CHECK70.runCheck((f) => Word.run(f), mode, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.caseCount + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => buttons.forEach((b) => { b.disabled = false; }));
  };
  buttons[0].addEventListener('click', () => start('report'));
  buttons[1].addEventListener('click', () => start('run'));
}
