// Check 58 (test/README.md check 58): what CR-002 section 58, CR-012 sections 18 and 19 and CR-002 section 32's item
// 1 left to measure, on 58-measures.docx (test/fixtures/check58/build.mjs). WordApi 1.6. Every case is a region of
// the document, BEGIN<n> <op> <args> ... END<n> (build.mjs lists the ops). 'Read' reads every case and changes
// nothing; 'Edit' accepts, rejects or fills in its edit cases only, and reads each region after. Nothing here depends
// on Script Lab but the last few lines, so a test can load CHECK58 with Word bound to another implementation.
const CHECK58 = (() => {
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && (e.code || e.message)) || String(e));
  const TOKEN = { matchCase: true, matchWholeWord: true };
  const WILD = { matchWildcards: true };
  const EDITS = new Set(['accept', 'reject', 'fill']);

  /** The queries a `search` region is searched with, paragraph by paragraph: [query, options]. */
  const QUERIES = {
    // a control showing its placeholder: Word's own text, a word of it, and a hit from plain text into it
    placeholder: (n) => [['Click or tap here to enter text.', {}], ['tap', TOKEN], ['PLAIN' + n + 'A Click', {}]],
    // a surrogate pair (U+1F600) and an Arabic word with its fathas (combining marks)
    wildcards: () => [['?', WILD], ['[!a-z ]', WILD], ['a?b', WILD], ['x?y', WILD], ['[!a-z ]@', WILD]],
  };

  /** A BEGIN paragraph's text as a case, or null; a word argument upper-cased, as the content has it. */
  function parseRegion(text) {
    const m = /^BEGIN(\d+) (\w+)(?: (.*))?$/.exec(text.trim());
    if (!m) return null;
    const args = (m[3] || '').split(' ').filter((a) => a !== '');
    return { n: Number(m[1]), op: m[2], args: m[2] === 'range' ? args.map((a) => a.toUpperCase()) : args };
  }

  async function attempt(f) {
    try {
      return await f();
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
  const controlOf = (context, tag) => context.document.body.contentControls.getByTag(tag).getFirst();
  const regionOf = async (context, n) => (await paragraphOf(context, 'BEGIN' + n)).getRange('Whole')
    .expandTo((await paragraphOf(context, 'END' + n)).getRange('Whole'));
  /** The region's paragraphs between its BEGIN and END paragraphs. */
  async function contentParagraphs(context, n) {
    const ps = (await regionOf(context, n)).paragraphs;
    ps.load('items/text');
    await context.sync();
    return ps.items.slice(1, -1);
  }

  /** Tracked changes as type, text, and their range's text and paragraph count. */
  async function changesOf(context, changes) {
    changes.load('items/type,items/text');
    await context.sync();
    const ranges = changes.items.map((c) => { const r = c.getRange(); r.load('text'); const ps = r.paragraphs; ps.load('items'); return { r, ps }; });
    await context.sync();
    return changes.items.map((c, i) => ({ type: c.type, text: vis(c.text), range: vis(ranges[i].r.text), rangeParagraphs: ranges[i].ps.items.length }));
  }

  /** The region's changes, and each paragraph's, its whole range's and its content range's. */
  async function changes(context, n) {
    const region = await regionOf(context, n);
    const out = {};
    out.region = await attempt(() => changesOf(context, region.getTrackedChanges()));
    out.paragraphs = await attempt(async () => {
      const ps = region.paragraphs;
      ps.load('items/text');
      await context.sync();
      const each = [];
      for (const p of ps.items) {
        each.push({
          text: vis(p.text),
          changes: await attempt(() => changesOf(context, p.getTrackedChanges())),
          wholeChanges: await attempt(() => changesOf(context, p.getRange('Whole').getTrackedChanges())),
          contentChanges: await attempt(() => changesOf(context, p.getRange('Content').getTrackedChanges())),
        });
      }
      return each;
    });
    return out;
  }

  /** A control showing its placeholder: its text and placeholder text, its paragraph's text and ranges, its own ranges. */
  async function placeholder(context, tag) {
    const control = controlOf(context, tag);
    const out = {};
    out.control = await attempt(async () => {
      control.load('text,placeholderText');
      const p = control.getRange('Whole').paragraphs.getFirst();
      p.load('text');
      const paragraphWhole = p.getRange('Whole');
      paragraphWhole.load('text');
      const whole = control.getRange('Whole');
      whole.load('text');
      const content = control.getRange('Content');
      content.load('text');
      await context.sync();
      return { text: vis(control.text), placeholderText: vis(control.placeholderText), paragraph: vis(p.text), paragraphWhole: vis(paragraphWhole.text), whole: vis(whole.text), content: vis(content.text) };
    });
    out.showingPlaceholderText = await attempt(async () => { control.load('showingPlaceholderText'); await context.sync(); return control.showingPlaceholderText; });
    return out;
  }

  /** After an edit: the region's paragraphs and its changes. */
  async function regionAfter(context, n) {
    const out = {};
    out.paragraphs = await attempt(async () => (await contentParagraphs(context, n)).map((p) => vis(p.text)));
    out.changes = await attempt(async () => changesOf(context, (await regionOf(context, n)).getTrackedChanges()));
    return out;
  }

  const OPS = {
    changes: (context, args, n) => changes(context, n),
    async range(context, [from, to]) {
      const range = (await find(context, from)).expandTo(await find(context, to));
      const out = {};
      out.range = await attempt(async () => { range.load('text,isEmpty'); await context.sync(); return { text: vis(range.text), isEmpty: range.isEmpty }; });
      out.paragraphs = await attempt(async () => { const ps = range.paragraphs; ps.load('items/text'); await context.sync(); return ps.items.map((p) => vis(p.text)); });
      out.whole = await attempt(async () => { const w = range.getRange('Whole'); w.load('text'); await context.sync(); return vis(w.text); });
      return out;
    },
    async cell(context, args, n) {
      const table = (await regionOf(context, n)).tables.getFirst();
      const out = {};
      out.table = await attempt(async () => {
        table.load('values');
        const whole = table.getRange('Whole');
        whole.load('text');
        await context.sync();
        return { values: table.values.map((row) => row.map(vis)), whole: vis(whole.text) };
      });
      const cell = table.getCell(0, 0);
      out.cell = await attempt(async () => {
        cell.load('value');
        cell.body.load('text');
        const whole = cell.body.getRange('Whole');
        whole.load('text');
        const ps = cell.body.paragraphs;
        ps.load('items/text');
        await context.sync();
        const wholes = ps.items.map((p) => { const r = p.getRange('Whole'); r.load('text'); return r; });
        await context.sync();
        return { value: vis(cell.value), bodyText: vis(cell.body.text), whole: vis(whole.text), paragraphs: ps.items.map((p, i) => ({ text: vis(p.text), whole: vis(wholes[i].text) })) };
      });
      return out;
    },
    placeholder: (context, [tag]) => placeholder(context, tag),
    async search(context, [id], n) {
      const paragraphs = await contentParagraphs(context, n);
      const out = {};
      for (const [query, options] of QUERIES[id](n)) {
        out[vis(query)] = await attempt(async () => {
          const hits = paragraphs.map((p) => { const h = p.search(query, options); h.load('items/text'); return h; });
          await context.sync();
          return hits.map((h) => h.items.map((x) => vis(x.text)));
        });
      }
      return out;
    },
    async accept(context, [i], n, mode) {
      if (mode !== 'edit') return changes(context, n);
      return resolvePart(context, n, Number(i), 'accept');
    },
    async reject(context, [i], n, mode) {
      if (mode !== 'edit') return changes(context, n);
      return resolvePart(context, n, Number(i), 'reject');
    },
    async fill(context, [tag, how], n, mode) {
      if (mode !== 'edit') return placeholder(context, tag);
      const out = {};
      out.inserted = await attempt(async () => {
        const r = controlOf(context, tag).insertText('FILL' + n, how === 'replace' ? 'Replace' : 'End');
        r.load('text');
        await context.sync();
        return vis(r.text);
      });
      out.after = await placeholder(context, tag);
      return out;
    },
  };

  /** The first change paragraph i of region n lists (its BEGIN paragraph is 0), accepted or rejected; the region after. */
  async function resolvePart(context, n, i, verb) {
    const out = {};
    out.resolved = await attempt(async () => {
      const ps = (await regionOf(context, n)).paragraphs;
      ps.load('items');
      await context.sync();
      const list = ps.items[i].getTrackedChanges();
      list.load('items/type,items/text');
      await context.sync();
      if (list.items.length === 0) return 'none listed';
      const change = list.items[0];
      const what = { type: change.type, text: vis(change.text), listed: list.items.length };
      change[verb]();
      await context.sync();
      return what;
    });
    out.after = await regionAfter(context, n);
    return out;
  }

  /** The check: `run(fn)` is Word.run; `label` names the view; `mode` is 'read' or 'edit'; `progress(text)` is told each case. */
  async function runCheck(run, label, mode = 'read', progress = () => {}, host = null) {
    const result = { check: 58, label, mode, host, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    for (const region of regions) {
      if (mode === 'edit' && !EDITS.has(region.op)) continue;
      progress('case ' + region.n + ' of ' + regions.length + ': ' + region.op);
      const op = OPS[region.op];
      const entry = { n: region.n, op: region.op, args: region.args };
      entry.reads = op ? await run((context) => op(context, region.args, region.n, mode)).catch(failed) : 'error: no op ' + region.op;
      result.cases.push(entry);
    }
    return result;
  }

  return { parseRegion, runCheck, vis, QUERIES };
})();

if (typeof document !== 'undefined' && document.getElementById('read')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const buttons = [document.getElementById('read'), document.getElementById('edit')];
  const start = (mode) => {
    for (const b of buttons) b.disabled = true;
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK58.runCheck((f) => Word.run(f), document.getElementById('label').value, mode, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.cases.length + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => { for (const b of buttons) b.disabled = false; });
  };
  buttons[0].addEventListener('click', () => start('read'));
  buttons[1].addEventListener('click', () => start('edit'));
}
