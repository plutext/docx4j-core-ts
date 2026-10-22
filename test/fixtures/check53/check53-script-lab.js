// Check 53 (CR-014 phase 1's choices where Word was not measured, section 9; test/README.md check 53), on
// 53-reads.docx (test/fixtures/check53/build.mjs). WordApi 1.6. Read only: every case is a region of the
// document, BEGIN<n> <op> <args> ... END<n>, and the snippet reads what the op names (build.mjs lists them).
// Run it once in Simple Markup and once in All Markup, the label saying which. Nothing here depends on Script
// Lab but the last few lines, so a test can load CHECK53 with Word bound to another implementation.
const CHECK53 = (() => {
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && (e.code || e.message)) || String(e));
  const TOKEN = { matchCase: true, matchWholeWord: true };

  /** A BEGIN paragraph's text as a case, or null; the arguments upper-cased where they are words of the content. */
  function parseRegion(text) {
    const m = /^BEGIN(\d+) (\w+)(?: (.*))?$/.exec(text.trim());
    if (!m) return null;
    const args = (m[3] || '').split(' ').filter((a) => a !== '');
    return { n: Number(m[1]), op: m[2], args: m[2] === 'control' || m[2] === 'placeholder' ? args : args.map((a) => a.toUpperCase()) };
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

  /** A collapsed or short range: its text and the paragraph it is in. */
  async function where(context, range) {
    range.load('text,isEmpty');
    const p = range.paragraphs.getFirstOrNullObject();
    p.load('text');
    await context.sync();
    return { text: vis(range.text), isEmpty: range.isEmpty, paragraph: p.isNullObject ? null : vis(p.text) };
  }

  /** Every read a range has, each synced alone. */
  async function rangeReads(context, range) {
    const out = {};
    out.range = await attempt(async () => {
      range.load('text,isEmpty,style,styleBuiltIn');
      await context.sync();
      return { text: vis(range.text), isEmpty: range.isEmpty, style: range.style, styleBuiltIn: range.styleBuiltIn };
    });
    out.font = await attempt(async () => {
      range.font.load('bold,italic,underline,name,size,color,highlightColor');
      await context.sync();
      const f = range.font;
      return { bold: f.bold, italic: f.italic, underline: f.underline, name: f.name, size: f.size, color: f.color, highlightColor: f.highlightColor };
    });
    for (const location of ['Whole', 'Content', 'Start', 'End', 'After']) out[location] = await attempt(() => where(context, range.getRange(location)));
    out.paragraphs = await attempt(async () => {
      const ps = range.paragraphs;
      ps.load('items/text,items/style,items/styleBuiltIn');
      await context.sync();
      return ps.items.map((p) => ({ text: vis(p.text), style: p.style, styleBuiltIn: p.styleBuiltIn }));
    });
    out.tables = await attempt(async () => {
      const tables = range.tables;
      tables.load('items/rowCount,items/values');
      await context.sync();
      return tables.items.map((x) => ({ rowCount: x.rowCount, values: x.values.map((row) => row.map(vis)) }));
    });
    out.contentControls = await attempt(async () => {
      const controls = range.contentControls;
      controls.load('items/tag,items/text');
      await context.sync();
      return controls.items.map((c) => ({ tag: c.tag, text: vis(c.text) }));
    });
    return out;
  }

  /** Tracked changes as type, text, and their range's text and paragraph count. */
  async function changesOf(context, changes) {
    changes.load('items/type,items/text');
    await context.sync();
    const ranges = changes.items.map((c) => { const r = c.getRange(); r.load('text'); const ps = r.paragraphs; ps.load('items'); return { r, ps }; });
    await context.sync();
    return changes.items.map((c, i) => ({ type: c.type, text: vis(c.text), range: vis(ranges[i].r.text), rangeParagraphs: ranges[i].ps.items.length }));
  }

  const OPS = {
    async range(context, [from, to]) {
      const range = (await find(context, from)).expandTo(await find(context, to));
      return rangeReads(context, range);
    },
    async para(context, [token]) {
      const p = await paragraphOf(context, token);
      const out = {};
      out.paragraph = await attempt(async () => { p.load('text,style,styleBuiltIn'); await context.sync(); return { text: vis(p.text), style: p.style, styleBuiltIn: p.styleBuiltIn }; });
      out.whole = await attempt(async () => { const w = p.getRange('Whole'); w.load('text,style,styleBuiltIn'); await context.sync(); return { text: vis(w.text), style: w.style, styleBuiltIn: w.styleBuiltIn }; });
      out.trackedChanges = await attempt(() => changesOf(context, p.getTrackedChanges()));
      out.wholeTrackedChanges = await attempt(() => changesOf(context, p.getRange('Whole').getTrackedChanges()));
      return out;
    },
    async control(context, [tag]) {
      const control = controlOf(context, tag);
      const out = {};
      out.control = await attempt(async () => {
        control.load('text,type');
        const ps = control.paragraphs;
        ps.load('items/text');
        await context.sync();
        return { text: vis(control.text), type: control.type, paragraphs: ps.items.map((p) => vis(p.text)) };
      });
      for (const location of ['Whole', 'Content', 'Start', 'End']) {
        out[location] = await attempt(async () => {
          const r = control.getRange(location);
          r.load('text');
          const ps = r.paragraphs;
          ps.load('items/text');
          await context.sync();
          return { text: vis(r.text), paragraphs: ps.items.map((p) => vis(p.text)) };
        });
      }
      return out;
    },
    async table(context, [token]) {
      const table = (await find(context, token)).parentTable;
      const out = {};
      for (const location of ['Whole', 'Start', 'End', 'After']) out[location] = await attempt(() => where(context, table.getRange(location)));
      return out;
    },
    async cellbody(context, [token]) {
      const body = (await find(context, token)).parentTableCell.body;
      const out = {};
      for (const location of ['Whole', 'Content', 'Start', 'End']) {
        out[location] = await attempt(async () => {
          const r = body.getRange(location);
          const reads = await rangeReads(context, r);
          return { text: reads.range.text, paragraphs: reads.paragraphs, tables: reads.tables, contentControls: reads.contentControls };
        });
      }
      return out;
    },
    async changes(context, args, n) {
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
          });
        }
        return each;
      });
      return out;
    },
    async stories(context, [token]) {
      return attempt(async () => {
        const notes = context.document.body.footnotes;
        notes.load('items');
        await context.sync();
        const across = (await find(context, token)).expandTo(notes.items[0].body.getRange('Whole'));
        across.load('text');
        await context.sync();
        return { text: vis(across.text) };
      });
    },
    async body(context) {
      const body = context.document.body;
      const out = {};
      out.text = await attempt(async () => { body.load('text'); await context.sync(); return { length: body.text.length, tail: vis(body.text.slice(-16)) }; });
      for (const location of ['Whole', 'Content', 'Start', 'End']) {
        out[location] = await attempt(async () => {
          const r = body.getRange(location);
          r.load('text');
          const first = r.paragraphs.getFirst();
          first.load('text');
          await context.sync();
          return { length: r.text.length, tail: vis(r.text.slice(-16)), firstParagraph: vis(first.text) };
        });
      }
      return out;
    },
    async placeholder(context, [tag]) {
      const control = controlOf(context, tag);
      const out = {};
      out.control = await attempt(async () => {
        control.load('text,placeholderText');
        const p = control.paragraphs.getFirst();
        p.load('text');
        const whole = control.getRange('Whole');
        whole.load('text');
        await context.sync();
        return { text: vis(control.text), placeholderText: vis(control.placeholderText), paragraph: vis(p.text), whole: vis(whole.text) };
      });
      out.showingPlaceholderText = await attempt(async () => { control.load('showingPlaceholderText'); await context.sync(); return control.showingPlaceholderText; });
      return out;
    },
  };

  /** The check: `run(fn)` is Word.run; `label` names the view; `progress(text)` is told each case. */
  async function runCheck(run, label, progress = () => {}, host = null) {
    const result = { check: 53, label, host, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    for (const region of regions) {
      progress('case ' + region.n + ' of ' + regions.length + ': ' + region.op);
      const op = OPS[region.op];
      const entry = { n: region.n, op: region.op, args: region.args };
      entry.reads = op ? await run((context) => op(context, region.args, region.n)).catch(failed) : 'error: no op ' + region.op;
      result.cases.push(entry);
    }
    return result;
  }

  return { parseRegion, runCheck, vis };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const button = document.getElementById('run');
  button.addEventListener('click', () => {
    button.disabled = true;
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK53.runCheck((f) => Word.run(f), document.getElementById('label').value, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.cases.length + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => { button.disabled = false; });
  });
}
