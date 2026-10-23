// Check 59 (test/README.md check 59, CR-002 section 32's K5): where Word's wildcards match a paragraph mark, on
// 59-wildcard-marks.docx (test/fixtures/check59/build.mjs). WordApi 1.6. Read only: every case is a region of the
// document, BEGIN<n> <shape> ... END<n>, searched with wildcards in four scopes. Nothing here depends on Script Lab but
// the last few lines, so a test can load CHECK59 with Word bound to another implementation.
const CHECK59 = (() => {
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && (e.code || e.message)) || String(e));
  const WILD = { matchWildcards: true };
  /** Across the region's content: a mark between two letters, Word's paragraph mark code alone. */
  const REGION_QUERIES = ['Q?Z', 'Q^13Z', 'Q*Z', '^13'];
  /** Within a paragraph, a cell's body, a control: every character, every one but a capital, the mark code. */
  const PART_QUERIES = ['?', '[!A-Z]', '^13'];

  function parseRegion(text) {
    const m = /^BEGIN(\d+) ([\w-]+)$/.exec(text.trim());
    return m ? { n: Number(m[1]), shape: m[2] } : null;
  }

  async function attempt(f) {
    try {
      return await f();
    } catch (e) {
      return failed(e);
    }
  }
  async function paragraphOf(context, token) {
    const hit = context.document.body.search(token, { matchCase: true, matchWholeWord: true }).getFirstOrNullObject();
    hit.load('text');
    await context.sync();
    if (hit.isNullObject) throw new Error('not found: ' + token);
    return hit.paragraphs.getFirst();
  }
  const regionOf = async (context, n) => (await paragraphOf(context, 'BEGIN' + n)).getRange('Whole')
    .expandTo((await paragraphOf(context, 'END' + n)).getRange('Whole'));

  /** Each query's hits in a scope, as their text. */
  async function searched(context, scope, queries) {
    const out = {};
    for (const q of queries) {
      out[vis(q)] = await attempt(async () => {
        const hits = scope.search(q, WILD);
        hits.load('items/text');
        await context.sync();
        return hits.items.map((h) => vis(h.text));
      });
    }
    return out;
  }

  async function read(context, n) {
    const region = await regionOf(context, n);
    const out = {};
    out.region = await searched(context, region, REGION_QUERIES);
    out.paragraphs = await attempt(async () => {
      const ps = region.paragraphs;
      ps.load('items/text');
      await context.sync();
      const each = [];
      for (const p of ps.items.slice(1, -1)) each.push({ text: vis(p.text), hits: await searched(context, p, PART_QUERIES) });
      return each;
    });
    const tables = region.tables;
    tables.load('items');
    const controls = region.contentControls;
    controls.load('items');
    await context.sync();
    if (tables.items.length > 0) {
      out.cell = await attempt(() => searched(context, tables.items[0].getCell(0, 0).body, PART_QUERIES));
      out.table = await attempt(() => searched(context, tables.items[0].getRange('Whole'), [...PART_QUERIES, 'Q?Z', 'Q*Z']));
    }
    if (controls.items.length > 0) out.control = await attempt(() => searched(context, controls.items[0], [...PART_QUERIES, 'Q?Z', 'Q*Z']));
    return out;
  }

  /** The check: `run(fn)` is Word.run; `label` names the view; `progress(text)` is told each case. */
  async function runCheck(run, label, progress = () => {}, host = null) {
    const result = { check: 59, label, host, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    for (const region of regions) {
      progress('case ' + region.n + ' of ' + regions.length + ': ' + region.shape);
      result.cases.push({ n: region.n, shape: region.shape, reads: await run((context) => read(context, region.n)).catch(failed) });
    }
    return result;
  }

  return { parseRegion, runCheck, vis, REGION_QUERIES, PART_QUERIES };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const button = document.getElementById('run');
  button.addEventListener('click', () => {
    button.disabled = true;
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK59.runCheck((f) => Word.run(f), document.getElementById('label').value, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.cases.length + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => { button.disabled = false; });
  });
}
