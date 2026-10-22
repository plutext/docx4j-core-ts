// Check 54 (CR-014 phase 3, and phase 2's choices check 52 didn't measure, CR-014 section 10.6; test/README.md
// check 54): Office JS's writes of a Word.Range, on 54-writes.docx (test/fixtures/check54/build.mjs). WordApi 1.6.
// As check 52's: every case is a region, BEGIN<n> <kind> <op> <tracking> <range> ... END<n>, so one run measures
// every case on content no other case has touched. "Run the check" applies each case's operation in its region,
// with change tracking as the case says, then reports the region; "Report only" reports every region and changes
// nothing (for after Review > Accept All Changes). Nothing here depends on Script Lab but the last few lines, so a
// test can load CHECK54 with Word bound to another implementation.
const CHECK54 = (() => {
  // every character outside printable ASCII as \uXXXX (a surrogate pair as two), so a control character shows
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && (e.code || e.message)) || String(e));
  const TOKEN = { matchCase: true, matchWholeWord: true };
  /** What the ooxml operation puts in place of the range: a flat OPC package whose body is one paragraph, NEWOOXML. */
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const PARAGRAPH_OOXML = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
    + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
    + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    + '</pkg:xmlData></pkg:part>'
    + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>'
    + '<w:document xmlns:w="' + W + '"><w:body><w:p><w:r><w:t>NEWOOXML</w:t></w:r></w:p></w:body></w:document>'
    + '</pkg:xmlData></pkg:part></pkg:package>';

  /** A BEGIN paragraph's text as a case, or null. The range's tokens are in lower case there; the content's are upper case. */
  function parseRegion(text) {
    const m = /^BEGIN(\d+) (\w+) ([\w-]+) (untracked|tracked) (.*)$/.exec(text.trim());
    if (!m) return null;
    const region = { n: Number(m[1]), kind: m[2], op: m[3], tracked: m[4] === 'tracked', spec: m[5] };
    const spec = region.spec;
    let r;
    if ((r = /^from (\S+) to (\S+)$/.exec(spec))) region.range = { from: r[1].toUpperCase(), to: r[2].toUpperCase() };
    else if ((r = /^from (\S+) to-mark (\S+)$/.exec(spec))) region.range = { from: r[1].toUpperCase(), toMark: r[2].toUpperCase() };
    else if ((r = /^from (\S+) to-table (\S+)$/.exec(spec))) region.range = { from: r[1].toUpperCase(), toTable: r[2].toUpperCase() };
    else if ((r = /^start (\S+) to-mark (\S+)$/.exec(spec))) region.range = { start: r[1].toUpperCase(), toMark: r[2].toUpperCase() };
    else if ((r = /^whole (\S+)$/.exec(spec))) region.range = { whole: r[1].toUpperCase() };
    else if ((r = /^word (\S+)$/.exec(spec))) region.range = { word: r[1].toUpperCase() };
    else throw new Error('BEGIN' + region.n + ': no range in "' + spec + '"');
    return region;
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

  /** The case's range, made afresh. */
  async function rangeOf(context, region) {
    const r = region.range;
    if (r.whole) return (await paragraphOf(context, r.whole)).getRange('Whole');
    if (r.word) return find(context, r.word);
    const to = r.toMark ? (await paragraphOf(context, r.toMark)).getRange('Whole')
      : r.toTable ? (await find(context, r.toTable)).parentTable.getRange('Whole')
        : await find(context, r.to);
    const from = r.start ? (await paragraphOf(context, r.start)).getRange('Start') : await find(context, r.from);
    return from.expandTo(to);
  }
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

  /** The case's operation, on its range made afresh; what the operation returns. */
  async function apply(context, region) {
    const returned = async (r) => { r.load('text'); await context.sync(); return { returned: vis(r.text) }; };
    if (region.op === 'own-delete') {
      // a paragraph inserted under tracking by whoever runs this, then a range from the heading into it deleted
      (await paragraphOf(context, 'HEAD' + region.n + 'A')).insertParagraph('OWN' + region.n + 'A OWN' + region.n + 'B', 'After');
      await context.sync();
      const range = await rangeOf(context, region);
      range.delete();
      await context.sync();
      return {};
    }
    const range = await rangeOf(context, region);
    switch (region.op) {
      case 'delete': range.delete(); await context.sync(); return {};
      case 'replace': return returned(range.insertText('NEWTEXT', 'Replace'));
      case 'paragraph-before': return returned(range.insertParagraph('NEWPARA', 'Before'));
      case 'paragraph-after': return returned(range.insertParagraph('NEWPARA', 'After'));
      case 'italic': range.font.italic = true; await context.sync(); return {};
      case 'size14': range.font.size = 14; await context.sync(); return {};
      case 'underline': range.font.underline = 'Single'; await context.sync(); return {};
      case 'style': range.styleBuiltIn = 'Heading2'; await context.sync(); return {};
      case 'hyperlink': range.hyperlink = 'https://example.com/check54'; await context.sync(); return {};
      case 'ooxml': return returned(range.insertOoxml(PARAGRAPH_OOXML, 'Replace'));
      case 'control': {
        const control = range.insertContentControl();
        control.tag = 'new54-' + region.n;
        control.load('text,type');
        await context.sync();
        return { returned: vis(control.text), type: control.type };
      }
      default: throw new Error('no operation ' + region.op);
    }
  }

  /** The region as it stands: its text, paragraphs, tables, tracked changes, comments, controls and bookmarks. */
  async function report(context, region) {
    const out = {};
    const whole = await regionOf(context, region.n);
    out.text = await attempt(async () => { whole.load('text'); await context.sync(); return vis(whole.text); });
    out.paragraphs = await attempt(async () => {
      const paragraphs = whole.paragraphs;
      paragraphs.load('items/text,items/style,items/styleBuiltIn,items/alignment,items/isListItem,items/tableNestingLevel');
      await context.sync();
      const lists = paragraphs.items.map((p) => { const l = p.listItemOrNullObject; l.load('listString,level'); return l; });
      await context.sync();
      return paragraphs.items.map((p, i) => {
        const entry = { text: vis(p.text), style: p.style, styleBuiltIn: p.styleBuiltIn, alignment: p.alignment, tableNestingLevel: p.tableNestingLevel };
        if (!lists[i].isNullObject) entry.list = { listString: vis(lists[i].listString), level: lists[i].level };
        return entry;
      });
    });
    out.tables = await attempt(async () => {
      const tables = whole.tables;
      tables.load('items/rowCount,items/values');
      await context.sync();
      return tables.items.map((x) => ({ rowCount: x.rowCount, values: x.values.map((row) => row.map(vis)) }));
    });
    out.trackedChanges = await attempt(async () => {
      const changes = whole.getTrackedChanges();
      changes.load('items/type,items/text');
      await context.sync();
      const ranges = changes.items.map((c) => { const r = c.getRange(); r.load('text'); return r; });
      await context.sync();
      return changes.items.map((c, i) => ({ type: c.type, text: vis(c.text), range: vis(ranges[i].text) }));
    });
    out.comments = await attempt(async () => {
      const comments = whole.getComments();
      comments.load('items/content');
      await context.sync();
      const ranges = comments.items.map((c) => { const r = c.getRange(); r.load('text'); return r; });
      await context.sync();
      return comments.items.map((c, i) => ({ content: vis(c.content), range: vis(ranges[i].text) }));
    });
    out.contentControls = await attempt(async () => {
      const controls = whole.contentControls;
      controls.load('items/tag,items/text,items/type');
      await context.sync();
      return controls.items.map((c) => ({ tag: c.tag, type: c.type, text: vis(c.text) }));
    });
    out.bookmarks = await attempt(async () => {
      const names = whole.getBookmarks(true, true);
      await context.sync();
      return names.value;
    });
    out.hyperlink = await attempt(async () => { whole.load('hyperlink'); await context.sync(); return whole.hyperlink; });
    return out;
  }

  /**
   * The check: `mode` 'run' or 'report'. `run(fn)` is Word.run; `progress(text)` is told each case. Each step is a
   * Word.run of its own, so that a case Office JS refuses fails alone.
   */
  async function runCheck(run, mode, progress = () => {}, host = null) {
    const result = { check: 54, mode, host, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    result.caseCount = regions.length;
    for (const region of regions) {
      progress('case ' + region.n + ' of ' + regions.length + ': ' + region.kind + ' ' + region.op + (region.tracked ? ' tracked' : ''));
      const entry = { n: region.n, kind: region.kind, op: region.op, tracked: region.tracked, spec: region.spec };
      try {
        if (mode === 'run') {
          await run(async (context) => {
            context.document.changeTrackingMode = region.tracked ? 'TrackAll' : 'Off';
            await context.sync();
          });
          if (region.op !== 'own-delete') {
            entry.range = await run(async (context) => { const r = await rangeOf(context, region); r.load('text'); await context.sync(); return vis(r.text); }).catch(failed);
          }
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

  return { parseRegion, runCheck, vis, PARAGRAPH_OOXML };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const buttons = ['run', 'report'].map((id) => document.getElementById(id));
  const start = (mode) => {
    buttons.forEach((b) => { b.disabled = true; });
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK54.runCheck((f) => Word.run(f), mode, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.caseCount + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => buttons.forEach((b) => { b.disabled = false; }));
  };
  buttons[0].addEventListener('click', () => start('run'));
  buttons[1].addEventListener('click', () => start('report'));
}
