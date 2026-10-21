// Check 52 (CR-014 phase 0; test/README.md check 52): Office JS's ranges across paragraphs, on
// 52-ranges.docx (test/fixtures/check52/build.mjs). WordApi 1.6. Every case is a region of the document,
// BEGIN<n> <kind> <op> <tracking> <range> ... END<n>, so one run measures every case on content no other
// case has touched. "Run the check" reads each kind's range once, applies each case's operation in its
// region with change tracking as the case says, then reports the region; "Report only" reports every
// region and changes nothing (for after Review > Accept All Changes). Cases 86 to 95 are CR-002 section
// 55's (Word's table rule): Office JS's edits of a cell, its body or its paragraph, not of a range. Nothing
// here depends on Script Lab but the last few lines, so a test can load CHECK52 with Word bound to another
// implementation.
const CHECK52 = (() => {
  // every character outside printable ASCII as \uXXXX (a surrogate pair as two), so a control character shows
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && (e.code || e.message)) || String(e));
  const TOKEN = { matchCase: true, matchWholeWord: true };
  /** What ooxml-replace puts in a cell: a flat OPC package whose body is a table of one cell, NEWOOXML, and nothing after it. */
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const TABLE_OOXML = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
    + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
    + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    + '</pkg:xmlData></pkg:part>'
    + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>'
    + '<w:document xmlns:w="' + W + '"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
    + '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>NEWOOXML</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>'
    + '</pkg:xmlData></pkg:part></pkg:package>';

  /** A BEGIN paragraph's text as a case, or null. The range's tokens are in lower case there; the content's are upper case. */
  function parseRegion(text) {
    const m = /^BEGIN(\d+) (\w+) ([\w-]+) (untracked|tracked)(?: (.*))?$/.exec(text.trim());
    if (!m) return null;
    const region = { n: Number(m[1]), kind: m[2], op: m[3], tracked: m[4] === 'tracked', spec: m[5] || '' };
    const spec = region.spec;
    let r;
    if ((r = /^from (\S+) to (\S+)$/.exec(spec))) region.range = { from: r[1].toUpperCase(), to: r[2].toUpperCase() };
    else if ((r = /^from (\S+) to-mark (\S+)$/.exec(spec))) region.range = { from: r[1].toUpperCase(), toMark: r[2].toUpperCase() };
    else if ((r = /^whole (\S+)$/.exec(spec))) region.range = { whole: r[1].toUpperCase() };
    else if ((r = /^control (\S+)$/.exec(spec))) region.range = { control: r[1] };
    else if ((r = /^cell (\S+)$/.exec(spec))) region.range = { cell: r[1].toUpperCase() };
    else if ((r = /^paragraph (\S+)$/.exec(spec))) region.range = { paragraph: r[1].toUpperCase() };
    else if ((r = /^queries (.*)$/.exec(spec))) region.queries = r[1].split(';').map(queryText);
    else if (spec !== '') throw new Error('BEGIN' + region.n + ': no range in "' + spec + '"');
    return region;
  }
  /** A query as written in the marker, its tokens upper-cased and Word's codes (^p) kept. */
  const queryText = (q) => q.split(/(\^.)/).map((s, i) => (i % 2 ? s : s.toUpperCase())).join('');

  async function find(context, token) {
    const hit = context.document.body.search(token, TOKEN).getFirstOrNullObject();
    hit.load('text');
    await context.sync();
    if (hit.isNullObject) throw new Error('not found: ' + token);
    return hit;
  }
  /** The whole of the paragraph holding a token, its mark included. */
  const paragraphOf = async (context, token) => (await find(context, token)).paragraphs.getFirst().getRange('Whole');

  /** The table cell holding a token. */
  const cellOf = async (context, token) => (await find(context, token)).parentTableCell;

  /** The case's range, made afresh. */
  async function rangeOf(context, region) {
    const r = region.range;
    if (r.control) return context.document.body.contentControls.getByTag(r.control).getFirst().getRange('Content');
    if (r.whole) return paragraphOf(context, r.whole);
    if (r.cell) return (await cellOf(context, r.cell)).body.getRange('Whole');
    if (r.paragraph) return paragraphOf(context, r.paragraph);
    const from = await find(context, r.from);
    return from.expandTo(r.toMark ? await paragraphOf(context, r.toMark) : await find(context, r.to));
  }
  /** The region from its BEGIN paragraph to its END paragraph, both whole. */
  const regionOf = async (context, n) => (await paragraphOf(context, 'BEGIN' + n)).expandTo(await paragraphOf(context, 'END' + n));

  /** One read, synced alone, so that what Office JS refuses fails alone. */
  async function attempt(f) {
    try {
      return await f();
    } catch (e) {
      return failed(e);
    }
  }

  /** Everything a range reads, before anything is written: CR-014 section 2, items 5 and 7. */
  async function reads(context, region) {
    const out = {};
    const range = await rangeOf(context, region);
    out.range = await attempt(async () => {
      range.load('text,isEmpty,style,styleBuiltIn');
      await context.sync();
      return { text: vis(range.text), isEmpty: range.isEmpty, style: range.style, styleBuiltIn: range.styleBuiltIn };
    });
    out.font = await attempt(async () => {
      range.font.load('bold,italic,underline,name,size');
      await context.sync();
      const f = range.font;
      return { bold: f.bold, italic: f.italic, underline: f.underline, name: f.name, size: f.size };
    });
    for (const location of ['Whole', 'Start', 'End', 'Content', 'After']) {
      out[location] = await attempt(async () => {
        const part = range.getRange(location);
        part.load('text,isEmpty');
        const first = part.paragraphs.getFirstOrNullObject();
        first.load('text');
        await context.sync();
        return { text: vis(part.text), isEmpty: part.isEmpty, paragraph: first.isNullObject ? null : vis(first.text) };
      });
    }
    out.paragraphs = await attempt(async () => {
      const paragraphs = range.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => vis(p.text));
    });
    out.tables = await attempt(async () => {
      const tables = range.tables;
      tables.load('items/rowCount');
      await context.sync();
      return tables.items.map((x) => x.rowCount);
    });
    out.marks = await attempt(async () => {
      const hits = range.search('^p', { matchCase: true });
      hits.load('items/text');
      await context.sync();
      return hits.items.map((h) => vis(h.text));
    });
    if (region.range.to) {
      // expandTo with its ends the other way round
      out.swapped = await attempt(async () => {
        const swapped = (await find(context, region.range.to)).expandTo(await find(context, region.range.from));
        swapped.load('text');
        await context.sync();
        return vis(swapped.text);
      });
    }
    return out;
  }

  /** The case's operation, on its range made afresh; what the operation returns. */
  async function apply(context, region) {
    const returned = async (r) => { r.load('text'); await context.sync(); return { returned: vis(r.text) }; };
    // CR-002 section 55's: a cell's body, its paragraph or its control
    switch (region.op) {
      case 'table-end': {
        const table = (await cellOf(context, region.range.cell)).body.insertTable(1, 1, 'End', [['NEWCELL']]);
        table.load('values');
        await context.sync();
        return { returned: table.values.map((row) => row.map(vis)) };
      }
      case 'paragraph-delete': (await find(context, region.range.paragraph)).paragraphs.getFirst().delete(); await context.sync(); return {};
      case 'clear': (await cellOf(context, region.range.cell)).body.clear(); await context.sync(); return {};
      case 'ooxml-replace': return returned((await cellOf(context, region.range.cell)).body.insertOoxml(TABLE_OOXML, 'Replace'));
      case 'control-delete': context.document.body.contentControls.getByTag(region.range.control).getFirst().delete(false); await context.sync(); return {};
      default: break;
    }
    const range = await rangeOf(context, region);
    switch (region.op) {
      case 'delete': range.delete(); await context.sync(); return {};
      case 'replace': return returned(range.insertText('NEWTEXT', 'Replace'));
      case 'bold': range.font.bold = true; await context.sync(); return {};
      case 'paragraph-before': return returned(range.insertParagraph('NEWPARA', 'Before'));
      case 'paragraph-after': return returned(range.insertParagraph('NEWPARA', 'After'));
      case 'text-start': return returned(range.insertText('NEWSTART', 'Start'));
      case 'text-end': return returned(range.insertText('NEWEND', 'End'));
      case 'comment': return returned(range.insertComment('NEWCOMMENT').getRange());
      default: throw new Error('no operation ' + region.op);
    }
  }

  /** Each query's hits in the region: their text and the paragraphs each touches (CR-014 section 2, item 5). */
  async function searches(context, region) {
    const out = {};
    for (const q of region.queries) {
      out[q] = await attempt(async () => {
        const hits = (await regionOf(context, region.n)).search(q, { matchCase: true });
        hits.load('items/text');
        await context.sync();
        const paragraphs = hits.items.map((h) => { const ps = h.paragraphs; ps.load('items/text'); return ps; });
        await context.sync();
        return hits.items.map((h, i) => ({ text: vis(h.text), paragraphs: paragraphs[i].items.map((p) => vis(p.text)) }));
      });
    }
    return out;
  }

  /** The region as it stands: its text, paragraphs, tables, tracked changes, comments, controls and notes. */
  async function report(context, region) {
    const out = {};
    const whole = await regionOf(context, region.n);
    out.text = await attempt(async () => { whole.load('text'); await context.sync(); return vis(whole.text); });
    out.paragraphs = await attempt(async () => {
      const paragraphs = whole.paragraphs;
      paragraphs.load('items/text,items/style,items/styleBuiltIn,items/alignment,items/isListItem,items/tableNestingLevel,items/font/bold,items/font/italic,items/font/underline');
      await context.sync();
      const lists = paragraphs.items.map((p) => { const l = p.listItemOrNullObject; l.load('listString,level'); return l; });
      const words = region.op === 'bold'
        ? paragraphs.items.map((p) => { const w = p.getTextRanges([' '], false); w.load('items/text,items/font/bold'); return w; })
        : [];
      await context.sync();
      return paragraphs.items.map((p, i) => {
        const entry = {
          text: vis(p.text), style: p.style, styleBuiltIn: p.styleBuiltIn, alignment: p.alignment, tableNestingLevel: p.tableNestingLevel,
          bold: p.font.bold, italic: p.font.italic, underline: p.font.underline,
        };
        if (!lists[i].isNullObject) entry.list = { listString: vis(lists[i].listString), level: lists[i].level };
        if (words.length > 0) entry.words = words[i].items.map((w) => ({ text: vis(w.text), bold: w.font.bold }));
        return entry;
      });
    });
    out.tables = await attempt(async () => {
      const tables = whole.tables;
      tables.load('items/rowCount,items/values');
      await context.sync();
      const ranges = tables.items.map((x) => { const r = x.getRange('Whole'); r.load('text'); return r; });
      await context.sync();
      return tables.items.map((x, i) => ({ rowCount: x.rowCount, values: x.values.map((row) => row.map(vis)), whole: vis(ranges[i].text) }));
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
      controls.load('items/tag,items/text');
      await context.sync();
      return controls.items.map((c) => ({ tag: c.tag, text: vis(c.text) }));
    });
    out.footnotes = await attempt(async () => {
      const notes = whole.footnotes;
      notes.load('items/type');
      await context.sync();
      const bodies = notes.items.map((note) => { note.body.load('text'); return note.body; });
      await context.sync();
      return bodies.map((b) => vis(b.text));
    });
    return out;
  }

  /**
   * The check: `mode` 'run' or 'report'. `run(fn)` is Word.run; `progress(text)` is told each case. Each step is
   * a Word.run of its own, so that a case Office JS refuses fails alone.
   */
  async function runCheck(run, mode, progress = () => {}, host = null) {
    const result = { check: 52, mode, host, reads: {}, cases: [] };
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
        if (mode === 'run' && region.op !== 'read') {
          await run(async (context) => {
            context.document.changeTrackingMode = region.tracked ? 'TrackAll' : 'Off';
            await context.sync();
          });
          if (region.queries) {
            entry.searches = await run((context) => searches(context, region));
          } else {
            if (!result.reads[region.kind]) result.reads[region.kind] = await run((context) => reads(context, region)).catch(failed);
            entry.range = await run(async (context) => { const r = await rangeOf(context, region); r.load('text'); await context.sync(); return vis(r.text); }).catch(failed);
            entry.result = await run((context) => apply(context, region));
          }
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

  return { parseRegion, queryText, runCheck, vis, TABLE_OOXML };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const buttons = ['run', 'report'].map((id) => document.getElementById(id));
  const start = (mode) => {
    buttons.forEach((b) => { b.disabled = true; });
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK52.runCheck((f) => Word.run(f), mode, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.caseCount + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => buttons.forEach((b) => { b.disabled = false; }));
  };
  buttons[0].addEventListener('click', () => start('run'));
  buttons[1].addEventListener('click', () => start('report'));
}
