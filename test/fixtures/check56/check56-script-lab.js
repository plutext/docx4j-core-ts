// Check 56 (CR-002 section 61; test/README.md check 56): what Office JS's block writes do to a bound block content
// control, by kind, on 56-writes.docx (test/fixtures/check56/build.mjs). WordApi 1.5 for getContentControls, 1.3 for
// the rest. Every case is a region, BEGIN<n> <op> <kind> ... END<n>, around one control tagged c56-<n> holding one
// paragraph start<n>, bound to value<n>. The control is looked for several ways, each recorded; "Run the check" gives
// it one write through the first way that finds it (a paragraph, a table, text with a line feed, or two paragraphs of
// OOXML, each at its end) and reports the region, the control and its node. "Report only" reports and changes nothing.
// Nothing here depends on Script Lab but the last few lines, so a test can load CHECK56 with Word bound to another
// implementation.
const CHECK56 = (() => {
  // every character outside printable ASCII as \uXXXX (a surrogate pair as two), so a control character shows
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && e.code) || 'no code') + (e && e.message ? ': ' + e.message : '');
  const TOKEN = { matchCase: true, matchWholeWord: true };
  const DATA_NS = 'urn:check56';
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const tagOf = (n) => 'c56-' + n;

  /** Two paragraphs, OOXML<n>A and OOXML<n>B, as a flat OPC package. */
  const twoParagraphs = (n) => '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
    + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
    + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    + '</pkg:xmlData></pkg:part>'
    + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>'
    + '<w:document xmlns:w="' + W + '"><w:body><w:p><w:r><w:t>OOXML' + n + 'A</w:t></w:r></w:p><w:p><w:r><w:t>OOXML' + n + 'B</w:t></w:r></w:p></w:body></w:document>'
    + '</pkg:xmlData></pkg:part></pkg:package>';

  /** A BEGIN paragraph's text as a case, or null. */
  function parseRegion(text) {
    const m = /^BEGIN(\d+) ([\w-]+) ([\w-]+)$/.exec(text.trim());
    return m ? { n: Number(m[1]), op: m[2], kind: m[3] } : null;
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

  /** The ways a case's control is looked for: collections to search by tag; the write goes through the first that finds it. */
  const LOOKUPS = [
    ['body.getContentControls()', (context) => context.document.body.getContentControls()],
    ['body.getContentControls(PlainText)', (context) => context.document.body.getContentControls({ types: ['PlainText'] })],
    ['body.getContentControls(RichText)', (context) => context.document.body.getContentControls({ types: ['RichText'] })],
    ['body.contentControls', (context) => context.document.body.contentControls],
  ];

  /** The case's control through one way, or null; and how many the way lists. */
  async function lookup(context, n, way) {
    const controls = way(context);
    controls.load('items/tag,items/type');
    await context.sync();
    const mine = controls.items.filter((c) => c.tag === tagOf(n));
    return { control: mine[0] || null, listed: controls.items.length, type: mine[0] ? mine[0].type : null };
  }

  /** Every way, recorded; and the control's paragraph's parentContentControlOrNullObject. */
  async function lookups(context, region) {
    const out = {};
    for (const [name, way] of LOOKUPS) {
      out[name] = await attempt(async () => {
        const r = await lookup(context, region.n, way);
        return { listed: r.listed, found: r.control !== null, type: r.type };
      });
    }
    out['paragraph.parentContentControlOrNullObject'] = await attempt(async () => {
      const parent = (await paragraphOf(context, 'start' + region.n)).parentContentControlOrNullObject;
      parent.load('tag,type');
      await context.sync();
      return parent.isNullObject ? { found: false } : { found: parent.tag === tagOf(region.n), type: parent.type };
    });
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

  /** The case's write; what it returns. */
  async function apply(context, region) {
    const found = await controlOf(context, region.n);
    if (!found) throw new Error('no way found ' + tagOf(region.n));
    const c = found.control;
    const returned = async (r) => { r.load('text'); await context.sync(); return { via: found.via, returned: vis(r.text) }; };
    switch (region.op) {
      case 'paragraph-end': return returned(c.insertParagraph('NEWPARA' + region.n, 'End'));
      case 'table-end': {
        c.insertTable(1, 2, 'End', [['CELL' + region.n + 'A', 'CELL' + region.n + 'B']]);
        await context.sync();
        return { via: found.via };
      }
      case 'text-newline': return returned(c.insertText('x\ny', 'End'));
      case 'ooxml-end': return returned(c.insertOoxml(twoParagraphs(region.n), 'End'));
      default: throw new Error('no operation ' + region.op);
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

  /** The region as it stands: its text, paragraphs and OOXML; the control's type, text, paragraphs and tables; its node. */
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
    out.control = await attempt(async () => {
      const found = await controlOf(context, region.n);
      if (!found) return null;
      const c = found.control;
      c.load('text,type');
      const paragraphs = c.paragraphs;
      paragraphs.load('items/text');
      const tables = c.tables;
      tables.load('items');
      await context.sync();
      return { via: found.via, type: c.type, text: vis(c.text), paragraphs: paragraphs.items.map((p) => vis(p.text)), tables: tables.items.length };
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
    const result = { check: 56, mode, host, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    result.caseCount = regions.length;
    for (const region of regions) {
      progress('case ' + region.n + ' of ' + regions.length + ': ' + region.op + ' ' + region.kind);
      const entry = { n: region.n, op: region.op, kind: region.kind };
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
    CHECK56.runCheck((f) => Word.run(f), mode, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.caseCount + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => buttons.forEach((b) => { b.disabled = false; }));
  };
  buttons[0].addEventListener('click', () => start('run'));
  buttons[1].addEventListener('click', () => start('report'));
}
