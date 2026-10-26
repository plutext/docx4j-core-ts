// Check 71 (CR-002 section 65; test/README.md check 71): what Office JS's measurement setters do with a value past the
// range Word stores its attribute in (docx4j CR-066 section 3.1: 16 bits, wrapped on reading), on
// 71-measurement-setters.docx (test/fixtures/check71/build.mjs). WordApi 1.3 (setLevelIndents, startNewList, columnWidth).
// Every case is a region, BEGIN<n> <setter> <value>, its target paragraph TARGET<n>, END<n>. "Report only" reports every
// region and changes nothing; "Run the check" sets the case's setter to its value, in a Word.run of its own so that a
// refusal fails alone, and reports what the setter read back and what the region's OOXML holds. Developer > Design Mode
// stays off. Nothing here depends on Script Lab but the last few lines, so a test can load CHECK71 with Word bound to
// another implementation.
const CHECK71 = (() => {
  const failed = (e) => 'error: ' + ((e && e.code) || 'no code') + (e && e.message ? ': ' + e.message : '');
  const TOKEN = { matchCase: true, matchWholeWord: true };

  /** A BEGIN paragraph's text as a case, or null. */
  function parseRegion(text) {
    const m = /^BEGIN(\d+) (\w+) (-?[\d.]+)$/.exec(text.trim());
    return m ? { n: Number(m[1]), setter: m[2], value: Number(m[3]) } : null;
  }

  async function attempt(f) {
    try {
      const v = await f();
      return v === undefined ? 'undefined' : v;
    } catch (e) {
      return failed(e);
    }
  }
  async function paragraphOf(context, token) {
    const hit = context.document.body.search(token, TOKEN).getFirstOrNullObject();
    hit.load('text');
    await context.sync();
    if (hit.isNullObject) throw new Error('not found: ' + token);
    return hit.paragraphs.getFirst();
  }
  /** The region from its BEGIN paragraph to its END paragraph, both whole. */
  const regionOf = async (context, n) => (await paragraphOf(context, 'BEGIN' + n)).getRange('Whole')
    .expandTo((await paragraphOf(context, 'END' + n)).getRange('Whole'));
  /** The table's only cell, for columnWidth. */
  async function cellOf(context, n) {
    const tables = (await regionOf(context, n)).tables;
    tables.load('items');
    await context.sync();
    return tables.items[0].getCell(0, 0);
  }

  /** The case's write, synced: refused, or done. */
  async function apply(context, c) {
    const p = await paragraphOf(context, 'TARGET' + c.n);
    switch (c.setter) {
      case 'leftIndent': p.leftIndent = c.value; break;
      case 'rightIndent': p.rightIndent = c.value; break;
      case 'firstLineIndent': p.firstLineIndent = c.value; break;
      case 'spaceBefore': p.spaceBefore = c.value; break;
      case 'spaceAfter': p.spaceAfter = c.value; break;
      case 'lineSpacing': p.lineSpacing = c.value; break;
      case 'columnWidth': (await cellOf(context, c.n)).columnWidth = c.value; break;
      case 'fontSize': p.font.size = c.value; break;
      case 'listIndent': {
        const made = p.startNewList();
        await context.sync();
        // Office JS returns the list; the Word shim a result holding it, the call being asynchronous there
        const list = typeof made.setLevelIndents === 'function' ? made : made.value;
        list.setLevelIndents(0, c.value, 0);
        break;
      }
      default: throw new Error('no setter ' + c.setter);
    }
    await context.sync();
    return 'set';
  }

  /** The elements a case can change, as the region's OOXML holds them. */
  const ELEMENTS = /<w:ind\b[^>]*\/>|<w:spacing\b[^>]*\/>|<w:tcW\b[^>]*\/>|<w:sz\b[^>]*\/>|<w:numPr>[\s\S]*?<\/w:numPr>/g;
  const tidy = (xml) => xml.replace(/ w:rsid\w*="[^"]*"| w14:(?:paraId|textId)="[^"]*"| xmlns:\w+="[^"]*"/g, '');
  /** A part's xmlData in a range's OOXML, or ''. */
  function partOf(xml, name) {
    const at = xml.indexOf('pkg:name="' + name + '"');
    if (at < 0) return '';
    const start = xml.indexOf('<pkg:xmlData>', at);
    const end = xml.indexOf('</pkg:xmlData>', start);
    return start < 0 || end < 0 ? '' : xml.slice(start, end);
  }
  /** What the region's OOXML holds: the document's elements a case changes, and the numbering's level 0 indents. */
  function elementsOf(xml) {
    const document = partOf(xml, '/word/document.xml') || xml;
    const body = /<w:body>([\s\S]*?)<\/w:body>/.exec(document);
    const content = tidy((body ? body[1] : document).replace(/<w:sectPr\b[\s\S]*<\/w:sectPr>/, ''));
    const out = { document: content.match(ELEMENTS) || [] };
    const numbering = partOf(xml, '/word/numbering.xml');
    if (numbering) {
      out.numbering = [...numbering.matchAll(/<w:lvl\b[^>]*w:ilvl="0"[^>]*>([\s\S]*?)<\/w:lvl>/g)]
        .map((m) => (m[1].match(/<w:ind\b[^>]*\/>/) || ['(no w:ind)'])[0]).map(tidy);
    }
    return out;
  }

  /** What the setter reads back, and the region's OOXML. */
  async function report(context, c) {
    const out = {};
    out.read = await attempt(async () => {
      const p = await paragraphOf(context, 'TARGET' + c.n);
      if (c.setter === 'columnWidth') {
        const cell = await cellOf(context, c.n);
        cell.load('columnWidth');
        await context.sync();
        return cell.columnWidth;
      }
      if (c.setter === 'fontSize') {
        p.font.load('size');
        await context.sync();
        return p.font.size;
      }
      p.load('leftIndent,rightIndent,firstLineIndent,spaceBefore,spaceAfter,lineSpacing');
      await context.sync();
      if (c.setter === 'listIndent') return { leftIndent: p.leftIndent, firstLineIndent: p.firstLineIndent };
      return p[c.setter];
    });
    out.ooxml = await attempt(async () => {
      const x = (await regionOf(context, c.n)).getOoxml();
      await context.sync();
      return elementsOf(x.value);
    });
    return out;
  }

  /**
   * The check: `mode` 'run' or 'report'. `run(fn)` is Word.run; `progress(text)` is told each case. Each step is a
   * Word.run of its own, so that a case Office JS refuses fails alone.
   */
  async function runCheck(run, mode, progress = () => {}, host = null) {
    const result = { check: 71, mode, host, title: null, cases: [] };
    const regions = await run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      result.title = paragraphs.items.length > 0 ? paragraphs.items[0].text : null;
      return paragraphs.items.map((p) => parseRegion(p.text)).filter((r) => r !== null);
    });
    result.caseCount = regions.length;
    for (const region of regions) {
      progress('case ' + region.n + ' of ' + regions.length + ': ' + region.setter + ' ' + region.value);
      const entry = { ...region };
      if (mode === 'run') {
        try {
          entry.result = await run((context) => apply(context, region));
        } catch (e) {
          entry.error = failed(e);
        }
      }
      entry.after = await run((context) => report(context, region)).catch(failed);
      result.cases.push(entry);
    }
    return result;
  }

  return { parseRegion, runCheck, elementsOf };
})();

if (typeof document !== 'undefined' && document.getElementById('run')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const buttons = ['report', 'run'].map((id) => document.getElementById(id));
  const start = (mode) => {
    buttons.forEach((b) => { b.disabled = true; });
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    CHECK71.runCheck((f) => Word.run(f), mode, (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.caseCount + ' cases. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => buttons.forEach((b) => { b.disabled = false; }));
  };
  buttons[0].addEventListener('click', () => start('report'));
  buttons[1].addEventListener('click', () => start('run'));
}
