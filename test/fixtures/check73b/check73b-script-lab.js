// Check 73b (CR-020 D4 and D3; test/README.md check 73b): check 73 again, for what its run left open. In Word 365
// (16.0.20326.20232) body.contentControls gave none of the probes' controls, so the controls are looked for several
// ways, each recorded; and Word bound bind-plain-two-parts.docx's B7 to part 2, where its save of 23 October had part 1,
// so the run is repeated in fresh sessions to see whether Word's choice, and its order of the custom XML parts, holds.
// This reads, and changes nothing: the document's custom XML parts in Word's order (WordApi 1.4), each control's w:tag
// and w:dataBinding in the body's OOXML (WordApi 1.1), each lookup's count and tags, and for each control the first
// lookup finds, its text, its xmlMapping (WordApiDesktop 1.3) and the customXmlPart and customXmlNode that names. Each
// read is a Word.run of its own, so a read Office JS refuses fails alone. Nothing here depends on Script Lab but the
// last few lines, so a test can load CHECK73B with Word bound to another implementation.
const CHECK73B = (() => {
  // every character outside printable ASCII as \uXXXX (a surrogate pair as two), so a control character shows
  const vis = (s) => (s === null || s === undefined ? s : String(s).replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')));
  const failed = (e) => 'error: ' + ((e && e.code) || 'no code') + (e && e.message ? ': ' + e.message : '');

  /** One read, synced alone, so that what Office JS refuses fails alone. */
  async function attempt(f) {
    try {
      return await f();
    } catch (e) {
      return failed(e);
    }
  }

  /** Every w:dataBinding and w15:dataBinding element in an OOXML string, as written. */
  const dataBindingsIn = (ooxml) => String(ooxml).match(/<w(?:15)?:dataBinding\b[^>]*>/g) || [];

  /** Each w:sdtPr in an OOXML string, in document order: its w:tag and its data bindings. */
  const sdtBindingsIn = (ooxml) => (String(ooxml).match(/<w:sdtPr\b[\s\S]*?<\/w:sdtPr>/g) || []).map((pr) => {
    const tag = /<w:tag\b[^>]*\bw:val="([^"]*)"/.exec(pr);
    return { tag: tag ? tag[1] : null, dataBinding: dataBindingsIn(pr) };
  });

  /** The ways the controls are looked for: each gives a collection, or an array of collections, to load. */
  const LOOKUPS = [
    ['document.contentControls', (context) => [context.document.contentControls]],
    ['body.contentControls', (context) => [context.document.body.contentControls]],
    ['body.getContentControls()', (context) => [context.document.body.getContentControls()]],
    ['paragraphs[i].contentControls', async (context) => {
      const ps = context.document.body.paragraphs;
      ps.load('items');
      await context.sync();
      return ps.items.map((p) => p.contentControls);
    }],
    ['paragraphs[i].getRange().contentControls', async (context) => {
      const ps = context.document.body.paragraphs;
      ps.load('items');
      await context.sync();
      return ps.items.map((p) => p.getRange('Whole').contentControls);
    }],
  ];

  /** The controls a lookup gives, as [collection, index] pairs with their tags. */
  async function lookup(context, at) {
    const collections = await LOOKUPS[at][1](context);
    for (const c of collections) c.load('items/tag');
    await context.sync();
    const found = [];
    collections.forEach((c, ci) => c.items.forEach((item, ii) => found.push({ ci, ii, tag: item.tag })));
    return found;
  }

  /** The control a lookup's [collection, index] names. */
  async function controlAt(context, at, ci, ii) {
    const collections = await LOOKUPS[at][1](context);
    const c = collections[ci];
    c.load('items');
    await context.sync();
    return c.items[ii];
  }

  async function control(context, at, ci, ii) {
    const c = await controlAt(context, at, ci, ii);
    c.load('id,tag,title,type,text');
    await context.sync();
    return { id: c.id, tag: c.tag, title: c.title, type: c.type, text: vis(c.text) };
  }

  async function mapping(context, at, ci, ii) {
    const m = (await controlAt(context, at, ci, ii)).xmlMapping;
    m.load('isMapped,xpath,prefixMappings');
    await context.sync();
    return { isMapped: m.isMapped, xpath: m.xpath, prefixMappings: m.prefixMappings };
  }

  async function part(context, at, ci, ii) {
    const p = (await controlAt(context, at, ci, ii)).xmlMapping.customXmlPart;
    p.load('id,namespaceUri');
    await context.sync();
    return p.isNullObject ? null : { id: p.id, namespaceUri: p.namespaceUri };
  }

  async function node(context, at, ci, ii) {
    const n = (await controlAt(context, at, ci, ii)).xmlMapping.customXmlNode;
    n.load('baseName,namespaceUri,text');
    await context.sync();
    return n.isNullObject ? null : { baseName: n.baseName, namespaceUri: n.namespaceUri, text: vis(n.text) };
  }

  async function bodyBindings(context) {
    const xml = context.document.body.getOoxml();
    await context.sync();
    return sdtBindingsIn(xml.value);
  }

  async function parts(context) {
    const ps = context.document.customXmlParts;
    ps.load('items/id,items/namespaceUri');
    await context.sync();
    return ps.items.map((p) => ({ id: p.id, namespaceUri: p.namespaceUri }));
  }

  /** The check: `run(fn)` is Word.run; `progress(text)` is told each step. */
  async function runCheck(run, progress = () => {}, host = null) {
    const result = { check: '73b', host };
    result.parts = await run(parts).catch(failed);
    result.bodyBindings = await run(bodyBindings).catch(failed);
    result.lookups = [];
    let first = null;
    for (let at = 0; at < LOOKUPS.length; at++) {
      progress('lookup ' + LOOKUPS[at][0]);
      const found = await attempt(() => run((context) => lookup(context, at)));
      result.lookups.push({ via: LOOKUPS[at][0], count: Array.isArray(found) ? found.length : null, tags: Array.isArray(found) ? found.map((f) => f.tag) : found });
      if (first === null && Array.isArray(found) && found.length > 0) first = { at, found };
    }
    result.controls = [];
    if (first !== null) {
      result.controlsVia = LOOKUPS[first.at][0];
      for (const { ci, ii } of first.found) {
        progress('control ' + (result.controls.length + 1) + ' of ' + first.found.length);
        const entry = await attempt(() => run((context) => control(context, first.at, ci, ii)));
        const out = typeof entry === 'string' ? { error: entry } : entry;
        out.mapping = await attempt(() => run((context) => mapping(context, first.at, ci, ii)));
        out.customXmlPart = await attempt(() => run((context) => part(context, first.at, ci, ii)));
        out.customXmlNode = await attempt(() => run((context) => node(context, first.at, ci, ii)));
        result.controls.push(out);
      }
    }
    result.controlCount = result.controls.length;
    return result;
  }

  return { runCheck, sdtBindingsIn, LOOKUPS, vis };
})();

if (typeof document !== 'undefined' && document.getElementById('read')) {
  const out = document.getElementById('out');
  const status = document.getElementById('status');
  const button = document.getElementById('read');
  button.addEventListener('click', () => {
    button.disabled = true;
    out.value = '';
    const host = typeof Office !== 'undefined' && Office.context
      ? { diagnostics: Office.context.diagnostics, url: Office.context.document && Office.context.document.url }
      : null;
    CHECK73B.runCheck((f) => Word.run(f), (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.controlCount + ' controls. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => { button.disabled = false; });
  });
}
