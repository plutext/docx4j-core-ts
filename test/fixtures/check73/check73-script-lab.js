// Check 73 (CR-020 D3; test/README.md check 73): what Office JS reports of a bound control whose w:storeItemID names
// no part, on docx4j CR-058's three probes (bind-plain.docx, bind-plain-two-parts.docx, bind-plain-no-part.docx).
// Word repairs such a binding on open: it binds the control to the custom XML part its XPath fits, and its save writes
// that part's ID (docx4j CR-058 section 8). This reads, and changes nothing: for each control in the body, its tag and
// text, its xmlMapping (WordApiDesktop 1.3: isMapped, xpath, prefixMappings, and the customXmlPart and customXmlNode it
// names); every control's w:tag and w:dataBinding in the body's OOXML (WordApi 1.1's getOoxml), which is the ID Word
// holds after the open; and the document's custom XML parts (WordApi 1.4). Each read is a Word.run of its own, so a read Office JS refuses
// fails alone. Nothing here depends on Script Lab but the last few lines, so a test can load CHECK73 with Word bound to
// another implementation.
const CHECK73 = (() => {
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

  /** The body's controls by index, with what identifies each. */
  async function controls(context) {
    const ccs = context.document.body.contentControls;
    ccs.load('items/id,items/tag,items/title,items/type,items/text');
    await context.sync();
    return ccs.items.map((c, index) => ({ index, id: c.id, tag: c.tag, title: c.title, type: c.type, text: vis(c.text) }));
  }

  /** The control at an index of the body's collection, loaded for the next read. */
  async function nth(context, index) {
    const ccs = context.document.body.contentControls;
    ccs.load('items');
    await context.sync();
    return ccs.items[index];
  }

  async function mapping(context, index) {
    const m = (await nth(context, index)).xmlMapping;
    m.load('isMapped,xpath,prefixMappings');
    await context.sync();
    return { isMapped: m.isMapped, xpath: m.xpath, prefixMappings: m.prefixMappings };
  }

  async function part(context, index) {
    const p = (await nth(context, index)).xmlMapping.customXmlPart;
    p.load('id,namespaceUri');
    await context.sync();
    return p.isNullObject ? null : { id: p.id, namespaceUri: p.namespaceUri };
  }

  async function node(context, index) {
    const n = (await nth(context, index)).xmlMapping.customXmlNode;
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

  /** The check: `run(fn)` is Word.run; `progress(text)` is told each control. */
  async function runCheck(run, progress = () => {}, host = null) {
    const result = { check: 73, host };
    result.parts = await run(parts).catch(failed);
    result.bodyBindings = await run(bodyBindings).catch(failed);
    const found = await run(controls);
    result.controlCount = found.length;
    result.controls = [];
    for (const c of found) {
      progress('control ' + (c.index + 1) + ' of ' + found.length + ': ' + c.tag);
      const entry = { ...c };
      entry.mapping = await attempt(() => run((context) => mapping(context, c.index)));
      entry.customXmlPart = await attempt(() => run((context) => part(context, c.index)));
      entry.customXmlNode = await attempt(() => run((context) => node(context, c.index)));
      result.controls.push(entry);
    }
    return result;
  }

  return { runCheck, sdtBindingsIn, vis };
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
    CHECK73.runCheck((f) => Word.run(f), (text) => { status.textContent = text; }, host)
      .then((result) => { out.value = JSON.stringify(result, null, 1); status.textContent = 'Done: ' + result.controlCount + ' controls. Copy the box below back.'; })
      .catch((e) => { out.value = String((e && e.stack) || e); status.textContent = 'Stopped: see the box below.'; })
      .finally(() => { button.disabled = false; });
  });
}
