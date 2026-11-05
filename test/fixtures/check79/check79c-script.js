// Word check 79c, the follow-up to 79a: is ParagraphCollection.getByParaId a route from a w14:paraId to a range, and which
// content makes Word renumber a file's ids at open? 79a's ids all changed there and 79b's didn't. Run on Word's own save of 79a
// and on 79d's files, one kind of content each. The function's source is reported so its signature isn't guessed at. Every
// form of an id is tried on every kind of paragraph, then a paragraph it reached is written to. #out is set once, last.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => {
  out.value = JSON.stringify({ check: '79c', failed: String((e && e.stack) || e) });
}));

const errorOf = (e) => ({ message: String((e && e.message) || e), code: e && e.code });
const bodyOf = (flat) => { const m = /<w:body[\s\S]*<\/w:body>/.exec(flat); return m ? m[0].replace(/ w:rsid\w*="[^"]*"/g, '') : null; };
const idsOf = (flat) => Array.from(flat.matchAll(/<w:p\b[^>]*?\bw14:paraId="([0-9A-Fa-f]+)"/g), (m) => m[1]);
const forms = (id) => [['as-written', id], ['lower-case', id.toLowerCase()], ['number', parseInt(id, 16)], ['decimal-string', String(parseInt(id, 16))]];

/** The members' own source, cut short: their parameters and the call each makes. */
function sources() {
  const of = (c, name) => { try { return c && c.prototype && c.prototype[name] ? String(c.prototype[name]).slice(0, 700) : null; } catch (e) { return String(e); } };
  return { getByParaId: of(Word.ParagraphCollection, 'getByParaId'), getById: of(Word.ParagraphCollection, 'getById'), getByIdOrNullObject: of(Word.ParagraphCollection, 'getByIdOrNullObject') };
}

async function tryReach(context, holderName, holder, value) {
  try {
    const p = holder.getByParaId(value);
    p.load('text');
    await context.sync();
    return { holder: holderName, text: p.text };
  } catch (e) { return { holder: holderName, error: errorOf(e) }; }
}

async function check() {
  const report = { check: '79c', ranAt: new Date().toISOString(), host: Office.context.diagnostics, sources: sources(), reached: [] };
  await Word.run(async (context) => {
    const flat = context.document.body.getOoxml();
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text');
    await context.sync();
    report.body = bodyOf(flat.value);
    report.ids = idsOf(flat.value);
    report.texts = paragraphs.items.map((p) => p.text);
    const holders = [['body.paragraphs', () => context.document.body.paragraphs], ['document.paragraphs', () => context.document.paragraphs]];
    // the last id is the paragraph getOoxml adds, so it's left for the unknown id's place
    for (const id of report.ids.slice(0, -1)) {
      const row = { id, forms: {} };
      for (const [form, value] of forms(id)) {
        row.forms[form] = [];
        for (const [name, holder] of holders) {
          let h;
          try { h = holder(); } catch (e) { row.forms[form].push({ holder: name, error: errorOf(e) }); continue; }
          row.forms[form].push(await tryReach(context, name, h, value));
        }
      }
      report.reached.push(row);
    }
    report.unknownId = await tryReach(context, 'body.paragraphs', context.document.body.paragraphs, '7FFFFFFE');
    // the collection's other two, given each kind of id
    report.byId = [];
    try {
      const local = context.document.body.paragraphs;
      local.load('items/uniqueLocalId');
      await context.sync();
      const given = [['paraId', report.ids[1]], ['uniqueLocalId', local.items[1].uniqueLocalId]];
      for (const [kind, value] of given) {
        for (const method of ['getById', 'getByIdOrNullObject']) {
          const row = { method, kind, value };
          try {
            const p = context.document.body.paragraphs[method](value);
            p.load('text,isNullObject');
            await context.sync();
            row.isNullObject = p.isNullObject;
            if (!p.isNullObject) row.text = p.text;
          } catch (e) { row.error = errorOf(e); }
          report.byId.push(row);
        }
      }
    } catch (e) { report.byIdError = errorOf(e); }
    // a range, and a write through it: the first id whose as-written form read back a paragraph
    const usable = report.reached.find((r) => r.forms['as-written'].some((x) => x.text !== undefined));
    report.write = { tried: usable ? usable.id : null };
    if (usable) {
      try {
        const p = context.document.body.paragraphs.getByParaId(usable.id);
        const range = p.getRange('Content');
        range.load('text');
        await context.sync();
        report.write.rangeText = range.text;
        range.insertText(' REACHED', 'End');
        await context.sync();
        const after = context.document.body.getOoxml();
        await context.sync();
        report.write.written = after.value.includes('REACHED');
        report.write.idAfter = idsOf(after.value).includes(usable.id);
      } catch (e) { report.write.error = errorOf(e); }
    }
  });
  out.value = JSON.stringify(report);
}
