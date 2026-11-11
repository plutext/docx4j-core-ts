// core-ts paragraph-joining check J9b (body End  pkg3): Office JS insertOoxml of a hand-built Flat OPC package into
// ij-template.docx, read back through Office JS. Built by the portfolio session's build_join.py, 2026-11-11.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const CASE = "J9b";
const TARGET = {"kind": "body", "loc": "End"};
const PKG_NAME = "pkg3";
const OOXML = "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body><w:p><w:pPr><w:pStyle w:val=\"FragFirst\"/></w:pPr><w:r><w:t>First package paragraph.</w:t></w:r></w:p><w:p><w:r><w:t>Middle package paragraph.</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val=\"FragLast\"/><w:rPr><w:u w:val=\"single\"/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val=\"0000FF\"/></w:rPr><w:t>Last package paragraph.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/_rels/document.xml.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/styles.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"><pkg:xmlData><w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style><w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"FragFirst\"><w:name w:val=\"FragFirst\"/><w:basedOn w:val=\"Normal\"/><w:qFormat/><w:pPr><w:spacing w:before=\"480\"/><w:jc w:val=\"left\"/></w:pPr></w:style><w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"FragLast\"><w:name w:val=\"FragLast\"/><w:basedOn w:val=\"Normal\"/><w:qFormat/><w:pPr><w:ind w:left=\"1440\"/><w:jc w:val=\"both\"/></w:pPr></w:style></w:styles></pkg:xmlData></pkg:part></pkg:package>";
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const tidy = (s) => s.replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ (w14:paraId|w14:textId|w:rsid\w*|w16du:dateUtc)="[^"]*"/g, '');
const err = (e) => String((e && (e.code ? e.code + ': ' : '') + (e.message || e)) || e);

function bodyXml(flat) {
  const doc = new DOMParser().parseFromString(flat, 'application/xml');
  const main = Array.from(doc.getElementsByTagNameNS(PKG, 'part')).find((p) => p.getAttribute('pkg:name') === '/word/document.xml');
  const body = main && main.getElementsByTagNameNS(W, 'body')[0];
  const r = { body: body ? tidy(new XMLSerializer().serializeToString(body)) : null };
  const st = Array.from(doc.getElementsByTagNameNS(PKG, 'part')).find((p) => p.getAttribute('pkg:name') === '/word/styles.xml');
  if (st) {
    r.styles = Array.from(st.getElementsByTagNameNS(W, 'style'))
      .filter((s) => /Target|Frag/.test(s.getAttribute('w:styleId') || ''))
      .map((s) => tidy(new XMLSerializer().serializeToString(s)));
    r.styleIds = Array.from(st.getElementsByTagNameNS(W, 'style')).map((s) => s.getAttribute('w:styleId'));
  }
  return r;
}

async function snapshot(context) {
  const s = { errors: {} };
  const paras = context.document.body.paragraphs;
  paras.load('items/text,items/style,items/alignment,items/leftIndent,items/firstLineIndent,items/spaceBefore,items/spaceAfter,items/tableNestingLevel');
  const flat = context.document.body.getOoxml();
  await context.sync();
  s.count = paras.items.length;
  s.paragraphs = paras.items.map((p) => ({ text: p.text, style: p.style, alignment: p.alignment, leftIndent: p.leftIndent,
    firstLineIndent: p.firstLineIndent, spaceBefore: p.spaceBefore, spaceAfter: p.spaceAfter, inTable: p.tableNestingLevel }));
  try { s.ooxml = bodyXml(flat.value); } catch (e) { s.errors.ooxml = err(e); }
  return s;
}

async function targetParagraph(context) {
  const paras = context.document.body.paragraphs;
  paras.load('items/text');
  await context.sync();
  const hits = paras.items.filter((p) => p.text === TARGET.para);
  if (hits.length !== 1) throw new Error('paragraphs "' + TARGET.para + '": ' + hits.length);
  return hits[0];
}

async function insert(context) {
  let returned;
  if (TARGET.kind === 'body') {
    returned = context.document.body.insertOoxml(OOXML, TARGET.loc);
  } else {
    const p = await targetParagraph(context);
    if (TARGET.kind === 'range') {
      returned = p.getRange('Whole').insertOoxml(OOXML, TARGET.loc);
    } else if (TARGET.kind === 'paragraph') {
      returned = p.insertOoxml(OOXML, TARGET.loc);
    } else if (TARGET.kind === 'search') {
      const found = p.search(TARGET.find, { matchCase: true, matchWholeWord: true });
      found.load('items');
      await context.sync();
      if (found.items.length !== 1) throw new Error('search "' + TARGET.find + '": ' + found.items.length);
      returned = found.items[0].insertOoxml(OOXML, TARGET.loc);
    }
  }
  returned.load('text');
  await context.sync();
  return returned.text;
}

async function check() {
  const report = { case: CASE, target: TARGET, package: PKG_NAME, host: Office.context.diagnostics, before: null,
    inserted: false, returnedText: null, after: null, errors: {} };
  await Word.run(async (context) => {
    try { report.before = await snapshot(context); } catch (e) { report.errors.before = err(e); }
    try { report.returnedText = await insert(context); report.inserted = true; } catch (e) { report.errors.insert = err(e); }
    try { report.after = await snapshot(context); } catch (e) { report.errors.after = err(e); }
  });
  out.value = JSON.stringify(report, null, 1);
}
