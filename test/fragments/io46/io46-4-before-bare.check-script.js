// core-ts insertOoxml check, case 4 (as case 1, the package without its styles and numbering parts): what Word 365 holds after Office JS's insertOoxml of the
// ED-005 proposal 46 fragment (Flat OPC), read back through Office JS. Built by the portfolio session's build_checks.py.
const out = document.getElementById('out');
document.getElementById('run').addEventListener('click', () => check().catch((e) => { out.value = String(e.stack || e); }));

const CASE = 4;
const FLAT = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><?mso-application progid=\"Word.Document\"?><pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body><w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t xml:space=\"preserve\">Fragment heading (the fragment's Heading 1: red, 20 pt)</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val=\"FragmentNote\"/></w:pPr><w:r><w:t xml:space=\"preserve\">A paragraph in FragmentNote, a style the document lacks (green, italic).</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t xml:space=\"preserve\">Fragment list, item one (the fragment's numId 1: upper-case Roman)</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t xml:space=\"preserve\">Fragment list, item two</w:t></w:r></w:p><w:p><w:r><w:t xml:space=\"preserve\">A footnote follows this sentence.</w:t></w:r><w:r><w:rPr><w:rStyle w:val=\"FootnoteReference\"/></w:rPr><w:footnoteReference w:id=\"1\"/></w:r></w:p><w:p><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"><wp:extent cx=\"1524000\" cy=\"762000\"/><wp:docPr id=\"1\" name=\"Fragment picture\"/><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"fragment.png\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rIdImage\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1524000\" cy=\"762000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t xml:space=\"preserve\">Service</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t xml:space=\"preserve\">Days</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t xml:space=\"preserve\">Standard</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t xml:space=\"preserve\">14</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sdt><w:sdtPr><w:alias w:val=\"Inner control\"/><w:tag w:val=\"inner\"/><w:id w:val=\"201\"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t xml:space=\"preserve\">A control inside the fragment (tag &quot;inner&quot;).</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t xml:space=\"preserve\">Last paragraph of the fragment. Its section is landscape A5; the document's is portrait A4.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w=\"11906\" w:h=\"8391\" w:orient=\"landscape\"/><w:pgMar w:top=\"720\" w:right=\"720\" w:bottom=\"720\" w:left=\"720\" w:header=\"360\" w:footer=\"360\" w:gutter=\"0\"/></w:sectPr></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/_rels/document.xml.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdFootnotes\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/><Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/footnotes.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"><pkg:xmlData><w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id=\"1\"><w:p><w:pPr><w:pStyle w:val=\"FootnoteText\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"FootnoteReference\"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space=\"preserve\"> The fragment's footnote.</w:t></w:r></w:p></w:footnote></w:footnotes></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/media/image1.png\" pkg:contentType=\"image/png\" pkg:compression=\"store\"><pkg:binaryData>iVBORw0KGgoAAAANSUhEUgAAAKAAAABQCAIAAAARP+ljAAAA1ElEQVR4nO3RAQkAIBDAwE9kO6uZ\nzxQijIMLMNisfQib7wU8ZXCcwXEGxxkcZ3CcwXEGxxkcZ3CcwXEGxxkcZ3CcwXEGxxkcZ3CcwXEG\nxxkcZ3CcwXEGxxkcZ3CcwXEGxxkcZ3CcwXEGxxkcZ3CcwXEGxxkcZ3CcwXEGxxkcZ3CcwXEGxxkc\nZ3CcwXEGxxkcZ3CcwXEGxxkcZ3CcwXEGxxkcZ3CcwXEGxxkcZ3CcwXEGxxkcZ3CcwXEGxxkcZ3Cc\nwXEGxxkcZ3CcwXEGxxkcZ3CcwXEXIcRl7l2Go4oAAAAASUVORK5CYII=</pkg:binaryData></pkg:part></pkg:package>";
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const tidy = (s) => s.replace(/ (w14:paraId|w14:textId|w:rsid\w*|w16du:dateUtc)="[^"]*"/g, '');
const ser = (node) => tidy(new XMLSerializer().serializeToString(node));
const err = (e) => String((e && (e.code ? e.code + ': ' : '') + (e.message || e)) || e);

function partsOf(flat) {
  const doc = new DOMParser().parseFromString(flat, 'application/xml');
  const parts = {};
  for (const p of Array.from(doc.getElementsByTagNameNS(PKG, 'part'))) parts[p.getAttribute('pkg:name')] = p;
  return parts;
}

function readOoxml(flat) {
  const parts = partsOf(flat);
  const r = { parts: Object.keys(parts).sort() };
  const main = parts['/word/document.xml'];
  if (main) {
    const body = main.getElementsByTagNameNS(W, 'body')[0];
    r.body = body ? ser(body) : null;
  }
  const st = parts['/word/styles.xml'];
  if (st) {
    r.styles = Array.from(st.getElementsByTagNameNS(W, 'style')).filter((s) => {
      const id = s.getAttributeNS(W, 'styleId') || s.getAttribute('w:styleId') || '';
      return /Heading1|FragmentNote|Footnote/i.test(id) || s.getAttribute('w:customStyle') === '1';
    }).map(ser);
    r.styleIds = Array.from(st.getElementsByTagNameNS(W, 'style')).map((s) => s.getAttribute('w:styleId'));
  }
  if (main) {
    r.listParagraphs = Array.from(main.getElementsByTagNameNS(W, 'p')).map((p) => {
      const pr = p.getElementsByTagNameNS(W, 'numPr')[0];
      if (!pr) return null;
      const numId = pr.getElementsByTagNameNS(W, 'numId')[0];
      const ilvl = pr.getElementsByTagNameNS(W, 'ilvl')[0];
      const text = Array.from(p.getElementsByTagNameNS(W, 't')).map((t) => t.textContent).join('');
      return { text, numId: numId ? numId.getAttribute('w:val') : null, ilvl: ilvl ? ilvl.getAttribute('w:val') : null };
    }).filter((x) => x);
  }
  const num = parts['/word/numbering.xml'];
  if (num) {
    const numbering = num.getElementsByTagNameNS(W, 'numbering')[0];
    r.numbering = ser(numbering);
    const abstracts = {};
    for (const a of Array.from(numbering.getElementsByTagNameNS(W, 'abstractNum'))) {
      const lvl0 = Array.from(a.getElementsByTagNameNS(W, 'lvl')).find((l) => l.getAttribute('w:ilvl') === '0');
      const fmt = lvl0 && lvl0.getElementsByTagNameNS(W, 'numFmt')[0];
      const start = lvl0 && lvl0.getElementsByTagNameNS(W, 'start')[0];
      abstracts[a.getAttribute('w:abstractNumId')] = { numFmt: fmt ? fmt.getAttribute('w:val') : null, start: start ? start.getAttribute('w:val') : null };
    }
    r.nums = Array.from(numbering.getElementsByTagNameNS(W, 'num')).map((n) => {
      const an = n.getElementsByTagNameNS(W, 'abstractNumId')[0];
      const id = an ? an.getAttribute('w:val') : null;
      const overrides = n.getElementsByTagNameNS(W, 'lvlOverride').length;
      return { numId: n.getAttribute('w:numId'), abstractNumId: id, lvl0: abstracts[id] || null, lvlOverrides: overrides };
    });
  }
  const fn = parts['/word/footnotes.xml'];
  if (fn) r.footnotes = ser(fn.getElementsByTagNameNS(W, 'footnotes')[0]);
  for (const name of Object.keys(parts)) {
    const bin = parts[name].getElementsByTagNameNS(PKG, 'binaryData')[0];
    if (bin) (r.binaries = r.binaries || {})[name] = bin.textContent.length;
  }
  return r;
}

async function snapshot(context, full) {
  const s = { errors: {} };
  const paras = context.document.body.paragraphs;
  paras.load('items/text,items/style,items/isListItem,items/tableNestingLevel');
  const ccs = context.document.contentControls;
  ccs.load('items/id,items/tag,items/title,items/type');
  const secs = context.document.sections;
  secs.load('items');
  const flat = full ? context.document.body.getOoxml() : null;
  await context.sync();
  const lis = paras.items.map((p) => p.listItemOrNullObject);
  lis.forEach((li) => li.load('listString,level'));
  await context.sync();
  s.paragraphs = paras.items.map((p, i) => ({ text: p.text, style: p.style, list: lis[i].isNullObject ? null : { string: lis[i].listString, level: lis[i].level }, inTable: p.tableNestingLevel }));
  s.contentControls = ccs.items.map((c) => ({ id: c.id, tag: c.tag, title: c.title, type: c.type }));
  s.sections = secs.items.length;
  if (!full) return s;
  try { s.ooxml = readOoxml(flat.value); } catch (e) { s.errors.ooxml = err(e); }
  try {
    const styles = context.document.getStyles();
    styles.load('items/nameLocal,items/builtIn,items/inUse,items/type');
    await context.sync();
    s.stylesInUse = styles.items.filter((x) => x.inUse || !x.builtIn).map((x) => ({ name: x.nameLocal, builtIn: x.builtIn, inUse: x.inUse, type: x.type }));
  } catch (e) { s.errors.getStyles = err(e); }
  try {
    const notes = context.document.body.footnotes;
    notes.load('items');
    await context.sync();
    notes.items.forEach((n) => n.body.load('text'));
    await context.sync();
    s.footnotes = notes.items.map((n) => n.body.text);
  } catch (e) { s.errors.footnotes = err(e); }
  return s;
}

async function insert(context) {
  if (CASE === 3) {
    const ccs = context.document.contentControls.getByTitle('Word fragment');
    ccs.load('items');
    await context.sync();
    if (ccs.items.length !== 1) throw new Error('controls titled Word fragment: ' + ccs.items.length);
    ccs.items[0].insertOoxml(FLAT, 'Replace');
    await context.sync();
    return;
  }
  const times = CASE === 2 ? 2 : 1;
  for (let t = 0; t < times; t++) {
    const paras = context.document.body.paragraphs;
    paras.load('items/text');
    await context.sync();
    const at = paras.items.filter((p) => p.text.trim() === 'After the control.');
    if (at.length !== 1) throw new Error('paragraphs "After the control.": ' + at.length);
    at[0].getRange('Whole').insertOoxml(FLAT, 'Before');
    await context.sync();
  }
}

async function check() {
  const report = { case: CASE, host: Office.context.diagnostics, before: null, inserted: false, after: null, errors: {} };
  await Word.run(async (context) => {
    try { report.before = await snapshot(context, false); } catch (e) { report.errors.before = err(e); }
    try { await insert(context); report.inserted = true; } catch (e) { report.errors.insert = err(e); }
    try { report.after = await snapshot(context, true); } catch (e) { report.errors.after = err(e); }
  });
  out.value = JSON.stringify(report, null, 1);
}
