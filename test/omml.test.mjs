// docx4j-core-tests org.docx4j.jaxb.OmmlInDrawingMLTextTest (docx4j CR-025): an equation in a
// DrawingML text body (a14:m holding m:oMathPara) has an a:rPr on every m:r and m:ctrlPr, which
// the OMML schema admits since CR-025 (objects 0.1.6; CR-001 section 17.5). docx4j types the
// equation only in its WordprocessingML context and keeps it as DOM in the PresentationML and
// SpreadsheetML ones; this package has one context, so it is typed in every part. Both fixture
// parts hold the same equation: 16 runs and 7 control-properties elements, each naming Cambria Math.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OpcPackage, ZipPartStore } from '../dist/index.mjs';
import { fixture } from './helpers.mjs';

const DML = 'http://schemas.openxmlformats.org/drawingml/2006/main';

function collect(root) {
  const runs = [];
  const ctrls = [];
  const walk = (v) => {
    if (!v || typeof v !== 'object' || v.nodeType) return;
    if (v.TYPE_NAME === 'org_docx4j_math.CTR') runs.push(v);
    if (v.TYPE_NAME === 'org_docx4j_math.CTCtrlPr') ctrls.push(v);
    for (const k of Object.keys(v)) if (k !== 'PARENT') walk(v[k]);
  };
  walk(root);
  return { runs, ctrls };
}

const count = (xml, s) => xml.split(s).length - 1;

for (const [name, partName] of [['loadAndSave.pptx', '/ppt/slides/slide2.xml'], ['loadAndSave.xlsx', '/xl/drawings/drawing1.xml']]) {
  test(`OMML inside DrawingML text is typed with its a:rPr and re-marshalled with it: ${name}`, async () => {
    const bytes = await fixture(name);
    const source = new TextDecoder().decode(new ZipPartStore(bytes).loadSync(partName.slice(1)));
    assert.equal(count(source, '<m:r><a:rPr'), 16);
    assert.equal(count(source, '<m:ctrlPr><a:rPr'), 7);

    const pkg = await OpcPackage.load(bytes);
    const { runs, ctrls } = collect(await pkg.parts.get(partName).getContents());
    assert.equal(runs.length, 16);
    assert.equal(ctrls.length, 7);
    for (const r of runs) {
      const first = r.content[0];
      assert.equal(first.name.namespaceURI, DML, "a run's first child is the a:rPr element");
      assert.equal(first.name.localPart, 'rPr');
      assert.equal(first.value.TYPE_NAME, 'org_docx4j_dml.CTTextCharacterProperties');
      assert.equal(first.value.latin.typeface, 'Cambria Math');
    }
    for (const c of ctrls) {
      assert.equal(c.rPrDml?.TYPE_NAME, 'org_docx4j_dml.CTTextCharacterProperties', 'm:ctrlPr keeps its a:rPr');
      assert.equal(c.rPrDml.latin.typeface, 'Cambria Math');
    }

    const out = new TextDecoder().decode(new ZipPartStore(await pkg.save()).loadSync(partName.slice(1)));
    assert.equal(count(out, '<m:r><a:rPr'), 16);
    assert.equal(count(out, '<m:ctrlPr><a:rPr'), 7);
    assert.equal(count(out, '<a14:m>'), 1);
  });
}
