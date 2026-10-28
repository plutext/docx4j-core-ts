// CR-017 (and CR-002 section 65.3): integer attribute values Word 365 reads differently, reported on load and on save,
// held to Word's own saves of docx4j CR-065's and CR-066's probes (test/integer-attributes/, copied unedited from the
// share with a SHA-1 manifest; Jason's runs of 2026-10-26). Word's save is the oracle: it writes Word's reading of each
// value, or leaves out a value that reads as the default. So:
// - a value past what Word stores (rounds 5 and 6) is a `value.out_of_range` warning on the part, and on `pkg.lastSave`
//   after a save, exactly where Word's save changed the value, its rule or its attribute, with `wordReads` what Word's
//   save wrote, and the value itself written as held;
// - a decimal or an odd form (round 2) reads as Word's save has it, `value.truncated`, `value.kept.invalid` or, for a
//   form Word refuses a document for, `value.refused`;
// - the setters' table (`WORD_STORAGE`, section 65.3) reads the objects package's table, and `keptRange` agrees with
//   its `wordIntegerReading` at every boundary.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, strFromU8 } from 'fflate';
import { WORD_INTEGER_STORAGE, wordIntegerReading } from '@docx4j/generated-objects-ts/helpers/wml';
import { WordprocessingMLPackage } from '../dist/index.mjs';
import { parseXml } from '../dist/xml/dom.mjs';
import { keptRange, rangeOf, UNMEASURED_RANGE } from '../dist/model/wordIntegers.mjs';
import { WORD_STORAGE } from '../dist/model/content/wordRange.mjs';

const DIR = new URL('./integer-attributes/', import.meta.url);
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const S = 'org_docx4j_wml.';
const bytesOf = (name) => readFileSync(new URL(name, DIR));
const load = async (name) => WordprocessingMLPackage.load(new Uint8Array(bytesOf(name)));
const partXmlOf = (name, part) => strFromU8(unzipSync(new Uint8Array(bytesOf(name)))[part]);
const OUT_OF_RANGE = (warnings) => warnings.filter((w) => w.code === 'value.out_of_range');

test('the probes and Word\'s saves are the share\'s, unedited (manifest.json)', () => {
  const manifest = JSON.parse(readFileSync(new URL('manifest.json', DIR), 'utf8'));
  assert.equal(manifest.source.docx4j.commit, '109a8d47b5ad916b75375643fed8911be917855d');
  const root = fileURLToPath(DIR);
  const files = [];
  const walk = (d) => { for (const n of readdirSync(d).sort()) { const p = join(d, n); if (statSync(p).isDirectory()) walk(p); else if (n !== 'manifest.json' && n !== 'README.md') files.push(relative(root, p)); } };
  walk(root);
  assert.deepEqual(files, manifest.files.map((f) => f.name), 'every file listed, nothing unlisted');
  for (const f of manifest.files) assert.equal(createHash('sha1').update(bytesOf(f.name)).digest('hex'), f.sha1, f.name);
  assert.equal(files.length, 114);
});

// ---- the range Word keeps as written, against the objects package's own reading ----

const RULES = [undefined, 'auto', 'exact', 'atLeast'];
const SAMPLES = [-70000, -65537, -65536, -65535, -32769, -32768, -32767, -31681, -31680, -256, -255, -32, -31, -2, -1, 0, 1, 31, 32, 143, 144, 255, 256,
  31680, 31681, 31682, 32766, 32767, 32768, 32769, 65534, 65535, 65536, 65537, 70000, 1205003, 7205001, 2147483647, 2147483648, -2147483649];

test('keptRange: every value in the range is kept by wordIntegerReading and none outside, for every entry of the objects table and every rule (CR-017 E1; CR-002 section 65.3)', () => {
  let entries = 0;
  for (const [typeName, properties] of Object.entries(WORD_INTEGER_STORAGE)) {
    for (const [property, entry] of Object.entries(properties)) {
      entries++;
      for (const rule of RULES) {
        const range = keptRange(typeName, property, rule);
        assert.deepEqual(range, rangeOf(entry, property, rule));
        if (range === null) {
          // a row height under hRule="auto": Word ignores it, so every value is "kept"
          for (const v of SAMPLES) assert.equal(wordIntegerReading(typeName, property, v, rule).kept, true, `${typeName}.${property}[${rule}] ${v}`);
          continue;
        }
        const [min, max] = range;
        const values = new Set([...SAMPLES, min - 1, min, min + 1, max - 1, max, max + 1]);
        for (const v of values) {
          const reading = wordIntegerReading(typeName, property, v, rule);
          const inRange = v >= min && v <= max;
          // Word's own spellings, kept though outside the range: -1 for 65535 where its save writes it so, and -1 a column space's sentinel
          if (v === -1 && (entry.maxWrittenAs !== undefined || entry.negative === -1)) { assert.equal(reading.kept, true, `${typeName}.${property}: -1 is Word's own spelling`); continue; }
          assert.equal(reading.kept, inRange, `${typeName}.${property}[${rule}] ${v}: kept ${reading.kept}, range ${min}..${max}`);
        }
        if (!entry.measured) assert.deepEqual(range, UNMEASURED_RANGE);
      }
    }
  }
  assert.equal(entries, 49, "the objects table's entries (its CR-016; w:legacy left out since objects 0.7.0, 5fa30a6)");
  assert.equal(keptRange('org_docx4j_wml.Nope', 'x'), undefined);
  assert.ok(Object.values(WORD_INTEGER_STORAGE).every((properties) => !('legacy' in properties) && !('legacyIndent' in properties) && !('legacySpace' in properties)), 'w:legacy is not in the table');
});

test("CR-002 section 65.3: the setters' table reads the objects table, each row Word's range narrowed only by the setter's stated policy", () => {
  const rows = Object.entries(WORD_STORAGE);
  assert.equal(rows.length, 11);
  for (const [name, row] of rows) {
    const { typeName, property, rule } = row.objects;
    const range = keptRange(typeName, property, rule);
    assert.ok(range, `${name}: in the objects table`);
    assert.ok(row.min >= range[0] && row.max <= range[1], `${name}: within Word's ${range}`);
    if (row.min !== range[0] || row.max !== range[1]) assert.ok(row.why, `${name}: a narrowing says why`);
    assert.equal(row.measured, true);
    // reads() is wordIntegerReading's answer
    for (const v of [-32769, -1, 0, 20, 32767, 32768, 32769, 65535, 65536, 8000000]) {
      const r = wordIntegerReading(typeName, property, v, rule);
      assert.ok(row.reads(v).includes(String(r.value)), `${name}.reads(${v}) = "${row.reads(v)}" names ${r.value}`);
    }
  }
  // Word's storage (`stored`), and the setters' policy over it: the indent and width setters at storage (D1), the spacing,
  // line, list and font setters at Word's own limits (D2, check 71b, section 65.9)
  assert.deepEqual([WORD_STORAGE.indLeft.min, WORD_STORAGE.indLeft.max, WORD_STORAGE.indFirstLine.max, WORD_STORAGE.indHanging.max, WORD_STORAGE.tcWDxa.min],
    [-32768, 32767, 32767, 32768, 0]);
  assert.deepEqual([WORD_STORAGE.spacingBefore.stored, WORD_STORAGE.spacingBefore.max, WORD_STORAGE.spacingAfter.max, WORD_STORAGE.spacingLineAuto.stored,
    WORD_STORAGE.spacingLineAuto.min, WORD_STORAGE.spacingLineAuto.max, WORD_STORAGE.lvlIndLeft.stored, WORD_STORAGE.lvlIndLeft.min, WORD_STORAGE.lvlIndLeft.max,
    WORD_STORAGE.sz.stored, WORD_STORAGE.sz.min, WORD_STORAGE.sz.max],
  [[0, 65535], 31680, 31680, [-32768, 32767], 15, 31680, [-32768, 32767], -31679, 31679, [0, 65535], 2, 3276]);
  assert.deepEqual([WORD_STORAGE.indFirstLine.reads(32768), WORD_STORAGE.indHanging.reads(32769), WORD_STORAGE.spacingLineAuto.reads(65535), WORD_STORAGE.spacingLineAuto.reads(5), WORD_STORAGE.indLeft.reads(7205001)],
    ['a hanging indent of 32768', 'a first-line indent of 32767', 'auto -1', 'auto 5', '-3959']);
});

// ---- rounds 5 and 6: Word's saves as the oracle for the out-of-range warnings ----

const ROUNDS = [
  { dir: 'cr066/', prefix: 'cr066-', files: 'groups/', valueColumn: 3, cases: 190 },
  { dir: 'cr066/round6/', prefix: 'cr066-r6-', files: 'groups/', singles: { groups: ['D'], dir: 'singles/' }, valueColumn: 4, cases: 478 },
];

/** Each attribute of the case lists: its class and property, the element steps from the case's block, the rule beside it, and the name Word writes it back under. */
const ATTRIBUTES = {
  'w:spacing/@w:before': [S + 'PPrBase.Spacing', 'before', ['pPr', 'spacing']],
  'w:spacing/@w:after': [S + 'PPrBase.Spacing', 'after', ['pPr', 'spacing']],
  'w:spacing/@w:line[auto]': [S + 'PPrBase.Spacing', 'line', ['pPr', 'spacing'], 'lineRule'],
  'w:spacing/@w:line[exact]': [S + 'PPrBase.Spacing', 'line', ['pPr', 'spacing'], 'lineRule'],
  'w:spacing/@w:line[atLeast]': [S + 'PPrBase.Spacing', 'line', ['pPr', 'spacing'], 'lineRule'],
  'w:ind/@w:left': [S + 'PPrBase.Ind', 'left', ['ind']],
  'w:ind/@w:right': [S + 'PPrBase.Ind', 'right', ['ind']],
  'w:ind/@w:start': [S + 'PPrBase.Ind', 'start', ['ind'], undefined, 'left'],
  'w:ind/@w:end': [S + 'PPrBase.Ind', 'end', ['ind'], undefined, 'right'],
  'w:ind/@w:firstLine': [S + 'PPrBase.Ind', 'firstLine', ['ind']],
  'w:ind/@w:hanging': [S + 'PPrBase.Ind', 'hanging', ['ind']],
  'w:tab/@w:pos': [S + 'CTTabStop', 'pos', ['tab']],
  'w:pBdr/w:top/@w:sz': [S + 'CTBorder', 'sz', ['pBdr', 'top']], 'w:pBdr/w:top/@w:space': [S + 'CTBorder', 'space', ['pBdr', 'top']],
  'w:bdr/@w:sz': [S + 'CTBorder', 'sz', ['rPr', 'bdr']], 'w:bdr/@w:space': [S + 'CTBorder', 'space', ['rPr', 'bdr']],
  'w:tblBorders/w:top/@w:sz': [S + 'CTBorder', 'sz', ['tblBorders', 'top']], 'w:tblBorders/w:top/@w:space': [S + 'CTBorder', 'space', ['tblBorders', 'top']],
  'w:tcBorders/w:top/@w:sz': [S + 'CTBorder', 'sz', ['tcBorders', 'top']], 'w:tcBorders/w:top/@w:space': [S + 'CTBorder', 'space', ['tcBorders', 'top']],
  'w:pgBorders/w:top/@w:sz': [S + 'CTBorder', 'sz', ['pgBorders', 'top']], 'w:pgBorders/w:top/@w:space': [S + 'CTBorder', 'space', ['pgBorders', 'top']],
  'w:sz/@w:val': [S + 'HpsMeasure', 'val', ['rPr', 'sz']], 'w:szCs/@w:val': [S + 'HpsMeasure', 'val', ['rPr', 'szCs']], 'w:kern/@w:val': [S + 'HpsMeasure', 'val', ['rPr', 'kern']],
  'w:position/@w:val': [S + 'CTSignedHpsMeasure', 'val', ['rPr', 'position']],
  'w:spacing/@w:val': [S + 'CTSignedTwipsMeasure', 'val', ['rPr', 'spacing']],
  'w:w/@w:val': [S + 'CTTextScale', 'val', ['rPr', 'w']],
  'w:fitText/@w:val': [S + 'CTFitText', 'val', ['rPr', 'fitText']],
  'w:tblW/@w:w[dxa]': [S + 'TblWidth', 'w', ['tblW']], 'w:tblW/@w:w[pct]': [S + 'TblWidth', 'w', ['tblW']],
  'w:tcW/@w:w[dxa]': [S + 'TblWidth', 'w', ['tcW']], 'w:tcW/@w:w[pct]': [S + 'TblWidth', 'w', ['tcW']],
  'w:tblInd/@w:w[dxa]': [S + 'TblWidth', 'w', ['tblInd']],
  'w:tblCellSpacing/@w:w': [S + 'TblWidth', 'w', ['tblCellSpacing']],
  'w:tblCellMar/w:left/@w:w': [S + 'TblWidth', 'w', ['tblCellMar', 'left']],
  'w:tcMar/w:top/@w:w': [S + 'TblWidth', 'w', ['tcMar', 'top']],
  'w:wBefore/@w:w': [S + 'TblWidth', 'w', ['wBefore']], 'w:wAfter/@w:w': [S + 'TblWidth', 'w', ['wAfter']],
  'w:gridCol/@w:w': [S + 'TblGridCol', 'w', ['gridCol']],
  'w:trHeight/@w:val': [S + 'CTHeight', 'val', ['trHeight'], 'hRule'],
  'w:trHeight/@w:val[exact]': [S + 'CTHeight', 'val', ['trHeight'], 'hRule'],
  'w:trHeight/@w:val[auto]': [S + 'CTHeight', 'val', ['trHeight'], 'hRule'],
  'w:framePr/@w:w': [S + 'CTFramePr', 'w', ['framePr']], 'w:framePr/@w:h': [S + 'CTFramePr', 'h', ['framePr'], 'hRule'],
  'w:framePr/@w:hSpace': [S + 'CTFramePr', 'hSpace', ['framePr']], 'w:framePr/@w:vSpace': [S + 'CTFramePr', 'vSpace', ['framePr']],
  'w:framePr/@w:x': [S + 'CTFramePr', 'x', ['framePr']], 'w:framePr/@w:y': [S + 'CTFramePr', 'y', ['framePr']],
  'w:tblpPr/@w:leftFromText': [S + 'CTTblPPr', 'leftFromText', ['tblpPr']], 'w:tblpPr/@w:rightFromText': [S + 'CTTblPPr', 'rightFromText', ['tblpPr']],
  'w:tblpPr/@w:topFromText': [S + 'CTTblPPr', 'topFromText', ['tblpPr']], 'w:tblpPr/@w:bottomFromText': [S + 'CTTblPPr', 'bottomFromText', ['tblpPr']],
  'w:tblpPr/@w:tblpX': [S + 'CTTblPPr', 'tblpX', ['tblpPr']], 'w:tblpPr/@w:tblpY': [S + 'CTTblPPr', 'tblpY', ['tblpPr']],
  'w:pgSz/@w:w': [S + 'SectPr.PgSz', 'w', ['pgSz']], 'w:pgSz/@w:h': [S + 'SectPr.PgSz', 'h', ['pgSz']],
  'w:pgMar/@w:top': [S + 'SectPr.PgMar', 'top', ['pgMar']], 'w:pgMar/@w:bottom': [S + 'SectPr.PgMar', 'bottom', ['pgMar']],
  'w:pgMar/@w:left': [S + 'SectPr.PgMar', 'left', ['pgMar']], 'w:pgMar/@w:right': [S + 'SectPr.PgMar', 'right', ['pgMar']],
  'w:pgMar/@w:header': [S + 'SectPr.PgMar', 'header', ['pgMar']], 'w:pgMar/@w:footer': [S + 'SectPr.PgMar', 'footer', ['pgMar']],
  'w:pgMar/@w:gutter': [S + 'SectPr.PgMar', 'gutter', ['pgMar']],
  'w:pgNumType/@w:start': [S + 'CTPageNumber', 'start', ['pgNumType']],
  'w:cols/@w:space': [S + 'CTColumns', 'space', ['cols']],
  'w:cols/w:col/@w:w': [S + 'CTColumn', 'w', ['cols', 'col']], 'w:cols/w:col/@w:space': [S + 'CTColumn', 'space', ['cols', 'col']],
  'w:lnNumType/@w:distance': [S + 'CTLineNumber', 'distance', ['lnNumType']],
  'w:defaultTabStop/@w:val': [S + 'CTTwipsMeasure', 'val', ['defaultTabStop']],
  'w:hyphenationZone/@w:val': [S + 'CTTwipsMeasure', 'val', ['hyphenationZone']],
};

/** What Word's saves don't let the test hold a case to, each stated, strict both ways: the warning then follows the objects table. */
const NOT_MEASURED = {
  'w:gridCol/@w:w': 'Word recomputes the grid, even under a fixed layout with no cell widths',
  'w:wBefore/@w:w': 'Word dropped it with the w:gridBefore that did not fit the row',
  'w:wAfter/@w:w': 'Word dropped it with the w:gridAfter that did not fit the row',
};

const textOf = (p) => Array.from(p.getElementsByTagNameNS(W, 't')).map((t) => t.textContent).join('');
const nextElement = (n) => { for (let x = n.nextSibling; x; x = x.nextSibling) if (x.nodeType === 1) return x; return undefined; };
const previousElement = (n) => { for (let x = n.previousSibling; x; x = x.previousSibling) if (x.nodeType === 1) return x; return undefined; };
const holdsSection = (p) => p.localName === 'p' && p.getElementsByTagNameNS(W, 'sectPr').length > 0 && textOf(p) === '';

/** The case's block: its label paragraph and, where the value is in it, the table or the section-ending paragraph after. */
function blockOf(doc, id) {
  if (doc.documentElement.localName === 'settings') return [doc.documentElement];
  const p = Array.from(doc.getElementsByTagNameNS(W, 'p')).find((x) => textOf(x).startsWith(`${id} `));
  if (!p) return undefined;
  const next = nextElement(p);
  return next && (next.localName === 'tbl' || holdsSection(next)) ? [p, next] : [p];
}

/** The label of the case a body child belongs to: itself, or the label paragraph before a table or a section-ending paragraph. */
function caseOfBodyChild(child) {
  const label = child.localName === 'tbl' || holdsSection(child) ? previousElement(child) : child;
  return label ? textOf(label).split(' ')[0] : undefined;
}

/** The first element the steps name in the block (each step the parent of the next). */
function elementIn(block, steps) {
  for (const root of block) {
    for (const e of [root, ...Array.from(root.getElementsByTagNameNS(W, steps.at(-1)))]) {
      if (e.localName !== steps.at(-1)) continue;
      let ok = true;
      for (let i = steps.length - 2, n = e.parentNode; i >= 0; i--, n = n?.parentNode) if (n?.localName !== steps[i]) { ok = false; break; }
      if (ok) return e;
    }
  }
  return undefined;
}

const integerAttribute = (e, name) => (e?.hasAttributeNS(W, name) ? Number(e.getAttributeNS(W, name)) : undefined);

/** The case an element of a part belongs to: the body child holding it (the label paragraph before a table or a section-ending paragraph). */
function caseOfElement(element) {
  let child = element;
  while (child.parentNode && child.parentNode.localName !== 'body') child = child.parentNode;
  return caseOfBodyChild(child);
}

/** The element a load-side warning names, by the line and column the DOM recorded for it in the source part. */
function elementAt(sourceDoc, warning) {
  const all = Array.from(sourceDoc.getElementsByTagNameNS(W, warning.element.replace(/^w:/, '')));
  return all.find((e) => e.lineNumber === warning.line && e.columnNumber === warning.column);
}

/** The body child a save-side path names: `/w:document/w:body[1]/w:p[5]/...` is the body's fifth w:p. */
function bodyChildOf(doc, path) {
  const m = /^\/w:document\/w:body\[1\]\/w:(\w+)\[(\d+)\]/.exec(path);
  if (!m) return undefined;
  const body = doc.getElementsByTagNameNS(W, 'body')[0];
  const children = Array.from(body.childNodes).filter((n) => n.nodeType === 1 && n.localName === m[1]);
  return children[Number(m[2]) - 1];
}

/** Every integer case of both rounds, with its file and part. */
const CASES = [];
for (const round of ROUNDS) {
  const list = readFileSync(new URL(`${round.dir}cases.txt`, DIR), 'utf8').trim().split('\n').slice(1).map((l) => l.split('\t'));
  for (const columns of list) {
    const [id, group, attribute] = columns;
    if (group === 'Z' || (group === 'N' && round.prefix === 'cr066-r6-')) continue; // round 6's N: Word dropped w:legacy whole
    const single = round.singles?.groups.includes(group);
    CASES.push({
      round: round.dir, id, group, attribute, value: columns[round.valueColumn],
      file: `${round.dir}${single ? round.singles.dir : round.files}${round.prefix}${single ? id : group}`,
      part: single ? 'word/settings.xml' : 'word/document.xml',
    });
  }
}
assert.deepEqual([CASES.filter((c) => c.round === 'cr066/').length, CASES.filter((c) => c.round === 'cr066/round6/').length], [190, 478]);
const INTEGER_CASES = CASES.filter((c) => c.group !== 'E');

test("rounds 5 and 6: a value.out_of_range warning exactly where Word's save changed the value, its rule or its attribute, with wordReads what Word wrote; on load and on save alike; the value written as held (CR-017 P2, P3, E5)", async () => {
  const files = [...new Set(INTEGER_CASES.map((c) => `${c.file}\t${c.part}`))].map((k) => k.split('\t'));
  const reportedOnLoad = new Set();
  const reportedOnSave = new Map();
  const written = new Map(); // file -> marshalled DOM of our save
  let onLoad = 0;
  for (const [file, partName] of files) {
    const pkg = await load(`${file}.docx`);
    const part = pkg.getPart(`/${partName}`);
    await part.getContents();
    const loadSide = OUT_OF_RANGE(part.warnings);
    onLoad += loadSide.length;
    for (const w of loadSide) {
      assert.equal(w.partName, `/${partName}`);
      assert.ok(typeof w.line === 'number' && w.path === undefined, `${file}: on load, a line and no path`);
    }
    const saved = unzipSync(await pkg.save());
    const ours = parseXml(strFromU8(saved[partName]));
    written.set(file, ours);
    const saveSide = OUT_OF_RANGE(pkg.lastSave.outOfRange).filter((w) => w.partName === `/${partName}`);
    assert.equal(OUT_OF_RANGE(pkg.lastSave.outOfRange).length, saveSide.length, `${file}: the probe's other parts hold nothing out of range`);
    // the same values on both sides
    const shape = (w) => `${w.element}/@${w.attribute}=${w.value} -> ${w.wordReads} ${w.readAs ?? ''} ${w.rule ?? ''} ${w.storage} ${JSON.stringify(w.range)}`;
    assert.deepEqual(loadSide.map(shape).sort(), saveSide.map(shape).sort(), `${file}: load and save report the same values`);
    for (const w of saveSide) {
      assert.ok(w.path && w.line === undefined, `${file}: on save, a path and no line`);
      const id = partName === 'word/settings.xml' ? file.split('/').pop().replace(/^cr066-(r6-)?/, '') : caseOfBodyChild(bodyChildOf(ours, w.path));
      assert.ok(INTEGER_CASES.some((c) => c.file === file && c.id === id && Number(c.value) === w.value), `${file}: ${w.path}=${w.value} belongs to case ${id}`);
      assert.ok(!reportedOnSave.has(`${file} ${id}`), `${file} ${id}: one warning per case`);
      reportedOnSave.set(`${file} ${id}`, w);
    }
    const source = parseXml(partXmlOf(`${file}.docx`, partName));
    for (const w of loadSide) {
      const id = partName === 'word/settings.xml' ? file.split('/').pop().replace(/^cr066-(r6-)?/, '') : caseOfElement(elementAt(source, w) ?? assert.fail(`${file}: ${w.message} names an element at its line and column`));
      assert.ok(!reportedOnLoad.has(`${file} ${id}`), `${file} ${id}: one warning per case on load`);
      reportedOnLoad.add(`${file} ${id}`);
    }
    // Word's own save holds nothing out of range: it writes only what it stores
    const word = await load(`${file}-word365.docx`);
    await word.getPart(`/${partName}`).getContents();
    assert.deepEqual(OUT_OF_RANGE(word.warnings), [], `${file}-word365: Word's save`);
  }
  assert.deepEqual([...reportedOnLoad].sort(), [...reportedOnSave.keys()].sort());

  // each case against Word's save
  let compared = 0;
  const leftOut = [];
  const stated = {};
  const parsedWord = new Map();
  for (const c of INTEGER_CASES) {
    const [typeName, property, steps, ruleAttribute, savedAs] = ATTRIBUTES[c.attribute] ?? assert.fail(`${c.id}: ${c.attribute} unmapped`);
    const value = Number(c.value);
    const key = `${c.file} ${c.id}`;
    const warning = reportedOnSave.get(key);
    // written as held, whatever Word reads
    assert.equal(integerAttribute(elementIn(blockOf(written.get(c.file), c.id), steps), property), value, `${c.id}: written as held`);
    if (NOT_MEASURED[c.attribute]) {
      // no oracle in Word's save: the warning follows the objects table (gridCol unmeasured, the certain range; wBefore and wAfter int16)
      const range = keptRange(typeName, property);
      assert.equal(warning !== undefined, value < range[0] || value > range[1], `${c.id}: ${c.attribute}, stated`);
      if (warning && typeName === S + 'TblGridCol') assert.deepEqual([warning.wordReads, warning.range], [null, UNMEASURED_RANGE], `${c.id}: unmeasured`);
      (stated[c.attribute] ??= []).push(c.id);
      continue;
    }
    const wordKey = `${c.file}-word365.docx ${c.part}`;
    if (!parsedWord.has(wordKey)) parsedWord.set(wordKey, parseXml(partXmlOf(`${c.file}-word365.docx`, c.part)));
    const saved = elementIn(blockOf(parsedWord.get(wordKey), c.id), steps);
    const storage = WORD_INTEGER_STORAGE[typeName][property];
    const writtenRule = c.attribute.match(/\[(auto|exact|atLeast)\]/)?.[1] ?? (storage.heightFlag ? 'exact' : undefined);
    let savedValue;
    let savedName = savedAs ?? property;
    if (storage.pair) {
      const positive = integerAttribute(saved, storage.pair.positive);
      const negative = integerAttribute(saved, storage.pair.negative);
      [savedName, savedValue] = negative !== undefined ? [storage.pair.negative, negative] : [storage.pair.positive, positive];
    } else {
      savedValue = integerAttribute(saved, savedName);
    }
    const savedRule = ruleAttribute ? (saved?.getAttributeNS(W, ruleAttribute) || undefined) : undefined;
    if (ruleAttribute === 'hRule' && writtenRule === 'auto' && storage.signRule) {
      assert.equal(savedValue, undefined, `${c.id}: a row height under auto, dropped by Word`);
      assert.equal(warning, undefined, `${c.id}: a row height under auto is ignored, not out of range`);
      compared++;
      continue;
    }
    if (savedValue === undefined) {
      // Word leaves out a value that reads as the default 0
      assert.ok(warning && warning.wordReads === 0, `${c.id} ${c.attribute}=${value}: Word left it out, so it read 0`);
      leftOut.push(key);
      continue;
    }
    if (storage.negative === 'computed' && warning && warning.wordReads === null) {
      assert.ok(savedValue > 0 && savedValue !== value, `${c.id}: a negative column width, computed by Word (${savedValue})`);
      assert.match(warning.message, /Word reads it as a width Word computes$/);
      compared++;
      continue;
    }
    // Word kept the value exactly when its save writes it back under the same name and rule, or 65535 where Word writes it -1
    const sameInSave = savedValue === value && savedName === (savedAs ?? property) && (writtenRule === undefined || savedRule === writtenRule
      || (writtenRule === 'atLeast' && ruleAttribute === 'lineRule' && savedRule === 'atLeast'));
    const rowRuleSame = ruleAttribute !== 'hRule' || storage.heightFlag || savedRule === (writtenRule === 'exact' ? 'exact' : undefined);
    const kept = (sameInSave && rowRuleSame) || (storage.maxWrittenAs !== undefined && value === 65535);
    assert.equal(warning === undefined, kept, `${c.id} ${c.attribute}=${value}: Word's save wrote ${savedName}=${savedValue}${savedRule ? ` ${ruleAttribute}=${savedRule}` : ''}; warning ${warning ? 'given' : 'none'}`);
    if (warning) {
      assert.equal(warning.wordReads, savedValue, `${c.id}: wordReads is what Word's save wrote`);
      if (storage.pair) assert.equal(warning.readAs, savedName === property ? undefined : `w:${savedName}`, `${c.id}: read as ${savedName}`);
      if (ruleAttribute === 'lineRule') assert.equal(warning.rule, savedRule, `${c.id}: the line rule`);
      if (ruleAttribute === 'hRule' && !storage.heightFlag) assert.equal(warning.rule === 'exact' ? 'exact' : undefined, savedRule, `${c.id}: the row rule`);
      if (storage.heightFlag) assert.equal(warning.rule === 'exact' ? 'exact' : undefined, savedRule, `${c.id}: the frame's rule`);
      assert.ok(warning.range === null || value < warning.range[0] || value > warning.range[1], `${c.id}: outside the range the warning names`);
    }
    compared++;
  }
  assert.deepEqual(Object.fromEntries(Object.entries(stated).map(([k, v]) => [k, v.length])), { 'w:gridCol/@w:w': 10, 'w:wBefore/@w:w': 10, 'w:wAfter/@w:w': 10 });
  assert.deepEqual([INTEGER_CASES.length, compared, leftOut.length, reportedOnSave.size, onLoad], [660, 614, 16, 432, 432],
    'the cases, the compared, Word\'s left-out zeros, the warnings (the objects package\'s own count, its CR-016 section 7)');
});

test("round 6's group E: the forms with no digit read as 0, value.kept.invalid, as Word's save writes them under a style's 100 (CR-017 P1)", async () => {
  const groupE = CASES.filter((c) => c.group === 'E');
  assert.equal(groupE.length, 8);
  const pkg = await load('cr066/round6/groups/cr066-r6-E.docx');
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  const invalid = main.warnings.filter((w) => w.code === 'value.kept.invalid');
  assert.deepEqual(invalid.map((w) => [w.text, w.value]).sort(), groupE.map((c) => [c.value, 0]).sort());
  assert.deepEqual(OUT_OF_RANGE(main.warnings), []);
  const word = parseXml(partXmlOf('cr066/round6/groups/cr066-r6-E-word365.docx', 'word/document.xml'));
  const ours = parseXml(await main.getXml());
  for (const c of groupE) {
    const [, property, steps] = ATTRIBUTES[c.attribute];
    assert.equal(integerAttribute(elementIn(blockOf(ours, c.id), steps), property), 0, `${c.id} "${c.value}": read as 0`);
    assert.equal(integerAttribute(elementIn(blockOf(word, c.id), steps), property), 0, `${c.id}: Word's save wrote 0`);
  }
});

test('the warning: its message, fields and OpcPackage.warnings, one of each kind (CR-017 E1, E4)', async () => {
  const pkg = await load('cr066/groups/cr066-P.docx');
  assert.equal(pkg.warnings.length, 0, 'nothing is read until a part is');
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  const warnings = OUT_OF_RANGE(main.warnings);
  assert.equal(pkg.warnings.length, main.warnings.length, 'OpcPackage.warnings gathers them');
  const find = (attribute, value) => warnings.find((w) => w.attribute === attribute && w.value === value);
  const wrap = find('w:left', 7205001);
  assert.deepEqual({ ...wrap, line: undefined, column: undefined, message: undefined }, {
    code: 'value.out_of_range', partName: '/word/document.xml', element: 'w:ind', attribute: 'w:left', value: 7205001, wordReads: -3959,
    storage: 'int16', range: [-32768, 32767], line: undefined, column: undefined, message: undefined,
  });
  assert.match(wrap.message, /^\/word\/document\.xml: w:ind\/@w:left=7205001 \(line \d+, column \d+\) is past what Word 365 stores there \(int16, -32768 to 32767\); Word reads it as -3959$/);
  const clamp = find('w:pos', 32767);
  assert.deepEqual([clamp.wordReads, clamp.storage, clamp.range], [31680, 'int16, then ±31680', [-31680, 31680]]);
  const rule = warnings.find((w) => w.attribute === 'w:line' && w.value === 65535 && w.storage.includes('rule'));
  assert.deepEqual([rule.wordReads, rule.rule, rule.storage, rule.range], [1, 'atLeast', 'int16, the sign the rule', [1, 32768]], 'exact 65535 reads as atLeast 1');
  assert.match(rule.message, /Word reads it as atLeast 1$/);
  const s = await load('cr066/groups/cr066-S.docx');
  await s.getMainDocumentPart().getContents();
  const page = OUT_OF_RANGE(s.warnings).find((w) => w.element === 'w:pgSz' && w.attribute === 'w:w' && w.value === 65536);
  assert.deepEqual([page.wordReads, page.storage, page.range], [144, 'uint16, then 144..31681', [144, 31681]]);

  const r6 = await load('cr066/round6/groups/cr066-r6-A.docx');
  await r6.getMainDocumentPart().getContents();
  const pair = OUT_OF_RANGE(r6.warnings).find((w) => w.attribute === 'w:firstLine' && w.value === 32768);
  assert.deepEqual([pair.wordReads, pair.readAs, pair.range], [32768, 'w:hanging', [0, 32767]]);
  assert.match(pair.message, /Word reads it as 32768 in w:hanging$/);
  const hanging = OUT_OF_RANGE(r6.warnings).find((w) => w.attribute === 'w:hanging' && w.value === 65535);
  assert.deepEqual([hanging.wordReads, hanging.readAs, hanging.range], [1, 'w:firstLine', [0, 32768]]);

  const tables = await load('cr066/round6/groups/cr066-r6-T.docx');
  await tables.getMainDocumentPart().getContents();
  const grid = OUT_OF_RANGE(tables.warnings).find((w) => w.element === 'w:gridCol' && w.value === 7205001);
  assert.deepEqual([grid.wordReads, grid.range, grid.storage], [null, [-32768, 65535], 'unmeasured: changed outside -32768..65535']);
  assert.match(grid.message, /is outside -32768 to 65535, where Word 365 changes every measured value of its kind \(w:gridCol\/@w:w is unmeasured\); what Word reads is not known$/);
  assert.equal(OUT_OF_RANGE(tables.warnings).find((w) => w.element === 'w:gridCol' && w.value === 65535), undefined, 'inside the certain range, not reported');

  const sections = await load('cr066/round6/groups/cr066-r6-S.docx');
  await sections.getMainDocumentPart().getContents();
  const column = OUT_OF_RANGE(sections.warnings).find((w) => w.element === 'w:col' && w.attribute === 'w:w' && w.value === -1);
  assert.deepEqual([column.wordReads, column.range, column.storage], [null, [0, 32767], 'int16, a negative computed']);
  assert.match(column.message, /Word reads it as a width Word computes$/);
  const space = OUT_OF_RANGE(sections.warnings).find((w) => w.element === 'w:col' && w.attribute === 'w:space' && w.value === 65535);
  assert.deepEqual([space.wordReads, space.range], [-1, [0, 32767]]);
  assert.equal(OUT_OF_RANGE(sections.warnings).find((w) => w.element === 'w:col' && w.attribute === 'w:space' && w.value === -1), undefined, "-1 is Word's own sentinel, kept");
  const margin = OUT_OF_RANGE(sections.warnings).find((w) => w.element === 'w:pgMar' && w.attribute === 'w:left' && w.value === 65535);
  assert.equal(margin, undefined, '65535 in an attribute Word writes back as -1 is read as written');
});

test("a value set by code is reported on save with its path, written as held, under the normaliser or not; a clean package reports [] (CR-017 P3, E3)", async () => {
  const pkg = await load('cr066/groups/cr066-P-word365.docx');
  const main = pkg.getMainDocumentPart();
  const doc = await main.getContents();
  assert.deepEqual(OUT_OF_RANGE(main.warnings), [], "Word's save holds nothing out of range");
  const unwrap = (x) => (x !== null && typeof x === 'object' && x.TYPE_NAME === undefined && 'value' in x ? x.value : x);
  const paragraphs = doc.body.content.map(unwrap).filter((x) => x?.TYPE_NAME === 'org_docx4j_wml.P');
  const withInd = paragraphs.find((p) => p.pPr?.ind?.left !== undefined && p.pPr.ind.left !== null);
  assert.ok(withInd, 'a paragraph with w:ind/@w:left');
  withInd.pPr.ind.left = 7205001;
  const first = paragraphs.find((p) => p.pPr?.spacing?.before !== undefined && p.pPr.spacing.before !== null);
  first.pPr.spacing.before = 65536;
  for (const normalise of [undefined, false]) {
    pkg.normaliseOnSave = normalise;
    const saved = unzipSync(await pkg.save());
    const reports = pkg.lastSave.outOfRange;
    assert.equal(reports.length, 2, `normaliseOnSave ${normalise}: both reported`);
    const left = reports.find((r) => r.attribute === 'w:left');
    assert.match(left.path, /^\/w:document\/w:body\[1\]\/w:p\[\d+\]\/w:pPr\[1\]\/w:ind\[1\]\/@w:left$/);
    assert.deepEqual({ ...left, path: undefined, message: undefined }, {
      code: 'value.out_of_range', partName: '/word/document.xml', element: 'w:ind', attribute: 'w:left', value: 7205001, wordReads: -3959,
      storage: 'int16', range: [-32768, 32767], path: undefined, message: undefined,
    });
    assert.equal(left.message, `/word/document.xml: ${left.path}=7205001 is past what Word 365 stores there (int16, -32768 to 32767); Word reads it as -3959`);
    assert.equal(left.line, undefined);
    const xml = strFromU8(saved['word/document.xml']);
    assert.ok(xml.includes('w:left="7205001"') && xml.includes('w:before="65536"'), 'written as held, never rewritten');
    assert.deepEqual(OUT_OF_RANGE(main.warnings), [], "the part's own warnings are the parse's; a save reports on lastSave");
  }
  // a package nothing is wrong with
  const clean = await WordprocessingMLPackage.createPackage();
  clean.body.insertParagraph('plain', 'End');
  await clean.save();
  assert.deepEqual(clean.lastSave.outOfRange, []);
  // a part nobody unmarshalled is copied, so nothing in it is reported
  const untouched = await load('cr066/groups/cr066-P.docx');
  await untouched.save();
  assert.deepEqual(untouched.lastSave.outOfRange, []);
});

// ---- round 2 (CR-065): decimals and odd forms, Word's saves as the oracle for the readings ----

const R2 = 'cr065/';
const R2_ALL = readdirSync(new URL(`${R2}families/`, DIR)).filter((d) => d.endsWith('.docx') && !d.endsWith('-word365.docx')).map((d) => d.slice(0, -'.docx'.length)).sort();
// the whole-family files Word refused (round 2's A, round 3's A2), split for rounds 3 and 4; the split files hold each case once
const R2_WHOLE = ['cr065-A', 'cr065-A2-forms'];
const R2_FILES = R2_ALL.filter((f) => !R2_WHOLE.includes(f));
const r2Saved = (file) => R2_FILES.length > 0 && statSync(new URL(`${R2}families/${file}-word365.docx`, DIR), { throwIfNoEntry: false }) !== undefined;

/** The cases whose value cases.txt does not give as one path. */
const R2_SPECIAL = {
  H56: { steps: ['pgSz'], attributes: ['w', 'h'] }, H57: { steps: ['pgSz'], attributes: ['w', 'h'] },
  H60: { steps: ['pgMar'], attributes: ['header', 'footer', 'gutter'] },
  I62: { part: 'styles', steps: ['ind'], attributes: ['left'] }, I63: { part: 'settings', steps: ['defaultTabStop'], attributes: ['val'] },
  I64: { part: 'header', steps: ['spacing'], attributes: ['before'] },
};
const R2_CASES = readFileSync(new URL(`${R2}cases.txt`, DIR), 'utf8').trim().split('\n').map((line) => line.split('\t')[1])
  .filter((c) => !c.startsWith('Z')).map((c) => {
    const id = c.split(' ')[0];
    const m = c.match(/^\S+ ((?:w:\w+\/)*w:\w+)\/@w:(\w+)="/);
    return m ? { id, steps: m[1].split('/').map((s) => s.slice(2)), attributes: [m[2]] } : { id, ...R2_SPECIAL[id] };
  });

/** Where this package's reading and Word's save differ, each stated, strict both ways: a value Word reads as what lies beneath and leaves out, and the grid it recomputes. */
const R2_KNOWN = {
  A08: 'truncated to 0, the default; Word leaves it out', A17: '"" read as 0, the default; Word leaves it out',
  B25: 'truncated to 0, the default; Word leaves it out', F41: 'truncated to 22, the document default; Word leaves it out',
  G50: 'w:gridCol, recomputed by Word', G51: 'w:gridCol, recomputed by Word',
};
/** The forms of the files Word refused, each dropped and reported as value.refused. */
const R2_REFUSED = { A11: ' 240.5 ', A12: '240,5', A13: '0x10', A14: '1.5e2', A15: '2.5e0', A16: '1e3', A18: 'abc', A19: ' 240 ' };

function valuesIn(scope, { steps, attributes }) {
  const out = [];
  for (const root of scope) {
    const candidates = [root, ...Array.from(root.getElementsByTagNameNS(W, steps.at(-1)))].filter((e) => e.localName === steps.at(-1));
    for (const e of candidates) {
      let ok = true;
      for (let i = steps.length - 2, n = e.parentNode; i >= 0; i--, n = n?.parentNode) if (n?.localName !== steps[i]) { ok = false; break; }
      if (!ok) continue;
      for (const a of attributes) if (e.hasAttributeNS(W, a)) out.push(`${a}=${e.getAttributeNS(W, a)}`);
    }
  }
  return out;
}

test("round 2: a decimal reads truncated toward zero, the no-digit forms as 0, a form Word refuses dropped and reported, each as Word's save has it; Word's saves need no reading (CR-017 P1, E5)", async () => {
  assert.deepEqual([R2_ALL.length, R2_FILES.length], [22, 20], 'A1 to A7 split seven ways (A2 six), and B to I, beside the two whole-family files Word refused');
  assert.equal(R2_CASES.length, 64);
  const partNameOf = (pkg, which) => (which === 'document' ? '/word/document.xml' : which === 'styles' ? '/word/styles.xml' : which === 'settings' ? '/word/settings.xml'
    : pkg.getMainDocumentPart().headerParts[0].partName.name);
  const loaded = new Map();
  async function read(file, which) {
    const key = `${file} ${which}`;
    if (!loaded.has(key)) {
      const pkg = await load(`${R2}families/${file}.docx`);
      const part = pkg.getPart(partNameOf(pkg, which));
      await part.getContents();
      const ours = { doc: parseXml(await part.getXml()), warnings: part.warnings };
      let word;
      if (r2Saved(file)) {
        const saved = await load(`${R2}families/${file}-word365.docx`);
        const savedPart = saved.getPart(partNameOf(saved, which));
        await savedPart.getContents();
        assert.deepEqual(savedPart.warnings.filter((w) => w.code.startsWith('value.')), [], `${file}: Word's save needs no reading of Word's`);
        word = { doc: parseXml(await savedPart.getXml()) };
      }
      loaded.set(key, { ours, word });
    }
    return loaded.get(key);
  }
  const fileOf = (c) => R2_FILES.find((f) => f.startsWith(`cr065-${c.id[0]}`) && new RegExp(`>${c.id} `).test(partXmlOf(`${R2}families/${f}.docx`, 'word/document.xml')));
  let equal = 0;
  const differing = {};
  const kinds = {};
  for (const c of R2_CASES) {
    const file = fileOf(c);
    assert.ok(file, `${c.id}: in a probe file`);
    const which = c.part ?? 'document';
    const { ours, word } = await read(file, which);
    const scope = (doc) => (which === 'document' ? blockOf(doc, c.id)
      : which === 'styles' ? Array.from(doc.getElementsByTagNameNS(W, 'style')).filter((st) => st.getAttributeNS(W, 'styleId') === 'ProbeStyle')
        : [doc.documentElement]);
    const ourValues = valuesIn(scope(ours.doc), c);
    if (!word) {
      assert.ok(c.id in R2_REFUSED, `${c.id}: a refused form`);
      assert.deepEqual(ourValues, [], `${c.id}: dropped`);
      const refused = ours.warnings.find((w) => w.code === 'value.refused' && w.text === R2_REFUSED[c.id]);
      assert.ok(refused && refused.value === undefined && /dropped: Word 365 refuses a document with this value$/.test(refused.message), `${c.id}: reported as refused`);
      kinds.refused = (kinds.refused ?? 0) + 1;
      continue;
    }
    const wordValues = valuesIn(scope(word.doc), c);
    if (JSON.stringify(ourValues) === JSON.stringify(wordValues)) equal++;
    else differing[c.id] = `${JSON.stringify(ourValues)} vs Word ${JSON.stringify(wordValues)}`;
  }
  assert.deepEqual(Object.keys(differing).sort(), Object.keys(R2_KNOWN).sort(), `the known differences, strict both ways: ${JSON.stringify(differing, null, 1)}`);
  assert.deepEqual([equal, kinds.refused], [50, 8]);
  // the warnings over the families Word opened: a decimal is value.truncated with Word's integer, "" is value.kept.invalid read as 0
  const b = await read('cr065-B', 'document');
  const truncated = b.ours.warnings.filter((w) => w.code === 'value.truncated');
  assert.deepEqual(truncated.map((w) => [w.text, w.value]), [['720.5', 720], ['720.4', 720], ['721.5', 721], ['-120.5', -120], ['-121.5', -121], ['-0.5', 0], ['-1.5', -1], ['10206.0', 10206]]);
  assert.match(truncated[0].message, /w:ind\/@w:left="720.5" \(line \d+, column \d+\) is not a value its type allows; read as 720, as Word 365 reads it$/);
  const a6 = await read('cr065-A6-empty', 'document');
  assert.deepEqual(a6.ours.warnings.filter((w) => w.code.startsWith('value.')).map((w) => [w.code, w.text, w.value]), [['value.kept.invalid', '', 0]]);
  // the all-cases file Word refused outright: every part loads here, the refused forms dropped
  const all = await load(`${R2}cr065-decimals-probe.docx`);
  await all.unmarshalAll();
  const codes = {};
  for (const w of all.warnings) codes[w.code] = (codes[w.code] ?? 0) + 1;
  assert.deepEqual(codes, { 'value.truncated': 55, 'value.kept.invalid': 5, 'value.refused': 8 }, 'A17 and H60\'s four empties read as 0; A11 to A16, A18 and A19 refused');
  assert.deepEqual(all.warnings.filter((w) => w.code === 'value.refused').map((w) => w.text).sort(), Object.values(R2_REFUSED).sort());
  // and the two whole-family files Word refused: family A's eight forms, A2's two with whitespace
  for (const [file, refused] of [['cr065-A', Object.values(R2_REFUSED)], ['cr065-A2-forms', [' 240.5 ', ' 240 ']]]) {
    const pkg = await load(`${R2}families/${file}.docx`);
    const main = pkg.getMainDocumentPart();
    await main.getContents();
    assert.deepEqual(main.warnings.filter((w) => w.code === 'value.refused').map((w) => w.text).sort(), refused.sort(), file);
  }
});

test('no fixture holds a value past what Word stores (CR-017 E4)', async () => {
  const fixtures = new URL('./fixtures/', import.meta.url);
  const files = [];
  const walk = (d) => { for (const n of readdirSync(d).sort()) { const p = join(d, n); if (statSync(p).isDirectory()) walk(p); else if (n.endsWith('.docx')) files.push(p); } };
  walk(fileURLToPath(fixtures));
  let parts = 0;
  for (const f of files) {
    let pkg;
    try { pkg = await WordprocessingMLPackage.load(new Uint8Array(readFileSync(f))); } catch { continue; }
    const main = pkg.getMainDocumentPart();
    if (!main) continue;
    try { await main.getContents(); } catch { continue; }
    parts++;
    assert.deepEqual(OUT_OF_RANGE(main.warnings).map((w) => w.message), [], f);
  }
  assert.ok(parts > 200, `${parts} main parts read`);
});
