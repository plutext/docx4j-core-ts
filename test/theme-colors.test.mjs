// CR-008 phase 2: the colour resolver (docx4j org.docx4j.model.color, its CR-034) against Office. The vectors are
// docx4j's (test/theme-colors/, copied unedited at the commit its manifest.json names): seven probe files of 1,068
// labelled cases, what Word, PowerPoint and Excel 365 drew for each (readings.tsv, read from their PDFs), and the 117
// distinct tint and shade triples of four real-document corpora (word-tint-shade.tsv). Each probe is loaded, each case
// resolved through its host's ThemeColors and compared with Office's reading, as docx4j's ThemeColourProbesTest does,
// at its bar: exact, or within one unit where CR-034 section 11 records it; the luminance transforms exact but T239;
// the saturation set on a grey (T290 to T298) the one known family off. Then docx4j's ThemeColorsTest: which theme and
// colour map a host takes.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync } from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import {
  OpcPackage, WordprocessingMLPackage, ThemeColors, WordColors, DrawingMLColors, SpreadsheetColors, ResolvedColor,
  ChartPart, ThemeOverridePart, parseXml, unmarshalNode,
} from '../dist/index.mjs';

const dir = new URL('./theme-colors/', import.meta.url);
const bytesOf = (name) => readFileSync(new URL(name, dir));
const tsv = (name) => readFileSync(new URL(name, dir), 'utf8').split(/\r?\n/).slice(1).filter((line) => line.length > 0);

/** Office's reading per probe and label: page, fill, stroke, text, rules. */
const READINGS = new Map();
for (const line of tsv('readings.tsv')) {
  const f = `${line}\t\t\t\t\t\t`.split('\t');
  if (!READINGS.has(f[0])) READINGS.set(f[0], new Map());
  READINGS.get(f[0]).set(f[1], [f[2], f[3], f[4], f[5], f[6]]);
}
/** cases.tsv's note per probe/label: a P3 label's family. */
const NOTES = new Map();
for (const line of tsv('cases.tsv')) {
  const f = line.split('\t');
  NOTES.set(`${f[0]}/${f[1]}`, f[4]);
}

const expected = (probe, label, column) => READINGS.get(probe).get(label)?.[column];

/** Saturation set on the grey bg2 (E8E8E8): Office gives a grey a hue of about 40 degrees; docx4j's (and this) has 0. */
const known = (label) => NOTES.get(`tc-p3-transforms.docx/${label}`)?.startsWith('bg2 sat') === true;

function family(label) {
  const note = NOTES.get(`tc-p3-transforms.docx/${label}`);
  if (note === undefined) return label.substring(0, 2);
  if (note.includes('picker')) return 'picker';
  const w = note.split(' ');
  return w.length > 1 && /^(accent\d|tx2|bg2)$/.test(w[0]) ? w[1] : w[0];
}

/** exact, within one unit, further: counted per family, failures listed (docx4j's Tally). */
class Tally {
  families = new Map();
  off = [];

  add(fam, label, got, want) {
    let f = this.families.get(fam);
    if (f === undefined) this.families.set(fam, f = [0, 0, 0]);
    const hex = want.includes('/') ? want.substring(0, want.indexOf('/')) : want;
    const w = ResolvedColor.ofHex(hex);
    if (got !== undefined && got.toHex() === hex) f[0]++;
    else if (got !== undefined && got.within(w, 1)) f[1]++;
    else {
      f[2]++;
      this.off.push(`${label} ${fam}: core-ts ${got}, Office ${want}`);
    }
    if (want.includes('/') && got !== undefined) {
      const a = Number.parseFloat(want.substring(want.indexOf('/') + 1));
      if (Math.abs(got.alpha / 100000.0 - a) > 0.004) this.off.push(`${label} alpha: core-ts ${got}, Office ${want}`);
    }
  }

  report() {
    return [...this.families.entries()].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
      .map(([k, [e, o, m]]) => `${k.padEnd(34)} exact ${String(e).padStart(3)}, one unit ${String(o).padStart(2)}, more ${String(m).padStart(2)}`).join('\n');
  }

  count(column) {
    let n = 0;
    for (const f of this.families.values()) n += f[column];
    return n;
  }

  /** the luminance transforms (and the picker's, which are luminance) not exact */
  luminanceOff() {
    let n = 0;
    for (const fam of ['lum', 'lumMod', 'lumOff', 'picker']) {
      const c = this.families.get(fam);
      if (c !== undefined) n += c[1] + c[2];
    }
    return n;
  }
}

/** every case reproduced exactly or within one unit, except the known family; at least `cases` read; docx4j's tallies */
function check(t, what, tally, cases, docx4j) {
  t.diagnostic(`${what}:\n${tally.report()}`);
  assert.deepEqual(Object.fromEntries(tally.families), docx4j, `${what}: docx4j's tallies`);
  const unexpected = tally.off.filter((o) => !known(o.substring(0, o.indexOf(' '))));
  assert.deepEqual(unexpected, [], what);
  const read = tally.count(0) + tally.count(1) + tally.count(2);
  assert.ok(read >= cases, `${what} cases read: ${read}`);
}

/**
 * docx4j's own tallies, family by family (exact, within one unit, further): ThemeColourProbesTest's output at docx4j
 * 1fffc20a3 (VERSION_17_3_5, 2026-10-24; org.docx4j.model.color unchanged since 691620fb7). P3's are the same in Word and
 * PowerPoint. The port computes as docx4j does, so its tallies are these, not just within docx4j's bar.
 */
const DOCX4J = {
  P3: {
    'alpha': [15, 0, 0], 'alphaMod': [10, 0, 0], 'alphaOff': [10, 0, 0], 'blue': [15, 0, 0], 'blueMod': [10, 0, 0],
    'blueOff': [10, 0, 0], 'chart': [2, 0, 0], 'comp': [5, 0, 0], 'gamma': [5, 0, 0], 'gray': [5, 0, 0],
    'green': [15, 0, 0], 'greenMod': [10, 0, 0], 'greenOff': [10, 0, 0], 'hslClr': [1, 0, 0], 'hue': [15, 0, 0],
    'hueMod': [13, 2, 0], 'hueOff': [9, 6, 0], 'inv': [5, 0, 0], 'invGamma': [5, 0, 0], 'lum': [15, 0, 0],
    'lumMod': [16, 1, 0], 'lumOff': [16, 0, 0], 'picker': [50, 0, 0], 'prstClr': [1, 0, 0], 'red': [15, 0, 0],
    'redMod': [10, 0, 0], 'redOff': [10, 0, 0], 'sat': [11, 2, 2], 'satMod': [13, 3, 0], 'satOff': [12, 0, 3],
    'schemeClr': [2, 0, 0], 'scrgbClr': [1, 0, 0], 'shade': [16, 0, 0], 'srgbClr': [2, 0, 0],
    'style: fill': [3, 0, 0], 'style: font': [4, 0, 0], 'style: line': [2, 0, 0], 'sysClr': [2, 0, 0],
    'the': [1, 0, 0], 'tint': [17, 0, 0],
  },
  P1: {
    'paragraph shading': [2, 1, 0], 'pattern blend': [1, 0, 0], 'run shading': [1, 0, 0], 'run text': [8, 0, 0],
    'top border': [3, 0, 0], 'underline': [1, 0, 0],
  },
  P2: {
    'P2': [7, 0, 0], 'run text': [21, 0, 0],
  },
  P4: {
    'indexed': [66, 0, 0], 'indexed, custom palette': [66, 0, 0], 'theme': [10, 0, 0], 'theme + tint': [80, 0, 0],
  },
  P5: {
    'fill': [10, 0, 0], 'text (tx1)': [10, 0, 0],
  },
};

// ------------------------------------------------------------------ the docx probes, read as XML

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

const xml = (probe, part) => parseXml(strFromU8(unzipSync(bytesOf(probe))[part]));

function child(e, ns, name) {
  for (let n = e?.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && n.namespaceURI === ns && n.localName === name) return n;
  }
  return undefined;
}
const elements = (e, ns, name) => Array.from(e.getElementsByTagNameNS(ns, name));
const text = (e) => elements(e, W, 't').map((t) => t.textContent).join('');
function attr(e, name) {
  if (e === undefined) return undefined;
  const v = e.getAttributeNS(W, name);
  return v === null || v === '' ? undefined : v;
}

function styled(probe, label, s, theme, tally) {
  const fill = DrawingMLColors.fillReference(s.fillRef, theme);
  if (fill.kind === 'SOLID') tally.add('style: fill', label, fill.color, expected(probe, label, 1));
  // not solid: Office drew no single-colour fill for it either
  else assert.ok(expected(probe, label, 1) === '' || fill.kind !== 'UNKNOWN', `${label} ${fill}`);
  const line = DrawingMLColors.lineReference(s.lnRef, theme);
  const stroke = expected(probe, label, 2);
  if (line.kind === 'SOLID' && stroke !== '') tally.add('style: line', label, line.color, stroke);
  else if (line.kind === 'NONE') assert.equal(stroke, '', `${label} has no line`);
  tally.add('style: font', label, DrawingMLColors.fontReference(s.fontRef, theme), expected(probe, label, 3));
}

/** The wps shapes of a docx probe, each unmarshalled as docx4j's test does: label, fill or style. */
async function wordShapes(probe, tally) {
  const theme = await ThemeColors.of(await OpcPackage.load(bytesOf(probe)));
  for (const wsp of elements(xml(probe, 'word/document.xml'), WPS, 'wsp')) {
    const label = text(wsp).trim();
    if (expected(probe, label, 1) === undefined) continue;
    const shape = (await unmarshalNode(wsp)).value;
    if (child(child(wsp, WPS, 'spPr'), A, 'solidFill') !== undefined) {
      tally.add(family(label), label, DrawingMLColors.resolve(shape.spPr.solidFill, theme), expected(probe, label, 1));
    } else if (shape.style !== undefined) {
      styled(probe, label, shape.style, theme, tally);
    }
  }
}

/** A slide shape's label: its text. */
const label = (sp) => (sp.txBody?.p ?? []).flatMap((p) => (p.egTextRun ?? []).filter((r) => r.TYPE_NAME === 'org_docx4j_dml.CTRegularTextRun').map((r) => r.t)).join('').trim();

/** Each slide's shapes with its ThemeColors. */
async function slideShapes(probe) {
  const pptx = await OpcPackage.load(bytesOf(probe));
  const out = [];
  for (const slide of pptx.slideParts) {
    const theme = await ThemeColors.of(slide);
    for (const sp of (await slide.readContents()).cSld.spTree.spOrGrpSpOrGraphicFrame ?? []) {
      if (sp.TYPE_NAME === 'org_pptx4j_pml.Shape' && expected(probe, label(sp), 1) !== undefined) out.push({ sp, theme, label: label(sp) });
    }
  }
  return out;
}

// ------------------------------------------------------------------ the vectors

test('the vectors are docx4j\'s, as copied: every file listed, every SHA-1 matching', () => {
  const manifest = JSON.parse(readFileSync(new URL('manifest.json', dir), 'utf8'));
  const listed = new Map(manifest.files.map((f) => [f.name, f.sha1]));
  const present = readdirSync(dir).filter((n) => n !== 'manifest.json').sort();
  assert.deepEqual(present, [...listed.keys()].sort());
  for (const name of present) assert.equal(createHash('sha1').update(bytesOf(name)).digest('hex'), listed.get(name), name);
  assert.equal(tsv('cases.tsv').length, 1068);
  assert.equal(tsv('readings.tsv').length, 1068);
  assert.equal(tsv('word-tint-shade.tsv').length, 117);
});

test('P3 in Word: the 403 DrawingML transform cases, the luminance families exact but T239 (docx4j CR-034 section 11, follow-up)', async (t) => {
  const tally = new Tally();
  await wordShapes('tc-p3-transforms.docx', tally);
  check(t, 'P3, Word', tally, 403, DOCX4J.P3);
  assert.ok(tally.luminanceOff() <= 1, `P3, Word: luminance exact but T239\n${tally.report()}`);
});

test('P3 in PowerPoint: the same 403 cases on slides, each through its slide\'s ThemeColors', async (t) => {
  const probe = 'tc-p3-transforms.pptx';
  const tally = new Tally();
  for (const { sp, theme, label: l } of await slideShapes(probe)) {
    if (sp.spPr?.solidFill !== undefined) tally.add(family(l), l, DrawingMLColors.resolve(sp.spPr.solidFill, theme), expected(probe, l, 1));
    else if (sp.style !== undefined) styled(probe, l, sp.style, theme, tally);
  }
  check(t, 'P3, PowerPoint', tally, 403, DOCX4J.P3);
  assert.ok(tally.luminanceOff() <= 1, `P3, PowerPoint: luminance exact but T239\n${tally.report()}`);
});

test('P1: a WordprocessingML theme reference wins over the value beside it; tint and shade in HSL luminance; exact but P1-10', async (t) => {
  const probe = 'tc-p1-precedence.docx';
  const theme = await ThemeColors.of(await OpcPackage.load(bytesOf(probe)));
  const tally = new Tally();
  for (const p of elements(xml(probe, 'word/document.xml'), W, 'p')) {
    const txt = text(p);
    if (!txt.startsWith('P1-')) continue;
    const l = txt.substring(0, 5);
    const pPr = child(p, W, 'pPr');
    const rPr = child(child(p, W, 'r'), W, 'rPr');
    const [color, u, rShd, pShd] = [child(rPr, W, 'color'), child(rPr, W, 'u'), child(rPr, W, 'shd'), child(pPr, W, 'shd')];
    const top = child(child(pPr, W, 'pBdr'), W, 'top');
    if (color !== undefined) {
      tally.add('run text', l, WordColors.resolve(attr(color, 'val'), attr(color, 'themeColor'), attr(color, 'themeTint'),
        attr(color, 'themeShade'), theme), expected(probe, l, 3));
    }
    if (pShd !== undefined && attr(pShd, 'val') === 'clear') {
      tally.add('paragraph shading', l, WordColors.resolve(attr(pShd, 'fill'), attr(pShd, 'themeFill'),
        attr(pShd, 'themeFillTint'), attr(pShd, 'themeFillShade'), theme), expected(probe, l, 1));
    }
    if (pShd !== undefined && attr(pShd, 'val') === 'pct25') {
      // the pattern's colour, then Word's 25% blend of it over the fill
      const pattern = WordColors.resolve(attr(pShd, 'color'), attr(pShd, 'themeColor'), undefined, undefined, theme);
      const fill = ResolvedColor.ofHex(attr(pShd, 'fill'));
      const blend = ResolvedColor.rgb(Math.round(0.25 * pattern.red + 0.75 * fill.red),
        Math.round(0.25 * pattern.green + 0.75 * fill.green), Math.round(0.25 * pattern.blue + 0.75 * fill.blue));
      tally.add('pattern blend', l, blend, expected(probe, l, 1));
    }
    if (rShd !== undefined) {
      tally.add('run shading', l, WordColors.resolve(attr(rShd, 'fill'), attr(rShd, 'themeFill'), undefined, undefined, theme), expected(probe, l, 1));
    }
    if (top !== undefined) {
      tally.add('top border', l, WordColors.resolve(attr(top, 'color'), attr(top, 'themeColor'), attr(top, 'themeTint'),
        attr(top, 'themeShade'), theme), expected(probe, l, 4).split(' ')[0]);
    }
    if (u !== undefined) {
      tally.add('underline', l, WordColors.resolve(attr(u, 'color'), attr(u, 'themeColor'), undefined, undefined, theme), expected(probe, l, 4).split(' ')[0]);
    }
  }
  check(t, 'P1', tally, 17, DOCX4J.P1);
  // exact but for P1-10, a paragraph shading's themeFillTint whose written value is unusable (FF0000): Word drew
  // F6C5AC, one unit below the computation that gives its three run cases exactly (docx4j CR-034 section 11)
  assert.ok(tally.count(1) <= 1 && tally.count(2) === 0, tally.report());
});

test('P2: w:clrSchemeMapping applies to WordprocessingML theme colours and DrawingML scheme colours alike; every case exact', async (t) => {
  const probe = 'tc-p2-colour-map.docx';
  const theme = await ThemeColors.of(await OpcPackage.load(bytesOf(probe)));
  const tally = new Tally();
  for (const p of elements(xml(probe, 'word/document.xml'), W, 'p')) {
    const txt = text(p);
    if (!txt.startsWith('P2-')) continue;
    const l = txt.substring(0, 5);
    const color = child(child(child(p, W, 'r'), W, 'rPr'), W, 'color');
    const c = color === undefined ? undefined : WordColors.resolve(attr(color, 'val'), attr(color, 'themeColor'), undefined, undefined, theme);
    tally.add('run text', l, c ?? ResolvedColor.BLACK, expected(probe, l, 3)); // automatic text is black
  }
  await wordShapes(probe, tally);
  check(t, 'P2', tally, 13, DOCX4J.P2);
  assert.equal(tally.count(1) + tally.count(2), 0, 'every P2 case exact');
});

test('P5: a slide overriding its master\'s colour map, and a slide keeping it; every case exact', async (t) => {
  const probe = 'tc-p5-map-override.pptx';
  const tally = new Tally();
  for (const { sp, theme, label: l } of await slideShapes(probe)) {
    tally.add('fill', l, DrawingMLColors.resolve(sp.spPr.solidFill, theme), expected(probe, l, 1));
    // the label states no colour: the master's text styles give tx1, through the slide's map
    tally.add('text (tx1)', l, theme.schemeColor('tx1'), expected(probe, l, 3));
  }
  check(t, 'P5', tally, 20, DOCX4J.P5);
  assert.equal(tally.count(1) + tally.count(2), 0, 'every P5 case exact');
});

test('P4: SpreadsheetML theme indexes (lt1, dk1, lt2, dk2 first), tints in Windows\' integer HLS, the indexed palettes; every case exact', async (t) => {
  const tally = new Tally();
  for (const probe of ['tc-p4-spreadsheet.xlsx', 'tc-p4-spreadsheet-palette.xlsx']) {
    const xlsx = await OpcPackage.load(bytesOf(probe));
    const theme = await ThemeColors.of(xlsx);
    const workbook = xlsx.getWorkbookPart();
    const ss = await workbook.stylesPart.readContents();
    const indexed = ss.colors?.indexedColors;
    for (const ws of workbook.worksheetParts) {
      for (const row of (await ws.readContents()).sheetData.row ?? []) {
        for (const c of row.c ?? []) {
          const l = c.is.t.value;
          const fillId = ss.cellXfs.xf[c.s].fillId;
          const fg = ss.fills.fill[fillId].patternFill.fgColor;
          const fam = l.startsWith('X') ? ((fg.tint ?? 0) === 0 ? 'theme' : 'theme + tint') : l.startsWith('Y') ? 'indexed' : 'indexed, custom palette';
          tally.add(fam, l, SpreadsheetColors.resolve(fg, theme, indexed), expected(probe, l, 1));
        }
      }
    }
  }
  check(t, 'P4', tally, 222, DOCX4J.P4);
  assert.equal(tally.count(1) + tally.count(2), 0, 'every P4 case exact');
});

test('Word\'s themeTint and themeShade: the computation within one unit of every distinct value Word wrote in four corpora', async (t) => {
  let exact = 0, near = 0;
  const far = [];
  for (const line of tsv('word-tint-shade.tsv')) {
    const f = line.split('\t');
    const got = WordColors.luminance(ResolvedColor.ofHex(f[0]), Number.parseInt(f[2], 16) / 255.0, f[1] === 'tint');
    if (got.toHex() === f[3]) exact++;
    else if (got.within(ResolvedColor.ofHex(f[3]), 1)) near++;
    else far.push(`${line} -> ${got}`);
  }
  t.diagnostic(`Word tint/shade: ${exact + near + far.length} distinct, exact ${exact}, one unit ${near}`);
  assert.deepEqual(far, []);
  assert.deepEqual([exact, near], [76, 41], 'docx4j\'s answer at 1fffc20a3');
  // with the written value beside it (a current cache), Word's own value comes back exactly
  const theme = ThemeColors.ofTheme(undefined).withScheme(await colorScheme({ accent1: '5B9BD5' }));
  assert.equal(WordColors.resolve('2E74B5', 'accent1', undefined, 'BF', theme).toHex(), '2E74B5');
});

/** An Office 2013 a:clrScheme with some slots replaced, unmarshalled. */
async function colorScheme(slots) {
  const values = { dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5',
    accent4: 'FFC000', accent5: '4472C4', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72', ...slots };
  const body = Object.entries(values).map(([k, v]) => `<a:${k}><a:srgbClr val="${v}"/></a:${k}>`).join('');
  const theme = `<a:theme xmlns:a="${A}" name="t"><a:themeElements><a:clrScheme name="t">${body}</a:clrScheme>`
    + '<a:fontScheme name="t"><a:majorFont><a:latin typeface="A"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    + '<a:minorFont><a:latin typeface="A"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
    + '<a:fmtScheme name="t"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>';
  return (await unmarshalNode(parseXml(theme))).value.themeElements.clrScheme;
}

test('hexBinary (a:sysClr/@lastClr, SpreadsheetML rgb) is byte numbers at run time, though the objects declarations say string', async () => {
  // A dependency defect, asserted as it stands (CR-008 section 7): jsonix-schema-compiler declares hexBinary `string`
  // (its CR-005 table), and jsonix's HexBinary type unmarshals it to an array of bytes. The resolver reads both
  // (`bytesOf`); when either side changes, this fails and the note in CR-008 goes.
  const docx = await OpcPackage.load(bytesOf('tc-p1-precedence.docx'));
  const theme = await docx.getMainDocumentPart().themePart.readContents();
  assert.deepEqual(theme.themeElements.clrScheme.dk1.sysClr.lastClr, [0, 0, 0]);
  const xlsx = await OpcPackage.load(bytesOf('tc-p4-spreadsheet-palette.xlsx'));
  const rgb = (await xlsx.getWorkbookPart().stylesPart.readContents()).colors.indexedColors.rgbColor[0].rgb;
  assert.ok(Array.isArray(rgb) && rgb.length === 4, String(rgb));
  // a string is read as well, as the declarations have it
  const asString = ThemeColors.ofTheme(undefined).withScheme({ ...theme.themeElements.clrScheme, dk1: { sysClr: { val: 'unknown', lastClr: '123456' } } });
  assert.equal(asString.slot('dk1').toHex(), '123456');
});

// ------------------------------------------------------------------ which theme and map a host takes (docx4j ThemeColorsTest)

test('a new docx has the Office 2023 theme', async () => {
  const t = await ThemeColors.of(await WordprocessingMLPackage.createPackage());
  assert.equal(t.schemeColor('accent1').toHex(), '156082');
  assert.equal(t.wordThemeColor('text1').toHex(), '000000');
  assert.equal(t.schemeColor('bg1').toHex(), 'FFFFFF');
  assert.equal(t.wordThemeColor('none'), undefined);
});

test('a package with no theme part takes the default theme its fonts.defaultTheme names (the 2007 setting has the 2013 colours, docx4j CR-034 D2)', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const main = pkg.getMainDocumentPart();
  main.relationshipsPart.removePart(main.themePart.partName);
  main.themePart = undefined;
  assert.equal(main.themePart, undefined);
  assert.equal((await ThemeColors.of(pkg)).schemeColor('accent1').toHex(), '156082');
  pkg.fonts.defaultTheme = '2013';
  assert.equal((await ThemeColors.of(pkg)).schemeColor('accent1').toHex(), '4472C4');
  pkg.fonts.defaultTheme = '2007';
  assert.equal((await ThemeColors.of(pkg)).schemeColor('accent1').toHex(), '4472C4');
});

test('SpreadsheetML theme indexes swap the first pairs', async () => {
  await ThemeColors.loadDefaultTheme();
  const t = ThemeColors.ofTheme(undefined);
  assert.deepEqual([0, 1, 2, 3, 4].map((i) => t.sheetThemeColor(i).toHex()), ['FFFFFF', '000000', 'E8E8E8', '0E2841', '156082']);
  assert.equal(t.sheetThemeColor(12), undefined);
});

test('a chart is drawn in its host\'s theme, under its theme override part and its c:clrMapOvr where it has them', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const host = await ThemeColors.of(pkg);
  const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const chartPart = (clrMapOvr = '') => {
    const part = new ChartPart('/word/charts/chart1.xml');
    part.setXml(`<c:chartSpace xmlns:c="${C}" xmlns:a="${A}">${clrMapOvr}<c:chart><c:plotArea><c:layout/></c:plotArea></c:chart></c:chartSpace>`);
    pkg.getMainDocumentPart().addTargetPart(part);
    return part;
  };
  // no override of its own: the host's
  const plain = await host.forChart(chartPart());
  assert.equal(plain.schemeColor('accent1').toHex(), '156082');
  assert.ok(plain.schemeColor('tx1').equals(host.schemeColor('tx1')));
  // its c:clrMapOvr swaps the light and dark slots
  const mapped = await host.forChart(chartPart('<c:clrMapOvr bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" '
    + 'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'));
  assert.deepEqual([mapped.schemeColor('tx1').toHex(), mapped.schemeColor('bg1').toHex()], ['FFFFFF', '000000']);
  // a theme override part replaces the scheme
  const overridden = chartPart();
  const override = new ThemeOverridePart('/word/theme/themeOverride1.xml');
  override.setXml(`<a:themeOverride xmlns:a="${A}"><a:clrScheme name="o">${['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4',
    'accent5', 'accent6', 'hlink', 'folHlink'].map((k) => `<a:${k}><a:srgbClr val="${k === 'accent1' ? '123456' : '000000'}"/></a:${k}>`).join('')}</a:clrScheme></a:themeOverride>`);
  overridden.addTargetPart(override);
  assert.equal((await host.forChart(overridden)).schemeColor('accent1').toHex(), '123456');
});

test('a slide, a layout and a master take the master\'s theme; a part of a docx its package\'s', async () => {
  const probe = 'tc-p5-map-override.pptx';
  const pptx = await OpcPackage.load(bytesOf(probe));
  const [slide] = pptx.slideParts;
  const layout = slide.slideLayoutPart;
  const master = layout.slideMasterPart;
  for (const part of [slide, layout, master]) assert.equal((await ThemeColors.of(part)).schemeColor('accent1').toHex(), '156082', part.partName.name);
  // a part of a docx: its package's
  const docx = await OpcPackage.load(bytesOf('tc-p2-colour-map.docx'));
  const viaPart = await ThemeColors.of(docx.getMainDocumentPart());
  const viaPackage = await ThemeColors.of(docx);
  assert.ok(viaPart.schemeColor('tx1').equals(viaPackage.schemeColor('tx1')));
});

// ------------------------------------------------------------------ through the PropertyResolver and Font (docx4j ThemeColourOutputTest)

const WNS = `xmlns:w="${W}"`;

/**
 * docx4j's fixture: a new document in the Office 2023 theme (accent1 156082, accent2 E97132, accent4 0F9ED5, accent5
 * A02B93, accent6 4EA72E), whose stale values are Office 2007's accent1, 4F81BD. Each paragraph's text names its case.
 */
async function outputDoc() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const main = pkg.getMainDocumentPart();
  const styles = (await main.styleDefinitionsPart.getContents()).style;
  const style = async (id, pr) => (await unmarshalNode(parseXml(`<w:style ${WNS} w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/>`
    + `<w:basedOn w:val="Normal"/>${pr}</w:style>`))).value;
  styles.push(await style('StaleColour', '<w:rPr><w:color w:val="4F81BD" w:themeColor="accent1"/></w:rPr>'));
  styles.push(await style('ThemeOnly', '<w:rPr><w:color w:themeColor="accent2"/></w:rPr>'));
  styles.push(await style('CurrentCache', '<w:rPr><w:color w:val="156082" w:themeColor="accent1"/></w:rPr>'));
  main.setXml(`<w:document ${WNS}><w:body>`
    + '<w:p><w:pPr><w:pStyle w:val="StaleColour"/></w:pPr><w:r><w:t>STALE</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:pStyle w:val="StaleColour"/></w:pPr><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>RED</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:pStyle w:val="ThemeOnly"/></w:pPr><w:r><w:t>THEMEONLY</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:pStyle w:val="CurrentCache"/></w:pPr><w:r><w:t>CURRENT</w:t></w:r></w:p>'
    + '<w:p><w:r><w:rPr><w:color w:val="auto" w:themeColor="accent5"/></w:rPr><w:t>AUTO</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="auto" w:themeFill="accent2"/></w:pPr><w:r><w:t>SHADEDAUTO</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>UNBORDERED</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="4F81BD" w:themeColor="accent6"/></w:pBdr></w:pPr><w:r><w:t>BORDERED</w:t></w:r></w:p>'
    + '<w:p><w:r><w:rPr><w:u w:val="single" w:color="4F81BD" w:themeColor="accent4"/></w:rPr><w:t>UNDERLINED</w:t></w:r></w:p>'
    + '</w:body></w:document>');
  pkg.refreshPropertyResolver();
  return { pkg, body: await pkg.getBody() };
}

const runOf = (p) => { const r = p.p.content[0]; return r.value ?? r; };

test('PropertyResolver: the colour Word draws, the theme reference over the value beside it (docx4j resolverGivesTheColourWordDraws)', async () => {
  const { pkg, body } = await outputDoc();
  const pr = pkg.propertyResolver;
  const ps = body.paragraphs;
  const rPr = (i) => pr.getEffectiveRPr(runOf(ps[i]).rPr, ps[i].p.pPr);
  assert.deepEqual([0, 1, 2, 3, 4].map((i) => pr.resolveColor(rPr(i).color).toHex()), ['156082', 'FF0000', 'E97132', '156082', 'A02B93']);
  const shaded = pr.getEffectivePPr(ps[5].p.pPr);
  assert.equal(pr.resolveFill(shaded.shd).toHex(), 'E97132');
  assert.equal(pr.resolvePatternColor(shaded.shd), undefined, 'automatic');
  assert.equal(pr.resolveBorderColor(pr.getEffectivePPr(ps[7].p.pPr).pBdr.top).toHex(), '4EA72E');
  assert.equal(pr.resolveUnderlineColor(rPr(8).u).toHex(), '0F9ED5');
  assert.equal(pr.resolveColor(rPr(6).color), undefined, 'no colour anywhere: automatic');
  assert.equal(pr.getThemeColors(), pr.getThemeColors(), 'made once');
  // a run's own value, stated without a theme reference, replaces the style's reference with it; a style stating only
  // a theme colour is inherited (CR-008 phase 1, docx4j aDirectValueReplacesTheStylesThemeReference and the next)
  assert.deepEqual([rPr(1).color.val, rPr(1).color.themeColor], ['FF0000', undefined]);
  assert.equal(rPr(2).color.themeColor, 'accent2');
  // refresh() reads the theme again
  const before = pr.getThemeColors();
  pkg.refreshPropertyResolver();
  assert.notEqual(pr.getThemeColors(), before);
});

test('Font.color reads the colour Word draws in the effective view, and w:val as written in the direct view (CR-008 phase 2)', async () => {
  const { body } = await outputDoc();
  const ps = body.paragraphs;
  assert.deepEqual([0, 1, 2, 3, 4, 6].map((i) => ps[i].font.color), ['#156082', '#FF0000', '#E97132', '#156082', '#A02B93', '#000000']);
  assert.deepEqual([0, 1, 2, 3, 4, 6].map((i) => ps[i].getRange('Whole').font.color), ['#156082', '#FF0000', '#E97132', '#156082', '#A02B93', '#000000']);
  // the direct view: the run's own w:color as written, a theme reference not read
  assert.deepEqual([0, 1, 4].map((i) => ps[i].getFont({ direct: true }).color), ['', '#FF0000', 'auto']);
  // across paragraphs whose runs' colours differ, '' (MixedFont); where they agree, the colour drawn
  assert.equal(ps[0].getRange('Whole').expandTo(ps[1].getRange('Whole')).font.color, '');
  assert.equal(ps[0].getRange('Whole').expandTo(ps[3].getRange('Whole')).font.color, '');
  assert.equal(ps[3].getRange('Whole').expandTo(ps[3].getRange('Whole')).font.color, '#156082');
});

test('Font.color through the content API is what Word drew for every P1 and P2 paragraph\'s text, w:clrSchemeMapping included', async () => {
  for (const probe of ['tc-p1-precedence.docx', 'tc-p2-colour-map.docx']) {
    const pkg = await OpcPackage.load(bytesOf(probe));
    const body = await pkg.getBody();
    let n = 0;
    for (const p of body.paragraphs) {
      const l = p.text.substring(0, 5);
      const want = /^P[12]-\d\d$/.test(l) ? expected(probe, l, 3) : undefined;
      if (want === undefined || want === '') continue;
      assert.equal(p.font.color, `#${want}`, `${probe} ${l}`);
      n++;
    }
    // P1's 17 paragraphs, the shaded and bordered ones' text automatic; P2's six runs (docx4j's 21 count the shapes' anchor and labels too)
    assert.equal(n, probe.startsWith('tc-p1') ? 17 : 6, probe);
  }
});

test('a package with no theme part reads its theme colours from the default theme it names', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const main = pkg.getMainDocumentPart();
  main.relationshipsPart.removePart(main.themePart.partName);
  main.themePart = undefined;
  pkg.fonts.defaultTheme = '2013';
  const pr = await pkg.refresh();
  assert.equal(pr.getThemeColors().schemeColor('accent1').toHex(), '4472C4');
  assert.equal(pr.resolveColor({ val: '365F91', themeColor: 'accent1', themeShade: 'BF' }).toHex(), '2F5496');
});
