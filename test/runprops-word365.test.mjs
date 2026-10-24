// CR-015: run properties read as Word 365 reads them (on/off, enumerations, colours). The probes are docx4j's
// (test/runprops-word365/, copied unedited at the commit its manifest.json names): Jason's seven rounds, each probe
// with Word 365's save of it. Word writes back only values it reads one way (an on value bare, an off value 0 or
// dropped, a colour as upper-case hex, a repaired element dropped), so Word's save is the oracle: each marker must
// read in the probe as it reads in the save, through the effective resolver, the text model and the Office JS Font,
// and after a round trip here. Where Word's own save changes what Word drew (H7, U4, M6) its PDF is stated instead.
// docx4j's RunPropertiesWord365Test (its HTML and PDF halves have no counterpart here), OnOffAttributesTest,
// ColourNormalFormTest and ColourNormalFormOptOutTest (CR-015 phase 2b: the objects facade's save normaliser, its
// CR-013 phase C); and check 60, Office JS's own Font readings over the probes (D4).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync } from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { textOf } from '@docx4j/generated-objects-ts/builders/wml';
import * as wmlModule from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import {
  WordprocessingMLPackage, Texts, WordColors, unmarshalString, setNormaliseOnSave,
} from '../dist/index.mjs';

/** A `BooleanDefaultTrue`'s value: absent `w:val` means true (docx4j `isVal()`). */
const isTrue = (value) => value !== undefined && value !== null && value.val !== false;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const dir = new URL('./runprops-word365/', import.meta.url);
const bytesOf = (name) => readFileSync(new URL(name, dir));

const FILES = [
  'onoff-valid', 'onoff-empty', 'onoff-space', 'onoff-TRUE', 'onoff-yes',
  'props-b-case', 'props-b-styles', 'props-toggles-1', 'props-toggles-2',
  'props-u-valid', 'props-u-case', 'props-u-other', 'props-enum-missing', 'props-enum-other',
  'colour-names-highlight', 'colour-names-css', 'colour-case', 'colour-hex',
  'colour-table-grey', 'colour-table-other',
  'props-u-styles', 'colour-styles', 'props-enum-styles', 'colour-elsewhere',
  'props-u-onoff-styles', 'colour-auto-styles', 'highlight-char-styles', 'colour-elsewhere-styles',
];

/**
 * Where Word's own save changes what Word drew (docx4j CR-060 section 13, item 5), Word's reading from its PDF: a
 * direct highlight over a character style's, drawn but dropped by the save; an underline colour "auto" over a style's
 * red, drawn black but dropped.
 */
const WORD_DREW = new Map([
  ['props-enum-styles M6_HIGHLIGHT_GREEN', { highlight: 'green' }],
  ['highlight-char-styles H7_CHAR_STYLE_DIRECT_GREEN', { highlight: 'green' }],
  ['colour-elsewhere-styles U4_UNDERLINE_AUTO', { 'u-color': 'auto' }],
]);

/** The on/off run properties the probes test (CR-015 section 2, rule 1). */
const ON_OFF = ['b', 'i', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'vanish'];

const unwrap = (x) => (x !== null && typeof x === 'object' && x.TYPE_NAME === undefined && 'value' in x ? x.value : x);
const isA = (x, type) => unwrap(x)?.TYPE_NAME === `org_docx4j_wml.${type}`;
const runsOf = (p) => (p.content ?? []).filter((x) => isA(x, 'R')).map(unwrap);

async function load(name) {
  return WordprocessingMLPackage.load(bytesOf(name));
}

/** A probe paragraph's marker run, in the probe and in Word's save: the middle of three runs, its text the marker. */
class Probe {
  static async of(name) {
    return new Probe(name, await load(`${name}.docx`), await load(`${name}-word365.docx`)).ready();
  }

  constructor(name, pkg, saved) {
    this.name = name;
    this.pkg = pkg;
    this.saved = saved;
    this.markers = [];
  }

  async ready() {
    const body = await this.pkg.getBody(), savedBody = await this.saved.getBody();
    const ps = body.paragraphs, savedPs = savedBody.paragraphs;
    assert.equal(ps.length, savedPs.length, `${this.name}: paragraphs in the save`);
    for (let i = 1; i < ps.length; i++) { // the first paragraph is the probe's title
      const runs = runsOf(ps[i].p);
      if (runs.length !== 3) continue;
      const text = textOf(runs[1]);
      const savedRun = runsOf(savedPs[i].p).filter((r) => textOf(r).includes(text)).at(-1);
      assert.ok(savedRun, `${this.name}: ${text} as one run in the save`);
      this.markers.push({ file: this.name, index: i, text, run: runs[1], p: ps[i].p, para: ps[i], savedRun, savedP: savedPs[i].p, savedPara: savedPs[i], key: `${this.name} ${text}` });
    }
    assert.ok(this.markers.length > 0, `${this.name}: markers`);
    this.resolver = await this.pkg.getPropertyResolver();
    this.savedResolver = await this.saved.getPropertyResolver();
    return this;
  }

  marker(text) {
    const mk = this.markers.find((m) => m.text === text);
    assert.ok(mk, `${this.name}: no marker ${text}`);
    return mk;
  }
}

const probes = new Map();
async function probe(name) {
  if (!probes.has(name)) probes.set(name, await Probe.of(name));
  return probes.get(name);
}

/** What a run's effective properties say, as far as the probes test them (docx4j's `reading`). */
function reading(resolver, p, r) {
  const eff = resolver.getEffectiveRPr(r.rPr, p.pPr);
  const m = {};
  for (const k of ON_OFF) m[k] = isTrue(eff?.[k]) ? 'on' : 'off';
  const u = eff?.u;
  m.u = u === undefined || u.val === undefined || u.val === 'none' ? 'none' : u.val;
  const h = eff?.highlight;
  m.highlight = h === undefined || h.val === undefined || h.val === 'none' ? 'none' : h.val;
  m.vertAlign = eff?.vertAlign?.val ?? 'none';
  const em = eff?.em;
  m.em = em === undefined || em.val === undefined || em.val === 'none' ? 'none' : em.val;
  m.color = resolver.resolveColor(eff?.color)?.toHex() ?? 'auto';
  m['u-color'] = m.u === 'none' ? '-' : resolver.resolveUnderlineColor(u)?.toHex() ?? 'auto';
  m.fill = eff?.shd === undefined ? 'none' : resolver.resolveFill(eff.shd)?.toHex() ?? 'none';
  const bdr = eff?.bdr;
  const bordered = bdr !== undefined && bdr.val !== undefined && bdr.val !== 'none' && bdr.val !== 'nil';
  m.border = bordered ? resolver.resolveBorderColor(bdr)?.toHex() ?? '000000' : 'none';
  return m;
}

/** Word's reading of a marker: its save's, or its PDF's where the save changes what Word drew. */
function wordReading(pr, mk) {
  return { ...reading(pr.savedResolver, mk.savedP, mk.savedRun), ...WORD_DREW.get(mk.key) };
}

function compare(what, mk, ours, word, diffs) {
  for (const k of Object.keys(word)) {
    if (word[k] !== ours[k]) diffs.push(`${mk.key} ${what} ${k}: core-ts ${ours[k]}, Word ${word[k]}`);
  }
}

function assertNone(diffs) {
  assert.ok(diffs.length === 0, `${diffs.length} readings differ from Word's:\n${diffs.join('\n')}`);
}

test('the probes are docx4j\'s, unedited (manifest.json)', () => {
  const manifest = JSON.parse(readFileSync(new URL('manifest.json', dir), 'utf8'));
  assert.equal(manifest.docx4j.commit, '5163c29862ed6c6b5f6a707fc30c98e009e3c05d');
  const names = readdirSync(dir).filter((n) => n !== 'manifest.json').sort();
  assert.deepEqual(names, manifest.files.map((f) => f.name));
  for (const f of manifest.files) {
    assert.equal(createHash('sha1').update(bytesOf(f.name)).digest('hex'), f.sha1, f.name);
  }
  for (const f of FILES) assert.ok(names.includes(`${f}.docx`) && names.includes(`${f}-word365.docx`), f);
});

test('every probe and Word\'s save of it loads and unmarshals (the facade reads refused values as Word does)', async () => {
  for (const name of readdirSync(dir).filter((n) => n.endsWith('.docx'))) {
    const pkg = await load(name);
    await pkg.unmarshalAll();
  }
});

test('the resolver\'s effective properties, for every marker, as Word reads them', async () => {
  const diffs = [];
  let n = 0;
  for (const f of FILES) {
    const pr = await probe(f);
    for (const mk of pr.markers) {
      compare('effective', mk, reading(pr.resolver, mk.p, mk.run), wordReading(pr, mk), diffs);
      n++;
    }
  }
  assertNone(diffs);
  assert.equal(n, 376, 'markers');
});

/** The text model's runs, by the marker each holds. */
async function textRuns(pkg) {
  return (await Texts.of(pkg)).blocks.flatMap((b) => b.runs ?? []);
}

test('the text model\'s bold, italic and struck through, for every marker, as Word reads them', async () => {
  const diffs = [];
  for (const f of FILES) {
    const pr = await probe(f);
    const runs = await textRuns(pr.pkg), savedRuns = await textRuns(pr.saved);
    for (const mk of pr.markers) {
      const t = runs.find((r) => r.text?.includes(mk.text)), w = savedRuns.find((r) => r.text?.includes(mk.text));
      if (t === undefined && w === undefined) continue; // hidden (w:vanish): no text in either
      if (t === undefined || w === undefined) {
        diffs.push(`${mk.key} text model: core-ts ${t === undefined ? 'hidden' : 'shown'}, Word ${w === undefined ? 'hidden' : 'shown'}`);
        continue;
      }
      for (const k of ['bold', 'italic', 'strike']) if (t[k] !== w[k]) diffs.push(`${mk.key} text model ${k}: core-ts ${t[k]}, Word ${w[k]}`);
    }
  }
  assertNone(diffs);
});

/** Office JS's Font of the marker's text (`Paragraph.search`), effective; `hidden` where Word's view has no such text. */
function fontReading(para, text, direct = false) {
  const hit = para.search(text, { matchCase: true }).at(-1);
  if (hit === undefined) return { hidden: 'hidden' };
  const font = direct ? hit.getFont({ direct: true }) : hit.font;
  const out = {};
  for (const k of ['bold', 'italic', 'underline', 'strikeThrough', 'doubleStrikeThrough', 'superscript', 'subscript', 'color', 'highlightColor']) {
    out[k] = String(font[k]);
  }
  return out;
}

/**
 * Word's Font reading of a marker: the save's, with what its PDF shows where the save changes it (a green highlight,
 * in the objects package's `highlightHexValue`; what Office JS itself reports for Word's highlight names is check
 * 60's to measure, CR-015 D4).
 */
const FONT_DREW = new Map([
  ['props-enum-styles M6_HIGHLIGHT_GREEN', { highlightColor: '#008000' }],
  ['highlight-char-styles H7_CHAR_STYLE_DIRECT_GREEN', { highlightColor: '#008000' }],
]);

test('Office JS\'s Font (effective) of every marker, as Word reads it', async () => {
  const diffs = [];
  for (const f of FILES) {
    const pr = await probe(f);
    for (const mk of pr.markers) {
      const word = { ...fontReading(mk.savedPara, mk.text), ...FONT_DREW.get(mk.key) };
      compare('Font', mk, fontReading(mk.para, mk.text), word, diffs);
    }
  }
  assertNone(diffs);
});

// The direct view differs from Word's save only where the save rewrites direct markup without changing what is drawn,
// or changes what is drawn (H7, M6): an automatic colour stated where Word drops it, and Word dropping a direct
// w:u single that repeats the style's (the objects package's three known, CR-015 section 6).
const DIRECT_DIFFERS = [
  'colour-case K1_auto direct Font color: core-ts auto, Word ',
  'props-enum-styles M6_HIGHLIGHT_GREEN direct Font highlightColor: core-ts #008000, Word null',
  'highlight-char-styles H7_CHAR_STYLE_DIRECT_GREEN direct Font highlightColor: core-ts #008000, Word null',
  'colour-elsewhere-styles U2_UNDERLINE_F00 direct Font underline: core-ts Single, Word None',
  'colour-elsewhere-styles U3_UNDERLINE_NAVY direct Font underline: core-ts Single, Word None',
  'colour-elsewhere-styles U4_UNDERLINE_AUTO direct Font underline: core-ts Single, Word None',
];

test('Office JS\'s Font (direct) of every marker, as Word\'s save writes it, but where the save rewrites it', async () => {
  const diffs = [];
  for (const f of FILES) {
    const pr = await probe(f);
    for (const mk of pr.markers) compare('direct Font', mk, fontReading(mk.para, mk.text, true), fontReading(mk.savedPara, mk.text, true), diffs);
  }
  assert.deepEqual(diffs, DIRECT_DIFFERS);
});

test('a w:u stating only a colour is no underline (CR-015 section 2, rule 2)', async () => {
  const pr = await probe('props-u-valid');
  const mk = pr.marker('U7_COLOR_ONLY');
  assert.equal(reading(pr.resolver, mk.p, mk.run).u, 'none');
  assert.equal(mk.para.search('U7_COLOR_ONLY')[0].font.underline, 'None');
  const red = pr.marker('U8_SINGLE_RED');
  assert.equal(reading(pr.resolver, red.p, red.run).u, 'single');
  assert.equal(red.para.search('U8_SINGLE_RED')[0].font.underline, 'Single');
});

test('the markers whose save changes what Word drew (H7, M6, U4), as its PDF shows', async () => {
  for (const [file, text] of [['props-enum-styles', 'M6_HIGHLIGHT_GREEN'], ['highlight-char-styles', 'H7_CHAR_STYLE_DIRECT_GREEN']]) {
    const pr = await probe(file);
    const mk = pr.marker(text);
    assert.equal(reading(pr.resolver, mk.p, mk.run).highlight, 'green', mk.key);
    assert.equal(mk.para.search(text)[0].font.highlightColor, '#008000', mk.key);
  }
  const pr = await probe('colour-elsewhere-styles');
  const u4 = pr.marker('U4_UNDERLINE_AUTO');
  assert.equal(reading(pr.resolver, u4.p, u4.run)['u-color'], 'auto');
});

test('an invalid highlight, vertAlign or em over a style leaves the style\'s; a character style\'s highlight is not drawn', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const styles = pkg.getMainDocumentPart().styleDefinitionsPart;
  const contents = await styles.getContents();
  const style = async (xml) => (await unmarshalString(`<w:styles xmlns:w="${W}">${xml}</w:styles>`)).value.style[0];
  contents.style.push(await style('<w:style w:type="character" w:customStyle="1" w:styleId="Marked"><w:name w:val="Marked"/><w:rPr><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/><w:em w:val="dot"/></w:rPr></w:style>'));
  contents.style.push(await style('<w:style w:type="paragraph" w:customStyle="1" w:styleId="ParaMarked"><w:name w:val="ParaMarked"/><w:rPr><w:highlight w:val="yellow"/></w:rPr></w:style>'));
  const resolver = await pkg.getPropertyResolver();
  resolver.refresh();
  const direct = ['<w:highlight w:val="YELLOW"/><w:vertAlign w:val="Superscript"/><w:em w:val="DOT"/>',
    '<w:highlight w:val="bogus"/><w:vertAlign w:val="bogus"/><w:em w:val="bogus"/>',
    '<w:highlight/><w:vertAlign/><w:em/>'];
  const pOf = async (xml) => (await unmarshalString(`<w:document xmlns:w="${W}"><w:body>${xml}</w:body></w:document>`)).value.body.content.map(unwrap)[0];
  for (const d of direct) {
    const p = await pOf(`<w:p><w:pPr><w:pStyle w:val="ParaMarked"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Marked"/>${d}</w:rPr><w:t>x</w:t></w:r></w:p>`);
    const m = reading(resolver, p, runsOf(p)[0]);
    assert.equal(m.highlight, 'yellow', `${d}: the paragraph style's`);
    assert.equal(m.vertAlign, 'superscript', d);
    assert.equal(m.em, 'dot', d);
    const alone = await pOf(`<w:p><w:r><w:rPr><w:rStyle w:val="Marked"/>${d}</w:rPr><w:t>x</w:t></w:r></w:p>`);
    assert.equal(reading(resolver, alone, runsOf(alone)[0]).highlight, 'none', `${d}: the character style's isn't drawn`);
  }
});

function documentXml(bytes) {
  return strFromU8(unzipSync(bytes)['word/document.xml']);
}

/** ST_HighlightColor's values, for the round trip's check. */
const HIGHLIGHT_VALUES = new Set(['black', 'blue', 'cyan', 'green', 'magenta', 'red', 'yellow', 'white', 'darkBlue', 'darkCyan',
  'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'none']);


/**
 * A round trip keeps Word's reading of every marker (a save never changes meaning), writes on/off as true or false
 * (D2), leaves out a highlight, vertAlign or em Word repairs away (docx4j CR-060's D3), and writes a colour Word reads
 * as upper-case hex (CR-061).
 */
test('a round trip keeps Word\'s reading of every marker, writes on/off as true or false and leaves out what Word repairs', async () => {
  const diffs = [];
  for (const f of FILES) {
    const pr = await probe(f);
    const bytes = await pr.pkg.save();
    const again = await new Probe(f, await WordprocessingMLPackage.load(bytes), pr.saved).ready();
    for (const mk of again.markers) compare('after a round trip', mk, reading(again.resolver, mk.p, mk.run), wordReading(pr, mk), diffs);
    const xml = documentXml(bytes);
    for (const m of xml.matchAll(new RegExp(`<w:(${ON_OFF.join('|')})\\b[^>]*?w:val="([^"]*)"`, 'g'))) {
      if (m[2] !== 'true' && m[2] !== 'false') diffs.push(`${f} saved w:${m[1]} w:val="${m[2]}" (D2: true or false)`);
    }
    for (const m of xml.matchAll(/<w:(highlight|vertAlign|em)\s*\/>/g)) diffs.push(`${f} saved a bare w:${m[1]} (D3: left out, as Word repairs)`);
    for (const m of xml.matchAll(/<w:highlight w:val="([^"]*)"/g)) if (!HIGHLIGHT_VALUES.has(m[1])) diffs.push(`${f} saved w:highlight w:val="${m[1]}" (D3)`);
  }
  const other = await load('props-enum-other.docx');
  await other.getBody();
  if (!runXml(documentXml(await other.save()), 'N6_COLOR_NAME').includes('<w:color w:val="FF0000"/>')) diffs.push('props-enum-other: color "red" not written FF0000 (CR-061)');
  assertNone(diffs);
});

// ---------------------------------------------------------------- D8: the record

test('each value read by Word\'s rule is a warning on its part (CR-015 D8)', async () => {
  const warningsOf = async (name) => {
    const pkg = await load(`${name}.docx`);
    await pkg.getMainDocumentPart().getContents();
    return pkg.getMainDocumentPart().warnings.filter((w) => w.code.startsWith('value.kept.'));
  };
  const yes = await warningsOf('onoff-yes');
  assert.ok(yes.length > 0);
  assert.ok(yes.every((w) => w.code === 'value.kept.invalid' && w.element === 'w:b' && w.attribute === 'w:val' && w.text === 'yes' && w.value === true), JSON.stringify(yes[0]));
  assert.match(yes[0].message, /^\/word\/document\.xml: w:b\/@w:val="yes" \(line \d+, column \d+\) is not a value its type allows; read as on, as Word 365 reads it$/);
  const cased = await warningsOf('props-b-case');
  assert.ok(cased.some((w) => w.code === 'value.kept.exact' && w.text === ' false' && w.value === true), 'a padded form is kept.exact');
  assert.ok(cased.some((w) => w.code === 'value.kept.invalid' && w.text === 'FALSE'));
  const u = await warningsOf('props-u-onoff-styles');
  assert.ok(u.some((w) => w.element === 'w:u' && w.text === 'true' && w.value === 'none'), 'w:u\'s on/off words read none');
  assert.ok(u.some((w) => w.element === 'w:u' && w.text === 'TRUE' && w.value === undefined), 'another invalid value is dropped');
  // A part with none has no such warning, and an untouched part none at all.
  assert.deepEqual(await warningsOf('onoff-valid'), []);
  const pkg = await load('onoff-yes.docx');
  assert.equal(pkg.warnings.length, 0, 'nothing is read until a part is');
  await pkg.getMainDocumentPart().getContents();
  assert.equal(pkg.warnings.length, yes.length, 'OpcPackage.warnings gathers them');
});

test('a private read records its warnings too, and leaves the part untouched', async () => {
  const pkg = await load('props-enum-styles.docx');
  const styles = pkg.getMainDocumentPart().styleDefinitionsPart;
  await pkg.getPropertyResolver();
  assert.equal(styles.isUntouched, true);
  const doc = pkg.getMainDocumentPart();
  await doc.readContents();
  assert.ok(doc.warnings.some((w) => w.code === 'value.kept.invalid'));
  assert.equal(doc.isUntouched, true);
});

// ---------------------------------------------------------------- docx4j's OnOffAttributesTest

/** A value as written, and Word's reading of it. */
const VALUES = [
  ['false', false], ['off', false], ['0', false],
  ['true', true], ['on', true], ['1', true],
  ['', true], [' ', true], [' false', true], ['false ', true], ['FALSE', true], ['Off', true],
  ['00', true], ['no', true], ['yes', true],
];

/** The wml mapping: every attribute property typed Boolean, by owner type. */
function booleanAttributes() {
  const mod = wmlModule.org_docx4j_wml ?? wmlModule.default ?? wmlModule;
  const out = [];
  for (const t of mod.typeInfos) {
    for (const p of t.propertyInfos ?? []) if (p.type === 'attribute' && p.typeInfo === 'Boolean') out.push({ owner: t.localName, name: p.name, local: p.attributeName?.localPart ?? p.attributeName ?? p.name });
  }
  return { mod, out };
}

/** A path of element names from a global wml element down to an element of `owner`'s type, by breadth-first search. */
function pathTo(mod, owner) {
  const types = new Map(mod.typeInfos.map((t) => [t.localName, t]));
  const local = (ti) => (typeof ti === 'string' && ti.startsWith('.') ? ti.substring(1) : undefined);
  const children = (name) => {
    const out = [];
    for (let t = types.get(name); t !== undefined; t = types.get(local(t.baseTypeInfo))) {
      for (const p of t.propertyInfos ?? []) {
        const kind = p.type ?? 'element';
        if (p.wrapperElementName !== undefined) continue;
        const entries = kind === 'element' ? [{ elementName: p.elementName ?? p.name, typeInfo: p.typeInfo }]
          : kind === 'elements' || kind === 'elementRefs' ? p.elementTypeInfos ?? []
            : kind === 'elementRef' ? [{ elementName: p.elementName ?? p.name, typeInfo: p.typeInfo }] : [];
        for (const e of entries) {
          const en = typeof e.elementName === 'string' ? e.elementName : e.elementName?.namespaceURI === W ? e.elementName.localPart : undefined;
          const ti = local(e.typeInfo);
          if (en !== undefined && ti !== undefined) out.push([en, ti]);
        }
      }
    }
    return out;
  };
  const queue = [];
  const seen = new Set();
  for (const e of mod.elementInfos) {
    if (e.scope !== undefined) continue;
    const en = typeof e.elementName === 'string' ? e.elementName : e.elementName.localPart;
    const ti = local(e.typeInfo);
    if (ti !== undefined && !seen.has(ti)) { seen.add(ti); queue.push([ti, [en]]); }
  }
  while (queue.length > 0) {
    const [type, path] = queue.shift();
    if (type === owner) return path;
    for (const [en, ti] of children(type)) if (!seen.has(ti)) { seen.add(ti); queue.push([ti, [...path, en]]); }
  }
  return undefined;
}

function findTyped(node, typeName, seen = new Set()) {
  if (node === null || typeof node !== 'object' || seen.has(node)) return undefined;
  seen.add(node);
  if (node.TYPE_NAME === typeName) return node;
  for (const [k, v] of Object.entries(node)) {
    if (k === 'PARENT') continue;
    const hit = findTyped(v, typeName, seen);
    if (hit !== undefined) return hit;
  }
  return undefined;
}

test('every on/off attribute in the wml mapping reads as Word 365 reads it (docx4j OnOffAttributesTest)', async () => {
  const { mod, out } = booleanAttributes();
  // docx4j's 46, and w:activeWritingStyle/@w:checkStyle and w:readModeInkLockDown/@w:actualPg, which docx4j binds as a
  // primitive boolean its adapter can't reach (its CR-060 section 9) and jsonix as Boolean like the rest
  assert.equal(out.length, 48);
  const diffs = [], unreachable = [];
  for (const a of out) {
    const path = pathTo(mod, a.owner);
    if (path === undefined) { unreachable.push(a.owner); continue; }
    for (const [text, word] of VALUES) {
      const open = path.map((n, i) => `<w:${n}${i === 0 ? ` xmlns:w="${W}"` : ''}${i === path.length - 1 ? ` w:${a.local}="${text}"` : ''}>`).join('');
      const close = [...path].reverse().map((n) => `</w:${n}>`).join('');
      const element = findTyped((await unmarshalString(open + close)).value, `org_docx4j_wml.${a.owner}`);
      assert.ok(element, `${a.owner} in ${open}`);
      if (element[a.name] !== word) diffs.push(`${a.owner}.${a.name} for ${JSON.stringify(text)}: ${element[a.name]}, Word ${word}`);
    }
  }
  assertNone(diffs);
  // declared, but no element of the mapping has the type, so no document reaches it (docx4j's test unmarshals it by
  // declared type, which jsonix has no counterpart of)
  assert.deepEqual(unreachable, ['BooleanDefaultFalse']);
});

// ---------------------------------------------------------------- docx4j's ColourNormalFormTest, as it stands

const COLOUR_PROBES = ['colour-case', 'colour-hex', 'colour-names-css', 'colour-names-highlight', 'colour-table-grey',
  'colour-table-other', 'props-enum-other', 'props-enum-missing', 'props-u-valid', 'colour-styles', 'colour-elsewhere',
  'colour-auto-styles', 'colour-elsewhere-styles', 'colour-pattern', 'colour-background-red', 'colour-background-navy',
  'colour-background-f00', 'colour-background-hex'];

/** The colour values in a part's XML that aren't in the schema's form, and any w:color without a w:val (docx4j's `outOfForm`). */
function outOfForm(xml) {
  const bad = [];
  for (const re of [/<w:color\b[^>]*?\bw:val="([^"]*)"/g, /<w:u\b[^>]*?\bw:color="([^"]*)"/g, /<w:shd\b[^>]*?\bw:fill="([^"]*)"/g,
    /<w:shd\b[^>]*?\bw:color="([^"]*)"/g, /<w:(?:top|left|bottom|right|start|end|between|bar|bdr|insideH|insideV|tl2br|tr2bl)\b[^>]*?\bw:color="([^"]*)"/g,
    /<w:background\b[^>]*?\bw:color="([^"]*)"/g]) {
    for (const m of xml.matchAll(re)) if (!/^(auto|[0-9A-F]{6})$/.test(m[1])) bad.push(m[0]);
  }
  for (const m of xml.matchAll(/<w:color(?:\s+(?!w:val=)[\w:]+="[^"]*")*\s*\/>/g)) bad.push(`${m[0]} (no w:val)`);
  return bad;
}

/** A part of the package as a save writes it, every XML part read first, as a caller who edited the file would have. */
async function saved(pkg, part) {
  await pkg.unmarshalAll();
  const file = unzipSync(await pkg.save())[part];
  assert.ok(file, `no ${part}`);
  return strFromU8(file);
}

/** A marker's run in a part's XML. */
function runXml(xml, marker) {
  const i = xml.indexOf(`>${marker}<`);
  assert.ok(i > 0, marker);
  const start = Math.max(xml.lastIndexOf('<w:r>', i), xml.lastIndexOf('<w:r ', i));
  return xml.substring(start, xml.indexOf('</w:r>', i));
}

test('every saved colour is in the schema\'s form (docx4j ColourNormalFormTest)', async () => {
  const bad = [];
  for (const f of COLOUR_PROBES) {
    const pkg = await load(`${f}.docx`);
    await pkg.unmarshalAll();
    const files = unzipSync(await pkg.save());
    for (const part of ['word/document.xml', 'word/styles.xml']) {
      if (files[part] === undefined) continue;
      for (const b of outOfForm(strFromU8(files[part]))) bad.push(`${f} ${part}: ${b}`);
    }
  }
  assert.ok(bad.length === 0, `${bad.length} colours out of the schema's form:\n${bad.join('\n')}`);
});

test('each colour rule as Word saves it (docx4j ColourNormalFormTest)', async () => {
  const has = (xml, marker, text) => assert.ok(runXml(xml, marker).includes(text), `${marker}: ${runXml(xml, marker)}`);
  const lacks = (xml, marker, text) => assert.ok(!runXml(xml, marker).includes(text), `${marker}: ${runXml(xml, marker)}`);
  const caseXml = await saved(await load('colour-case.docx'), 'word/document.xml');
  has(caseXml, 'K4_Red', 'w:val="FF0000"');
  has(caseXml, 'K1_auto', 'w:val="auto"');
  const hex = await saved(await load('colour-hex.docx'), 'word/document.xml');
  has(hex, 'X1_ff0000', 'w:val="FF0000"');
  has(hex, 'X10_FF0000', 'w:val="FF0000"');
  lacks(hex, 'X3_F00', '<w:color'); // rejected: absent
  const styles = await saved(await load('colour-styles.docx'), 'word/document.xml');
  lacks(styles, 'R2_F00', '<w:color'); // the style's shows
  has(styles, 'R9_AUTO', 'w:val="auto"');
  const elsewhere = await saved(await load('colour-elsewhere.docx'), 'word/document.xml');
  has(elsewhere, 'U1_UNDERLINE_RED', 'w:color="FF0000"');
  has(elsewhere, 'U3_UNDERLINE_F00', '<w:u w:val="single"/>');
  has(elsewhere, 'S2_SHADING_AQUA', 'w:fill="00FFFF"');
  has(elsewhere, 'S3_SHADING_F00', 'w:fill="auto"');
  has(elsewhere, 'B1_BORDER_RED', 'w:color="FF0000"');
  has(elsewhere, 'B2_BORDER_F00', 'w:color="auto"');
  // round 7 (docx4j CR-061 phase 3): a pattern colour as a fill's, over a style too; a page colour named
  const pattern = await saved(await load('colour-pattern.docx'), 'word/document.xml');
  has(pattern, 'P2_PATTERN_NAVY', 'w:color="000080"');
  has(pattern, 'P3_PATTERN_F00', 'w:color="auto"');
  has(pattern, 'P7_STYLE_F00', 'w:color="auto"');
  const page = await saved(await load('colour-background-red.docx'), 'word/document.xml');
  assert.ok(page.includes('<w:background w:color="FF0000"/>'), page);
  // Word drops a w:background whose colour it rejects; docx4j writes "auto" (no page colour either way), and so does this
  const rejected = await saved(await load('colour-background-f00.docx'), 'word/document.xml');
  assert.ok(rejected.includes('<w:background w:color="auto"/>'), rejected);
});

async function synthetic() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const rev = ' w:author="A" w:date="2026-10-24T00:00:00Z"';
  const ps = [
    '<w:p><w:r><w:rPr><w:color w:val="bogus" w:themeColor="accent1"/></w:rPr><w:t>T1_THEME</w:t></w:r></w:p>',
    `<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="1"${rev}><w:rPr><w:color w:val="F00"/><w:highlight w:val="YELLOW"/><w:vertAlign/><w:em w:val="DOT"/><w:u w:val="single" w:color="red"/><w:shd w:val="clear" w:color="auto" w:fill="bogus"/></w:rPr></w:rPrChange></w:rPr><w:t>T2_REVISION</w:t></w:r></w:p>`,
    `<w:p><w:r><w:rPr><w:i/><w:rPrChange w:id="2"${rev}><w:rPr><w:color w:val="red"/><w:highlight w:val="green"/></w:rPr></w:rPrChange></w:rPr><w:t>T3_REVISION_NAME</w:t></w:r></w:p>`,
    `<w:p><w:pPr><w:rPr><w:b/><w:rPrChange w:id="3"${rev}><w:rPr><w:color w:val="F00"/><w:highlight w:val="bogus"/><w:color w:val="navy"/></w:rPr></w:rPrChange></w:rPr></w:pPr><w:r><w:t>T4_MARK_REVISION</w:t></w:r></w:p>`,
  ];
  const content = (await unmarshalString(`<w:document xmlns:w="${W}"><w:body>${ps.join('')}</w:body></w:document>`)).value.body.content;
  const body = (await pkg.getMainDocumentPart().getContents()).body;
  body.content.splice(0, body.content.length, ...content);
  return pkg;
}

test('beside a theme colour, a rejected w:val is written "auto" (it is required; the theme colour is drawn)', async () => {
  const xml = await saved(await synthetic(), 'word/document.xml');
  assert.match(runXml(xml, 'T1_THEME'), /<w:color (?=[^>]*w:val="auto")(?=[^>]*w:themeColor="accent1")[^>]*\/>/);
});

test('a revision\'s former run properties, which Reject All restores, are written as live ones are (D9)', async () => {
  const xml = await saved(await synthetic(), 'word/document.xml');
  const t2 = runXml(xml, 'T2_REVISION');
  for (const e of ['<w:color', '<w:highlight', '<w:vertAlign', '<w:em']) assert.ok(!t2.includes(e), `${e} in ${t2}`);
  assert.ok(t2.includes('w:color="FF0000"') && t2.includes('w:fill="auto"'), t2);
  const t3 = runXml(xml, 'T3_REVISION_NAME');
  assert.ok(t3.includes('<w:color w:val="FF0000"/>') && t3.includes('<w:highlight w:val="green"/>'), t3);
  const p4 = xml.indexOf('T4_MARK_REVISION');
  const mark = xml.substring(xml.lastIndexOf('<w:p>', p4), p4);
  assert.ok(!mark.includes('w:val="F00"') && !mark.includes('<w:highlight') && mark.includes('<w:color w:val="000080"/>'), mark);
});

test('only what is written is normalised: the tree keeps the value it read', async () => {
  const pkg = await synthetic();
  await saved(pkg, 'word/document.xml');
  const content = (await pkg.getMainDocumentPart().getContents()).body.content.map(unwrap);
  const former = (p) => runsOf(p)[0].rPr.rPrChange.rPr.egrPrBase.map(unwrap);
  assert.ok(former(content[2]).some((e) => e.TYPE_NAME === 'org_docx4j_wml.Color' && e.val === 'red'), 'the former w:color still reads "red"');
  assert.equal(former(content[1]).length, 6, 'the elements left out are still in the tree');
});

// docx4j ColourNormalFormOptOutTest: a package's normaliseOnSave = false, or the objects package's process-wide
// setNormaliseOnSave(false), writes colours as they were read and keeps what Word repairs; a package's setting wins.
test('a package that opts out keeps what it read (docx4j ColourNormalFormOptOutTest)', async () => {
  const pkg = await load('colour-hex.docx');
  pkg.normaliseOnSave = false;
  const xml = await saved(pkg, 'word/document.xml');
  assert.ok(runXml(xml, 'X1_ff0000').includes('w:val="ff0000"'));
  assert.ok(runXml(xml, 'X3_F00').includes('w:val="F00"'));
  const enums = await load('props-enum-missing.docx');
  enums.normaliseOnSave = false;
  assert.ok(runXml(await saved(enums, 'word/document.xml'), 'M1_HIGHLIGHT_BARE').includes('<w:highlight/>'));
  assert.deepEqual(pkg.lastSave.normalised, [], 'nothing normalised, nothing reported');
  const clone = await (await load('colour-hex.docx')).clone();
  assert.equal(clone.normaliseOnSave, undefined);
  pkg.normaliseOnSave = false;
  assert.equal((await pkg.clone()).normaliseOnSave, false, 'a clone carries the setting');
});

test('the process-wide setting off keeps what it read, and a package\'s own setting wins', async () => {
  setNormaliseOnSave(false);
  try {
    const xml = await saved(await load('colour-case.docx'), 'word/document.xml');
    assert.ok(runXml(xml, 'K4_Red').includes('w:val="Red"'));
    const on = await load('colour-case.docx');
    on.normaliseOnSave = true;
    assert.ok(runXml(await saved(on, 'word/document.xml'), 'K4_Red').includes('w:val="FF0000"'));
  } finally {
    setNormaliseOnSave(true);
  }
  assert.ok(!runXml(await saved(await load('colour-case.docx'), 'word/document.xml'), 'K4_Red').includes('w:val="Red"'));
});

test('a save reports each value it normalised (CR-015 D8)', async () => {
  const pkg = await load('colour-case.docx');
  assert.equal(pkg.lastSave, undefined);
  await saved(pkg, 'word/document.xml');
  const red = pkg.lastSave.normalised.find((n) => n.before === 'Red');
  assert.deepEqual({ ...red, message: undefined }, {
    code: 'value.normalised', partName: '/word/document.xml',
    path: red.path, before: 'Red', after: 'FF0000', message: undefined,
  });
  assert.match(red.path, /^\/w:document\/w:body\[1\]\/w:p\[\d+\]\/w:r\[\d+\]\/w:rPr\[1\]\/w:color\[1\]\/@w:val$/);
  assert.equal(red.message, `/word/document.xml: ${red.path} "Red" written "FF0000", as Word reads it`);
  const enums = await load('props-enum-missing.docx');
  await saved(enums, 'word/document.xml');
  assert.ok(enums.lastSave.normalised.some((n) => n.path.endsWith('/w:highlight') && n.after === null), 'an element left out');
  // a part nobody unmarshalled is copied, so nothing in it is normalised
  const untouched = await load('colour-case.docx');
  await untouched.save();
  assert.deepEqual(untouched.lastSave.normalised, []);
});

// The round-trip contract (re-marshalled parts deep-equal after reload) under the normaliser: the one tolerated
// difference is a value written in its normal form (CR-015 section 8; CR-018's word-normal-form). Under
// normaliseOnSave false the contract stands unchanged. pandoc-code.docx's styles part holds 20 lower-case hex colours.
test('a round trip is deep-equal but for values in their normal form, and exact with normaliseOnSave false', async () => {
  const bytes = readFileSync(new URL('./fixtures/onoff/pandoc-code.docx', import.meta.url));
  const differing = [];
  const compareTrees = (a, b, path) => {
    if (typeof a === 'string' && typeof b === 'string' && a !== b) {
      assert.equal(WordColors.normalise(a), b, `${path}: ${a} against ${b}`);
      differing.push(path);
      return;
    }
    if (a === null || b === null || typeof a !== 'object' || typeof b !== 'object') { assert.deepEqual(b, a, path); return; }
    const keys = new Set([...Object.keys(a), ...Object.keys(b)].filter((k) => k !== 'PARENT'));
    for (const k of keys) compareTrees(a[k], b[k], `${path}.${k}`);
  };
  for (const normalise of [undefined, false]) {
    const pkg = await WordprocessingMLPackage.load(bytes);
    pkg.normaliseOnSave = normalise;
    const styles = await pkg.getMainDocumentPart().styleDefinitionsPart.getContents();
    const again = await WordprocessingMLPackage.load(await pkg.save());
    differing.length = 0;
    compareTrees(styles, await again.getMainDocumentPart().styleDefinitionsPart.getContents(), 'styles');
    assert.equal(differing.length, normalise === false ? 0 : 20, String(normalise));
    assert.equal(pkg.lastSave.normalised.length, differing.length);
  }
});

test('WordColors.normalise and stated are docx4j\'s WordColorValues', () => {
  for (const v of ['FF0000', 'ff0000', '#ff0000', ' FF0000 ', 'red', 'RED', ' Red ']) assert.equal(WordColors.normalise(v), 'FF0000', v);
  for (const v of ['auto', 'Auto', 'F00', 'grey', '', undefined, 'lightGray']) assert.equal(WordColors.normalise(v), undefined, String(v));
  assert.equal(WordColors.stated('auto'), true, 'exactly auto overrides');
  for (const v of ['Auto', 'AUTO', ' auto', '#auto', 'F00', 'bogus', '', undefined]) assert.equal(WordColors.stated(v), false, String(v));
  assert.equal(WordColors.normalise('lightgrey'), 'D3D3D3');
  assert.equal(WordColors.normalise('lightgoldenrodyellow'), 'FAFAD2', 'not DrawingML\'s ltGoldenrodYellow');
});

// ---------------------------------------------------------------- check 60 (CR-015 D4): Office JS's own Font readings

const { check60Bytes, CHECK60_PROBES } = await import('./fixtures/check60/build.mjs');
const { Word } = await import('../dist/office-js/index.mjs');

async function check60() {
  const source = readFileSync(new URL('./fixtures/check60/check60-script-lab.js', import.meta.url), 'utf8');
  return new Function(`${source}\nreturn CHECK60;`)();
}

test('check 60: its file holds every marker of its probes, and the snippet run through the Word shim finds each', async () => {
  const { runCheck } = await check60();
  const pkg = await WordprocessingMLPackage.load(check60Bytes());
  await pkg.getBody();
  const ours = await runCheck((f) => Word.run(pkg, f), 'core-ts');
  assert.deepEqual([...new Set(ours.markers.map((m) => m.probe))], CHECK60_PROBES);
  for (const name of CHECK60_PROBES) {
    const pr = await probe(name);
    const entries = ours.markers.filter((m) => m.probe === name);
    const paragraphs = (await pr.pkg.getBody()).paragraphs;
    assert.equal(entries.length, paragraphs.length - 1, `${name}: a paragraph each, but the title`);
    for (const mk of pr.markers) {
      const entry = entries[mk.index - 1];
      const hidden = reading(pr.resolver, mk.p, mk.run).vanish === 'on';
      if (hidden) assert.equal(entry.marker, undefined, `${mk.key}: hidden, so no marker in the paragraph's text`);
      else {
        assert.equal(entry.marker, mk.text, mk.key);
        assert.ok(entry.reads.hits >= 1, `${mk.key}: found`);
      }
    }
  }
});

test('check 60 through the Word shim reads each marker as the probe\'s Font does', async () => {
  const { runCheck } = await check60();
  const pkg = await WordprocessingMLPackage.load(check60Bytes());
  await pkg.getBody();
  const ours = await runCheck((f) => Word.run(pkg, f), 'core-ts');
  const diffs = [];
  let n = 0;
  for (const entry of ours.markers.filter((m) => m.reads?.font !== undefined)) {
    const pr = await probe(entry.probe);
    const mk = pr.marker(entry.marker);
    const want = fontReading(mk.para, mk.text);
    for (const [k, v] of Object.entries(want)) {
      if (String(entry.reads.font[k]) !== v) diffs.push(`${mk.key} ${k}: check 60 ${entry.reads.font[k]}, the probe ${v}`);
    }
    n++;
  }
  assertNone(diffs);
  assert.ok(n > 300, `${n} markers read`);
});
