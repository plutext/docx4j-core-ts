// CR-001 Phase B step 2: the property catalogue, StyleUtil and PropertyResolver.
//
// The cases are docx4j's own, ported with their expected values: PropertyCatalogueTest,
// StyleUtilMergeRulesTest, TogglePropertiesTest, PropertyResolverOrderTest,
// PropertyResolverNoMutationTest and PropertyResolverTableStyleTest under
// docx4j-core-tests/src/test/java/org/docx4j/model/.  Parity against docx4j over 45 real
// documents is test/parity.test.mjs; this file is the rules one at a time.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import {
  PropertyResolver, WordprocessingMLPackage, PropertyResolverNotCreatedException,
  RUN, PARAGRAPH, TABLE, CELL, TOGGLES, TOGGLE_NAMES,
  RUN_EXCLUDED, PARAGRAPH_EXCLUDED, TABLE_EXCLUDED, CELL_EXCLUDED,
  isEmptyOf, hasDirectFormattingOf, applyCatalogue, unsetCatalogue, namedProperty,
  applyRPr, applyPPrBase, applyStyleLevel, toggle, isEmptyRPr, isEmptyPPrBase,
  hasDirectFormattingRPr, hasDirectFormattingPPr, throwOnCyclicStyles,
  headingLevelByName, WORD_DEFAULT_CELL_MARGIN_TWIPS,
  parseXml, unmarshalNode, marshalString, getContext, getContextSync, setLogger,
} from '../dist/index.mjs';
import { fixturesDir, plain } from './helpers.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function unmarshalStyles(inner) {
  const element = await unmarshalNode(parseXml(`<w:styles xmlns:w="${W}">${inner}</w:styles>`));
  return element.value;
}

async function unmarshalIn(container, inner, pick) {
  const element = await unmarshalNode(parseXml(`<w:${container} xmlns:w="${W}">${inner}</w:${container}>`));
  return pick(element.value);
}

const pPrOf = (inner) => unmarshalIn('p', `<w:pPr>${inner}</w:pPr>`, (v) => v.pPr);
const rPrOf = (inner) => unmarshalIn('r', `<w:rPr>${inner}</w:rPr>`, (v) => v.rPr);

// ---------------------------------------------------------------- the catalogue

/**
 * The members each generated type declares, from the runtime's own mapping model - the same
 * source the declarations in `modules/org_docx4j_wml.d.ts` are generated from - so a member
 * added to the schema and not to a table fails this test, which is what the four hand-kept
 * lists docx4j replaced used to miss.
 */
function declaredMembers(typeName) {
  const context = getContextSync();
  const names = [];
  for (let info = context.getTypeInfoByName(typeName); info; info = info.baseTypeInfo) {
    for (const property of info.properties ?? []) names.push(property.name);
  }
  return names;
}

await getContext();

test('catalogue: every declared member of every type is in its table or named as excluded', () => {
  const cases = [
    ['RUN (w:rPr)', 'org_docx4j_wml.RPr', RUN, RUN_EXCLUDED],
    ['RUN (w:pPr/w:rPr)', 'org_docx4j_wml.ParaRPr', RUN, RUN_EXCLUDED],
    ['PARAGRAPH', 'org_docx4j_wml.PPrBase', PARAGRAPH, PARAGRAPH_EXCLUDED],
    ['TABLE', 'org_docx4j_wml.CTTblPrBase', TABLE, TABLE_EXCLUDED],
    ['CELL', 'org_docx4j_wml.TcPr', CELL, CELL_EXCLUDED],
  ];
  for (const [label, typeName, catalogue, excluded] of cases) {
    const covered = new Set([...catalogue.map((p) => p.name), ...excluded]);
    const missing = declaredMembers(typeName).filter((name) => !covered.has(name));
    assert.deepEqual(missing, [], `${label} lacks ${missing.join(', ')}`);
    // and nothing in a table that the type does not declare
    const declared = new Set(declaredMembers(typeName));
    const unknown = catalogue.map((p) => p.name).filter((name) => !declared.has(name));
    assert.deepEqual(unknown, [], `${label} names ${unknown.join(', ')}, which ${typeName} does not declare`);
  }
});

test('catalogue: schema order, and the twelve toggles', () => {
  assert.deepEqual(RUN.slice(0, 4).map((p) => p.name), ['rStyle', 'rFonts', 'b', 'bCs']);
  assert.deepEqual(PARAGRAPH.slice(0, 3).map((p) => p.name), ['pStyle', 'keepNext', 'keepLines']);
  assert.deepEqual([...TOGGLE_NAMES],
    ['b', 'bCs', 'caps', 'emboss', 'i', 'iCs', 'imprint', 'outline', 'shadow', 'smallCaps', 'strike', 'vanish']);
  assert.equal(TOGGLES.length, 12);
  // Boolean, but not toggles: the last value in the order wins for these
  for (const notAToggle of ['dstrike', 'noProof', 'snapToGrid', 'webHidden', 'rtl', 'cs', 'specVanish', 'oMath']) {
    assert.ok(namedProperty(RUN, notAToggle), notAToggle);
    assert.ok(!TOGGLE_NAMES.has(notAToggle), notAToggle);
  }
});

/**
 * A value that states this member and nothing else: the first candidate the entry's own
 * `isEmpty` calls non-empty. Probing rather than a table per type keeps the test honest -
 * every entry must have *some* value its emptiness rule accepts.
 */
const CANDIDATES = [
  {},                                     // w:b and the w14 members: the element itself is the statement
  { val: 1 },                             // w:sz, w:jc, w:outlineLvl, w:gridSpan, ...
  { val: 'single' },                      // w:u, w:highlight, w:rStyle, w:vMerge, ...
  { ascii: 'Arial' },                     // w:rFonts
  { bidi: 'he-IL' },                      // w:lang
  { ilvl: { val: 1 } },                   // w:numPr, stating only the level
  { left: 720 },                          // w:ind
  { after: 120 },                         // w:spacing (paragraph)
  { tab: [{ pos: 720, val: 'left' }] },   // w:tabs
  { top: { val: 'single' } },             // w:pBdr, w:tblBorders, w:tcBorders
  { top: { w: 100, type: 'dxa' } },       // w:tblCellMar, w:tcMar
  { w: 3600 },                            // w:framePr, w:tblW, w:tblInd
  { type: 'fixed' },                      // w:tblLayout
  { firstRow: '1' },                      // w:tblLook
  { tblpX: 100 },                         // w:tblpPr
];

function sampleFor(p) {
  for (const candidate of CANDIDATES) {
    const value = JSON.parse(JSON.stringify(candidate));
    if (!p.isEmpty(value)) return value;
  }
  throw new Error(`no sample value satisfies ${p.name}'s isEmpty`);
}

/** Members docx4j's hand-written `apply` for the owning element does not carry. */
const NOT_CARRIED = new Set(['tblCaption', 'tblDescription']);

for (const [label, catalogue] of [['RUN', RUN], ['PARAGRAPH', PARAGRAPH], ['TABLE', TABLE], ['CELL', CELL]]) {
  test(`catalogue ${label}: every member on its own is stated, carried as a copy, and unset`, () => {
    for (const p of catalogue) {
      const sample = sampleFor(p);
      const source = {};
      p.set(source, sample);

      if (p.countsEmpty) {
        assert.equal(isEmptyOf(catalogue, source), false, `${p.name}: a source stating it is not empty`);
      }
      assert.equal(hasDirectFormattingOf(catalogue, source), p.isFormatting, `${p.name}: direct formatting`);

      const destination = {};
      applyCatalogue(catalogue, source, destination);
      if (NOT_CARRIED.has(p.name)) {
        assert.equal(p.get(destination), undefined, `${p.name}: docx4j's apply does not carry it`);
        continue;
      }
      const carried = p.get(destination);
      assert.notEqual(carried, undefined, `${p.name}: apply carries it`);
      assert.equal(p.isEmpty(carried), false, `${p.name}: carried non-empty`);
      assert.notEqual(carried, sample, `${p.name}: the carried value is a copy, not the source's instance`);
      // w:framePr is the one member whose merge writes a default out: docx4j reads w:anchorLock
      // through isAnchorLock(), which answers true when the attribute is absent
      const expected = p.name === 'framePr' ? { ...plain(sample), anchorLock: true } : plain(sample);
      assert.deepEqual(plain(carried), expected, `${p.name}: and equal to it`);

      unsetCatalogue(catalogue, source, destination);
      assert.equal(p.get(destination), undefined, `${p.name}: unset removes it`);
    }
  });
}

test('catalogue: an empty element is empty and states no direct formatting', () => {
  assert.equal(isEmptyRPr({}), true);
  assert.equal(isEmptyRPr(undefined), true);
  assert.equal(isEmptyPPrBase({}), true);
  assert.equal(hasDirectFormattingRPr({}), false);
  assert.equal(hasDirectFormattingRPr(undefined), false);
  assert.equal(hasDirectFormattingPPr({}), false);
  // the members that are not formatting do not make a source "directly formatted"
  assert.equal(hasDirectFormattingRPr({ rStyle: { val: 'Strong' } }), false);
  assert.equal(hasDirectFormattingPPr({ pStyle: { val: 'Heading1' } }), false);
  // ... but w:rtl-only and w:position-only direct formatting is direct formatting
  assert.equal(hasDirectFormattingRPr({ rtl: {} }), true);
  assert.equal(hasDirectFormattingRPr({ position: { val: 6 } }), true);
  assert.equal(hasDirectFormattingPPr({ mirrorIndents: {} }), true);
});

// ---------------------------------------------------------------- the merge rules

test('merge: w:lineRule is taken only beside a w:line the source states (CR-015 probe styles-linerule)', async () => {
  const style = await pPrOf('<w:spacing w:line="480" w:lineRule="exact"/>');
  const direct = await pPrOf('<w:spacing w:after="0"/>');
  const destination = {};
  applyPPrBase(style, destination);
  applyPPrBase(direct, destination);
  assert.equal(destination.spacing.line, 480);
  assert.equal(destination.spacing.lineRule, 'exact', 'the exact rule is inherited, not reset to auto');
  assert.equal(destination.spacing.after, 0);

  // a source stating w:line and no rule states auto
  const auto = await pPrOf('<w:spacing w:line="240"/>');
  applyPPrBase(auto, destination);
  assert.equal(destination.spacing.lineRule, 'auto');
});

test('merge: w:numId and w:ilvl each inherit (CR-015 probe styles-numpr-ilvl-only)', async () => {
  const destination = {};
  applyPPrBase(await pPrOf('<w:numPr><w:numId w:val="3"/><w:ilvl w:val="0"/></w:numPr>'), destination);
  applyPPrBase(await pPrOf('<w:numPr><w:ilvl w:val="1"/></w:numPr>'), destination);
  assert.equal(destination.numPr.numId.val, 3, 'the id is inherited');
  assert.equal(destination.numPr.ilvl.val, 1, 'the level is the more specific one');
});

test('merge: w:lang, w:u and w:tblpPr inherit per attribute', async () => {
  const destination = {};
  applyRPr(await rPrOf('<w:lang w:val="en-US" w:eastAsia="ja-JP"/>'), destination);
  applyRPr(await rPrOf('<w:lang w:bidi="he-IL"/>'), destination);
  assert.deepEqual([destination.lang.val, destination.lang.eastAsia, destination.lang.bidi], ['en-US', 'ja-JP', 'he-IL']);

  const u = {};
  applyRPr(await rPrOf('<w:u w:val="single" w:color="FF0000"/>'), u);
  applyRPr(await rPrOf('<w:u w:val="double"/>'), u);
  assert.deepEqual([u.u.val, u.u.color], ['double', 'FF0000']);
});

test('merge: w:tabs are cumulative and w:val="clear" removes', async () => {
  const destination = {};
  applyPPrBase(await pPrOf('<w:tabs><w:tab w:val="left" w:pos="284"/><w:tab w:val="left" w:pos="6804"/><w:tab w:val="right" w:pos="10348"/></w:tabs>'), destination);
  applyPPrBase(await pPrOf('<w:tabs><w:tab w:val="clear" w:pos="6804"/><w:tab w:val="left" w:pos="6379"/></w:tabs>'), destination);
  assert.deepEqual(destination.tabs.tab.map((t) => [t.pos, t.val]), [[284, 'left'], [6379, 'left'], [10348, 'right']]);
});

test('merge: w:ind firstLine and hanging are one property', async () => {
  const destination = {};
  applyPPrBase(await pPrOf('<w:ind w:left="283" w:hanging="283"/>'), destination);
  applyPPrBase(await pPrOf('<w:ind w:left="0" w:firstLine="709"/>'), destination);
  assert.equal(destination.ind.firstLine, 709);
  assert.equal(destination.ind.hanging, undefined, 'the inherited hanging indent is gone');
  assert.equal(destination.ind.left, 0);
});

test('merge: a leaf carried into the destination is a copy, not the source object', async () => {
  const source = await rPrOf('<w:b/><w:sz w:val="28"/>');
  const destination = {};
  applyRPr(source, destination);
  assert.notEqual(destination.b, source.b);
  assert.notEqual(destination.sz, source.sz);
  destination.b.val = false;
  assert.equal(source.b.val, undefined, 'editing the result leaves the source alone');
});

// ---------------------------------------------------------------- toggle properties

const on = () => ({ val: true });
const off = () => ({ val: false });

test('toggles: the rule of ECMA-376-1 17.7.3 as Word settled it', () => {
  assert.equal(toggle(on(), on(), undefined).val, false, 'two levels saying true cancel');
  assert.equal(toggle(on(), undefined, undefined).val, true, 'one level saying true applies');
  assert.equal(toggle(undefined, on(), undefined).val, true);
  // three levels: true for an odd number
  assert.equal(toggle(on(), toggle(on(), on(), undefined), undefined).val, true);
  // an explicit false is a term of the XOR like any other value (the toggle-levels golden)
  assert.equal(toggle(off(), on(), undefined).val, true, 'false XOR true = true: the lower level stands');
  assert.equal(toggle(off(), off(), undefined).val, false);
  assert.equal(toggle(off(), undefined, undefined).val, false, 'the only level stating it states false');
  // a level saying nothing leaves what is beneath alone
  assert.equal(toggle(undefined, undefined, undefined), undefined);
  assert.equal(toggle(undefined, off(), undefined).val, false);
  // the document defaults' true wins where the upper level states the property
  assert.equal(toggle(on(), on(), on()).val, true);
  assert.equal(toggle(off(), off(), on()).val, true);
  assert.equal(toggle(on(), on(), off()).val, false, 'a default of false is no default for this rule');
  // ... and a silent upper level is no boundary, so it does not force the default's true back
  assert.equal(toggle(undefined, off(), on()).val, false);
  assert.equal(toggle(undefined, on(), on()).val, true);
});

test('toggles: applyStyleLevel XORs the toggles and overrides everything else', async () => {
  const level = await rPrOf('<w:b/><w:i/><w:dstrike/><w:sz w:val="24"/>');
  const beneath = await rPrOf('<w:b/><w:dstrike/><w:sz w:val="20"/>');
  applyStyleLevel(level, beneath, undefined);
  assert.equal(beneath.b.val, false, 'b is a toggle: true XOR true');
  assert.equal(beneath.i.val ?? true, true, 'i is stated at one level only');
  assert.equal(beneath.dstrike.val ?? true, true, 'dstrike is NOT a toggle: the last value wins');
  assert.equal(beneath.sz.val, 24, 'nor is the size');
});

const TOGGLE_STYLES = `
  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="BoldPara"><w:name w:val="Bold Para"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="BoldChild"><w:name w:val="Bold Child"/><w:basedOn w:val="BoldPara"/><w:rPr><w:b/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/><w:bCs/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="NotStrong"><w:name w:val="Not Strong"/><w:rPr><w:b w:val="0"/></w:rPr></w:style>`;

test('toggles: through the resolver, at the paragraph-style to character-style boundary', async () => {
  const resolver = PropertyResolver.fromSource({ styles: await unmarshalStyles(TOGGLE_STYLES) });
  const boldPara = await pPrOf('<w:pStyle w:val="BoldPara"/>');
  const boldChild = await pPrOf('<w:pStyle w:val="BoldChild"/>');
  const strong = await rPrOf('<w:rStyle w:val="Strong"/>');
  const notStrong = await rPrOf('<w:rStyle w:val="NotStrong"/>');

  assert.equal(resolver.getEffectiveRPr(undefined, boldPara).b.val ?? true, true, 'the paragraph style alone');
  assert.equal(resolver.getEffectiveRPr(strong, undefined).b.val ?? true, true, 'the character style alone');
  assert.equal(resolver.getEffectiveRPr(strong, boldPara).b.val, false, 'both levels: bold XOR bold');
  assert.equal(resolver.getEffectiveRPr(strong, boldPara).bCs.val ?? true, true,
    'w:bCs is counted on its own, not with w:b');

  // a w:basedOn chain is ONE level and is not XORed
  assert.equal(resolver.getEffectiveRPr(undefined, boldChild).b.val ?? true, true);
  assert.equal(resolver.getEffectiveRPr(strong, boldChild).b.val, false);

  // a character style stating false does not unbold a bold paragraph style (the golden)
  assert.equal(resolver.getEffectiveRPr(notStrong, boldPara).b.val ?? true, true);

  // direct formatting is not a level: an explicit value there is used as it stands
  assert.equal(resolver.getEffectiveRPr(await rPrOf('<w:b w:val="0"/>'), boldPara).b.val, false);
  assert.equal(resolver.getEffectiveRPr(await rPrOf('<w:rStyle w:val="Strong"/><w:b/>'), boldPara).b.val ?? true, true);
});

// ---------------------------------------------------------------- the resolution order

/**
 * docx4j's `PropertyResolverOrderTest` styles part: docDefaults Calibri 11pt lang en-US;
 * Normal (default) 14pt Times New Roman; H based on Normal, Arial 20pt; character styles S
 * (w:b) and DefaultParagraphFont.
 */
const ORDER_STYLES = `
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>
    <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="DefaultParagraphFont"/></w:style>
  <w:style w:type="paragraph" w:styleId="H"><w:name w:val="H"/><w:basedOn w:val="Normal"/>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="40"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="S"><w:name w:val="S"/><w:basedOn w:val="DefaultParagraphFont"/>
    <w:rPr><w:b w:val="true"/></w:rPr></w:style>`;

async function orderResolver() {
  return PropertyResolver.fromSource({ styles: await unmarshalStyles(ORDER_STYLES) });
}

const sz = (rPr) => {
  assert.ok(rPr, 'an rPr');
  assert.ok(rPr.sz, 'a w:sz');
  return rPr.sz.val;
};

test('order: a run in a paragraph naming no style gets the default style\'s run properties', async () => {
  const resolver = await orderResolver();
  assert.equal(sz(resolver.getEffectiveRPr(undefined, undefined)), 28, 'null pPr');
  assert.equal(sz(resolver.getEffectiveRPr(undefined, {})), 28, 'pPr without w:pStyle');
  assert.equal(sz(resolver.getEffectiveRPr(undefined, await pPrOf('<w:pStyle w:val="Normal"/>'))), 28, 'w:pStyle Normal');
  const effective = resolver.getEffectiveRPr(await rPrOf('<w:b/>'), {});
  assert.equal(sz(effective), 28, 'a bold run, no w:pStyle (golden (c))');
  assert.equal(effective.rFonts.ascii, 'Times New Roman');
  assert.ok(effective.b);
});

test('order: a paragraph naming a missing style resolves as the default style', async () => {
  const resolver = await orderResolver();
  const missing = await pPrOf('<w:pStyle w:val="Missing"/>');
  assert.equal(sz(resolver.getEffectiveRPr(undefined, missing)), 28, 'run (golden (d))');
  const effective = resolver.getEffectivePPr(missing);
  assert.ok(effective, 'pPr no longer null');
  assert.equal(effective.spacing.after, 160, 'the document defaults are in it');
  assert.equal(resolver.getEffectivePPr('Normal'), effective, 'and it is the default style\'s own entry');
  const direct = await pPrOf('<w:pStyle w:val="Missing"/><w:jc w:val="center"/>');
  assert.equal(resolver.getEffectivePPr(direct).spacing.after, 160, 'with direct formatting too');
});

test('order: a character style in a styled paragraph keeps the paragraph style\'s font and size', async () => {
  const resolver = await orderResolver();
  const effective = resolver.getEffectiveRPr(await rPrOf('<w:rStyle w:val="S"/>'), await pPrOf('<w:pStyle w:val="H"/>'));
  assert.equal(effective.rFonts.ascii, 'Arial');
  assert.equal(sz(effective), 40, '20pt (golden (e))');
  assert.ok(effective.b);
  assert.equal(effective.lang.val, 'en-US');
});

test('order: the answer does not depend on which caller came first', async () => {
  const s = await rPrOf('<w:rStyle w:val="S"/>');
  const h = await pPrOf('<w:pStyle w:val="H"/>');

  const first = await orderResolver();
  const byStyle = first.getEffectiveRPr('S');
  assert.equal(byStyle.rFonts.ascii, 'Calibri', 'a character style on its own: the document defaults under it');
  assert.equal(sz(byStyle), 22);
  const effective = first.getEffectiveRPr(s, h);
  assert.equal(effective.rFonts.ascii, 'Arial');
  assert.equal(sz(effective), 40);

  const second = await orderResolver();
  second.getEffectiveRPr(s, h);
  const byStyle2 = second.getEffectiveRPr('S');
  assert.equal(byStyle2.rFonts.ascii, 'Calibri');
  assert.equal(sz(byStyle2), 22);
});

test('order: DefaultParagraphFont adds nothing, and a missing character style is ignored', async () => {
  const resolver = await orderResolver();
  const h = await pPrOf('<w:pStyle w:val="H"/>');
  const effective = resolver.getEffectiveRPr(await rPrOf('<w:rStyle w:val="DefaultParagraphFont"/>'), h);
  assert.equal(sz(effective), 40, '20pt, not Normal\'s 14pt (golden (f))');
  assert.equal(effective.rFonts.ascii, 'Arial');
  assert.equal(sz(resolver.getEffectiveRPr(await rPrOf('<w:rStyle w:val="Nope"/>'), h)), 40);
  assert.equal(resolver.getEffectiveRPr('Nope'), undefined);
});

test('order: the paragraph mark formats the mark, not the runs', async () => {
  const resolver = await orderResolver();
  const marked = await pPrOf('<w:rPr><w:b/><w:sz w:val="40"/></w:rPr>');

  const run = resolver.getEffectiveRPr(undefined, marked);
  assert.equal(run.b, undefined, 'a run with no rPr is not bold (ECMA 17.3.1.29)');
  assert.equal(sz(run), 28);

  const mark = resolver.getEffectiveParagraphMarkRPr(marked);
  assert.ok(mark.b);
  assert.equal(sz(mark), 40, 'the mark: defaults, Normal, then its own rPr');
  assert.equal(mark.rFonts.ascii, 'Times New Roman');

  assert.equal(sz(resolver.getEffectiveParagraphMarkRPr({})), 28, 'a mark with no rPr is the run baseline');
  assert.equal(sz(resolver.getEffectiveParagraphMarkRPr(undefined)), 28);
});

test('order: the no-paragraph overload uses the default paragraph style', async () => {
  const resolver = await orderResolver();
  const effective = resolver.getEffectiveRPr(await rPrOf('<w:rStyle w:val="S"/>'));
  assert.equal(effective.rFonts.ascii, 'Times New Roman');
  assert.equal(sz(effective), 28);
  assert.ok(effective.b);
  assert.equal(sz(resolver.getEffectiveRPr({})), 28, 'with no rStyle either');
});

test('order: the style overloads include the document defaults, and are cached', async () => {
  const resolver = await orderResolver();
  const h = resolver.getEffectiveRPr('H');
  assert.equal(sz(h), 40);
  assert.equal(h.lang.val, 'en-US');
  const hp = resolver.getEffectivePPr('H');
  assert.equal(hp.spacing.after, 160);
  assert.equal(resolver.getEffectivePPr('H'), hp, 'cached');
  assert.equal(resolver.getResolvedDefaultParagraphStyle(), resolver.getEffectivePPr('Normal'));
  assert.equal(resolver.getDefaultParagraphStyleId(), 'Normal');
  // the chains carry no document defaults
  assert.equal(resolver.getChainRPr('H').lang, undefined);
  assert.equal(resolver.getChainRPr('H').sz.val, 40);
  assert.deepEqual(resolver.ancestry('H').map((s) => s.styleId), ['Normal', 'H'], 'root first');
});

// ---------------------------------------------------------------- cycles, headings, refresh

const CYCLE_STYLES = `
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="A"><w:name w:val="A"/><w:basedOn w:val="B"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="B"><w:name w:val="B"/><w:basedOn w:val="A"/><w:rPr><w:b/></w:rPr></w:style>`;

test('cycles: a w:basedOn cycle stops the walk, and throws when asked to', async () => {
  const resolver = PropertyResolver.fromSource({ styles: await unmarshalStyles(CYCLE_STYLES) });
  setLogger({ warn() {} });
  try {
    // docx4j's default: degrade gracefully, using as much of the hierarchy as there is
    assert.deepEqual(resolver.ancestry('A').map((s) => s.styleId), ['B', 'A']);
    assert.equal(resolver.getEffectiveRPr('A').sz.val, 24);

    throwOnCyclicStyles(true);
    const again = PropertyResolver.fromSource({ styles: await unmarshalStyles(CYCLE_STYLES) });
    assert.throws(() => again.ancestry('A'), (e) => e.name === 'CyclicStylesException');
  } finally {
    throwOnCyclicStyles(false);
    setLogger();
  }
});

test('headings: the outline level comes from the style name, into the resolved pPr only', async () => {
  const styles = await unmarshalStyles(`
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
    <w:style w:type="paragraph" w:styleId="berschrift2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
      <w:pPr><w:outlineLvl w:val="5"/></w:pPr></w:style>`);
  const resolver = PropertyResolver.fromSource({ styles });
  const effective = resolver.getEffectivePPr('berschrift2');
  assert.equal(effective.outlineLvl.val, 1, "level 1 for 'heading 2', whatever the style declares");
  const style = styles.style.find((s) => s.styleId === 'berschrift2');
  assert.equal(style.pPr.outlineLvl.val, 5, 'the style still declares 5');
  assert.equal(headingLevelByName(style), 2);
  assert.equal(headingLevelByName(styles.style[0]), -1);
  assert.equal(resolver.getLvlFromHeadingStyle('Heading3'), 3);
  assert.equal(resolver.getLvlFromHeadingStyle('berschrift2'), -1, 'the id rule is English-only, by design');
});

test('refresh: a style added to the part is found, and a changed one after refresh()', async () => {
  const styles = await unmarshalStyles(`
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>`);
  const resolver = PropertyResolver.fromSource({ styles });
  assert.equal(resolver.getEffectiveRPr('Added'), undefined);

  const added = (await unmarshalStyles('<w:style w:type="paragraph" w:styleId="Added"><w:name w:val="Added"/><w:rPr><w:sz w:val="48"/></w:rPr></w:style>')).style[0];
  styles.style.push(added);
  assert.equal(resolver.getEffectiveRPr('Added').sz.val, 48, 'an added style is found without refresh()');

  added.rPr.sz.val = 60;
  assert.equal(resolver.getEffectiveRPr('Added').sz.val, 48, 'a change to a resolved style is cached');
  resolver.refresh();
  assert.equal(resolver.getEffectiveRPr('Added').sz.val, 60, 'refresh() re-reads it');
});

// ---------------------------------------------------------------- no mutation

const NO_MUTATION_STYLES = `
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="DefaultParagraphFont"/></w:style>
  <w:style w:type="paragraph" w:styleId="berschrift2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="5"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="L"><w:name w:val="L"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:numPr><w:numId w:val="90"/><w:ilvl w:val="0"/></w:numPr></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="L2"><w:name w:val="L2"/><w:basedOn w:val="L"/>
    <w:pPr><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr></w:style>
  <w:style w:type="character" w:styleId="S"><w:name w:val="S"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:b/></w:rPr></w:style>`;

test('no mutation: resolving everything leaves the styles part exactly as it was', async () => {
  const styles = await unmarshalStyles(NO_MUTATION_STYLES);
  const asElement = () => ({ name: { namespaceURI: W, localPart: 'styles' }, value: styles });
  const before = await marshalString(asElement());

  const resolver = PropertyResolver.fromSource({ styles });
  for (const style of styles.style) {
    if (style.type === 'paragraph') resolver.getEffectivePPr(style.styleId);
    resolver.getEffectiveRPr(style.styleId);
  }
  resolver.getEffectivePPr(await pPrOf('<w:pStyle w:val="berschrift2"/>'));
  resolver.getEffectiveRPr(await rPrOf('<w:rStyle w:val="S"/>'), await pPrOf('<w:pStyle w:val="L2"/>'));
  resolver.getEffectiveParagraphMarkRPr(await pPrOf('<w:pStyle w:val="Normal"/>'));
  setLogger({ warn() {} });
  try {
    resolver.getEffectiveRPr(undefined, await pPrOf('<w:pStyle w:val="Missing"/>'));
  } finally {
    setLogger();
  }

  assert.equal(await marshalString(asElement()), before, 'the styles part is unchanged');
  assert.equal(styles.docDefaults.rPrDefault.rPr.sz, undefined,
    "the 10pt default is the resolver's, not the part's");
  assert.equal(resolver.getDocumentDefaultRPr().sz.val, 20);
});

test('no mutation: an inherited w:numId reaches the effective pPr, and no leaf is shared', async () => {
  const styles = await unmarshalStyles(NO_MUTATION_STYLES);
  const resolver = PropertyResolver.fromSource({ styles });
  const find = (id) => styles.style.find((s) => s.styleId === id);

  const effectivePPr = resolver.getEffectivePPr('L2');
  assert.equal(effectivePPr.numPr.numId.val, 90);
  assert.equal(effectivePPr.numPr.ilvl.val, 1);
  assert.equal(find('L2').pPr.numPr.numId, undefined, "L2's own w:numPr still names no w:numId");
  assert.notEqual(effectivePPr.numPr.numId, find('L').pPr.numPr.numId, 'w:numId is a copy');
  assert.notEqual(effectivePPr.numPr.ilvl, find('L2').pPr.numPr.ilvl, 'w:ilvl is a copy');

  const effective = resolver.getEffectiveRPr(await rPrOf('<w:rStyle w:val="S"/>'), await pPrOf('<w:pStyle w:val="Normal"/>'));
  assert.ok(effective.b);
  assert.notEqual(effective.b, find('S').rPr.b, 'w:b is a copy of the style\'s');
  assert.notEqual(effective.sz, find('Normal').rPr.sz, 'w:sz is a copy of Normal\'s');
  assert.notEqual(effective.rFonts, styles.docDefaults.rPrDefault.rPr.rFonts, 'w:rFonts is a copy of the defaults\'');

  effective.b.val = false;
  assert.equal(find('S').rPr.b.val, undefined, 'editing the effective object changes nothing in the part');
});

// ---------------------------------------------------------------- table styles

test('tables: the built-in Normal Table underlies a chain that reaches the default style', async () => {
  const bytes = new Uint8Array(await readFile(join(fixturesDir, 'parity', 'styles-table-default.docx')));
  const pkg = await WordprocessingMLPackage.load(bytes);
  const resolver = await pkg.getPropertyResolver();
  const document = await pkg.getMainDocumentPart().getContents();
  const tables = (document.body.content ?? []).map((e) => e.value).filter((v) => v.TYPE_NAME === 'org_docx4j_wml.Tbl');
  assert.equal(tables.length, 3, 'the probe has three tables: (a) no style, (b) Custom, (c) Grid2');

  const left = (style) => style.tblPr?.tblCellMar?.left?.w;
  const right = (style) => style.tblPr?.tblCellMar?.right?.w;

  // the document's own Table Normal says 300; Word applies its built-in 108 regardless
  const tableNormal = resolver.getStyle(resolver.getEffectiveTableStyle(undefined).styleId);
  assert.equal(tableNormal.tblPr.tblCellMar.left.w, 300, 'the document says 300');

  // (a) a table naming no style
  assert.equal(resolver.reachesDefaultTableStyle(tables[0].tblPr), true);
  assert.equal(left(resolver.getEffectiveTableStyle(tables[0].tblPr)), WORD_DEFAULT_CELL_MARGIN_TWIPS);
  assert.equal(right(resolver.getEffectiveTableStyle(tables[0].tblPr)), 108);
  assert.equal(resolver.getEffectiveTableStyle(tables[0].tblPr).tblPr.tblInd.w, 0);

  // (b) a style with no w:basedOn: no cell margin at all
  assert.equal(resolver.reachesDefaultTableStyle(tables[1].tblPr), false);
  assert.equal(resolver.getEffectiveTableStyle(tables[1].tblPr).tblPr.tblCellMar, undefined,
    'a chain that does not reach the default gets no margin (golden (b))');

  // (c) a style based on Table Normal: the built-in, not the document's 300
  assert.equal(resolver.reachesDefaultTableStyle(tables[2].tblPr), true);
  assert.equal(left(resolver.getEffectiveTableStyle(tables[2].tblPr)), 108, 'not the document\'s 300 (golden (c))');

  // a missing table style is the built-in; and the table's own properties win
  assert.equal(left(resolver.getEffectiveTableStyle({ tblStyle: { val: 'Nope' } })), 108);
  const own = { tblStyle: { val: tables[2].tblPr.tblStyle.val }, tblCellMar: { left: { w: 50, type: 'dxa' } } };
  const effective = resolver.getEffectiveTableStyle(own);
  assert.equal(left(effective), 50);
  assert.equal(right(effective), 108, 'the right margin is still the built-in\'s');
  assert.equal(tableNormal.tblPr.tblCellMar.left.w, 300, 'the document\'s Table Normal is untouched');
});

// ---------------------------------------------------------------- wiring

test('wiring: getPropertyResolver caches, the synchronous accessor names itself when it cannot', async () => {
  const bytes = new Uint8Array(await readFile(join(fixturesDir, 'loadAndSave.docx')));
  const pkg = await WordprocessingMLPackage.load(bytes);
  assert.throws(() => pkg.propertyResolver, PropertyResolverNotCreatedException);
  assert.throws(() => pkg.propertyResolver, /getPropertyResolver\(\)/);
  const resolver = await pkg.getPropertyResolver();
  assert.equal(await pkg.getPropertyResolver(), resolver, 'created once');
  assert.equal(pkg.propertyResolver, resolver);
  // reading the styles part did not cost it its byte-for-byte round trip
  assert.equal(pkg.getMainDocumentPart().styleDefinitionsPart.isUnmarshalled, false);
});

test('wiring: createPackage builds a resolver, and the content API reads effective values', async () => {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const plainParagraph = body.insertParagraph('Plain', 'End');
  assert.equal(plainParagraph.font.size, 11, 'the document defaults say 22 half-points');
  assert.equal(plainParagraph.getFont({ direct: true }).size, 0, 'and nothing is set directly');
  const p = body.insertParagraph('Hello', 'End');
  p.styleBuiltIn = 'Heading1';
  assert.equal(p.font.size, 14, "Heading1's own w:sz 28");
  assert.equal(p.getFont({ direct: true }).size, 0, 'and nothing is set directly');
  assert.equal(p.alignment, 'Left');
  assert.equal(p.formatting({ direct: true }).alignment, 'Unknown');
  assert.equal(p.spaceAfter, 10, 'the default styles say w:after 200');
  assert.equal(p.formatting({ direct: true }).spaceAfter, 0);
  assert.equal(p.outlineLevel, 1, 'Heading1 supplies w:outlineLvl 0');
  assert.equal(p.effectivePPr.outlineLvl.val, 0);
});
