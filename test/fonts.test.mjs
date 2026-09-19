// CR-001 Phase B step 4: fonts.
//
// The docx4j CR-016 cases, ported: the character-range dispatch (phase 2), `w:cs` and `w:rtl`
// by value, the theme language, a run with no `w:rFonts`, the symbol fonts and emoji (phase 1);
// `LanguageTagToScriptMappingTest`, `MapperPrecedenceTest`, the altName chain, `NoBoldFaceTest`,
// `MetricallyCompatibleSubstituteTest`, `ClassBasedSubstituteTest` and `FontsInUseTest`
// (phases 3 and 4); and the themeless default and `createPackage`'s theme part, which docx4j
// gained in CR-001 batch 49.
//
// **Document fonts, not physical ones.**  docx4j's tests map their probe names onto the two
// Liberation faces and assert the `font-family` the FO carries, because its selector's answer
// *is* a fragment.  Here `spans()` answers document font names, so the assertions read the
// names themselves; the mapper's tests are the ones that name physical faces, over the default
// registry (the docx4j font jars).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  WordprocessingMLPackage, MainDocumentPart, StyleDefinitionsPart, DocumentSettingsPart,
  NumberingDefinitionsPart, ThemePart, HeaderPart, FontTablePart,
  IdentityPlusMapper, SimpleFontRegistry, PhysicalFont, DEFAULT_FONT_REGISTRY,
  RunFontSelector, symbolFontName, isEmoji, spanScript, coverageGroupOf, isSymbol,
  getScriptForLanguageTag, classOf, substitutionClass, selectByClass, isKnownFamily,
  hasBoldFace, wordDefaultFor, fontTableOf, NOBOLD_SUFFIX, toEnglish,
  THEME_2007, THEME_2013, THEME_2023, defaultThemeOf, defaultThemeSetting,
} from '../dist/index.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

// ------------------------------------------------------------------ the test harness
//
// docx4j's FontsTestSupport, over this package's parts: everything is set as XML, which is what
// the selector reads (`readContents()`), so nothing has to be built by hand.

function styles(docDefaultsRPr, moreStyles) {
  return `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${docDefaultsRPr}`
    + '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    + '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
    + '<w:name w:val="Default Paragraph Font"/></w:style>'
    + `${moreStyles ?? ''}</w:styles>`;
}

function settings(themeFontLangAttrs) {
  return `<w:settings xmlns:w="${W}"><w:themeFontLang ${themeFontLangAttrs}/></w:settings>`;
}

/** `a:fontScheme` with these Latin faces and these `a:font` script entries in both collections. */
function fontScheme(majorLatin, minorLatin, scriptFonts = '') {
  const collection = (tag, latin) => `<a:${tag}><a:latin typeface="${latin}"/><a:ea typeface=""/>`
    + `<a:cs typeface=""/>${scriptFonts}</a:${tag}>`;
  return `<a:fontScheme name="Office">${collection('majorFont', majorLatin)}`
    + `${collection('minorFont', minorLatin)}</a:fontScheme>`;
}

function themeXml(scheme) {
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const ln = `<a:ln w="9525">${fill}</a:ln>`;
  return `<a:theme xmlns:a="${A}" name="Office"><a:themeElements>`
    + '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    + '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2>'
    + '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1>'
    + '<a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3>'
    + '<a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5>'
    + '<a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink>'
    + '<a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>'
    + scheme
    + `<a:fmtScheme name="Office"><a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst>`
    + `<a:lnStyleLst>${ln}${ln}${ln}</a:lnStyleLst>`
    + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>'
    + '<a:effectStyle><a:effectLst/></a:effectStyle>'
    + '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    + `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst></a:fmtScheme>`
    + '</a:themeElements></a:theme>';
}

/** A `w:p` holding one run with these run properties (null for none) and this text. */
function p(rPrXml, text) {
  return `<w:p><w:r>${rPrXml === null ? '' : `<w:rPr>${rPrXml}</w:rPr>`}`
    + `<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

/**
 * A package built from XML: the styles, settings, theme (a bare `a:fontScheme`; undefined for
 * no theme part at all) and body.  No `createPackage`, so "no theme part" is the default here
 * rather than something to undo.
 */
async function packageWith({ stylesXml, settingsXml, fontSchemeXml, numberingXml, fontTableXml,
  headerXml, bodyXml = '', defaultTheme } = {}) {
  const pkg = new WordprocessingMLPackage();
  if (defaultTheme !== undefined) pkg.fonts.defaultTheme = defaultTheme;
  const main = new MainDocumentPart();
  main.setXml(`<w:document xmlns:w="${W}"><w:body>${bodyXml}</w:body></w:document>`);
  pkg.addTargetPart(main);
  if (stylesXml !== undefined) {
    const part = new StyleDefinitionsPart();
    part.setXml(stylesXml);
    main.addTargetPart(part);
  }
  if (settingsXml !== undefined) {
    const part = new DocumentSettingsPart();
    part.setXml(settingsXml);
    main.addTargetPart(part);
  }
  if (fontSchemeXml !== undefined) {
    const part = new ThemePart();
    part.setXml(themeXml(fontSchemeXml));
    main.addTargetPart(part);
  }
  if (numberingXml !== undefined) {
    const part = new NumberingDefinitionsPart();
    part.setXml(numberingXml);
    main.addTargetPart(part);
  }
  if (fontTableXml !== undefined) {
    const part = new FontTablePart();
    part.setXml(fontTableXml);
    main.addTargetPart(part);
  }
  if (headerXml !== undefined) {
    const part = new HeaderPart();
    part.setXml(headerXml);
    main.addTargetPart(part);
  }
  await pkg.getPropertyResolver();
  await main.getRunFontSelector();
  return pkg;
}

/** The document fonts of paragraph i's first run, span by span. */
async function fontsOf(pkg, i) {
  const main = pkg.getMainDocumentPart();
  const selector = await main.getRunFontSelector();
  const paragraph = (await main.getContents()).body.content[i].value;
  const run = paragraph.content[0].value;
  const text = run.content[0].value.value;
  return selector.spans(paragraph.pPr, run.rPr, text).map((s) => s.documentFont);
}

const fontTable = (entries) => `<w:fonts xmlns:w="${W}">${entries}</w:fonts>`;

async function unmarshalFontTable(entries) {
  const part = new FontTablePart();
  part.setXml(fontTable(entries));
  return part.getContents();
}

// ------------------------------------------------------------- phase 1: one resolution

test('w:cs and w:rtl are values, not flags (RunFontSelectorCsValueTest)', async () => {
  const TEXT = 'The quick brown fox';
  const pkg = await packageWith({
    stylesXml: styles(
      '<w:rFonts w:ascii="FontA" w:hAnsi="FontA" w:cs="FontB" w:eastAsia="FontA"/><w:sz w:val="22"/>',
      '<w:style w:type="character" w:styleId="CsOn"><w:name w:val="CsOn"/>'
      + '<w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:cs/></w:rPr></w:style>'),
    bodyXml: p('<w:cs/>', TEXT)                                       // 0: the cs font
      + p('<w:cs w:val="0"/>', TEXT)                                  // 1: off
      + p('<w:rStyle w:val="CsOn"/><w:cs w:val="0"/>', TEXT)          // 2: the style's cs overridden
      + p('<w:rtl w:val="0"/>', TEXT)                                 // 3: off
      + p('<w:rtl/>', TEXT)                                           // 4: the cs font
      + p('<w:rStyle w:val="CsOn"/>', TEXT)                           // 5: the style's cs applies
      + p(null, TEXT),                                                // 6: nothing: the ascii font
  });
  assert.deepEqual(await fontsOf(pkg, 0), ['FontB'], 'w:cs');
  assert.deepEqual(await fontsOf(pkg, 1), ['FontA'], 'w:cs w:val=0');
  assert.deepEqual(await fontsOf(pkg, 2), ['FontA'], 'style cs, direct w:val=0');
  assert.deepEqual(await fontsOf(pkg, 3), ['FontA'], 'w:rtl w:val=0');
  assert.deepEqual(await fontsOf(pkg, 4), ['FontB'], 'w:rtl');
  assert.deepEqual(await fontsOf(pkg, 5), ['FontB'], 'style cs');
  assert.deepEqual(await fontsOf(pkg, 6), ['FontA'], 'nothing');
  // the flags the spans report are the effective properties', by value
  const main = pkg.getMainDocumentPart();
  const selector = await main.getRunFontSelector();
  const body = (await main.getContents()).body.content;
  const spanOf = (i) => selector.spans(body[i].value.pPr, body[i].value.content[0].value.rPr, 'x')[0];
  assert.equal(spanOf(0).cs, true);
  assert.equal(spanOf(1).cs, false);
  assert.equal(spanOf(4).rtl, true);
  assert.equal(spanOf(3).rtl, false);
});

test('the theme language is an exact subtag, so Estonian is not Ethiopic (RunFontSelectorThemeLangTest)', async () => {
  const SCRIPTS = '<a:font script="Ethi" typeface="FontEthiopic"/>'
    + '<a:font script="Jpan" typeface="FontJapanese"/>';
  const DEFAULTS = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" '
    + 'w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/><w:sz w:val="22"/>';

  const estonian = await packageWith({
    stylesXml: styles(DEFAULTS, null),
    settingsXml: settings('w:val="et-EE"'),
    fontSchemeXml: fontScheme('FontMajor', 'FontMinor', SCRIPTS),
    bodyXml: p(null, 'Tere') + p('<w:lang w:val="et-EE"/>', 'Tere'),
  });
  assert.deepEqual(await fontsOf(estonian, 0), ['FontMinor']);
  assert.deepEqual(await fontsOf(estonian, 1), ['FontMinor']);
  assert.equal((await estonian.getMainDocumentPart().getRunFontSelector()).defaultFont, 'FontMinor');

  // ja-JP as the Latin language selects the theme's Jpan face for minorHAnsi, and the default
  // font must agree with what the runs get (docx4j computed it before reading the language)
  const japanese = await packageWith({
    stylesXml: styles(DEFAULTS, null),
    settingsXml: settings('w:val="ja-JP"'),
    fontSchemeXml: fontScheme('FontMajor', 'FontMinor', SCRIPTS),
    bodyXml: p(null, 'Hello'),
  });
  assert.deepEqual(await fontsOf(japanese, 0), ['FontJapanese']);
  assert.equal((await japanese.getMainDocumentPart().getRunFontSelector()).defaultFont, 'FontJapanese');
});

test('no w:rFonts anywhere: Times New Roman, through the dispatch (RunFontSelectorNoRFontsTest)', async () => {
  const pkg = await packageWith({
    stylesXml: styles('<w:sz w:val="22"/>', null),
    bodyXml: p(null, 'Hello') + p(null, 'Hello 日本 world'),
  });
  assert.deepEqual(await fontsOf(pkg, 0), ['Times New Roman']);
  // the East Asian slot is deliberately left unnamed, so the preamble rule does not fire and
  // the text is still dispatched by range; every slot answers the same default here
  const mixed = await fontsOf(pkg, 1);
  assert.equal(mixed[0], 'Times New Roman');
  assert.equal((await pkg.getMainDocumentPart().getRunFontSelector()).defaultFont, 'Times New Roman');
});

test('a theme reference with no theme part resolves to the Office theme, not the explicit name', async () => {
  // docx4j CR-016 probe fonts-missing-slots (b): Word supplies its own theme to a themeless
  // document and the explicit w:ascii beside the reference is unused
  const pkg = await packageWith({
    stylesXml: styles('<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/>', null),
    bodyXml: p('<w:rFonts w:ascii="Decoy" w:asciiTheme="minorHAnsi"/>', 'Hello'),
  });
  assert.deepEqual(await fontsOf(pkg, 0), ['Aptos']);
});

// ---------------------------------------------------------- phase 2: the range dispatch

test('the character-range table (RunFontSelectorDispatchTest)', async () => {
  const pkg = await packageWith({ stylesXml: styles('<w:sz w:val="22"/>', null), bodyXml: p(null, 'x') });
  const rfs = await pkg.getMainDocumentPart().getRunFontSelector();
  const SLOTS = ['ea', 'ascii', 'hAnsi', 'cs'];
  const NO_EA = [undefined, 'ascii', 'hAnsi', 'cs'];
  const at = (cp, hint, lang, slots = SLOTS) => rfs.fontFor(cp, hint, lang, ...slots);

  // the space is ASCII: Word draws it in the ascii font between two East Asian words
  assert.equal(at(0x20), 'ascii');
  assert.equal(at(0x65E5), 'ea');            // CJK
  assert.equal(at(0x32), 'ascii');           // a digit
  assert.equal(at(0x3001), 'ea');            // the ideographic comma

  // Hebrew and Arabic without w:cs: the ascii font, not Times New Roman
  assert.equal(at(0x05E9), 'ascii');
  assert.equal(at(0x0645), 'ascii');
  // the Indic ranges: cs (docx4j issues 622, 666); hAnsi where the run names none
  assert.equal(at(0x0915), 'cs');
  assert.equal(rfs.fontFor(0x0915, undefined, undefined, 'ea', 'ascii', 'hAnsi', undefined), 'hAnsi');
  assert.equal(at(0x0E01), 'cs');            // Thai
  assert.equal(at(0x1780), 'cs');            // Khmer

  // hint=eastAsia with no East Asian font is hAnsi in every branch, never nothing
  assert.equal(at(0x0430, 'eastAsia'), 'ea');
  assert.equal(at(0x0430, 'eastAsia', undefined, NO_EA), 'hAnsi');
  assert.equal(at(0x0430), 'hAnsi');
  assert.equal(at(0x1100, undefined, undefined, NO_EA), 'hAnsi');  // Hangul Jamo
  assert.equal(at(0x65E5, undefined, undefined, NO_EA), 'hAnsi');  // CJK
  assert.equal(at(0xFF1B, undefined, undefined, NO_EA), 'hAnsi');  // fullwidth

  // the Latin-1 exceptions under the hint
  assert.equal(at(0x00A7), 'hAnsi');
  assert.equal(at(0x00A7, 'eastAsia'), 'ea');
  assert.equal(at(0x00E9, 'eastAsia'), 'hAnsi');
  assert.equal(at(0x00E9, 'eastAsia', 'zh-CN'), 'ea');
  assert.equal(at(0x00A7, 'eastAsia', undefined, NO_EA), 'hAnsi');

  // Georgian is unlisted, so hAnsi; the symbol blocks take hAnsi here (docx4j names Segoe UI
  // Symbol only where it has that face - a glyph check this CR does not port)
  assert.equal(at(0x10D2), 'hAnsi');
  assert.equal(at(0x2192), 'hAnsi');
  // an astral code point: hAnsi, unless an emoji font is configured
  assert.equal(at(0x1F600), 'hAnsi');
  const withEmoji = new RunFontSelector({ emojiFont: 'Segoe UI Emoji' });
  assert.equal(withEmoji.fontFor(0x1F600, undefined, undefined, 'ea', 'ascii', 'hAnsi', 'cs'), 'Segoe UI Emoji');
  assert.equal(withEmoji.fontFor(0x20001, undefined, undefined, 'ea', 'ascii', 'hAnsi', 'cs'), 'hAnsi');
});

test('the ASCII letters return to the ascii font after a Latin-1 one', async () => {
  // docx4j CR-016 probe fonts-missing-slots (c): "caf" ascii, "e-acute" hAnsi, "ber" ascii again
  const pkg = await packageWith({
    stylesXml: styles('<w:rFonts w:ascii="FontA" w:hAnsi="FontH"/><w:sz w:val="22"/>', null),
    bodyXml: p(null, 'café über'),
  });
  assert.deepEqual(await fontsOf(pkg, 0), ['FontA', 'FontH', 'FontA', 'FontH', 'FontA']);
});

test('spans join by font across ranges, and cut where the font changes', async () => {
  const pkg = await packageWith({
    stylesXml: styles('<w:rFonts w:ascii="FontA" w:hAnsi="FontH" w:eastAsia="FontE" w:cs="FontC"/><w:sz w:val="22"/>', null),
    bodyXml: p(null, 'გაда')            // 0: Georgian then Cyrillic: both hAnsi
      + p(null, '日本 日本')                // 1: CJK, an ascii space, CJK
      + p(null, 'naïve café, 12'),                // 2: ascii with two hAnsi accents
  });
  // Georgian (unlisted) and Cyrillic (listed) both take hAnsi: one span, since spans join by font
  assert.deepEqual(await fontsOf(pkg, 0), ['FontH']);
  // the space keeps the ascii font, so the CJK words are cut apart
  assert.deepEqual(await fontsOf(pkg, 1), ['FontE', 'FontA', 'FontE']);
  assert.deepEqual(await fontsOf(pkg, 2), ['FontA', 'FontH', 'FontA', 'FontH', 'FontA']);
  // and the scripts a span reports
  const rfs = await pkg.getMainDocumentPart().getRunFontSelector();
  assert.deepEqual(rfs.spans(undefined, { rFonts: { ascii: 'A', hAnsi: 'A' } }, 'a日',
    { rPrIsEffective: true }).map((s) => s.script),
    ['LATIN']); // one font, one span; its script is the first non-shared character's
  assert.equal(spanScript(0x20), null, 'a space is shared');
  assert.equal(spanScript(0x65E5), 'CJK');
  assert.equal(spanScript(0x3042), 'CJK', 'hiragana groups with kanji');
  assert.equal(spanScript(0x10D2), 'GEORGIAN');
  assert.equal(spanScript(0x2192), 'SYMBOL');
  assert.equal(spanScript(0x1F600), 'EMOJI');
});

test('the preamble rule: eastAsia Times New Roman with ascii = hAnsi is one span in ascii', async () => {
  const pkg = await packageWith({
    stylesXml: styles('<w:rFonts w:ascii="FontA" w:hAnsi="FontA" w:eastAsia="Times New Roman"/><w:sz w:val="22"/>', null),
    bodyXml: p(null, 'abc 日本'),
  });
  assert.deepEqual(await fontsOf(pkg, 0), ['FontA']);
});

test('symbol fonts by case-insensitive name, and emoji by code point', async () => {
  assert.equal(symbolFontName('symbol'), 'Symbol');
  assert.equal(symbolFontName('WINGDINGS 2'), 'Wingdings 2');
  assert.equal(symbolFontName('Webdings'), 'Webdings');
  assert.equal(symbolFontName('Calibri'), undefined);
  assert.equal(symbolFontName(undefined), undefined);

  assert.equal(isEmoji(0x1F600), true);
  assert.equal(isEmoji(0x2751), false);
  assert.equal(coverageGroupOf(0x1F600), 'EMOJI');
  assert.equal(coverageGroupOf(0x2751), 'SYMBOL');
  assert.equal(coverageGroupOf(0x20), 'COMMON');
  assert.equal(isSymbol(0x2192), true);

  // a run whose hAnsi is a symbol font is set in it whole, whatever the characters
  const pkg = await packageWith({
    stylesXml: styles('<w:rFonts w:ascii="FontA" w:hAnsi="FontA"/><w:sz w:val="22"/>', null),
    bodyXml: p('<w:rFonts w:ascii="symbol" w:hAnsi="symbol"/>', 'abg 日'),
  });
  assert.deepEqual(await fontsOf(pkg, 0), ['Symbol']);
});

test('LanguageTagToScriptMappingTest', () => {
  assert.equal(getScriptForLanguageTag('ja-JP'), 'Jpan');
  assert.equal(getScriptForLanguageTag('ko-KR'), 'Hang');
  assert.equal(getScriptForLanguageTag('zh-CN'), 'Hans');
  assert.equal(getScriptForLanguageTag('zh-SG'), 'Hans');
  assert.equal(getScriptForLanguageTag('zh-TW'), 'Hant');
  assert.equal(getScriptForLanguageTag('he-IL'), 'Hebr');
  assert.equal(getScriptForLanguageTag('ti-ET'), 'Ethi');
  assert.equal(getScriptForLanguageTag('bn-IN'), 'Beng');
  assert.equal(getScriptForLanguageTag('hi-IN'), 'Deva');
  assert.equal(getScriptForLanguageTag('kok'), 'Deva');
  assert.equal(getScriptForLanguageTag('vi-VN'), 'Viet');
  assert.equal(getScriptForLanguageTag('ka-GE'), 'Geor');
  // and no substring matches
  assert.equal(getScriptForLanguageTag('et-EE'), undefined, 'Estonian is not Ethiopic');
  assert.equal(getScriptForLanguageTag('mn-MN'), undefined, 'Mongolian is not Bengali');
  assert.equal(getScriptForLanguageTag('wo-SN'), undefined, 'Wolof is not Ethiopic');
  assert.equal(getScriptForLanguageTag('en-US'), undefined);
  assert.equal(getScriptForLanguageTag('es-ES'), undefined);
  assert.equal(getScriptForLanguageTag('i'), undefined);
  assert.equal(getScriptForLanguageTag(undefined), undefined);
});

// --------------------------------------------------------------- phase 3: the mapping

/** A registry of our own, so a mapping assertion does not depend on the default one. */
function registry(...names) {
  return new SimpleFontRegistry(names.map((n) => new PhysicalFont(n)));
}

const SANS = 'Liberation Sans';
const SERIF = 'Liberation Serif';

test('MapperPrecedenceTest: the installed font wins over the embedded one', async () => {
  const reg = registry(SANS, SERIF);
  const m = new IdentityPlusMapper(reg);
  // the document embeds a "Liberation Sans" which is really the Serif file
  m.registerRegularForm(SANS, reg.get(SERIF));
  m.populateFontMappings([SANS], undefined);
  assert.equal(m.get(SANS), reg.get(SANS));
  assert.equal(m.getDecision(SANS).source, 'INSTALLED');
});

test('MapperPrecedenceTest: the embedded form where the environment lacks the font', async () => {
  const reg = registry(SANS, SERIF);
  const m = new IdentityPlusMapper(reg);
  m.registerBoldForm('Docx4j Embedded Only', reg.get(SERIF));
  m.populateFontMappings(['Docx4j Embedded Only'], undefined);
  assert.equal(m.get('Docx4j Embedded Only'), reg.get(SERIF));
  assert.equal(m.getDecision('Docx4j Embedded Only').source, 'EMBEDDED');
});

test('MapperPrecedenceTest: the identity face order is regular, bold, italic', async () => {
  // a family with no plain face: bold before italic (with italic first, upright text came out italic)
  const reg = registry('Docx4j Faceorder bold', 'Docx4j Faceorder italic');
  const m = new IdentityPlusMapper(reg);
  m.populateFontMappings(['Docx4j Faceorder'], undefined);
  assert.equal(m.get('Docx4j Faceorder').name, 'Docx4j Faceorder bold');
  assert.equal(m.getDecision('Docx4j Faceorder').source, 'MAPPER_OWN');
  assert.equal(m.getDecision('Docx4j Faceorder').via,
    'a variant of the name: Docx4j Faceorder bold');
});

test('MapperPrecedenceTest: the altName chain, and a cycle', async () => {
  const reg = registry(SANS, SERIF);
  const m = new IdentityPlusMapper(reg);
  const table = await unmarshalFontTable(
    '<w:font w:name="Docx4j Alt A"><w:altName w:val="Docx4j Alt B"/></w:font>'
    + `<w:font w:name="Docx4j Alt B"><w:altName w:val="${SERIF}"/></w:font>`
    + '<w:font w:name="Docx4j Alt Loop"><w:altName w:val="Docx4j Alt Loop"/></w:font>');
  const names = ['Docx4j Alt A', 'Docx4j Alt Loop'];
  m.populateFontMappings(names, table);
  m.addAltNameSubstitutes(names, table);
  assert.equal(m.get('Docx4j Alt A'), reg.get(SERIF), 'two hops');
  assert.equal(m.getDecision('Docx4j Alt A').via, `w:altName Docx4j Alt B -> ${SERIF}`);
  assert.equal(m.get('Docx4j Alt Loop'), undefined, 'a cycle resolves to nothing');
});

test("MapperPrecedenceTest: Word's default for an unknown family", async () => {
  const table = await unmarshalFontTable(
    '<w:font w:name="Docx4j Probe B"><w:panose1 w:val="020B0604020202020204"/><w:charset w:val="00"/><w:family w:val="swiss"/></w:font>'
    + '<w:font w:name="Docx4j Probe C"><w:charset w:val="00"/><w:family w:val="roman"/></w:font>'
    + '<w:font w:name="Docx4j Probe E"><w:altName w:val="Docx4j Probe X"/><w:charset w:val="00"/></w:font>'
    + '<w:font w:name="Docx4j Probe F"><w:altName w:val="Docx4j Probe Y"/></w:font>'
    + '<w:font w:name="Docx4j Probe Y"><w:family w:val="swiss"/></w:font>'
    + '<w:font w:name="Docx4j Probe M"><w:family w:val="modern"/></w:font>');
  const t = fontTableOf(table);
  // the golden fonts-unresolvable: (a) no entry Cambria, (b) swiss Calibri (panose unread),
  // (c) roman Cambria, (e) an entry without a family Calibri, (f) the altName's family
  assert.equal(wordDefaultFor('Docx4j Probe A', t), 'Cambria');
  assert.equal(wordDefaultFor('Docx4j Probe B', t), 'Calibri');
  assert.equal(wordDefaultFor('Docx4j Probe C', t), 'Cambria');
  assert.equal(wordDefaultFor('Docx4j Probe E', t), 'Calibri');
  assert.equal(wordDefaultFor('Docx4j Probe F', t), 'Calibri');
  assert.equal(wordDefaultFor('Docx4j Probe M', t), 'Courier New');

  assert.equal(isKnownFamily('Arial'), true);
  assert.equal(isKnownFamily('Calibri Light'), true);
  assert.equal(isKnownFamily('Trebuchet MS'), true);
  assert.equal(isKnownFamily('Docx4j Probe B'), false);

  const reg = registry(SANS, SERIF);
  const m = new IdentityPlusMapper(reg);
  m.put('Calibri', reg.get(SANS));
  m.put('Cambria', reg.get(SERIF));
  const names = ['Docx4j Probe A', 'Docx4j Probe B', 'Docx4j Probe C', 'Docx4j Probe E', 'Arial Narrow'];
  m.addWordDefaultSubstitutes(names, table);
  assert.equal(m.get('Docx4j Probe A'), reg.get(SERIF));
  assert.equal(m.get('Docx4j Probe B'), reg.get(SANS));
  assert.equal(m.get('Docx4j Probe C'), reg.get(SERIF));
  assert.equal(m.get('Docx4j Probe E'), reg.get(SANS));
  assert.equal(m.get('Arial Narrow'), undefined, 'a known family is not Word-defaulted');
  // and the line box is the face Word uses
  assert.equal(m.lineMetricsFamily('Docx4j Probe A'), 'Cambria');
});

test('MapperPrecedenceTest: a name whose class is only a guess is Word-defaulted', async () => {
  const reg = registry(SANS, SERIF);
  assert.equal(selectByClass(reg, 'Docx4j Probe Sans Light'), undefined,
    'the class pass has nothing for such a name');
  assert.equal(isKnownFamily('Docx4j Probe Sans Light'), false,
    'so this pass must not stand back from it');

  const table = await unmarshalFontTable('<w:font w:name="Docx4j Probe Sans Light">'
    + '<w:panose1 w:val="020B0300040303060204"/><w:family w:val="swiss"/>'
    + '<w:pitch w:val="variable"/></w:font>');
  const names = ['Docx4j Probe Sans Light'];

  const m = new IdentityPlusMapper(reg);
  m.put('Calibri', reg.get(SANS));
  m.addClassBasedSubstitutes(names);
  assert.equal(m.get('Docx4j Probe Sans Light'), undefined, 'nothing from the class pass');

  m.addWordDefaultSubstitutes(names, table);
  assert.equal(m.get('Docx4j Probe Sans Light'), reg.get(SANS), 'w:family swiss is Calibri');
  assert.equal(m.lineMetricsFamily('Docx4j Probe Sans Light'), 'Calibri');

  // the families deliberately left to the document default are still left there
  assert.equal(isKnownFamily('Docx4j Probe Sans Narrow'), true, 'a condensed family');
  assert.equal(isKnownFamily('Docx4jProbeSans-Light'), true, 'a PostScript name');
});

test('NoBoldFaceTest: a family with no bold face of its own', async () => {
  assert.equal(hasBoldFace('Calibri'), true);
  assert.equal(hasBoldFace('Arial'), true);
  assert.equal(hasBoldFace('Calibri Light'), false, 'a weight-named family');
  assert.equal(hasBoldFace('Segoe UI Semibold'), false);
  assert.equal(hasBoldFace('Arial Black'), false, 'the registry entry has no bold child');
  assert.equal(hasBoldFace('Some Unknown Family'), true);

  const reg = registry(SANS, SERIF);
  const m = new IdentityPlusMapper(reg);
  m.put('Calibri Light', reg.get(SANS));
  m.put('Calibri', reg.get(SANS));
  m.addNoBoldFaceAliases(['Calibri Light', 'Calibri']);
  const alias = m.get('Calibri Light');
  assert.equal(alias.noBoldFace, true);
  assert.equal(alias.name, SANS + NOBOLD_SUFFIX);
  assert.equal(alias.familyName, SANS, 'the alias keeps its file\'s family');
  // the alias's name resolves to the file's own entry, and the FO layer's suffixes stack
  assert.equal(reg.get(alias.name), reg.get(SANS));
  assert.equal(reg.get(`${alias.name}+noliga`), reg.get(SANS));
  assert.equal(reg.get(`${alias.name}+kerned+noliga`), reg.get(SANS));
  assert.equal(m.get('Calibri'), reg.get(SANS), 'Calibri keeps its real bold');
  assert.equal(m.get('Calibri').noBoldFace, false);

  // a font Word itself could not find is substituted whole, bold face included
  const wordDefaulted = new IdentityPlusMapper(reg);
  wordDefaulted.put('Calibri', reg.get(SANS));
  const table = await unmarshalFontTable('<w:font w:name="Docx4j DIN Pro Light"><w:family w:val="swiss"/></w:font>');
  wordDefaulted.addWordDefaultSubstitutes(['Docx4j DIN Pro Light'], table);
  wordDefaulted.addNoBoldFaceAliases(['Docx4j DIN Pro Light']);
  assert.equal(wordDefaulted.get('Docx4j DIN Pro Light').noBoldFace, false);
});

test('MetricallyCompatibleSubstituteTest: serifs get a serif and sans a sans', async () => {
  const m = new IdentityPlusMapper(DEFAULT_FONT_REGISTRY);
  m.addMetricallyCompatibleSubstitutes();
  const substitute = (font) => m.get(font)?.name.toLowerCase();
  const assertOneOf = (font, ...anyOf) => {
    const name = substitute(font);
    assert.ok(name !== undefined, `${font} left unmapped`);
    assert.ok(anyOf.some((c) => name.includes(c)),
      `${font} mapped to ${name}, expected one of ${anyOf.join(', ')}`);
  };
  assertOneOf('Times New Roman', 'tinos', 'liberation serif');
  assertOneOf('Georgia', 'gelasio', 'p052', 'tinos', 'liberation serif');
  assertOneOf('Book Antiqua', 'p052', 'tinos', 'liberation serif');
  assertOneOf('Garamond', 'tinos', 'liberation serif');
  assertOneOf('Arial', 'arimo', 'liberation sans');
  assertOneOf('Tahoma', 'arimo', 'liberation sans');
  assertOneOf('Verdana', 'dejavu sans', 'arimo', 'liberation sans');
  assertOneOf('Comic Sans MS', 'noto sans', 'dejavu sans', 'arimo', 'liberation sans');
  assertOneOf('Segoe UI', 'selawik', 'arimo', 'liberation sans');
  assertOneOf('Helvetica', 'arimo', 'liberation sans');
  assertOneOf('Trebuchet MS', 'droid sans', 'arimo', 'liberation sans');
  assertOneOf('Arial Black', 'noto sans black', 'arimo', 'liberation sans');
  assertOneOf('Tw Cen MT', 'arimo', 'liberation sans');
  assertOneOf('Calibri Light', 'carlito', 'liberation sans');
  assertOneOf('Calibri', 'carlito', 'liberation sans');
  assertOneOf('Cambria', 'caladea', 'liberation serif');
  assertOneOf('Aptos', 'akasia', 'carlito');
  assertOneOf('Aptos Display', 'intos display', 'carlito');

  // the pair whose secondary substitutes were once swapped: neither becomes the other's class
  assert.ok(!substitute('Times New Roman').includes('sans'), 'Times New Roman became a sans');
  assert.ok(!substitute('Arial').includes('serif'), 'Arial became a serif');
  // and the sources the table's quality words give
  assert.equal(m.getDecision('Times New Roman').source, 'METRIC_CLONE');
  assert.equal(m.getDecision('Cambria').source, 'MEASURED_STAND_IN');
  assert.equal(m.getDecision('Consolas').source, 'CLASS');
});

test('ClassBasedSubstituteTest: a face of the same class, and the families left alone', async () => {
  const classOfSubstitute = (font) => {
    const m = new IdentityPlusMapper(DEFAULT_FONT_REGISTRY);
    m.addClassBasedSubstitutes([font]);
    const pf = m.get(font);
    assert.ok(pf !== undefined, `${font} was left unmapped`);
    return classOf(pf.name);
  };
  assert.equal(classOfSubstitute('Tahoma'), 'SANS');
  assert.equal(classOfSubstitute('Century Gothic'), 'SANS');
  assert.equal(classOfSubstitute('HelveticaNeue LT 55 Roman'), 'SANS');
  assert.equal(classOfSubstitute('Georgia'), 'SERIF');
  assert.equal(classOfSubstitute('Calisto MT'), 'SERIF');
  assert.equal(classOfSubstitute('Baskerville Old Face'), 'SERIF');
  assert.equal(classOfSubstitute('Consolas'), 'MONO');
  assert.equal(classOfSubstitute('Andale Mono'), 'MONO');

  // a condensed face is deliberately left unmapped, and so is Lato and a PostScript name
  const left = new IdentityPlusMapper(DEFAULT_FONT_REGISTRY);
  left.addClassBasedSubstitutes(['Arial Narrow', 'Lato', 'MyriadPro-Regular']);
  assert.equal(left.get('Arial Narrow'), undefined);
  assert.equal(left.get('Lato'), undefined);
  assert.equal(left.get('MyriadPro-Regular'), undefined);
  // substitutionClass does not guess from a trailing "Sans", where classOf does
  assert.equal(substitutionClass('Docx4j Probe Sans'), 'UNKNOWN');
  assert.equal(classOf('Docx4j Probe Sans'), 'SANS');
});

// -------------------------------------------------------------- phase 4: fonts in use

test('FontsInUseTest: every slot of everything in use', async () => {
  const docDefaults = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi" w:cs=""/>';
  const moreStyles =
    '<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:rPr><w:rFonts w:ascii="Base Font"/></w:rPr></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Base"/><w:rPr><w:rFonts w:cs="Quote CS Font"/></w:rPr></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Unused"><w:name w:val="Unused"/><w:rPr><w:rFonts w:ascii="Unused Font"/></w:rPr></w:style>'
    + '<w:style w:type="character" w:styleId="Emph"><w:name w:val="Emph"/><w:rPr><w:rFonts w:ascii="Char Style Font"/></w:rPr></w:style>';
  const body =
    // a paragraph mark's font, a paragraph style with a base, all four slots of a run
    '<w:p><w:pPr><w:pStyle w:val="Quote"/><w:rPr><w:rFonts w:ascii="Mark Font"/></w:rPr></w:pPr>'
    + '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="MS Mincho" w:cs="Arial Unicode MS"/></w:rPr><w:t>x</w:t></w:r></w:p>'
    // a character style, a CJK name, a w:sym
    + '<w:p><w:r><w:rPr><w:rStyle w:val="Emph"/><w:rFonts w:ascii="宋体"/></w:rPr><w:t>y</w:t></w:r>'
    + '<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r></w:p>'
    // a theme reference beside an explicit name: the theme's face is what is collected
    + '<w:p><w:r><w:rPr><w:rFonts w:ascii="Decoy Font" w:asciiTheme="minorHAnsi" w:cs=""/></w:rPr><w:t>z</w:t></w:r></w:p>';
  const parts = {
    stylesXml: styles(docDefaults, moreStyles),
    bodyXml: body,
    numberingXml: `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">`
      + '<w:rPr><w:rFonts w:ascii="Symbol" w:cs="Numbering CS Font"/></w:rPr></w:lvl></w:abstractNum></w:numbering>',
    headerXml: `<w:hdr xmlns:w="${W}"><w:p><w:r><w:rPr><w:rFonts w:ascii="Header Font"/></w:rPr><w:t>h</w:t></w:r></w:p></w:hdr>`,
  };

  const pkg = await packageWith({ ...parts, fontSchemeXml: fontScheme('Cambria', 'Calibri') });
  const fonts = await pkg.getMainDocumentPart().fontsInUse();
  for (const expected of [
    'Calibri',                                    // the theme's minor face, from the defaults and the run
    'Arial', 'MS Mincho', 'Arial Unicode MS',     // the run's three named slots
    'Mark Font',                                  // the paragraph mark's
    'Base Font', 'Quote CS Font',                 // the style in use and what it is based on
    'Char Style Font',                            // the character style in use
    'SimSun',                                     // the CJK name, in English
    'Wingdings',                                  // w:sym
    'Header Font',                                // a header part
    'Symbol', 'Numbering CS Font',                // a numbering level, both slots
  ]) {
    assert.ok(fonts.has(expected), `${expected} missing from ${[...fonts]}`);
  }
  assert.ok(!fonts.has('Unused Font'), 'a style the document never uses');
  assert.ok(!fonts.has('Decoy Font'), 'the explicit name beside a theme reference');
  assert.ok(!fonts.has('宋体'), 'the CJK name itself');
  assert.ok(!fonts.has(''), 'a blank name');
  assert.equal(toEnglish('宋体'), 'SimSun');

  // the styles in use, without what they are based on
  const stylesInUse = await pkg.getMainDocumentPart().getStylesInUse();
  assert.deepEqual([...stylesInUse].sort(), ['DefaultParagraphFont', 'Emph', 'Normal', 'Quote']);

  // no theme part: the Office theme's faces, which since docx4j 17.1.1 are Word 365's Aptos
  const themeless = await packageWith(parts);
  const withoutTheme = await themeless.getMainDocumentPart().fontsInUse();
  assert.ok(withoutTheme.has('Aptos'), [...withoutTheme].join(', '));
  assert.ok(!withoutTheme.has('Calibri'), 'the Office 2007 face is no longer the default');
});

// ------------------------------------------------------- the themeless default, and createPackage

test('the themeless default: 2023, 2013 and 2007', async () => {
  assert.equal(defaultThemeSetting().value, '2023', 'Word 365 is the default');
  assert.equal(defaultThemeOf('2013'), THEME_2013);
  assert.equal(defaultThemeOf('nope'), undefined);
  assert.deepEqual([THEME_2023.majorLatin, THEME_2023.minorLatin], ['Aptos Display', 'Aptos']);
  assert.deepEqual([THEME_2013.majorLatin, THEME_2013.minorLatin], ['Calibri Light', 'Calibri']);
  assert.deepEqual([THEME_2007.majorLatin, THEME_2007.minorLatin], ['Cambria', 'Calibri']);

  const themeReferences = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/>';
  for (const [value, minor, major] of [
    ['2023', 'Aptos', 'Aptos Display'],
    ['2013', 'Calibri', 'Calibri Light'],
    ['2007', 'Calibri', 'Cambria'],
  ]) {
    const pkg = await packageWith({
      defaultTheme: value,
      stylesXml: styles(themeReferences, null),
      bodyXml: p(null, 'Hello') + p('<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>', 'Heading'),
    });
    assert.deepEqual(await fontsOf(pkg, 0), [minor], `${value} minor`);
    assert.deepEqual(await fontsOf(pkg, 1), [major], `${value} major`);
    assert.equal((await pkg.getMainDocumentPart().getRunFontSelector()).defaultFont, minor);
    // only the two Latin slots: the Office theme's East Asian and complex-script entries are
    // empty, so those references resolve to nothing and an explicit attribute stands
    const rfs = await pkg.getMainDocumentPart().getRunFontSelector();
    assert.equal(rfs.themeFont('minorEastAsia'), undefined);
    assert.equal(rfs.themeFont('minorBidi'), undefined);
  }
});

test('createPackage adds a theme part whose faces are the default theme\'s', async () => {
  for (const [value, minor, major] of [
    [undefined, 'Aptos', 'Aptos Display'],
    ['2013', 'Calibri', 'Calibri Light'],
    ['2007', 'Calibri', 'Cambria'],
  ]) {
    const pkg = await WordprocessingMLPackage.createPackage(value === undefined ? {} : { defaultTheme: value });
    const main = pkg.getMainDocumentPart();
    assert.ok(main.themePart, 'a theme part');
    assert.equal(main.themePart.partName.name, '/word/theme/theme1.xml');
    assert.equal(pkg.fonts.defaultTheme, value ?? '2023');
    const selector = await main.getRunFontSelector();
    // resolved through the part, not the themeless fallback: the same faces either way
    assert.equal(selector.themeFont('minorHAnsi'), minor);
    assert.equal(selector.themeFont('majorHAnsi'), major);
    // the default styles reference minorHAnsi, so the document default follows
    assert.equal(selector.defaultFont, minor);
    // and a paragraph's font reads the theme's face rather than ''
    const body = await pkg.getBody();
    const paragraph = body.insertParagraph('Hello', 'End');
    assert.equal(paragraph.font.name, minor);
    assert.equal(paragraph.getFont({ direct: true }).name, '', 'the direct read is unchanged');
  }
});

test('a theme part with script fonts answers per script, as docx4j\'s ThemePart.getFont does', async () => {
  const pkg = await packageWith({
    stylesXml: styles('<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/><w:sz w:val="22"/>', null),
    settingsXml: settings('w:val="en-US" w:eastAsia="ja-JP"'),
    fontSchemeXml: fontScheme('FontMajor', 'FontMinor',
      '<a:font script="Jpan" typeface="FontJapanese"/><a:font script="Hans" typeface="FontSimplified"/>'),
    bodyXml: p(null, 'x'),
  });
  const rfs = await pkg.getMainDocumentPart().getRunFontSelector();
  assert.equal(rfs.themeFont('minorHAnsi'), 'FontMinor', 'en-US selects no script');
  assert.equal(rfs.themeFont('minorEastAsia'), 'FontJapanese', 'the eastAsia language does');
  assert.equal(rfs.themeFont('majorHAnsi'), 'FontMajor');
  assert.equal(rfs.themeFont(undefined), undefined);
});
