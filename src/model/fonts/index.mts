// CR-001 Phase B step 4: fonts (docx4j `org.docx4j.fonts`).
//
// Two halves, as in docx4j:
//
// - {@link RunFontSelector} decides which *document* font formats each character of a run: the
//   effective `w:rFonts` with its theme references resolved, `w:cs` / `w:rtl` by value, the
//   symbol fonts, and the [MS-OI29500] 17.3.2.26 character-range table.  Its answer is
//   `FontSpan[]`, which is what a renderer, a text extractor or a measurement routine needs.
// - {@link Mapper} maps a document font name to a {@link PhysicalFont}.  This CR ships
//   {@link IdentityPlusMapper} over a {@link FontRegistry} the caller supplies; docx4j's
//   `BestMatchingMapper`, the glyph-coverage pass and font metrics need font files and are a
//   later CR, but the interface is fixed here so a consumer with `fontkit` can plug its own in.
//
// The theme a package with no theme part is read as having is {@link defaultThemeSetting} /
// `pkg.fonts.defaultTheme`, and `createPackage()` gives a new package that theme's part.
export { PhysicalFont, NOBOLD_SUFFIX, stripSuffixes } from './PhysicalFont.mjs';
export {
  type FontRegistry, SimpleFontRegistry, DEFAULT_FONT_REGISTRY, EMPTY_FONT_REGISTRY,
} from './registry.mjs';
export {
  Mapper, type FontDecision, type FontDecisionSource, UNKNOWN_ERROR, FONT_FALLBACK,
  isKnownFamily, hasBoldFace, isEastAsianFamily, wordDefaultFor, fontTableOf,
} from './Mapper.mjs';
export { IdentityPlusMapper } from './IdentityPlusMapper.mjs';
export {
  type FontClass, classOf, classFromName, substitutionClass, isCondensed,
  leftToTheDocumentDefault, selectByClass, shortKey, vclEntryFor,
  isSymbol, isEastAsianForm, coverageGroupOf, SYMBOL_GROUP, EMOJI_GROUP,
} from './FontFallback.mjs';
export { scriptOf } from './scripts.mjs';
export {
  RunFontSelector, type RunFontSelectorSource, type FontSpan, type SpansOptions,
  BUILT_IN_DEFAULT_FONT, isOn, symbolFontName, isEmoji, spanScript, preambleRule, runLocale,
} from './RunFontSelector.mjs';
export { themeFontOf, type ThemeFontLang, type ThemeSource } from './ThemeFonts.mjs';
export { getScriptForLanguageTag } from './LanguageTagToScriptMapping.mjs';
export { toEnglish } from './CJKToEnglish.mjs';
export {
  type DefaultTheme, type DefaultThemeValue, type FontSettings,
  THEME_2007, THEME_2013, THEME_2023, DEFAULT_THEMES,
  defaultThemeOf, defaultThemeSetting, newFontSettings, themeOfSettings,
} from './defaultTheme.mjs';
export { fontsInUse, walkStories, type FontsAndStyles, type StyleSource } from './fontsInUse.mjs';
export { runFontSelectorOf } from './lookup.mjs';
export {
  type Substitute, type SubstituteRow, type VclEntry, SUBSTITUTE_ROWS, VCL_ENTRIES,
} from './substitutions.generated.mjs';
