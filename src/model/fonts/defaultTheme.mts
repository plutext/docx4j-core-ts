// Which Office theme a package that has **no theme part** is read as having, and which theme
// part `WordprocessingMLPackage.createPackage` puts in a new one.
//
// docx4j `org.docx4j.Docx4jProperties.DefaultTheme` and the property
// `docx4j.fonts.defaultTheme` (docx4j CR-001 batch 49 item 6).  Word supplies its own default
// theme to a themeless document, and which faces that names has changed twice: measured on a
// Word 365 rendering of a package with no theme part (probe `theme-fonts-no-theme-part`), the
// body face is **Aptos** and the heading face **Aptos Display**, where Office 2013 to 2022
// answered Calibri and Calibri Light and Office 2007 to 2010 answered Calibri and Cambria.  So
// this says which version of Word a themeless document is read as having come from - not a
// preference - and the default is the current one.
//
// **The shape here.**  docx4j has one JVM-wide property; this package has no properties file,
// and a browser or an add-in may hold several packages at once, so the setting lives on the
// package: `pkg.fonts.defaultTheme = '2013'`.  {@link defaultThemeSetting} is the fallback for
// a selector built without a package, and is what a caller changes to move every package at
// once.
import { THEME_2007_XML, THEME_2013_XML, THEME_2023_XML } from './themes.generated.mjs';

/** The values `docx4j.fonts.defaultTheme` takes. */
export type DefaultThemeValue = '2023' | '2013' | '2007';

/** A default theme: the Latin faces its font scheme names, and the theme part docx4j bundles. */
export interface DefaultTheme {
  readonly value: DefaultThemeValue;
  /** The face its font scheme names for the major (heading) Latin slot. */
  readonly majorLatin: string;
  /** The face its font scheme names for the minor (body) Latin slot. */
  readonly minorLatin: string;
  /** That theme's own `theme1.xml`, as docx4j ships it. */
  readonly xml: string;
}

/** Word 365's Office theme: Aptos Display / Aptos.  The default. */
export const THEME_2023: DefaultTheme = {
  value: '2023', majorLatin: 'Aptos Display', minorLatin: 'Aptos', xml: THEME_2023_XML,
};
/** Office 2013 to 2022: Calibri Light / Calibri. */
export const THEME_2013: DefaultTheme = {
  value: '2013', majorLatin: 'Calibri Light', minorLatin: 'Calibri', xml: THEME_2013_XML,
};
/** Office 2007 to 2010: Cambria / Calibri; docx4j's answer up to 17.1.0. */
export const THEME_2007: DefaultTheme = {
  value: '2007', majorLatin: 'Cambria', minorLatin: 'Calibri', xml: THEME_2007_XML,
};

export const DEFAULT_THEMES: readonly DefaultTheme[] = [THEME_2023, THEME_2013, THEME_2007];

/** The theme a value names; undefined where it names none. */
export function defaultThemeOf(value: string | undefined): DefaultTheme | undefined {
  if (value === undefined || value === null) return undefined;
  const v = value.trim();
  return DEFAULT_THEMES.find((t) => t.value.toLowerCase() === v.toLowerCase());
}

let setting: DefaultTheme = THEME_2023;

/**
 * The package-level default, for a selector or a `createPackage()` that is not told otherwise.
 * `'2023'` unless set; an unknown value throws rather than being silently ignored.
 */
export function defaultThemeSetting(): DefaultTheme;
export function defaultThemeSetting(value: DefaultThemeValue | DefaultTheme): DefaultTheme;
export function defaultThemeSetting(value?: DefaultThemeValue | DefaultTheme): DefaultTheme {
  if (value !== undefined) {
    const theme = typeof value === 'string' ? defaultThemeOf(value) : value;
    if (!theme) throw new Error(`docx4j.fonts.defaultTheme: not one of 2023, 2013, 2007: ${String(value)}`);
    setting = theme;
  }
  return setting;
}

/** What a package carries: `pkg.fonts.defaultTheme`. */
export interface FontSettings {
  /** Which Office theme this package's themeless font references resolve against, and which
   *  theme part `createPackage()` gives it.  Defaults to {@link defaultThemeSetting}. */
  defaultTheme: DefaultThemeValue;
}

/** A package's font settings, defaulting to the process-wide setting when they are created. */
export function newFontSettings(): FontSettings {
  return { defaultTheme: defaultThemeSetting().value };
}

/** The theme a package's settings name, falling back to the process-wide setting. */
export function themeOfSettings(settings: FontSettings | undefined): DefaultTheme {
  return (settings === undefined ? undefined : defaultThemeOf(settings.defaultTheme)) ?? defaultThemeSetting();
}
