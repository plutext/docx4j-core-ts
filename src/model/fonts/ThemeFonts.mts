// The theme part's answer to a font reference: docx4j `ThemePart.getFont(STTheme, CTLanguage)`.
//
// A `w:rFonts` slot may name a *theme* font (`w:asciiTheme="minorHAnsi"`) instead of a face.
// The theme part's font scheme holds a major and a minor collection, each with a Latin, an East
// Asian and a complex-script face and a list of per-script faces; `w:themeFontLang` in the
// settings part says which language the document is in, and the script that language selects
// (`LanguageTagToScriptMapping`) picks from the list.
//
// Split out from the part class because the selector, font discovery and the content API all
// resolve references, and none of them should have to hold a `ThemePart`: what they need is the
// `a:fontScheme`.
import type * as dml from '@docx4j/generated-objects-ts/modules/org_docx4j_dml';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { getScriptForLanguageTag } from './LanguageTagToScriptMapping.mjs';

/** `w:themeFontLang`. */
export type ThemeFontLang = wml.CTLanguage;

/** A theme's font scheme, or a whole theme (`a:theme`), or nothing. */
export type ThemeSource = dml.Theme | dml.BaseStyles.FontScheme | undefined;

function fontScheme(theme: ThemeSource): dml.BaseStyles.FontScheme | undefined {
  if (theme === undefined) return undefined;
  if ('majorFont' in theme) return theme as dml.BaseStyles.FontScheme;
  return (theme as dml.Theme).themeElements?.fontScheme;
}

function isMajor(type: wml.STTheme): boolean {
  return type === 'majorAscii' || type === 'majorBidi' || type === 'majorEastAsia' || type === 'majorHAnsi';
}

/** `ThemePart.getLang`: which of `w:themeFontLang`'s three attributes this reference reads. */
function langFor(themeFontLang: ThemeFontLang, type: wml.STTheme): string | undefined {
  if (type === 'majorAscii' || type === 'minorAscii' || type === 'majorHAnsi' || type === 'minorHAnsi') {
    return themeFontLang.val ?? undefined;
  }
  if (type === 'majorBidi' || type === 'minorBidi') return themeFontLang.bidi ?? undefined;
  return themeFontLang.eastAsia ?? undefined;
}

/** The collection's own `a:latin` / `a:ea` / `a:cs` for this reference. */
function textFontFromTheme(scheme: dml.BaseStyles.FontScheme, type: wml.STTheme): dml.TextFont | undefined {
  const collection = isMajor(type) ? scheme.majorFont : scheme.minorFont;
  if (collection === undefined) return undefined;
  switch (type) {
    case 'majorEastAsia': case 'minorEastAsia': return collection.ea;
    case 'majorBidi': case 'minorBidi': return collection.cs;
    // majorAscii / minorAscii / majorHAnsi / minorHAnsi, and anything unknown: a:latin
    default: return collection.latin;
  }
}

/** The typeface of a collection's own face for this reference; undefined where it is absent or
 *  empty (`<a:ea typeface=""/>`, which the Office theme writes). */
function fontFromTheme(scheme: dml.BaseStyles.FontScheme, type: wml.STTheme): string | undefined {
  const typeface = textFontFromTheme(scheme, type)?.typeface;
  if (typeface === undefined || typeface.trim().length === 0) return undefined;
  return typeface;
}

/**
 * The face a theme reference names in this theme, for this `w:themeFontLang`.
 *
 * With no `w:themeFontLang` (or none that applies to this reference), the collection's own
 * `a:latin` / `a:ea` / `a:cs`.  With one, the language's script (`Jpan`, `Hans`, `Geor` ...)
 * picks from the collection's `a:font` list, falling back to the collection's own face where
 * the language selects no script or the list has no entry for it.
 *
 * @returns the typeface, or undefined where the theme answers nothing.
 */
export function themeFontOf(theme: ThemeSource, type: wml.STTheme | undefined,
  themeFontLang?: ThemeFontLang | undefined): string | undefined {

  if (type === undefined || type === null) return undefined;
  const scheme = fontScheme(theme);
  if (scheme === undefined) return undefined;
  if (themeFontLang === undefined || themeFontLang === null) return fontFromTheme(scheme, type);

  const lang = langFor(themeFontLang, type);
  if (lang === undefined) return fontFromTheme(scheme, type);
  const script = getScriptForLanguageTag(lang);
  if (script === undefined) return fontFromTheme(scheme, type);

  const collection = isMajor(type) ? scheme.majorFont : scheme.minorFont;
  const entry = collection?.font?.find((f) => f.script === script);
  // as docx4j: only a script the list does not carry falls back; an entry with an empty
  // typeface answers the empty string, and the caller's explicit attribute does not stand
  return entry === undefined ? fontFromTheme(scheme, type) : entry.typeface;
}
