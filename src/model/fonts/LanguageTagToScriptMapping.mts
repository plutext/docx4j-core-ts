// docx4j `org.docx4j.fonts.LanguageTagToScriptMapping`.
//
// Maps `w:themeFontLang` in the settings part - `<w:themeFontLang w:val="en-US"
// w:eastAsia="ko-KR"/>` - to the `script` attribute of an `a:font` in the theme part's font
// scheme (`<a:font script="Jpan" typeface="ＭＳ 明朝"/>`), so that `ThemePart.getFont` can pick
// the face for the document's language.
//
// **By exact subtag.**  Until docx4j 17.1.1 this was a substring test over comma-separated
// lists, so Estonian (`et`) matched inside `eth` and took the Ethiopic face (Nyala), Mongolian
// (`mn`) inside `mni` the Bengali one, and Wolof (`wo`) inside `bwo`.  CR-016 probe
// `fonts-theme-lang` measured Word: with `<w:themeFontLang w:val="et-EE"/>` and docDefaults
// naming `minorHAnsi`, Word sets the text in Calibri, the theme's Latin face, on every line.

/** The script a primary language subtag selects in a theme's font list. */
const SCRIPT_BY_LANGUAGE = new Map<string, string>();

function put(script: string, ...langs: string[]): void {
  for (const lang of langs) SCRIPT_BY_LANGUAGE.set(lang, script);
}

put('Jpan', 'ja');
put('Hang', 'ko');
put('Arab', 'ar');
put('Hebr', 'he', 'yi', 'iw');
put('Thai', 'th');
put('Ethi', 'ti', 'bwo', 'eth', 'kxh', 'mdy');
put('Beng', 'bn', 'as', 'mni');
put('Gujr', 'gu');
put('Khmr', 'km');
put('Knda', 'kn');
put('Guru', 'pa');
put('Cans', 'iu');
put('Cher', 'chr');
// Yiii (Microsoft Yi Baiti) omitted; see http://en.wikipedia.org/wiki/Yi_script
put('Tibt', 'bo');
put('Thaa', 'dv');
put('Deva', 'hi', 'ks', 'kok', 'mr', 'ne', 'sa', 'sd');
put('Telu', 'te');
put('Taml', 'ta');
put('Syrc', 'syr');
put('Orya', 'or');
put('Mlym', 'ml');
put('Laoo', 'lo');
put('Sinh', 'si');
// Mong (Mongolian Baiti) and Uigh (Microsoft Uighur) omitted, as in docx4j
put('Viet', 'vi', 'lha', 'nut');
put('Geor', 'ka');

/**
 * The theme's `script` for a language tag, or undefined where the tag selects none (and the
 * theme's `a:latin`, `a:ea` or `a:cs` answers instead).  Chinese is decided on the region as
 * well: mainland China and Singapore use simplified characters (`Hans`), everywhere else
 * traditional (`Hant`).
 */
export function getScriptForLanguageTag(langTag: string | undefined): string | undefined {
  if (langTag === undefined || langTag === null) return undefined;
  let lang = langTag.trim();
  let pos = lang.indexOf('-');
  if (pos < 0) pos = lang.indexOf('_');
  if (pos > -1) lang = lang.slice(0, pos);
  lang = lang.toLowerCase();

  if (lang === 'zh') {
    const tag = langTag.toLowerCase();
    return tag === 'zh-cn' || tag === 'zh-sg' ? 'Hans' : 'Hant';
  }
  return SCRIPT_BY_LANGUAGE.get(lang);
}
