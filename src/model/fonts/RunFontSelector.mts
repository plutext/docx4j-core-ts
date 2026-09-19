// docx4j `org.docx4j.fonts.RunFontSelector`: which of a run's fonts formats each character.
//
// **The selection core only.**  docx4j's class is 2,572 lines, most of them XSL-FO output:
// element creation, the kerned and no-ligature twins, character scaling, small caps, line
// metrics, and the glyph-coverage pass that replaces a face which lacks a script.  What CR-001
// section 6.3 promised, and what a browser, a text extractor or a measurer needs, is the
// decision per character, so what is ported is `documentFontsOf`, `defaultFontOf`, the theme
// lookup by `themeFontLang`, `resolvedSlots`, `complexScriptFont` (`w:cs` and `w:rtl` **by
// value**), `preambleRule`, the code-point-to-slot table `fontFor` verbatim, `spanScript`,
// `isEmoji`, `symbolFontName` and the `hint` rules.
//
// **One resolution** (docx4j CR-016 phase 1): the selector consumes the run's *effective*
// properties.  It resolves them itself through the `PropertyResolver` unless the caller says it
// has already done so (`rPrIsEffective`), which is what the parity harness and every walk over
// a whole document do.
//
// What the goldens settled, and is reproduced here: `<w:cs w:val="0"/>` turns an inherited
// complex-script flag **off** (until docx4j 17.1.1 the element's presence was tested, so a
// false value took the cs font too); `w:rtl` alone on Latin text *does* take the cs font; the
// theme language is matched by exact subtag, so Estonian is not Ethiopic; U+0020 is in the
// ASCII range and takes the ascii font, so the space between two East Asian words is the Latin
// font's; U+0590-U+07BF in a run with no `w:cs` takes the ascii font; a symbol font is
// recognised whatever case the document wrote it in; an East Asian reference with no East Asian
// font falls to hAnsi in every branch.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { PropertyResolver } from '../properties/PropertyResolver.mjs';
import { themeFontOf, type ThemeFontLang, type ThemeSource } from './ThemeFonts.mjs';
import { defaultThemeSetting, type DefaultTheme } from './defaultTheme.mjs';
import { coverageGroupOf } from './FontFallback.mjs';

/** Word's built-in default where the document defaults name no font at all (docx4j CR-016
 *  probe `fonts-missing-slots` (a)). */
export const BUILT_IN_DEFAULT_FONT = 'Times New Roman';

/** The fonts `symbolFontName` recognises; docx4j's `SymbolMapper` knows these five. */
const SYMBOL_FONTS = ['Symbol', 'Webdings', 'Wingdings', 'Wingdings 2', 'Wingdings 3'];

/** One stretch of a run's text set in one document font. */
export interface FontSpan {
  /** The text of the span. */
  readonly text: string;
  /** The document font (the name a `w:rFonts` slot gives, theme reference resolved); null
   *  where the run names no font at all and there is no default either. */
  readonly documentFont: string | null;
  readonly bold: boolean;
  readonly italic: boolean;
  /** `w:cs` of the effective run properties, by value. */
  readonly cs: boolean;
  /** `w:rtl` of the effective run properties, by value. */
  readonly rtl: boolean;
  /** The span's script, as `spanScript` groups it (`LATIN`, `CJK`, `SYMBOL`, `EMOJI`, ...);
   *  null where every character of the span is shared (a space, a digit, punctuation). */
  readonly script: string | null;
}

/** What the selector reads.  Structural, so this module imports no part class. */
export interface RunFontSelectorSource {
  /** The package's resolver; without one, the properties passed in are taken as they are. */
  resolver?: PropertyResolver | undefined;
  /** The theme part's `a:theme` (or its font scheme); absent for a package with no theme part. */
  theme?: ThemeSource;
  /** `w:themeFontLang` from the settings part. */
  themeFontLang?: ThemeFontLang | undefined;
  /** Which Office theme answers a reference where there is no theme part; the process-wide
   *  `defaultThemeSetting()` by default. */
  defaultTheme?: DefaultTheme | undefined;
  /** The face to use for an astral emoji, docx4j's `docx4j.fonts.RunFontSelector.EmojiFont`.
   *  Unset by default, as docx4j's property is, so an emoji takes the hAnsi font. */
  emojiFont?: string | undefined;
}

/** `spans()` options. */
export interface SpansOptions {
  /** True where the caller has resolved the run's effective properties already. */
  rPrIsEffective?: boolean;
}

/** `w:cs` and `w:rtl` are `ST_OnOff`: present without `w:val` means true, `w:val="0"` false. */
export function isOn(value: wml.BooleanDefaultTrue | undefined): boolean {
  return value !== undefined && value !== null && value.val !== false;
}

/** The canonical name of one of the symbol fonts, whatever case the document wrote it in;
 *  undefined for any other font.  Font names are case-insensitive, and Word draws 'symbol'
 *  with Symbol (docx4j CR-016 probe `fonts-symbol-and-emoji` (a), (b)). */
export function symbolFontName(documentFontName: string | undefined): string | undefined {
  if (documentFontName === undefined || documentFontName === null) return undefined;
  const name = documentFontName.trim().toLowerCase();
  return SYMBOL_FONTS.find((known) => known.toLowerCase() === name);
}

/** The emoji blocks - Mahjong, Domino and Playing Card tiles, the enclosed alphanumerics and
 *  ideographs, Miscellaneous Symbols and Pictographs, Emoticons, Transport and Map, the
 *  supplemental and extended symbol blocks.  Word sets them in Segoe UI Emoji. */
export function isEmoji(cp: number): boolean {
  return cp >= 0x1F000 && cp <= 0x1FAFF;
}

/**
 * The group a span is cut by: the coverage group of the code point (its Unicode script, or
 * SYMBOL / EMOJI), the CJK scripts as one, and null for a shared character (COMMON,
 * INHERITED), which joins whatever it follows.
 *
 * docx4j's rendering walk cuts a span where the *script* changes between two non-shared
 * characters as well as where the font does, so that a Greek stretch inside a Latin font is
 * substituted as one top-level span; measured, nesting it changed FOP's line stacking and cost
 * a corpus document a page.  {@link RunFontSelector.spans} folds by font alone (which is what
 * `documentFontFor` per code point answers, and what the parity goldens record) and reports the
 * script, so a renderer can cut further.
 */
export function spanScript(cp: number): string | null {
  const group = coverageGroupOf(cp);
  if (group === 'COMMON' || group === 'INHERITED') return null;
  if (group === 'HAN' || group === 'HIRAGANA' || group === 'KATAKANA' || group === 'HANGUL'
    || group === 'BOPOMOFO') return 'CJK';
  return group;
}

/** The table's preamble: an eastAsia of Times New Roman with ascii and hAnsi equal sets the
 *  whole run in the ascii font. */
export function preambleRule(eastAsia: string | undefined, ascii: string | undefined,
  hAnsi: string | undefined): boolean {
  return eastAsia === 'Times New Roman' && ascii !== undefined && ascii === hAnsi;
}

/** The run's language, for case conversion: `w:lang/@w:val`, else undefined (the root locale). */
export function runLocale(rPr: wml.RPr | undefined): string | undefined {
  const tag = rPr?.lang?.val;
  return tag !== undefined && tag.length > 0 ? tag : undefined;
}

function contains(langEastAsia: string | undefined, lang: string): boolean {
  // eg <w:lang w:eastAsia="zh-CN" .. />
  return langEastAsia !== undefined && langEastAsia.includes(lang);
}

/** The four slots and the run's hint, as one run's dispatch needs them. */
interface Slots {
  /** Set where the whole run goes in one font: a symbol font, the cs font, the preamble rule. */
  single?: string;
  eastAsia?: string;
  ascii?: string;
  hAnsi?: string;
  cs?: string;
  hint?: wml.STHint;
  langEastAsia?: string;
}

export class RunFontSelector {

  constructor(private readonly source: RunFontSelectorSource = {}) {}

  private defaultFontCache: string | undefined;

  /** Which Office theme answers a reference where the package has no theme part. */
  get defaultTheme(): DefaultTheme {
    return this.source.defaultTheme ?? defaultThemeSetting();
  }

  /**
   * The face a theme reference names: the theme part's answer for the document's
   * `themeFontLang`; or, where the document has **no theme part**, the Office theme's own Latin
   * faces - Word supplies that theme to such a document, and the explicit `w:ascii` beside the
   * reference is unused (docx4j CR-016 probe `fonts-missing-slots` (b)).
   *
   * Only the two Latin slots are answered in the themeless case: the Office theme's East Asian
   * and complex-script entries are empty, so those references resolve to nothing and the
   * explicit attribute, if any, stands.  Which Latin faces they are is
   * {@link RunFontSelector.defaultTheme}.
   */
  themeFont(type: wml.STTheme | undefined): string | undefined {
    if (type === undefined || type === null) return undefined;
    if (this.source.theme !== undefined) {
      return themeFontOf(this.source.theme, type, this.source.themeFontLang);
    }
    switch (type) {
      case 'minorAscii': case 'minorHAnsi': return this.defaultTheme.minorLatin;
      case 'majorAscii': case 'majorHAnsi': return this.defaultTheme.majorLatin;
      default: return undefined;
    }
  }

  /**
   * The document fonts an `rFonts` names once its theme references are resolved: for each of
   * the four slots, the theme's face, else the explicit attribute.  Blank names are left out.
   * This is what font discovery collects (`MainDocumentPart.fontsInUse()`).
   */
  documentFontsOf(rFonts: wml.RFonts | undefined): string[] {
    const names: string[] = [];
    if (rFonts === undefined || rFonts === null) return names;
    const refs = [rFonts.asciiTheme, rFonts.hAnsiTheme, rFonts.eastAsiaTheme, rFonts.cstheme];
    const explicit = [rFonts.ascii, rFonts.hAnsi, rFonts.eastAsia, rFonts.cs];
    for (let i = 0; i < 4; i++) {
      const name = this.themeFont(refs[i]) ?? explicit[i];
      if (name !== undefined && name.trim().length > 0) names.push(name.trim());
    }
    return names;
  }

  /**
   * The document's default font for Latin text: the document defaults' `w:ascii` (a theme
   * reference resolved first), else `w:hAnsi`; Times New Roman where the defaults name no font
   * at all.  A theme reference the theme cannot answer gives Calibri, as docx4j's does.
   */
  defaultFontOf(docDefaultsRFonts: wml.RFonts | undefined): string {
    let f: string | undefined;
    if (docDefaultsRFonts !== undefined && docDefaultsRFonts !== null) {
      const rFonts = docDefaultsRFonts;
      if (rFonts.asciiTheme !== undefined) f = this.themeFont(rFonts.asciiTheme);
      if (f === undefined) f = rFonts.ascii;
      if (f === undefined && rFonts.asciiTheme !== undefined) {
        // a theme reference the theme part cannot answer (an empty typeface)
        f = 'Calibri';
      }
      if (f === undefined && rFonts.hAnsiTheme !== undefined) f = this.themeFont(rFonts.hAnsiTheme);
      if (f === undefined) f = rFonts.hAnsi;
    }
    return f ?? BUILT_IN_DEFAULT_FONT;
  }

  /** {@link defaultFontOf} of the document defaults the resolver holds, computed once. */
  get defaultFont(): string {
    this.defaultFontCache ??= this.defaultFontOf(this.source.resolver?.getDocumentDefaultRPr()?.rFonts);
    return this.defaultFontCache;
  }

  /** Forget the cached default font: the styles part changed under us. */
  refresh(): void {
    this.defaultFontCache = undefined;
  }

  /**
   * The document font an `rPr` asks for its ASCII text: `w:ascii`, or the face `w:asciiTheme`
   * names; failing those `w:hAnsi`, and failing that the document default.  Never empty.
   *
   * This is what a *name* is - what Office JS's `font.name` reports and what docx4j uses for
   * the line metrics of generated text and paragraph marks.  It is not the dispatch: a run of
   * Arabic in a font named here may still be drawn in the cs font
   * ({@link RunFontSelector.documentFontFor}).
   */
  asciiFontName(rPr: wml.RPr | undefined): string {
    const rFonts = rPr?.rFonts;
    if (rFonts !== undefined && rFonts !== null) {
      if (rFonts.asciiTheme !== undefined) {
        const f = this.themeFont(rFonts.asciiTheme);
        if (f !== undefined && f.length > 0) return f;
      }
      if (rFonts.ascii !== undefined) return rFonts.ascii;
      if (rFonts.hAnsi !== undefined) return rFonts.hAnsi;
    }
    return this.defaultFont;
  }

  /** The run's effective properties, never undefined. */
  private effectiveRPr(pPr: wml.PPr | undefined, rPr: wml.RPr | undefined,
    rPrIsEffective: boolean): wml.RPr {
    if (!rPrIsEffective && this.source.resolver !== undefined) {
      return this.source.resolver.getEffectiveRPr(rPr, pPr) ?? {};
    }
    return rPr ?? {};
  }

  /**
   * The run's `w:rFonts`, or where nothing anywhere names a font, the document default for the
   * ascii and hAnsi slots, the East Asian and complex-script slots left unnamed as for any run
   * that never named them, and then the range dispatch as for any other run.  (Not the eastAsia
   * slot: a Times New Roman there would fire the preamble rule and set the whole run in one
   * span.)
   */
  private rFontsOf(rPr: wml.RPr): wml.RFonts {
    const rFonts = rPr.rFonts;
    if (rFonts !== undefined && rFonts !== null) return rFonts;
    return { ascii: this.defaultFont, hAnsi: this.defaultFont };
  }

  /**
   * The complex-script font the run is set in as a whole, where `w:cs` or `w:rtl` is **on** and
   * the cs slot names a font (the theme reference resolved first, else the attribute);
   * undefined where the run is dispatched by range instead - neither is on, or the reference
   * resolves to nothing and there is no `w:cs`.
   */
  complexScriptFont(rPr: wml.RPr, rFonts: wml.RFonts): string | undefined {
    if (!isOn(rPr.cs) && !isOn(rPr.rtl)) return undefined;
    if (rFonts.cstheme !== undefined) return this.themeFont(rFonts.cstheme) ?? rFonts.cs;
    return rFonts.cs; // undefined: no CS value, dispatched by range
  }

  /**
   * The four slots resolved: each the theme reference where there is one, else the explicit
   * attribute (which also stands where the reference resolves to nothing).  Order: eastAsia,
   * ascii, hAnsi, cs.
   */
  resolvedSlots(rFonts: wml.RFonts): [string | undefined, string | undefined, string | undefined, string | undefined] {
    const eastAsia = this.themeFont(rFonts.eastAsiaTheme) ?? rFonts.eastAsia;
    const ascii = this.themeFont(rFonts.asciiTheme) ?? rFonts.ascii;
    const hAnsi = this.themeFont(rFonts.hAnsiTheme) ?? rFonts.hAnsi;
    let cs = this.themeFont(rFonts.cstheme) ?? rFonts.cs;
    // eg LibreOffice writes w:cs="" in docDefaults
    if (cs !== undefined && cs.trim().length === 0) cs = undefined;
    return [eastAsia, ascii, hAnsi, cs];
  }

  /** The run's slots, ready for {@link fontFor}: the body of docx4j's `fontSelector` before the
   *  walk, shared by {@link documentFontFor} and {@link spans}. */
  private slotsOf(pPr: wml.PPr | undefined, rPr: wml.RPr): Slots {
    const rFonts = this.rFontsOf(rPr);

    const symbolFont = symbolFontName(rFonts.hAnsi);
    if (symbolFont !== undefined) return { single: symbolFont };

    const csFont = this.complexScriptFont(rPr, rFonts);
    if (csFont !== undefined) return { single: csFont };

    const [eastAsia, ascii0, hAnsi0, cs] = this.resolvedSlots(rFonts);
    if (preambleRule(eastAsia, ascii0, hAnsi0)) return { single: ascii0 };

    return {
      eastAsia,
      ascii: ascii0 ?? this.defaultFont,
      hAnsi: hAnsi0 ?? this.defaultFont,
      cs,
      hint: rFonts.hint,
      langEastAsia: rPr.lang?.eastAsia,
    };
  }

  private fontOf(slots: Slots, cp: number): string | undefined {
    if (slots.single !== undefined) return slots.single;
    return this.fontFor(cp, slots.hint, slots.langEastAsia, slots.eastAsia, slots.ascii,
      slots.hAnsi, slots.cs);
  }

  /**
   * The document font the selector picks for one code point in a run with these properties: the
   * symbol font by its canonical name where the run's hAnsi is one, the complex-script font
   * where `w:cs` or `w:rtl` is on and names one, the ascii font under the preamble rule, else
   * the font the character-range table gives the code point.  Never null: the document default
   * at least.
   *
   * @param rPrIsEffective true where the caller has resolved the run's properties already
   */
  documentFontFor(pPr: wml.PPr | undefined, rPr: wml.RPr | undefined, codePoint: number,
    rPrIsEffective = false): string {
    const effective = this.effectiveRPr(pPr, rPr, rPrIsEffective);
    return this.fontOf(this.slotsOf(pPr, effective), codePoint) ?? this.defaultFont;
  }

  /**
   * The run's text folded into maximal stretches of one document font.
   *
   * Every code point is asked for its font ({@link documentFontFor}) and joins the span being
   * built while the answer is the same, whatever its range - docx4j's walk since 17.1.1, which
   * replaced cutting by *range* (a Latin run was cut at every exception character, a space went
   * to the East Asian font of the word before it, and the ASCII letters after an accented one
   * stayed in the hAnsi font).  `bold`, `italic`, `cs` and `rtl` are the effective run
   * properties' and are the same for every span of a run.
   */
  spans(pPr: wml.PPr | undefined, rPr: wml.RPr | undefined, text: string | undefined,
    options: SpansOptions = {}): FontSpan[] {
    const spans: FontSpan[] = [];
    if (text === undefined || text === null || text.length === 0) return spans;

    const effective = this.effectiveRPr(pPr, rPr, options.rPrIsEffective === true);
    const bold = isOn(effective.b);
    const italic = isOn(effective.i);
    const cs = isOn(effective.cs);
    const rtl = isOn(effective.rtl);
    const slots = this.slotsOf(pPr, effective);

    const memo = new Map<number, string | null>();
    let current = '';
    let currentFont: string | null | undefined;
    let currentScript: string | null = null;
    const flush = (): void => {
      if (currentFont === undefined) return;
      spans.push({ text: current, documentFont: currentFont, bold, italic, cs, rtl, script: currentScript });
    };
    for (const ch of text) {
      const cp = ch.codePointAt(0) as number;
      let font = memo.get(cp);
      if (font === undefined) {
        font = this.fontOf(slots, cp) ?? null;
        memo.set(cp, font);
      }
      if (currentFont !== undefined && currentFont !== font) {
        flush();
        current = '';
        currentScript = null;
      }
      currentFont = font;
      const script = spanScript(cp);
      if (script !== null && currentScript === null) currentScript = script;
      current += ch;
    }
    flush();
    return spans;
  }

  /**
   * The document font which formats one code point, given the run's four fonts (each already
   * resolved through the theme) and its hint: the table in [MS-OI29500] 17.3.2.26, with these
   * departures from it, each measured by docx4j:
   *
   * - the Indic, Thai, Lao, Myanmar and Khmer ranges use cs, where the table says hAnsi
   *   (docx4j issues 622 and 666);
   * - an East Asian reference is "eastAsia (or eastAsiaTheme) if defined": a run naming none
   *   gets hAnsi in those places, never nothing;
   * - characters outside the Basic Multilingual Plane, which the table does not cover: the
   *   {@link RunFontSelectorSource.emojiFont} where one is configured, else hAnsi (Word uses
   *   Segoe UI Emoji, which docx4j's coverage pass finds).
   *
   * What the table says and Word confirmed: U+0020 is ASCII and takes the ascii font, so the
   * space between two East Asian words is the Latin font's; Hebrew, Arabic and the rest of
   * U+0590-U+07BF in a run with no `w:cs` take the ascii font.  The table's other condition for
   * U+0100-U+02AF and U+1E00-U+1EFF under `hint=eastAsia`, "or the character set of the
   * eastAsia font is Chinese5 or GB2312", needs the font's OS/2 code-page bits and is not
   * implemented here either.
   *
   * **U+2190-U+2BFF**: docx4j asks whether the hAnsi font has the glyph and names Segoe UI
   * Symbol where it does not.  That is a glyph check over a font file, which this CR does not
   * port (CR-001 section 14.2), so the hAnsi font is answered - which is also docx4j's answer
   * wherever Segoe UI Symbol is not among the fonts it has, the environment the parity goldens
   * record.
   *
   * @returns a document font name; undefined only where the run names no hAnsi font either
   */
  fontFor(cp: number, hint: wml.STHint | undefined, langEastAsia: string | undefined,
    eastAsia: string | undefined, ascii: string | undefined, hAnsi: string | undefined,
    cs: string | undefined): string | undefined {

    // "eastAsia (or eastAsiaTheme if defined)": where the run names none, the table's hAnsi
    const ea = eastAsia ?? hAnsi;
    const hintEA = hint === 'eastAsia';

    if (cp > 0xFFFF) {
      if (isEmoji(cp) && this.source.emojiFont !== undefined) return this.source.emojiFont;
      return hAnsi;
    }
    const c = cp;
    // Basic Latin: "@ascii is used to format all characters in the ASCII range (0-127)"
    if (c <= 0x7F) return ascii;
    // Latin-1 Supplement: hAnsi, with the table's exceptions under hint=eastAsia
    if (c >= 0xA0 && c <= 0xFF) {
      if (hintEA && eastAsia !== undefined) {
        // "If hint is eastAsia, the following characters use eastAsia: A1, A4, A7-A8, AA, AD,
        //  AF, B0-B4, B6-BA, BC-BF, D7, F7"
        if (c === 0xA1 || c === 0xA4 || (c >= 0xA7 && c <= 0xA8) || c === 0xAA
          || c === 0xAD || c === 0xAF || (c >= 0xB0 && c <= 0xB4)
          || (c >= 0xB6 && c <= 0xBA) || (c >= 0xBC && c <= 0xBF)
          || c === 0xD7 || c === 0xF7) return eastAsia;
        // "If hint is eastAsia and the language of the run is either Chinese Traditional or
        //  Chinese Simplified, the following characters use eastAsia: E0-E1, E8-EA, EC-ED,
        //  F2-F3, F9-FA, FC"
        if (contains(langEastAsia, 'zh') && ((c >= 0xE0 && c <= 0xE1) || (c >= 0xE8 && c <= 0xEA)
          || (c >= 0xEC && c <= 0xED) || (c >= 0xF2 && c <= 0xF3)
          || (c >= 0xF9 && c <= 0xFA) || c === 0xFC)) return eastAsia;
      }
      return hAnsi;
    }
    // Latin Extended-A, Latin Extended-B, IPA Extensions
    if (c >= 0x100 && c <= 0x2AF) return (hintEA && contains(langEastAsia, 'zh')) ? ea : hAnsi;
    // Spacing Modifier Letters .. Cyrillic Supplement
    if (c >= 0x2B0 && c <= 0x4FF) return hintEA ? ea : hAnsi;
    // Hebrew, Arabic, Syriac, Thaana, NKo: ascii (measured, fonts-hebrew-no-cs)
    if (c >= 0x590 && c <= 0x7BF) return ascii;
    // the Indic scripts, Thai, Lao, Myanmar, Khmer and the Khmer symbols: cs (issues 622, 666)
    if ((c >= 0x900 && c <= 0xDFF) || (c >= 0xE00 && c <= 0xEFF) || (c >= 0x1000 && c <= 0x109F)
      || (c >= 0x1780 && c <= 0x17FF) || (c >= 0x19E0 && c <= 0x19FF)) {
      return cs ?? hAnsi;
    }
    // Hangul Jamo
    if (c >= 0x1100 && c <= 0x11FF) return ea;
    // Latin Extended Additional
    if (c >= 0x1E00 && c <= 0x1EFF) return (hintEA && contains(langEastAsia, 'zh')) ? ea : hAnsi;
    // General Punctuation .. Number Forms: ordinary text
    if (c >= 0x2000 && c <= 0x218F) return hintEA ? ea : hAnsi;
    // the symbol blocks: Arrows .. Braille (docx4j glyph-checks these; see the javadoc above)
    if (c >= 0x2190 && c <= 0x2BFF) return hintEA ? ea : hAnsi;
    // Glagolitic .. CJK Radicals Supplement
    if (c >= 0x2C00 && c <= 0x2EFF) return hintEA ? ea : hAnsi;
    // the CJK ranges (and the surrogate range, which a paired surrogate never reaches here)
    if (c >= 0x2F00 && c <= 0xDFFF) return ea;
    // Private Use Area; F000-F0FF is where the symbol fonts live
    if (c >= 0xE000 && c <= 0xF8FF) return hintEA ? ea : hAnsi;
    // CJK Compatibility Ideographs
    if (c >= 0xF900 && c <= 0xFAFF) return ea;
    // Alphabetic Presentation Forms: "If the hint is eastAsia then eastAsia is used for
    // FB00-FB1C.  For FB1D-FB4F, ascii is used."
    if (c >= 0xFB00 && c <= 0xFB4F) {
      if (c >= 0xFB1D) return ascii;
      return hintEA ? ea : hAnsi;
    }
    // Arabic Presentation Forms-A
    if (c >= 0xFB50 && c <= 0xFDFF) return ascii;
    // CJK Compatibility Forms, Small Form Variants
    if (c >= 0xFE30 && c <= 0xFE6F) return ea;
    // Arabic Presentation Forms-B
    if (c >= 0xFE70 && c <= 0xFEFE) return ascii;
    // Halfwidth and Fullwidth Forms
    if (c >= 0xFF00 && c <= 0xFFEF) return ea;
    // "for all ranges not listed in the above, the hAnsi font shall be used"
    // (Georgian, Armenian, Ethiopic, Tibetan, Mongolian, Greek Extended, the C1 controls ...)
    return hAnsi;
  }
}
