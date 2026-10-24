// WordprocessingML colours resolved as Word draws them: docx4j `org.docx4j.model.color.WordColors` (CR-034).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { wordColor } from '@docx4j/generated-objects-ts/helpers/wml';
import * as ColorMath from './ColorMath.mjs';
import { ResolvedColor } from './ResolvedColor.mjs';
import type { ThemeColors } from './ThemeColors.mjs';

function hexByte(s: string | undefined | null): number | undefined {
  if (s === undefined || s === null || s.length === 0) return undefined;
  // Java's Integer.parseInt(s, 16): an optional sign, then hex digits
  if (!/^[+-]?[0-9A-Fa-f]+$/.test(s)) return undefined;
  const v = Number.parseInt(s, 16);
  return v >= 0 && v <= 255 ? v : undefined;
}

function truncate(v: number): number {
  return Math.floor(ColorMath.clamp01(v) * 255 + 1e-9);
}

/**
 * WordprocessingML colours (`w:color`, `w:shd`, the borders, `w:u`) resolved as Word draws them (docx4j CR-034
 * section 11, P1).
 *
 * - A theme reference (`w:themeColor`, `w:themeFill`) wins over the value written beside it (`w:val`, `w:color`,
 *   `w:fill`), which is a cache Word keeps: Word draws the theme colour even where the value disagrees (an older
 *   theme's, or anything else).
 * - `w:themeTint` and `w:themeShade` change the HSL luminance: tint t, L' = L·t + (1 − t); shade t, L' = L·t, t
 *   being the byte over 255 (MS-OI29500 2.1.72). Where both are present Word applies the tint (the same note).
 * - Word's own rounding of that arithmetic is not reproduced exactly (within one unit per channel in all 7,700 tints
 *   and shades of four corpora), so where the written value is within one unit of the computed one - a current
 *   cache - Word's value is returned: exact for every attribute Word itself wrote against the document's theme.
 * - A value is read as Word 365 reads it ({@link WordColors.normalise}, docx4j CR-060 and core-ts CR-015): trimmed,
 *   hex in either case with an optional `#`, or one of the 140 classic HTML colour names. One Word rejects is
 *   automatic. A fill, a pattern colour and a border's colour are read the same way, as Word reads them (round 5: a
 *   name there is drawn): docx4j's `WordColors` since `a9e0303ff` (until then it normalised in `of(Color)` and `of(U)`
 *   only; CR-015 section 11.4).
 *
 * docx4j's overloads `of(Color)`, `of(CTBorder)` and `of(U)` are {@link WordColors.of}, {@link WordColors.ofBorder}
 * and {@link WordColors.ofUnderline} here, since the three objects can't be told apart by shape; its `null` (for
 * automatic) is `undefined`.
 */
export const WordColors = {
  /**
   * The colour Word draws for a WordprocessingML colour's attributes.
   *
   * @param val the written value: RRGGBB, "auto", or undefined
   * @param themeColor the theme reference, or undefined (or "none")
   * @param themeTint a byte in hex ("99"), or undefined
   * @param themeShade a byte in hex ("BF"), or undefined
   * @returns undefined for automatic (no theme reference and val "auto" or absent): the caller's default
   */
  resolve(val: string | undefined | null, themeColor: string | undefined | null, themeTint: string | undefined | null,
    themeShade: string | undefined | null, theme: ThemeColors | undefined): ResolvedColor | undefined {
    const written = ResolvedColor.ofHex(val);
    if (themeColor === undefined || themeColor === null || themeColor === 'none' || theme === undefined) return written;
    // Word's value where it is a current cache: Word's own rounding, which the computation below matches for Word 365
    // but not for every older version that wrote the file
    const base = theme.wordThemeColor(themeColor);
    if (base === undefined) return written;
    const tint = hexByte(themeTint), shade = hexByte(themeShade);
    const computed = tint !== undefined ? WordColors.luminance(base, tint / 255.0, true)
      : shade !== undefined ? WordColors.luminance(base, shade / 255.0, false) : base;
    return written !== undefined && written.within(computed, 1) ? written : computed;
  },

  /**
   * A colour value (`w:color/@w:val`, `w:u/@w:color`, `w:shd/@w:fill`, a border's `w:color`) as Word 365 reads it
   * (docx4j CR-060 section 1, rule 3, and sections 12 and 13): surrounding whitespace trimmed, six hex digits in
   * either case with an optional leading `#`; else a classic HTML colour name, in any case. The rule is the objects
   * package's `wordColor` (its CR-013, held to docx4j's `WordColorValues`).
   *
   * @returns six upper-case hex digits; undefined for `auto` and for a value Word rejects, which Word reads as
   * automatic in a fill or a border colour, and as absent in a text or underline colour ({@link WordColors.stated})
   */
  normalise(val: string | undefined | null): string | undefined {
    const w = typeof val === 'string' ? wordColor(val) : undefined;
    return w === undefined || w === 'auto' ? undefined : w.substring(1);
  },

  /**
   * Whether a `w:color/@w:val` or `w:u/@w:color` value states a colour, as Word reads it (docx4j CR-060 sections 12
   * and 13): one {@link WordColors.normalise} accepts, or exactly `auto` (automatic, which overrides a style's
   * colour). Word reads any other value (`F00`, `bogus`, `""`, `Auto`, `" auto"`) as absent: a style's colour beneath
   * still shows. (In a fill or a border colour it reads such a value as automatic instead.)
   */
  stated(val: string | undefined | null): boolean {
    return typeof val === 'string' && wordColor(val) !== undefined;
  },

  /** A run's colour (`w:color`); undefined for automatic. */
  of(c: wml.Color | undefined | null, theme: ThemeColors | undefined): ResolvedColor | undefined {
    return c === undefined || c === null ? undefined
      : WordColors.resolve(WordColors.normalise(c.val), c.themeColor, c.themeTint, c.themeShade, theme);
  },

  /** A shading's fill (`w:fill`, `w:themeFill`); undefined for automatic. */
  fill(shd: wml.CTShd | undefined | null, theme: ThemeColors | undefined): ResolvedColor | undefined {
    return shd === undefined || shd === null ? undefined
      : WordColors.resolve(WordColors.normalise(shd.fill), shd.themeFill, shd.themeFillTint, shd.themeFillShade, theme);
  },

  /** A shading's pattern colour (`w:color`, `w:themeColor`); undefined for automatic. */
  pattern(shd: wml.CTShd | undefined | null, theme: ThemeColors | undefined): ResolvedColor | undefined {
    return shd === undefined || shd === null ? undefined
      : WordColors.resolve(WordColors.normalise(shd.color), shd.themeColor, shd.themeTint, shd.themeShade, theme);
  },

  /** A border's colour; undefined for automatic. docx4j `of(CTBorder)`. */
  ofBorder(border: wml.CTBorder | undefined | null, theme: ThemeColors | undefined): ResolvedColor | undefined {
    return border === undefined || border === null ? undefined
      : WordColors.resolve(WordColors.normalise(border.color), border.themeColor, border.themeTint, border.themeShade, theme);
  },

  /** An underline's colour; undefined for automatic (the text's colour). docx4j `of(U)`. */
  ofUnderline(u: wml.U | undefined | null, theme: ThemeColors | undefined): ResolvedColor | undefined {
    return u === undefined || u === null ? undefined
      : WordColors.resolve(WordColors.normalise(u.color), u.themeColor, u.themeTint, u.themeShade, theme);
  },

  /**
   * Word's themeTint (tint true) or themeShade on a colour: float HSL, the channels truncated. That is what Word 365
   * draws (docx4j CR-034 section 11: the three P1 cases with no usable written value, all exact); the values older
   * versions of Word wrote are within one unit of it (all 119 distinct ones in four corpora, 77 exact).
   * Package-private in docx4j; public here for the corpus vectors' test.
   */
  luminance(c: ResolvedColor, t: number, tint: boolean): ResolvedColor {
    const hsl = ColorMath.rgbToHsl(c.red / 255.0, c.green / 255.0, c.blue / 255.0);
    const l = tint ? hsl[2] * t + (1 - t) : hsl[2] * t;
    const rgb = ColorMath.hslToRgb(hsl[0], hsl[1], ColorMath.clamp01(l));
    return ResolvedColor.rgb(truncate(rgb[0]), truncate(rgb[1]), truncate(rgb[2]));
  },
};
