// SpreadsheetML colours resolved as Excel draws them: docx4j `org.docx4j.model.color.SpreadsheetColors` (CR-034).
import type * as sml from '@docx4j/generated-objects-ts/modules/org_xlsx4j_sml';
import * as ColorMath from './ColorMath.mjs';
import { INDEXED } from './ColorTables.mjs';
import { bytesOf } from './DrawingMLColors.mjs';
import { ResolvedColor } from './ResolvedColor.mjs';
import type { ThemeColors } from './ThemeColors.mjs';

/** ARGB (4 bytes) or RGB (3 bytes); the alpha byte is not drawn by Excel and is ignored. */
function argb(v: unknown): ResolvedColor | undefined {
  const b = bytesOf(v);
  if (b === undefined) return undefined;
  if (b.length === 4) return ResolvedColor.rgb(b[1]!, b[2]!, b[3]!);
  if (b.length === 3) return ResolvedColor.rgb(b[0]!, b[1]!, b[2]!);
  return undefined;
}

/**
 * SpreadsheetML colours (`x:color`, `x:fgColor`, `x:bgColor` and the font and border colours, ISO/IEC 29500-1
 * section 18.8) resolved as Excel draws them (docx4j CR-034 section 11, P4).
 *
 * - `theme`: Excel's index swaps the first two pairs (0 lt1, 1 dk1, 2 lt2, 3 dk2, then the accents and hyperlinks),
 *   as {@link ThemeColors.sheetThemeColor} does.
 * - `indexed`: the legacy 64-colour palette, or all 64 entries of the workbook's own `x:indexedColors`; 64 and 65
 *   the system foreground and background (black, white).
 * - `tint`, −1 to 1: the HLS luminance, L' = L·(1 + tint) below 0, L·(1 − tint) + HLSMAX·tint above, in Windows'
 *   integer HLS (HLSMAX 240), the luminance truncated after adding 0.02; a tint of 0 leaves the colour alone. All 90
 *   theme-and-tint cells measured exact. (Not POI's per-channel rule.)
 *
 * The `rgb` attributes are hexBinary, which the objects model holds as byte numbers (its declarations say a string;
 * either is read). docx4j's `null` is `undefined` here.
 */
export const SpreadsheetColors = {
  /**
   * @param c the colour element
   * @param theme the workbook's theme
   * @param indexedColors the workbook's `x:indexedColors`, or undefined for the legacy palette
   * @returns undefined for `auto` or nothing stated: the caller's automatic colour
   */
  resolve(c: sml.CTColor | undefined | null, theme: ThemeColors | undefined, indexedColors?: sml.CTIndexedColors | null): ResolvedColor | undefined {
    if (c === undefined || c === null || c.auto === true) return undefined;
    let base: ResolvedColor | undefined;
    if (c.rgb !== undefined && c.rgb !== null) {
      base = argb(c.rgb);
    } else if (c.theme !== undefined && c.theme !== null) {
      base = theme === undefined ? undefined : theme.sheetThemeColor(Math.trunc(c.theme));
    } else if (c.indexed !== undefined && c.indexed !== null) {
      base = SpreadsheetColors.indexed(Math.trunc(c.indexed), indexedColors);
    }
    if (base === undefined) return undefined;
    const tint = c.tint ?? 0;
    return tint === 0 ? base : SpreadsheetColors.tint(base, tint);
  },

  /** An indexed colour: the workbook's palette for 0 to 63 where it has one, else the legacy one. */
  indexed(i: number, indexedColors?: sml.CTIndexedColors | null): ResolvedColor | undefined {
    if (indexedColors !== undefined && indexedColors !== null && i >= 0 && i < 64) {
      const custom = indexedColors.rgbColor ?? [];
      if (i < custom.length && custom[i]!.rgb !== undefined && custom[i]!.rgb !== null) return argb(custom[i]!.rgb);
    }
    return i >= 0 && i < INDEXED.length ? ResolvedColor.ofHex(INDEXED[i]) : undefined;
  },

  /**
   * Excel's tint: Windows' integer HLS, the luminance truncated after adding 0.02. Plain truncation is two cells out
   * of 90 short (both at tint 0.3333, whose luminance comes to just under a whole number and which Excel rounds up);
   * any tolerance from 0.01 to 0.05 gives all 90, presumably the precision Excel keeps a tint at.
   * Package-private in docx4j.
   */
  tint(c: ResolvedColor, tint: number): ResolvedColor {
    const hls = ColorMath.rgbToHls240(c.red, c.green, c.blue);
    const l = tint < 0 ? hls[1] * (1 + tint) : hls[1] * (1 - tint) + (ColorMath.HLSMAX - ColorMath.HLSMAX * (1 - tint));
    let li = Math.trunc(l + 0.02);
    if (li < 0) li = 0;
    if (li > ColorMath.HLSMAX) li = ColorMath.HLSMAX;
    const rgb = ColorMath.hls240ToRgb(hls[0], li, hls[2]);
    return ResolvedColor.rgb(rgb[0], rgb[1], rgb[2]);
  },
};
