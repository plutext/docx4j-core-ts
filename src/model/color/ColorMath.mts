// The colour arithmetic Office uses, as docx4j CR-034's probes measured it (section 11): sRGB's transfer function,
// HSL with an unclamped saturation, and Windows' integer HLS for SpreadsheetML tints. docx4j
// `org.docx4j.model.color.ColorMath`, expression for expression: JS numbers are IEEE doubles as Java's are, so the
// same expressions in the same order give the same bits, which the luminance transforms' rounding depends on
// (CR-008 section 3). Java's integer division is `Math.trunc` of the quotient here; nothing is simplified.

// ------------------------------------------------------------------ sRGB and linear RGB

/** sRGB (0 to 1) to linear RGB (IEC 61966-2-1). */
export function toLinear(c: number): number {
  return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
}

/** linear RGB to sRGB, clamped to 0 to 1. */
export function toSrgb(l: number): number {
  const c = clamp01(l);
  return c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055;
}

export function clamp01(v: number): number {
  return v < 0 ? 0 : v > 1 ? 1 : v;
}

/** a channel, 0 to 1, as 0 to 255, rounded half up */
export function channel(v: number): number {
  return Math.floor(clamp01(v) * 255 + 0.5);
}

/** a channel, 0 to 1, as 0 to 255, rounded half down (Office's luminance transforms) */
export function channelHalfDown(v: number): number {
  // Math.ceil(-0.5) is -0 in JS, 0 in Java's (int) cast; ResolvedColor's clamp makes either 0
  return Math.ceil(clamp01(v) * 255 - 0.5);
}

// ------------------------------------------------------------------ HSL

/** [h, s, l], each 0 to 1, from sRGB channels 0 to 1; a grey's hue is 0 */
export function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, Math.max(g, b)), min = Math.min(r, Math.min(g, b));
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l <= 0.5 ? d / (max + min) : d / (2 - max - min);
  let h: number;
  if (max === r) {
    h = ((g - b) / d) % 6;
    if (h < 0) h += 6;
  } else if (max === g) {
    h = (b - r) / d + 2;
  } else {
    h = (r - g) / d + 4;
  }
  return [h / 6, s, l];
}

/**
 * sRGB channels (0 to 1, clamped) from HSL; the saturation is not clamped, as Office does not clamp it (a
 * saturation above 1 pushes a channel past its range, which is then clamped).
 */
export function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  return [hslChannel(0, h, s, l), hslChannel(8, h, s, l), hslChannel(4, h, s, l)];
}

/**
 * sRGB channels (0 to 1, not clamped) from HSL as Foley and van Dam give it: the same colour as {@link hslToRgb}
 * but not the same rounding error, and the one Office's luminance transforms round as.
 */
export function hslToRgbFoley(h: number, s: number, l: number): [number, number, number] {
  // keep this form and its order of operations, and don't fold it into hslToRgb: a channel that is exactly half a
  // unit rounds (channelHalfDown) the way this form's floating-point error leans, which is the way Office rounds it
  // (docx4j CR-034 section 11, follow-up: 29 of 30 halves; hslToRgb's error leans otherwise)
  const m2 = l <= 0.5 ? l * (1 + s) : l + s - l * s, m1 = 2 * l - m2;
  return [foley(m1, m2, h + 1 / 3.0), foley(m1, m2, h), foley(m1, m2, h - 1 / 3.0)];
}

function foley(m1: number, m2: number, h: number): number {
  h -= Math.floor(h);
  if (h < 1 / 6.0) return m1 + (m2 - m1) * 6 * h;
  if (h < 1 / 2.0) return m2;
  if (h < 2 / 3.0) return m1 + (m2 - m1) * (2 / 3.0 - h) * 6;
  return m1;
}

function hslChannel(n: number, h: number, s: number, l: number): number {
  let k = (n + h * 12) % 12;
  if (k < 0) k += 12;
  const a = s * Math.min(l, 1 - l);
  return clamp01(l - a * Math.max(-1, Math.min(Math.min(k - 3, 9 - k), 1)));
}

// ------------------------------------------------------------------ Windows' integer HLS (Excel)

export const HLSMAX = 240;
const RGBMAX = 255;

/** Java's int division. */
const div = (a: number, b: number): number => Math.trunc(a / b);

/** [h, l, s] on 0 to 240, as Windows' ColorRGBToHLS (KB29240) computes them */
export function rgbToHls240(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, Math.max(g, b)), min = Math.min(r, Math.min(g, b));
  const l = div((max + min) * HLSMAX + RGBMAX, 2 * RGBMAX);
  if (max === min) return [div(HLSMAX * 2, 3), l, 0];
  const s = l <= div(HLSMAX, 2)
    ? div((max - min) * HLSMAX + div(max + min, 2), max + min)
    : div((max - min) * HLSMAX + div(2 * RGBMAX - max - min, 2), 2 * RGBMAX - max - min);
  const d = max - min;
  const rd = div((max - r) * div(HLSMAX, 6) + div(d, 2), d);
  const gd = div((max - g) * div(HLSMAX, 6) + div(d, 2), d);
  const bd = div((max - b) * div(HLSMAX, 6) + div(d, 2), d);
  let h: number;
  if (r === max) h = bd - gd;
  else if (g === max) h = div(HLSMAX, 3) + rd - bd;
  else h = div(2 * HLSMAX, 3) + gd - rd;
  if (h < 0) h += HLSMAX;
  if (h > HLSMAX) h -= HLSMAX;
  return [h, l, s];
}

/**
 * RGB from {@link rgbToHls240}'s values, as Windows' ColorHLSToRGB, except that a grey is rounded rather than
 * truncated, which is what Excel's fills show (docx4j CR-034 section 11, P4).
 */
export function hls240ToRgb(h: number, l: number, s: number): [number, number, number] {
  if (s === 0) {
    const v = div(l * RGBMAX + div(HLSMAX, 2), HLSMAX);
    return [v, v, v];
  }
  const m2 = l <= div(HLSMAX, 2) ? div(l * (HLSMAX + s) + div(HLSMAX, 2), HLSMAX) : l + s - div(l * s + div(HLSMAX, 2), HLSMAX);
  const m1 = 2 * l - m2;
  return [
    div(hueToRgb(m1, m2, h + div(HLSMAX, 3)) * RGBMAX + div(HLSMAX, 2), HLSMAX),
    div(hueToRgb(m1, m2, h) * RGBMAX + div(HLSMAX, 2), HLSMAX),
    div(hueToRgb(m1, m2, h - div(HLSMAX, 3)) * RGBMAX + div(HLSMAX, 2), HLSMAX)];
}

function hueToRgb(n1: number, n2: number, hue: number): number {
  if (hue < 0) hue += HLSMAX;
  if (hue > HLSMAX) hue -= HLSMAX;
  if (hue < div(HLSMAX, 6)) return n1 + div((n2 - n1) * hue + div(HLSMAX, 12), div(HLSMAX, 6));
  if (hue < div(HLSMAX, 2)) return n2;
  if (hue < div(HLSMAX * 2, 3)) return n1 + div((n2 - n1) * (div(HLSMAX * 2, 3) - hue) + div(HLSMAX, 12), div(HLSMAX, 6));
  return n1;
}
