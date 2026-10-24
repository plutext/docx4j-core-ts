// DrawingML colours (ISO/IEC 29500-1 section 20.1.2.3) resolved as Office draws them: docx4j
// `org.docx4j.model.color.DrawingMLColors` (CR-034), with the luminance rounding of its follow-up (`691620fb7`).
import type * as dml from '@docx4j/generated-objects-ts/modules/org_docx4j_dml';
import type { Jsonix } from '@docx4j/generated-objects-ts';
import * as ColorMath from './ColorMath.mjs';
import { PRESET, SYSTEM } from './ColorTables.mjs';
import { ResolvedColor } from './ResolvedColor.mjs';
import { StyleFill } from './StyleFill.mjs';
import type { ThemeColors } from './ThemeColors.mjs';

/** One of the six colour kinds. */
type ColourKind = dml.CTSRgbColor | dml.CTSchemeColor | dml.CTScRgbColor | dml.CTHslColor | dml.CTSystemColor | dml.CTPresetColor;
/** What holds one of them as `srgbClr`, `schemeClr`, ...: a scheme slot, a solid fill, a style or font reference. */
type ColourHolder = Partial<Pick<dml.CTColor, 'srgbClr' | 'schemeClr' | 'scrgbClr' | 'hslClr' | 'sysClr' | 'prstClr'>>;
type Transform = Jsonix.TypedNamedValue<{ val?: number } | object>;

const KIND_NAMES: Readonly<Record<string, string>> = {
  srgbClr: 'org_docx4j_dml.CTSRgbColor', schemeClr: 'org_docx4j_dml.CTSchemeColor', scrgbClr: 'org_docx4j_dml.CTScRgbColor',
  hslClr: 'org_docx4j_dml.CTHslColor', sysClr: 'org_docx4j_dml.CTSystemColor', prstClr: 'org_docx4j_dml.CTPresetColor',
};
const HOLDERS = new Set(['org_docx4j_dml.CTColor', 'org_docx4j_dml.CTSolidColorFillProperties',
  'org_docx4j_dml.CTStyleMatrixReference', 'org_docx4j_dml.CTFontReference']);
const FILL_NAMES: Readonly<Record<string, string>> = {
  noFill: 'org_docx4j_dml.CTNoFillProperties', solidFill: 'org_docx4j_dml.CTSolidColorFillProperties',
  gradFill: 'org_docx4j_dml.CTGradientFillProperties', pattFill: 'org_docx4j_dml.CTPatternFillProperties',
  blipFill: 'org_docx4j_dml.CTBlipFillProperties', grpFill: 'org_docx4j_dml.CTGroupFillProperties',
};

const typeNameOf = (o: unknown): string | undefined => (o as { TYPE_NAME?: string } | undefined)?.TYPE_NAME;

/** A JAXBElement's counterpart, `{ name, value }`: the value, with the type its element name gives where it has none. */
function unwrap(o: unknown, names: Readonly<Record<string, string>>): { value: unknown; type: string | undefined } {
  const named = o as Partial<Jsonix.TypedNamedValue<unknown>> | undefined;
  if (named !== undefined && named !== null && typeof named === 'object' && named.name !== undefined && 'value' in named) {
    const local = typeof named.name === 'string' ? named.name : named.name.localPart;
    return { value: named.value, type: typeNameOf(named.value) ?? names[local] };
  }
  return { value: o, type: typeNameOf(o) };
}

const holds = (o: object): boolean =>
  ['srgbClr', 'schemeClr', 'scrgbClr', 'hslClr', 'sysClr', 'prstClr'].some((k) => (o as Record<string, unknown>)[k] !== undefined && (o as Record<string, unknown>)[k] !== null);

/** The colour element inside a holder, or the element itself, with its kind's type name. */
function kind(o: unknown): { value: ColourKind; type: string } | undefined {
  const { value, type } = unwrap(o, KIND_NAMES);
  if (value === undefined || value === null || typeof value !== 'object') return undefined;
  if ((type === undefined || HOLDERS.has(type)) && holds(value)) {
    const h = value as ColourHolder;
    for (const [key, name] of [['srgbClr', KIND_NAMES.srgbClr], ['schemeClr', KIND_NAMES.schemeClr], ['scrgbClr', KIND_NAMES.scrgbClr],
      ['hslClr', KIND_NAMES.hslClr], ['sysClr', KIND_NAMES.sysClr], ['prstClr', KIND_NAMES.prstClr]] as const) {
      const c = h[key];
      if (c !== undefined && c !== null) return { value: c, type: name! };
    }
  }
  if (type !== undefined && Object.values(KIND_NAMES).includes(type)) return { value: value as ColourKind, type };
  return undefined;
}

/** hexBinary as the objects model has it at run time (byte numbers), or as its declarations say (hex digits). */
export function bytesOf(v: unknown): number[] | undefined {
  if (v === undefined || v === null) return undefined;
  if (typeof v === 'string') {
    if (!/^([0-9A-Fa-f]{2})*$/.test(v)) return undefined;
    const out: number[] = [];
    for (let i = 0; i < v.length; i += 2) out.push(Number.parseInt(v.substring(i, i + 2), 16));
    return out;
  }
  if (Array.isArray(v) || ArrayBuffer.isView(v)) return Array.from(v as ArrayLike<number>, (b) => b & 0xFF);
  return undefined;
}

function base(c: ColourKind, type: string, theme: ThemeColors | undefined): ResolvedColor | undefined {
  switch (type) {
    case KIND_NAMES.srgbClr: {
      const val = (c as dml.CTSRgbColor).val as unknown;
      if (typeof val === 'string') return ResolvedColor.ofHex(val);
      const b = bytesOf(val);
      return b !== undefined && b.length === 3 ? ResolvedColor.rgb(b[0]!, b[1]!, b[2]!) : undefined;
    }
    case KIND_NAMES.schemeClr: {
      const s = c as dml.CTSchemeColor;
      // docx4j logs at debug "Couldn't resolve schemeClr ..."
      return theme === undefined || s.val === undefined || s.val === null ? undefined : theme.schemeColor(s.val);
    }
    case KIND_NAMES.scrgbClr: {
      const s = c as dml.CTScRgbColor;
      return ResolvedColor.rgb(ColorMath.channel(ColorMath.toSrgb(s.r / 100000.0)),
        ColorMath.channel(ColorMath.toSrgb(s.g / 100000.0)), ColorMath.channel(ColorMath.toSrgb(s.b / 100000.0)));
    }
    case KIND_NAMES.hslClr: {
      const h = c as dml.CTHslColor;
      const rgb = ColorMath.hslToRgb(h.hue / 21600000.0, h.sat / 100000.0, h.lum / 100000.0);
      return ResolvedColor.rgb(ColorMath.channel(rgb[0]), ColorMath.channel(rgb[1]), ColorMath.channel(rgb[2]));
    }
    case KIND_NAMES.sysClr: {
      const s = c as dml.CTSystemColor;
      // Office draws the machine's colour for the name and ignores lastClr (docx4j CR-034 section 11)
      const hex = s.val === undefined || s.val === null ? undefined : SYSTEM.get(s.val);
      if (hex !== undefined) return ResolvedColor.ofHex(hex);
      const last = bytesOf(s.lastClr);
      return last !== undefined && last.length === 3 ? ResolvedColor.rgb(last[0]!, last[1]!, last[2]!) : undefined;
    }
    case KIND_NAMES.prstClr: {
      const p = c as dml.CTPresetColor;
      return p.val === undefined || p.val === null ? undefined : ResolvedColor.ofHex(PRESET.get(p.val));
    }
    default: return undefined;
  }
}

/** The transforms' value, in thousandths of a percent or 60000ths of a degree; 0 for those without one. */
function value(t: unknown): number {
  const v = (t as { val?: unknown } | undefined)?.val;
  return typeof v === 'number' ? v : 0;
}

/** A colour with transforms applied in document order. */
function transform(baseColor: ResolvedColor, transforms: readonly Transform[] | undefined): ResolvedColor {
  if (transforms === undefined || transforms === null || transforms.length === 0) return baseColor;
  let r = baseColor.red / 255.0, g = baseColor.green / 255.0, b = baseColor.blue / 255.0;
  let alpha = baseColor.alpha / 100000.0;
  let lumLast = false; // the last change to the colour a luminance transform: it rounds half down
  for (const t of transforms) {
    const name = typeof t.name === 'string' ? t.name : t.name.localPart;
    const v = value(t.value) / 100000.0;
    if (!name.startsWith('alpha')) lumLast = name.startsWith('lum');
    switch (name) {
      case 'tint':
        r = ColorMath.toSrgb(1 - (1 - ColorMath.toLinear(r)) * v);
        g = ColorMath.toSrgb(1 - (1 - ColorMath.toLinear(g)) * v);
        b = ColorMath.toSrgb(1 - (1 - ColorMath.toLinear(b)) * v);
        break;
      case 'shade':
        r = ColorMath.toSrgb(ColorMath.toLinear(r) * v);
        g = ColorMath.toSrgb(ColorMath.toLinear(g) * v);
        b = ColorMath.toSrgb(ColorMath.toLinear(b) * v);
        break;
      case 'inv':
        r = ColorMath.toSrgb(1 - ColorMath.toLinear(r));
        g = ColorMath.toSrgb(1 - ColorMath.toLinear(g));
        b = ColorMath.toSrgb(1 - ColorMath.toLinear(b));
        break;
      case 'gray': {
        const y = 0.2126 * r + 0.7152 * g + 0.0722 * b;
        r = g = b = ColorMath.clamp01(y);
        break;
      }
      case 'gamma':
        r = ColorMath.toSrgb(r);
        g = ColorMath.toSrgb(g);
        b = ColorMath.toSrgb(b);
        break;
      case 'invGamma':
        r = ColorMath.toLinear(r);
        g = ColorMath.toLinear(g);
        b = ColorMath.toLinear(b);
        break;
      case 'red': case 'redOff': case 'redMod':
        r = channel(name, r, v);
        break;
      case 'green': case 'greenOff': case 'greenMod':
        g = channel(name, g, v);
        break;
      case 'blue': case 'blueOff': case 'blueMod':
        b = channel(name, b, v);
        break;
      case 'alpha':
        alpha = v;
        break;
      case 'alphaOff':
        alpha += v;
        break;
      case 'alphaMod':
        alpha *= v;
        break;
      case 'comp': case 'hue': case 'hueOff': case 'hueMod':
      case 'sat': case 'satOff': case 'satMod':
      case 'lum': case 'lumOff': case 'lumMod': {
        let [h, s, l] = ColorMath.rgbToHsl(r, g, b);
        const raw = value(t.value);
        switch (name) {
          case 'comp': h += 0.5; break;
          case 'hue': h = raw / 21600000.0; break;
          case 'hueOff': h += raw / 21600000.0; break;
          case 'hueMod': h *= v; break;
          case 'sat': s = v; break;
          case 'satOff': s += v; break;
          case 'satMod': s *= v; break;
          case 'lum': l = v; break;
          case 'lumOff': l += v; break;
          // l x val / 100000 in this order, not l *= v: with hslToRgbFoley and channelHalfDown, which way an exact
          // half rounds follows this form's floating-point error, as Office's does (docx4j CR-034 section 11, follow-up)
          case 'lumMod': l = l * raw / 100000.0; break;
          default: break;
        }
        h = h % 1;
        if (h < 0) h += 1;
        const rgb = lumLast ? ColorMath.hslToRgbFoley(h, s, ColorMath.clamp01(l)) : ColorMath.hslToRgb(h, s, ColorMath.clamp01(l));
        r = rgb[0];
        g = rgb[1];
        b = rgb[2];
        break;
      }
      default:
        // docx4j logs at debug "Ignoring an unknown colour transform ..."
        break;
    }
  }
  if (lumLast) {
    return ResolvedColor.rgb(ColorMath.channelHalfDown(r), ColorMath.channelHalfDown(g), ColorMath.channelHalfDown(b))
      .withAlpha(Math.round(ColorMath.clamp01(alpha) * 100000));
  }
  return ResolvedColor.rgb(ColorMath.channel(r), ColorMath.channel(g), ColorMath.channel(b))
    .withAlpha(Math.round(ColorMath.clamp01(alpha) * 100000));
}

/** red, redOff, redMod (and green, blue): the channel in linear RGB */
function channel(name: string, c: number, v: number): number {
  let l = ColorMath.toLinear(c);
  if (name.endsWith('Off')) l += v;
  else if (name.endsWith('Mod')) l *= v;
  else l = v;
  return ColorMath.toSrgb(l);
}

/** The fill's type name: its own, its element name's (`{ name, value }`), or a solid fill's by its colour. */
function fillOf(o: unknown): { value: object; type: string | undefined } | undefined {
  const { value, type } = unwrap(o, FILL_NAMES);
  if (value === undefined || value === null || typeof value !== 'object') return undefined;
  return { value, type: type ?? (holds(value) ? FILL_NAMES.solidFill : undefined) };
}

/**
 * DrawingML colours (ISO/IEC 29500-1 section 20.1.2.3) resolved as Office draws them. The six kinds: `srgbClr`,
 * `scrgbClr` (linear percentages), `hslClr`, `sysClr` (the machine's colour by name, not `lastClr`), `prstClr` and
 * `schemeClr` (through {@link ThemeColors}). Their transforms apply in document order, in the spaces Word and
 * PowerPoint use, measured on 403 cases each (docx4j CR-034 section 11, P3):
 * - `tint`, `shade`, `inv`, `red`/`green`/`blue` and their `Off`/`Mod`: linear RGB; `gamma` and `invGamma`: the
 *   sRGB transfer function each way;
 * - `hue`, `sat`, `lum` and their `Off`/`Mod`, `comp`: HSL of the sRGB colour, the saturation not clamped (the
 *   channels are, afterwards); a colour whose last change is a luminance transform converted back as Foley and van
 *   Dam do and rounded half down, as Office rounds it (so `tx1` at `lumMod 50000 lumOff 50000` is 7F7F7F);
 * - `gray`: Rec.709 luma of the sRGB channels;
 * - `alpha`, `alphaOff`, `alphaMod`: the alpha.
 *
 * MS-OI29500 says `lumOff` and `satOff` are not supported by Office; Office 365 applies both as above in every case
 * measured. Known differences from Office: a saturation set on a grey (Office gives a grey's hue about 40 degrees),
 * and some hue, saturation and luminance cases one unit off.
 *
 * An object is known by its `TYPE_NAME` (which the objects package's unmarshal and factories give it), or, wrapped
 * as `{ name, value }`, by its element name; a holder of a colour also by its `srgbClr`, `schemeClr`, ... member.
 * docx4j's `null` is `undefined` here.
 */
export const DrawingMLColors = {
  /**
   * The colour a DrawingML colour element, or an element holding one, resolves to.
   *
   * @param colour one of the six kinds (`CTSchemeColor`, `CTSRgbColor`, ...), or a holder of one: `CTColor` (a
   *   scheme slot), `CTSolidColorFillProperties`, `CTStyleMatrixReference`, `CTFontReference`; `{ name, value }`
   *   is unwrapped
   * @param theme the theme and colour map for `schemeClr`; may be undefined where there is none
   * @returns undefined where there is no colour, or a `schemeClr` cannot be resolved
   */
  resolve(colour: unknown, theme: ThemeColors | undefined): ResolvedColor | undefined {
    const c = kind(colour);
    if (c === undefined) return undefined;
    const b = base(c.value, c.type, theme);
    if (b === undefined) return undefined;
    return transform(b, c.value.egColorTransform as Transform[] | undefined);
  },

  /** A fill: `a:noFill`, `a:solidFill` (its colour resolved), or a gradient, pattern, picture or group fill reported by kind. */
  fill(fillProperties: unknown, theme: ThemeColors | undefined): StyleFill {
    const f = fillOf(fillProperties);
    return f === undefined ? StyleFill.UNKNOWN : fillOfType(f.type, f.value, theme);
  },

  /** A line's fill: none, solid, gradient or pattern. */
  lineFill(ln: dml.CTLineProperties | undefined | null, theme: ThemeColors | undefined): StyleFill {
    if (ln === undefined || ln === null) return StyleFill.UNKNOWN;
    if (ln.noFill !== undefined && ln.noFill !== null) return StyleFill.NONE;
    if (ln.solidFill !== undefined && ln.solidFill !== null) return fillOfType(FILL_NAMES.solidFill, ln.solidFill, theme);
    if (ln.gradFill !== undefined && ln.gradFill !== null) return fillOfType(FILL_NAMES.gradFill, ln.gradFill, theme);
    if (ln.pattFill !== undefined && ln.pattFill !== null) return fillOfType(FILL_NAMES.pattFill, ln.pattFill, theme);
    return StyleFill.UNKNOWN;
  },

  /**
   * A shape style's `a:fillRef`: index 0 is no fill, 1 to 999 the theme's `a:fillStyleLst`, 1001 and up its
   * `a:bgFillStyleLst`, with `phClr` the reference's own colour.
   */
  fillReference(ref: dml.CTStyleMatrixReference | undefined | null, theme: ThemeColors | undefined): StyleFill {
    if (ref === undefined || ref === null) return StyleFill.UNKNOWN;
    const idx = ref.idx;
    if (idx === 0) return StyleFill.NONE;
    const fmt = theme?.formatScheme;
    if (theme === undefined || fmt === undefined) return StyleFill.UNKNOWN;
    const list = idx >= 1001 ? fmt.bgFillStyleLst?.egFillProperties : fmt.fillStyleLst?.egFillProperties;
    const i = Math.trunc(idx >= 1001 ? idx - 1001 : idx - 1);
    if (list === undefined || list === null || i < 0 || i >= list.length) return StyleFill.UNKNOWN;
    return DrawingMLColors.fill(list[i], theme.withPlaceholder(DrawingMLColors.resolve(holder(ref), theme)));
  },

  /** A shape style's `a:lnRef`: index 0 is no line, else the theme's `a:lnStyleLst` entry's fill. */
  lineReference(ref: dml.CTStyleMatrixReference | undefined | null, theme: ThemeColors | undefined): StyleFill {
    if (ref === undefined || ref === null) return StyleFill.UNKNOWN;
    const idx = ref.idx;
    if (idx === 0) return StyleFill.NONE;
    const fmt = theme?.formatScheme;
    if (theme === undefined || fmt === undefined || fmt.lnStyleLst === undefined || fmt.lnStyleLst === null) return StyleFill.UNKNOWN;
    const list = fmt.lnStyleLst.ln ?? [];
    const i = Math.trunc(idx) - 1;
    if (i < 0 || i >= list.length) return StyleFill.UNKNOWN;
    return DrawingMLColors.lineFill(list[i], theme.withPlaceholder(DrawingMLColors.resolve(holder(ref), theme)));
  },

  /** A shape style's `a:fontRef`: its colour, the text's where the text states none. */
  fontReference(ref: dml.CTFontReference | undefined | null, theme: ThemeColors | undefined): ResolvedColor | undefined {
    return ref === undefined || ref === null ? undefined : DrawingMLColors.resolve(holder(ref), theme);
  },
};

function fillOfType(type: string | undefined, f: object, theme: ThemeColors | undefined): StyleFill {
  switch (type) {
    case FILL_NAMES.noFill: return StyleFill.NONE;
    case FILL_NAMES.solidFill: return new StyleFill('SOLID', DrawingMLColors.resolve(holder(f as ColourHolder), theme), f);
    case FILL_NAMES.gradFill: return new StyleFill('GRADIENT', undefined, f);
    case FILL_NAMES.pattFill: return new StyleFill('PATTERN', undefined, f);
    case FILL_NAMES.blipFill: return new StyleFill('PICTURE', undefined, f);
    case FILL_NAMES.grpFill: return new StyleFill('GROUP', undefined, f);
    default: return StyleFill.UNKNOWN;
  }
}

/** A solid fill, a style or a font reference as a holder of its colour, whatever its TYPE_NAME. */
function holder(ref: ColourHolder): ColourHolder {
  return { srgbClr: ref.srgbClr, schemeClr: ref.schemeClr, scrgbClr: ref.scrgbClr, hslClr: ref.hslClr, sysClr: ref.sysClr, prstClr: ref.prstClr };
}
