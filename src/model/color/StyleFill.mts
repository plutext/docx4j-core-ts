// A DrawingML fill as far as colour resolution goes: docx4j `org.docx4j.model.color.StyleFill` (CR-034).
import type { ResolvedColor } from './ResolvedColor.mjs';

/**
 * docx4j's `StyleFill.Kind`, as its names: `NONE` (`a:noFill`, or a style reference with index 0), `SOLID`
 * (`a:solidFill`: `color` is its colour), `GRADIENT` (`a:gradFill`), `PATTERN` (`a:pattFill`), `PICTURE`
 * (`a:blipFill`), `GROUP` (`a:grpFill`: the group's fill), `UNKNOWN` (nothing stated, or a style reference whose
 * index is outside the theme's lists).
 */
export type StyleFillKind = 'NONE' | 'SOLID' | 'GRADIENT' | 'PATTERN' | 'PICTURE' | 'GROUP' | 'UNKNOWN';

/**
 * A DrawingML fill as far as colour resolution goes: none, a solid colour (resolved), or a gradient, pattern,
 * picture or group fill, which is reported as such and not guessed at - drawing one is a renderer's business.
 * For a line, the fill is the line's. docx4j's getters are properties here (`getKind()` is `kind`).
 */
export class StyleFill {
  static readonly NONE = new StyleFill('NONE', undefined, undefined);
  static readonly UNKNOWN = new StyleFill('UNKNOWN', undefined, undefined);

  constructor(
    readonly kind: StyleFillKind,
    /** The colour of a `SOLID` fill, else undefined. */
    readonly color: ResolvedColor | undefined,
    /** The fill object (`a:gradFill` and so on), for a renderer; undefined for none. */
    readonly fill: unknown,
  ) {}

  toString(): string {
    return this.kind === 'SOLID' ? `solid ${String(this.color)}` : this.kind.toLowerCase();
  }
}
