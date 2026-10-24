// A colour as Office draws it: docx4j `org.docx4j.model.color.ResolvedColor` (CR-034).

function clamp(v: number, max: number): number {
  return v < 0 ? 0 : v > max ? max : v;
}

/**
 * A colour as Office draws it: red, green and blue in sRGB (0 to 255) and an alpha in OOXML's unit
 * (0 to 100000, where 100000 is opaque). Immutable.
 *
 * docx4j's getters are properties here (`getRed()` is `red`); `equals` compares the four values.
 */
export class ResolvedColor {
  /** OOXML's opaque alpha, 100% in thousandths of a percent. */
  static readonly OPAQUE = 100000;

  static readonly BLACK = new ResolvedColor(0, 0, 0, ResolvedColor.OPAQUE);
  static readonly WHITE = new ResolvedColor(255, 255, 255, ResolvedColor.OPAQUE);

  readonly red: number;
  readonly green: number;
  readonly blue: number;
  /** 0 (transparent) to 100000 (opaque). */
  readonly alpha: number;

  private constructor(red: number, green: number, blue: number, alpha: number) {
    // Java's int arguments: a caller here could pass a fraction
    this.red = clamp(Math.trunc(red), 255);
    this.green = clamp(Math.trunc(green), 255);
    this.blue = clamp(Math.trunc(blue), 255);
    this.alpha = clamp(Math.trunc(alpha), ResolvedColor.OPAQUE);
  }

  /** An opaque colour; each channel is clamped to 0 to 255. */
  static rgb(red: number, green: number, blue: number): ResolvedColor {
    return new ResolvedColor(red, green, blue, ResolvedColor.OPAQUE);
  }

  /**
   * An opaque colour from six hex digits (RRGGBB, either case; a leading '#' is allowed).
   *
   * @returns undefined where the string is not six hex digits ("auto", for instance)
   */
  static ofHex(hex: string | null | undefined): ResolvedColor | undefined {
    if (hex === undefined || hex === null) return undefined;
    const h = hex.startsWith('#') ? hex.substring(1) : hex;
    if (!/^[0-9A-Fa-f]{6}$/.test(h)) return undefined;
    const v = Number.parseInt(h, 16);
    return ResolvedColor.rgb((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF);
  }

  /** This colour with another alpha (0 to 100000). */
  withAlpha(alpha: number): ResolvedColor {
    return new ResolvedColor(this.red, this.green, this.blue, alpha);
  }

  get isOpaque(): boolean {
    return this.alpha === ResolvedColor.OPAQUE;
  }

  /** RRGGBB, upper case, without the alpha. */
  toHex(): string {
    return [this.red, this.green, this.blue].map((c) => c.toString(16).toUpperCase().padStart(2, '0')).join('');
  }

  /** Whether each channel is within `units` of the other colour's (alpha not compared). */
  within(other: ResolvedColor | undefined, units: number): boolean {
    return other !== undefined && Math.abs(this.red - other.red) <= units && Math.abs(this.green - other.green) <= units
      && Math.abs(this.blue - other.blue) <= units;
  }

  equals(other: unknown): boolean {
    return other instanceof ResolvedColor && this.red === other.red && this.green === other.green
      && this.blue === other.blue && this.alpha === other.alpha;
  }

  /** RRGGBB, or RRGGBB/alpha where not opaque. */
  toString(): string {
    return this.isOpaque ? this.toHex() : `${this.toHex()}/${this.alpha}`;
  }
}
