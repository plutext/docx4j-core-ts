import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { applyRunOptions, readRunOptions, type RunOptions, type RunFormatting, type UnderlineType } from '@docx4j/generated-objects-ts/builders/wml';
import type { ChangeTracker } from './tracking.mjs';

export type { UnderlineType };

/** Something with run properties: a run, or the paragraph mark. */
export interface RPrHolder {
  rPr?: wml.RPr;
}

/** What a `Font` needs to write a formatting revision (CR-002 phase F); undefined when tracking is off. */
export interface FontTracking {
  tracker: ChangeTracker;
  /** The runs that are this author's own insertion: Word changes those in place, with no w:rPrChange. */
  ownInsertions: ReadonlySet<object>;
}

/**
 * A subset of Office JS `Word.Font` over the run properties (`w:rPr`) of one or more runs.
 *
 * **Reads report effective formatting** (CR-001 Phase B step 2): the first run in scope
 * resolved through the document defaults, the paragraph style and the character style, which
 * is what Office JS reports. Pass an `effective` supplier for that; without one the view falls
 * back to direct formatting, which is what `getFont({ direct: true })` asks for.
 *
 * **Writes are always direct**, and apply to every run in scope. Both directions go through
 * the objects package's one name-to-`w:rPr` mapping (`builders/wml` applyRunOptions /
 * readRunOptions), which its `r(text, opts)` builder shares.
 *
 * `name` is effective too since step 4: it is `RunFontSelector.asciiFontName` of the resolved
 * properties, so a run whose font comes from the theme (`w:asciiTheme`) reads the face the
 * theme names, and one that names no font anywhere reads the document default rather than ''.
 * `getFont({ direct: true }).name` is still `w:rFonts/@w:ascii` of the run itself, and a
 * package whose selector has not been built yet falls back to that.
 */
export class Font {
  constructor(
    /** The runs in scope; `write` is true for a write, which may split runs to have them, and false for a read, which must not. */
    private readonly holders: (write: boolean) => RPrHolder[],
    /** The change tracker and this author's own insertions, when the package is tracking changes. */
    private readonly tracking?: () => FontTracking | undefined,
    /** The effective `w:rPr` of the first run in scope; absent for a direct-formatting view. */
    private readonly effective?: (rPr: wml.RPr | undefined) => wml.RPr | undefined,
    /** The document font of that effective `w:rPr` with its theme reference resolved
     *  (`RunFontSelector.asciiFontName`); absent for a direct-formatting view, and undefined
     *  where the package has no selector yet. */
    private readonly asciiFontName?: (rPr: wml.RPr | undefined) => string | undefined,
  ) {}

  private effectiveRPr(): wml.RPr | undefined {
    const direct = this.holders(false)[0]?.rPr;
    return this.effective ? this.effective(direct) : direct;
  }

  private read(): RunFormatting {
    return readRunOptions(this.effectiveRPr());
  }

  private apply(opts: RunOptions): void {
    const tracking = this.tracking?.();
    for (const h of this.holders(true)) {
      h.rPr ??= { TYPE_NAME: 'org_docx4j_wml.RPr' };
      // the properties as they are now become w:rPrChange, unless this run is our own insertion
      if (tracking && !tracking.ownInsertions.has(h)) tracking.tracker.recordRPrChange(h.rPr);
      applyRunOptions(h.rPr, opts);
      if (tracking) tracking.tracker.settleRPrChange(h);
    }
  }

  get bold(): boolean { return this.read().bold; }
  set bold(v: boolean) { this.apply({ bold: v }); }
  get italic(): boolean { return this.read().italic; }
  set italic(v: boolean) { this.apply({ italic: v }); }
  get underline(): UnderlineType { return this.read().underline; }
  set underline(v: UnderlineType) { this.apply({ underline: v }); }
  get strikeThrough(): boolean { return this.read().strikeThrough; }
  set strikeThrough(v: boolean) { this.apply({ strikeThrough: v }); }
  get doubleStrikeThrough(): boolean { return this.read().doubleStrikeThrough; }
  set doubleStrikeThrough(v: boolean) { this.apply({ doubleStrikeThrough: v }); }
  get subscript(): boolean { return this.read().subscript; }
  set subscript(v: boolean) { this.apply({ subscript: v }); }
  get superscript(): boolean { return this.read().superscript; }
  set superscript(v: boolean) { this.apply({ superscript: v }); }
  /**
   * The document font formatting this run's Latin text: `w:rFonts/@w:ascii` of the effective
   * properties, the face `w:asciiTheme` names, else `w:hAnsi`, else the document default
   * (docx4j `RunFontSelector.asciiFontName`). '' only in a direct-formatting view with no
   * `w:ascii`, or where the package's `RunFontSelector` has not been built yet.
   */
  get name(): string {
    if (this.asciiFontName !== undefined) {
      const resolved = this.asciiFontName(this.effectiveRPr());
      if (resolved !== undefined) return resolved;
    }
    return this.read().name;
  }
  set name(v: string) { this.apply({ name: v }); }
  /** Size in points (w:sz is half-points); 0 only in a direct-formatting view with no w:sz. */
  get size(): number { return this.read().size; }
  set size(v: number) { this.apply({ size: v }); }
  /** '#RRGGBB', or the raw value ('auto', a theme colour name) when it is not a hex triple; '' when not set. */
  get color(): string { return this.read().color; }
  set color(v: string) { this.apply({ color: v }); }
  /** '#RRGGBB' of the highlight, or null when there is none. */
  get highlightColor(): string | null { return this.read().highlightColor; }
  set highlightColor(v: string | null) { this.apply({ highlightColor: v }); }
}

/** A run's formatting as a `MixedFont` reads it: its properties, and its document font's name. */
export interface RunReading {
  formatting: RunFormatting;
  name: string;
}

/**
 * The `Font` of a range across paragraphs (CR-014 phase 1): each run read in its own paragraph, and a property
 * where the runs differ read as Office JS reads it (checks 52 and 53): `null`, `''` for the name, the colour and
 * the highlight, and `Mixed` for the underline. Office JS's own types say `boolean` and `number` there too. Its writes are CR-014 phase 2's; until
 * then `refuse` throws.
 */
export class MixedFont extends Font {
  constructor(private readonly readings: () => RunReading[], private readonly refuse: () => void) {
    super(() => []);
  }

  private value<K extends keyof RunFormatting>(key: K): RunFormatting[K] | null {
    const all = this.readings().map((r) => r.formatting[key]);
    if (all.length === 0) return null;
    return all.every((v) => v === all[0]) ? all[0]! : null;
  }

  private write(): never {
    this.refuse();
    throw new Error('unreachable');
  }

  override get bold(): boolean { return this.value('bold') as boolean; }
  override set bold(_v: boolean) { this.write(); }
  override get italic(): boolean { return this.value('italic') as boolean; }
  override set italic(_v: boolean) { this.write(); }
  override get underline(): UnderlineType { return (this.value('underline') ?? 'Mixed') as UnderlineType; }
  override set underline(_v: UnderlineType) { this.write(); }
  override get strikeThrough(): boolean { return this.value('strikeThrough') as boolean; }
  override set strikeThrough(_v: boolean) { this.write(); }
  override get doubleStrikeThrough(): boolean { return this.value('doubleStrikeThrough') as boolean; }
  override set doubleStrikeThrough(_v: boolean) { this.write(); }
  override get subscript(): boolean { return this.value('subscript') as boolean; }
  override set subscript(_v: boolean) { this.write(); }
  override get superscript(): boolean { return this.value('superscript') as boolean; }
  override set superscript(_v: boolean) { this.write(); }
  override get name(): string {
    const all = this.readings().map((r) => r.name);
    return all.length > 0 && all.every((n) => n === all[0]) ? all[0]! : '';
  }
  override set name(_v: string) { this.write(); }
  override get size(): number { return this.value('size') as number; }
  override set size(_v: number) { this.write(); }
  override get color(): string { return (this.value('color') ?? '') as string; }
  override set color(_v: string) { this.write(); }
  override get highlightColor(): string | null {
    const all = this.readings().map((r) => r.formatting.highlightColor);
    if (all.length === 0) return null;
    return all.every((v) => v === all[0]) ? all[0]! : '';
  }
  override set highlightColor(_v: string | null) { this.write(); }
}
