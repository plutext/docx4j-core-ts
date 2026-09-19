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
    private readonly holders: () => RPrHolder[],
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
    const direct = this.holders()[0]?.rPr;
    return this.effective ? this.effective(direct) : direct;
  }

  private read(): RunFormatting {
    return readRunOptions(this.effectiveRPr());
  }

  private apply(opts: RunOptions): void {
    const tracking = this.tracking?.();
    for (const h of this.holders()) {
      h.rPr ??= { TYPE_NAME: 'org_docx4j_wml.RPr' };
      // the properties as they are now become w:rPrChange, unless this run is our own insertion
      if (tracking && !tracking.ownInsertions.has(h)) tracking.tracker.recordRPrChange(h.rPr);
      applyRunOptions(h.rPr, opts);
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
