import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { applyRunOptions, readRunOptions, type RunOptions, type RunFormatting, type UnderlineType } from '@docx4j/generated-objects-ts/builders/wml';
import type { ChangeTracker } from './tracking.mjs';
import { withinWord, WORD_STORAGE } from './wordRange.mjs';
import { WordColors } from '../color/WordColors.mjs';
import type { ThemeColors } from '../color/ThemeColors.mjs';

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

/** A run's formatting as a `Font` reads it: its properties, and its document font's name. */
export interface RunReading {
  formatting: RunFormatting;
  name: string;
}

/** A paragraph mark's run properties as a `Font` reads them: `w:pPr/w:rPr` without the mark's own revisions. */
export function markHolderOf(p: wml.P): RPrHolder {
  const rPr = p.pPr?.rPr;
  if (!rPr) return {};
  const { ins: _ins, del: _del, moveFrom: _from, moveTo: _to, rPrChange: _change, ...rest } = rPr;
  return { rPr: { ...rest, TYPE_NAME: 'org_docx4j_wml.RPr' } as wml.RPr };
}

/**
 * A run's reading from its properties: in the effective view a theme colour the colour Word draws, given the
 * document's theme colours (`w:themeColor`, tint and shade, over the `w:val` beside it, which is a cache: check 53,
 * CR-008 phase 2), and an automatic colour (`w:val="auto"`, or none) `#000000`, as Office JS reads them (CR-002
 * section 59); and `name` its document font where one is given.
 */
export function readingOf(rPr: wml.RPr | undefined, effective: boolean, name?: string, themeColors?: ThemeColors): RunReading {
  const formatting = readRunOptions(rPr);
  if (effective) {
    const themed = rPr?.color?.themeColor;
    const drawn = themed === undefined || themed === null || themed === 'none' || themeColors === undefined
      ? undefined : WordColors.of(rPr!.color, themeColors);
    if (drawn !== undefined) formatting.color = `#${drawn.toHex()}`;
    else if (formatting.color === '' || formatting.color === 'auto') formatting.color = '#000000';
  }
  return { formatting, name: name ?? formatting.name };
}

/**
 * A subset of Office JS `Word.Font` over the run properties (`w:rPr`) of one or more runs.
 *
 * **Reads report effective formatting** (CR-001 Phase B step 2): each run in scope resolved through
 * the document defaults, the paragraph style and the character style. Pass an `effective` supplier
 * for that; without one the view reads direct formatting, which is what `getFont({ direct: true })`
 * asks for. Where the runs differ a property reads as Office JS reads it (Word 365, checks 52 and 53;
 * CR-002 section 59): `null`, `''` for the name, the colour and the highlight, and `Mixed` for the
 * underline. Office JS's own types say `boolean` and `number` there, and so do these. With no run in
 * scope, a paragraph with no text reads its mark's run properties (`w:pPr/w:rPr`, check 52 case 60),
 * and an empty range its paragraph's style. An automatic colour (`w:val="auto"`, or none) reads
 * `#000000` in the effective view, and a theme colour the colour Word draws (CR-008 phase 2), as Office JS reads them
 * (check 53).
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
    /** The effective `w:rPr` of a run in scope; absent for a direct-formatting view. */
    private readonly effective?: (rPr: wml.RPr | undefined) => wml.RPr | undefined,
    /** The document font of that effective `w:rPr` with its theme reference resolved
     *  (`RunFontSelector.asciiFontName`); absent for a direct-formatting view, and undefined
     *  where the package has no selector yet. */
    private readonly asciiFontName?: (rPr: wml.RPr | undefined) => string | undefined,
    /** The paragraph mark, read where no run is in scope: a paragraph with no text (CR-002 section 59). */
    private readonly mark?: () => RPrHolder | undefined,
    /** The document's theme colours, which a theme colour is read through (CR-008 phase 2); absent for a direct view. */
    private readonly themeColors?: () => ThemeColors | undefined,
  ) {}

  /** Each run in scope, read; the mark's properties where there is none, else the paragraph's style's. */
  protected readings(): RunReading[] {
    let holders = this.holders(false);
    if (holders.length === 0) holders = [this.mark?.() ?? {}];
    const themeColors = this.effective ? this.themeColors?.() : undefined;
    return holders.map((h) => {
      const rPr = this.effective ? this.effective(h.rPr) : h.rPr;
      return readingOf(rPr, this.effective !== undefined, this.asciiFontName?.(rPr), themeColors);
    });
  }

  /** A property every reading agrees on, else `null` (also where nothing is read). */
  private value<K extends keyof RunFormatting>(key: K): RunFormatting[K] | null {
    const all = this.readings().map((r) => r.formatting[key]);
    if (all.length === 0) return null;
    return all.every((v) => v === all[0]) ? all[0]! : null;
  }

  protected apply(opts: RunOptions): void {
    const tracking = this.tracking?.();
    for (const h of this.holders(true)) {
      h.rPr ??= { TYPE_NAME: 'org_docx4j_wml.RPr' };
      // the properties as they are now become w:rPrChange, unless this run is our own insertion
      if (tracking && !tracking.ownInsertions.has(h)) tracking.tracker.recordRPrChange(h.rPr);
      applyRunOptions(h.rPr, opts);
      if (tracking) tracking.tracker.settleRPrChange(h);
    }
  }

  get bold(): boolean { return this.value('bold') as boolean; }
  set bold(v: boolean) { this.apply({ bold: v }); }
  get italic(): boolean { return this.value('italic') as boolean; }
  set italic(v: boolean) { this.apply({ italic: v }); }
  get underline(): UnderlineType { return (this.value('underline') ?? 'Mixed') as UnderlineType; }
  set underline(v: UnderlineType) { this.apply({ underline: v }); }
  get strikeThrough(): boolean { return this.value('strikeThrough') as boolean; }
  set strikeThrough(v: boolean) { this.apply({ strikeThrough: v }); }
  get doubleStrikeThrough(): boolean { return this.value('doubleStrikeThrough') as boolean; }
  set doubleStrikeThrough(v: boolean) { this.apply({ doubleStrikeThrough: v }); }
  get subscript(): boolean { return this.value('subscript') as boolean; }
  set subscript(v: boolean) { this.apply({ subscript: v }); }
  get superscript(): boolean { return this.value('superscript') as boolean; }
  set superscript(v: boolean) { this.apply({ superscript: v }); }
  /**
   * The document font formatting this run's Latin text: `w:rFonts/@w:ascii` of the effective
   * properties, the face `w:asciiTheme` names, else `w:hAnsi`, else the document default
   * (docx4j `RunFontSelector.asciiFontName`). '' where the runs differ, in a direct-formatting
   * view with no `w:ascii`, or where the package's `RunFontSelector` has not been built yet.
   */
  get name(): string {
    const all = this.readings().map((r) => r.name);
    return all.length > 0 && all.every((n) => n === all[0]) ? all[0]! : '';
  }
  set name(v: string) { this.apply({ name: v }); }
  /**
   * Size in points (w:sz is half-points); 0 only in a direct-formatting view with no w:sz; `null` where the runs differ. A
   * size written is held to the range Word 365 stores w:sz in, 0 to 32767.5 pt (CR-002 section 65): outside it, `InvalidArgument`.
   */
  get size(): number { return this.value('size') as number; }
  set size(v: number) { withinWord('Font.size', v, WORD_STORAGE.sz); this.apply({ size: v }); }
  /**
   * '#RRGGBB': in the effective view the colour Word draws, so a theme colour (`w:themeColor`, with its tint or shade)
   * through the document's theme, as Office JS reads it (check 53; CR-008 phase 2), and `#000000` for automatic; in
   * the direct view `w:val` as written, '' when not set; the raw value where it is not a hex triple and no theme
   * colour applies; '' where the runs differ.
   */
  get color(): string { return (this.value('color') ?? '') as string; }
  set color(v: string) { this.apply({ color: v }); }
  /** '#RRGGBB' of the highlight, or null when there is none; '' where the runs differ. */
  get highlightColor(): string | null {
    const all = this.readings().map((r) => r.formatting.highlightColor);
    if (all.length === 0) return null;
    return all.every((v) => v === all[0]) ? all[0]! : '';
  }
  set highlightColor(v: string | null) { this.apply({ highlightColor: v }); }
}

/**
 * The `Font` of a range across paragraphs (CR-014): each run read in its own paragraph, as `Font` reads them. A write
 * goes to `write` (phase 2: every run and mark in the range, as Word formats them).
 */
export class MixedFont extends Font {
  constructor(private readonly reads: () => RunReading[], private readonly write: (opts: RunOptions) => void) {
    super(() => []);
  }

  protected override readings(): RunReading[] {
    return this.reads();
  }

  protected override apply(opts: RunOptions): void {
    this.write(opts);
  }
}
