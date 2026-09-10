import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { applyRunOptions, readRunOptions, type RunOptions, type RunFormatting, type UnderlineType } from '@docx4j/generated-objects-ts/builders/wml';

export type { UnderlineType };

/** Something with run properties: a run, or the paragraph mark. */
export interface RPrHolder {
  rPr?: wml.RPr;
}

/**
 * A subset of Office JS `Word.Font`, over the direct run properties (w:rPr) of one or more
 * runs. Reads report the first run's direct formatting (effective formatting through styles
 * is CR-001 Phase B); writes apply to every run in scope. Both go through the objects
 * package's one name-to-w:rPr mapping (`builders/wml` applyRunOptions / readRunOptions),
 * which its `r(text, opts)` builder shares.
 */
export class Font {
  constructor(private readonly holders: () => RPrHolder[]) {}

  private read(): RunFormatting {
    return readRunOptions(this.holders()[0]?.rPr);
  }

  private apply(opts: RunOptions): void {
    for (const h of this.holders()) {
      h.rPr ??= { TYPE_NAME: 'org_docx4j_wml.RPr' };
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
  /** The ASCII font name (w:rFonts/@w:ascii); '' when not set directly. */
  get name(): string { return this.read().name; }
  set name(v: string) { this.apply({ name: v }); }
  /** Size in points (w:sz is half-points); 0 when not set directly. */
  get size(): number { return this.read().size; }
  set size(v: number) { this.apply({ size: v }); }
  /** '#RRGGBB', or the raw value ('auto', a theme colour name) when it is not a hex triple; '' when not set. */
  get color(): string { return this.read().color; }
  set color(v: string) { this.apply({ color: v }); }
  /** '#RRGGBB' of the highlight, or null when there is none. */
  get highlightColor(): string | null { return this.read().highlightColor; }
  set highlightColor(v: string | null) { this.apply({ highlightColor: v }); }
}
