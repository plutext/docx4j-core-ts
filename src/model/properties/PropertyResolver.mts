// docx4j `org.docx4j.model.PropertyResolver` (CR-001 section 6.1 as CR-015 revised it).
//
// Works out the properties which actually apply to a paragraph, a run, a paragraph mark or a
// table, following the order ECMA-376 17.7.2 gives and Word applies:
//
//   effectivePPr(direct)       = docDefaults.pPr ⊕ chainPPr(styleOf(direct)) ⊕ direct
//   effectiveRPr(direct, pPr)  = docDefaults.rPr ⊕ chainRPr(styleOf(pPr)) ⊕ chainRPr(direct.rStyle) ⊕ direct
//   paragraphMarkRPr(pPr)      = docDefaults.rPr ⊕ chainRPr(styleOf(pPr)) ⊕ pPr.rPr
//   tableStyle(tblPr)          = built-in Normal Table (where its chain reaches the default
//                                table style, or it names none) ⊕ the chain ⊕ tblPr
//
// where `styleOf(pPr)` is the paragraph's `w:pStyle` if it names a style that exists, else the
// `w:default="1"` paragraph style (Word writes no `w:pStyle` for it and treats a missing style
// as it), and `chainPPr`/`chainRPr` are a style's `w:basedOn` chain merged root-first *without*
// the document defaults, cached per style id.  The merging itself is `styleUtil.mts`, driven by
// the property catalogue.  Numbering level indents are folded in per layer; table styles are
// not applied to paragraphs here (the resolver is handed a `w:pPr` and does not know the
// table).
//
// **Live objects.** What the style overloads and `getEffectivePPr(pPr)` return is cached and
// shared: copy it before changing it.  The cached objects share no leaf with the styles part,
// and resolution writes nothing into the part (test/properties.test.mjs proves it by
// marshalling the part before and after).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy } from '@docx4j/generated-objects-ts';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { log } from './log.mjs';
import { NumberingLevels } from './numberingInd.mjs';
import {
  applyPPrBase, applyRPr, applyStyle, applyStyleLevel, applyTblPr,
  hasDirectFormattingPPr, hasDirectFormattingRPr, isCyclic, type NumberingIndents,
} from './styleUtil.mjs';

/** The parts the resolver reads. Structural, so this module imports no part class. */
export interface ResolverSource {
  /** The styles part's `w:styles`, or undefined when the package has none. */
  styles?: wml.Styles;
  /** The numbering part's `w:numbering`, or undefined. */
  numbering?: wml.Numbering;
}

/** As much of an `XmlPart<T>` as the resolver uses. */
interface ReadablePart<T> {
  readContents(): Promise<T>;
  readonly isUnmarshalled: boolean;
  readonly contents: T;
}

/** What `create` accepts: a WordprocessingMLPackage, or the trees on their own. */
interface PackageLike {
  getMainDocumentPart(): {
    styleDefinitionsPart?: ReadablePart<wml.Styles>;
    numberingDefinitionsPart?: ReadablePart<wml.Numbering>;
  };
}

/** The cache key for "no style" (a styles part with no default paragraph style). */
const NO_STYLE = '';

/** 108 twips (0.08in): the left and right cell margin of Word's built-in Normal Table. */
export const WORD_DEFAULT_CELL_MARGIN_TWIPS = 108;

const HEADING_NAME = /^heading ([1-9])$/i;
const HEADING_STYLE = 'Heading';

/**
 * The outline level (1-9) of a built-in heading style, from its `w:name` ("heading 1" ...
 * "heading 9"), or -1. Word identifies built-in styles by name, so this holds in every locale;
 * {@link PropertyResolver.getLvlFromHeadingStyle} keys on the English style id.
 */
export function headingLevelByName(style: wml.Style | undefined): number {
  const name = style?.name?.val;
  if (name === undefined) return -1;
  const match = HEADING_NAME.exec(name.trim());
  return match ? Number(match[1]) : -1;
}

export class PropertyResolver {

  private documentDefaultPPr: wml.PPr = { TYPE_NAME: 'org_docx4j_wml.PPr' };
  private documentDefaultRPr: wml.RPr = { TYPE_NAME: 'org_docx4j_wml.RPr' };

  /** All styles in the styles part, keyed by style id (docx4j's `liveStyles`). */
  private readonly liveStyles = new Map<string, wml.Style>();
  /** How many styles the part held when `liveStyles` was last scanned; a miss rescans only when that has changed. */
  private scannedStyleCount = -1;

  /** A style's `w:basedOn` chain merged root-first, WITHOUT the document defaults, per style id. */
  private readonly chainPPrCache = new Map<string, wml.PPr>();
  private readonly chainRPrCache = new Map<string, wml.RPr>();
  /** The document defaults with the chain applied over them: what the style overloads return. */
  private readonly effectivePPrByStyle = new Map<string, wml.PPr>();
  private readonly effectiveRPrByStyle = new Map<string, wml.RPr>();

  /** Missing styles are logged once each per resolver. */
  private readonly missingLogged = new Set<string>();

  private numberingLevels: NumberingIndents | undefined;

  defaultParagraphStyleId: string | undefined;
  private defaultCharacterStyleId: string | undefined;
  private defaultTableStyleId: string | undefined;

  private source: ResolverSource;

  private constructor(private readonly provider: () => ResolverSource,
    /** Set when the styles tree the resolver holds is not the styles part's own. */
    private readonly stylesArePrivate: () => boolean = () => false) {
    this.source = provider();
    this.init();
  }

  /**
   * Builds a resolver over a package's styles and numbering parts (docx4j's constructor,
   * which can do this synchronously because JAXB is).
   *
   * The parts are read with `readContents()`, which does **not** mark them unmarshalled, so a
   * document whose styles nobody has edited still saves byte for byte. Once something else
   * unmarshals the part - the comment styles, a caller's `getContents()` - `refresh()` picks
   * up that live tree instead of the private copy.
   */
  static async create(pkg: PackageLike | ResolverSource): Promise<PropertyResolver> {
    if (typeof (pkg as PackageLike).getMainDocumentPart !== 'function') {
      const source = pkg as ResolverSource;
      return new PropertyResolver(() => source);
    }
    let main;
    try {
      main = (pkg as PackageLike).getMainDocumentPart();
    } catch (e) {
      throw new Docx4JException('A PropertyResolver needs a main document part', { cause: e });
    }
    const stylesPart = main.styleDefinitionsPart;
    const numberingPart = main.numberingDefinitionsPart;
    const privateStyles = stylesPart === undefined ? undefined : await stylesPart.readContents();
    const privateNumbering = numberingPart === undefined ? undefined : await numberingPart.readContents();
    return new PropertyResolver(
      () => ({
        styles: stylesPart?.isUnmarshalled === true ? stylesPart.contents : privateStyles,
        numbering: numberingPart?.isUnmarshalled === true ? numberingPart.contents : privateNumbering,
      }),
      () => stylesPart !== undefined && !stylesPart.isUnmarshalled,
    );
  }

  /** Builds a resolver over object trees directly (a test, or parts already in hand). */
  static fromSource(source: ResolverSource): PropertyResolver {
    return new PropertyResolver(() => source);
  }

  // ------------------------------------------------------------------ init

  private init(): void {
    const styles = this.source.styles;
    this.liveStyles.clear();
    this.scannedStyleCount = -1;
    this.defaultParagraphStyleId = undefined;
    this.defaultCharacterStyleId = undefined;
    this.defaultTableStyleId = undefined;
    if (styles !== undefined) {
      this.scanStyles();
      for (const style of styles.style ?? []) {
        if (style._default !== true || style.styleId === undefined) continue;
        if (style.type === 'paragraph') this.defaultParagraphStyleId ??= style.styleId;
        else if (style.type === 'character') this.defaultCharacterStyleId ??= style.styleId;
        else if (style.type === 'table') this.defaultTableStyleId ??= style.styleId;
      }
    }
    if (this.defaultParagraphStyleId === undefined) log.warn('No default paragraph style!!');

    // private copies: the resolver's defaults are its own, so nothing below writes into the
    // styles part (until CR-015 phase 3 the w:sz 20 default went into the part, and a docx
    // saved after an export carried it)
    const docDefaults = styles?.docDefaults;
    this.documentDefaultPPr = docDefaults?.pPrDefault?.pPr === undefined
      ? { TYPE_NAME: 'org_docx4j_wml.PPr' } : deepCopy(docDefaults.pPrDefault.pPr);
    this.documentDefaultRPr = docDefaults?.rPrDefault?.rPr === undefined
      ? { TYPE_NAME: 'org_docx4j_wml.RPr' } : deepCopy(docDefaults.rPrDefault.rPr);
    if (this.documentDefaultRPr.sz === undefined) {
      // Word's default made explicit: 10pt where nothing states a size (measured, CR-015 probe
      // styles-no-size-anywhere: identical to an explicit 10pt run)
      this.documentDefaultRPr.sz = { TYPE_NAME: 'org_docx4j_wml.HpsMeasure', val: 20 };
    }

    this.numberingLevels = this.source.numbering === undefined
      ? undefined
      : new NumberingLevels(this.source.numbering, (styleId) => this.getLiveStyle(styleId));
  }

  /** (Re)read the styles part into `liveStyles`. A style with no id is skipped. */
  private scanStyles(): void {
    const list = this.source.styles?.style ?? [];
    for (const style of list) {
      if (style.styleId === undefined) continue;
      this.liveStyles.set(style.styleId, style);
    }
    this.scannedStyleCount = list.length;
  }

  /**
   * The style with this id, rescanning the styles part on a miss - and only when the list's
   * size has changed - so that a style *added* to the part after this resolver was built is
   * still found. A style *modified* or *removed* needs {@link refresh}.
   */
  private getLiveStyle(styleId: string | undefined): wml.Style | undefined {
    if (styleId === undefined) return undefined;
    let style = this.liveStyles.get(styleId);
    if (style === undefined && (this.source.styles?.style?.length ?? 0) !== this.scannedStyleCount) {
      this.scanStyles();
      style = this.liveStyles.get(styleId);
    }
    return style;
  }

  private logMissing(styleId: string): void {
    if (this.missingLogged.has(styleId)) return;
    this.missingLogged.add(styleId);
    log.warn(`Style definition not found: ${styleId} (logged once)`);
  }

  /**
   * Discard cached state and re-read the styles part. A style merely *added* is picked up
   * automatically; call this when a style has been modified or removed, or the styles part's
   * contents replaced. The content API calls it where it creates a style.
   */
  refresh(): void {
    this.chainPPrCache.clear();
    this.chainRPrCache.clear();
    this.effectivePPrByStyle.clear();
    this.effectiveRPrByStyle.clear();
    this.missingLogged.clear();
    this.source = this.provider();
    this.init();
  }

  // ------------------------------------------------------------------ defaults

  getDocumentDefaultPPr(): wml.PPr {
    return this.documentDefaultPPr;
  }

  getDocumentDefaultRPr(): wml.RPr {
    return this.documentDefaultRPr;
  }

  /** The style id of the `w:default="1"` paragraph style, or undefined when the part declares none. */
  getDefaultParagraphStyleId(): string | undefined {
    return this.defaultParagraphStyleId;
  }

  /** The default paragraph style's effective `w:pPr`, as `getEffectivePPr(styleId)` gives it. */
  getResolvedDefaultParagraphStyle(): wml.PPr {
    return this.getEffectivePPr(this.defaultParagraphStyleId);
  }

  /** The style with this id, or undefined (docx4j `getStyle`). */
  getStyle(styleId: string | undefined): wml.Style | undefined {
    return this.getLiveStyle(styleId);
  }

  // ------------------------------------------------------------------ the chains

  /**
   * The style and the styles it is based on, root first (the base of the chain at index 0, the
   * style itself last); empty for an unknown id. A cycle, or a chain deeper than 32, ends the
   * walk where it is detected (and throws `CyclicStylesException` when
   * `throwOnCyclicStyles(true)` is set).
   */
  ancestry(styleId: string | undefined): wml.Style[] {
    const leafFirst: wml.Style[] = [];
    const seen: string[] = [];
    let id = styleId;
    while (id !== undefined) {
      if (isCyclic(id, seen)) break;
      seen.push(id);
      const style = this.getLiveStyle(id);
      if (style === undefined) {
        // "DocDefaults" is StyleTree's virtual style for the document defaults, which this
        // resolver applies itself; anything else is a reference to nothing
        if (id !== 'DocDefaults') this.logMissing(id);
        break;
      }
      leafFirst.push(style);
      id = style.basedOn?.val;
    }
    return leafFirst.reverse();
  }

  /**
   * A paragraph style's `w:basedOn` chain merged root-first, WITHOUT the document defaults:
   * what the style contributes on its own. A live, cached object: copy before changing.
   */
  getChainPPr(styleId: string | undefined): wml.PPr {
    if (styleId === undefined) return { TYPE_NAME: 'org_docx4j_wml.PPr' };
    const cached = this.chainPPrCache.get(styleId);
    if (cached !== undefined) return cached;
    const chain: wml.PPr = { TYPE_NAME: 'org_docx4j_wml.PPr' };
    for (const style of this.ancestry(styleId)) {
      this.applyPPrLayer(this.headingLayer(style), chain);
    }
    this.chainPPrCache.set(styleId, chain);
    return chain;
  }

  /**
   * A style's `w:basedOn` chain merged root-first, WITHOUT the document defaults: for a
   * character style, what it contributes over the paragraph's run properties. A live, cached
   * object: copy before changing.
   */
  getChainRPr(styleId: string | undefined): wml.RPr {
    if (styleId === undefined) return { TYPE_NAME: 'org_docx4j_wml.RPr' };
    const cached = this.chainRPrCache.get(styleId);
    if (cached !== undefined) return cached;
    const chain: wml.RPr = { TYPE_NAME: 'org_docx4j_wml.RPr' };
    for (const style of this.ancestry(styleId)) {
      applyRPr<wml.RPr>(style.rPr, chain);
    }
    this.chainRPrCache.set(styleId, chain);
    return chain;
  }

  /**
   * A style's own `w:pPr` as a layer of its chain. The built-in heading styles have a fixed
   * outline level, which Word takes from the style's built-in NAME ("heading 1" ...
   * "heading 9", the same in every locale; the style id is localised: "Überschrift1",
   * "Titre1"), whatever `w:outlineLvl` the style declares: where they differ the layer is a
   * copy with the level from the name, and the styles part is not touched.
   */
  private headingLayer(style: wml.Style): wml.PPr | undefined {
    const layer = style.pPr;
    const headingLevel = headingLevelByName(style);
    if (headingLevel > 0 && layer !== undefined
      && layer.outlineLvl !== undefined && layer.outlineLvl.val !== undefined
      && layer.outlineLvl.val !== headingLevel - 1) {
      const copy = deepCopy(layer);
      copy.outlineLvl!.val = headingLevel - 1;
      return copy;
    }
    return layer;
  }

  private applyPPrLayer(source: wml.PPr | undefined, destination: wml.PPr): void {
    if (source === undefined) return;
    applyPPrBase(source, destination, this.numberingLevels);
    destination.rPr = applyRPr<wml.ParaRPr>(source.rPr, destination.rPr,
      () => ({ TYPE_NAME: 'org_docx4j_wml.ParaRPr' }));
  }

  /** The paragraph's style: its `w:pStyle` where that names a style that exists, else the default paragraph style. */
  private paragraphStyleOf(pPr: wml.PPr | undefined): string | undefined {
    if (pPr === undefined || pPr.pStyle === undefined) return this.defaultParagraphStyleId;
    const styleId = pPr.pStyle.val;
    if (styleId === undefined) return this.defaultParagraphStyleId;
    return this.existingParagraphStyle(styleId);
  }

  /** `styleId` if it names a style that exists, else the default paragraph style's id (logged once). */
  private existingParagraphStyle(styleId: string | undefined): string | undefined {
    if (styleId === undefined) return this.defaultParagraphStyleId;
    if (this.getLiveStyle(styleId) === undefined) {
      this.logMissing(styleId);
      return this.defaultParagraphStyleId;
    }
    return styleId;
  }

  // ------------------------------------------------------------------ paragraphs

  /**
   * The paragraph properties which actually apply, given a paragraph's own `w:pPr`: the
   * document defaults, the paragraph style's chain, then the direct formatting. Run properties
   * are not resolved here; the paragraph mark's `w:rPr` is carried as `pPr/rPr` but is not a
   * run's formatting ({@link getEffectiveParagraphMarkRPr}).
   *
   * What is returned is a live object when the paragraph states no direct formatting: copy it
   * before changing it.
   */
  getEffectivePPr(pPr: wml.PPr | undefined): wml.PPr;
  /**
   * The paragraph properties a style resolves to (document defaults included). A style id
   * naming no style resolves as the default paragraph style does.
   */
  getEffectivePPr(styleId: string | undefined): wml.PPr;
  getEffectivePPr(arg: wml.PPr | string | undefined): wml.PPr {
    if (arg === undefined || typeof arg === 'string') return this.effectivePPrOfStyle(arg);
    const resolved = this.effectivePPrOfStyle(this.paragraphStyleOf(arg));
    if (!hasDirectFormattingPPr(arg)) return resolved;
    const effective = deepCopy(resolved);
    this.applyPPrLayer(arg, effective);
    return effective;
  }

  private effectivePPrOfStyle(styleId: string | undefined): wml.PPr {
    const existing = this.existingParagraphStyle(styleId);
    const key = existing ?? NO_STYLE;
    const cached = this.effectivePPrByStyle.get(key);
    if (cached !== undefined) return cached;
    const resolved = deepCopy(this.documentDefaultPPr);
    this.applyPPrLayer(this.getChainPPr(existing), resolved);
    this.effectivePPrByStyle.set(key, resolved);
    return resolved;
  }

  // ------------------------------------------------------------------ runs

  /**
   * The run properties which apply to a run, given its own `w:rPr` and the `w:pPr` of the
   * paragraph it is in: document defaults, the paragraph style's run properties, the run's
   * character style, then its direct formatting. A new object each call.
   *
   * The paragraph mark's `w:rPr` is never applied to a run; for the mark itself see
   * {@link getEffectiveParagraphMarkRPr}.
   */
  getEffectiveRPr(rPr: wml.RPr | undefined, pPr?: wml.PPr | undefined): wml.RPr;
  /**
   * The run properties a style resolves to on its own: document defaults, then its `w:basedOn`
   * chain. For a paragraph style, the run properties of its paragraphs; for a character style,
   * what it contributes with nothing under it. Undefined when no such style exists.
   */
  getEffectiveRPr(styleId: string): wml.RPr | undefined;
  getEffectiveRPr(arg: wml.RPr | string | undefined, pPr?: wml.PPr | undefined): wml.RPr | undefined {
    if (typeof arg === 'string') return this.effectiveRPrOfStyle(arg);
    const effective = deepCopy(this.documentDefaultRPr);
    applyRPr<wml.RPr>(this.getChainRPr(this.paragraphStyleOf(pPr)), effective);
    this.applyCharacterStyleAndDirect(arg, effective);
    return effective;
  }

  private effectiveRPrOfStyle(styleId: string): wml.RPr | undefined {
    const cached = this.effectiveRPrByStyle.get(styleId);
    if (cached !== undefined) return cached;
    if (this.getLiveStyle(styleId) === undefined) {
      this.logMissing(styleId);
      return undefined;
    }
    const resolved = deepCopy(this.documentDefaultRPr);
    applyRPr<wml.RPr>(this.getChainRPr(styleId), resolved);
    this.effectiveRPrByStyle.set(styleId, resolved);
    return resolved;
  }

  /**
   * The character style's chain, then the direct formatting.
   *
   * The character style is a **level** of the style hierarchy over the paragraph's, so its
   * toggle properties are XORed with what is beneath rather than overriding it (ECMA-376-1
   * §17.7.3). The direct formatting after it is not a level: an explicit value there is used
   * as it stands.
   */
  private applyCharacterStyleAndDirect(rPr: wml.RPr | undefined, effective: wml.RPr): void {
    const runStyleId = rPr?.rStyle?.val;
    if (runStyleId !== undefined) {
      if (this.getLiveStyle(runStyleId) === undefined) {
        this.logMissing(runStyleId);
      } else {
        applyStyleLevel(this.getChainRPr(runStyleId), effective, this.documentDefaultRPr);
      }
    }
    if (hasDirectFormattingRPr(rPr)) applyRPr<wml.RPr>(rPr, effective);
  }

  /**
   * The run properties of the paragraph mark: document defaults, the paragraph style's run
   * properties, then the `w:pPr`'s own `w:rPr`. What sizes an empty paragraph, and what a list
   * label starts from.
   */
  getEffectiveParagraphMarkRPr(pPr: wml.PPr | undefined): wml.RPr {
    const effective = deepCopy(this.documentDefaultRPr);
    applyRPr<wml.RPr>(this.getChainRPr(this.paragraphStyleOf(pPr)), effective);
    if (pPr?.rPr !== undefined) applyRPr<wml.RPr>(pPr.rPr, effective);
    return effective;
  }

  /** Whether the run states any formatting of its own (docx4j's deprecated `hasDirectRPrFormatting`). */
  hasDirectRPrFormatting(rPr: wml.RPr | undefined): boolean {
    return hasDirectFormattingRPr(rPr);
  }

  // ------------------------------------------------------------------ tables

  /**
   * The table style which applies, merged root-first down its `w:basedOn` chain, then the
   * table's own `w:tblPr` over it.
   *
   * Word's built-in "Normal Table" (`w:tblInd` 0; cell margins 108 twips left and right, 0 top
   * and bottom) underlies a table naming no style and a table whose chain reaches the
   * document's default table style - and it is the built-in that applies, not the document's
   * definition of that style (measured, CR-015 probe styles-table-default). So the default
   * style's own layer is skipped in favour of the built-in, and a chain that does not reach it
   * starts from nothing.
   *
   * A new `Style` each call, with a non-null `w:tblPr`.
   */
  getEffectiveTableStyle(tblPr: wml.CTTblPrBase | undefined): wml.Style {
    const styleId = tblPr?.tblStyle?.val;
    const chain = styleId === undefined ? [] : this.ancestry(styleId);
    const builtIn = chain.length === 0
      || (this.defaultTableStyleId !== undefined && chain.some((s) => s.styleId === this.defaultTableStyleId));

    const result = builtIn ? this.builtInTableNormal() : emptyTableStyle();
    for (const layer of chain) {
      // the built-in stands in for the document's definition of the default table style
      if (this.defaultTableStyleId !== undefined && layer.styleId === this.defaultTableStyleId) continue;
      applyStyle(layer, result);
    }
    if (tblPr !== undefined) {
      result.tblPr = applyTblPr(tblPr, result.tblPr);
    }
    result.tblPr ??= { TYPE_NAME: 'org_docx4j_wml.CTTblPrBase' };
    return result;
  }

  /**
   * Whether a table's style chain reaches the document's default table style - the flag
   * {@link getEffectiveTableStyle} decides Word's built-in Normal Table by (a table naming no
   * style counts as reaching it).
   */
  reachesDefaultTableStyle(tblPr: wml.CTTblPrBase | undefined): boolean {
    const styleId = tblPr?.tblStyle?.val;
    const chain = styleId === undefined ? [] : this.ancestry(styleId);
    return chain.length === 0
      || (this.defaultTableStyleId !== undefined && chain.some((s) => s.styleId === this.defaultTableStyleId));
  }

  /** Word's built-in Normal Table, as a style: what it applies whatever the document's own definition says. */
  private builtInTableNormal(): wml.Style {
    const style = emptyTableStyle();
    style.styleId = this.defaultTableStyleId ?? 'TableNormal';
    style.name = { TYPE_NAME: 'org_docx4j_wml.Style.Name', val: 'Normal Table' };
    const tblPr = style.tblPr!;
    tblPr.tblInd = twips(0);
    tblPr.tblCellMar = {
      TYPE_NAME: 'org_docx4j_wml.CTTblCellMar',
      top: twips(0),
      left: twips(WORD_DEFAULT_CELL_MARGIN_TWIPS),
      bottom: twips(0),
      right: twips(WORD_DEFAULT_CELL_MARGIN_TWIPS),
    };
    return style;
  }

  // ------------------------------------------------------------------ headings and known styles

  /**
   * The level of a heading style from its **id** ("Heading2" -> 2), or -1. docx4j keeps this
   * because callers pass a style id; {@link headingLevelByName} is the rule Word applies, and
   * is what resolution uses.
   */
  getLvlFromHeadingStyle(styleId: string): number {
    if (!styleId.startsWith(HEADING_STYLE)) return -1;
    const rest = styleId.slice(HEADING_STYLE.length).trim();
    if (!/^\d+$/.test(rest)) return -1;
    return Number(rest);
  }

  /**
   * Make a style available for use in the document, adding it to the styles part - and,
   * recursively, whatever it is `w:basedOn` and whatever it is linked to. Given a style id,
   * the style must be one this package knows (there is no KnownStyles.xml here yet; a style id
   * that is not already live and not supplied as an object returns false).
   */
  activateStyle(style: wml.Style | string): boolean {
    if (typeof style !== 'string' && this.stylesArePrivate()) {
      throw new Docx4JException('activateStyle writes into the styles part: await styleDefinitionsPart.getContents() and refresh() the resolver first');
    }
    if (typeof style === 'string') {
      if (this.getLiveStyle(style) !== undefined) return true;
      log.warn(`Unknown style: ${style}`);
      return false;
    }
    return this.activate(style, true);
  }

  private activate(style: wml.Style, replace: boolean): boolean {
    if (style.styleId === undefined) return false;
    const styles = this.source.styles;
    if (styles === undefined) return false;
    const list = (styles.style ??= []);
    const existing = this.getLiveStyle(style.styleId);
    if (existing !== undefined) {
      if (!replace) return false;
      const index = list.indexOf(existing);
      if (index >= 0) list.splice(index, 1);
    }
    list.push(style);
    this.liveStyles.set(style.styleId, style);
    this.scannedStyleCount = list.length;
    // resolution of this style (and of anything based on it) has to be recomputed
    this.chainPPrCache.clear();
    this.chainRPrCache.clear();
    this.effectivePPrByStyle.clear();
    this.effectiveRPrByStyle.clear();

    let ok = true;
    const basedOn = style.basedOn?.val;
    if (basedOn !== undefined) ok = this.activateStyle(basedOn);
    const link = style.link?.val;
    if (link !== undefined) ok = this.activateStyle(link) && ok;
    return ok;
  }
}

function twips(w: number): wml.TblWidth {
  return { TYPE_NAME: 'org_docx4j_wml.TblWidth', type: 'dxa', w };
}

function emptyTableStyle(): wml.Style {
  return {
    TYPE_NAME: 'org_docx4j_wml.Style',
    type: 'table',
    tblPr: { TYPE_NAME: 'org_docx4j_wml.CTTblPrBase' },
  };
}
