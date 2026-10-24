// A theme's colour scheme and the colour map that leads to it: docx4j `org.docx4j.model.color.ThemeColors` (CR-034).
//
// Which theme and which map apply depends on the host (docx4j CR-034 section 2.1): a docx's main document part's
// theme and the settings' `w:clrSchemeMapping`; a slide's master's theme with the master's `p:clrMap`, then the
// layout's and the slide's `p:clrMapOvr`; an xlsx's workbook theme; a chart its host's, under its theme override
// part and its `c:clrMapOvr`. A package with no theme part gets the theme Office supplies, the bundled one
// `pkg.fonts.defaultTheme` names (docx4j's `docx4j.fonts.defaultTheme`; Office 2023's by default).
//
// Reading a package's or a part's theme is asynchronous here (`ThemeColors.of`), since its parts are read; the
// parts are read with `readContents()`, so a package whose theme nobody edited still saves byte for byte.
import type * as dml from '@docx4j/generated-objects-ts/modules/org_docx4j_dml';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as chart from '@docx4j/generated-objects-ts/modules/org_docx4j_dml_chart';
import { getContext, getContextSync, type Jsonix } from '@docx4j/generated-objects-ts';
import { Part } from '../../parts/Part.mjs';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import { Namespaces } from '../../parts/Namespaces.mjs';
import { ThemePart, ThemeOverridePart } from '../../parts/dml/index.mjs';
import {
  SlidePart, SlideLayoutPart, SlideMasterPart, NotesSlidePart, NotesMasterPart, HandoutMasterPart,
} from '../../parts/pml/index.mjs';
import { Docx4JException } from '../../opc/exceptions.mjs';
import type { OpcPackage } from '../../packages/OpcPackage.mjs';
import { themeOfSettings, defaultThemeSetting, type DefaultTheme, type FontSettings } from '../fonts/defaultTheme.mjs';
import { ResolvedColor } from './ResolvedColor.mjs';
import { DrawingMLColors } from './DrawingMLColors.mjs';

/** The scheme's slots. */
const SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'] as const;
/** What a colour map takes to a slot. */
const MAPPED = ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'] as const;

/** WordprocessingML's names for the scheme's slots. */
const WORD_NAMES: ReadonlyMap<string, string> = new Map([
  ['dark1', 'dk1'], ['light1', 'lt1'], ['dark2', 'dk2'], ['light2', 'lt2'],
  ['accent1', 'accent1'], ['accent2', 'accent2'], ['accent3', 'accent3'], ['accent4', 'accent4'], ['accent5', 'accent5'], ['accent6', 'accent6'],
  ['hyperlink', 'hlink'], ['followedHyperlink', 'folHlink'],
]);

/** Excel's theme indexes: the first two pairs swapped (docx4j CR-034 section 11, P4). */
const SHEET_ORDER = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'] as const;

/** The default colour map: bg1 to lt1, tx1 to dk1, bg2 to lt2, tx2 to dk2, the rest to themselves. */
function defaultMap(): ReadonlyMap<string, string> {
  const m = new Map<string, string>([['bg1', 'lt1'], ['tx1', 'dk1'], ['bg2', 'lt2'], ['tx2', 'dk2']]);
  for (let i = 4; i < MAPPED.length; i++) m.set(MAPPED[i]!, MAPPED[i]!);
  return m;
}

function scheme(cs: dml.CTColorScheme | undefined): ReadonlyMap<string, ResolvedColor> {
  const m = new Map<string, ResolvedColor>();
  if (cs !== undefined && cs !== null) {
    for (const slot of SLOTS) {
      const c = cs[slot];
      const r = c === undefined || c === null ? undefined : DrawingMLColors.resolve(c, undefined);
      if (r !== undefined) m.set(slot, r);
    }
  }
  return m;
}

// ------------------------------------------------------------------ the bundled themes

const bundled = new Map<string, dml.Theme>();

function parseTheme(context: Jsonix.Context, xml: string): dml.Theme {
  const element = context.createUnmarshaller().unmarshalString(xml) as Jsonix.TypedNamedValue<dml.Theme>;
  return element.value;
}

/** A bundled theme, parsed; the objects package's context must be built (it is once anything has been read). */
function bundledThemeSync(theme: DefaultTheme): dml.Theme {
  let t = bundled.get(theme.value);
  if (t === undefined) {
    let context: Jsonix.Context;
    try {
      context = getContextSync();
    } catch (e) {
      throw new Docx4JException(`The default theme ${theme.value} needs the objects context: await ThemeColors.loadDefaultTheme() first`, { cause: e });
    }
    t = parseTheme(context, theme.xml);
    bundled.set(theme.value, t);
  }
  return t;
}

async function readQuietly<T>(part: XmlPart<T> | undefined): Promise<T | undefined> {
  if (part === undefined) return undefined;
  try {
    return await part.readContents();
  } catch {
    return undefined;
  }
}

/** The part a relationship of this type leads to, if it is one of this class. */
function related<P extends Part>(part: Part, type: string, kind: abstract new (...args: never[]) => P): P | undefined {
  const rp = part.relationshipsPart;
  const rel = rp?.getRelationshipByType(type);
  const target = rel === undefined ? undefined : rp!.getPart(rel);
  return target instanceof kind ? target : undefined;
}

interface WordPackageLike {
  getMainDocumentPart(): { themePart?: ThemePart; documentSettingsPart?: XmlPart<wml.CTSettings> };
  fonts?: FontSettings;
}
interface SpreadsheetPackageLike {
  getWorkbookPart(): { themePart?: ThemePart };
  fonts?: FontSettings;
}
const isWord = (pkg: unknown): pkg is WordPackageLike => typeof (pkg as WordPackageLike | undefined)?.getMainDocumentPart === 'function';
const isSpreadsheet = (pkg: unknown): pkg is SpreadsheetPackageLike => typeof (pkg as SpreadsheetPackageLike | undefined)?.getWorkbookPart === 'function';

/** The default theme a package names (`pkg.fonts.defaultTheme`), else the process-wide setting. */
function defaultThemeOfPackage(pkg: unknown): DefaultTheme {
  return themeOfSettings((pkg as { fonts?: FontSettings } | undefined)?.fonts);
}

/**
 * A theme's colour scheme and the colour map that leads to it: all a colour reference needs. Immutable; each
 * `with...` returns a new one. Which theme and which map apply depends on the host: {@link ThemeColors.of} takes a
 * package (docx or xlsx) or a part (a slide, layout, master or notes part; a part of a docx or xlsx), and
 * {@link ThemeColors.forChart} a chart in any of them - drawn in its host's theme.
 *
 * The scheme's slots are `dk1`, `lt1`, `dk2`, `lt2`, `accent1` to `accent6`, `hlink` and `folHlink`; the map takes
 * `bg1`, `tx1`, `bg2`, `tx2` (and, in principle, the accents and hyperlinks) to a slot. A slot holding a system
 * colour is the machine's colour, not its `lastClr` (docx4j CR-034 section 11).
 *
 * docx4j's names, but: `of(Theme)` is {@link ThemeColors.ofTheme} here (synchronous), and `of(WordprocessingMLPackage)`,
 * `of(SpreadsheetMLPackage)` and `of(Part)` are one asynchronous {@link ThemeColors.of}; `forChart(Chart)` is
 * asynchronous; `getFormatScheme()` and `getPlaceholder()` are the properties `formatScheme` and `placeholder`.
 */
export class ThemeColors {
  private constructor(
    private readonly slots: ReadonlyMap<string, ResolvedColor>,
    private readonly map: ReadonlyMap<string, string>,
    /** The theme's format scheme, which style references index into; undefined where the theme has none. */
    readonly formatScheme: dml.CTStyleMatrix | undefined,
    /** `phClr`'s colour, where a style reference supplies one. */
    readonly placeholder: ResolvedColor | undefined,
  ) {}

  // ------------------------------------------------------------------ making one

  /**
   * A theme with the default colour map (docx4j `of(Theme)`).
   *
   * @param theme undefined for the theme Office supplies where a package has none: the bundled theme `defaultTheme`
   *   names, else {@link defaultThemeSetting}'s. Parsing it needs the objects package's context, which is built once
   *   anything has been read; {@link ThemeColors.loadDefaultTheme} builds it.
   */
  static ofTheme(theme: dml.Theme | undefined, defaultTheme?: DefaultTheme): ThemeColors {
    theme ??= bundledThemeSync(defaultTheme ?? defaultThemeSetting());
    const elements = theme.themeElements;
    return new ThemeColors(scheme(elements?.clrScheme), defaultMap(), elements?.fmtScheme ?? undefined, undefined);
  }

  /** Parses a bundled theme ahead of a synchronous {@link ThemeColors.ofTheme}(undefined), building the context if need be. */
  static async loadDefaultTheme(defaultTheme: DefaultTheme = defaultThemeSetting()): Promise<dml.Theme> {
    let t = bundled.get(defaultTheme.value);
    if (t === undefined) {
      t = parseTheme(await getContext(), defaultTheme.xml);
      bundled.set(defaultTheme.value, t);
    }
    return t;
  }

  /**
   * The theme colours of a package or a part:
   * - a docx: the main document part's theme (or the default theme) and `w:clrSchemeMapping` from the settings,
   *   which applies to WordprocessingML theme colours and DrawingML scheme colours alike (docx4j CR-034 section 11, P2);
   * - an xlsx: the workbook's theme (or the default theme), with the default map for its drawings;
   * - a slide: its master's theme, its master's `p:clrMap`, then its layout's `p:clrMapOvr`, then its own; a layout,
   *   a master, a notes slide, a notes or handout master likewise;
   * - a part of a docx or xlsx: its package's.
   * Anything else, or a part that can't be read, gets the default theme with the default map.
   */
  static async of(source: OpcPackage | Part): Promise<ThemeColors> {
    if (source instanceof Part) return ThemeColors.ofPart(source);
    const fallback = defaultThemeOfPackage(source);
    if (isWord(source)) {
      const main = source.getMainDocumentPart();
      const tc = await ThemeColors.themed(await readQuietly(main.themePart), fallback);
      const settings = await readQuietly(main.documentSettingsPart);
      const m = settings?.clrSchemeMapping;
      return m === undefined || m === null ? tc : tc.withWordMapping(m);
    }
    if (isSpreadsheet(source)) {
      return ThemeColors.themed(await readQuietly(source.getWorkbookPart().themePart), fallback);
    }
    return ThemeColors.themed(undefined, fallback);
  }

  private static async themed(theme: dml.Theme | undefined, fallback: DefaultTheme): Promise<ThemeColors> {
    if (theme === undefined) await ThemeColors.loadDefaultTheme(fallback);
    return ThemeColors.ofTheme(theme, fallback);
  }

  private static async ofPart(part: Part): Promise<ThemeColors> {
    const fallback = defaultThemeOfPackage(part.package);
    try {
      if (part instanceof SlidePart) {
        const layout = part.slideLayoutPart;
        const tc = layout === undefined ? await ThemeColors.themed(undefined, fallback) : await ThemeColors.ofPart(layout);
        return tc.withOverride((await part.readContents()).clrMapOvr);
      }
      if (part instanceof SlideLayoutPart) {
        const master = part.slideMasterPart;
        const tc = master === undefined ? await ThemeColors.themed(undefined, fallback) : await ThemeColors.ofPart(master);
        return tc.withOverride((await part.readContents()).clrMapOvr);
      }
      if (part instanceof SlideMasterPart || part instanceof NotesMasterPart || part instanceof HandoutMasterPart) {
        const theme = await readQuietly(related(part, Namespaces.THEME, ThemePart));
        return (await ThemeColors.themed(theme, fallback)).withMapping((await part.readContents()).clrMap);
      }
      if (part instanceof NotesSlidePart) {
        const master = related(part, Namespaces.PRESENTATIONML_NOTES_MASTER, NotesMasterPart);
        const tc = master === undefined ? await ThemeColors.themed(undefined, fallback) : await ThemeColors.ofPart(master);
        return tc.withOverride((await part.readContents()).clrMapOvr);
      }
    } catch {
      // docx4j logs "Couldn't read the colour map of ..." and takes the default theme
      return ThemeColors.themed(undefined, fallback);
    }
    const pkg = part.package;
    if (isWord(pkg) || isSpreadsheet(pkg)) return ThemeColors.of(pkg);
    return ThemeColors.themed(undefined, fallback);
  }

  // ------------------------------------------------------------------ deriving one

  /** This theme with a DrawingML colour map (`p:clrMap`, `c:clrMapOvr`); undefined leaves it. */
  withMapping(m: dml.CTColorMapping | undefined | null): ThemeColors {
    if (m === undefined || m === null) return this;
    const mm = new Map(this.map);
    for (const key of MAPPED) put(mm, key, m[key]);
    return new ThemeColors(this.slots, mm, this.formatScheme, this.placeholder);
  }

  /** `a:overrideClrMapping` replaces the map; `a:masterClrMapping` (or nothing) keeps it. */
  withOverride(o: dml.CTColorMappingOverride | undefined | null): ThemeColors {
    return o === undefined || o === null || o.overrideClrMapping === undefined || o.overrideClrMapping === null
      ? this : this.withMapping(o.overrideClrMapping);
  }

  /** This theme with WordprocessingML's `w:clrSchemeMapping` (whose names differ: dark1, t1, ...). */
  withWordMapping(m: wml.CTColorSchemeMapping | undefined | null): ThemeColors {
    if (m === undefined || m === null) return this;
    const mm = new Map(this.map);
    put(mm, 'bg1', wordSlot(m.bg1));
    put(mm, 'tx1', wordSlot(m.t1));
    put(mm, 'bg2', wordSlot(m.bg2));
    put(mm, 'tx2', wordSlot(m.t2));
    put(mm, 'accent1', wordSlot(m.accent1));
    put(mm, 'accent2', wordSlot(m.accent2));
    put(mm, 'accent3', wordSlot(m.accent3));
    put(mm, 'accent4', wordSlot(m.accent4));
    put(mm, 'accent5', wordSlot(m.accent5));
    put(mm, 'accent6', wordSlot(m.accent6));
    put(mm, 'hlink', wordSlot(m.hyperlink));
    put(mm, 'folHlink', wordSlot(m.followedHyperlink));
    return new ThemeColors(this.slots, mm, this.formatScheme, this.placeholder);
  }

  /** This theme with another colour scheme and format scheme (a theme override part's). */
  withScheme(clrScheme: dml.CTColorScheme | undefined | null, fmt?: dml.CTStyleMatrix | null): ThemeColors {
    return new ThemeColors(clrScheme === undefined || clrScheme === null ? this.slots : scheme(clrScheme), this.map,
      fmt === undefined || fmt === null ? this.formatScheme : fmt, this.placeholder);
  }

  /**
   * A chart's colours: this (its host's) theme, under the chart's theme override part if it has one, and the
   * chart's `c:clrMapOvr` if it states one.
   */
  async forChart(chartPart: XmlPart<chart.CTChartSpace>): Promise<ThemeColors> {
    let tc: ThemeColors = this;
    try {
      const override = await readQuietly(related(chartPart, Namespaces.THEME_OVERRIDE, ThemeOverridePart));
      if (override !== undefined) tc = tc.withScheme(override.clrScheme, override.fmtScheme);
      tc = tc.withMapping((await chartPart.readContents()).clrMapOvr);
    } catch {
      // docx4j logs "Couldn't read ...'s theme override or colour map" and keeps what it has
    }
    return tc;
  }

  /** This theme with `phClr` standing for the given colour (a style reference's). */
  withPlaceholder(color: ResolvedColor | undefined): ThemeColors {
    return new ThemeColors(this.slots, this.map, this.formatScheme, color);
  }

  // ------------------------------------------------------------------ reading it

  /** A scheme slot's colour (dk1, lt1, ..., folHlink), not through the map. */
  slot(slot: string): ResolvedColor | undefined {
    return this.slots.get(slot);
  }

  /**
   * A DrawingML `schemeClr` value's colour: bg1, tx1, bg2, tx2 (and the accents and hyperlinks) through the map,
   * dk1 ... lt2 directly, `phClr` the placeholder.
   */
  schemeColor(val: string | undefined | null): ResolvedColor | undefined {
    if (val === undefined || val === null) return undefined;
    if (val === 'phClr') return this.placeholder;
    const slot = this.map.get(val);
    return this.slots.get(slot ?? val);
  }

  /**
   * A WordprocessingML `w:themeColor` value's colour: text1, background1, text2, background2 through the map,
   * dark1 ... followedHyperlink directly; undefined for "none".
   */
  wordThemeColor(themeColor: string | undefined | null): ResolvedColor | undefined {
    if (themeColor === undefined || themeColor === null) return undefined;
    switch (themeColor) {
      case 'text1': return this.schemeColor('tx1');
      case 'background1': return this.schemeColor('bg1');
      case 'text2': return this.schemeColor('tx2');
      case 'background2': return this.schemeColor('bg2');
      default: {
        const slot = WORD_NAMES.get(themeColor);
        return slot === undefined ? undefined : this.slots.get(slot);
      }
    }
  }

  /**
   * A SpreadsheetML `theme` index's colour. Excel swaps the first two pairs: 0 is lt1, 1 dk1, 2 lt2, 3 dk2, then
   * accent1 to accent6, hlink, folHlink (docx4j CR-034 section 11, P4).
   */
  sheetThemeColor(index: number): ResolvedColor | undefined {
    return index < 0 || index >= SHEET_ORDER.length ? undefined : this.slots.get(SHEET_ORDER[Math.trunc(index)]!);
  }
}

function wordSlot(i: wml.STColorSchemeIndex | undefined | null): string | undefined {
  return i === undefined || i === null ? undefined : WORD_NAMES.get(i);
}

function put(m: Map<string, string>, key: string, slot: string | undefined | null): void {
  if (slot !== undefined && slot !== null) m.set(key, slot);
}
