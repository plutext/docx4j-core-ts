// docx4j `org.docx4j.fonts.Mapper`: a document font name to a physical font.
//
// One order of precedence for every mapper (docx4j CR-016 phase 3), `populateFontMappings`
// first and then the passes `WordprocessingMLPackage.setFontMapper` runs after it:
//
//   1. the font itself where the environment has it, by its name             INSTALLED
//   2. the document's embedded form of it (regular, bold, italic, bold italic) EMBEDDED
//   3. the mapper's own answer - a variant of the name for IdentityPlusMapper  MAPPER_OWN
//   4. addMetricallyCompatibleSubstitutes: font-substitutes.xml, best first    METRIC_CLONE
//                                                                             MEASURED_STAND_IN
//                                                                             CLASS
//   5. addAltNameSubstitutes: the document's own w:altName, chain followed     ALT_NAME
//   6. addClassBasedSubstitutes: a face of the same class                      CLASS
//   7. addWordDefaultSubstitutes: what Word itself draws a font it cannot find WORD_DEFAULT
//   8. addNoBoldFaceAliases: re-maps a family which has no bold face of its own
//
// and what is left is UNMAPPED.  `BestMatchingMapper`'s own step (a panose match, then
// `FontSubstitutions.xml`) would go at 3 and `addMapperSubstitutes` between 6 and 7; it needs
// a font file's panose and is a later CR, so `addMapperSubstitutes` is the empty hook docx4j's
// is and `IdentityPlusMapper` is the only mapper here.
//
// **What is not ported** (CR-001 section 14.2): everything about *files* - the embedded-font
// metrics, the glyph checks, the FOP configuration, `WordLineMetrics`, `WidthFactors` and the
// per-script coverage choices a conversion records against a decision.  `registerLineMetricsAlias`
// is kept because the alt-name and Word-default passes are written over it and a consumer that
// does measure lines needs to know whose metrics Word used, but nothing here reads the metrics.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { PhysicalFont } from './PhysicalFont.mjs';
import { DEFAULT_FONT_REGISTRY, type FontRegistry } from './registry.mjs';
import { SUBSTITUTE_ROWS, type Substitute } from './substitutions.generated.mjs';
import { MS_FONT_NAMES, MS_FONTS_WITH_BOLD, LINE_METRIC_FAMILIES, EAST_ASIAN_FAMILIES } from './families.generated.mjs';
import { leftToTheDocumentDefault, selectByClass, substitutionClass } from './FontFallback.mjs';

/** Which pass answered for a document font (docx4j `FontDecision.Source`). */
export type FontDecisionSource =
  /** The environment has a font of that name. */
  | 'INSTALLED'
  /** The document embeds it (`w:embedRegular` and the rest). */
  | 'EMBEDDED'
  /** A metric-compatible clone from `font-substitutes.xml`: Word's advances. */
  | 'METRIC_CLONE'
  /** No clone exists; the closest face measured against Word's own PDF. */
  | 'MEASURED_STAND_IN'
  /** The document's own `w:altName`, resolved. */
  | 'ALT_NAME'
  /** A face of the same class: the widths are not the document font's. */
  | 'CLASS'
  /** What Word itself draws a font it cannot find. */
  | 'WORD_DEFAULT'
  /** The mapper's own answer: a variant of the name. */
  | 'MAPPER_OWN'
  /** A symbol font, drawn in the face that has the glyphs. */
  | 'SYMBOL'
  /** Nothing mapped it. */
  | 'UNMAPPED';

/** What was decided for one document font, and why. */
export interface FontDecision {
  /** The name as the document has it. */
  readonly documentFont: string;
  readonly source: FontDecisionSource;
  /** How it got there - the `w:altName` chain, the class, Word's default family - or null
   *  where the source says it all. */
  readonly via: string | null;
  /** What is known of the substitute's width error, or null. */
  readonly widthError: string | null;
  /** The physical font as the mapping now stands; null where nothing mapped it. */
  readonly physicalFont: PhysicalFont | null;
}

/** What the report says where nothing was measured. */
export const UNKNOWN_ERROR = 'unknown';

/** Word's own fallback where a document font resolves to nothing at all. */
export const FONT_FALLBACK = 'Times New Roman';

const MS_FONT_NAME_SET = new Set(MS_FONT_NAMES);
const MS_FONTS_WITH_BOLD_SET = new Set(MS_FONTS_WITH_BOLD);
const LINE_METRIC_FAMILY_SET = new Set(LINE_METRIC_FAMILIES);

/** docx4j `WordLineMetrics.ALIASES`: families the table reaches through a built-in alias. */
const LINE_METRIC_ALIASES = new Map<string, string>([
  ['helvetica', 'arial'], ['helvetica neue', 'arial'], ['helveticaneue', 'arial'],
  ['helv', 'arial'], ['arialmt', 'arial'], ['timesnewromanpsmt', 'times new roman'],
]);

const EAST_ASIAN_FAMILY_SET = new Set(EAST_ASIAN_FAMILIES);

/** docx4j `WordLineMetrics.isEastAsianFamily`: the table's optional field 7. */
export function isEastAsianFamily(documentFont: string | undefined): boolean {
  if (documentFont === undefined) return false;
  const key = documentFont.trim().toLowerCase();
  const alias = LINE_METRIC_ALIASES.get(key);
  return EAST_ASIAN_FAMILY_SET.has(key) || (alias !== undefined && EAST_ASIAN_FAMILY_SET.has(alias));
}

/** docx4j `WordLineMetrics.isTableFamily`: the table's own families and its built-in aliases. */
function isLineMetricFamily(documentFontName: string | undefined): boolean {
  if (documentFontName === undefined) return false;
  const key = documentFontName.trim().toLowerCase();
  if (LINE_METRIC_FAMILY_SET.has(key)) return true;
  const alias = LINE_METRIC_ALIASES.get(key);
  return alias !== undefined && LINE_METRIC_FAMILY_SET.has(alias);
}

/**
 * Whether docx4j's tables know this family: `MicrosoftFonts.xml`, `word-line-metrics` (512
 * Microsoft and Office cloud families) or `FontSubstitutions.xml`.
 *
 * A known family the environment merely lacks is substituted for its widths (the passes before
 * `addWordDefaultSubstitutes`), since Word on the author's machine had the font; an unknown one
 * is what Word itself could not find either.  "Known" is `substitutionClass`, **not** `classOf`,
 * which will also guess a class from a name that merely ends in "Sans": docx4j measured a
 * corporate face that fell between the two and was drawn in the document default's serif
 * throughout.  A family deliberately left to the document default is "known" for this purpose
 * too, the measurement that put it there being that the default beat a stand-in.
 */
export function isKnownFamily(documentFontName: string | undefined): boolean {
  if (documentFontName === undefined || documentFontName === null) return false;
  const name = documentFontName.trim();
  if (MS_FONT_NAME_SET.has(name)) return true;
  if (isLineMetricFamily(name)) return true;
  if (substitutionClass(name) !== 'UNKNOWN') return true;
  return leftToTheDocumentDefault(name);
}

const WEIGHT_WORDS = ['light', 'semilight', 'semibold', 'demibold', 'medium', 'black', 'thin',
  'extralight', 'ultralight', 'heavy'];

/**
 * Whether this document font has a bold face of its own: the `MicrosoftFonts.xml` entry where
 * there is one (Franklin Gothic Book has none), else a family whose name ends in a weight word
 * (Calibri Light, Segoe UI Semibold, Arial Black) is a single weight in Windows' font model and
 * has none; any other family is taken to have one.
 */
export function hasBoldFace(documentFontName: string | undefined): boolean {
  if (documentFontName === undefined || documentFontName === null) return true;
  const name = documentFontName.trim();
  if (MS_FONT_NAME_SET.has(name)) return MS_FONTS_WITH_BOLD_SET.has(name);
  let last = name.toLowerCase();
  const sp = last.lastIndexOf(' ');
  if (sp > 0) last = last.slice(sp + 1);
  return !WEIGHT_WORDS.includes(last);
}

/** The source a table row's quality word says. */
function sourceOf(substitute: Substitute | undefined): FontDecisionSource {
  if (substitute === undefined) return 'CLASS';
  if (substitute.quality === 'metric') return 'METRIC_CLONE';
  if (substitute.quality === 'measured') return 'MEASURED_STAND_IN';
  return 'CLASS';
}

/** The font table keyed by lower-cased name. */
export function fontTableOf(fonts: wml.Fonts | undefined): Map<string, wml.Fonts.Font> {
  const table = new Map<string, wml.Fonts.Font>();
  for (const font of fonts?.font ?? []) {
    if (font?.name !== undefined) table.set(font.name.trim().toLowerCase(), font);
  }
  return table;
}

/**
 * The font Word draws an unknown font in, by its fontTable entry: measured (docx4j CR-016
 * probe `fonts-unresolvable`, Word 365) a face whose entry says `w:family="roman"` is drawn in
 * **Cambria**, one saying `swiss` in **Calibri**, one with an entry that names no family in
 * Calibri, one with no fontTable entry at all in Cambria; an exact panose does not change that
 * (Arial's gave Calibri), and the document's default font is never used.
 */
export function wordDefaultFor(documentFontName: string, table: Map<string, wml.Fonts.Font>): string {
  let entry: wml.Fonts.Font | undefined = table.get(documentFontName.trim().toLowerCase());
  if (entry === undefined) return 'Cambria';
  // the family: the entry's own, else the first along its altName chain
  const seen = new Set<string>();
  while (entry !== undefined && entry.name !== undefined && !seen.has(entry.name.trim().toLowerCase())) {
    seen.add(entry.name.trim().toLowerCase());
    const family = entry.family?.val;
    if (family !== undefined) {
      const f = family.trim().toLowerCase();
      if (f === 'roman') return 'Cambria';
      if (f === 'modern') return 'Courier New';
      return 'Calibri'; // swiss, script, decorative, auto
    }
    const alt: string | undefined = entry.altName?.val;
    entry = alt === undefined ? undefined : table.get(alt.trim().toLowerCase());
  }
  return 'Calibri';
}

interface DecisionRecord {
  documentFont: string;
  source: FontDecisionSource;
  via: string | null;
  widthError: string | null;
}

/**
 * Maps the document's font names to physical fonts, over a {@link FontRegistry} the caller
 * supplies.  One per package, as in docx4j.
 */
export abstract class Mapper {

  constructor(
    /** The fonts the consumer has; the docx4j font jars' faces where nothing is given, which
     *  is the environment the parity goldens were made in. */
    readonly registry: FontRegistry = DEFAULT_FONT_REGISTRY,
  ) {}

  /** The mapping, keyed by the lower-cased document font name. */
  private readonly fontMappings = new Map<string, PhysicalFont>();
  /** The document's embedded faces, keyed by the name the font table part uses. */
  private readonly regularForms = new Map<string, PhysicalFont>();
  private readonly boldForms = new Map<string, PhysicalFont>();
  private readonly italicForms = new Map<string, PhysicalFont>();
  private readonly boldItalicForms = new Map<string, PhysicalFont>();
  private readonly decisions = new Map<string, DecisionRecord>();
  /** The document fonts `addWordDefaultSubstitutes` mapped: Word substitutes such a font whole,
   *  bold face included, so `addNoBoldFaceAliases` leaves them alone. */
  private readonly wordDefaulted = new Set<string>();
  private readonly lineMetricsAliases = new Map<string, string>();

  /** The physical font a document font maps to, by case-insensitive name; undefined for a slot
   *  nothing names (docx4j `Mapper.get`, which tolerates null since 17.1.1). */
  get(key: string | undefined | null): PhysicalFont | undefined {
    if (key === undefined || key === null) return undefined;
    return this.fontMappings.get(key.toLowerCase());
  }

  put(key: string, pf: PhysicalFont): void {
    this.fontMappings.set(key.toLowerCase(), pf);
  }

  get size(): number {
    return this.fontMappings.size;
  }

  /** What was decided for this document font, or undefined where no pass met it. */
  getDecision(documentFont: string | undefined): FontDecision | undefined {
    if (documentFont === undefined || documentFont === null) return undefined;
    const record = this.decisions.get(documentFont.trim().toLowerCase());
    if (record === undefined) return undefined;
    return { ...record, physicalFont: this.get(record.documentFont) ?? null };
  }

  /** Every decision, ordered by the document font's name (case-insensitively), so that two
   *  runs over the same document give the same list. */
  getDecisions(): FontDecision[] {
    return [...this.decisions.values()]
      .map((r) => ({ ...r, physicalFont: this.get(r.documentFont) ?? null }))
      .sort((a, b) => a.documentFont.toLowerCase().localeCompare(b.documentFont.toLowerCase()));
  }

  /** Whose Word line metrics a document font takes, where a pass resolved it through another
   *  family; undefined where it takes its own.  Per mapper, i.e. per package: docx4j's was a
   *  JVM-wide map until 17.1.1, and one document's `w:altName` answered for every later one. */
  lineMetricsFamily(documentFont: string | undefined): string | undefined {
    if (documentFont === undefined) return undefined;
    return this.lineMetricsAliases.get(documentFont.trim().toLowerCase());
  }

  registerLineMetricsAlias(documentFont: string | undefined, family: string | undefined): void {
    if (documentFont === undefined || family === undefined) return;
    const key = documentFont.trim().toLowerCase();
    // a font with metrics of its own takes them; there is nothing to alias
    if (isLineMetricFamily(documentFont)) return;
    this.lineMetricsAliases.set(key, family.trim());
  }

  protected decide(documentFont: string, source: FontDecisionSource, via: string | null = null,
    widthError: string | null = null): void {
    this.decisions.set(documentFont.trim().toLowerCase(),
      { documentFont: documentFont.trim(), source, via, widthError });
  }

  // --- embedded faces ----------------------------------------------------------------

  registerRegularForm(name: string, pf: PhysicalFont | undefined): void { set(this.regularForms, name, pf); }
  registerBoldForm(name: string, pf: PhysicalFont | undefined): void { set(this.boldForms, name, pf); }
  registerItalicForm(name: string, pf: PhysicalFont | undefined): void { set(this.italicForms, name, pf); }
  registerBoldItalicForm(name: string, pf: PhysicalFont | undefined): void { set(this.boldItalicForms, name, pf); }

  /** Whether the document embeds this font. */
  isEmbedded(name: string): boolean {
    return this.regularForms.has(name) || this.boldForms.has(name)
      || this.italicForms.has(name) || this.boldItalicForms.has(name);
  }

  /** The installed font of this name, else the document's embedded form of it (docx4j
   *  CR-016 Decisions 2: the installed font wins, for all four faces and in both mappers). */
  protected installedOrEmbedded(name: string): PhysicalFont | undefined {
    return this.registry.get(name) ?? this.regularForms.get(name) ?? this.boldForms.get(name)
      ?? this.italicForms.get(name) ?? this.boldItalicForms.get(name);
  }

  // --- the passes --------------------------------------------------------------------

  /**
   * The mapper's own answer for a document font the environment has neither under that name
   * nor embedded; undefined where it has none.  {@link resolvedVia} says how it was reached.
   */
  protected resolveDocumentFont(_documentFontName: string, _fontTableEntry: wml.Fonts.Font | undefined): PhysicalFont | undefined {
    return undefined;
  }

  /** How the mapper's own answer was reached, set by {@link resolveDocumentFont} as it answers. */
  protected resolvedVia: string | undefined;

  /** An entry for each of the document's fonts: steps 1 to 3 of the precedence. */
  populateFontMappings(documentFontNames: Iterable<string>, fonts: wml.Fonts | undefined): void {
    const table = fontTableOf(fonts);
    for (const documentFontName of documentFontNames) {
      if (documentFontName === undefined || documentFontName.trim().length === 0) continue;
      if (this.get(documentFontName) !== undefined) continue; // already mapped; its decision stands
      let pf = this.installedOrEmbedded(documentFontName);
      let source: FontDecisionSource | undefined = pf === undefined ? undefined
        : (this.registry.get(documentFontName) !== undefined ? 'INSTALLED' : 'EMBEDDED');
      if (pf === undefined) {
        this.resolvedVia = undefined;
        pf = this.resolveDocumentFont(documentFontName, table.get(documentFontName.trim().toLowerCase()));
        if (pf !== undefined) source = 'MAPPER_OWN';
      }
      if (pf === undefined) {
        // recorded, not merely logged: a later pass overwrites this where it maps the font,
        // and what is left says which fonts reached the consumer with nothing
        this.decide(documentFontName, 'UNMAPPED');
      } else {
        this.put(documentFontName, pf);
        this.decide(documentFontName, source as FontDecisionSource,
          source === 'MAPPER_OWN'
            ? (this.resolvedVia ?? `the mapper's own answer: ${pf.name}`)
            : null);
      }
    }
  }

  /**
   * `font-substitutes.xml`: for each document font the table knows, the first of its open
   * substitutes this environment has.  The measurements that chose each substitute are in
   * docx4j's `Mapper.addMetricallyCompatibleSubstitutes` javadoc, and the table's `error`
   * attributes quote them.
   */
  addMetricallyCompatibleSubstitutes(): void {
    for (const row of SUBSTITUTE_ROWS) {
      const before = this.get(row.documentFont);
      let taken: Substitute | undefined;
      for (const candidate of row.substitutes) {
        if (this.registry.get(candidate.font) !== undefined) { taken = candidate; break; }
      }
      if (taken === undefined) continue;
      this.addMetricallyCompatibleSubstitute(row.documentFont, taken.font);
      const after = this.get(row.documentFont);
      if (after !== undefined && after !== before) {
        this.decide(row.documentFont, sourceOf(taken), null,
          taken.error ?? (taken.quality === 'class' ? UNKNOWN_ERROR : null));
      }
    }
  }

  /** Map this document font to an open substitute, unless the environment has the font itself
   *  or the document embeds it (the embedded font is what the author intended, and it is the
   *  only thing certain to be available). */
  protected addMetricallyCompatibleSubstitute(documentFont: string, substitute: string): void {
    if (this.isEmbedded(documentFont)) return;
    if (this.registry.get(documentFont) !== undefined) return;
    const pf = this.registry.get(substitute);
    if (pf !== undefined) this.put(documentFont, pf);
  }

  /**
   * `w:altName` in `word/fontTable.xml` (ECMA-376 17.8.3.1): "the name of an alternate font
   * which shall be used if the font specified is not available".  That is the document author's
   * own answer to a missing font, and Word takes it.
   *
   * After the metric clones - a clone of the document's own font is a better answer than any
   * alias - and before the class pass, which is only a guess from the name.  The chain is
   * followed: the alternate may itself be absent and name an alternate (CR-016 probe
   * `fonts-unresolvable` (e), (f)); a few hops, never a cycle.
   */
  addAltNameSubstitutes(documentFontNames: Iterable<string> | undefined, fonts: wml.Fonts | undefined): void {
    const altNames = new Map<string, string>();
    for (const font of fonts?.font ?? []) {
      const alt = font?.altName?.val;
      if (font?.name === undefined || alt === undefined || alt.trim().length === 0) continue;
      altNames.set(font.name.trim().toLowerCase(), alt.trim());
    }
    if (altNames.size === 0) return;

    const names = documentFontNames ?? altNames.keys();
    for (const documentFontName of names) {
      if (documentFontName === undefined || documentFontName.trim().length === 0) continue;
      if (this.get(documentFontName) !== undefined) continue;   // already mapped
      if (this.isEmbedded(documentFontName)) continue;
      if (this.registry.get(documentFontName) !== undefined) continue; // installed; identity

      const first: string | undefined = altNames.get(documentFontName.trim().toLowerCase());
      let alt: string | undefined = first;
      const seen = new Set<string>([documentFontName.trim().toLowerCase()]);
      let pf: PhysicalFont | undefined;
      let resolvedAlt: string | undefined;
      let eastAsianHop: string | undefined;
      const chain: string[] = [];
      while (alt !== undefined && !seen.has(alt.trim().toLowerCase())) {
        seen.add(alt.trim().toLowerCase());
        chain.push(alt);
        /* A hop this environment lacks but whose family the line-metrics table flags East
         * Asian: Word had that font and drew it, and its line is 1.3 x its usWin box, which no
         * substitute comes near.  The chain goes on for the *width* substitution; only the line
         * box is taken here. */
        if (eastAsianHop === undefined && isEastAsianFamily(alt) && this.registry.get(alt) === undefined) {
          eastAsianHop = alt;
        }
        pf = this.registry.get(alt) ?? this.registry.get(`${alt} Regular`) ?? this.get(alt);
        if (pf !== undefined) { resolvedAlt = alt; break; }
        alt = altNames.get(alt.trim().toLowerCase());
      }
      /* The line box is registered even where the chain found no face: Word had the alternate
       * and drew it, so the line is its whatever renders the glyphs. */
      if (eastAsianHop !== undefined) this.registerLineMetricsAlias(documentFontName, eastAsianHop);
      if (pf === undefined) continue;

      this.put(documentFontName, pf);
      if (eastAsianHop === undefined) this.registerLineMetricsAlias(documentFontName, resolvedAlt);
      this.decide(documentFontName, 'ALT_NAME', `w:altName ${chain.join(' -> ')}`);
    }
  }

  /** Whether {@link addClassBasedSubstitutes} applies to this mapper.  True for every mapper
   *  since docx4j 17.1.1. */
  wantsClassBasedSubstitutes(): boolean {
    return true;
  }

  /**
   * Map whatever is still unmapped to a font of the same class.  Without it an unmapped font
   * falls back to whatever the document's *default* font maps to, which is a Times clone
   * standing in for a sans as often as not.  Deliberately conservative: a condensed face is
   * left unmapped, its widths being further from an ordinary condensed face's than the
   * default's are.
   */
  addClassBasedSubstitutes(documentFontNames: Iterable<string> | undefined): void {
    if (documentFontNames === undefined) return;
    for (const documentFontName of documentFontNames) {
      if (documentFontName === undefined || documentFontName.trim().length === 0) continue;
      if (this.get(documentFontName) !== undefined) continue;
      if (this.isEmbedded(documentFontName)) continue;
      if (this.registry.get(documentFontName) !== undefined) continue;

      const pf = selectByClass(this.registry, documentFontName);
      if (pf !== undefined) {
        this.put(documentFontName, pf);
        this.decide(documentFontName, 'CLASS',
          `a face of the same class: ${substitutionClass(documentFontName)}`, UNKNOWN_ERROR);
      }
    }
  }

  /** The mapper's own guesses for what the measured passes left (`BestMatchingMapper`'s panose
   *  match and `FontSubstitutions.xml`).  Nothing here: that mapper needs font files. */
  addMapperSubstitutes(_documentFontNames: Iterable<string> | undefined, _fonts: wml.Fonts | undefined): void {
  }

  /**
   * Word's own answer for a font it cannot find, for whatever is still unmapped: Cambria or
   * Calibri (or Courier New) by the fontTable entry's family ({@link wordDefaultFor}), mapped
   * to whatever this mapper maps that font to.  Applied only to a family none of docx4j's
   * tables know ({@link isKnownFamily}).
   */
  addWordDefaultSubstitutes(documentFontNames: Iterable<string> | undefined, fonts: wml.Fonts | undefined): void {
    if (documentFontNames === undefined) return;
    const table = fontTableOf(fonts);
    for (const documentFontName of documentFontNames) {
      if (documentFontName === undefined || documentFontName.trim().length === 0) continue;
      if (this.get(documentFontName) !== undefined) continue;
      if (this.isEmbedded(documentFontName)) continue;
      if (this.registry.get(documentFontName) !== undefined) continue;
      if (isKnownFamily(documentFontName)) continue;

      const wordFont = wordDefaultFor(documentFontName, table);
      const pf = this.get(wordFont) ?? this.registry.get(wordFont) ?? selectByClass(this.registry, wordFont);
      if (pf === undefined) continue;
      this.put(documentFontName, pf);
      this.wordDefaulted.add(documentFontName.trim().toLowerCase());
      this.registerLineMetricsAlias(documentFontName, wordFont);
      this.decide(documentFontName, 'WORD_DEFAULT',
        `Word's own default for a font it cannot find: ${wordFont}`, UNKNOWN_ERROR);
    }
  }

  /**
   * For each document font with no bold face of its own, an alias of its mapped physical font
   * reporting none, so that the consumer synthesises bold at the regular advances as Word does.
   * Last of the passes: it re-maps, and it is not a source of its own - the decision says so by
   * reporting a synthetic bold face.
   */
  addNoBoldFaceAliases(documentFontNames: Iterable<string> | undefined): void {
    if (documentFontNames === undefined) return;
    for (const documentFontName of documentFontNames) {
      if (documentFontName === undefined || documentFontName.trim().length === 0) continue;
      if (hasBoldFace(documentFontName)) continue;
      if (this.isEmbedded(documentFontName)) continue; // the embedded forms say what the document has
      /* A font Word itself could not find is substituted whole - Calibri's real bold for its
       * w:b - however its name reads: "EnBW DIN Pro Light" is Calibri Bold in Word, not a
       * synthesised Calibri.  Measured (CR-016 phase 4 gate): with the alias that document
       * scored 0.8441, without it 0.8783. */
      if (this.wordDefaulted.has(documentFontName.trim().toLowerCase())) continue;
      const pf = this.get(documentFontName);
      if (pf === undefined || pf.noBoldFace) continue;
      this.fontMappings.set(documentFontName.toLowerCase(), pf.noBoldFaceAlias());
    }
  }

  /**
   * Every pass in order, as docx4j's `WordprocessingMLPackage.setFontMapper` runs them.
   *
   * @param documentFontNames the document's font names (`MainDocumentPart.fontsInUse()`)
   * @param fonts the font table part's content, for `w:altName` and `w:family`
   */
  populate(documentFontNames: Iterable<string>, fonts: wml.Fonts | undefined): void {
    const names = [...documentFontNames];
    this.populateFontMappings(names, fonts);
    this.addMetricallyCompatibleSubstitutes();
    this.addAltNameSubstitutes(names, fonts);
    if (this.wantsClassBasedSubstitutes()) this.addClassBasedSubstitutes(names);
    this.addMapperSubstitutes(names, fonts);
    this.addWordDefaultSubstitutes(names, fonts);
    this.addNoBoldFaceAliases(names);
  }
}

function set(map: Map<string, PhysicalFont>, name: string, pf: PhysicalFont | undefined): void {
  if (pf === undefined) map.delete(name); else map.set(name, pf);
}
