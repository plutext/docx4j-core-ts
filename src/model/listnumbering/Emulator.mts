// docx4j `org.docx4j.model.listnumbering.Emulator` (CR-001 section 6.2 as CR-014 revised it).
//
// Emulates Word's list numbering: given a paragraph, the number (or the bullet) Word would put
// in front of it.  Three things happen, in this order:
//
//   1. **resolve** ({@link Emulator.resolve}) the paragraph's `w:numPr` - its own where it has
//      one, else its paragraph style's effective `w:pPr` through the `PropertyResolver`,
//      falling back to the `w:default="1"` paragraph style - into a {@link NumRef}: the numId
//      and ilvl to count with, whether the numId is direct, or why Word numbers nothing;
//   2. **count** in a {@link NumberingState}, one per story of one traversal;
//   3. **format** the level's `w:lvlText` with the counters filled in.
//
// The static methods are docx4j's, over a package; an {@link Emulator} instance is what
// `NumberingDefinitionsPart.getEmulator()` hands out, and carries the definitions, the resolver
// and the part's default state.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { PropertyResolver } from '../properties/PropertyResolver.mjs';
import { log } from '../properties/log.mjs';
import { NumberingState } from './state.mjs';
import type { LevelDefinition, NumberingDefinitions } from './definitions.mjs';

/**
 * Where a paragraph's numbering comes from once its own `w:numPr` and its style chain have been
 * read. docx4j `Emulator.NumRef` (17.1.1, CR-014 phase 3).
 */
export class NumRef {

  /** The list, or undefined when not numbered. */
  readonly numId: string | undefined;

  /** The level ("0" where nothing states one); undefined when not numbered. */
  readonly ilvl: string | undefined;

  /** Whether the numId is the paragraph's own `w:numPr`, not one its style contributed. */
  readonly direct: boolean;

  /**
   * Word paints no label and does not count the paragraph: nothing names a list (no `w:numPr`,
   * and no paragraph style - not even the default - carrying one), a `w:numId` of 0 turns
   * numbering off, or the level is linked to a different paragraph style than the one that
   * brought the numbering.
   */
  readonly notNumbered: boolean;

  /** Why not numbered; undefined otherwise. */
  readonly reason: string | undefined;

  private constructor(numId: string | undefined, ilvl: string | undefined, direct: boolean,
    notNumbered: boolean, reason: string | undefined) {
    this.numId = numId;
    this.ilvl = ilvl;
    this.direct = direct;
    this.notNumbered = notNumbered;
    this.reason = reason;
  }

  /** A paragraph Word numbers, at this list and level. */
  static numbered(numId: string, ilvl: string, direct: boolean): NumRef {
    return new NumRef(numId, ilvl, direct, false, undefined);
  }

  /** A paragraph Word does not number, and why. */
  static not(reason: string): NumRef {
    return new NumRef(undefined, undefined, false, true, reason);
  }

  toString(): string {
    return this.notNumbered
      ? `not numbered: ${this.reason}`
      : `numId ${this.numId} ilvl ${this.ilvl} (${this.direct ? 'direct' : 'from style'})`;
  }
}

/**
 * What a paragraph is numbered with: the label text (or the bullet), the level definition it
 * came from, the indent and the rPr that level contributes. docx4j `Emulator.NumberingResult`
 * (which `ResultTriple`, an acknowledged misnomer, subclasses until 17.2).
 */
export interface NumberingResult {
  /** The label: the level's `w:lvlText` with its counters filled in ("1.2.", "(c)"), or the
   *  bullet's character; undefined for a list or level with no definition. */
  numString: string | undefined;
  /** The `w:hAnsi` font the level's `w:rPr` names, or undefined.
   *  @deprecated docx4j deprecated `getNumFont()`: use {@link labelRPr} and its `w:rFonts`. */
  numFont: string | undefined;
  /** Whether the level's `w:numFmt` is `bullet` (docx4j's `getBullet() != null`). */
  isBullet: boolean;
  /** The level's `w:lvlText` where its `w:numFmt` is bullet, else undefined. docx4j `getBullet()`. */
  bullet: string | undefined;
  /** The level's **own** `w:ind`: the override level's where it states one, else the abstract
   *  level's.  `indResolved` is the one that also follows the level's linked `w:pStyle`. */
  ind: wml.PPrBase.Ind | undefined;
  /** `NumberingDefinitionsPart.getInd(numId, ilvl)`: the level's indent with the linked
   *  paragraph style (and its `w:basedOn` chain) followed where the level states none. */
  indResolved: wml.PPrBase.Ind | undefined;
  /** The **abstract** level's `w:rPr` (never the override's). docx4j `getRPr()`. */
  rPr: wml.RPr | undefined;
  /** The rPr the label is drawn with: the override level's where it has one, else the abstract
   *  level's - a replacement, not a merge (CR-014 probe P3). docx4j `getLabelRPr()`. */
  labelRPr: wml.RPr | undefined;
  /** The abstract level definition (`w:abstractNum/w:lvl`) this number came from. docx4j `getLvl()`. */
  lvl: wml.Lvl | undefined;
  /** The `w:numId` the label was counted in. */
  numId: string | undefined;
  /** The `w:ilvl`. */
  ilvl: string | undefined;
  /** The level definition, merged. Not docx4j's: a port convenience for CR-002 phase H. */
  level: LevelDefinition | undefined;
  /** The counter's value at this level, before formatting. Not docx4j's. */
  count: number | undefined;
  /** No label was produced. docx4j returns an *empty* `ResultTriple` where a `w:num` or its
   *  level has no definition; this flags that case rather than leaving every member undefined
   *  and the caller to guess. */
  notNumbered: boolean;
  /** Why there is no label; undefined otherwise. */
  reason: string | undefined;
}

/** As much of a `NumberingDefinitionsPart` as the static entry points use. */
export interface NumberingPartLike {
  readonly definitions: NumberingDefinitions;
  readonly numberingState: NumberingState;
  getEmulator(reset?: boolean): Emulator;
}

/** As much of a `WordprocessingMLPackage` as the static entry points use. */
export interface NumberedPackageLike {
  getMainDocumentPart(): { numberingDefinitionsPart?: NumberingPartLike | undefined };
  readonly propertyResolverOrUndefined?: PropertyResolver | undefined;
}

function numberingPartOf(pkg: NumberedPackageLike | undefined): NumberingPartLike | undefined {
  try {
    return pkg?.getMainDocumentPart().numberingDefinitionsPart;
  } catch {
    return undefined;   // no main document part
  }
}

/**
 * The numbering of one package's numbering part: the definitions, the `PropertyResolver` the
 * resolution goes through, and the part's default {@link NumberingState}.
 *
 * docx4j's `Emulator` is a marker object with static methods over maps the part holds; here the
 * instance carries what those statics look up, and the statics delegate to it.
 */
export class Emulator {

  private state: NumberingState;

  constructor(
    /** The definitions this emulator counts in. */
    readonly definitions: NumberingDefinitions,
    /** The resolver a style-contributed `w:numPr` is read through; undefined disables that path. */
    readonly resolver: PropertyResolver | undefined,
    state: NumberingState = new NumberingState()) {
    this.state = state;
  }

  /** The counters the state-less calls use. docx4j `NumberingDefinitionsPart.getNumberingState()`. */
  get numberingState(): NumberingState {
    return this.state;
  }

  /** Every list starts again: what `getEmulator(true)` does to the part's default state. */
  newState(): NumberingState {
    this.state = new NumberingState();
    return this.state;
  }

  // ---------------------------------------------------------------- resolution

  /**
   * Whether this level belongs to a **different** paragraph style than the one which brought the
   * numbering to this paragraph, in which case Word paints no label and does not count the
   * paragraph. docx4j `Emulator.styleLinkedElsewhere` (17.1.0, measured on `numbering-label-ilvl0`).
   *
   * ECMA-376 17.9.24's `w:pStyle` inside a `w:lvl` names the paragraph style the level is linked
   * to.  The style itself, or any style it is based on, counts as the level's own.  Direct
   * formatting is never suppressed: a paragraph whose *own* `w:numPr` names such a level is
   * numbered.
   */
  styleLinkedElsewhere(numId: string, ilvl: string, pStyleVal: string | undefined): boolean {
    const linked = this.definitions.getLinkedStyleId(numId, ilvl);
    if (linked === undefined || linked === '') return false;
    if (pStyleVal === undefined) return true;
    const seen = new Set<string>();
    let id: string | undefined = pStyleVal;
    while (id !== undefined && !seen.has(id)) {
      seen.add(id);
      if (linked === id) return false;
      const style: wml.Style | undefined = this.resolver?.getStyle(id);
      id = style?.basedOn?.val;
    }
    return true;
  }

  /**
   * The one resolution of numId and ilvl for a paragraph. docx4j `Emulator.resolve`
   * (CR-014 phase 3, which collapsed `getNumber`'s and `getInd`'s two copies onto it).
   *
   * From the paragraph's own `w:numPr` where it has one; otherwise from the *effective* `w:pPr`
   * of its paragraph style - the `w:default="1"` paragraph style when it names none, since that
   * style may itself be numbered (measured, CR-014 P6) - then the `w:numId` 0 rule and the
   * §2.8 rule for a style-contributed level linked to another style.
   *
   * @param numId       the paragraph's numId, or undefined/"" to read the style
   * @param ilvl        the paragraph's ilvl, or undefined/"" for the style's, else "0"
   * @param directNumPr whether a numId given here is the paragraph's own
   */
  resolve(pStyleVal: string | undefined, numId: string | undefined, ilvl: string | undefined,
    directNumPr: boolean): NumRef {

    let styleId = pStyleVal === undefined || pStyleVal === '' ? undefined : pStyleVal;
    let direct = directNumPr;
    let list = numId === undefined || numId === '' ? undefined : numId;
    let level = ilvl === undefined || ilvl === '' ? undefined : ilvl;

    if (list === undefined) {
      // no explicit numId: is it provided by the style (ie does this style, or the default
      // paragraph style, have a list associated with it)?
      direct = false;
      if (styleId === undefined) {
        styleId = this.resolver?.getDefaultParagraphStyleId();
        if (styleId === undefined) {
          return NumRef.not('no numId, no paragraph style and no default paragraph style');
        }
      }
      let pPr: wml.PPr | undefined;
      try {
        pPr = this.resolver?.getEffectivePPr(styleId);
      } catch (e) {
        log.warn(String(e));
        return NumRef.not(`cyclic styles at ${styleId}`);
      }
      if (pPr === undefined) return NumRef.not(`style '${styleId}' has no pPr`);
      const numPr = pPr.numPr;
      if (numPr === undefined) {
        // no numbering set on the style either; that's ok
        return NumRef.not(`no numId, and style '${styleId}' is not numbered`);
      }
      if (numPr.numId?.val === undefined) {
        return NumRef.not(`style '${styleId}' has a w:numPr without a w:numId val`);
      }
      list = String(numPr.numId.val);
      if (level === undefined) {
        // w:ilvl (and its w:val) is optional; its absence means level 0
        level = numPr.ilvl?.val === undefined ? '0' : String(numPr.ilvl.val);
      }
    }

    if (list === '0') {
      // ECMA-376 17.9.18: a w:numId of 0 never references a definition; it designates the
      // removal of numbering at this level (a paragraph or style switching an inherited list
      // off).  docx4j adopted this reading, and these words, on 2026-09-19 (01d661547).
      return NumRef.not(`${direct ? "the paragraph's" : `style '${styleId}'s`} w:numId 0 turns numbering off`);
    }

    if (level === undefined) {
      log.warn('No level id?! Default to 0.');
      level = '0';
    }

    if (!direct && this.styleLinkedElsewhere(list, level, styleId)) {
      return NumRef.not(`level ${level} of numId ${list} is linked to a paragraph style other than '${styleId}'`);
    }
    return NumRef.numbered(list, level, direct);
  }

  /**
   * The numbering reference a paragraph resolves to, without taking a number and without
   * touching any state. docx4j `Emulator.numRefFor` (17.1.1); never undefined.
   */
  numRefFor(pPr: wml.PPr | undefined): NumRef {
    if (pPr === undefined) return NumRef.not('no pPr');
    const [pStyleVal, numId, ilvl] = readNumPr(pPr);
    return this.resolve(pStyleVal, numId, ilvl, numId !== undefined && numId !== '');
  }

  // ---------------------------------------------------------------- counting

  /**
   * The next number for a paragraph, from its `w:pPr` (its own `w:numPr`, else its style's),
   * counted in `state` or in this emulator's default state.
   *
   * @returns undefined where the paragraph is not numbered
   */
  getNumber(pPr: wml.PPr | undefined, state?: NumberingState): NumberingResult | undefined {
    if (pPr === undefined) return undefined;
    const [pStyleVal, numId, ilvl] = readNumPr(pPr);
    return this.getNumberOf(pStyleVal, numId, ilvl, numId !== undefined && numId !== '', state);
  }

  /**
   * The next number given the paragraph's style and, where it has one, its own numId and ilvl.
   * docx4j's `getNumber(pkg, pStyleVal, numId, levelId, directNumPr, state)`.
   */
  getNumberOf(pStyleVal: string | undefined, numId: string | undefined, ilvl: string | undefined,
    directNumPr = numId !== undefined && numId !== '', state?: NumberingState): NumberingResult | undefined {

    const counters = state ?? this.state;
    const ref = this.resolve(pStyleVal, numId, ilvl, directNumPr);
    // docx4j logs the reason at DEBUG here; this package has no debug level, and a paragraph
    // that is simply not in a list is the common case, so nothing is logged
    if (ref.notNumbered) return undefined;
    return this.number(ref, counters);
  }

  /** Counts a resolved reference. The half of `getNumber` after `resolve`. */
  private number(ref: NumRef, state: NumberingState): NumberingResult {
    const list = this.definitions.list(ref.numId);
    const levelId = ref.ilvl ?? '0';
    if (list === undefined || !list.levelExists(levelId)) {
      // Word has been seen to write a w:num whose abstract definition is missing, or which
      // lacks the level; docx4j logs and returns an empty ResultTriple
      const reason = list === undefined
        ? `Couldn't find list ${ref.numId}`
        : `Couldn't find level ${levelId} in list ${ref.numId}`;
      log.warn(reason);
      return emptyResult(ref, reason);
    }

    list.incrementCounter(levelId, state);
    const level = list.level(levelId)!;
    const numString = list.currentNumberString(levelId, state);
    const font = level.font;
    return {
      numString,
      numFont: font === undefined || font === '' ? undefined : font,
      isBullet: level.isBullet,
      bullet: level.isBullet ? level.lvlText : undefined,
      ind: level.ind,
      indResolved: this.definitions.indOf(list.numId, levelId),
      rPr: level.rPr,
      labelRPr: level.labelRPr,
      lvl: level.abstractLvl,
      numId: list.numId,
      ilvl: level.ilvl,
      level,
      count: level.counter(state).value,
      notNumbered: false,
      reason: undefined,
    };
  }

  /** What {@link getNumber} would return, leaving the state as it was. docx4j `peek`. */
  peek(pPr: wml.PPr | undefined, state?: NumberingState): NumberingResult | undefined {
    return this.getNumber(pPr, (state ?? this.state).copy());
  }

  /**
   * The indent a paragraph's numbering level contributes, resolved the way {@link getNumber}
   * resolves the level itself and read as `NumberingDefinitionsPart.getInd` reads it: the
   * level's own `w:ind` first, then the linked style's, following `w:basedOn`.  Undefined where
   * the paragraph is not numbered or the level states none. docx4j `Emulator.getInd`.
   */
  getInd(pStyleVal: string | undefined, numId: string | undefined, ilvl: string | undefined): wml.PPrBase.Ind | undefined {
    const ref = this.resolve(pStyleVal, numId, ilvl, numId !== undefined && numId !== '');
    if (ref.notNumbered) return undefined;
    return this.definitions.indOf(ref.numId!, ref.ilvl);
  }

  // ---------------------------------------------------------------- docx4j's static entry points

  /** The package's emulator, or undefined where it has no numbering part. */
  static of(pkg: NumberedPackageLike | undefined): Emulator | undefined {
    return numberingPartOf(pkg)?.getEmulator();
  }

  /**
   * docx4j `Emulator.getNumber(wmlPackage, pPr[, state])` and
   * `Emulator.getNumber(wmlPackage, pStyleVal, numId, levelId[, directNumPr[, state]])`.
   *
   * The two are told apart as docx4j's overloads are by their signatures: a `string` second
   * argument, or more than three arguments, is the style form; anything else is the `w:pPr` one.
   */
  static getNumber(pkg: NumberedPackageLike, pPr: wml.PPr | undefined, state?: NumberingState): NumberingResult | undefined;
  static getNumber(pkg: NumberedPackageLike, pStyleVal: string | undefined, numId: string | undefined,
    ilvl?: string | undefined, directNumPr?: boolean, state?: NumberingState): NumberingResult | undefined;
  static getNumber(pkg: NumberedPackageLike, second: wml.PPr | string | undefined,
    ...rest: unknown[]): NumberingResult | undefined {
    const emulator = Emulator.of(pkg);
    if (emulator === undefined) return undefined;
    if (typeof second === 'string' || rest.length > 1) {
      const [numId, ilvl, directNumPr, state] = rest as [string | undefined, string | undefined, boolean | undefined, NumberingState | undefined];
      return emulator.getNumberOf(second as string | undefined, numId, ilvl,
        directNumPr ?? (numId !== undefined && numId !== ''), state);
    }
    return emulator.getNumber(second as wml.PPr | undefined, rest[0] as NumberingState | undefined);
  }

  /** docx4j `Emulator.peek(wmlPackage, pPr, state)`: the number, without taking it. */
  static peek(pkg: NumberedPackageLike, pPr: wml.PPr | undefined, state?: NumberingState): NumberingResult | undefined {
    const emulator = Emulator.of(pkg);
    return emulator?.peek(pPr, state);
  }

  /**
   * docx4j `Emulator.numRefFor(wmlPackage, pPr)`: where this paragraph's numbering resolves to,
   * before any counting.  Never undefined; a package with no numbering part, a null `w:pPr` or
   * a package whose `PropertyResolver` has not been built each answer {@link NumRef.notNumbered}
   * with the reason.
   */
  static numRefFor(pkg: NumberedPackageLike | undefined, pPr: wml.PPr | undefined): NumRef {
    if (pPr === undefined) return NumRef.not('no pPr');
    const part = numberingPartOf(pkg);
    if (part === undefined) return NumRef.not('no numbering part');
    const resolver = pkg?.propertyResolverOrUndefined;
    if (resolver === undefined) {
      return NumRef.not('no property resolver: await getPropertyResolver() first');
    }
    return part.getEmulator().numRefFor(pPr);
  }

  /** docx4j `Emulator.getInd(wmlPackage, pStyleVal, numId, levelId)`. */
  static getInd(pkg: NumberedPackageLike, pStyleVal: string | undefined, numId: string | undefined,
    ilvl: string | undefined): wml.PPrBase.Ind | undefined {
    return Emulator.of(pkg)?.getInd(pStyleVal, numId, ilvl);
  }

  toString(): string {
    return `Emulator(${this.definitions.instanceListDefinitions.size} lists, `
      + `${this.definitions.abstractListDefinitions.size} definitions)`;
  }
}

/** The paragraph style and the two `w:numPr` values a `w:pPr` states, each as a string. */
function readNumPr(pPr: wml.PPr): [string | undefined, string | undefined, string | undefined] {
  const pStyleVal = pPr.pStyle?.val;
  const numPr = pPr.numPr;
  const numId = numPr?.numId?.val === undefined ? undefined : String(numPr.numId.val);
  const ilvl = numPr?.ilvl?.val === undefined ? undefined : String(numPr.ilvl.val);
  return [pStyleVal, numId, ilvl];
}

function emptyResult(ref: NumRef, reason: string): NumberingResult {
  return {
    numString: undefined, numFont: undefined, isBullet: false, bullet: undefined,
    ind: undefined, indResolved: undefined, rPr: undefined, labelRPr: undefined, lvl: undefined,
    numId: ref.numId, ilvl: ref.ilvl, level: undefined, count: undefined,
    notNumbered: true, reason,
  };
}
