// The list definitions (docx4j `ListLevel`, `AbstractListNumberingDefinition`,
// `ListNumberingDefinition` and `NumberingDefinitionsPart.initialiseMaps`, CR-014 phase 4).
//
// Immutable once built and holding no counters: what a level *is*, never what it is up to.  The
// counting is `state.mts` and the resolution `Emulator.mts`.
//
//   NumberingDefinitions        one `w:numbering`
//     AbstractListDefinition    one `w:abstractNum` -> levels 0..8
//     ListDefinition            one `w:num` -> its abstract definition, with `w:lvlOverride`s
//     LevelDefinition           one `w:lvl`, the abstract level with the override over it
//
// Built from the tree the part holds - which the `PropertyResolver` reads with
// `readContents()`, so a document whose labels are only *read* still saves byte for byte.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { log } from '../properties/log.mjs';
import type { NumberingIndents } from '../properties/styleUtil.mjs';
import type { Counter, NumberingState } from './state.mjs';
import { formatValue } from './formats.mjs';

/** How many levels a list definition may have (ECMA-376: `w:ilvl` 0 to 8). */
export const LEVELS = 9;

/** What the definitions need of the styles: a style by id (the resolver's `getStyle`). */
export type StyleLookup = (styleId: string) => wml.Style | undefined;

/**
 * One `w:lvl`: the abstract definition's level, and - for a level of a `w:num` - the instance's
 * `w:lvlOverride/w:lvl` applied over it. docx4j `ListLevel`; the count lives in a
 * {@link NumberingState}, keyed by {@link abstractNumId} and {@link ilvl}.
 *
 * Five members are **merged** attribute by attribute, as docx4j's `setOverrides` merges them -
 * the override replaces each only where it states one: `w:start`, `w:lvlRestart`, `w:lvlText`,
 * `w:numFmt` (and so `isBullet`), and the `w:hAnsi` of the level's `w:rPr/w:rFonts`.  Everything
 * else is read off **one** `w:lvl`, the override where the instance has one and the abstract
 * level otherwise ({@link controllingLvl}): `w:isLgl` (docx4j's rule, measured through
 * `IsLglTest`), `w:pStyle`, `w:suff`, `w:lvlJc` and `w:lvlPicBulletId`.  `w:ind` and the label's
 * `w:rPr` have rules of their own; see {@link ind} and {@link labelRPr}.
 */
export class LevelDefinition {

  /** `w:ilvl`, as a string ("0" to "8"): the counter's key, and docx4j's map key. */
  readonly ilvl: string;

  /** The `w:abstractNum/w:lvl` this level was read from. */
  abstractLvl: wml.Lvl | undefined;

  /** The instance's `w:lvlOverride/w:lvl` for this level, or undefined. */
  overrideLvl: wml.Lvl | undefined;

  /** The **referencing** `w:abstractNum` (the counter's key; never the `w:numStyleLink` target's). */
  abstractNumId = '';

  /** The `w:num` an instance level belongs to; undefined for an abstract one. */
  ownerNumId: string | undefined;

  /** `w:start` **less one**, since every fetch increments first. */
  startValue = 0;

  /** This instance level carries a `w:startOverride`, spent on the `w:num`'s first use. */
  hasStartOverride = false;

  /** `w:lvlRestart` (ECMA-376 17.9.11), or undefined where the level states none. */
  lvlRestart: number | undefined;

  /** `w:lvlText`: the label pattern ("%1.%2."), or a bullet level's character. */
  lvlText: string | undefined;

  /** `w:numFmt`, or undefined where the level states none. */
  numFmt: wml.NumberFormat | undefined;

  /** Whether `w:numFmt` is `bullet`. */
  isBullet = false;

  /**
   * The `w:hAnsi` of the level's `w:rPr/w:rFonts`, or undefined.
   *
   * @deprecated docx4j deprecated `ListLevel.getFont()` in 17.1.1: use {@link labelRPr} and its
   *   `w:rFonts`.  Kept because `NumberingResult.numFont` is what the goldens record.
   */
  font: string | undefined;

  constructor(ilvl: string, lvl?: wml.Lvl) {
    this.ilvl = ilvl;
    this.abstractLvl = lvl;
    if (lvl !== undefined) this.read(lvl);
  }

  /** Take from a `w:lvl` whatever it states; leave the rest alone. docx4j's constructor body. */
  private read(lvl: wml.Lvl): void {
    if (lvl.start !== undefined) {
      // one less than the user set it to, since whenever we fetch the number we first increment
      this.startValue = lvl.start.val - 1;
    }
    if (lvl.lvlRestart?.val !== undefined) this.lvlRestart = lvl.lvlRestart.val;
    if (lvl.lvlText !== undefined) this.lvlText = lvl.lvlText.val;
    const hAnsi = lvl.rPr?.rFonts?.hAnsi;
    if (hAnsi !== undefined) this.font = hAnsi;
    if (lvl.numFmt !== undefined) {
      this.numFmt = lvl.numFmt.val;
      this.isBullet = lvl.numFmt.val === 'bullet';
    }
  }

  /** A copy for an instance definition. docx4j's copy constructor: no override, no owner. */
  copy(): LevelDefinition {
    const other = new LevelDefinition(this.ilvl);
    other.abstractLvl = this.abstractLvl;
    other.abstractNumId = this.abstractNumId;   // the counter is shared through the state, by this key
    other.startValue = this.startValue;
    other.lvlRestart = this.lvlRestart;
    other.lvlText = this.lvlText;
    other.numFmt = this.numFmt;
    other.isBullet = this.isBullet;
    other.font = this.font;
    return other;
  }

  /** Applies the instance's `w:lvlOverride/w:lvl`. docx4j `setOverrides`. */
  applyOverride(lvl: wml.Lvl): void {
    this.overrideLvl = lvl;
    this.read(lvl);
  }

  /** A `w:startOverride`: given effect on the first use of this `w:num` in a story. */
  setStartValue(value: number): void {
    this.startValue = value;
    this.hasStartOverride = true;
  }

  /** `w:ilvl` as a number. */
  get level(): number {
    const n = Number(this.ilvl);
    return Number.isFinite(n) ? n : 0;
  }

  /**
   * The one `w:lvl` the members docx4j does not merge are read from: the instance's override
   * where there is one, else the abstract level's.
   */
  get controllingLvl(): wml.Lvl | undefined {
    return this.overrideLvl ?? this.abstractLvl;
  }

  /** `w:start` as the document states it (one more than {@link startValue}). */
  get start(): number | undefined {
    return this.abstractLvl === undefined && this.overrideLvl === undefined ? undefined : this.startValue + 1;
  }

  /**
   * `w:isLgl`: the levels this number *inherits* print in decimal whatever their own `w:numFmt`
   * (Word's "legal style numbering": "Section 1.01" and not "Section I.01").
   */
  get isLgl(): boolean {
    const isLgl = this.controllingLvl?.isLgl;
    return isLgl !== undefined && isLgl.val !== false;
  }

  /** The `w:pStyle` this level is linked to (ECMA-376 17.9.24), or undefined. */
  get pStyle(): string | undefined {
    const override = this.overrideLvl;
    if (override !== undefined && override.pStyle !== undefined) return override.pStyle.val;
    return this.abstractLvl?.pStyle?.val;
  }

  /** `w:suff`: what separates the label from the text (tab, space, nothing). */
  get suff(): string | undefined {
    return this.controllingLvl?.suff?.val;
  }

  /** `w:lvlJc`: how the label is justified in its column. */
  get lvlJc(): wml.JcEnumeration | undefined {
    return this.controllingLvl?.lvlJc?.val;
  }

  /** `w:lvlPicBulletId`: the `w:numPicBullet` a picture bullet level names. */
  get lvlPicBulletId(): number | undefined {
    return this.controllingLvl?.lvlPicBulletId?.val;
  }

  /** The level's `w:pPr`: the override's where it has one, else the abstract level's. */
  get pPr(): wml.PPr | undefined {
    return this.overrideLvl?.pPr ?? this.abstractLvl?.pPr;
  }

  /**
   * The level's **own** `w:ind`: the instance's override level's where *it* states one, else the
   * abstract level's (docx4j `Emulator.getNumber`'s `triple.ind`, since 17.1.0).  The linked
   * `w:pStyle`'s is not consulted here; {@link NumberingDefinitions.getInd} does that.
   */
  get ind(): wml.PPrBase.Ind | undefined {
    return this.overrideLvl?.pPr?.ind ?? this.abstractLvl?.pPr?.ind;
  }

  /** The **abstract** level's `w:rPr`, never the override's. docx4j `NumberingResult.getRPr()`. */
  get rPr(): wml.RPr | undefined {
    return this.abstractLvl?.rPr;
  }

  /**
   * The `w:rPr` the label is drawn with: the `w:lvlOverride/w:lvl`'s where the instance
   * overrides this level with one carrying an rPr, else the abstract level's - one or the other,
   * as Word applies them (CR-014 probe P3, measured: an abstract `w:i` under an override
   * `w:b` + `w:sz 36` gives a bold 18pt label with no italic anywhere).  It formats the number
   * alone, never the paragraph's text (ECMA-376 17.9.24).
   */
  get labelRPr(): wml.RPr | undefined {
    const override = this.overrideLvl;
    return override !== undefined && override.rPr !== undefined ? override.rPr : this.rPr;
  }

  /** This level's counter in the given state. */
  counter(state: NumberingState): Counter {
    return state.counter(this.abstractNumId, this.ilvl, this.startValue);
  }

  /**
   * One more item at this level. docx4j `incrementCounter(state)`.
   *
   * The shared counter takes this level's start value the first time it is met in the state, and
   * again the first time this instance level's `w:num` is met there when that `w:num` overrides
   * the start - deferred until then, since otherwise earlier numbering over the same abstract
   * list would use it.
   */
  incrementCounter(state: NumberingState): void {
    const counter = this.counter(state);
    const overridePending = this.hasStartOverride && this.ownerNumId !== undefined
      && !state.startOverrideApplied(this.ownerNumId, this.ilvl);
    if (overridePending || !counter.encounteredAlready) {
      counter.value = this.startValue;
      counter.encounteredAlready = true;
      counter.resetPending = false;
      if (this.ownerNumId !== undefined) state.markStartOverrideApplied(this.ownerNumId, this.ilvl);
    }
    if (counter.resetPending) {
      // the reset already placed the counter at its start value (see resetCounter)
      counter.resetPending = false;
      return;
    }
    counter.increment();
  }

  /**
   * A shallower level was used: the level shows its start value from now until it is next used,
   * and that first use does not increment it (CR-014 probe P8; before, the counter went to
   * start-1 and a deeper label printed "2.0.1" where Word prints "2.1.1").
   */
  resetCounter(state: NumberingState): void {
    const counter = this.counter(state);
    counter.value = this.startValue + 1;
    counter.resetPending = true;
  }

  /**
   * Whether using `shallowerIlvl` (0-based) restarts this level. ECMA-376 17.9.11: without
   * `w:lvlRestart` any shallower level restarts it; `w:val="0"` means none does; `w:val="n"`
   * means levels 1..n (1-based, so ilvl 0..n-1) do and deeper ones do not.  Measured, P8.
   */
  restartsAfter(shallowerIlvl: number): boolean {
    if (this.lvlRestart === undefined) return true;
    if (this.lvlRestart <= 0) return false;
    return shallowerIlvl <= this.lvlRestart - 1;
  }

  /** The count at this level, in this level's `w:numFmt`. */
  currentValueFormatted(state: NumberingState, where?: string): string {
    return formatValue(this.numFmt, this.counter(state).value,
      where === undefined ? undefined : `${where} ilvl ${this.ilvl}`);
  }

  /** The count at this level as a decimal, whatever its `w:numFmt` (what `w:isLgl` shows). */
  currentValueUnformatted(state: NumberingState): string {
    return String(this.counter(state).value);
  }

  toString(): string {
    return `LevelDefinition ${this.ilvl} ${this.numFmt ?? '-'} ${JSON.stringify(this.lvlText ?? '')}`;
  }
}

/**
 * One `w:abstractNum`: its levels, and - where it carries `w:numStyleLink` - the numbering style
 * whose definition it takes them from. docx4j `AbstractListNumberingDefinition`.
 */
export class AbstractListDefinition {

  /** `w:abstractNumId`. */
  readonly id: string;

  /** The levels, keyed by `w:ilvl` ("0" to "8"). */
  readonly levels = new Map<string, LevelDefinition>();

  /** `w:numStyleLink`: the numbering style this definition takes its levels from. */
  readonly linkedStyleId: string | undefined;

  /** `w:styleLink`: the inverse pointer, read and not acted on, as docx4j. */
  readonly styleLink: string | undefined;

  constructor(readonly abstractNum: wml.Numbering.AbstractNum) {
    this.id = String(abstractNum.abstractNumId);
    this.linkedStyleId = abstractNum.numStyleLink?.val;
    this.styleLink = abstractNum.styleLink?.val;
    this.readLevels(abstractNum);
  }

  private readLevels(abstractNum: wml.Numbering.AbstractNum): void {
    for (const lvl of abstractNum.lvl ?? []) this.readLevel(lvl);
  }

  /** Adds (or replaces) the level read from this `w:lvl`. docx4j `readLevel`. */
  readLevel(lvl: wml.Lvl): void {
    const level = new LevelDefinition(String(lvl.ilvl), lvl);
    // the REFERENCING abstract list: a w:numStyleLink definition counts on its own (P2)
    level.abstractNumId = this.id;
    this.levels.set(level.ilvl, level);
  }

  /**
   * The second pass for a definition carrying `w:numStyleLink`: read its levels from the
   * abstract definition the numbering style's own `w:num` names.  They count under **this**
   * definition's id, not the linked one's. docx4j `updateDefinitionFromLinkedStyle`.
   */
  updateFromLinkedStyle(linked: wml.Numbering.AbstractNum): void {
    if (!this.hasLinkedStyle) return;
    this.readLevels(linked);
  }

  /** Whether this definition carries `w:numStyleLink`. */
  get hasLinkedStyle(): boolean {
    return this.linkedStyleId !== undefined && this.linkedStyleId !== '';
  }

  /** How many levels are defined; 0 for an unresolved `w:numStyleLink` definition. */
  get levelCount(): number {
    return this.levels.size;
  }
}

/**
 * One `w:num`: the abstract definition it names, with its `w:lvlOverride`s applied to copies of
 * that definition's levels, and the counting. docx4j `ListNumberingDefinition`.
 */
export class ListDefinition {

  /** `w:numId`. */
  readonly numId: string;

  /** The abstract definition this names, or undefined. */
  readonly abstractDefinition: AbstractListDefinition | undefined;

  /** The levels, with this instance's overrides applied. */
  readonly levels = new Map<string, LevelDefinition>();

  constructor(readonly num: wml.Numbering.Num,
    abstracts: ReadonlyMap<string, AbstractListDefinition>,
    resolveLinkedStyle: boolean) {

    this.numId = String(num.numId);

    const abstractNumId = num.abstractNumId as wml.Numbering.Num.AbstractNumId | undefined;
    if (abstractNumId === undefined) {
      log.warn(`No abstractNumId on w:numId=${this.numId}`);
      return;
    }
    const abstract = abstracts.get(String(abstractNumId.val));
    this.abstractDefinition = abstract;
    if (abstract === undefined) {
      log.warn(`No abstractListDefinition for w:numId=${this.numId}`);
      return;
    }
    if (abstract.levelCount === 0 && abstract.hasLinkedStyle && !resolveLinkedStyle) {
      // resolved on the second pass, once the numbering style is known
      return;
    }

    for (const [ilvl, level] of abstract.levels) {
      const instance = level.copy();
      instance.ownerNumId = this.numId;
      this.levels.set(ilvl, instance);
    }

    for (const override of num.lvlOverride ?? []) {
      if (override.ilvl === undefined) {
        log.warn(`Missing @w:ilvl on a w:lvlOverride of w:numId=${this.numId}`);
        continue;
      }
      const ilvl = String(override.ilvl);
      const level = this.levels.get(ilvl);
      // the w:startOverride first, as docx4j applies them: an override w:lvl stating its own
      // w:start then replaces the value, but the level still counts as start-overridden
      const start = override.startOverride?.val;
      if (start !== undefined) {
        if (level === undefined) {
          log.warn(`level ${ilvl} missing for abstractNum ${abstract.id}, referenced from w:num ${this.numId}`);
        } else {
          level.setStartValue(start - 1);
        }
      }
      if (override.lvl !== undefined && level !== undefined) level.applyOverride(override.lvl);
    }
  }

  /** The level with this `w:ilvl`, or undefined. */
  level(ilvl: string | number): LevelDefinition | undefined {
    return this.levels.get(String(ilvl));
  }

  /** Whether the level exists: Word writes `w:num`s whose abstract definition or level is missing. */
  levelExists(ilvl: string | number): boolean {
    return this.levels.has(String(ilvl));
  }

  /** Whether the level's `w:numFmt` is `bullet`. */
  isBullet(ilvl: string | number): boolean {
    return this.level(ilvl)?.isBullet ?? false;
  }

  /** The `w:hAnsi` font the level's `w:rPr` names, or undefined. */
  font(ilvl: string | number): string | undefined {
    return this.level(ilvl)?.font;
  }

  /**
   * One more item at a level, and the deeper levels that restart after it reset.
   * docx4j `incrementCounter(level, state)`.
   */
  incrementCounter(ilvl: string | number, state: NumberingState): void {
    const levelInt = Number(ilvl);
    const thisLevel = this.level(levelInt);
    if (thisLevel === undefined) return;   // guarded by levelExists

    if (!thisLevel.counter(state).encounteredAlready) {
      // make sure the shallower levels have been initialised
      for (let shallower = levelInt - 1; shallower >= 0; shallower--) {
        const level = this.level(shallower);
        if (level === undefined || level.counter(state).encounteredAlready) break;
        level.incrementCounter(state);
      }
    }

    thisLevel.incrementCounter(state);

    // the deeper levels go back to their start - each unless its w:lvlRestart says this level
    // does not restart it (ECMA-376 17.9.11, CR-014 phase 2)
    for (let deeper = levelInt + 1; ; deeper++) {
      const level = this.level(deeper);
      if (level === undefined) break;
      if (level.restartsAfter(levelInt)) level.resetCounter(state);
    }
  }

  /**
   * The label of a level: its `w:lvlText` with the counters filled in ("1.2.", "(c)").
   * docx4j `getCurrentNumberString(level, state)`.
   *
   * `w:isLgl` at **any** level shows the levels this number *inherits* in decimal whatever their
   * own `w:numFmt`, the level carrying it keeping its own: Word's Article / Section numbering
   * prints "Section 1.01" where the plain reading would print "Section I.01".
   */
  currentNumberString(ilvl: string | number, state: NumberingState): string {
    const controlling = this.level(ilvl);
    if (controlling === undefined) return '';   // guarded by levelExists
    const isLegal = controlling.isLgl;
    const thisLevel = Number.isFinite(Number(ilvl)) ? Number(ilvl) : -1;

    const format = controlling.lvlText ?? '';
    let out = '';
    for (let i = 0; i < format.length; i++) {
      const char = format.charAt(i);
      if (char !== '%') { out += char; continue; }
      // a trailing '%' names no level and is dropped, as docx4j's loop drops it
      if (i >= format.length - 1) continue;
      const digit = format.charAt(i + 1);
      if (digit < '0' || digit > '9') {
        // docx4j throws NumberFormatException here; a malformed w:lvlText is not worth an
        // exception, so the two characters stand as they are written
        out += char;
        continue;
      }
      i++;
      const levelId = Number(digit) - 1;   // the format string is 1-based
      const level = this.level(levelId);
      if (level === undefined) continue;   // Word's referential-integrity bugs (docx4j NPEs here)
      out += (isLegal && levelId < thisLevel)
        ? level.currentValueUnformatted(state)
        : level.currentValueFormatted(state, `numId ${this.numId}`);
    }
    return out;
  }
}

/**
 * The definitions of one `w:numbering`: every `w:abstractNum` and every `w:num`, with
 * `w:numStyleLink` resolved. docx4j `NumberingDefinitionsPart.initialiseMaps` and the two maps
 * it fills, plus `getInd` and `getLinkedStyleId`.
 *
 * Immutable once built; the numbering part rebuilds it when its tree changes.
 */
export class NumberingDefinitions implements NumberingIndents {

  /** `w:abstractNumId` -> the definition. docx4j `getAbstractListDefinitions()`. */
  readonly abstractListDefinitions = new Map<string, AbstractListDefinition>();

  /** `w:numId` -> the definition. docx4j `getInstanceListDefinitions()`. */
  readonly instanceListDefinitions = new Map<string, ListDefinition>();

  /**
   * @param numbering the `w:numbering` tree, or undefined for a package with no numbering part
   * @param getStyle  a style by id, for `w:numStyleLink` and a level's linked `w:pStyle`
   */
  constructor(readonly numbering: wml.Numbering | undefined, private readonly getStyle: StyleLookup = () => undefined) {
    if (numbering === undefined) return;

    // pass 1: every w:abstractNum, then every w:num
    let needSecondPass = false;
    for (const abstractNum of numbering.abstractNum ?? []) {
      const definition = new AbstractListDefinition(abstractNum);
      this.abstractListDefinitions.set(definition.id, definition);
      if (abstractNum.numStyleLink !== undefined) needSecondPass = true;
    }
    this.readInstances(numbering, false);
    if (!needSecondPass) return;

    // pass 2: resolveLinkedAbstractNum, then the instances again (which overwrites them)
    for (const definition of this.abstractListDefinitions.values()) this.resolveLinkedAbstractNum(definition);
    this.readInstances(numbering, true);
  }

  private readInstances(numbering: wml.Numbering, resolveLinkedStyle: boolean): void {
    for (const num of numbering.num ?? []) {
      const definition = new ListDefinition(num, this.abstractListDefinitions, resolveLinkedStyle);
      this.instanceListDefinitions.set(definition.numId, definition);
    }
  }

  /**
   * docx4j `resolveLinkedAbstractNum`: the numbering style `w:numStyleLink` names, the `w:num`
   * its `w:numPr` points at, and *that* `w:num`'s abstract definition, whose levels this one
   * takes while counting under its own id - "treated as a separate list by Word (ie its numbers
   * are incremented independently), and this code honours that".
   */
  private resolveLinkedAbstractNum(definition: AbstractListDefinition): void {
    if (!definition.hasLinkedStyle) return;
    const numStyleId = definition.linkedStyleId!;
    const style = this.getStyle(numStyleId);
    if (style === undefined) {
      log.warn(`For w:numStyleLink, couldn't find style ${numStyleId}`);
      return;
    }
    const numPr = style.pPr?.numPr;
    if (numPr === undefined) {
      log.warn(`For w:numStyleLink, style ${numStyleId} has no w:numPr`);
      return;
    }
    if (numPr.numId?.val === undefined) {
      log.warn(`For w:numStyleLink, style ${numStyleId} w:numPr has no w:numId`);
      return;
    }
    const concrete = String(numPr.numId.val);
    const target = this.instanceListDefinitions.get(concrete);
    if (target === undefined) {
      log.warn(`No ListDefinition entry with ID ${concrete}`);
      return;
    }
    const linked = target.abstractDefinition;
    if (linked === undefined || linked === definition) return;
    definition.updateFromLinkedStyle(linked.abstractNum);
    // docx4j also copies the levels into the underlying w:abstractNum object here; this port
    // does not, so that a numbering part nobody has edited still saves byte for byte
  }

  /** The definition of a `w:numId`, or undefined. */
  list(numId: string | number | undefined): ListDefinition | undefined {
    return numId === undefined ? undefined : this.instanceListDefinitions.get(String(numId));
  }

  /** The level of a `w:numId` / `w:ilvl`, or undefined. */
  level(numId: string | number | undefined, ilvl: string | number | undefined): LevelDefinition | undefined {
    return this.list(numId)?.level(ilvl === undefined || ilvl === '' ? '0' : ilvl);
  }

  /**
   * The paragraph style this level is linked to - the `w:pStyle` of ECMA-376 17.9.24, on the
   * instance's `w:lvlOverride/w:lvl` where it has one and on the abstract level otherwise - or
   * undefined. docx4j `NumberingDefinitionsPart.getLinkedStyleId`.
   */
  getLinkedStyleId(numId: string | undefined, ilvl: string | undefined): string | undefined {
    return this.level(numId, ilvl)?.pStyle;
  }

  /** docx4j `NumberingDefinitionsPart.getInd(NumPr)`: `w:ilvl` is optional and means level 0. */
  getInd(numPr: wml.PPrBase.NumPr): wml.PPrBase.Ind | undefined {
    if (numPr.numId?.val === undefined) return undefined;
    const ilvl = numPr.ilvl?.val === undefined ? '0' : String(numPr.ilvl.val);
    return this.indOf(String(numPr.numId.val), ilvl);
  }

  /**
   * The indent this level contributes. docx4j `NumberingDefinitionsPart.getInd(String, String)`:
   * the instance's override level first, then the abstract level, each read by
   * {@link indFromLvl}.
   */
  indOf(numId: string, ilvl: string | undefined): wml.PPrBase.Ind | undefined {
    const list = this.list(numId);
    if (list === undefined) return undefined;
    const level = list.level(ilvl ?? '0');
    if (level === undefined) return undefined;
    if (level.overrideLvl !== undefined) {
      const ind = this.indFromLvl(level.overrideLvl);
      if (ind !== undefined) return ind;
    }
    return level.abstractLvl === undefined ? undefined : this.indFromLvl(level.abstractLvl);
  }

  /**
   * **The level's own `w:pPr/w:ind` comes first** (ECMA-376 17.9.24: a `w:lvl/w:pPr` states the
   * paragraph properties applied to a paragraph at this level).  A `w:lvl/w:pStyle` only *links*
   * the level to a paragraph style; it does not make that style's indent the level's.  Where the
   * level states no indent of its own the linked style's is still used, following `w:basedOn` as
   * every other property a style contributes does (docx4j 17.1.0 and 17.1.1; CR-014 claim 7,
   * measured, and probe P5).
   *
   * A style reference in the level's `w:pPr`, but not also as a sibling of `w:pPr`, contributes
   * nothing here - and numbers nothing (CR-014 P4; `Emulator.resolve` decides "not numbered").
   */
  private indFromLvl(lvl: wml.Lvl): wml.PPrBase.Ind | undefined {
    if (lvl.pPr?.ind !== undefined) return lvl.pPr.ind;
    const linked = lvl.pStyle?.val;
    if (linked === undefined) return undefined;
    const seen = new Set<string>();
    let id: string | undefined = linked;
    while (id !== undefined && !seen.has(id)) {
      seen.add(id);
      const style: wml.Style | undefined = this.getStyle(id);
      if (style === undefined) {
        log.warn(`Couldn't find style ${id}`);
        return undefined;
      }
      if (style.pPr?.ind !== undefined) return style.pPr.ind;
      id = style.basedOn?.val;
    }
    return undefined;
  }
}
