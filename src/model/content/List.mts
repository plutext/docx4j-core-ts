// CR-002 phase H (section 3.9): `Word.List` and `Word.ListItem` over `w:numPr` and the
// numbering part.
//
// Views like the rest of the content API: a `List` holds a `w:numId` and the body it was found
// through, a `ListItem` holds a paragraph, and nothing is cached.  Reads go through the
// numbering part's **definitions** (CR-001 section 15.2), which are built from a private read,
// so a document whose lists are only read still saves `word/numbering.xml` byte for byte; the
// first write promotes that tree to the part's live contents (`makeLive`) and, where another
// `w:num` shares the `w:abstractNum`, copies the abstract definition first so that the change
// stays local to this list.
//
// The label a list item shows (`listString`) is `Emulator.getNumber` counted in a fresh
// `NumberingState` walked over the paragraph's **story** in document order, which is the story
// rule of CR-001 section 15.2: headers and footers share one state, each notes part has its
// own, and a text box is a story of its own.  That costs a walk of the story per call;
// `Body.listLabels()` does the walk once and answers for every paragraph.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy, unmarshalNode } from '@docx4j/generated-objects-ts';
import * as f from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { parseXml } from '../../xml/dom.mjs';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { DEFAULT_NUMBERING_XML } from '../../parts/wml/defaultNumbering.mjs';
import { NumberingState, storyKindOf } from '../listnumbering/state.mjs';
import type { NumberingDefinitions, ListDefinition, LevelDefinition } from '../listnumbering/definitions.mjs';
import { LEVELS } from '../listnumbering/definitions.mjs';
import type { Emulator, NumRef } from '../listnumbering/Emulator.mjs';
import { isElement, linkParents, type Element } from './tree.mjs';
import { Font, type RPrHolder } from './Font.mjs';
import type { Body } from './Body.mjs';
import type { Paragraph } from './Paragraph.mjs';
import type { AlignmentOrUnknown } from './Paragraph.mjs';

/** Office JS `Word.ListLevelType`: what a level paints in front of its paragraphs. */
export type ListLevelType = 'Bullet' | 'Number' | 'Picture';

/** Office JS `Word.ListNumbering`: the number formats `setLevelNumbering` takes. */
export type ListNumbering = 'None' | 'Arabic' | 'UpperRoman' | 'LowerRoman' | 'UpperLetter' | 'LowerLetter';

/** Office JS `Word.ListBullet`: the bullets `setLevelBullet` takes. */
export type ListBullet = 'Custom' | 'Solid' | 'Hollow' | 'Square' | 'Diamonds' | 'Arrow' | 'Checkmark';

/** What `Paragraph.startNewList()` builds (extension: Office JS's takes no options). */
export interface StartListOptions {
  /** docx4j's default **bullet** set rather than its decimal set (`w:abstractNum` 0, not 1). */
  bullet?: boolean;
  /** The level to put this paragraph at; 0 by default. */
  level?: number;
}

/** One paragraph's place in its list, as one walk of the story answers it (`Body.listLabels`). */
export interface ListLabel {
  /** The label Word paints: the level's `w:lvlText` with the counters filled in, or the bullet. */
  listString: string;
  /** The `w:ilvl` the paragraph counts at, style-contributed levels included. */
  level: number;
  /** The `w:numId` it counts in. */
  numId: string;
  /** Its 0-based index among the items of this list and level since the level last restarted. */
  siblingIndex: number;
  /** Whether the level's `w:numFmt` is `bullet`. */
  isBullet: boolean;
}

// --------------------------------------------------------------- reaching the package, by shape

/**
 * The content API must not import `WordprocessingMLPackage` or `parts/wml` at runtime (that
 * would be a cycle, since `parts/wml` builds the `Body` views), so the numbering part is
 * reached by shape, as `Body.propertyResolver` and `fonts/lookup` are.
 */
interface NumberingPartLike {
  readonly isUnmarshalled: boolean;
  readonly contents: wml.Numbering;
  readonly definitions: NumberingDefinitions;
  getEmulator(reset?: boolean): Emulator;
  makeLive(): wml.Numbering;
  getContents(): Promise<wml.Numbering>;
}

/** As much of an `XmlPart` as the story walk reads. */
interface StoryPartLike {
  readonly contentType: string;
  readonly isUnmarshalled: boolean;
  readonly contents: unknown;
}

/** As much of a `MainDocumentPart` as the list views use. */
interface MainPartLike {
  numberingDefinitionsPart?: NumberingPartLike;
  readonly headerParts: StoryPartLike[];
  readonly footerParts: StoryPartLike[];
  addTargetPart(part: unknown): unknown;
}

/** As much of a `WordprocessingMLPackage` as the list views use. */
interface ListPackageLike {
  getMainDocumentPart(): MainPartLike;
  refreshPropertyResolver(): void;
  refresh(): Promise<unknown>;
}

function packageOf(body: Body): ListPackageLike | undefined {
  const pkg = body.package_ as unknown as ListPackageLike | undefined;
  return typeof pkg?.getMainDocumentPart === 'function' ? pkg : undefined;
}

function mainOf(body: Body): MainPartLike | undefined {
  try {
    return packageOf(body)?.getMainDocumentPart();
  } catch {
    return undefined;   // no main document part
  }
}

/** The numbering part, or undefined where the document has none. */
function numberingPartOf(body: Body): NumberingPartLike | undefined {
  return mainOf(body)?.numberingDefinitionsPart;
}

/** The definitions, or undefined where there is no numbering part or nothing has read it. */
function definitionsOf(body: Body): NumberingDefinitions | undefined {
  try {
    return numberingPartOf(body)?.definitions;
  } catch {
    return undefined;   // the part has not been read (await pkg.getPropertyResolver())
  }
}

/**
 * The package's numbering emulator, or undefined where the document has no lists (or nothing
 * has read the numbering part: `await pkg.getBody()` builds the resolver, which reads it).
 */
export function numberingEmulatorOf(body: Body): Emulator | undefined {
  try {
    return numberingPartOf(body)?.getEmulator();
  } catch {
    return undefined;
  }
}

/** Where a paragraph's numbering resolves to (docx4j `Emulator.numRefFor`), or undefined. */
function numRefOf(paragraph: Paragraph): NumRef | undefined {
  const ref = numberingEmulatorOf(paragraph.parentBody)?.numRefFor(paragraph.p.pPr);
  return ref === undefined || ref.notNumbered ? undefined : ref;
}

// --------------------------------------------------------------------------- the story walk

/**
 * Every paragraph under `root` in document order, each with the {@link NumberingState} it
 * counts in. `test/parity.test.mjs`'s `walkNumbering`, which is the harness's own walk: a text
 * box's paragraphs stay in this story's list (they are in the part, in document order) but
 * count in a story of their own, which is docx4j's rule (CR-014 probe P7).
 */
function walkStory(root: unknown, state: NumberingState, emit: (p: wml.P, state: NumberingState) => void): void {
  const walk = (value: unknown, current: NumberingState): void => {
    if (Array.isArray(value)) {
      for (const item of value) walk(item, current);
      return;
    }
    if (typeof value !== 'object' || value === null) return;
    if (isElement(value)) { walk((value as Element).value, current); return; }
    const node = value as Record<string, unknown> & { TYPE_NAME?: string };
    let inner = current;
    if (node.TYPE_NAME === 'org_docx4j_wml.P') emit(node as unknown as wml.P, current);
    else if (node.TYPE_NAME === 'org_docx4j_wml.CTTxbxContent') inner = new NumberingState();
    for (const key of Object.keys(node)) {
      if (key === 'PARENT' || key === 'TYPE_NAME') continue;
      walk(node[key], inner);
    }
  };
  walk(root, state);
}

/**
 * The parts whose paragraphs share one story with this one, in docx4j's order: every header
 * then every footer for a header or a footer, and the part alone otherwise (CR-001 section
 * 15.2, `NumberingStates.forPart`).
 */
function storyPartsOf(body: Body, part: StoryPartLike): StoryPartLike[] {
  if (storyKindOf(part) !== 'headersFooters') return [part];
  const main = mainOf(body);
  return main === undefined ? [part] : [...main.headerParts, ...main.footerParts];
}

/**
 * Every numbered paragraph of this body's **story**, with the label and the counters, in one
 * walk. A header or footer that has not been unmarshalled contributes nothing to the count,
 * since it cannot be read synchronously: `await part.getBody()` on each first where a footer's
 * numbers must continue the headers' (the story rule).
 */
export function listLabelsOf(body: Body): Map<wml.P, ListLabel> {
  const out = new Map<wml.P, ListLabel>();
  const emulator = numberingEmulatorOf(body);
  const part = body.part as unknown as StoryPartLike | undefined;
  if (emulator === undefined || part === undefined) return out;
  const state = new NumberingState();
  for (const story of storyPartsOf(body, part)) {
    if (!story.isUnmarshalled) continue;
    walkStory(story.contents, state, (p, counters) => {
      const result = emulator.getNumber(p.pPr, counters);
      if (result === undefined || result.numString === undefined) return;
      out.set(p, {
        listString: result.numString,
        level: Number(result.ilvl ?? '0'),
        numId: result.numId ?? '',
        siblingIndex: (result.count ?? 1) - (result.level?.start ?? 1),
        isBullet: result.isBullet,
      });
    });
  }
  return out;
}

/** The label of one paragraph, from a walk of its story. */
function labelOf(paragraph: Paragraph): ListLabel | undefined {
  return listLabelsOf(paragraph.parentBody).get(paragraph.p);
}

// ------------------------------------------------------------------------- the null objects

/**
 * Office JS's `*OrNullObject` result: `isNullObject` is true and every other member throws.
 * `src/office-js/proxy.mts` has the same shape for collections; it is repeated here rather
 * than imported so that the content API does not pull the shim's proxies into a bundle.
 */
function nullObject<T>(className: string): T {
  const fail = (member: string): never => {
    throw new ItemNotFound(`Word.${className} is a null object (this paragraph is not a list item); check isNullObject before using ${member}`);
  };
  const proxy: unknown = new Proxy({ isNullObject: true } as object, {
    get(target, prop): unknown {
      if (typeof prop === 'symbol') return Reflect.get(target, prop);
      if (prop === 'isNullObject') return true;
      if (prop === 'toJSON') return () => ({ isNullObject: true });
      if (prop === 'then') return undefined;
      return fail(prop);
    },
    set(_target, prop): boolean {
      return fail(String(prop));
    },
  });
  return proxy as T;
}

/**
 * Nothing matched, as Office JS's `ItemNotFound` says it: the same `code` as the shim's
 * `ItemNotFoundError`, so add-in code branching on `error.code` behaves the same.
 */
class ItemNotFound extends Error {
  readonly code = 'ItemNotFound';
  constructor(message: string) {
    super(message);
    this.name = 'ItemNotFoundError';
  }
}

export { ItemNotFound as ListItemNotFoundError };

// -------------------------------------------------------------------------------- the views

/**
 * A subset of Office JS `Word.List` over one `w:num` and the `w:abstractNum` it names. A light
 * view: it holds the `w:numId` and the body it was found through, and resolves both elements
 * through the numbering part's definitions on every call.
 *
 * Every `setLevel*` method writes on the **abstract** definition, copying it first where
 * another `w:num` shares it ({@link separate}), so that the change stays local to this list;
 * each then refreshes the package's resolver, which rebuilds the numbering definitions.
 */
export class List {
  constructor(
    /** `w:numId`, as a string (docx4j's key; {@link id} is Office JS's number). */
    readonly numId: string,
    /** The body this list was found through: where {@link paragraphs} looks. */
    readonly body: Body,
  ) {}

  /** Office JS `list.id`: the `w:numId`. */
  get id(): number {
    return Number(this.numId);
  }

  /** The definitions this list is read from; throws where the numbering part has not been read. */
  private get definitions(): NumberingDefinitions {
    const definitions = definitionsOf(this.body);
    if (definitions === undefined) {
      throw new Docx4JException('This document has no numbering definitions; await pkg.getBody() (or getPropertyResolver()) first');
    }
    return definitions;
  }

  /** The `w:num`'s definition (docx4j `ListNumberingDefinition`). */
  get definition(): ListDefinition {
    const definition = this.definitions.list(this.numId);
    if (definition === undefined) throw new ItemNotFound(`No w:num for numId ${this.numId}`);
    return definition;
  }

  /** The `w:num` (extension: the tree). */
  get element(): wml.Numbering.Num {
    return this.definition.num;
  }

  /** The `w:abstractNum` this list resolves to (extension: the tree). */
  get abstractElement(): wml.Numbering.AbstractNum | undefined {
    return this.definition.abstractDefinition?.abstractNum;
  }

  /** The level definition (the abstract level with this `w:num`'s override over it), or undefined. */
  level(level: number): LevelDefinition | undefined {
    return this.definition.level(level);
  }

  /** Office JS `list.levelExistences`: whether each of the nine levels is defined. */
  get levelExistences(): boolean[] {
    return Array.from({ length: LEVELS }, (_, i) => this.definition.levelExists(i));
  }

  /** Whether the level is defined (extension; Office JS has {@link levelExistences}). */
  levelExists(level: number): boolean {
    return this.definition.levelExists(level);
  }

  /**
   * Office JS `list.levelTypes`: `Bullet`, `Number` or `Picture` for each of the nine levels,
   * from the level's `w:numFmt` and its `w:lvlPicBulletId`. A level that is not defined reads
   * `Number`, as Office JS reports for a level with no format of its own.
   */
  get levelTypes(): ListLevelType[] {
    return Array.from({ length: LEVELS }, (_, i) => {
      const level = this.level(i);
      if (level === undefined) return 'Number';
      if (level.lvlPicBulletId !== undefined) return 'Picture';
      return level.isBullet ? 'Bullet' : 'Number';
    });
  }

  /** Office JS `list.getLevelString(level)`: the level's `w:lvlText` ('%1.', '(%2)', a bullet). */
  getLevelString(level: number): string {
    return this.level(level)?.lvlText ?? '';
  }

  /**
   * Office JS `list.getLevelFont(level)`: the `Font` the level's label is drawn with, over the
   * level's `w:rPr`.
   *
   * **It is a writable view**, and Office JS pairs it with `resetLevelFont`, so it
   * {@link separate}s first: the numbering part becomes live and a shared abstract definition
   * is copied, exactly as the `setLevel*` methods do. A read that must not cost the part its
   * byte-for-byte round trip is `list.level(n).labelRPr` (CR-002 section 17).
   */
  getLevelFont(level: number): Font {
    this.separate();
    const holder = this.writableLvl(level) as RPrHolder;
    return new Font(() => [holder]);
  }

  /** Office JS `list.getLevelParagraphs(level)`: this body's paragraphs at that level of this list. */
  getLevelParagraphs(level: number): Paragraph[] {
    return this.paragraphs.filter((p) => levelOf(p) === level);
  }

  /**
   * Office JS `list.paragraphs`: every paragraph of this body whose effective `w:numPr` names
   * this list, in document order (a style-contributed `w:numPr` counts).
   */
  get paragraphs(): Paragraph[] {
    return this.body.paragraphs.filter((p) => numRefOf(p)?.numId === this.numId);
  }

  /**
   * Office JS `list.insertParagraph(text, location)`: a new paragraph of this list before its
   * first item or after its last, at that item's level.
   */
  insertParagraph(text: string, location: 'Start' | 'End'): Paragraph {
    const items = this.paragraphs;
    if (items.length === 0) {
      throw new ItemNotFound(`List ${this.id} has no paragraphs in this body to insert relative to`);
    }
    const anchor = location === 'Start' ? items[0]! : items[items.length - 1]!;
    const paragraph = anchor.insertParagraph(text, location === 'Start' ? 'Before' : 'After');
    paragraph.attachToList(this.id, levelOf(anchor) ?? 0);
    return paragraph;
  }

  // --- the level writers ------------------------------------------------------------------

  /**
   * Office JS `list.setLevelNumbering(level, listNumbering, formatString)`: the level's
   * `w:numFmt` and, from `formatString`, its `w:lvlText` - each number in the array is a
   * 0-based level whose counter goes there (`[0, '.', 1]` is `'%1.%2'`), each string a literal.
   * Without one the level numbers itself and takes a trailing full stop (`'%<n>.'`).
   */
  setLevelNumbering(level: number, listNumbering: ListNumbering, formatString?: (string | number)[]): void {
    const lvl = this.writableLvl(level);
    lvl.numFmt = f.createNumFmt({ val: NUMBERING_FORMATS[listNumbering] });
    linkParents(lvl.numFmt, lvl);
    lvl.lvlText = f.createLvlLvlText({ val: lvlTextOf(level, formatString) });
    linkParents(lvl.lvlText, lvl);
    this.written();
  }

  /**
   * Office JS `list.setLevelBullet(level, listBullet, charCode, fontName)`: the level's
   * `w:numFmt` becomes `bullet` and its `w:lvlText` the bullet character, with the font in the
   * level's `w:rPr/w:rFonts`. `Custom` takes `charCode` and `fontName`; the other six are the
   * characters and faces Word's bullet library uses.
   */
  setLevelBullet(level: number, listBullet: ListBullet, charCode?: number, fontName?: string): void {
    const preset = BULLETS[listBullet];
    const code = listBullet === 'Custom' ? charCode : preset?.charCode;
    if (code === undefined) throw new Docx4JException("setLevelBullet('Custom') needs a charCode");
    const font = fontName ?? preset?.fontName;
    const lvl = this.writableLvl(level);
    lvl.numFmt = f.createNumFmt({ val: 'bullet' });
    linkParents(lvl.numFmt, lvl);
    lvl.lvlText = f.createLvlLvlText({ val: String.fromCharCode(code) });
    linkParents(lvl.lvlText, lvl);
    if (font !== undefined && font !== '') {
      const rPr = (lvl.rPr ??= { TYPE_NAME: 'org_docx4j_wml.RPr' });
      const rFonts = (rPr.rFonts ??= { TYPE_NAME: 'org_docx4j_wml.RFonts' });
      rFonts.ascii = font;
      rFonts.hAnsi = font;
      rFonts.hint = 'default';
      linkParents(rPr, lvl);
    }
    this.written();
  }

  /**
   * Office JS `list.setLevelIndents(level, textIndent, bulletNumberPictureIndent)`, both in
   * points: the level's `w:ind/@w:left` is where the text sits and the label's position is the
   * hanging indent from it (or `w:firstLine` where the label sits to the right of the text).
   */
  setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number): void {
    const lvl = this.writableLvl(level);
    const pPr = (lvl.pPr ??= { TYPE_NAME: 'org_docx4j_wml.PPr' });
    const ind = (pPr.ind ??= f.createPPrBaseInd({}));
    const left = Math.round(textIndent * TWIPS_PER_POINT);
    const label = Math.round(bulletNumberPictureIndent * TWIPS_PER_POINT);
    ind.left = left;
    if (label <= left) { ind.hanging = left - label; delete ind.firstLine; }
    else { ind.firstLine = label - left; delete ind.hanging; }
    linkParents(pPr, lvl);
    this.written();
  }

  /** Office JS `list.setLevelAlignment(level, alignment)`: the level's `w:lvlJc`. */
  setLevelAlignment(level: number, alignment: AlignmentOrUnknown): void {
    const lvl = this.writableLvl(level);
    const val = LVL_JC[alignment];
    if (val === undefined) delete lvl.lvlJc;
    else { lvl.lvlJc = f.createJc({ val }); linkParents(lvl.lvlJc, lvl); }
    this.written();
  }

  /**
   * Office JS `list.setLevelStartingNumber(level, startingNumber)`: the level's `w:start`. A
   * `w:startOverride` this `w:num` carries for the level would still win the first use, so it
   * is removed with it.
   */
  setLevelStartingNumber(level: number, startingNumber: number): void {
    const lvl = this.writableLvl(level);
    lvl.start = f.createLvlStart({ val: startingNumber });
    linkParents(lvl.start, lvl);
    for (const override of this.element.lvlOverride ?? []) {
      if (Number(override.ilvl) === level) delete override.startOverride;
    }
    this.written();
  }

  /**
   * A new `w:num` over this list's `w:abstractNum` with a `w:startOverride` on level 0, which
   * is how Word restarts a list: the counters are shared by abstract definition, and the
   * override is spent on the new `w:num`'s first use in a story (CR-001 section 15.2). Attach
   * the paragraphs that are to start again to the list it returns.
   *
   * Extension name (Office JS has no restart verb); `Paragraph.restartList()` is the sugar.
   */
  restart(): List {
    const numbering = this.live();
    const num = this.element;
    const start = this.level(0)?.start ?? 1;
    const created = f.createNumberingNum({
      numId: nextNumId(numbering),
      abstractNumId: f.createNumberingNumAbstractNumId({ val: num.abstractNumId.val }),
      lvlOverride: [f.createNumberingNumLvlOverride({
        ilvl: 0,
        startOverride: f.createNumberingNumLvlOverrideStartOverride({ val: start }),
      })],
    });
    (numbering.num ??= []).push(created);
    linkParents(created, numbering);
    this.written();
    return new List(String(created.numId), this.body);
  }

  /**
   * Gives this list a `w:abstractNum` of its own where another `w:num` shares the one it names
   * (CR-002 section 3.9), so that a change to its levels stays local to it. The numbering part
   * becomes live either way; the `setLevel*` methods call it for you.
   */
  separate(): void {
    const numbering = this.live();
    const num = this.element;
    const abstractId = num.abstractNumId.val;
    const shared = (numbering.num ?? []).some((other) => other !== num && other.abstractNumId?.val === abstractId);
    if (!shared) return;
    const source = (numbering.abstractNum ?? []).find((a) => a.abstractNumId === abstractId);
    if (source === undefined) return;
    const copy = deepCopy(source);
    copy.abstractNumId = nextAbstractNumId(numbering);
    copy.nsid = f.createCTLongHexNumber({ val: hexId() });
    // w:styleLink maps a numbering style to one definition; the copy must not claim it too
    delete copy.styleLink;
    (numbering.abstractNum ??= []).push(copy);
    linkParents(copy, numbering);
    num.abstractNumId = f.createNumberingNumAbstractNumId({ val: copy.abstractNumId });
    linkParents(num.abstractNumId, num);
    this.written();
  }

  // --- internals --------------------------------------------------------------------------

  /** The numbering part's live tree; the first write promotes the tree the definitions hold. */
  private live(): wml.Numbering {
    const part = numberingPartOf(this.body);
    if (part === undefined) throw new Docx4JException('This document has no numbering part');
    return part.makeLive();
  }

  /**
   * The `w:lvl` a write goes to: this list's own abstract definition's, created where the
   * definition has no level for it. A `w:lvlOverride/w:lvl` this `w:num` carries for the level
   * would mask the abstract one, so it goes (the `w:startOverride` beside it stays).
   */
  private writableLvl(level: number): wml.Lvl {
    this.separate();
    const abstract = this.abstractElement;
    if (abstract === undefined) {
      throw new ItemNotFound(`w:num ${this.numId} names no w:abstractNum that this document defines`);
    }
    for (const override of this.element.lvlOverride ?? []) {
      if (Number(override.ilvl) === level) delete override.lvl;
    }
    const levels = (abstract.lvl ??= []);
    let lvl = levels.find((l) => Number(l.ilvl) === level);
    if (lvl === undefined) {
      lvl = f.createLvl({ ilvl: level });
      levels.push(lvl);
      levels.sort((a, b) => Number(a.ilvl) - Number(b.ilvl));
      linkParents(lvl, abstract);
    }
    return lvl;
  }

  /** The numbering tree changed: the resolver (and with it the definitions) is rebuilt. */
  private written(): void {
    packageOf(this.body)?.refreshPropertyResolver();
  }
}

/**
 * A subset of Office JS `Word.ListItem` over one paragraph's place in its list. A view: the
 * paragraph is the state.
 */
export class ListItem {
  constructor(
    /** The paragraph this is the list item of (Office JS: `paragraph.listItem`). */
    readonly paragraph: Paragraph,
  ) {}

  /** Where this paragraph's numbering resolves to; throws when it is not a list item any more. */
  private get ref(): NumRef {
    const ref = numRefOf(this.paragraph);
    if (ref === undefined) throw new ItemNotFound('This paragraph is not a list item');
    return ref;
  }

  /** The list this item is in. */
  get list(): List {
    return new List(this.ref.numId!, this.paragraph.parentBody);
  }

  /**
   * Office JS `listItem.level`: the `w:ilvl` the paragraph counts at, a style-contributed level
   * included. Setting it writes the paragraph's own `w:numPr/w:ilvl` (with the `w:numId` beside
   * it where the list came from the style), through the accessor that records `w:pPrChange`
   * while the package tracks changes.
   */
  get level(): number {
    return Number(this.ref.ilvl ?? '0');
  }

  set level(value: number) {
    const ref = this.ref;
    const numPr = this.paragraph.numPr();
    if (numPr.numId?.val === undefined && ref.numId !== undefined) {
      numPr.numId = f.createPPrBaseNumPrNumId({ val: Number(ref.numId) });
      linkParents(numPr.numId, numPr);
    }
    numPr.ilvl = f.createPPrBaseNumPrIlvl({ val: value });
    linkParents(numPr.ilvl, numPr);
  }

  /**
   * Office JS `listItem.listString`: the label Word paints in front of the paragraph.
   *
   * It is counted by walking the paragraph's **story** from its start in document order
   * (CR-001 section 15.2), so it costs one walk of the story per read; `Body.listLabels()`
   * answers for every paragraph of the story in one walk.
   */
  get listString(): string {
    return labelOf(this.paragraph)?.listString ?? '';
  }

  /**
   * Office JS `listItem.siblingIndex`: the 0-based index of this item among the items of the
   * same list and level, counted from where the level last restarted. The same walk as
   * {@link listString}.
   */
  get siblingIndex(): number {
    return labelOf(this.paragraph)?.siblingIndex ?? 0;
  }

  /**
   * Office JS `listItem.getAncestor(parentOnly)`: the nearest preceding paragraph of this list
   * at a shallower level - the parent when `parentOnly` is true, which throws when the nearest
   * one is not exactly one level up. Throws `ItemNotFound` when there is no ancestor.
   */
  getAncestor(parentOnly = false): Paragraph {
    const level = this.level;
    const items = this.list.paragraphs;
    const at = items.findIndex((p) => p.p === this.paragraph.p);
    for (let i = at - 1; i >= 0; i--) {
      const other = levelOf(items[i]!);
      if (other === undefined || other >= level) continue;
      if (parentOnly && other !== level - 1) break;
      return items[i]!;
    }
    throw new ItemNotFound(`This list item (level ${level}) has no ${parentOnly ? 'parent' : 'ancestor'} in list ${this.list.id}`);
  }

  /**
   * Office JS `listItem.getDescendants(directChildrenOnly)`: the paragraphs of this list which
   * follow this one at a deeper level, up to the next one at this level or shallower; only the
   * ones one level down when `directChildrenOnly` is true.
   */
  getDescendants(directChildrenOnly = false): Paragraph[] {
    const level = this.level;
    const items = this.list.paragraphs;
    const at = items.findIndex((p) => p.p === this.paragraph.p);
    const out: Paragraph[] = [];
    for (let i = at + 1; i < items.length; i++) {
      const other = levelOf(items[i]!);
      if (other === undefined || other <= level) break;
      if (!directChildrenOnly || other === level + 1) out.push(items[i]!);
    }
    return out;
  }
}

// ------------------------------------------------------------- what Paragraph and Body call

/** The effective `w:ilvl` of a paragraph, or undefined when it is not a list item. */
export function levelOf(paragraph: Paragraph): number | undefined {
  const ref = numRefOf(paragraph);
  return ref === undefined ? undefined : Number(ref.ilvl ?? '0');
}

/** Office JS `paragraph.isListItem`: docx4j's `Emulator.numRefFor` answers it. */
export function isListItem(paragraph: Paragraph): boolean {
  return numRefOf(paragraph) !== undefined;
}

/** Office JS `paragraph.listItem`; throws `ItemNotFound` when the paragraph is not one. */
export function listItemOf(paragraph: Paragraph): ListItem {
  if (!isListItem(paragraph)) {
    throw new ItemNotFound('This paragraph is not a list item (listItemOrNullObject does not throw)');
  }
  return new ListItem(paragraph);
}

/** Office JS `paragraph.listItemOrNullObject`. */
export function listItemOrNullObjectOf(paragraph: Paragraph): ListItem {
  return isListItem(paragraph) ? new ListItem(paragraph) : nullObject<ListItem>('ListItem');
}

/** Office JS `paragraph.list`; throws `ItemNotFound` when the paragraph is not a list item. */
export function listOf(paragraph: Paragraph): List {
  const ref = numRefOf(paragraph);
  if (ref === undefined) {
    throw new ItemNotFound('This paragraph is not a list item (listOrNullObject does not throw)');
  }
  return new List(ref.numId!, paragraph.parentBody);
}

/** Office JS `paragraph.listOrNullObject`. */
export function listOrNullObjectOf(paragraph: Paragraph): List {
  const ref = numRefOf(paragraph);
  return ref === undefined ? nullObject<List>('List') : new List(ref.numId!, paragraph.parentBody);
}

/** Office JS `body.lists`: every `w:num` a paragraph of this body names, in order of first use. */
export function listsOf(body: Body): List[] {
  const emulator = numberingEmulatorOf(body);
  if (emulator === undefined) return [];
  const seen = new Set<string>();
  const out: List[] = [];
  for (const paragraph of body.paragraphs) {
    const ref = emulator.numRefFor(paragraph.p.pPr);
    if (ref.notNumbered || ref.numId === undefined || seen.has(ref.numId)) continue;
    seen.add(ref.numId);
    out.push(new List(ref.numId, body));
  }
  return out;
}

/** Office JS `paragraph.attachToList(listId, level)`: the paragraph's own `w:numPr`. */
export function attachToList(paragraph: Paragraph, listId: number | string, level = 0): List {
  const numPr = paragraph.numPr();
  numPr.numId = f.createPPrBaseNumPrNumId({ val: Number(listId) });
  numPr.ilvl = f.createPPrBaseNumPrIlvl({ val: level });
  linkParents(numPr, paragraph.p.pPr as object);
  return new List(String(listId), paragraph.parentBody);
}

/**
 * Office JS `paragraph.detachFromList()`: the paragraph's own `w:numPr` goes. Where its style
 * still numbers it, a `w:numPr` with `w:numId` 0 takes its place, which is how Word switches an
 * inherited list off (ECMA-376 17.9.18, and what `Emulator.resolve` reads).
 */
export function detachFromList(paragraph: Paragraph): void {
  if (paragraph.p.pPr?.numPr !== undefined) {
    delete paragraph.pPr().numPr;
  }
  if (!isListItem(paragraph)) return;
  const numPr = paragraph.numPr();
  numPr.numId = f.createPPrBaseNumPrNumId({ val: 0 });
  linkParents(numPr.numId, numPr);
  delete numPr.ilvl;
}

/**
 * Office JS `paragraph.startNewList()`: a `w:abstractNum` copied from docx4j's default
 * definitions (its `numbering.xml` resource: `w:abstractNum` 1 is Word's decimal set and 0 its
 * bullet set), a `w:num` for it, and the paragraph attached to that list.
 *
 * Asynchronous where Office JS's is synchronous: the default definitions are XML which has to
 * be unmarshalled, and a document with no numbering part gets one here (CR-002 section 17).
 */
export async function startNewList(paragraph: Paragraph, options?: StartListOptions): Promise<List> {
  const body = paragraph.parentBody;
  const part = await ensureNumberingPart(body);
  const numbering = await part.getContents();
  const defaults = await defaultNumbering();
  const template = (defaults.abstractNum ?? [])[options?.bullet === true ? 0 : 1];
  if (template === undefined) throw new Docx4JException('The default numbering definitions are missing an abstractNum');

  const abstract = deepCopy(template);
  abstract.abstractNumId = nextAbstractNumId(numbering);
  abstract.nsid = f.createCTLongHexNumber({ val: hexId() });
  delete abstract.styleLink;
  delete abstract.numStyleLink;
  (numbering.abstractNum ??= []).push(abstract);
  linkParents(abstract, numbering);

  const num = f.createNumberingNum({
    numId: nextNumId(numbering),
    abstractNumId: f.createNumberingNumAbstractNumId({ val: abstract.abstractNumId }),
  });
  (numbering.num ??= []).push(num);
  linkParents(num, numbering);

  packageOf(body)?.refreshPropertyResolver();
  return attachToList(paragraph, num.numId, options?.level ?? 0);
}

/**
 * The numbering part, created with docx4j's default numbering where the document has none
 * (docx4j's own `createNumbering` idiom), and unmarshalled either way so that what is written
 * to it is saved.
 */
async function ensureNumberingPart(body: Body): Promise<NumberingPartLike> {
  const pkg = packageOf(body);
  const main = mainOf(body);
  if (pkg === undefined || main === undefined) {
    throw new Docx4JException('This body has no package, so a list cannot be added to it');
  }
  const existing = main.numberingDefinitionsPart;
  if (existing !== undefined) {
    // the resolver read the part privately; unmarshalling it gives a second tree, so the
    // definitions (and the `getInd` the resolver folds into every layer) are rebuilt over the
    // live one
    const before = existing.isUnmarshalled;
    await existing.getContents();
    if (!before) pkg.refreshPropertyResolver();
    return existing;
  }
  // Imported here so that the content API does not import the parts statically: parts/wml
  // builds the Body views, so a static import would be a runtime cycle (as `ooxml.mts` does).
  const { NumberingDefinitionsPart } = await import('../../parts/wml/index.mjs');
  const created = new NumberingDefinitionsPart();
  main.addTargetPart(created);
  await created.unmarshalDefaultNumbering();
  // the resolver was built before the part existed, and refreshPropertyResolver() cannot find
  // a part added since (CR-001 section 15.2 departure 6)
  await pkg.refresh();
  return created as unknown as NumberingPartLike;
}

// ----------------------------------------------------------------------------- the mechanics

const TWIPS_PER_POINT = 20;

/** Office JS `Word.ListNumbering` as `w:numFmt`. */
const NUMBERING_FORMATS: Readonly<Record<ListNumbering, wml.NumberFormat>> = {
  None: 'none',
  Arabic: 'decimal',
  UpperRoman: 'upperRoman',
  LowerRoman: 'lowerRoman',
  UpperLetter: 'upperLetter',
  LowerLetter: 'lowerLetter',
};

/** Office JS `Word.ListBullet` as the character and face Word's bullet library uses. */
const BULLETS: Readonly<Record<ListBullet, { charCode: number; fontName: string } | undefined>> = {
  Custom: undefined,
  Solid: { charCode: 0xf0b7, fontName: 'Symbol' },            // a filled round bullet
  Hollow: { charCode: 0x006f, fontName: 'Courier New' },      // a lower-case o
  Square: { charCode: 0xf0a7, fontName: 'Wingdings' },        // a filled square
  Diamonds: { charCode: 0xf076, fontName: 'Wingdings' },      // a four-diamond bullet
  Arrow: { charCode: 0xf0d8, fontName: 'Wingdings' },         // a solid right-pointing arrowhead
  Checkmark: { charCode: 0xf0fc, fontName: 'Wingdings' },     // a tick
};

/** `w:lvlJc`, which takes the same values as a paragraph's `w:jc`. */
const LVL_JC: Readonly<Record<string, wml.JcEnumeration | undefined>> = {
  Left: 'left', Centered: 'center', Right: 'right', Justified: 'both', Unknown: undefined,
};

/** Office JS's `formatString` as a `w:lvlText`: a number is a 0-based level, a string a literal. */
function lvlTextOf(level: number, formatString?: (string | number)[]): string {
  if (formatString === undefined || formatString.length === 0) return `%${level + 1}.`;
  return formatString.map((part) => (typeof part === 'number' ? `%${part + 1}` : part)).join('');
}

/** A free `w:numId`: one past the highest the document uses (Word numbers from 1). */
function nextNumId(numbering: wml.Numbering): number {
  return (numbering.num ?? []).reduce((max, num) => Math.max(max, Number(num.numId)), 0) + 1;
}

/** A free `w:abstractNumId`. */
function nextAbstractNumId(numbering: wml.Numbering): number {
  return (numbering.abstractNum ?? []).reduce((max, a) => Math.max(max, Number(a.abstractNumId)), -1) + 1;
}

/** An eight-digit `w:nsid`, which is how Word tells two list families apart. */
function hexId(): string {
  return Math.floor(Math.random() * 0x7fffffff).toString(16).toUpperCase().padStart(8, '0');
}

/** docx4j's default numbering definitions, unmarshalled once. */
let defaults: Promise<wml.Numbering> | undefined;

function defaultNumbering(): Promise<wml.Numbering> {
  defaults ??= (async () => ((await unmarshalNode(parseXml(DEFAULT_NUMBERING_XML))).value as wml.Numbering))();
  return defaults;
}
