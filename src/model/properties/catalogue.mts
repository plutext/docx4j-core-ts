// The property catalogue (docx4j `org.docx4j.model.styles.PropertyCatalogue`, CR-015 phase 1)
// and the leaf merge rules it calls (the `apply`/`isEmpty` halves of `StyleUtil`).
//
// One table per properties element, one entry per schema member in schema order, each saying
// how to read it, write it, merge a more specific value over an inherited one, whether it is
// empty and whether it is formatting.  `apply`, `isEmpty`, `unset` and `hasDirectFormatting`
// are derived from the tables by iteration (`styleUtil.mts`), so the four cannot drift apart
// as the hand-kept lists docx4j replaced had.
//
// Every carried leaf is a *copy* (`copyOf`), never the style definition's own object: an
// effective `w:rPr`'s `w:b` must not be the styles part's `w:b`, or a caller editing the
// effective object would edit the document.  The objects facade's `deepCopy` is not used per
// leaf (it allocates a Map for cycle detection on every call); these are one- and
// two-attribute objects.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';

/** Run properties: `w:rPr` and the paragraph mark's `w:pPr/w:rPr`, which have the same members. */
export type RunProps = wml.RPr | wml.ParaRPr;

/** How a more specific value merges over an inherited one. Returns the value to store. */
export type Merge<V> = (source: V | undefined, destination: V | undefined) => V | undefined;

/**
 * One member of a properties element: its name, accessors, merge, emptiness and whether it
 * formats the content (as opposed to naming a style, recording a revision, or caching a
 * computed condition).
 */
export interface Property<O, V> {
  readonly name: string;
  /** False for a style reference, a revision record or Word's condition cache: `hasDirectFormatting` leaves those out. */
  readonly isFormatting: boolean;
  /**
   * Whether `isEmpty(owner)` counts this member. False for the few members docx4j's
   * hand-written `isEmpty` omits (`w:bidiVisual`, `w:tblCaption`, `w:tblDescription`), kept
   * so that the tables stay one-entry-per-member and the quirk stays visible.
   */
  readonly countsEmpty: boolean;
  get(owner: O): V | undefined;
  set(owner: O, value: V | undefined): void;
  merge(source: V | undefined, destination: V | undefined): V | undefined;
  isEmpty(value: V | undefined): boolean;
  /** A copy of a value of this member's type, sharing nothing with it. */
  copyOf(value: V): V;
}

// ---------------------------------------------------------------- copying

/**
 * A structural copy of a leaf value, without its `PARENT` link: what the merges store, so
 * that nothing in an effective `w:pPr` or `w:rPr` is a styles-part object.
 */
export function copyLeaf<T>(value: T): T {
  if (value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) return value.map((item) => copyLeaf(item)) as unknown as T;
  const out: Record<string, unknown> = {};
  for (const key of Object.keys(value as Record<string, unknown>)) {
    if (key === 'PARENT') continue;
    out[key] = copyLeaf((value as Record<string, unknown>)[key]);
  }
  return out as T;
}

// ---------------------------------------------------------------- emptiness of scalars

/** docx4j `StyleUtil.isEmpty(String)`: null or the empty string. */
function emptyString(value: string | undefined): boolean {
  return value === undefined || value === '';
}

/** docx4j `StyleUtil.isEmpty(STThemeColor)`: null or `none`. */
function emptyThemeColor(value: wml.STThemeColor | undefined): boolean {
  return value === undefined || value === 'none';
}

/** docx4j `StyleUtil.apply(String, String)` and friends: the source where it states one. */
function attr<V>(source: V | undefined, destination: V | undefined): V | undefined {
  return source === undefined ? destination : source;
}

/** docx4j `StyleUtil.apply(STThemeColor, STThemeColor)`: `none` does not override. */
function themeColor(source: wml.STThemeColor | undefined, destination: wml.STThemeColor | undefined): wml.STThemeColor | undefined {
  return emptyThemeColor(source) ? destination : source;
}

/** A merge with no per-attribute rule: a copy of the source where it has one, else the inherited. */
export function mergeReplace<V>(source: V | undefined, destination: V | undefined): V | undefined {
  return source === undefined ? destination : copyLeaf(source);
}

/** Never carried: docx4j's `apply` for the owning element does not mention this member. */
function notCarried<V>(_source: V | undefined, destination: V | undefined): V | undefined {
  return destination;
}

/** Never empty in itself: `undefined` only. */
function isNull(value: unknown): boolean {
  return value === undefined;
}

// ---------------------------------------------------------------- leaf emptiness

/** `<w:b/>` and `<w:b w:val="0"/>` both count: an explicit false is applied. */
export function isEmptyBool(value: wml.BooleanDefaultTrue | undefined): boolean {
  return value === undefined;
}

/** A `BooleanDefaultTrue`'s value: absent `w:val` means true (docx4j `isVal()`). */
export function isTrue(value: wml.BooleanDefaultTrue | undefined): boolean {
  return value !== undefined && value.val !== false;
}

function isEmptyVal(value: { val?: unknown } | undefined): boolean {
  return value === undefined || value.val === undefined;
}

function isEmptyStringVal(value: { val?: string } | undefined): boolean {
  return value === undefined || emptyString(value.val);
}

export function isEmptyRStyle(value: wml.RStyle | undefined): boolean {
  return isEmptyStringVal(value);
}

/** docx4j `isEmpty(RFonts)`: insensitive to a `w:hint` (which `apply` still carries). */
export function isEmptyRFonts(value: wml.RFonts | undefined): boolean {
  return value === undefined
    || (emptyString(value.ascii) && value.asciiTheme === undefined
      && emptyString(value.cs) && value.cstheme === undefined
      && emptyString(value.eastAsia) && value.eastAsiaTheme === undefined
      && emptyString(value.hAnsi) && value.hAnsiTheme === undefined);
}

export function isEmptyColor(value: wml.Color | undefined): boolean {
  return value === undefined || emptyString(value.val);
}

export function isEmptyU(value: wml.U | undefined): boolean {
  return value === undefined || (value.val === undefined && emptyString(value.color));
}

export function isEmptyHighlight(value: wml.Highlight | undefined): boolean {
  return value === undefined || emptyString(value.val);
}

/** @since CR-015 phase 1: a style stating only `w:lang` used to count as empty and was never applied. */
export function isEmptyLang(value: wml.CTLanguage | undefined): boolean {
  return value === undefined || (emptyString(value.val) && emptyString(value.eastAsia) && emptyString(value.bidi));
}

export function isEmptyShd(value: wml.CTShd | undefined): boolean {
  return value === undefined
    || (emptyString(value.color) && emptyString(value.themeTint) && emptyString(value.themeShade)
      && emptyString(value.fill) && emptyString(value.themeFillTint) && emptyString(value.themeFillShade)
      && value.themeColor === undefined && value.themeFill === undefined && value.val === undefined);
}

export function isEmptyBorder(value: wml.CTBorder | undefined): boolean {
  return value === undefined
    || (emptyString(value.color) && value.space === undefined && value.sz === undefined
      && value.themeColor === undefined && emptyString(value.themeShade) && emptyString(value.themeTint)
      && value.val === undefined);
}

export function isEmptyPBdr(value: wml.PPrBase.PBdr | undefined): boolean {
  return value === undefined
    || (isEmptyBorder(value.top) && isEmptyBorder(value.left) && isEmptyBorder(value.bottom)
      && isEmptyBorder(value.right) && isEmptyBorder(value.between) && isEmptyBorder(value.bar));
}

/**
 * docx4j `isEmpty(Ind)`: `w:start`/`w:end` (the strict-conformance spellings) count too
 * (17.0.5); dropping them lost them in style resolution.
 */
export function isEmptyInd(value: wml.PPrBase.Ind | undefined): boolean {
  return value === undefined
    || (value.firstLine === undefined && value.firstLineChars === undefined
      && value.hanging === undefined && value.hangingChars === undefined
      && value.left === undefined && value.leftChars === undefined
      && value.right === undefined && value.rightChars === undefined
      && value.start === undefined && value.startChars === undefined
      && value.end === undefined && value.endChars === undefined);
}

/** An explicit `w:beforeAutospacing`/`w:afterAutospacing` of either value is a statement too. */
export function isEmptySpacing(value: wml.PPrBase.Spacing | undefined): boolean {
  return value === undefined
    || (value.after === undefined && value.afterLines === undefined
      && value.before === undefined && value.beforeLines === undefined
      && value.line === undefined && value.lineRule === undefined
      && value.beforeAutospacing === undefined && value.afterAutospacing === undefined);
}

function isEmptyTabStop(value: wml.CTTabStop | undefined): boolean {
  return value === undefined || (value.pos === undefined && value.val === undefined && value.leader === undefined);
}

export function isEmptyTabs(value: wml.Tabs | undefined): boolean {
  if (value === undefined || value.tab === undefined || value.tab.length === 0) return true;
  return !value.tab.some((stop) => !isEmptyTabStop(stop));
}

/**
 * A `w:numPr` is empty only when it states neither `w:numId` nor `w:ilvl`: each inherits on
 * its own (CR-015 probe styles-numpr-ilvl-only).
 */
export function isEmptyNumPr(value: wml.PPrBase.NumPr | undefined): boolean {
  return value === undefined
    || ((value.numId === undefined || value.numId.val === undefined)
      && (value.ilvl === undefined || value.ilvl.val === undefined));
}

/** `w:anchorLock` defaults to true, so only an explicit `w:anchorLock="0"` makes a frame non-empty by itself. */
export function isEmptyFramePr(value: wml.CTFramePr | undefined): boolean {
  return value === undefined
    || (value.dropCap === undefined && value.lines === undefined && value.w === undefined
      && value.h === undefined && value.vSpace === undefined && value.hSpace === undefined
      && value.wrap === undefined && value.hAnchor === undefined && value.vAnchor === undefined
      && value.x === undefined && value.xAlign === undefined
      && value.y === undefined && value.yAlign === undefined && value.hRule === undefined
      && (value.anchorLock ?? true));
}

function isEmptyTblWidth(value: wml.TblWidth | undefined): boolean {
  // @w:type is ignored, as docx4j does
  return value === undefined || value.w === undefined;
}

function isEmptyTblCellMar(value: wml.CTTblCellMar | undefined): boolean {
  // w:start/w:end are not read, as docx4j's isEmpty/apply pair does not read them
  return value === undefined
    || (isEmptyTblWidth(value.bottom) && isEmptyTblWidth(value.left)
      && isEmptyTblWidth(value.right) && isEmptyTblWidth(value.top));
}

function isEmptyTblBorders(value: wml.TblBorders | undefined): boolean {
  return value === undefined
    || (isEmptyBorder(value.bottom) && isEmptyBorder(value.left) && isEmptyBorder(value.right)
      && isEmptyBorder(value.top) && isEmptyBorder(value.insideH) && isEmptyBorder(value.insideV));
}

function isEmptyTcBorders(value: wml.TcPrInner.TcBorders | undefined): boolean {
  return value === undefined
    || (isEmptyBorder(value.bottom) && isEmptyBorder(value.left) && isEmptyBorder(value.right)
      && isEmptyBorder(value.top) && isEmptyBorder(value.insideH) && isEmptyBorder(value.insideV)
      && isEmptyBorder(value.tl2Br) && isEmptyBorder(value.tr2Bl));
}

function isEmptyTcMar(value: wml.TcMar | undefined): boolean {
  return value === undefined
    || (isEmptyTblWidth(value.bottom) && isEmptyTblWidth(value.left)
      && isEmptyTblWidth(value.right) && isEmptyTblWidth(value.top));
}

function isEmptyTblLook(value: wml.CTTblLook | undefined): boolean {
  return value === undefined
    || (value.firstColumn === undefined && value.firstRow === undefined
      && value.lastColumn === undefined && value.lastRow === undefined
      && value.noHBand === undefined && value.noVBand === undefined && value.val === undefined);
}

function isEmptyTblPPr(value: wml.CTTblPPr | undefined): boolean {
  return value === undefined
    || (value.leftFromText === undefined && value.rightFromText === undefined
      && value.topFromText === undefined && value.bottomFromText === undefined
      && value.vertAnchor === undefined && value.horzAnchor === undefined
      && value.tblpXSpec === undefined && value.tblpX === undefined
      && value.tblpYSpec === undefined && value.tblpY === undefined);
}

function isEmptyCnf(value: wml.CTCnf | undefined): boolean {
  return isEmptyStringVal(value);
}

// ---------------------------------------------------------------- leaf merges

/** docx4j `apply(BooleanDefaultTrue, BooleanDefaultTrue)`: a copy of the source where it states one. */
export function mergeBool(source: wml.BooleanDefaultTrue | undefined, destination: wml.BooleanDefaultTrue | undefined): wml.BooleanDefaultTrue | undefined {
  return source === undefined ? destination : copyLeaf(source);
}

/** A leaf whose only content is `w:val`, taken whole when the source states it. */
function mergeVal<V extends { val?: unknown }>(isEmptyOf: (v: V | undefined) => boolean): Merge<V> {
  return (source, destination) => {
    if (source === undefined || isEmptyOf(source)) return destination;
    const out = (destination ?? {}) as V;
    (out as { val?: unknown }).val = (source as { val?: unknown }).val;
    return out;
  };
}

export const mergeRStyle: Merge<wml.RStyle> = mergeVal<wml.RStyle>(isEmptyRStyle);
export const mergeHighlight: Merge<wml.Highlight> = mergeVal<wml.Highlight>(isEmptyHighlight);
export const mergeEm: Merge<wml.CTEm> = mergeVal<wml.CTEm>(isEmptyVal);
export const mergeVertAlign: Merge<wml.CTVerticalAlignRun> = mergeVal<wml.CTVerticalAlignRun>(isEmptyVal);
export const mergeEffect: Merge<wml.CTTextEffect> = mergeVal<wml.CTTextEffect>(isEmptyVal);
export const mergeHps: Merge<wml.HpsMeasure> = mergeVal<wml.HpsMeasure>(isEmptyVal);
export const mergeSignedHps: Merge<wml.CTSignedHpsMeasure> = mergeVal<wml.CTSignedHpsMeasure>(isEmptyVal);
export const mergeSignedTwips: Merge<wml.CTSignedTwipsMeasure> = mergeVal<wml.CTSignedTwipsMeasure>(isEmptyVal);
export const mergeTextScale: Merge<wml.CTTextScale> = mergeVal<wml.CTTextScale>(isEmptyVal);
export const mergeJc: Merge<wml.Jc> = mergeVal<wml.Jc>(isEmptyVal);
export const mergeOutlineLvl: Merge<wml.PPrBase.OutlineLvl> = mergeVal<wml.PPrBase.OutlineLvl>(isEmptyVal);
export const mergeTextAlignment: Merge<wml.PPrBase.TextAlignment> = mergeVal<wml.PPrBase.TextAlignment>(isEmptyVal);
export const mergeTextDirection: Merge<wml.TextDirection> = mergeVal<wml.TextDirection>(isEmptyVal);
export const mergeTextboxTightWrap: Merge<wml.CTTextboxTightWrap> = mergeVal<wml.CTTextboxTightWrap>(isEmptyVal);
export const mergeCnf: Merge<wml.CTCnf> = mergeVal<wml.CTCnf>(isEmptyCnf);
export const mergePStyle: Merge<wml.PPrBase.PStyle> = mergeVal<wml.PPrBase.PStyle>(isEmptyStringVal);
const mergeTblStyle: Merge<wml.CTTblPrBase.TblStyle> = mergeVal<wml.CTTblPrBase.TblStyle>(isEmptyStringVal);
const mergeRowBandSize: Merge<wml.CTTblPrBase.TblStyleRowBandSize> = mergeVal<wml.CTTblPrBase.TblStyleRowBandSize>(isEmptyVal);
const mergeColBandSize: Merge<wml.CTTblPrBase.TblStyleColBandSize> = mergeVal<wml.CTTblPrBase.TblStyleColBandSize>(isEmptyVal);
const mergeTblOverlap: Merge<wml.CTTblOverlap> = mergeVal<wml.CTTblOverlap>(isEmptyVal);
const mergeVAlign: Merge<wml.CTVerticalJc> = mergeVal<wml.CTVerticalJc>(isEmptyVal);
const mergeGridSpan: Merge<wml.TcPrInner.GridSpan> = mergeVal<wml.TcPrInner.GridSpan>(isEmptyVal);
const mergeHMerge: Merge<wml.TcPrInner.HMerge> = mergeVal<wml.TcPrInner.HMerge>(isEmptyStringVal);
const mergeVMerge: Merge<wml.TcPrInner.VMerge> = mergeVal<wml.TcPrInner.VMerge>(isEmptyStringVal);
const mergeTblLayout: Merge<wml.CTTblLayoutType> = (source, destination) => {
  if (source === undefined || source.type === undefined) return destination;
  const out = destination ?? {};
  out.type = source.type;
  return out;
};

/**
 * `w:rFonts`: a theme reference trumps an explicit face within one element, and a source with
 * only a `w:hint` still says something. A null source leaves the destination alone (until
 * CR-015 phase 1 it was given an empty `w:rFonts`, so every merged `w:rPr` read as directly
 * formatted).
 */
export const mergeRFonts: Merge<wml.RFonts> = (source, destination) => {
  if (source === undefined) return destination;
  const out: wml.RFonts = destination ?? {};
  if (isEmptyRFonts(source)) {
    if (source.hint !== undefined) out.hint = source.hint;
    // else: keep the (possibly empty) element, which RunFontSelector relies on
    return out;
  }
  if (source.ascii !== undefined) { out.ascii = source.ascii; delete out.asciiTheme; }
  if (source.cs !== undefined) { out.cs = source.cs; delete out.cstheme; }
  if (source.eastAsia !== undefined) { out.eastAsia = source.eastAsia; delete out.eastAsiaTheme; }
  if (source.hAnsi !== undefined) { out.hAnsi = source.hAnsi; delete out.hAnsiTheme; }
  if (source.asciiTheme !== undefined) { out.asciiTheme = source.asciiTheme; delete out.ascii; }
  if (source.cstheme !== undefined) { out.cstheme = source.cstheme; delete out.cs; }
  if (source.eastAsiaTheme !== undefined) { out.eastAsiaTheme = source.eastAsiaTheme; delete out.eastAsia; }
  if (source.hAnsiTheme !== undefined) { out.hAnsiTheme = source.hAnsiTheme; delete out.hAnsi; }
  if (source.hint !== undefined) out.hint = source.hint;
  return out;
};

export const mergeColor: Merge<wml.Color> = (source, destination) => {
  if (isEmptyColor(source)) return destination;
  const s = source as wml.Color;
  const out = (destination ?? {}) as wml.Color;
  out.themeColor = themeColor(s.themeColor, out.themeColor);
  out.themeShade = attr(s.themeShade, out.themeShade);
  out.themeTint = attr(s.themeTint, out.themeTint);
  out.val = attr(s.val, out.val) as string;
  return out;
};

/** `w:val` and the colour attributes each inherit on their own (CR-015 phase 1). */
export const mergeU: Merge<wml.U> = (source, destination) => {
  if (isEmptyU(source)) return destination;
  const s = source as wml.U;
  const out = destination ?? {};
  out.val = attr(s.val, out.val);
  out.color = attr(s.color, out.color);
  out.themeColor = themeColor(s.themeColor, out.themeColor);
  out.themeTint = attr(s.themeTint, out.themeTint);
  out.themeShade = attr(s.themeShade, out.themeShade);
  return out;
};

/** `w:val`, `w:eastAsia` and `w:bidi` each name the language of one script range and each inherits. */
export const mergeLang: Merge<wml.CTLanguage> = (source, destination) => {
  if (isEmptyLang(source)) return destination;
  const s = source as wml.CTLanguage;
  const out = destination ?? {};
  out.val = attr(s.val, out.val);
  out.eastAsia = attr(s.eastAsia, out.eastAsia);
  out.bidi = attr(s.bidi, out.bidi);
  return out;
};

export const mergeShd: Merge<wml.CTShd> = (source, destination) => {
  if (isEmptyShd(source)) return destination;
  const s = source as wml.CTShd;
  const out = (destination ?? {}) as wml.CTShd;
  out.color = attr(s.color, out.color);
  out.fill = attr(s.fill, out.fill);
  out.val = attr(s.val, out.val) as wml.STShd;
  out.themeTint = attr(s.themeTint, out.themeTint);
  out.themeShade = attr(s.themeShade, out.themeShade);
  out.themeFillTint = attr(s.themeFillTint, out.themeFillTint);
  out.themeFillShade = attr(s.themeFillShade, out.themeFillShade);
  // the enums only when the source states them (until CR-015 a null source cleared them)
  out.themeColor = themeColor(s.themeColor, out.themeColor);
  out.themeFill = themeColor(s.themeFill, out.themeFill);
  return out;
};

export const mergeBorder: Merge<wml.CTBorder> = (source, destination) => {
  if (isEmptyBorder(source)) return destination;
  const s = source as wml.CTBorder;
  const out = (destination ?? {}) as wml.CTBorder;
  out.color = attr(s.color, out.color);
  out.space = attr(s.space, out.space);
  out.sz = attr(s.sz, out.sz);
  out.themeColor = themeColor(s.themeColor, out.themeColor);
  out.themeShade = attr(s.themeShade, out.themeShade);
  out.themeTint = attr(s.themeTint, out.themeTint);
  out.val = attr(s.val, out.val) as wml.STBorder;
  return out;
};

export const mergePBdr: Merge<wml.PPrBase.PBdr> = (source, destination) => {
  if (isEmptyPBdr(source)) return destination;
  const s = source as wml.PPrBase.PBdr;
  const out = destination ?? {};
  out.top = mergeBorder(s.top, out.top);
  out.left = mergeBorder(s.left, out.left);
  out.bottom = mergeBorder(s.bottom, out.bottom);
  out.right = mergeBorder(s.right, out.right);
  out.between = mergeBorder(s.between, out.between);
  out.bar = mergeBorder(s.bar, out.bar);
  return out;
};

/** `w:numId` and `w:ilvl` each inherit on their own; see {@link isEmptyNumPr}. */
export const mergeNumPr: Merge<wml.PPrBase.NumPr> = (source, destination) => {
  if (isEmptyNumPr(source)) return destination;
  const s = source as wml.PPrBase.NumPr;
  const out = destination ?? {};
  if (s.numId !== undefined && s.numId.val !== undefined) out.numId = copyLeaf(s.numId);
  if (s.ilvl !== undefined && s.ilvl.val !== undefined) out.ilvl = copyLeaf(s.ilvl);
  return out;
};

/**
 * `w:ind`: `w:firstLine` and `w:hanging` are two spellings of one property (ECMA-376
 * 17.3.1.12), so whichever of the two the more specific `w:ind` states replaces both; the
 * rest inherit per attribute.
 */
export const mergeInd: Merge<wml.PPrBase.Ind> = (source, destination) => {
  if (isEmptyInd(source)) return destination;
  const s = source as wml.PPrBase.Ind;
  const out = destination ?? {};
  const sourceFirstLine = s.firstLine !== undefined || s.firstLineChars !== undefined;
  const sourceHanging = s.hanging !== undefined || s.hangingChars !== undefined;
  if (sourceFirstLine && !sourceHanging) { delete out.hanging; delete out.hangingChars; }
  else if (sourceHanging && !sourceFirstLine) { delete out.firstLine; delete out.firstLineChars; }
  out.firstLine = attr(s.firstLine, out.firstLine);
  out.firstLineChars = attr(s.firstLineChars, out.firstLineChars);
  out.hanging = attr(s.hanging, out.hanging);
  out.hangingChars = attr(s.hangingChars, out.hangingChars);
  out.left = attr(s.left, out.left);
  out.leftChars = attr(s.leftChars, out.leftChars);
  out.right = attr(s.right, out.right);
  out.rightChars = attr(s.rightChars, out.rightChars);
  out.start = attr(s.start, out.start);
  out.startChars = attr(s.startChars, out.startChars);
  out.end = attr(s.end, out.end);
  out.endChars = attr(s.endChars, out.endChars);
  return out;
};

/**
 * `w:spacing`: the autospacing attributes survive the merge, and `w:lineRule` qualifies
 * `w:line` - a source stating `w:line` states a rule too (auto when it names none), and a
 * source stating no `w:line` leaves both inherited (CR-015 probe styles-linerule).
 */
export const mergeSpacing: Merge<wml.PPrBase.Spacing> = (source, destination) => {
  if (source === undefined) return destination;
  let out = destination;
  if (source.beforeAutospacing !== undefined || source.afterAutospacing !== undefined) {
    out ??= {};
    if (source.beforeAutospacing !== undefined) out.beforeAutospacing = source.beforeAutospacing;
    if (source.afterAutospacing !== undefined) out.afterAutospacing = source.afterAutospacing;
  }
  if (!isEmptySpacing(source)) {
    out ??= {};
    out.after = attr(source.after, out.after);
    out.afterLines = attr(source.afterLines, out.afterLines);
    out.before = attr(source.before, out.before);
    out.beforeLines = attr(source.beforeLines, out.beforeLines);
    if (source.line !== undefined) {
      out.line = source.line;
      out.lineRule = source.lineRule ?? 'auto';
    } else if (source.lineRule !== undefined) {
      out.lineRule = source.lineRule;
    }
  }
  return out;
};

/**
 * `w:tabs`: the custom tab stops are the union of those declared and those inherited, a stop
 * at the same position replaces the inherited one, and `w:val="clear"` only removes
 * (ECMA-376 17.3.1.38, 17.3.1.37). Word holds them in ascending position order.
 */
export const mergeTabs: Merge<wml.Tabs> = (source, destination) => {
  if (isEmptyTabs(source)) return destination;
  const out: wml.Tabs = destination ?? { tab: [] };
  out.tab ??= [];
  for (const stop of (source as wml.Tabs).tab) {
    const pos = stop.pos;
    if (pos !== undefined) {
      for (let i = out.tab.length - 1; i >= 0; i--) {
        if (out.tab[i]!.pos === pos) out.tab.splice(i, 1);
      }
    }
    if (stop.val === 'clear') continue;
    out.tab.push({ leader: stop.leader, pos, val: stop.val } as wml.CTTabStop);
  }
  out.tab.sort((a, b) => {
    if (a.pos === undefined) return b.pos === undefined ? 0 : -1;
    if (b.pos === undefined) return 1;
    return a.pos - b.pos;
  });
  return out;
};

/**
 * `w:framePr`'s attributes each inherit on their own: a paragraph whose direct `w:framePr`
 * gives only `w:x`/`w:y` keeps the anchors its style states. `w:anchorLock` defaults to true,
 * and docx4j's merge writes that default out explicitly.
 */
export const mergeFramePr: Merge<wml.CTFramePr> = (source, destination) => {
  if (isEmptyFramePr(source)) return destination;
  const s = source as wml.CTFramePr;
  const out = destination ?? {};
  out.dropCap = attr(s.dropCap, out.dropCap);
  out.lines = attr(s.lines, out.lines);
  out.w = attr(s.w, out.w);
  out.h = attr(s.h, out.h);
  out.vSpace = attr(s.vSpace, out.vSpace);
  out.hSpace = attr(s.hSpace, out.hSpace);
  out.wrap = attr(s.wrap, out.wrap);
  out.hAnchor = attr(s.hAnchor, out.hAnchor);
  out.vAnchor = attr(s.vAnchor, out.vAnchor);
  out.x = attr(s.x, out.x);
  out.xAlign = attr(s.xAlign, out.xAlign);
  out.y = attr(s.y, out.y);
  out.yAlign = attr(s.yAlign, out.yAlign);
  out.hRule = attr(s.hRule, out.hRule);
  // docx4j reads this through isAnchorLock(), which answers true when the attribute is absent
  out.anchorLock = s.anchorLock ?? true;
  return out;
};

const mergeTblWidth: Merge<wml.TblWidth> = (source, destination) => {
  if (source === undefined || source.w === undefined) return destination;
  const out = destination ?? {};
  out.w = source.w;
  // "if @w:type is omitted, its value shall be assumed to be dxa"
  out.type = source.type;
  return out;
};

const mergeTblCellMar: Merge<wml.CTTblCellMar> = (source, destination) => {
  if (isEmptyTblCellMar(source)) return destination;
  const s = source as wml.CTTblCellMar;
  const out = destination ?? {};
  out.bottom = mergeTblWidth(s.bottom, out.bottom);
  out.left = mergeTblWidth(s.left, out.left);
  out.right = mergeTblWidth(s.right, out.right);
  out.top = mergeTblWidth(s.top, out.top);
  return out;
};

const mergeTblBorders: Merge<wml.TblBorders> = (source, destination) => {
  if (isEmptyTblBorders(source)) return destination;
  const s = source as wml.TblBorders;
  const out = destination ?? {};
  out.bottom = mergeBorder(s.bottom, out.bottom);
  out.left = mergeBorder(s.left, out.left);
  out.right = mergeBorder(s.right, out.right);
  out.top = mergeBorder(s.top, out.top);
  out.insideH = mergeBorder(s.insideH, out.insideH);
  out.insideV = mergeBorder(s.insideV, out.insideV);
  return out;
};

const mergeTcBorders: Merge<wml.TcPrInner.TcBorders> = (source, destination) => {
  if (isEmptyTcBorders(source)) return destination;
  const s = source as wml.TcPrInner.TcBorders;
  const out = destination ?? {};
  out.bottom = mergeBorder(s.bottom, out.bottom);
  out.left = mergeBorder(s.left, out.left);
  out.right = mergeBorder(s.right, out.right);
  out.top = mergeBorder(s.top, out.top);
  out.insideH = mergeBorder(s.insideH, out.insideH);
  out.insideV = mergeBorder(s.insideV, out.insideV);
  out.tl2Br = mergeBorder(s.tl2Br, out.tl2Br);
  out.tr2Bl = mergeBorder(s.tr2Bl, out.tr2Bl);
  return out;
};

const mergeTcMar: Merge<wml.TcMar> = (source, destination) => {
  if (isEmptyTcMar(source)) return destination;
  const s = source as wml.TcMar;
  const out = destination ?? {};
  out.bottom = mergeTblWidth(s.bottom, out.bottom);
  out.left = mergeTblWidth(s.left, out.left);
  out.right = mergeTblWidth(s.right, out.right);
  out.top = mergeTblWidth(s.top, out.top);
  return out;
};

const mergeTblLook: Merge<wml.CTTblLook> = (source, destination) => {
  if (isEmptyTblLook(source)) return destination;
  const s = source as wml.CTTblLook;
  const out = destination ?? {};
  out.firstColumn = attr(s.firstColumn, out.firstColumn);
  out.firstRow = attr(s.firstRow, out.firstRow);
  out.lastColumn = attr(s.lastColumn, out.lastColumn);
  out.lastRow = attr(s.lastRow, out.lastRow);
  out.noHBand = attr(s.noHBand, out.noHBand);
  out.noVBand = attr(s.noVBand, out.noVBand);
  out.val = attr(s.val, out.val);
  return out;
};

/**
 * `w:tblpPr`: the anchors and the `*Spec` enums only when the source states them - a
 * `w:tblpPr` stating only `w:tblpX` keeps its inherited anchors in Word.
 */
const mergeTblPPr: Merge<wml.CTTblPPr> = (source, destination) => {
  if (isEmptyTblPPr(source)) return destination;
  const s = source as wml.CTTblPPr;
  const out = destination ?? {};
  out.leftFromText = attr(s.leftFromText, out.leftFromText);
  out.rightFromText = attr(s.rightFromText, out.rightFromText);
  out.topFromText = attr(s.topFromText, out.topFromText);
  out.bottomFromText = attr(s.bottomFromText, out.bottomFromText);
  if (s.vertAnchor !== undefined) out.vertAnchor = s.vertAnchor;
  if (s.horzAnchor !== undefined) out.horzAnchor = s.horzAnchor;
  if (s.tblpXSpec !== undefined) out.tblpXSpec = s.tblpXSpec;
  out.tblpX = attr(s.tblpX, out.tblpX);
  if (s.tblpYSpec !== undefined) out.tblpYSpec = s.tblpYSpec;
  out.tblpY = attr(s.tblpY, out.tblpY);
  return out;
};

// ---------------------------------------------------------------- the tables

interface Options {
  formatting?: boolean;
  countsEmpty?: boolean;
}

/**
 * One entry. The accessors are by member name, which is the same on `RPr` and `ParaRPr`
 * (docx4j needs an accessor pair per class there; here the two generated interfaces have the
 * same members, so one pair serves both).
 */
function prop<O, V>(name: string, merge: Merge<V>, isEmpty: (value: V | undefined) => boolean,
  copyOf: (value: V) => V, options?: Options): Property<O, V> {
  return {
    name,
    isFormatting: options?.formatting ?? true,
    countsEmpty: options?.countsEmpty ?? true,
    get: (owner) => (owner as Record<string, unknown>)[name] as V | undefined,
    set: (owner, value) => {
      const record = owner as Record<string, unknown>;
      if (value === undefined) delete record[name]; else record[name] = value;
    },
    merge,
    isEmpty,
    copyOf,
  };
}

/** A `BooleanDefaultTrue` member. */
function bool<O>(name: string, options?: Options): Property<O, wml.BooleanDefaultTrue> {
  return prop<O, wml.BooleanDefaultTrue>(name, mergeBool, isEmptyBool, copyLeaf, options);
}

/** A w14 member (Office 2010 text effects and OpenType features): replaced whole. */
function w14<O>(name: string): Property<O, object> {
  return prop<O, object>(name, mergeReplace, isNull, copyLeaf);
}

/**
 * The members of `w:rPr` (17.3.2) and of the paragraph mark's `w:pPr/w:rPr`, in schema order,
 * plus the w14 extensions the generated interfaces carry. Not listed: `w:rPrChange` and the
 * paragraph mark's `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`, which record revisions and are
 * never inherited.
 */
export const RUN: readonly Property<RunProps, never>[] = ([
  prop<RunProps, wml.RStyle>('rStyle', mergeRStyle, isEmptyRStyle, copyLeaf, { formatting: false }),
  prop<RunProps, wml.RFonts>('rFonts', mergeRFonts, isEmptyRFonts, copyLeaf),
  bool<RunProps>('b'),
  bool<RunProps>('bCs'),
  bool<RunProps>('i'),
  bool<RunProps>('iCs'),
  bool<RunProps>('caps'),
  bool<RunProps>('smallCaps'),
  bool<RunProps>('strike'),
  bool<RunProps>('dstrike'),
  bool<RunProps>('outline'),
  bool<RunProps>('shadow'),
  bool<RunProps>('emboss'),
  bool<RunProps>('imprint'),
  bool<RunProps>('noProof'),
  bool<RunProps>('snapToGrid'),
  bool<RunProps>('vanish'),
  bool<RunProps>('webHidden'),
  prop<RunProps, wml.Color>('color', mergeColor, isEmptyColor, copyLeaf),
  prop<RunProps, wml.CTSignedTwipsMeasure>('spacing', mergeSignedTwips, isEmptyVal, copyLeaf),
  prop<RunProps, wml.CTTextScale>('w', mergeTextScale, isEmptyVal, copyLeaf),
  prop<RunProps, wml.HpsMeasure>('kern', mergeHps, isEmptyVal, copyLeaf),
  prop<RunProps, wml.CTSignedHpsMeasure>('position', mergeSignedHps, isEmptyVal, copyLeaf),
  prop<RunProps, wml.HpsMeasure>('sz', mergeHps, isEmptyVal, copyLeaf),
  prop<RunProps, wml.HpsMeasure>('szCs', mergeHps, isEmptyVal, copyLeaf),
  prop<RunProps, wml.Highlight>('highlight', mergeHighlight, isEmptyHighlight, copyLeaf),
  prop<RunProps, wml.U>('u', mergeU, isEmptyU, copyLeaf),
  prop<RunProps, wml.CTTextEffect>('effect', mergeEffect, isEmptyVal, copyLeaf),
  prop<RunProps, wml.CTBorder>('bdr', mergeBorder, isEmptyBorder, copyLeaf),
  prop<RunProps, wml.CTShd>('shd', mergeShd, isEmptyShd, copyLeaf),
  prop<RunProps, wml.CTFitText>('fitText', mergeReplace, isNull, copyLeaf),
  prop<RunProps, wml.CTVerticalAlignRun>('vertAlign', mergeVertAlign, isEmptyVal, copyLeaf),
  bool<RunProps>('rtl'),
  bool<RunProps>('cs'),
  prop<RunProps, wml.CTEm>('em', mergeEm, isEmptyVal, copyLeaf),
  prop<RunProps, wml.CTLanguage>('lang', mergeLang, isEmptyLang, copyLeaf),
  prop<RunProps, wml.CTEastAsianLayout>('eastAsianLayout', mergeReplace, isNull, copyLeaf),
  bool<RunProps>('specVanish'),
  bool<RunProps>('oMath'),
  // w14 (Office 2010 text effects and OpenType features)
  w14<RunProps>('glow'),
  w14<RunProps>('shadow14'),
  w14<RunProps>('reflection'),
  w14<RunProps>('textOutline'),
  w14<RunProps>('textFill'),
  w14<RunProps>('scene3D'),
  w14<RunProps>('props3D'),
  w14<RunProps>('ligatures'),
  w14<RunProps>('numForm'),
  w14<RunProps>('numSpacing'),
  w14<RunProps>('stylisticSets'),
  w14<RunProps>('cntxtAlts'),
] as unknown[] as Property<RunProps, never>[]);

/** Members of `w:rPr` and `w:pPr/w:rPr` that are not in {@link RUN} (revision records). */
export const RUN_EXCLUDED: readonly string[] = ['rPrChange', 'ins', 'del', 'moveFrom', 'moveTo'];

/**
 * The names of the **toggle properties** of `w:rPr`, in the order ECMA-376-1 §17.7.3 lists
 * them. Twelve, and no more: `w:dstrike`, `w:noProof`, `w:snapToGrid`, `w:webHidden`,
 * `w:rtl`, `w:cs`, `w:specVanish` and `w:oMath` are Boolean but are not toggles.
 */
export const TOGGLE_NAMES: ReadonlySet<string> = new Set([
  'b', 'bCs', 'caps', 'emboss', 'i', 'iCs', 'imprint', 'outline',
  'shadow', 'smallCaps', 'strike', 'vanish',
]);

/** {@link TOGGLE_NAMES} as members of {@link RUN}. */
export const TOGGLES: readonly Property<RunProps, wml.BooleanDefaultTrue>[] = [...TOGGLE_NAMES].map((name) => {
  const found = RUN.find((p) => p.name === name);
  if (!found) throw new Error(`no w:rPr member named ${name}`);
  return found as unknown as Property<RunProps, wml.BooleanDefaultTrue>;
});

/** The `w:ind` member's name: the paragraph merge applies it last, after the numbering level's. */
export const IND = 'ind';

/**
 * The members of `w:pPr` shared with a style's `w:pPr` (`PPrBase`, 17.3.1), in schema order.
 * Not formatting: `w:pStyle` (the style reference), `w:divId` (an HTML import artefact),
 * `w:cnfStyle` (Word's cache of the table conditions that apply) and `w14:collapsed` (the
 * outline view's state). A `w:pPr`'s own `w:rPr` (the paragraph mark) and `w:sectPr` are
 * outside `PPrBase` and outside this list.
 */
export const PARAGRAPH: readonly Property<wml.PPrBase, never>[] = ([
  prop<wml.PPrBase, wml.PPrBase.PStyle>('pStyle', mergePStyle, isEmptyStringVal, copyLeaf, { formatting: false }),
  bool<wml.PPrBase>('keepNext'),
  bool<wml.PPrBase>('keepLines'),
  bool<wml.PPrBase>('pageBreakBefore'),
  prop<wml.PPrBase, wml.CTFramePr>('framePr', mergeFramePr, isEmptyFramePr, copyLeaf),
  bool<wml.PPrBase>('widowControl'),
  prop<wml.PPrBase, wml.PPrBase.NumPr>('numPr', mergeNumPr, isEmptyNumPr, copyLeaf),
  bool<wml.PPrBase>('suppressLineNumbers'),
  prop<wml.PPrBase, wml.PPrBase.PBdr>('pBdr', mergePBdr, isEmptyPBdr, copyLeaf),
  prop<wml.PPrBase, wml.CTShd>('shd', mergeShd, isEmptyShd, copyLeaf),
  prop<wml.PPrBase, wml.Tabs>('tabs', mergeTabs, isEmptyTabs, copyLeaf),
  bool<wml.PPrBase>('suppressAutoHyphens'),
  bool<wml.PPrBase>('kinsoku'),
  bool<wml.PPrBase>('wordWrap'),
  bool<wml.PPrBase>('overflowPunct'),
  bool<wml.PPrBase>('topLinePunct'),
  bool<wml.PPrBase>('autoSpaceDE'),
  bool<wml.PPrBase>('autoSpaceDN'),
  bool<wml.PPrBase>('bidi'),
  bool<wml.PPrBase>('adjustRightInd'),
  bool<wml.PPrBase>('snapToGrid'),
  prop<wml.PPrBase, wml.PPrBase.Spacing>('spacing', mergeSpacing, isEmptySpacing, copyLeaf),
  prop<wml.PPrBase, wml.PPrBase.Ind>(IND, mergeInd, isEmptyInd, copyLeaf),
  bool<wml.PPrBase>('contextualSpacing'),
  bool<wml.PPrBase>('mirrorIndents'),
  bool<wml.PPrBase>('suppressOverlap'),
  prop<wml.PPrBase, wml.Jc>('jc', mergeJc, isEmptyVal, copyLeaf),
  prop<wml.PPrBase, wml.TextDirection>('textDirection', mergeTextDirection, isEmptyVal, copyLeaf),
  prop<wml.PPrBase, wml.PPrBase.TextAlignment>('textAlignment', mergeTextAlignment, isEmptyVal, copyLeaf),
  prop<wml.PPrBase, wml.CTTextboxTightWrap>('textboxTightWrap', mergeTextboxTightWrap, isEmptyVal, copyLeaf),
  prop<wml.PPrBase, wml.PPrBase.OutlineLvl>('outlineLvl', mergeOutlineLvl, isEmptyVal, copyLeaf),
  prop<wml.PPrBase, wml.PPrBase.DivId>('divId', mergeReplace, isNull, copyLeaf, { formatting: false }),
  prop<wml.PPrBase, wml.CTCnf>('cnfStyle', mergeCnf, isEmptyCnf, copyLeaf, { formatting: false }),
  bool<wml.PPrBase>('collapsed', { formatting: false }),
] as unknown[] as Property<wml.PPrBase, never>[]);

/** Members of `w:pPr` that are not in {@link PARAGRAPH}: the mark, the section, the revision. */
export const PARAGRAPH_EXCLUDED: readonly string[] = ['rPr', 'sectPr', 'pPrChange'];

/**
 * The members of `w:tblPr` (`CTTblPrBase`, 17.4.60), in schema order.
 *
 * Three quirks of docx4j's hand-written pair are reproduced and flagged: `w:bidiVisual` is
 * applied but not counted by `isEmpty`, and `w:tblCaption` and `w:tblDescription` are neither
 * applied nor counted (they are a table's accessibility text, not formatting a style carries).
 */
export const TABLE: readonly Property<wml.CTTblPrBase, never>[] = ([
  prop<wml.CTTblPrBase, wml.CTTblPrBase.TblStyle>('tblStyle', mergeTblStyle, isEmptyStringVal, copyLeaf, { formatting: false }),
  prop<wml.CTTblPrBase, wml.CTTblPPr>('tblpPr', mergeTblPPr, isEmptyTblPPr, copyLeaf),
  prop<wml.CTTblPrBase, wml.CTTblOverlap>('tblOverlap', mergeTblOverlap, isEmptyVal, copyLeaf),
  bool<wml.CTTblPrBase>('bidiVisual', { countsEmpty: false }),
  prop<wml.CTTblPrBase, wml.CTTblPrBase.TblStyleRowBandSize>('tblStyleRowBandSize', mergeRowBandSize, isEmptyVal, copyLeaf),
  prop<wml.CTTblPrBase, wml.CTTblPrBase.TblStyleColBandSize>('tblStyleColBandSize', mergeColBandSize, isEmptyVal, copyLeaf),
  prop<wml.CTTblPrBase, wml.TblWidth>('tblW', mergeTblWidth, isEmptyTblWidth, copyLeaf),
  prop<wml.CTTblPrBase, wml.Jc>('jc', mergeJc, isEmptyVal, copyLeaf),
  prop<wml.CTTblPrBase, wml.TblWidth>('tblCellSpacing', mergeTblWidth, isEmptyTblWidth, copyLeaf),
  prop<wml.CTTblPrBase, wml.TblWidth>('tblInd', mergeTblWidth, isEmptyTblWidth, copyLeaf),
  prop<wml.CTTblPrBase, wml.TblBorders>('tblBorders', mergeTblBorders, isEmptyTblBorders, copyLeaf),
  prop<wml.CTTblPrBase, wml.CTShd>('shd', mergeShd, isEmptyShd, copyLeaf),
  prop<wml.CTTblPrBase, wml.CTTblLayoutType>('tblLayout', mergeTblLayout, (v) => v === undefined || v.type === undefined, copyLeaf),
  prop<wml.CTTblPrBase, wml.CTTblCellMar>('tblCellMar', mergeTblCellMar, isEmptyTblCellMar, copyLeaf),
  prop<wml.CTTblPrBase, wml.CTTblLook>('tblLook', mergeTblLook, isEmptyTblLook, copyLeaf),
  prop<wml.CTTblPrBase, wml.CTString>('tblCaption', notCarried, isNull, copyLeaf, { countsEmpty: false }),
  prop<wml.CTTblPrBase, wml.CTString>('tblDescription', notCarried, isNull, copyLeaf, { countsEmpty: false }),
] as unknown[] as Property<wml.CTTblPrBase, never>[]);

export const TABLE_EXCLUDED: readonly string[] = ['tblPrChange'];

/**
 * The members of `w:tcPr` (`TcPr`, 17.4.68), in schema order. `w:cellIns`, `w:cellDel`,
 * `w:cellMerge` and `w:tcPrChange` record revisions and are excluded.
 */
export const CELL: readonly Property<wml.TcPr, never>[] = ([
  prop<wml.TcPr, wml.CTCnf>('cnfStyle', mergeCnf, isEmptyCnf, copyLeaf, { formatting: false }),
  prop<wml.TcPr, wml.TblWidth>('tcW', mergeTblWidth, isEmptyTblWidth, copyLeaf),
  prop<wml.TcPr, wml.TcPrInner.GridSpan>('gridSpan', mergeGridSpan, isEmptyVal, copyLeaf),
  prop<wml.TcPr, wml.TcPrInner.HMerge>('hMerge', mergeHMerge, isEmptyStringVal, copyLeaf),
  prop<wml.TcPr, wml.TcPrInner.VMerge>('vMerge', mergeVMerge, isEmptyStringVal, copyLeaf),
  prop<wml.TcPr, wml.TcPrInner.TcBorders>('tcBorders', mergeTcBorders, isEmptyTcBorders, copyLeaf),
  prop<wml.TcPr, wml.CTShd>('shd', mergeShd, isEmptyShd, copyLeaf),
  bool<wml.TcPr>('noWrap'),
  prop<wml.TcPr, wml.TcMar>('tcMar', mergeTcMar, isEmptyTcMar, copyLeaf),
  prop<wml.TcPr, wml.TextDirection>('textDirection', mergeTextDirection, isEmptyVal, copyLeaf),
  bool<wml.TcPr>('tcFitText'),
  prop<wml.TcPr, wml.CTVerticalJc>('vAlign', mergeVAlign, isEmptyVal, copyLeaf),
  bool<wml.TcPr>('hideMark'),
] as unknown[] as Property<wml.TcPr, never>[]);

export const CELL_EXCLUDED: readonly string[] = ['cellIns', 'cellDel', 'cellMerge', 'tcPrChange'];

// ---------------------------------------------------------------- generic operations

/** The member of that name, or undefined. */
export function namedProperty<O>(catalogue: readonly Property<O, never>[], name: string): Property<O, never> | undefined {
  return catalogue.find((p) => p.name === name);
}

/** Whether every member of `owner` is empty. */
export function isEmptyOf<O>(catalogue: readonly Property<O, never>[], owner: O | undefined): boolean {
  if (owner === undefined || owner === null) return true;
  for (const p of catalogue) {
    if (!p.countsEmpty) continue;
    if (!p.isEmpty(p.get(owner))) return false;
  }
  return true;
}

/** Whether any formatting member is present (an element that is there, even with no attributes, counts: it was stated). */
export function hasDirectFormattingOf<O>(catalogue: readonly Property<O, never>[], owner: O | undefined): boolean {
  if (owner === undefined || owner === null) return false;
  for (const p of catalogue) {
    if (p.isFormatting && p.get(owner) !== undefined) return true;
  }
  return false;
}

/** Apply every member of `source` over `destination`, in place, except those named. */
export function applyCatalogue<O>(catalogue: readonly Property<O, never>[], source: O, destination: O,
  except?: ReadonlySet<string>): void {
  for (const p of catalogue) {
    if (except?.has(p.name)) continue;
    applyProperty(p, source, destination);
  }
}

/** Apply one member of `source` over `destination`. */
export function applyProperty<O, V>(p: Property<O, V>, source: O, destination: O): void {
  p.set(destination, p.merge(p.get(source), p.get(destination)));
}

/** Unset in `destination` every member that `source` states. */
export function unsetCatalogue<O>(catalogue: readonly Property<O, never>[], source: O, destination: O): void {
  for (const p of catalogue) {
    if (p.get(source) !== undefined) p.set(destination, undefined);
  }
}
