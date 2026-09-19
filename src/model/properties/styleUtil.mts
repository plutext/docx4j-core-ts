// docx4j `org.docx4j.model.styles.StyleUtil`, the catalogue-driven half: `apply`,
// `applyStyleLevel`, `isEmpty`, `unset` and `hasDirectFormatting`, derived from the tables of
// `catalogue.mts` by iteration.  Not ported: the 73 `areEqual` overloads and `StyleTree` /
// `Node` / `Tree` / `BrokenStyleRemediator`, which serve docx4j's HTML and CSS export.
//
// Names are docx4j's; where Java overloads on the argument type TypeScript cannot, the type is
// in the name (`applyRPr`, `applyPPrBase`, `applyTblPr`, ...).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import {
  type Property, type RunProps, RUN, PARAGRAPH, TABLE, CELL, TOGGLE_NAMES, TOGGLES, IND,
  applyCatalogue, unsetCatalogue, isEmptyOf, hasDirectFormattingOf, isTrue, copyLeaf,
} from './catalogue.mjs';
import { CyclicStylesException } from '../../opc/exceptions.mjs';
import { log } from './log.mjs';

/** What the paragraph merge needs from the numbering part: a level's `w:ind`. */
export interface NumberingIndents {
  /** The indent the level a `w:numPr` names contributes, or undefined. */
  getInd(numPr: wml.PPrBase.NumPr): wml.PPrBase.Ind | undefined;
}

// ---------------------------------------------------------------- isEmpty

export function isEmptyRPr(rPr: RunProps | undefined): boolean {
  return isEmptyOf(RUN, rPr);
}

export function isEmptyPPrBase(pPrBase: wml.PPrBase | undefined): boolean {
  return isEmptyOf(PARAGRAPH, pPrBase);
}

/** docx4j `isEmpty(PPr)`: the base members, the paragraph mark and the section. */
export function isEmptyPPr(pPr: wml.PPr | undefined): boolean {
  // isEmpty(SectPr) is "false for anything non-null" in docx4j ("implementation is quite basic")
  return pPr === undefined
    || (isEmptyPPrBase(pPr) && isEmptyRPr(pPr.rPr) && pPr.sectPr === undefined);
}

export function isEmptyTblPr(tblPr: wml.CTTblPrBase | undefined): boolean {
  return isEmptyOf(TABLE, tblPr);
}

export function isEmptyTcPr(tcPr: wml.TcPr | undefined): boolean {
  return isEmptyOf(CELL, tcPr);
}

export function isEmptyTrPr(trPr: wml.TrPr | undefined): boolean {
  return trPr === undefined || trPr.cnfStyleOrDivIdOrGridBefore === undefined
    || trPr.cnfStyleOrDivIdOrGridBefore.length === 0;
}

export function isEmptyTblStylePr(pr: wml.CTTblStylePr | undefined): boolean {
  return pr === undefined
    || (isEmptyPPr(pr.pPr) && isEmptyRPr(pr.rPr) && isEmptyTblPr(pr.tblPr)
      && isEmptyTrPr(pr.trPr) && isEmptyTcPr(pr.tcPr) && pr.type === undefined);
}

export function isEmptyTblStylePrList(list: readonly wml.CTTblStylePr[] | undefined): boolean {
  if (list === undefined || list.length === 0) return true;
  return !list.some((pr) => !isEmptyTblStylePr(pr));
}

const CHARACTER_STYLE = 'character';
const PARAGRAPH_STYLE = 'paragraph';
const NUMBERING_STYLE = 'numbering';
const TABLE_STYLE = 'table';

export function isEmptyStyle(style: wml.Style | undefined): boolean {
  if (style === undefined) return true;
  if (style.type === CHARACTER_STYLE) return isEmptyRPr(style.rPr);
  if (style.type === PARAGRAPH_STYLE || style.type === NUMBERING_STYLE) {
    return isEmptyPPr(style.pPr) && isEmptyRPr(style.rPr);
  }
  return isEmptyPPr(style.pPr) && isEmptyRPr(style.rPr) && isEmptyTblPr(style.tblPr)
    && isEmptyTcPr(style.tcPr) && isEmptyTblStylePrList(style.tblStylePr);
}

// ---------------------------------------------------------------- direct formatting

/**
 * Whether a run's `w:rPr` states any formatting of its own: any member of the run table other
 * than the style reference is present (an element with no attributes counts: it was stated).
 */
export function hasDirectFormattingRPr(rPr: RunProps | undefined): boolean {
  return hasDirectFormattingOf(RUN, rPr);
}

/**
 * Whether a paragraph's `w:pPr` states any formatting of its own. The paragraph mark's
 * `w:rPr` and a `w:sectPr` are not formatting of the paragraph.
 */
export function hasDirectFormattingPPr(pPr: wml.PPrBase | undefined): boolean {
  return hasDirectFormattingOf(PARAGRAPH, pPr);
}

// ---------------------------------------------------------------- unset

export function unsetRPr(source: RunProps, destination: RunProps): void {
  unsetCatalogue(RUN, source, destination);
}

export function unsetPPrBase(source: wml.PPrBase, destination: wml.PPrBase): void {
  unsetCatalogue(PARAGRAPH, source, destination);
}

// ---------------------------------------------------------------- apply: runs

/** A run properties source with nothing to say - except that a `w:rFonts` carrying only a `w:hint` still says something. */
function skipRun(source: RunProps | undefined): boolean {
  if (source === undefined) return true;
  const hint = source.rFonts !== undefined && source.rFonts.hint !== undefined;
  return isEmptyOf(RUN, source) && !hint;
}

/**
 * Apply run properties over run properties (a `w:rPr` or a `w:pPr/w:rPr` over either), in
 * place. Returns the destination, created when it was undefined and the source had something
 * to say.
 */
export function applyRPr<D extends RunProps>(source: RunProps | undefined, destination: D | undefined,
  create?: () => D): D | undefined {
  if (skipRun(source)) return destination;
  const out = destination ?? (create ? create() : ({ TYPE_NAME: 'org_docx4j_wml.RPr' } as unknown as D));
  applyCatalogue(RUN, source as D, out);
  return out;
}

/**
 * Apply one **level** of the style hierarchy over the levels beneath it, combining the
 * **toggle properties** as ECMA-376-1 §17.7.3 requires instead of overriding them.
 *
 * Only a level boundary goes through here - the table style's contribution against the
 * paragraph style's, and the paragraph style's against the character style's. Within one
 * level a style's `w:basedOn` chain is *not* XORed (§17.7.3 takes the first value found
 * walking that chain, which is what merging the chain root-first with {@link applyRPr}
 * gives), and direct formatting is not a level either: an explicit value there is used as it
 * stands, so it goes through {@link applyRPr} too.
 */
export function applyStyleLevel(source: RunProps | undefined, destination: RunProps | undefined,
  documentDefaults: RunProps | undefined): void {
  if (source === undefined || destination === undefined) return;
  applyCatalogue(RUN, source as never, destination as never, TOGGLE_NAMES);
  applyToggles(source, destination, documentDefaults, destination);
}

/**
 * Combine the twelve toggle properties of two levels of the style hierarchy and write the
 * result into a third `w:rPr`. Nothing else of either is touched.
 */
export function applyToggles(upper: RunProps | undefined, lower: RunProps | undefined,
  documentDefaults: RunProps | undefined, destination: RunProps | undefined): void {
  if (destination === undefined) return;
  for (const p of TOGGLES) {
    const value = toggle(
      upper === undefined ? undefined : p.get(upper),
      lower === undefined ? undefined : p.get(lower),
      documentDefaults === undefined ? undefined : p.get(documentDefaults),
    );
    p.set(destination, value);
  }
}

/**
 * One toggle property at one level boundary (ECMA-376-1 §17.7.3), as Word's `toggle-levels`
 * golden settled it (docx4j CR-015, 2026-09-18):
 *
 * - an upper level that says nothing is no boundary at all, so the lower value stands and the
 *   document defaults' true is not forced back on (confirmed by the golden);
 * - `w:docDefaults` is a base value, never a term of the XOR: where the defaults say true and
 *   the upper level states the property, the effective value is true;
 * - an explicit false at a style level is a term of the XOR like any other value, so
 *   false XOR lower = lower and only a stated *true* inverts what is beneath it. Word draws a
 *   run in a character style stating `<w:b w:val="0"/>` over a bold paragraph style **bold**.
 *   Where no level beneath states the property, the one level's own false stands rather than
 *   the property going absent.
 *
 * Direct formatting never comes through here: an explicit value there wins outright.
 */
export function toggle(upper: wml.BooleanDefaultTrue | undefined, lower: wml.BooleanDefaultTrue | undefined,
  documentDefault: wml.BooleanDefaultTrue | undefined): wml.BooleanDefaultTrue | undefined {
  if (upper === undefined) return lower;
  if (documentDefault !== undefined && isTrue(documentDefault)) {
    return copyLeaf(documentDefault);
  }
  // false XOR lower = lower; with no level beneath it, the level's own false stands
  if (!isTrue(upper)) return lower !== undefined ? lower : upper;
  const out: wml.BooleanDefaultTrue = { TYPE_NAME: 'org_docx4j_wml.BooleanDefaultTrue' };
  out.val = !isTrue(lower);
  return out;
}

// ---------------------------------------------------------------- apply: paragraphs

/** Whether this `w:numPr` says "not numbered" - `w:numId w:val="0"` (ECMA-376 17.9.18). */
export function numberingOff(numPr: wml.PPrBase.NumPr | undefined): boolean {
  return numPr !== undefined && numPr.numId !== undefined && numPr.numId.val === 0;
}

/** Take out of an indent the components a numbering level contributed. */
function dropNumberingInd(ind: wml.PPrBase.Ind, fromLevel: wml.PPrBase.Ind): void {
  if (fromLevel.left !== undefined && fromLevel.left === ind.left) delete ind.left;
  if (fromLevel.right !== undefined && fromLevel.right === ind.right) delete ind.right;
  if (fromLevel.hanging !== undefined && fromLevel.hanging === ind.hanging) delete ind.hanging;
  if (fromLevel.firstLine !== undefined && fromLevel.firstLine === ind.firstLine) delete ind.firstLine;
  if (fromLevel.start !== undefined && fromLevel.start === ind.start) delete ind.start;
  if (fromLevel.end !== undefined && fromLevel.end === ind.end) delete ind.end;
}

const IND_ONLY: ReadonlySet<string> = new Set([IND]);

/**
 * Apply paragraph properties over paragraph properties, in place: every member but `w:ind`,
 * then the indent the numbering level this layer brings in contributes, then `w:ind` itself
 * (an indent stated by the layer overrides the level's).
 */
export function applyPPrBase(source: wml.PPrBase | undefined, destination: wml.PPrBase,
  numbering?: NumberingIndents): void {
  if (source === undefined || isEmptyPPrBase(source)) return;

  /* A snapshot: the numPr merge writes into the destination object, so holding a reference
   * here would hand the numbering-off rule below the w:numId 0 it is looking past. */
  let inheritedNumPr: wml.PPrBase.NumPr | undefined;
  if (destination.numPr !== undefined) {
    inheritedNumPr = { numId: destination.numPr.numId, ilvl: destination.numPr.ilvl };
  }

  applyCatalogue(PARAGRAPH, source, destination, IND_ONLY);

  /* ECMA-376 17.9.18: w:numId w:val="0" takes the paragraph out of the list, so neither the
   * level's label nor its w:ind applies - only the paragraph's own.  Only the components the
   * inherited level contributed are dropped, so a w:ind the style states itself survives. */
  if (numbering !== undefined && numberingOff(source.numPr)
    && inheritedNumPr !== undefined && inheritedNumPr.numId !== undefined
    && !numberingOff(inheritedNumPr) && destination.ind !== undefined) {
    const fromLevel = numbering.getInd(inheritedNumPr);
    if (fromLevel !== undefined) {
      const own = copyLeaf(destination.ind);
      dropNumberingInd(own, fromLevel);
      destination.ind = own;
    }
  }

  /* The indent of the level of the MERGED w:numPr, since a layer may state only w:ilvl (the
   * list from the style) or only w:numId (the level from the style): CR-015 probe
   * styles-numpr-ilvl-only. */
  if (numbering !== undefined && source.numPr !== undefined
    && destination.numPr !== undefined && !numberingOff(destination.numPr)) {
    const numInd = numbering.getInd(destination.numPr);
    if (numInd !== undefined) {
      destination.ind = indProperty().merge(numInd, destination.ind);
    }
  }
  // an indent in the layer overrides any the numbering brought in
  destination.ind = indProperty().merge(source.ind, destination.ind);
}

let indEntry: Property<wml.PPrBase, wml.PPrBase.Ind> | undefined;
function indProperty(): Property<wml.PPrBase, wml.PPrBase.Ind> {
  indEntry ??= PARAGRAPH.find((p) => p.name === IND) as unknown as Property<wml.PPrBase, wml.PPrBase.Ind>;
  return indEntry;
}

/**
 * Apply a `w:pPr` over a `w:pPr`: the base members and the paragraph mark's `w:rPr`. A
 * `w:sectPr` is not a paragraph property and is not merged (CR-015 phase 4).
 */
export function applyPPr(source: wml.PPr | undefined, destination: wml.PPr | undefined,
  numbering?: NumberingIndents): wml.PPr | undefined {
  if (source === undefined || isEmptyPPr(source)) return destination;
  const out: wml.PPr = destination ?? { TYPE_NAME: 'org_docx4j_wml.PPr' };
  applyPPrBase(source, out, numbering);
  out.rPr = applyRPr<wml.ParaRPr>(source.rPr, out.rPr, () => ({ TYPE_NAME: 'org_docx4j_wml.ParaRPr' }));
  return out;
}

// ---------------------------------------------------------------- apply: tables

export function applyTblPr(source: wml.CTTblPrBase | undefined, destination: wml.CTTblPrBase | undefined): wml.CTTblPrBase | undefined {
  if (source === undefined || isEmptyTblPr(source)) return destination;
  const out: wml.CTTblPrBase = destination ?? { TYPE_NAME: 'org_docx4j_wml.CTTblPrBase' };
  applyCatalogue(TABLE, source, out);
  return out;
}

export function applyTcPr(source: wml.TcPr | undefined, destination: wml.TcPr | undefined): wml.TcPr | undefined {
  if (source === undefined || isEmptyTcPr(source)) return destination;
  const out: wml.TcPr = destination ?? { TYPE_NAME: 'org_docx4j_wml.TcPr' };
  applyCatalogue(CELL, source, out);
  return out;
}

/**
 * A row's properties, which the generated model keeps as one element list: an element of a
 * name the destination already has replaces it, and one it lacks is added.
 */
export function applyTrPr(source: wml.TrPr | undefined, destination: wml.TrPr | undefined): wml.TrPr | undefined {
  if (isEmptyTrPr(source)) return destination;
  const out: wml.TrPr = destination ?? { TYPE_NAME: 'org_docx4j_wml.TrPr' };
  const into = (out.cnfStyleOrDivIdOrGridBefore ??= []);
  for (const element of source!.cnfStyleOrDivIdOrGridBefore!) {
    const index = into.findIndex((existing) => existing.name.localPart === element.name.localPart
      && existing.name.namespaceURI === element.name.namespaceURI);
    if (index >= 0) into.splice(index, 1);
    into.push(copyLeaf(element));
  }
  return out;
}

/**
 * Merge a table style's conditional formatting (`w:tblStylePr`) into another's, *per
 * condition*: an entry for a condition the destination already has is applied on top of it,
 * and one it lacks is added. So a style `w:basedOn` another inherits the conditions it does
 * not restate, and restating a condition overrides only the properties it names.
 */
export function applyTblStylePrList(source: readonly wml.CTTblStylePr[] | undefined,
  destination: wml.CTTblStylePr[] | undefined): void {
  if (isEmptyTblStylePrList(source) || destination === undefined) return;
  for (const sourcePr of source!) {
    if (sourcePr === undefined) continue;
    let existing: wml.CTTblStylePr | undefined;
    if (sourcePr.type !== undefined) {
      existing = destination.find((d) => d !== undefined && d.type === sourcePr.type);
    }
    if (existing === undefined) {
      const copy = applyTblStylePr(sourcePr, undefined);
      if (copy !== undefined) destination.push(copy);
    } else {
      applyTblStylePr(sourcePr, existing);
    }
  }
}

export function applyTblStylePr(source: wml.CTTblStylePr | undefined, destination: wml.CTTblStylePr | undefined): wml.CTTblStylePr | undefined {
  if (isEmptyTblStylePr(source)) return destination;
  const s = source as wml.CTTblStylePr;
  const out: wml.CTTblStylePr = destination ?? ({ TYPE_NAME: 'org_docx4j_wml.CTTblStylePr' } as wml.CTTblStylePr);
  out.pPr = applyPPr(s.pPr, out.pPr);
  out.rPr = applyRPr<wml.RPr>(s.rPr, out.rPr);
  out.tblPr = applyTblPr(s.tblPr, out.tblPr);
  out.tcPr = applyTcPr(s.tcPr, out.tcPr);
  out.trPr = applyTrPr(s.trPr, out.trPr);
  out.type = s.type;
  return out;
}

/**
 * Apply one style's own properties over another's (docx4j `apply(Style, Style)`). It does not
 * climb the `w:basedOn` chain; `PropertyResolver.ancestry` does that.
 */
export function applyStyle(source: wml.Style | undefined, destination: wml.Style): wml.Style {
  if (isEmptyStyle(source)) return destination;
  const s = source as wml.Style;
  if (destination.type === undefined) destination.type = s.type;
  if (s.type === CHARACTER_STYLE) {
    destination.rPr = applyRPr<wml.RPr>(s.rPr, destination.rPr);
    return destination;
  }
  if (s.type === PARAGRAPH_STYLE || s.type === NUMBERING_STYLE) {
    destination.pPr = applyPPr(s.pPr, destination.pPr);
    destination.rPr = applyRPr<wml.RPr>(s.rPr, destination.rPr);
    return destination;
  }
  destination.tblPr = applyTblPr(s.tblPr, destination.tblPr);
  destination.trPr = applyTrPr(s.trPr, destination.trPr);
  destination.tcPr = applyTcPr(s.tcPr, destination.tcPr);
  destination.tblStylePr ??= [];
  applyTblStylePrList(s.tblStylePr, destination.tblStylePr);
  if (destination.tblStylePr.length === 0) delete destination.tblStylePr;
  destination.pPr = applyPPr(s.pPr, destination.pPr);
  destination.rPr = applyRPr<wml.RPr>(s.rPr, destination.rPr);
  return destination;
}

// ---------------------------------------------------------------- cycles

/** docx4j's hard limit on deep `w:basedOn` hierarchies. */
const CYCLE_LIMIT = 32;

/**
 * Whether the chain has come back to a style it has already seen (or has gone deeper than
 * {@link CYCLE_LIMIT}). The caller degrades gracefully by using as much of the hierarchy as it
 * has; set {@link throwOnCyclicStyles} to throw instead (docx4j's
 * `docx4j.openpackaging.exceptions.CyclicStylesException.throw` property).
 */
export function isCyclic(styleId: string, seen: readonly string[]): boolean {
  if (!seen.includes(styleId) && seen.length <= CYCLE_LIMIT) return false;
  log.warn(`Cycle detected in style basedOn hierarchy for: ${styleId} - stopping`);
  if (throwOnCyclic) throw new CyclicStylesException(`Cycle detected in style basedOn hierarchy for: ${styleId}`);
  return true;
}

let throwOnCyclic = false;

/** docx4j's `docx4j.openpackaging.exceptions.CyclicStylesException.throw`: off by default. */
export function throwOnCyclicStyles(value: boolean): void {
  throwOnCyclic = value;
}

export function throwsOnCyclicStyles(): boolean {
  return throwOnCyclic;
}
