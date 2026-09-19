// `NumberingDefinitionsPart.getInd(w:numPr)` - the indent a numbering level contributes,
// which `StyleUtil.apply(PPrBase, PPrBase, NumberingDefinitionsPart)` folds into every layer
// of the paragraph merge.  Without it a numbered paragraph's effective `w:pPr` has no indent
// at all, so property resolution cannot be measured against docx4j without it.
//
// **Scope.** This is the one piece of docx4j's `org.docx4j.model.listnumbering` that CR-001
// Phase B step 2 needs; step 3 ports the package proper (`definitions.mts`, `state.mts`,
// `formats.mts`, `Emulator.mts`) and this module goes, `NumberingDefinitionsPart.getInd`
// taking its place.  It is the same walk, read off `NumberingDefinitionsPart.getInd`,
// `getIndFromLvl`, `ListNumberingDefinition` and `AbstractListNumberingDefinition`:
//
//   numId -> w:num -> w:abstractNumId -> w:abstractNum (its levels, with w:numStyleLink
//   resolved) -> the level for w:ilvl, the instance's w:lvlOverride/w:lvl first;
//   a level's own w:pPr/w:ind is the level's indent, and where it has none, the w:ind of the
//   w:pStyle the level links to, or of the styles that style is based on.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { NumberingIndents } from './styleUtil.mjs';
import { log } from './log.mjs';

/** What the level lookup needs of the styles: a style by id (the resolver's `getStyle`). */
export type StyleLookup = (styleId: string) => wml.Style | undefined;

interface Level {
  /** The level as the instance's `w:lvlOverride` states it. */
  override?: wml.Lvl;
  /** The level as the abstract definition states it. */
  abstract?: wml.Lvl;
}

/**
 * The numbering part's levels, keyed by concrete `w:numId` then by `w:ilvl`, built once.
 * Cheap: it walks the part's `w:num` and `w:abstractNum` lists and holds references.
 */
export class NumberingLevels implements NumberingIndents {

  private readonly byNumId = new Map<string, Map<string, Level>>();

  constructor(numbering: wml.Numbering | undefined, private readonly getStyle: StyleLookup) {
    if (numbering === undefined) return;
    const abstracts = new Map<string, wml.Numbering.AbstractNum>();
    for (const abstractNum of numbering.abstractNum ?? []) {
      if (abstractNum.abstractNumId !== undefined) abstracts.set(String(abstractNum.abstractNumId), abstractNum);
    }
    // pass 1: the levels each abstract definition states itself
    const levelsOf = new Map<string, Map<string, wml.Lvl>>();
    for (const [id, abstractNum] of abstracts) {
      levelsOf.set(id, levelMap(abstractNum));
    }
    // pass 2: w:numStyleLink - an abstract definition with no levels of its own takes them
    // from the abstract definition of the w:num the numbering style names
    for (const [id, abstractNum] of abstracts) {
      const link = abstractNum.numStyleLink?.val;
      if (link === undefined || (levelsOf.get(id)?.size ?? 0) > 0) continue;
      const style = this.getStyle(link);
      const numId = style?.pPr?.numPr?.numId?.val;
      if (numId === undefined) {
        log.warn(`For w:numStyleLink, style ${link} has no w:numPr/w:numId`);
        continue;
      }
      const target = (numbering.num ?? []).find((n) => n.numId === numId);
      const targetAbstract = target?.abstractNumId?.val === undefined
        ? undefined : abstracts.get(String(target.abstractNumId.val));
      if (targetAbstract === undefined) {
        log.warn(`For w:numStyleLink, no w:abstractNum behind w:num ${String(numId)}`);
        continue;
      }
      levelsOf.set(id, levelMap(targetAbstract));
    }
    // the instances, with their w:lvlOverride levels over the abstract ones
    for (const num of numbering.num ?? []) {
      if (num.numId === undefined) continue;
      const abstractId = num.abstractNumId?.val;
      const abstractLevels = abstractId === undefined ? undefined : levelsOf.get(String(abstractId));
      const levels = new Map<string, Level>();
      for (const [ilvl, lvl] of abstractLevels ?? []) levels.set(ilvl, { abstract: lvl });
      for (const override of num.lvlOverride ?? []) {
        if (override.ilvl === undefined) continue;
        const ilvl = String(override.ilvl);
        if (override.lvl === undefined) continue;
        const level = levels.get(ilvl);
        if (level !== undefined) level.override = override.lvl;
      }
      this.byNumId.set(String(num.numId), levels);
    }
  }

  /** docx4j `NumberingDefinitionsPart.getInd(NumPr)`: `w:ilvl` is optional and means level 0. */
  getInd(numPr: wml.PPrBase.NumPr): wml.PPrBase.Ind | undefined {
    if (numPr.numId === undefined || numPr.numId.val === undefined) return undefined;
    const ilvl = numPr.ilvl?.val === undefined ? '0' : String(numPr.ilvl.val);
    return this.indOf(String(numPr.numId.val), ilvl);
  }

  /** docx4j `NumberingDefinitionsPart.getInd(String, String)`: the override level first. */
  indOf(numId: string, ilvl: string): wml.PPrBase.Ind | undefined {
    const level = this.byNumId.get(numId)?.get(ilvl);
    if (level === undefined) return undefined;
    if (level.override !== undefined) {
      const ind = this.indFromLvl(level.override);
      if (ind !== undefined) return ind;
    }
    return level.abstract === undefined ? undefined : this.indFromLvl(level.abstract);
  }

  /**
   * The indent this level contributes: **the level's own `w:pPr/w:ind` first** (ECMA-376
   * 17.9.24), and where it states none, the `w:ind` of the `w:pStyle` it links to, or of the
   * styles that style is based on.
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

function levelMap(abstractNum: wml.Numbering.AbstractNum): Map<string, wml.Lvl> {
  const out = new Map<string, wml.Lvl>();
  for (const lvl of abstractNum.lvl ?? []) {
    if (lvl.ilvl !== undefined) out.set(String(lvl.ilvl), lvl);
  }
  return out;
}
