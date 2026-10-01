// A subset of Office JS `Word.TrackedChange` over Word's revision markup (ECMA-376 17.13.5).
// Accepting and rejecting follow docx4j's AcceptTrackedChanges (docx4j-core
// org.docx4j.convert.out.common.preprocess): a w:ins is unwrapped, a w:del removed, a deleted
// paragraph mark joins its paragraph with the next. CR-002 phase F.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { textOf } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { type Element, type Located, type RevisionKind, typeNameOf, runItemsOf, revisionKindOf, linkParents, segmentsOf, textOfView, cellsOf, rowsOf, childrenOf, W_NS, itemTextOf } from './tree.mjs';
import { removeMarkers } from './comments.mjs';
import { Range } from './Range.mjs';
import type { Paragraph } from './Paragraph.mjs';
import { rPrFromElements } from '@docx4j/generated-objects-ts/builders/wml';
import { revisionDateOf, toRestoredText, toDeletedText, restoreRPr, restorePPr, pruneParagraphProperties, HOLDER_TYPES, holderItemsOf } from './tracking.mjs';
import { tableRecordsOf, hasTableRecords, tableChangeText, rowText as tableRowText, acceptTable, rejectTable } from './tableRevisions.mjs';
import { deepCopy } from '@docx4j/generated-objects-ts';

/** Office JS `Word.ChangeTrackingState` (`Unknown` is not produced here). */
export type TrackedChangeType = 'Added' | 'Deleted' | 'Formatted' | 'None' | 'Unknown';

/** Which piece of revision markup a `TrackedChange` is over (an extension: Office JS has no such member). */
export type TrackedChangeTarget =
  /** `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` around runs. */
  | { kind: 'run'; revision: RevisionKind; element: Element; owner: Element[]; value: wml.CTTrackChange }
  /** `w:pPr/w:rPr/w:ins` or `w:pPr/w:rPr/w:del`: the paragraph mark. */
  | { kind: 'mark'; mark: 'ins' | 'del'; value: wml.CTTrackChange }
  /** `w:rPr/w:rPrChange`: a run's formatting. */
  | { kind: 'runProperties'; run: wml.R; value: wml.CTRPrChange }
  /** `w:pPr/w:pPrChange`: the paragraph's properties. */
  | { kind: 'paragraphProperties'; value: wml.CTPPrChange }
  /**
   * `w:trPr/w:ins` or `w:trPr/w:del`: a table row. `inner` is the revision markup the row's own
   * cells carry (the `w:ins` or `w:del` around their runs and on their paragraph marks), which
   * `Body.getTrackedChanges` reports through the row rather than separately, as Office JS does:
   * one change per row. Accepting or rejecting the row applies them too. `table` is the table the
   * row is in, which goes with its last row.
   */
  | { kind: 'row'; row: 'ins' | 'del'; tr: Element<wml.Tr>; owner: Element[]; value: wml.CTTrackChange; inner: TrackedChange[]; table?: Located<wml.Tbl> }
  /**
   * Touching insertions, deletions or formatting changes by one author - `w:ins` or `w:del` around
   * runs and inserted or deleted paragraph marks, or runs' `w:rPrChange`, with nothing unrevised
   * between them - which Office JS lists as one change (CR-002 section 29, checks 22, 23 and 28): a
   * deleted paragraph's text and mark, a deletion across a mark, bold put on two runs of different
   * formatting. `pieces` are the single changes, in document order; `value` is the first's.
   */
  | { kind: 'group'; revision: 'ins' | 'del' | 'format'; pieces: TrackedChange[]; value: wml.CTTrackChange }
  /**
   * A table's property records - `w:tblPrChange`, `w:tblGridChange`, `w:tblPrExChange`, and every
   * `w:trPrChange` and `w:tcPrChange` in it - which Word writes as one set for one change and Office
   * JS lists as one `Formatted` change (CR-002 section 35, checks 27 and 28). `value` is the table's
   * record, else the first row's or cell's.
   */
  | { kind: 'tableProperties'; tbl: Element<wml.Tbl>; value: wml.CTTrackChange }
  /** `w:tcPr/w:cellIns` or `w:tcPr/w:cellDel`: an inserted or deleted cell (section 35, check 27). */
  | { kind: 'cell'; cell: 'ins' | 'del'; tc: Element<wml.Tc>; owner: Element[]; tr: wml.Tr; value: wml.CTTrackChange }
  /** `w:pPr/w:rPr/w:rPrChange`: a paragraph mark's formatting, which groups with its runs' (check 27). */
  | { kind: 'markProperties'; value: wml.ParaRPrChange }
  /**
   * `w:sectPr/w:sectPrChange`: a section's properties, at a paragraph's section break or the body's
   * last section. `text` is the section's, as Office JS reports it (check 27: the text, then `\f`).
   */
  | { kind: 'sectionProperties'; sectPr: wml.SectPr; value: wml.CTSectPrChange; text: string }
  /**
   * A content control or custom XML element that is itself an insertion: `w:customXmlInsRangeStart`
   * and `w:customXmlInsRangeEnd` around it, the end first inside its content, and a second pair the
   * other way round at its end - Word's form for a control put in under tracking (check 30, CR-002
   * section 37 item 7), which groups with the insertions around and inside it. `holder` is the
   * `w:sdt` or `w:customXml`, `owner` the list holding it, `outer` the pair in that list and `inner`
   * the pair in its content. Rejecting takes the element away and leaves its content where it
   * stood; accepting takes the markers away. Its text is its content's, which the pieces inside
   * report.
   */
  | { kind: 'holder'; holder: Element; owner: Element[]; outer: [Element, Element]; inner: [Element, Element]; value: wml.CTTrackChange };

/**
 * A subset of Office JS `Word.TrackedChange`: `type`, `author`, `date`, `text`, `accept()`,
 * `reject()` and `getRange()`, plus `element`, `id` and `target` as extensions. A view, like
 * `Paragraph` and `Range`: it holds the markup it is over, nothing is cached.
 */
export class TrackedChange {
  constructor(
    readonly target: TrackedChangeTarget,
    /** The paragraph the change is in; undefined for a row revision. */
    readonly paragraph: Paragraph | undefined,
  ) {}

  /** Office JS `Word.ChangeTrackingState`. */
  get type(): TrackedChangeType {
    const t = this.target;
    switch (t.kind) {
      case 'run': return t.revision === 'ins' || t.revision === 'moveTo' ? 'Added' : 'Deleted';
      case 'mark': return t.mark === 'ins' ? 'Added' : 'Deleted';
      case 'row': return t.row === 'ins' ? 'Added' : 'Deleted';
      case 'group': return t.revision === 'ins' ? 'Added' : t.revision === 'del' ? 'Deleted' : 'Formatted';
      case 'cell': return t.cell === 'ins' ? 'Added' : 'Deleted';
      case 'holder': return 'Added';
      default: return 'Formatted';
    }
  }

  get author(): string {
    return this.target.value.author;
  }

  /**
   * When the change was made: `w16du:dateUtc` if the markup has it, else `w:date` read as Word writes
   * it, local wall-clock time (`revisionDateOf`, CR-002 section 29); a group's is its first piece's,
   * as Office JS reports it.
   */
  get date(): Date | undefined {
    return revisionDateOf(this.target.value);
  }

  /** The annotation id (`w:id`; an extension). */
  get id(): number {
    return this.target.value.id;
  }

  /** The markup this change is over (an extension): the `w:ins` element, the `w:rPrChange`, ... */
  get element(): object {
    const t = this.target;
    switch (t.kind) {
      case 'run': return t.element;
      case 'row': return t.tr;
      case 'group': return t.pieces[0]!.element;
      case 'tableProperties': return t.tbl;
      case 'cell': return t.tc;
      case 'sectionProperties': return t.sectPr;
      case 'holder': return t.holder;
      default: return t.value;
    }
  }

  /** The text the change covers; `\r` for a paragraph mark, as Office JS gives it (check 22). */
  get text(): string {
    const t = this.target;
    switch (t.kind) {
      case 'mark': return '\r';
      case 'markProperties': return '\r';
      case 'group': return t.pieces.map((piece) => piece.text).join('');
      case 'tableProperties': return tableChangeText(t.tbl.value);
      case 'cell': return tableRowText(t.tr);
      case 'sectionProperties': return t.text;
      case 'run': return segmentsOf(t.element.value as object, { view: t.revision === 'ins' || t.revision === 'moveTo' ? 'accepted' : 'original' }).map((s) => s.text).join('');
      case 'runProperties': return textOf(t.run);
      case 'paragraphProperties': return this.paragraph?.text ?? '';
      // a deleted row's content is `w:delText` inside `w:del`, which the accepted view skips:
      // its text is what the row said before, as for a run deletion
      case 'row': return rowText(t.tr.value, t.row === 'del' ? 'original' : 'accepted');
      default: return '';
    }
  }

  /** The range the change covers; a deletion is a collapsed range where it sits in the accepted text. */
  getRange(): Range | undefined {
    const p = this.paragraph;
    if (!p) return undefined;
    const t = this.target;
    if (t.kind === 'run') {
      const [start, end] = spanOf(p, t.element);
      return new Range(p, start, end);
    }
    if (t.kind === 'group') {
      // a range is one paragraph's: the group's text in the paragraph it starts in
      const here = t.pieces.filter((piece) => (piece.target.kind === 'run' || piece.target.kind === 'runProperties') && piece.paragraph?.element === p.element).map((piece) => piece.getRange()!);
      if (here.length === 0) return t.pieces[0]!.getRange();
      return new Range(p, Math.min(...here.map((r) => r.start)), Math.max(...here.map((r) => r.end)));
    }
    if (t.kind === 'runProperties') {
      const segs = p.segments().filter((s) => s.run === t.run);
      if (segs.length === 0) return p.getRange('Start');
      return new Range(p, segs[0]!.start, segs[segs.length - 1]!.end);
    }
    return p.getRange('Whole');
  }

  /**
   * Keeps the change: a `w:ins` is unwrapped, a `w:del` removed, a deleted mark joins the paragraphs.
   * Either half of a move keeps the whole move, as Word's Review tab does (CR-002 section 29).
   */
  accept(): void {
    const t = this.target;
    switch (t.kind) {
      case 'run':
        if ((t.revision === 'moveFrom' || t.revision === 'moveTo') && (resolveMove(this, 'accept') || resolveRangelessHalf(this, 'accept'))) return;
        if (t.revision === 'ins' || t.revision === 'moveTo') unwrap(t); else removeRevision(t, this.paragraph);
        return;
      case 'mark':
        if (t.mark === 'ins') dropMark(this.paragraph, 'ins'); else joinWithNext(this.requireParagraph());
        return;
      case 'runProperties':
        delete t.run.rPr?.rPrChange;
        return;
      case 'paragraphProperties': {
        const p = this.requireParagraph().p;
        delete p.pPr?.pPrChange;
        pruneParagraphProperties(p);
        return;
      }
      case 'row':
        // an accepted insertion keeps the row, so its content's own w:ins must go with the
        // w:trPr/w:ins; an accepted deletion takes the row and everything in it away
        if (t.row === 'ins') { delete t.tr.value.trPr?.ins; pruneRowProperties(t.tr.value); applyInner(t.inner, 'accept'); }
        else removeRow(t);
        return;
      case 'group':
        resolvePieces(t.pieces, 'accept');
        return;
      case 'tableProperties':
        acceptTable(t.tbl.value);
        return;
      case 'cell':
        // an inserted cell stays, a deleted one goes (check 27)
        if (t.cell === 'ins') delete t.tc.value.tcPr?.cellIns; else removeCell(t.tc, t.owner);
        return;
      case 'markProperties': {
        const p = this.requireParagraph().p;
        delete p.pPr?.rPr?.rPrChange;
        pruneParagraphProperties(p);
        return;
      }
      case 'sectionProperties':
        delete t.sectPr.sectPrChange;
        return;
      case 'holder':
        // the element stays, the markers go (check 30: accept all leaves the control, no marker)
        removeHolderMarkers(t);
        return;
    }
  }

  /**
   * Puts back what was there: a `w:ins` is removed, a `w:del` restored, a `w:rPrChange` re-applied.
   * Either half of a move rejects the whole move, as Word's Review tab does (CR-002 section 29).
   */
  reject(): void {
    const t = this.target;
    switch (t.kind) {
      case 'run':
        if ((t.revision === 'moveFrom' || t.revision === 'moveTo') && (resolveMove(this, 'reject') || resolveRangelessHalf(this, 'reject'))) return;
        if (t.revision === 'ins' || t.revision === 'moveTo') removeRevision(t, this.paragraph); else restoreDeleted(t);
        return;
      case 'mark':
        if (t.mark === 'del') dropMark(this.paragraph, 'del'); else joinWithNext(this.requireParagraph(), true);
        return;
      case 'runProperties':
        restoreRPr(t.run, rPrFromElements(t.value.rPr));
        return;
      case 'paragraphProperties': {
        const paragraph = this.requireParagraph();
        restorePPr(paragraph.p, t.value.pPr);
        dropLevelIndent(paragraph);
        writeRecordedLevel(paragraph.p);
        return;
      }
      case 'row':
        // a rejected deletion keeps the row, so its content comes back too (w:delText to w:t,
        // the w:del unwrapped, the deleted marks dropped); a rejected insertion takes it away
        if (t.row === 'del') { delete t.tr.value.trPr?.del; pruneRowProperties(t.tr.value); applyInner(t.inner, 'reject'); }
        else removeRow(t);
        return;
      case 'group':
        resolvePieces(t.pieces, 'reject');
        return;
      case 'tableProperties':
        rejectTable(t.tbl.value);
        return;
      case 'cell':
        if (t.cell === 'del') delete t.tc.value.tcPr?.cellDel; else removeCell(t.tc, t.owner);
        return;
      case 'markProperties':
        restoreMarkProperties(this.requireParagraph().p, t.value);
        return;
      case 'sectionProperties':
        restoreSection(t.sectPr, t.value);
        return;
      case 'holder':
        // the element goes, its content stays where it stood: what is left of it after the pieces
        // inside were rejected (nothing, for content inserted with it; check 30: reject all removes the
        // control, and an existing control, which has no markers, stays)
        unwrapHolder(t);
        return;
    }
  }

  private requireParagraph(): Paragraph {
    if (!this.paragraph) throw new Docx4JException('This tracked change has no paragraph');
    return this.paragraph;
  }
}

// --- the operations ------------------------------------------------------------------------

/**
 * The text of a table row in one of the two views, a line per paragraph (docx4j TextUtils, which
 * the builders' `textOf` is - but that reads the accepted view only, and a deleted row's content
 * is all `w:delText`).
 */
function rowText(tr: wml.Tr, view: 'accepted' | 'original'): string {
  // Office JS's text for a row: its cells with a tab between them, then `\r\n`; a cell's paragraphs
  // with `\r` between them, a nested table's rows each with their own `\r\n` and no separator of their
  // own (checks 30 and 31: "row one\r\n", "nested\r\nouter cell\r\n")
  const cellText = (tc: wml.Tc): string => {
    let out = '';
    let afterParagraph = false;
    for (const el of childrenOf(tc) ?? []) {
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.P') {
        out += (afterParagraph ? '\r' : '') + textOfView(el.value as object, { view });
        afterParagraph = true;
      } else if (tn === 'org_docx4j_wml.Tbl') {
        for (const row of rowsOf(el.value as wml.Tbl)) out += rowText(row.element.value, view);
        afterParagraph = false;
      } else if (typeof el.value === 'object' && el.value !== null) {
        // a block control or custom XML element: what it holds, as if it were not there
        for (const inner of childrenOf(el.value) ?? []) {
          if (typeNameOf(inner) === 'org_docx4j_wml.P') { out += (afterParagraph ? '\r' : '') + textOfView(inner.value as object, { view }); afterParagraph = true; }
        }
      }
    }
    return out;
  };
  return cellsOf(tr).map((cell) => cellText(cell.element.value)).join('\t') + '\r\n';
}

/** The changes inside a row, last first, so that a paragraph join never disturbs one still to do. */
function applyInner(inner: TrackedChange[], what: 'accept' | 'reject'): void {
  for (let i = inner.length - 1; i >= 0; i--) inner[i]![what]();
}

/**
 * Word records a list paragraph's indent in `w:pPrChange` even where the list level gives it, and on
 * reject writes no indent back: it restored "The first numbered item." to its numbered list with no
 * `w:ind`, where the change recorded `left=720 hanging=360`, the level's own (`test/README.md` check
 * 18, `reject-all.docx`, section 9), and check 25's level-0 items with none where it recorded only
 * `hanging=360`, part of the level's `left=720 hanging=360` (`25b-bullets-reject.docx`; CR-002
 * section 29). So a restored `w:ind` whose every attribute is the level's is dropped, leaving the level
 * to give it; one with an attribute of its own - a genuinely direct indent - is restored. The numbering
 * part must have been read (`await pkg.getBody()` or `getPropertyResolver()` reads it); when it has
 * not, the recorded indent is restored as it was.
 */
function dropLevelIndent(paragraph: Paragraph): void {
  const pPr = paragraph.p.pPr;
  const numPr = pPr?.numPr;
  const ind = pPr?.ind;
  if (!pPr || !numPr || !ind) return;
  const part = (paragraph.parentBody.package_ as { numberingDefinitionsPart?: { getInd(numPr: wml.PPrBase.NumPr): wml.PPrBase.Ind | undefined } } | undefined)?.numberingDefinitionsPart;
  let level: wml.PPrBase.Ind | undefined;
  try {
    level = part?.getInd(numPr);
  } catch {
    return;                                                             // not read: nothing to compare with
  }
  if (!level || !indentWithin(ind, level)) return;
  delete pPr.ind;
  pruneParagraphProperties(paragraph.p);
}

/**
 * A restored `w:numPr` that names a list but no level is at level 0, and Word writes the 0: it
 * restored the level recorded, not the current one, when the two differed (check 25's
 * `25e-bullet-demote.docx`, recorded `numId 1` alone and current `ilvl 1`, rejected to `ilvl 0,
 * numId 1`), and wrote `w:ilvl="0"` every time the record had none (check 18's section 9, 25b, 25f;
 * CR-002 section 29).
 */
function writeRecordedLevel(p: wml.P): void {
  const numPr = p.pPr?.numPr;
  if (!numPr?.numId || Number(numPr.numId.val) === 0 || numPr.ilvl) return;
  numPr.ilvl = { TYPE_NAME: 'org_docx4j_wml.PPrBase.NumPr.Ilvl', val: 0 };
  linkParents(numPr.ilvl, numPr);
}

/** True when every attribute `ind` sets, `level` sets to the same value. */
function indentWithin(ind: wml.PPrBase.Ind, level: wml.PPrBase.Ind): boolean {
  const attributes = Object.entries(ind).filter(([key, value]) => key !== 'TYPE_NAME' && key !== 'PARENT' && value !== undefined);
  return attributes.length > 0 && attributes.every(([key, value]) => String((level as Record<string, unknown>)[key]) === String(value));
}

/**
 * A group's pieces resolved: its runs last first, then its marks last first, so that a paragraph join
 * never strands a piece still to do. Office JS and Word's Review tab resolve the group as one (checks
 * 22 and 23), and so, piece by piece, does this.
 */
function resolvePieces(pieces: TrackedChange[], what: 'accept' | 'reject'): void {
  const runs = pieces.filter((piece) => piece.target.kind !== 'mark');
  const marks = pieces.filter((piece) => piece.target.kind === 'mark');
  for (let i = runs.length - 1; i >= 0; i--) runs[i]![what]();
  for (let i = marks.length - 1; i >= 0; i--) marks[i]![what]();
}

/**
 * The list an element is in now: `owner` as the change was listed, else the list its `PARENT` keeps -
 * a piece listed inside a holder that was unwrapped since (its content moved out into the owner,
 * check 30) is still found.
 */
function currentOwner(owner: Element[], element: Element): Element[] {
  if (owner.includes(element)) return owner;
  const parent = (element.value as { PARENT?: object } | undefined)?.PARENT;
  const actual = parent ? runItemsOf(parent) ?? childrenOf(parent) : undefined;
  return actual?.includes(element) ? actual : owner;
}

function remove(owner: Element[], element: Element): void {
  const list = currentOwner(owner, element);
  const i = list.indexOf(element);
  if (i >= 0) list.splice(i, 1);
}

/** A holder's four insertion markers taken away; the element stays. */
function removeHolderMarkers(t: Extract<TrackedChangeTarget, { kind: 'holder' }>): void {
  remove(t.owner, t.outer[0]);
  remove(t.owner, t.outer[1]);
  const items = holderItemsOf(t.holder.value as object) ?? [];
  remove(items, t.inner[0]);
  remove(items, t.inner[1]);
}

/** A holder's insertion rejected: the markers go, and its content takes its place in the owner. */
function unwrapHolder(t: Extract<TrackedChangeTarget, { kind: 'holder' }>): void {
  removeHolderMarkers(t);
  const owner = currentOwner(t.owner, t.holder);
  const i = owner.indexOf(t.holder);
  if (i < 0) return;
  const items = holderItemsOf(t.holder.value as object) ?? [];
  const moved = items.splice(0);                                       // the holder keeps nothing
  owner.splice(i, 1, ...moved);
  linkParents(moved, (t.holder.value as { PARENT?: object }).PARENT);
}

/**
 * The insertion of a content control or custom XML element, when `items[i]` is a
 * `w:customXmlInsRangeStart` and what follows is Word's form for one (check 30): the element, whose
 * content starts with that marker's `w:customXmlInsRangeEnd` and ends with a second start, and that
 * one's end right after the element. Anything else is not one (a marker on its own is left alone).
 */
export function holderInsertionAt(items: Element[], i: number, paragraph: Paragraph | undefined): TrackedChange | undefined {
  const start = items[i]!;
  if (start.name?.localPart !== 'customXmlInsRangeStart') return undefined;
  const holder = items[i + 1];
  const outerEnd = items[i + 2];
  if (!holder || !outerEnd || !HOLDER_TYPES.has(typeNameOf(holder) ?? '') || outerEnd.name?.localPart !== 'customXmlInsRangeEnd') return undefined;
  const inner = childrenOf(holder.value as object) ?? [];
  const innerEnd = inner[0];
  const innerStart = inner[inner.length - 1];
  if (inner.length < 2 || !innerEnd || !innerStart || innerEnd.name?.localPart !== 'customXmlInsRangeEnd' || innerStart.name?.localPart !== 'customXmlInsRangeStart') return undefined;
  const idOf = (marker: Element): unknown => (marker.value as { id?: unknown }).id;
  if (idOf(innerEnd) !== idOf(start) || idOf(outerEnd) !== idOf(innerStart)) return undefined;
  return new TrackedChange({ kind: 'holder', holder, owner: items, outer: [start, outerEnd], inner: [innerEnd, innerStart], value: start.value as wml.CTTrackChange }, paragraph);
}

/**
 * A row taken away, and its table with it when it was the last: a table has no revision of its own,
 * so one whose rows were all inserted was inserted, and one whose rows were all deleted was deleted.
 * A `w:tbl` left with no rows is not a table any more (CR-002 section 37).
 */
function removeRow(t: Extract<TrackedChangeTarget, { kind: 'row' }>): void {
  remove(t.owner, t.tr as Element);
  if (t.table && rowsOf(t.table.element.value).length === 0) remove(t.table.container, t.table.element as Element);
}

/**
 * A revision taken away with its runs - a rejected insertion, an accepted deletion - and the comments
 * on it. Word writes a comment's range start inside the `w:ins` of the text it is on (CR-002 section
 * 38 item 1); when that text goes, Word removes the comment with it - its markers, its reference run
 * and its entry (check 32, run 2026-10-02: rejecting the insertion, rejecting all, and accepting the
 * deletion in the same shape, each left no comment). So a comment whose markers come out of the
 * revision with no text left between them is removed here: the markers and the reference now, the
 * entries in the comment parts when the comments are next read or the package saved
 * (`pendingCommentRemovals`). A comment with text left in its range - one that reached beyond the
 * revision - keeps its markers where the text was (not measured).
 */
function removeRevision(t: Extract<TrackedChangeTarget, { kind: 'run' }>, paragraph: Paragraph | undefined): void {
  const owner = currentOwner(t.owner, t.element);
  const i = owner.indexOf(t.element);
  if (i < 0) return;
  const kept = (runItemsOf(t.element.value as object) ?? []).filter((item) => item.name?.localPart === 'commentRangeStart' || item.name?.localPart === 'commentRangeEnd');
  // the comments whose range held the revision: a start inside it (Word's form), or one before it
  // with its end after it (the engine's own, the markers beside the runs)
  const idOf = (el: Element): unknown => (el.value as { id?: unknown }).id;
  const spanning = new Set<unknown>(kept.filter((m) => m.name?.localPart === 'commentRangeStart').map(idOf));
  for (let s = 0; s < i; s++) {
    const el = owner[s]!;
    if (el.name?.localPart !== 'commentRangeStart') continue;
    const id = idOf(el);
    if (owner.some((other, e) => e > i && other.name?.localPart === 'commentRangeEnd' && idOf(other) === id)) spanning.add(id);
  }
  owner.splice(i, 1, ...kept);
  linkParents(kept, (t.element.value as { PARENT?: object }).PARENT);
  const body = paragraph?.parentBody;
  if (!body) return;
  for (const id of spanning) {
    const start = owner.find((el) => el.name?.localPart === 'commentRangeStart' && idOf(el) === id);
    if (!start || !rangeIsEmpty(owner, start, id as number)) continue;
    removeMarkers(body.container, id as number);
    const pkg = body.package_ as { pendingCommentRemovals?: Set<number> } | undefined;
    pkg?.pendingCommentRemovals?.add(id as number);
  }
}

/** True when nothing with text stands between a comment's range start and its end in the same list. */
function rangeIsEmpty(owner: Element[], start: Element, id: number): boolean {
  const from = owner.indexOf(start);
  if (from < 0) return false;
  for (let i = from + 1; i < owner.length; i++) {
    const el = owner[i]!;
    if (el.name?.localPart === 'commentRangeEnd' && (el.value as { id?: unknown }).id === id) return true;
    if (hasText(el)) return false;
  }
  return false;                                                        // the end is elsewhere: text may be
}

function hasText(el: Element): boolean {
  const v = el.value;
  if (typeof v !== 'object' || v === null) return false;
  if (typeNameOf(el) === 'org_docx4j_wml.R') return ((v as wml.R).content ?? []).some((item) => itemTextOf(item as Element) !== undefined);
  const kind = revisionKindOf(el);
  if (kind === 'del' || kind === 'moveFrom') return false;             // deleted text is no text in the accepted view
  return (runItemsOf(v) ?? []).some(hasText);
}

/** A `w:ins` or `w:moveTo` accepted: its runs take its place. */
function unwrap(t: Extract<TrackedChangeTarget, { kind: 'run' }>): void {
  const i = t.owner.indexOf(t.element);
  if (i < 0) return;
  const items = runItemsOf(t.element.value as object) ?? [];
  t.owner.splice(i, 1, ...items);
  const parent = (t.value as { PARENT?: object }).PARENT;
  for (const item of items) linkParents(item, parent);
}

/** A `w:del` or `w:moveFrom` rejected: its runs come back, with `w:delText` turned into `w:t`. */
function restoreDeleted(t: Extract<TrackedChangeTarget, { kind: 'run' }>): void {
  for (const item of runItemsOf(t.element.value as object) ?? []) {
    if (typeNameOf(item) === 'org_docx4j_wml.R') toRestoredText(item.value as wml.R);
  }
  unwrap(t);
}

function dropMark(paragraph: Paragraph | undefined, which: 'ins' | 'del'): void {
  if (!paragraph) return;
  const rPr = paragraph.p.pPr?.rPr;
  if (rPr) delete rPr[which];
  pruneParagraphProperties(paragraph.p);
}

/**
 * docx4j AcceptTrackedChanges: the paragraph takes the next one's content and properties, so the
 * mark that survives is the next one's. Called when a deleted mark is accepted and when an
 * inserted one is rejected.
 *
 * `fallbackToPrevious` is for the second case at the end of a container, where this package's
 * inserted paragraph carries its own mark (section 13) and there is nothing after it to join
 * with: the paragraph then goes, its content joining the one before. With neither neighbour the
 * mark is simply dropped.
 */
export function joinWithNext(paragraph: Pick<Paragraph, 'container' | 'element' | 'p'>, fallbackToPrevious = false): void {
  const container = currentOwner(paragraph.container, paragraph.element as Element);
  const i = container.indexOf(paragraph.element);
  const next = i >= 0 ? container[i + 1] : undefined;
  const p = paragraph.p;
  if (!next || typeNameOf(next) !== 'org_docx4j_wml.P') {
    const previous = fallbackToPrevious && i > 0 ? container[i - 1] : undefined;
    if (previous && typeNameOf(previous) === 'org_docx4j_wml.P') {
      const into = previous.value as wml.P;
      const items = (into.content ??= []);
      for (const item of p.content ?? []) items.push(item as never);
      linkParents(items, into);
      container.splice(i, 1);
      return;
    }
    const rPr = p.pPr?.rPr;
    if (rPr) { delete rPr.ins; delete rPr.del; }
    pruneParagraphProperties(p);
    return;
  }
  const nextP = next.value as wml.P;
  const content = (p.content ??= []);
  for (const item of nextP.content ?? []) content.push(item as never);
  linkParents(content, p);
  p.pPr = nextP.pPr;
  if (p.pPr) linkParents(p.pPr, p);
  // The mark that survives is the next paragraph's, and so is the paragraph's identity: Word keeps the
  // second paragraph's w14:paraId when a mark goes (check 18's accept-all and reject-all, check 22's
  // 22-accept-whole; found by the editor's C5). Its w14:textId and session ids come with it; Word
  // writes a new textId, the text having changed, which Word will do again when it next saves.
  for (const key of PARAGRAPH_ATTRIBUTES) {
    const value = nextP[key];
    if (value === undefined) delete p[key]; else p[key] = value;
  }
  container.splice(i + 1, 1);
}

/** A `w:p`'s own attributes, which go with its mark. */
const PARAGRAPH_ATTRIBUTES = ['paraId', 'textId', 'rsidDel', 'rsidP', 'rsidR', 'rsidRDefault', 'rsidRPr', 'noSpellErr'] as const;

/** A row's `w:trPr` with nothing left in it goes, as Word writes none (check 18's accept-all and reject-all). */
function pruneRowProperties(tr: wml.Tr): void {
  const trPr = tr.trPr;
  if (trPr && !trPr.ins && !trPr.del && !trPr.trPrChange && !(trPr.cnfStyleOrDivIdOrGridBefore?.length)) delete tr.trPr;
}

/** The span of a run-level revision in the paragraph's accepted text; collapsed for a deletion. */
function spanOf(paragraph: Paragraph, element: Element): [number, number] {
  const segs = paragraph.segments();
  const inside = segs.filter((s) => s.revision?.element === element);
  if (inside.length > 0) return [inside[0]!.start, inside[inside.length - 1]!.end];
  const at = acceptedOffsetOf(paragraph.p, element);
  return [at, at];
}

/** How much accepted text comes before `target` in the paragraph. */
function acceptedOffsetOf(p: wml.P, target: Element): number {
  let pos = 0;
  let found = -1;
  const visit = (items: Element[] | undefined): void => {
    if (!items || found >= 0) return;
    for (const el of items) {
      if (found >= 0) return;
      if (el === target) { found = pos; return; }
      if (typeNameOf(el) === 'org_docx4j_wml.R') { pos += segmentsOf({ content: [el] }).map((s) => s.text).join('').length; continue; }
      const kind = revisionKindOf(el);
      if (kind === 'del' || kind === 'moveFrom') continue;
      visit(runItemsOf(el.value as object));
    }
  };
  visit(runItemsOf(p));
  return found >= 0 ? found : pos;
}

// --- moves -------------------------------------------------------------------------------------

/**
 * One piece of a move's markup, in document order: a range marker, a run-level half (`w:moveFrom`
 * or `w:moveTo` around runs), or a moved paragraph's mark (`w:pPr/w:rPr/w:moveFrom` or `w:moveTo`).
 */
type MovePiece =
  | { kind: 'start'; side: 'from' | 'to'; id: number; name: string; element: Element; owner: Element[] }
  | { kind: 'end'; side: 'from' | 'to'; id: number; element: Element; owner: Element[] }
  | { kind: 'half'; side: 'from' | 'to'; element: Element; owner: Element[]; names: string[] }
  | { kind: 'mark'; side: 'from' | 'to'; paragraph: Element<wml.P>; container: Element[]; names: string[] }
  /** Anything else directly in a paragraph while a range is open: text typed into moved text, say. */
  | { kind: 'inner'; side: 'from' | 'to'; element: Element; owner: Element[]; names: string[] };

const RANGE_MARKERS: Readonly<Record<string, { kind: 'start' | 'end'; side: 'from' | 'to' }>> = {
  moveFromRangeStart: { kind: 'start', side: 'from' }, moveFromRangeEnd: { kind: 'end', side: 'from' },
  moveToRangeStart: { kind: 'start', side: 'to' }, moveToRangeEnd: { kind: 'end', side: 'to' },
};

/**
 * Accepts or rejects the whole move a run-level half belongs to, as Word's Review tab does from
 * either half (`test/README.md` check 18, Word 16.0.20326.20158; Office JS cannot list a move at
 * all there, check 17). Word pairs the halves by the `w:name` on their range starts
 * (`w:moveFromRangeStart`, `w:moveToRangeStart`); `w:moveFrom` and `w:moveTo` carry none, so a half
 * belongs to the move whose range of its side is open where it stands. Accepting keeps the
 * destination and removes the source, paragraph mark and all; rejecting the other way round; the
 * four range markers go either way. Returns false for a half in no named range - markup without
 * range markers, or a half already resolved - which the caller then treats on its own, as before.
 */
function resolveMove(change: TrackedChange, action: 'accept' | 'reject'): boolean {
  const t = change.target;
  if (t.kind !== 'run' || !change.paragraph) return false;
  const story = storyOf(change.paragraph.p);
  const pieces = movePiecesOf(story);
  const half = pieces.find((piece) => piece.kind === 'half' && piece.element === t.element);
  if (!half || half.kind !== 'half' || half.names.length === 0) return false;
  const name = half.names[0]!;
  const ids = new Set(pieces.filter((piece) => piece.kind === 'start' && piece.name === name).map((piece) => (piece as { id: number }).id));
  const mine = pieces.filter((piece) =>
    (piece.kind === 'start' && piece.name === name) || (piece.kind === 'end' && ids.has(piece.id))
    || ((piece.kind === 'half' || piece.kind === 'mark' || piece.kind === 'inner') && piece.names.includes(name)));
  // A destination whose source is gone is kept whichever way it is resolved: Word rejected such a
  // move and left its text, plain, where it stood (test/README.md check 20 case 12).
  const orphan = !mine.some((piece) => piece.side === 'from');
  const kept = action === 'accept' || orphan ? 'to' : 'from';
  const broken = orphan || !mine.some((piece) => piece.side === 'to');

  // the markers first: a range end between two paragraphs would stop the paragraph join below
  for (const piece of mine) if (piece.kind === 'start' || piece.kind === 'end') remove(piece.owner, piece.element);
  // Rejected, a move takes with it everything typed into its destination: Word's default when moved
  // text was edited (the Tracked Moves Conflict Dialog's "keep original location text", check 20
  // cases 01 and 04). Text typed before the range starts is outside the move and stays (case 03).
  // What accepting does to text typed into the source was not measured, so it is left alone.
  if (kept === 'from') {
    for (const piece of mine) if (piece.kind === 'inner' && piece.side === 'to') remove(piece.owner, piece.element);
  }
  for (const piece of mine) {
    if (piece.kind !== 'half') continue;
    const target = { kind: 'run' as const, revision: piece.side === 'from' ? 'moveFrom' as const : 'moveTo' as const, element: piece.element, owner: piece.owner, value: piece.element.value as wml.CTTrackChange };
    if (piece.side === kept) { if (piece.side === 'from') restoreDeleted(target); else unwrap(target); }
    else remove(piece.owner, piece.element);
  }
  // the paragraph marks last, last first, so that a join never disturbs one still to do
  const marks = mine.filter((piece): piece is Extract<MovePiece, { kind: 'mark' }> => piece.kind === 'mark');
  for (let i = marks.length - 1; i >= 0; i--) {
    const mark = marks[i]!;
    const p = mark.paragraph.value;
    if (mark.side === kept) {
      const rPr = p.pPr?.rPr;
      if (rPr) delete rPr[mark.side === 'from' ? 'moveFrom' : 'moveTo'];
      pruneParagraphProperties(p);
    } else if ((p.content ?? []).length === 0) {
      remove(mark.container, mark.paragraph as Element);                 // the moved paragraph, emptied: gone, as in Word
    } else {
      const rPr = p.pPr?.rPr;
      if (rPr) delete rPr[mark.side === 'from' ? 'moveFrom' : 'moveTo'];
      joinWithNext({ container: mark.container, element: mark.paragraph, p }, mark.side === 'to');
    }
  }
  if (broken) dissolveBrokenMoves(story);
  return true;
}

/** The story a paragraph is in: the outermost array of block content above it (a body, a header, a note). */
function storyOf(p: wml.P): Element[] {
  let story: Element[] | undefined;
  for (let o: object | undefined = p; o !== undefined; o = (o as { PARENT?: object }).PARENT) {
    const content = (o as { content?: unknown }).content;
    if (Array.isArray(content) && (o as { TYPE_NAME?: string }).TYPE_NAME !== 'org_docx4j_wml.P') story = content as Element[];
  }
  return story ?? [];
}

/** Every piece of move markup in a story, in document order, each half and mark with the names of the ranges open over it. */
function movePiecesOf(story: Element[]): MovePiece[] {
  const out: MovePiece[] = [];
  const open = { from: new Map<number, string>(), to: new Map<number, string>() };
  const visit = (items: Element[], runLevel = false): void => {
    for (const el of items) {
      const marker = RANGE_MARKERS[el.name?.localPart ?? ''];
      const value = el.value as { id?: number; name?: string } | undefined;
      if (marker && value) {
        const id = Number(value.id);
        if (marker.kind === 'start') { open[marker.side].set(id, value.name ?? ''); out.push({ kind: 'start', side: marker.side, id, name: value.name ?? '', element: el, owner: items }); }
        else { open[marker.side].delete(id); out.push({ kind: 'end', side: marker.side, id, element: el, owner: items }); }
        continue;
      }
      const revision = revisionKindOf(el);
      if (revision === 'moveFrom' || revision === 'moveTo') {
        const side = revision === 'moveFrom' ? 'from' : 'to';
        out.push({ kind: 'half', side, element: el, owner: items, names: [...open[side].values()] });
      } else if (runLevel) {
        for (const side of ['from', 'to'] as const) {
          if (open[side].size > 0) out.push({ kind: 'inner', side, element: el, owner: items, names: [...open[side].values()] });
        }
      }
      const v = el.value;
      if (typeof v !== 'object' || v === null) continue;
      const children = childrenOf(v) ?? runItemsOf(v);
      if (children) visit(children, typeNameOf(el) === 'org_docx4j_wml.P');
      if (typeNameOf(el) === 'org_docx4j_wml.P') {
        const rPr = (v as wml.P).pPr?.rPr;
        for (const side of ['from', 'to'] as const) {
          if (rPr?.[side === 'from' ? 'moveFrom' : 'moveTo']) {
            out.push({ kind: 'mark', side, paragraph: el as Element<wml.P>, container: items, names: [...open[side].values()] });
          }
        }
      }
    }
  };
  visit(story);
  return out;
}

/**
 * A move half in no named range - its range markers gone - resolved as Word resolved one (check 20
 * cases 13 and 14, where Word showed it as an insertion or a deletion): a `w:moveTo` as an insertion,
 * a `w:moveFrom` as a deletion, its paragraph mark with it, so that a paragraph whose text and mark
 * both go goes whole. Then the rest of the broken pair is dissolved (`dissolveBrokenMoves`).
 */
function resolveRangelessHalf(change: TrackedChange, action: 'accept' | 'reject'): boolean {
  const t = change.target;
  const paragraph = change.paragraph;
  if (t.kind !== 'run' || (t.revision !== 'moveTo' && t.revision !== 'moveFrom') || !paragraph) return false;
  const story = storyOf(paragraph.p);
  const pieces = movePiecesOf(story);
  const half = pieces.find((piece) => piece.kind === 'half' && piece.element === t.element);
  if (!half || half.kind !== 'half' || half.names.length > 0) return false;

  const side = half.side;
  const key = side === 'from' ? 'moveFrom' : 'moveTo';
  const keep = (side === 'to') === (action === 'accept');     // an insertion accepted, a deletion rejected
  if (!keep) remove(t.owner, t.element);
  else if (side === 'to') unwrap(t);
  else restoreDeleted(t);
  const mark = pieces.find((piece) => piece.kind === 'mark' && piece.side === side && piece.paragraph === paragraph.element && piece.names.length === 0);
  const p = paragraph.p;
  if (mark && p.pPr?.rPr?.[key]) {
    if (keep) {
      delete p.pPr.rPr[key];
      pruneParagraphProperties(p);
    } else if ((p.content ?? []).length === 0) {
      remove(paragraph.container, paragraph.element);
    } else {
      delete p.pPr.rPr[key];
      joinWithNext(paragraph, side === 'to');
    }
  }
  dissolveBrokenMoves(story);
  return true;
}

/**
 * Every move half and moved paragraph mark in the story that cannot be paired - in no named range,
 * or in one whose `w:name` no range of the other side carries - turned into the plain `w:ins` or
 * `w:del` it stands for, pending, its range markers kept. That is how Word saved each broken pair
 * in check 20 (re-run, cases 12 to 14, saved unchanged) and what it left of the other side once one
 * was accepted or rejected. It runs after a broken move is resolved, and only then, so a document's
 * other moves are left as they are; but a broken pair's halves cannot be told from another's, so
 * this takes them all.
 */
function dissolveBrokenMoves(story: Element[]): void {
  const pieces = movePiecesOf(story);
  const named = { from: new Set<string>(), to: new Set<string>() };
  for (const piece of pieces) if (piece.kind === 'start') named[piece.side].add(piece.name);
  const unpaired = (side: 'from' | 'to', names: string[]): boolean => {
    const other = named[side === 'from' ? 'to' : 'from'];
    return names.every((name) => !other.has(name));          // true for no names too
  };
  for (const piece of pieces) {
    if (piece.kind === 'half' && unpaired(piece.side, piece.names)) toPlainRevision(piece.element, piece.owner, piece.side === 'from' ? 'del' : 'ins');
    if (piece.kind === 'mark' && unpaired(piece.side, piece.names)) {
      const rPr = piece.paragraph.value.pPr?.rPr;
      if (piece.side === 'from' && rPr?.moveFrom) { rPr.del = rPr.moveFrom; delete rPr.moveFrom; }
      if (piece.side === 'to' && rPr?.moveTo) { rPr.ins = rPr.moveTo; delete rPr.moveTo; }
    }
  }
}

/**
 * A `w:moveFrom` turned into the plain `w:del` of the same runs, their text become deleted text, or a
 * `w:moveTo` into the plain `w:ins`.
 */
function toPlainRevision(element: Element, owner: Element[], kind: 'ins' | 'del'): void {
  const value = element.value as unknown as wml.CTTrackChange;
  const items = runItemsOf(element.value as object) ?? [];
  if (kind === 'del') for (const item of items) if (typeNameOf(item) === 'org_docx4j_wml.R') toDeletedText(item.value as wml.R);
  const plain = {
    name: { namespaceURI: W_NS, localPart: kind },
    value: { TYPE_NAME: kind === 'del' ? 'org_docx4j_wml.RunDel' : 'org_docx4j_wml.RunIns', id: value.id, author: value.author, date: value.date, customXmlOrSmartTagOrSdt: items },
  } as unknown as Element;
  const i = owner.indexOf(element);
  if (i < 0) return;
  owner.splice(i, 1, plain);
  linkParents(items, plain.value as object);
  linkParents(plain, (value as { PARENT?: object }).PARENT);
}

// --- collecting ----------------------------------------------------------------------------

/**
 * The tracked changes of one paragraph, in document order, touching pieces grouped as Office JS
 * groups them (`groupTouching`). Office JS lists the changes that start in the paragraph (check 22):
 * pieces that carry on a change from the paragraph before - its inserted or deleted mark - are left
 * to that paragraph, and a change running on into the next is cut at this paragraph's end.
 */
export function trackedChangesOfParagraph(paragraph: Paragraph): TrackedChange[] {
  const tokens = trackedChangeTokensOfParagraph(paragraph);
  const before = markBefore(paragraph);
  if (before) {
    for (let i = 0; i < tokens.length;) {
      const token = tokens[i]!;
      if (token === BREAK) break;
      const kind = groupable(token);
      if (kind === undefined) { i++; continue; }
      if (kind !== before.kind || token.author !== before.author) break;
      tokens.splice(i, 1);
    }
  }
  return groupTouching(tokens);
}

/** The last revision on the mark of the paragraph before this one in its container, if any. */
function markBefore(paragraph: Paragraph): { kind: 'ins' | 'del' | 'format'; author: string } | undefined {
  const i = paragraph.container.indexOf(paragraph.element);
  const previous = i > 0 ? paragraph.container[i - 1] : undefined;
  if (!previous || typeNameOf(previous) !== 'org_docx4j_wml.P') return undefined;
  const rPr = (previous.value as wml.P).pPr?.rPr;
  if (rPr?.del) return { kind: 'del', author: rPr.del.author };
  if (rPr?.ins) return { kind: 'ins', author: rPr.ins.author };
  if (rPr?.rPrChange) return { kind: 'format', author: rPr.rPrChange.author };
  return undefined;
}

/** Between two tracked changes: something unrevised - text, a paragraph mark, a table - that keeps them apart. */
export const BREAK = 'break' as const;
export type TrackedChangeToken = TrackedChange | typeof BREAK;

/**
 * A paragraph's single changes in document order, its mark last, with a `BREAK` wherever unrevised
 * content stands between two of them: `groupTouching` makes the groups from these.
 */
export function trackedChangeTokensOfParagraph(paragraph: Paragraph): TrackedChangeToken[] {
  const out: TrackedChangeToken[] = [];
  const p = paragraph.p;
  const pPrChange = p.pPr?.pPrChange;
  if (pPrChange) out.push(new TrackedChange({ kind: 'paragraphProperties', value: pPrChange }, paragraph));
  tokensOfRunLevel(paragraph, (p.content ?? []) as Element[], out);
  const rPr = p.pPr?.rPr;
  // a mark whose formatting changed is a piece, grouping with its runs' (check 27: "Before.\r")
  if (rPr?.rPrChange) out.push(new TrackedChange({ kind: 'markProperties', value: rPr.rPrChange }, paragraph));
  if (rPr?.ins) out.push(new TrackedChange({ kind: 'mark', mark: 'ins', value: rPr.ins }, paragraph));
  if (rPr?.del) out.push(new TrackedChange({ kind: 'mark', mark: 'del', value: rPr.del }, paragraph));
  if (!rPr?.ins && !rPr?.del && !rPr?.rPrChange) out.push(BREAK);
  const sectPr = p.pPr?.sectPr;
  if (sectPr?.sectPrChange) {
    out.push(new TrackedChange({ kind: 'sectionProperties', sectPr, value: sectPr.sectPrChange, text: `${sectionText(paragraph.container, paragraph.container.indexOf(paragraph.element) + 1)}\f` }, paragraph));
  }
  return out;
}

/**
 * The text of the section that ends at `end` (exclusive) in a container: its paragraphs' text, a `\r`
 * between them, back to the previous section break.
 */
function sectionText(items: Element[], end: number): string {
  let from = end - 1;
  while (from > 0 && !(typeNameOf(items[from - 1]!) === 'org_docx4j_wml.P' && (items[from - 1]!.value as wml.P).pPr?.sectPr)) from--;
  return items.slice(Math.max(from, 0), end).filter((el) => typeNameOf(el) === 'org_docx4j_wml.P')
    .map((el) => textOfView(el.value as object, { view: 'accepted' })).join('\r');
}

function tokensOfRunLevel(paragraph: Paragraph, items: Element[], out: TrackedChangeToken[]): void {
  for (let i = 0; i < items.length; i++) {
    const el = items[i]!;
    // an inline control or custom XML element that is itself an insertion: a piece, then what is in it
    const inserted = holderInsertionAt(items, i, paragraph);
    if (inserted) {
      out.push(inserted);
      tokensOfRunLevel(paragraph, childrenOf(items[i + 1]!.value as object) ?? [], out);
      i += 2;
      continue;
    }
    const kind = revisionKindOf(el);
    if (kind !== undefined) {
      // moved text is content Office JS cannot list (checks 16, 17), and is kept apart
      const move = kind === 'moveFrom' || kind === 'moveTo';
      if (move) out.push(BREAK);
      out.push(new TrackedChange({ kind: 'run', revision: kind, element: el, owner: items, value: el.value as wml.CTTrackChange }, paragraph));
      if (move) out.push(BREAK);
      const nested = runItemsOf(el.value as object);
      if (nested) {
        // a revision inside a revision (a deletion of inserted text) was not measured: listed apart
        const inner: TrackedChangeToken[] = [];
        tokensOfRunLevel(paragraph, nested, inner);
        if (inner.some((token) => token !== BREAK && token.target.kind === 'run')) out.push(BREAK, ...inner, BREAK);
        else out.push(...inner.filter((token) => token !== BREAK));
      }
      continue;
    }
    if (typeNameOf(el) === 'org_docx4j_wml.R') {
      // a run whose formatting changed is a piece of its own; one that did not keeps pieces apart -
      // if it has text: a run holding only a comment reference does not (check 32: one author's two
      // w:ins with the reference run between them listed as one change)
      const run = el.value as wml.R;
      if (run.rPr?.rPrChange) out.push(new TrackedChange({ kind: 'runProperties', run, value: run.rPr.rPrChange }, paragraph));
      else if ((run.content ?? []).some((item) => itemTextOf(item as Element) !== undefined)) out.push(BREAK);
      continue;
    }
    // a run holder that is not itself an insertion - a hyperlink, a field, a control or custom XML
    // element that was there - keeps the insertions inside it apart from those outside: Office JS
    // listed "before ", "link" and " after" as three changes over a w:ins in a w:hyperlink, and the same
    // over a w:fldSimple and an unmarked inline w:sdt (check 30, part C), against one change over an
    // inserted control (above) and over Word's own one-w:ins forms
    const nested = runItemsOf(el.value as object);
    if (nested) { out.push(BREAK); tokensOfRunLevel(paragraph, nested, out); out.push(BREAK); }
  }
}

/**
 * What a piece groups as: a `w:ins` or `w:del` around runs, or an inserted or deleted mark; a run's
 * `w:rPrChange` (check 28: bold on an italic run and a plain one, one change). Moves and a
 * paragraph's `w:pPrChange` do not group.
 */
function groupable(change: TrackedChange): 'ins' | 'del' | 'format' | undefined {
  const t = change.target;
  if (t.kind === 'mark') return t.mark;
  if (t.kind === 'run' && (t.revision === 'ins' || t.revision === 'del')) return t.revision;
  if (t.kind === 'runProperties' || t.kind === 'markProperties') return 'format';
  if (t.kind === 'holder') return 'ins';
  // an inserted or deleted row groups with rows and text touching it: Office JS lists an inserted
  // table of three rows, with the inserted paragraphs before and after it, as one change, and rows
  // added to a table that was there as one change per touching run of them (check 31)
  if (t.kind === 'row') return t.row;
  return undefined;
}

/**
 * Tokens made into the changes Office JS lists: consecutive pieces of one kind (inserted, deleted, or
 * a run's formatting) by one author, nothing unrevised between them, are one change (checks 22, 23
 * and 28: text and marks alike, dates regardless, formatting whatever it recorded; another author, or
 * one kind against another, starts a new one). A change that does not group - a paragraph's
 * properties, a move - is listed where it stands and does not break a group. A group of one piece is
 * that piece.
 */
export function groupTouching(tokens: TrackedChangeToken[]): TrackedChange[] {
  const out: (TrackedChange | TrackedChange[])[] = [];
  let open: TrackedChange[] | undefined;
  for (const token of tokens) {
    if (token === BREAK) { open = undefined; continue; }
    const kind = groupable(token);
    if (kind === undefined) { out.push(token); continue; }
    const last = open?.[open.length - 1];
    if (open && last && groupable(last) === kind && last.author === token.author) { open.push(token); continue; }
    open = [token];
    out.push(open);
  }
  return out.map((item) => {
    if (!Array.isArray(item)) return item;
    const first = item[0]!;
    if (item.length === 1) return first;
    return new TrackedChange({ kind: 'group', revision: groupable(first)!, pieces: item, value: first.target.value as wml.CTTrackChange }, first.paragraph);
  });
}

/**
 * The `w:trPr/w:ins` and `w:trPr/w:del` revisions of a table row. `inner` is the revision markup
 * of the row's own cells, which the row's change owns rather than reporting separately.
 */
export function trackedChangesOfRow(tr: Element<wml.Tr>, owner: Element[], inner: TrackedChange[] = [], table?: Located<wml.Tbl>): TrackedChange[] {
  const trPr = tr.value.trPr;
  const out: TrackedChange[] = [];
  if (trPr?.ins) out.push(new TrackedChange({ kind: 'row', row: 'ins', tr, owner, value: trPr.ins, inner, table }, undefined));
  if (trPr?.del) out.push(new TrackedChange({ kind: 'row', row: 'del', tr, owner, value: trPr.del, inner, table }, undefined));
  return out;
}

// --- CR-002 section 35: tables, cells, marks and sections --------------------------------------------

/** The change a table's property records make, one per table, or undefined when they make none. */
export function tablePropertiesChangeOf(tbl: Element<wml.Tbl>): TrackedChange | undefined {
  const records = tableRecordsOf(tbl.value);
  return records ? new TrackedChange({ kind: 'tableProperties', tbl, value: records.value }, undefined) : undefined;
}

/** A cell's own revision, `w:cellIns` or `w:cellDel`, or undefined. */
export function cellChangeOf(tc: Element<wml.Tc>, owner: Element[], tr: wml.Tr): TrackedChange | undefined {
  const tcPr = tc.value.tcPr;
  if (tcPr?.cellIns) return new TrackedChange({ kind: 'cell', cell: 'ins', tc, owner, tr, value: tcPr.cellIns }, undefined);
  if (tcPr?.cellDel) return new TrackedChange({ kind: 'cell', cell: 'del', tc, owner, tr, value: tcPr.cellDel }, undefined);
  return undefined;
}

/** The change to the body's last section (`w:body/w:sectPr/w:sectPrChange`), its text the section's paragraphs'. */
export function bodySectionChangeOf(container: { content?: Element[]; sectPr?: wml.SectPr }): TrackedChange | undefined {
  const sectPr = container.sectPr;
  if (!sectPr?.sectPrChange) return undefined;
  const items = container.content ?? [];
  return new TrackedChange({ kind: 'sectionProperties', sectPr, value: sectPr.sectPrChange, text: sectionText(items, items.length) }, undefined);
}

/**
 * A cell taken out of its row, the cell before it (else after it) widened over its grid columns, as
 * Word did rejecting an inserted cell and accepting a deleted one (check 27: `w:gridSpan` 2 and the
 * two widths added).
 */
function removeCell(tc: Element<wml.Tc>, owner: Element[]): void {
  const i = owner.indexOf(tc);
  if (i < 0) return;
  const cells = owner.filter((el) => typeNameOf(el) === 'org_docx4j_wml.Tc') as Element<wml.Tc>[];
  const at = cells.indexOf(tc);
  const neighbour = cells[at - 1] ?? cells[at + 1];
  owner.splice(i, 1);
  if (!neighbour) return;
  const removed = tc.value.tcPr;
  const kept = (neighbour.value.tcPr ??= { TYPE_NAME: 'org_docx4j_wml.TcPr' } as wml.TcPr);
  const span = (tcPr: wml.TcPrInner | undefined): number => Number(tcPr?.gridSpan?.val ?? 1) || 1;
  kept.gridSpan = { ...(kept.gridSpan ?? {}), val: span(kept) + span(removed) } as wml.TcPrInner.GridSpan;
  if (kept.tcW?.type === 'dxa' && removed?.tcW?.type === 'dxa') {
    kept.tcW = { ...kept.tcW, w: Number(kept.tcW.w ?? 0) + Number(removed.tcW.w ?? 0) } as wml.TblWidth;
  }
  linkParents(kept, neighbour.value);
}

/** A mark's formatting put back as its `w:rPrChange` recorded it, its own revisions (inserted, deleted, moved) kept. */
function restoreMarkProperties(p: wml.P, change: wml.ParaRPrChange): void {
  const rPr = p.pPr?.rPr;
  if (!rPr) return;
  const keep = ['ins', 'del', 'moveFrom', 'moveTo'];
  const target = rPr as unknown as Record<string, unknown>;
  for (const key of Object.keys(target)) if (key !== 'TYPE_NAME' && key !== 'PARENT' && !keep.includes(key)) delete target[key];
  const original = rPrFromElements(change.rPr?.egrPrBase ?? []) as unknown as Record<string, unknown>;
  for (const [key, value] of Object.entries(original)) if (key !== 'TYPE_NAME' && key !== 'PARENT') target[key] = value;
  if (p.pPr) linkParents(rPr, p.pPr);
  pruneParagraphProperties(p);
}

/**
 * A section's properties put back as its `w:sectPrChange` recorded them. Word records only what
 * changed, and rejecting laid those over the section, the rest left as it stood (check 18's saved
 * file: the margins came back, the page size and columns stayed); a recorded `w:docGrid` of pitch 0
 * stands for none, and Word wrote none (check 18, and check 27's section break).
 */
function restoreSection(sectPr: wml.SectPr, change: wml.CTSectPrChange): void {
  const target = sectPr as unknown as Record<string, unknown>;
  for (const [key, value] of Object.entries(change.sectPr ?? {})) {
    if (key !== 'TYPE_NAME' && key !== 'PARENT') target[key] = deepCopy(value);
  }
  const grid = sectPr.docGrid as { linePitch?: unknown; type?: unknown; charSpace?: unknown } | undefined;
  if (grid && Number(grid.linePitch ?? 0) === 0 && grid.type === undefined && grid.charSpace === undefined) delete sectPr.docGrid;
  delete sectPr.sectPrChange;
  linkParents(sectPr, (sectPr as { PARENT?: object }).PARENT);
}

/**
 * What `acceptAll` and `rejectAll` resolve that is never listed: a `w:cellMerge`, which Office JS
 * cannot list (check 27: `getTrackedChanges` throws); a `w:numberingChange`, which Word no longer
 * writes and drops when one is put in, and whose record is the old number's text, so that a reject
 * too can only drop it; and a table's records that change nothing. A merge rejected gets the merge it
 * recorded back (docx4j's reading; unmeasured). A `w:customXmlInsRangeStart` or `End` left without
 * the form `holderInsertionAt` lists marks nothing that can be found, and goes either way.
 */
export function resolveUnlisted(items: Element[], what: 'accept' | 'reject'): void {
  const visit = (list: Element[] | undefined): void => {
    if (!list) return;
    for (let i = list.length - 1; i >= 0; i--) {
      const name = list[i]!.name?.localPart;
      if (name === 'customXmlInsRangeStart' || name === 'customXmlInsRangeEnd') list.splice(i, 1);
    }
    for (const el of list) {
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.P') {
        const numPr = (el.value as wml.P).pPr?.numPr;
        if (numPr?.numberingChange) delete numPr.numberingChange;
        sweepMarkers((el.value as wml.P).content as Element[] | undefined);
        continue;
      }
      if (tn === 'org_docx4j_wml.Tbl') {
        const tbl = el.value as wml.Tbl;
        if (hasTableRecords(tbl) && !tableRecordsOf(tbl)) acceptTable(tbl);
        for (const row of rowsOf(tbl)) {
          for (const cell of cellsOf(row.element.value)) {
            const tcPr = cell.element.value.tcPr;
            if (tcPr?.cellMerge) resolveCellMerge(tcPr, what);
            visit(childrenOf(cell.element.value));
          }
        }
        continue;
      }
      const v = el.value;
      if (typeof v === 'object' && v !== null) visit(childrenOf(v));
    }
  };
  visit(items);
}

/** Stray `w:customXmlInsRangeStart` / `End` markers taken out of run-level content, through its holders and revisions. */
function sweepMarkers(items: Element[] | undefined): void {
  if (!items) return;
  for (let i = items.length - 1; i >= 0; i--) {
    const name = items[i]!.name?.localPart;
    if (name === 'customXmlInsRangeStart' || name === 'customXmlInsRangeEnd') { items.splice(i, 1); continue; }
    const v = items[i]!.value;
    if (typeof v === 'object' && v !== null) sweepMarkers(runItemsOf(v));
  }
}

function resolveCellMerge(tcPr: wml.TcPr, what: 'accept' | 'reject'): void {
  const change = tcPr.cellMerge!;
  delete tcPr.cellMerge;
  if (what === 'accept') return;
  const orig = (change as { vMergeOrig?: string }).vMergeOrig;
  if (orig === 'rest') tcPr.vMerge = { val: 'restart' } as wml.TcPrInner.VMerge;
  else if (orig === 'cont') tcPr.vMerge = {} as wml.TcPrInner.VMerge;
  else delete tcPr.vMerge;
}
