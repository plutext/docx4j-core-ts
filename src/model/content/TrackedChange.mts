// A subset of Office JS `Word.TrackedChange` over Word's revision markup (ECMA-376 17.13.5).
// Accepting and rejecting follow docx4j's AcceptTrackedChanges (docx4j-core
// org.docx4j.convert.out.common.preprocess): a w:ins is unwrapped, a w:del removed, a deleted
// paragraph mark joins its paragraph with the next. CR-002 phase F.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { textOf } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { type Element, type RevisionKind, typeNameOf, runItemsOf, revisionKindOf, linkParents, segmentsOf, textOfView, cellsOf, childrenOf, W_NS } from './tree.mjs';
import { Range } from './Range.mjs';
import type { Paragraph } from './Paragraph.mjs';
import { rPrFromElements } from '@docx4j/generated-objects-ts/builders/wml';
import { revisionDateOf, toRestoredText, toDeletedText, restoreRPr, restorePPr, pruneParagraphProperties } from './tracking.mjs';

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
   * one change per row. Accepting or rejecting the row applies them too.
   */
  | { kind: 'row'; row: 'ins' | 'del'; tr: Element<wml.Tr>; owner: Element[]; value: wml.CTTrackChange; inner: TrackedChange[] }
  /**
   * Touching insertions or deletions by one author - `w:ins` or `w:del` around runs and inserted or
   * deleted paragraph marks, with nothing unrevised between them - which Office JS lists as one
   * change (CR-002 section 29, checks 22 and 23): a deleted paragraph's text and mark, a deletion
   * across a mark. `pieces` are the single changes, in document order; `value` is the first's.
   */
  | { kind: 'group'; revision: 'ins' | 'del'; pieces: TrackedChange[]; value: wml.CTTrackChange };

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
      case 'group': return t.revision === 'ins' ? 'Added' : 'Deleted';
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
    return t.kind === 'run' ? t.element : t.kind === 'row' ? t.tr : t.kind === 'group' ? t.pieces[0]!.element : t.value;
  }

  /** The text the change covers; `\r` for a paragraph mark, as Office JS gives it (check 22). */
  get text(): string {
    const t = this.target;
    switch (t.kind) {
      case 'mark': return '\r';
      case 'group': return t.pieces.map((piece) => piece.text).join('');
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
      const here = t.pieces.filter((piece) => piece.target.kind === 'run' && piece.paragraph?.element === p.element).map((piece) => piece.getRange()!);
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
        if (t.revision === 'ins' || t.revision === 'moveTo') unwrap(t); else remove(t.owner, t.element);
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
        if (t.row === 'ins') { delete t.tr.value.trPr?.ins; applyInner(t.inner, 'accept'); }
        else remove(t.owner, t.tr as Element);
        return;
      case 'group':
        resolvePieces(t.pieces, 'accept');
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
        if (t.revision === 'ins' || t.revision === 'moveTo') remove(t.owner, t.element); else restoreDeleted(t);
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
        return;
      }
      case 'row':
        // a rejected deletion keeps the row, so its content comes back too (w:delText to w:t,
        // the w:del unwrapped, the deleted marks dropped); a rejected insertion takes it away
        if (t.row === 'del') { delete t.tr.value.trPr?.del; applyInner(t.inner, 'reject'); }
        else remove(t.owner, t.tr as Element);
        return;
      case 'group':
        resolvePieces(t.pieces, 'reject');
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
  const lines: string[] = [];
  const visit = (items: Element[] | undefined): void => {
    for (const el of items ?? []) {
      if (typeNameOf(el) === 'org_docx4j_wml.P') { lines.push(textOfView(el.value as object, { view })); continue; }
      const v = el.value;
      if (typeof v === 'object' && v !== null) visit(childrenOf(v));
    }
  };
  for (const cell of cellsOf(tr)) visit(childrenOf(cell.element.value));
  return lines.join('\n');
}

/** The changes inside a row, last first, so that a paragraph join never disturbs one still to do. */
function applyInner(inner: TrackedChange[], what: 'accept' | 'reject'): void {
  for (let i = inner.length - 1; i >= 0; i--) inner[i]![what]();
}

/**
 * Word records a list paragraph's indent in `w:pPrChange` even where the list level gives it, and on
 * reject writes no indent back: it restored "The first numbered item." to its numbered list with no
 * `w:ind`, where the change recorded `left=720 hanging=360`, the level's own (`test/README.md` check
 * 18, `reject-all.docx`, section 9; CR-002 section 29). So a restored `w:ind` equal to the indent of
 * the level the restored `w:numPr` names is dropped, leaving the level to give it. The numbering part
 * must have been read (`await pkg.getBody()` or `getPropertyResolver()` reads it); when it has not,
 * the recorded indent is restored as it was.
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
  if (!level || !sameIndent(ind, level)) return;
  delete pPr.ind;
  pruneParagraphProperties(paragraph.p);
}

/** Two `w:ind` with the same attributes and values. */
function sameIndent(a: wml.PPrBase.Ind, b: wml.PPrBase.Ind): boolean {
  const own = (ind: wml.PPrBase.Ind): string[] => Object.entries(ind)
    .filter(([key, value]) => key !== 'TYPE_NAME' && key !== 'PARENT' && value !== undefined)
    .map(([key, value]) => `${key}=${String(value)}`).sort();
  return own(a).join(' ') === own(b).join(' ');
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

function remove(owner: Element[], element: Element): void {
  const i = owner.indexOf(element);
  if (i >= 0) owner.splice(i, 1);
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
  const container = paragraph.container;
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
  container.splice(i + 1, 1);
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
function markBefore(paragraph: Paragraph): { kind: 'ins' | 'del'; author: string } | undefined {
  const i = paragraph.container.indexOf(paragraph.element);
  const previous = i > 0 ? paragraph.container[i - 1] : undefined;
  if (!previous || typeNameOf(previous) !== 'org_docx4j_wml.P') return undefined;
  const rPr = (previous.value as wml.P).pPr?.rPr;
  if (rPr?.del) return { kind: 'del', author: rPr.del.author };
  if (rPr?.ins) return { kind: 'ins', author: rPr.ins.author };
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
  if (rPr?.ins) out.push(new TrackedChange({ kind: 'mark', mark: 'ins', value: rPr.ins }, paragraph));
  if (rPr?.del) out.push(new TrackedChange({ kind: 'mark', mark: 'del', value: rPr.del }, paragraph));
  if (!rPr?.ins && !rPr?.del) out.push(BREAK);
  return out;
}

function tokensOfRunLevel(paragraph: Paragraph, items: Element[], out: TrackedChangeToken[]): void {
  for (const el of items) {
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
      const run = el.value as wml.R;
      if ((run.content ?? []).length > 0) out.push(BREAK);
      if (run.rPr?.rPrChange) out.push(new TrackedChange({ kind: 'runProperties', run, value: run.rPr.rPrChange }, paragraph));
      continue;
    }
    const nested = runItemsOf(el.value as object);
    if (nested) tokensOfRunLevel(paragraph, nested, out);
  }
}

/** What a piece groups as: a `w:ins` or `w:del` around runs, or an inserted or deleted mark. Moves and formatting do not group. */
function groupable(change: TrackedChange): 'ins' | 'del' | undefined {
  const t = change.target;
  if (t.kind === 'mark') return t.mark;
  if (t.kind === 'run' && (t.revision === 'ins' || t.revision === 'del')) return t.revision;
  return undefined;
}

/**
 * Tokens made into the changes Office JS lists: consecutive pieces of one kind (inserted or deleted)
 * by one author, nothing unrevised between them, are one change (checks 22 and 23: text and marks
 * alike, dates regardless; another author, or an insertion against a deletion, starts a new one). A
 * change that does not group - formatting, a move - is listed where it stands and does not break a
 * group. A group of one piece is that piece.
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
export function trackedChangesOfRow(tr: Element<wml.Tr>, owner: Element[], inner: TrackedChange[] = []): TrackedChange[] {
  const trPr = tr.value.trPr;
  const out: TrackedChange[] = [];
  if (trPr?.ins) out.push(new TrackedChange({ kind: 'row', row: 'ins', tr, owner, value: trPr.ins, inner }, undefined));
  if (trPr?.del) out.push(new TrackedChange({ kind: 'row', row: 'del', tr, owner, value: trPr.del, inner }, undefined));
  return out;
}
