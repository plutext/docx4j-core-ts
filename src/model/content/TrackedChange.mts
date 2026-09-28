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
import { dateOf, toRestoredText, toDeletedText, restoreRPr, restorePPr, pruneParagraphProperties } from './tracking.mjs';

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
  | { kind: 'row'; row: 'ins' | 'del'; tr: Element<wml.Tr>; owner: Element[]; value: wml.CTTrackChange; inner: TrackedChange[] };

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
      default: return 'Formatted';
    }
  }

  get author(): string {
    return this.target.value.author;
  }

  /** `w:date`, when the markup carries one. */
  get date(): Date | undefined {
    return dateOf(this.target.value.date);
  }

  /** The annotation id (`w:id`; an extension). */
  get id(): number {
    return this.target.value.id;
  }

  /** The markup this change is over (an extension): the `w:ins` element, the `w:rPrChange`, ... */
  get element(): object {
    return this.target.kind === 'run' ? this.target.element : this.target.kind === 'row' ? this.target.tr : this.target.value;
  }

  /** The text the change covers; '' for a paragraph mark, whose range is the paragraph. */
  get text(): string {
    const t = this.target;
    switch (t.kind) {
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
        if ((t.revision === 'moveFrom' || t.revision === 'moveTo') && resolveMove(this, 'accept')) return;
        if (t.revision === 'moveTo' && resolveRangelessDestination(this, 'accept')) return;
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
        if ((t.revision === 'moveFrom' || t.revision === 'moveTo') && resolveMove(this, 'reject')) return;
        if (t.revision === 'moveTo' && resolveRangelessDestination(this, 'reject')) return;
        if (t.revision === 'ins' || t.revision === 'moveTo') remove(t.owner, t.element); else restoreDeleted(t);
        return;
      case 'mark':
        if (t.mark === 'del') dropMark(this.paragraph, 'del'); else joinWithNext(this.requireParagraph(), true);
        return;
      case 'runProperties':
        restoreRPr(t.run, rPrFromElements(t.value.rPr));
        return;
      case 'paragraphProperties': {
        const p = this.requireParagraph().p;
        restorePPr(p, t.value.pPr);
        return;
      }
      case 'row':
        // a rejected deletion keeps the row, so its content comes back too (w:delText to w:t,
        // the w:del unwrapped, the deleted marks dropped); a rejected insertion takes it away
        if (t.row === 'del') { delete t.tr.value.trPr?.del; applyInner(t.inner, 'reject'); }
        else remove(t.owner, t.tr as Element);
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
  const pieces = movePiecesOf(storyOf(change.paragraph.p));
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
 * A `w:moveTo` in no named range - its range markers gone - resolved as Word resolved one (check 20
 * case 13, where Word showed it as an insertion): kept plain when accepted, taken away when rejected,
 * its paragraph mark with it. And the other side of the broken pair - a source range whose `w:name`
 * no destination range carries - is turned into a plain deletion, pending, its range markers kept,
 * as Word left it either way. That second part is read from one case; Word's own rule may be that it
 * dissolves any broken pair when it saves (case 14), which this does not attempt.
 */
function resolveRangelessDestination(change: TrackedChange, action: 'accept' | 'reject'): boolean {
  const t = change.target;
  const paragraph = change.paragraph;
  if (t.kind !== 'run' || t.revision !== 'moveTo' || !paragraph) return false;
  const pieces = movePiecesOf(storyOf(paragraph.p));
  const half = pieces.find((piece) => piece.kind === 'half' && piece.element === t.element);
  if (!half || half.kind !== 'half' || half.names.length > 0) return false;

  if (action === 'accept') unwrap(t); else remove(t.owner, t.element);
  const mark = pieces.find((piece) => piece.kind === 'mark' && piece.side === 'to' && piece.paragraph === paragraph.element && piece.names.length === 0);
  const p = paragraph.p;
  if (mark && p.pPr?.rPr?.moveTo) {
    if (action === 'accept' || (p.content ?? []).length > 0) {
      delete p.pPr.rPr.moveTo;
      if (action === 'accept') pruneParagraphProperties(p);
      else joinWithNext(paragraph, true);
    } else {
      remove(paragraph.container, paragraph.element);
    }
  }
  const destinations = new Set(pieces.filter((piece) => piece.kind === 'start' && piece.side === 'to').map((piece) => (piece as { name: string }).name));
  const orphaned = (names: string[]): boolean => names.length > 0 && names.every((name) => !destinations.has(name));
  for (const piece of pieces) {
    if (piece.kind === 'half' && piece.side === 'from' && orphaned(piece.names)) toDeletion(piece.element, piece.owner);
    if (piece.kind === 'mark' && piece.side === 'from' && orphaned(piece.names)) {
      const rPr = piece.paragraph.value.pPr?.rPr;
      if (rPr?.moveFrom) { rPr.del = rPr.moveFrom; delete rPr.moveFrom; }
    }
  }
  return true;
}

/** A `w:moveFrom` turned into the plain `w:del` of the same runs, their text become deleted text. */
function toDeletion(element: Element, owner: Element[]): void {
  const value = element.value as unknown as wml.CTTrackChange;
  const items = runItemsOf(element.value as object) ?? [];
  for (const item of items) if (typeNameOf(item) === 'org_docx4j_wml.R') toDeletedText(item.value as wml.R);
  const del = {
    name: { namespaceURI: W_NS, localPart: 'del' },
    value: { TYPE_NAME: 'org_docx4j_wml.RunDel', id: value.id, author: value.author, date: value.date, customXmlOrSmartTagOrSdt: items },
  } as unknown as Element;
  const i = owner.indexOf(element);
  if (i < 0) return;
  owner.splice(i, 1, del);
  linkParents(items, del.value as object);
  linkParents(del, (value as { PARENT?: object }).PARENT);
}

// --- collecting ----------------------------------------------------------------------------

/** The tracked changes of one paragraph, in document order. */
export function trackedChangesOfParagraph(paragraph: Paragraph): TrackedChange[] {
  const out: TrackedChange[] = [];
  const p = paragraph.p;
  const pPrChange = p.pPr?.pPrChange;
  if (pPrChange) out.push(new TrackedChange({ kind: 'paragraphProperties', value: pPrChange }, paragraph));
  const rPr = p.pPr?.rPr;
  if (rPr?.ins) out.push(new TrackedChange({ kind: 'mark', mark: 'ins', value: rPr.ins }, paragraph));
  if (rPr?.del) out.push(new TrackedChange({ kind: 'mark', mark: 'del', value: rPr.del }, paragraph));
  collectRunLevel(paragraph, (p.content ?? []) as Element[], out);
  return out;
}

function collectRunLevel(paragraph: Paragraph, items: Element[], out: TrackedChange[]): void {
  for (const el of items) {
    const kind = revisionKindOf(el);
    if (kind !== undefined) {
      out.push(new TrackedChange({ kind: 'run', revision: kind, element: el, owner: items, value: el.value as wml.CTTrackChange }, paragraph));
      const nested = runItemsOf(el.value as object);
      if (nested) collectRunLevel(paragraph, nested, out);
      continue;
    }
    if (typeNameOf(el) === 'org_docx4j_wml.R') {
      const run = el.value as wml.R;
      if (run.rPr?.rPrChange) out.push(new TrackedChange({ kind: 'runProperties', run, value: run.rPr.rPrChange }, paragraph));
      continue;
    }
    const nested = runItemsOf(el.value as object);
    if (nested) collectRunLevel(paragraph, nested, out);
  }
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
