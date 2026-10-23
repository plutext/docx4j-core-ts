// CR-012 phase 2: what Accept All or Reject All would leave of a document, worked out without resolving
// anything: docx4j's org.docx4j.model.revisions.ResolvedView (its CR-046 D2), over this package's own
// listing (`trackedChangesOfStory`) and resolution (`Body.acceptAll`, `TrackedChange.mts`), which is the
// copy route the text model's views are held equal to (WD REQ-406, the agreement test).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { type Element, typeNameOf, runItemsOf, childrenOf, revisionKindOf, itemTextOf } from './tree.mjs';
import { type TrackedChange, type MovePiece, movePiecesOf, storyOf, firstCellParagraphOf } from './TrackedChange.mjs';
import { trackedChangesOfStory, type Body } from './Body.mjs';
import { HOLDER_TYPES } from './tracking.mjs';

const valueOf = (item: unknown): unknown => (typeof item === 'object' && item !== null && 'value' in item && 'name' in item ? (item as Element).value : item);
const localName = (item: unknown): string => (typeof item === 'object' && item !== null && 'name' in item ? (item as Element).name?.localPart ?? '' : '');
const commentIdOf = (item: unknown): unknown => (valueOf(item) as { id?: unknown } | undefined)?.id;

/** The index in `list`, from `from` up to `to`, of a comment's range start or end (`local`), or -1. */
function markerAt(list: readonly Element[], from: number, to: number, local: 'commentRangeStart' | 'commentRangeEnd', id: unknown): number {
  for (let i = from; i < to; i++) if (localName(list[i]) === local && commentIdOf(list[i]) === id) return i;
  return -1;
}

/** A block-level control's content, which holds paragraphs. */
const SDT_CONTENTS: ReadonlySet<string> = new Set(['org_docx4j_wml.SdtContentBlock', 'org_docx4j_wml.CTSdtContentBlock']);

/** A marker that stands between two blocks without keeping them apart, as a join reads one. */
function isRangeMarker(item: unknown): boolean {
  const name = localName(item);
  return name.endsWith('RangeStart') || name.endsWith('RangeEnd') || name === 'bookmarkStart' || name === 'bookmarkEnd'
    || name === 'permStart' || name === 'permEnd' || name === 'proofErr';
}

/** The live list of run-level content a holder keeps, or of block content a container keeps. */
function listOf(value: unknown): Element[] | undefined {
  if (typeof value !== 'object' || value === null) return undefined;
  return runItemsOf(value) ?? childrenOf(value);
}

/**
 * What Word's Accept All or Reject All would leave of a document, as `acceptAllRevisions` and
 * `rejectAllRevisions` leave it, worked out without resolving anything: the content each removes and the
 * paragraphs each joins. A reader walks the unresolved tree and asks this view what is there, so that the
 * text it reads in a view is the text of the resolved document (WD REQ-402), and the package is untouched
 * (REQ-204). docx4j `org.docx4j.model.revisions.ResolvedView`.
 *
 * It lists each story's changes (`trackedChangesOfStory`, text boxes in their selected branch) and
 * records, change by change, what the resolution would do to the text: a revision taken away, a move's
 * halves and the text typed into a rejected move's destination, an inserted field rejected whole, a row
 * or cell taken away, an equation's run-level revision or structure, a paragraph mark taken away, and an
 * inserted content control rejected (its content left in its place). What a resolution does to
 * properties, range markers, comments' entries and parts is not recorded; the comments a resolution takes
 * whole are (CR-002 section 54).
 *
 * A snapshot over the trees it was made from: valid until they change.
 */
export class ResolvedView {
  /** List members gone, by identity of their values. */
  private readonly removed = new Set<unknown>();
  /** Content controls and custom XML elements gone, their content standing in their place. */
  private readonly unwrapped = new Set<unknown>();
  /** Paragraphs whose mark is gone. */
  private readonly marksGone = new Set<unknown>();
  /** Moved paragraphs whose mark is gone: gone themselves when nothing is left in them, else joined. */
  private readonly movedMarksGone = new Set<unknown>();
  private readonly commentCandidates: CommentCandidate[] = [];
  private readonly spanCandidates: SpanCandidate[] = [];
  private readonly commentsGone = new Set<number>();
  /** Move halves whose story has no piece for them (a text box's, whose paragraphs' story is the body around it). */
  private readonly unreplayed = new Set<TrackedChange>();

  private constructor(stories: readonly Body[], readonly accepted: boolean) {
    const changes: TrackedChange[] = [];
    for (const body of stories) changes.push(...trackedChangesOfStory(body));
    // moves first, story by story, in the resolution's order: their outcome depends on it
    const moves = new Map<Element[], TrackedChange[]>();
    for (const change of changes) this.collectMoves(change, moves);
    for (const [story, halves] of moves) for (const missing of new MoveReplay(this, story, halves).run()) this.unreplayed.add(missing);
    for (const change of changes) this.plan(change);
    for (const c of this.commentCandidates) if (c.rangeEmpty()) this.commentsGone.add(Number(c.id));
    for (const c of this.spanCandidates) if (c.rangeEmpty()) this.commentsGone.add(Number(c.id));
  }

  /** The document after Accept All, over these stories (each a `Body` over the tree a reader walks). */
  static accepted(stories: readonly Body[]): ResolvedView {
    return new ResolvedView(stories, true);
  }

  /** The document after Reject All. */
  static original(stories: readonly Body[]): ResolvedView {
    return new ResolvedView(stories, false);
  }

  private collectMoves(change: TrackedChange, moves: Map<Element[], TrackedChange[]>): void {
    const t = change.target;
    if (t.kind === 'group') { for (const piece of t.pieces) this.collectMoves(piece, moves); return; }
    if (t.kind === 'row') { for (const inner of t.inner) this.collectMoves(inner, moves); return; }
    if (t.kind !== 'run' || (t.revision !== 'moveFrom' && t.revision !== 'moveTo') || !change.paragraph) return;
    const story = storyOf(change.paragraph.p);
    let list = moves.get(story);
    if (!list) moves.set(story, (list = []));
    list.push(change);
  }

  /**
   * Whether a list member is gone in this view: a run-level revision taken away with its content, a run,
   * a row, a cell, a paragraph, an equation's run-level revision or structure. A revision that is not
   * gone is kept: its content is read as if it were not there. An element pair and its value alike.
   */
  removes(item: unknown): boolean {
    const v = valueOf(item);
    if (v === undefined || v === null) return false;
    if (this.removed.has(v)) return true;
    const tn = typeNameOf(v);
    if (tn === 'org_docx4j_wml.Tbl') return this.emptiedTable(v as wml.Tbl);
    if (tn !== 'org_docx4j_wml.P') return false;
    const p = v as wml.P;
    if (this.movedMarksGone.has(p) && !this.remainsIn(p)) return true;
    if (!this.marksGone.has(p)) return false;
    // an emptied paragraph whose mark is gone, with a table or a block control next in its list: gone (docx4j CR-043
    // section 17: check 54 case 9, the text-left probe's case 3). One with text left and a table next isn't removed but
    // moved: `joinsIntoTable`, which the walk applies
    const next = this.nextBlockAfter(p);
    return !this.remainsIn(p) && (typeNameOf(next) === 'org_docx4j_wml.Tbl' || HOLDER_TYPES.has(typeNameOf(next) ?? ''));
  }

  /**
   * The table a paragraph with text left joins into, where its mark is gone and a table is the next block in its list
   * (Word 365's Accept All, docx4j CR-043 section 17, the text-left probe's case 1: its content lands at the front of the
   * first cell's first paragraph, which keeps its own properties); undefined otherwise. docx4j `joinsIntoTable`.
   */
  joinsIntoTable(p: wml.P): wml.Tbl | undefined {
    if (!this.marksGone.has(p) || !this.remainsIn(p)) return undefined;
    const next = this.nextBlockAfter(p);
    return typeNameOf(next) === 'org_docx4j_wml.Tbl' && firstCellParagraphOf(next as wml.Tbl) ? next as wml.Tbl : undefined;
  }

  /**
   * A paragraph's properties in this view, where its mark is gone and no paragraph follows it in its list (Word 365's
   * Accept All, check 54; docx4j CR-043 section 17): a block control next gives it the control's first paragraph's
   * properties, as a join would have (case 19); a control's last paragraph keeps a bare mark, its properties gone, its
   * text left or not (case 13, the text-left probe's case 2); otherwise its own. docx4j `paragraphPropertiesOf`.
   */
  paragraphPropertiesOf(p: wml.P): wml.PPr | undefined {
    if (!this.marksGone.has(p)) return p.pPr;
    const parent = (p as { PARENT?: { TYPE_NAME?: string } }).PARENT;
    const container = listOf(parent);
    const k = container ? container.findIndex((item) => valueOf(item) === p) : -1;
    if (!container || k < 0) return p.pPr;
    const next = this.nextBlock(container, k);
    if (typeNameOf(next) === 'org_docx4j_wml.P') return p.pPr;    // a join: the walk takes the chain's last
    if (next && HOLDER_TYPES.has(typeNameOf(next) ?? '')) {
      for (const o of childrenOf(next as object) ?? []) {
        if (typeNameOf(o) === 'org_docx4j_wml.P') return (o.value as wml.P).pPr;
        if (typeNameOf(o) === 'org_docx4j_wml.Tbl') break;
      }
    }
    if (next === undefined && SDT_CONTENTS.has(parent?.TYPE_NAME ?? '')) return undefined;   // a bare mark, text or not (the probe's case 2)
    return p.pPr;
  }

  /**
   * A table whose every row this view removes is gone with them, as the resolution removes an emptied table
   * (check 54 case 6: a deletion from a heading across a whole table joins the heading with the paragraph after it).
   */
  private emptiedTable(tbl: wml.Tbl): boolean {
    let rows = false;
    for (const item of (tbl.content ?? []) as Element[]) {
      if (typeNameOf(item) === 'org_docx4j_wml.Tr') {
        rows = true;
        if (!this.removed.has(item.value)) return false;
      } else if (!isRangeMarker(item)) return false;
    }
    return rows;
  }

  /** The block after `p` in its list, past range markers and what this view removes; undefined at the end. */
  private nextBlockAfter(p: wml.P): unknown {
    const container = listOf((p as { PARENT?: unknown }).PARENT);
    const k = container ? container.findIndex((item) => valueOf(item) === p) : -1;
    return container && k >= 0 ? this.nextBlock(container, k) : undefined;
  }

  private nextBlock(container: readonly Element[], k: number): unknown {
    for (let i = k + 1; i < container.length; i++) {
      const item = container[i]!;
      if (isRangeMarker(item) || this.removes(item)) continue;
      return valueOf(item);
    }
    return undefined;
  }

  /** Whether a content control or custom XML element is gone with its content left in its place (one inserted, rejected). */
  unwraps(item: unknown): boolean {
    const v = valueOf(item);
    return v !== undefined && v !== null && this.unwrapped.has(v);
  }

  /**
   * Whether a paragraph's mark is gone in this view: a deletion accepted, an insertion rejected, a move's
   * mark on the side not kept. Such a paragraph joins the next one in its list, past range markers and
   * what this view removes, when that is a paragraph; with no paragraph next only the mark goes.
   */
  removesMark(p: wml.P | undefined): boolean {
    return p !== undefined && (this.marksGone.has(p) || this.movedMarksGone.has(p));
  }

  /**
   * Whether a comment goes in this view, entry and reference: one whose range started on a revision taken
   * away, or whose markers lie either side of one, once nothing with text is left in its range (CR-002
   * section 54).
   */
  removesComment(id: unknown): boolean {
    return id !== undefined && id !== null && this.commentsGone.has(Number(id));
  }

  // --- the plan, change by change, as the resolution resolves each ----------------------------

  private plan(change: TrackedChange): void {
    const t = change.target;
    switch (t.kind) {
      case 'run': {
        // a move is replayed, unless its story has no piece for it (a text box's, whose paragraphs' story is
        // the body around it): then the resolution takes it as the plain revision it stands for
        if ((t.revision === 'moveFrom' || t.revision === 'moveTo') && !this.unreplayed.has(change)) return;
        const kept = (t.revision === 'ins' || t.revision === 'moveTo') === this.accepted;
        if (!kept) this.removeRevision(t.element, t.owner);
        return;
      }
      case 'mark':
        if (change.paragraph && (t.mark === 'ins') !== this.accepted) this.marksGone.add(change.paragraph.p);
        return;
      case 'row':
        if ((t.row === 'ins') === this.accepted) { for (const inner of t.inner) this.plan(inner); }
        else this.removed.add(t.tr.value);
        return;
      case 'group':
        for (const piece of t.pieces) this.plan(piece);
        return;
      case 'field':
        // Word's inserted field: rejected, the whole field goes, its separator and result too
        for (const el of t.span) {
          if (revisionKindOf(el) === 'ins') { if (!this.accepted) this.removeRevision(el, t.owner); }
          else if (!this.accepted && typeNameOf(el) === 'org_docx4j_wml.R') this.removed.add(el.value);
        }
        return;
      case 'mathRun':
        if ((t.revision === 'ins') !== this.accepted) this.removed.add(t.item.value);
        return;
      case 'mathStructure':
        if ((t.revision === 'ins') !== this.accepted && t.structure) this.removed.add(t.structure.value);
        return;
      case 'cell':
        if ((t.cell === 'ins') !== this.accepted) this.removed.add(t.tc.value);
        return;
      case 'holder':
        // a content control inserted keeps its content either way; rejected, the control itself goes. One deleted
        // whole (Word's customXmlDelRange form), accepted, goes with everything in it; rejected, it stays (docx4j
        // CR-043 section 17, check 54 cases 15 and 17)
        if (t.revision === 'ins') { if (!this.accepted) this.unwrapped.add(t.holder.value); }
        else if (this.accepted) this.removed.add(t.holder.value);
        return;
      default:
        // property records change no text
    }
  }

  /**
   * `removeRevision`: the revision goes, and the comments whose range started on it, or whose markers
   * lie either side of it in its list, are candidates; so is one with a marker in its list on the
   * revision's side and the other in another paragraph (docx4j's U23).
   */
  removeRevision(item: Element, owner: Element[]): void {
    this.removed.add(item.value);
    const inner = listOf(item.value);
    const at = owner.indexOf(item);
    if (!inner || at < 0) return;
    for (const o of inner) {
      if (localName(o) !== 'commentRangeStart') continue;
      const id = commentIdOf(o);
      if (markerAt(owner, at + 1, owner.length, 'commentRangeEnd', id) >= 0) this.commentCandidates.push(new CommentCandidate(this, id, owner, at));
      else this.spanCandidates.push(new SpanCandidate(this, id, item, owner, at, true));
    }
    for (let s = 0; s < at; s++) {
      if (localName(owner[s]) !== 'commentRangeStart') continue;
      const id = commentIdOf(owner[s]);
      if (markerAt(owner, s + 1, owner.length, 'commentRangeEnd', id) < 0) this.spanCandidates.push(new SpanCandidate(this, id, item, owner, s + 1, true));
      else if (markerAt(owner, at + 1, owner.length, 'commentRangeEnd', id) >= 0) this.commentCandidates.push(new CommentCandidate(this, id, owner, s));
    }
    for (let e = at + 1; e < owner.length; e++) {
      if (localName(owner[e]) === 'commentRangeEnd' && markerAt(owner, 0, e, 'commentRangeStart', commentIdOf(owner[e])) < 0) {
        this.spanCandidates.push(new SpanCandidate(this, commentIdOf(owner[e]), item, owner, e, false));
      }
    }
  }

  /** Records a list member gone (a move's half, the text typed into a rejected move's destination). */
  remove(value: unknown): void {
    this.removed.add(value);
  }

  /** Records a paragraph mark gone. */
  markGone(p: wml.P): void {
    this.marksGone.add(p);
  }

  /** Records a moved paragraph's mark gone. */
  movedMarkGone(p: wml.P): void {
    this.movedMarksGone.add(p);
  }

  /** Whether run-level content has text in this view, once its removals are made. */
  hasText(item: unknown): boolean {
    const v = valueOf(item);
    if (v === undefined || v === null || this.removed.has(v)) return false;
    if (typeNameOf(v) === 'org_docx4j_wml.R') return ((v as wml.R).content ?? []).some((o) => itemTextOf(o as Element) !== undefined);
    const items = listOf(v);
    return items !== undefined && items.some((o) => this.hasText(o));
  }

  /**
   * Whether anything is left in a paragraph once this view's removals are made: the resolution removes
   * an emptied moved paragraph and joins one that is not. A comment's range markers inside a revision
   * taken away stay behind it.
   */
  private remainsIn(p: wml.P): boolean {
    for (const item of (p.content ?? []) as Element[]) {
      const v = item.value;
      if (!this.removed.has(v)) return true;
      const inner = listOf(v);
      if (inner && revisionKindOf(item) !== undefined && inner.some((o) => localName(o) === 'commentRangeStart' || localName(o) === 'commentRangeEnd')) return true;
    }
    return false;
  }
}

/** A comment whose range held a revision taken away, and where its range's text begins to be looked for. */
class CommentCandidate {
  constructor(private readonly view: ResolvedView, readonly id: unknown, private readonly owner: Element[], private readonly at: number) {}

  /** `rangeIsEmpty`, on what this view leaves: nothing with text before the range's end, in the same list. */
  rangeEmpty(): boolean {
    for (let i = this.at + 1; i < this.owner.length; i++) {
      const item = this.owner[i]!;
      if (localName(item) === 'commentRangeEnd' && commentIdOf(item) === this.id) return true;
      if (this.view.hasText(item)) return false;
    }
    return false;                                                     // the end is elsewhere: text may be
  }
}

/**
 * A comment whose range held a revision taken away, its start and end in different paragraphs of one
 * list: gone, as `settleSpanningComments` judges it once the joins are made, when every paragraph from
 * the start's up to the end's loses its mark (so that they join), with nothing between them but range
 * markers and what this view removes, and nothing with text in the range.
 */
class SpanCandidate {
  constructor(private readonly view: ResolvedView, readonly id: unknown, private readonly revision: Element,
    private readonly owner: Element[], private readonly at: number, private readonly startHere: boolean) {}

  rangeEmpty(): boolean {
    const here = (this.revision.value as { PARENT?: object }).PARENT;
    if (typeNameOf(here) !== 'org_docx4j_wml.P' || (here as wml.P).content !== (this.owner as unknown)) return false;
    const container = listOf((here as { PARENT?: object }).PARENT);
    if (!container) return false;
    let k = container.findIndex((el) => el.value === here);
    if (k < 0) return false;
    const view = this.view;
    const contentOf = (p: wml.P): Element[] => (p.content ?? []) as Element[];
    if (this.startHere) {
      for (let i = this.at; i < this.owner.length; i++) if (view.hasText(this.owner[i])) return false;
      let p = here as wml.P;
      for (;;) {
        if (!view.removesMark(p)) return false;                       // a mark left between start and end (U24)
        k++;
        while (k < container.length && this.skipped(container[k]!)) {
          if (this.isMarker(container[k]!, 'commentRangeEnd')) return true;
          k++;
        }
        if (k >= container.length || typeNameOf(container[k]) !== 'org_docx4j_wml.P') return false;
        p = container[k]!.value as wml.P;
        const items = contentOf(p);
        const end = markerAt(items, 0, items.length, 'commentRangeEnd', this.id);
        const to = end < 0 ? items.length : end;
        for (let i = 0; i < to; i++) if (view.hasText(items[i])) return false;
        if (end >= 0) return true;
      }
    }
    for (let i = 0; i < this.at; i++) if (view.hasText(this.owner[i])) return false;
    for (;;) {
      k--;
      while (k >= 0 && this.skipped(container[k]!)) {
        if (this.isMarker(container[k]!, 'commentRangeStart')) return true;
        k--;
      }
      if (k < 0 || typeNameOf(container[k]) !== 'org_docx4j_wml.P') return false;
      const p = container[k]!.value as wml.P;
      if (!view.removesMark(p)) return false;                         // a mark left between start and end (U24)
      const items = contentOf(p);
      const start = markerAt(items, 0, items.length, 'commentRangeStart', this.id);
      for (let i = start + 1; i < items.length; i++) if (view.hasText(items[i])) return false;
      if (start >= 0) return true;
    }
  }

  /** A range marker between paragraphs, which a join carries along, or a member this view removes. */
  private skipped(item: Element): boolean {
    const name = localName(item);
    return name.endsWith('RangeStart') || name.endsWith('RangeEnd') || name === 'bookmarkStart' || name === 'bookmarkEnd'
      || name === 'permStart' || name === 'permEnd' || name === 'proofErr'
      || (typeNameOf(item) !== 'org_docx4j_wml.P' && this.view.removes(item));
  }

  private isMarker(item: Element, kind: 'commentRangeStart' | 'commentRangeEnd'): boolean {
    return localName(item) === kind && commentIdOf(item) === this.id;
  }
}

/**
 * The moves of one story resolved as `Body.acceptAll` and `rejectAll` resolve them (`resolveMove`,
 * `resolveRangelessHalf`, `dissolveBrokenMoves`): last half first; a half in a named range resolves its
 * whole move, a half in none is an insertion or a deletion, and after a broken move what is left unpaired
 * becomes the plain `w:ins` or `w:del` it stands for, resolved as such. docx4j `ResolvedView.MoveReplay`.
 */
class MoveReplay {
  private readonly pieces: MovePiece[];
  /** Pieces resolved, or turned plain: no longer move markup. */
  private readonly done = new Set<MovePiece>();
  /** Halves and marks turned into plain revisions. */
  private readonly plain: MovePiece[] = [];

  constructor(private readonly view: ResolvedView, story: Element[], private readonly halves: TrackedChange[]) {
    this.pieces = movePiecesOf(story);
  }

  /** Replays the moves; returns the halves the story has no piece for. */
  run(): TrackedChange[] {
    const missing: TrackedChange[] = [];
    for (let i = this.halves.length - 1; i >= 0; i--) {
      const change = this.halves[i]!;
      const t = change.target;
      if (t.kind !== 'run') continue;
      const half = this.pieces.find((piece) => piece.kind === 'half' && piece.element === t.element);
      if (!half || half.kind !== 'half') { missing.push(change); continue; }
      if (this.done.has(half)) continue;
      if (half.names.length > 0) this.named(half);
      else this.rangeless(half, change);
    }
    // what became plain is resolved as plain markup, in the next round
    for (const piece of this.plain) {
      const ins = piece.side === 'to';
      if (ins === this.view.accepted) continue;
      if (piece.kind === 'half') this.view.removeRevision(piece.element, piece.owner);
      else if (piece.kind === 'mark') this.view.markGone(piece.paragraph.value);
    }
    return missing;
  }

  private live(piece: MovePiece): boolean {
    return !this.done.has(piece);
  }

  /** `resolveMove`. */
  private named(half: Extract<MovePiece, { kind: 'half' }>): void {
    const name = half.names[0]!;
    const ids = new Set<number>();
    for (const piece of this.pieces) if (this.live(piece) && piece.kind === 'start' && piece.name === name) ids.add(piece.id);
    const mine = this.pieces.filter((piece) => this.live(piece) && (
      (piece.kind === 'start' && piece.name === name) || (piece.kind === 'end' && ids.has(piece.id))
      || ((piece.kind === 'half' || piece.kind === 'mark' || piece.kind === 'inner') && piece.names.includes(name))));
    const orphan = !mine.some((piece) => piece.side === 'from');
    const keptFrom = !(this.view.accepted || orphan);
    const broken = orphan || !mine.some((piece) => piece.side === 'to');
    for (const piece of mine) {
      this.done.add(piece);
      if (piece.kind === 'inner' && piece.side === 'to' && keptFrom) this.view.remove(piece.element.value);
      if (piece.kind === 'half' && (piece.side === 'from') !== keptFrom) this.view.remove(piece.element.value);
      if (piece.kind === 'mark' && (piece.side === 'from') !== keptFrom) this.view.movedMarkGone(piece.paragraph.value);
    }
    if (broken) this.dissolveBroken();
  }

  /** `resolveRangelessHalf`. */
  private rangeless(half: Extract<MovePiece, { kind: 'half' }>, change: TrackedChange): void {
    const keep = (half.side === 'to') === this.view.accepted;
    this.done.add(half);
    if (!keep) this.view.remove(half.element.value);
    const paragraph = change.paragraph;
    if (paragraph) {
      for (const piece of this.pieces) {
        if (!this.live(piece) || piece.kind !== 'mark' || piece.side !== half.side || piece.paragraph !== paragraph.element || piece.names.length > 0) continue;
        this.done.add(piece);
        if (!keep) this.view.movedMarkGone(piece.paragraph.value);
        break;
      }
    }
    this.dissolveBroken();
  }

  /** `dissolveBrokenMoves`: every half and mark that cannot be paired becomes plain. */
  private dissolveBroken(): void {
    const named = { from: new Set<string>(), to: new Set<string>() };
    for (const piece of this.pieces) if (this.live(piece) && piece.kind === 'start') named[piece.side].add(piece.name);
    for (const piece of this.pieces) {
      if (!this.live(piece) || (piece.kind !== 'half' && piece.kind !== 'mark')) continue;
      const other = named[piece.side === 'from' ? 'to' : 'from'];
      if (piece.names.some((n) => other.has(n))) continue;
      this.done.add(piece);
      this.plain.push(piece);
    }
  }
}
