// Comments in the shape of Office JS `Word.Comment`, over the four parts a current Word writes
// (CR-002 section 3.8, phase G). The part plumbing is in `comments.mts` (and its implementation
// in `src/parts/wml/comments.mts`), so that nothing here imports a part at runtime.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import * as wmlFactory from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import * as w15Factory from '@docx4j/generated-objects-ts/factory/org_docx4j_w15';
import * as w16cidFactory from '@docx4j/generated-objects-ts/factory/org_docx4j_w16cid';
import { r as textRun, applyRunOptions, linkParents } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { Body } from './Body.mjs';
import { Paragraph } from './Paragraph.mjs';
import { Range } from './Range.mjs';
import { runOf, paragraphOf, typeNameOf, selectedBranchOf, BRANCH_TYPES, type Element, type TextSegment } from './tree.mjs';
import { storyLayout } from './story.mjs';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import { revisionDateOf, wordCalendarOf, copyRPr } from './tracking.mjs';
import { relationshipIdsIn, queueRelationshipRemovals, type RelationshipRemovals } from './relationships.mjs';
import {
  COMMENT_REFERENCE_STYLE, COMMENT_TEXT_STYLE, type CommentContent, type CommentOptions, type CommentEntry,
  commentExOf, commentIdsOf, commentPartsAccess, commentParaId,
  initialsOf, isDone, markersOf, markersOfParagraph, nextCommentId, paraIdTaken, randomHexId, removeMarkers,
  setCommentApi, type CommentMarker, type CommentParts,
} from './comments.mjs';

/**
 * A subset of Office JS `Word.Comment` over a `w:comment` and the markers that anchor it. A view:
 * the tree is the state, and a view may be created and discarded freely. Replies are `Comment`s
 * too (Office JS has a separate `CommentReply`; see the CR's section 9), threaded through
 * `w15:commentsEx`.
 */
export class Comment {
  private parentComment: Comment | undefined;
  private readonly replyList: Comment[] = [];
  private bodyView: Body | undefined;

  constructor(
    /** The `w:comment` (extension; the model holds comments as values, not `{ name, value }` pairs). */
    readonly element: wml.Comments.Comment,
    /** The document's comment parts. */
    readonly parts: CommentParts,
    /** The body the comment is anchored in. */
    readonly body: Body,
  ) {}

  /** `w:id` (extension: Office JS's `id` is a string; this is the OOXML id docx4j users work with). */
  get id(): number {
    return this.element.id;
  }

  get authorName(): string {
    return this.element.author ?? '';
  }

  /** `w:initials` (extension; Office JS has no initials). */
  get initials(): string {
    return this.element.initials ?? '';
  }

  /** The author's email from `w:people`, '' when the document records none. */
  get authorEmail(): string {
    const person = (this.parts.people?.person ?? []).find((p) => p.author === this.authorName);
    const userId = person?.presenceInfo?.userId ?? '';
    const ad = /^S::([^:]+)::/.exec(userId);
    return ad ? ad[1]! : userId;
  }

  /**
   * When the comment was made, as Word means it (CR-002 section 29, the revision dates' rule): its
   * `w16cex:dateUtc` when the document has one for it, which is UTC; else `w:date`, which Word writes
   * on the author's local wall clock whatever its `Z` says, read as local time. Undefined when the
   * comment carries neither.
   */
  get creationDate(): Date | undefined {
    for (const el of [...(this.element.content ?? [])].reverse()) {
      const paraId = typeNameOf(el) === 'org_docx4j_wml.P' ? (el.value as wml.P).paraId : undefined;
      const utc = paraId ? this.parts.datesUtc.get(paraId.toUpperCase()) : undefined;
      if (utc) return utc;
    }
    return revisionDateOf({ date: this.element.date });
  }

  /** The comment's paragraphs as a `Body` (extension: the same content API as the document's). */
  get commentBody(): Body {
    return (this.bodyView ??= new Body(this.parts.commentsPart, this.element as { content?: Element[] }, `comment:${this.id}`, this.body.package_));
  }

  /** The comment's paragraphs (extension). */
  get paragraphs(): Paragraph[] {
    return this.commentBody.paragraphs;
  }

  /** The comment's text, a paragraph per line; setting it replaces the paragraphs. */
  get content(): string {
    // Office JS's comment.content: its body's text, the paragraphs a mark \r apart (check 51)
    return this.commentBody.text;
  }
  set content(text: string) {
    const paraId = commentParaId(this.element);
    // a paragraph per line, \r (as the getter gives it) or \n
    const paragraphs = text.split(/\r\n|\r|\n/).map((line, i) => commentParagraph(line, i === 0));
    if (paraId) paragraphs[0]!.value.paraId = paraId;
    this.element.content = paragraphs as wml.Comments.Comment['content'];
    linkParents(this.element.content, this.element);
    this.bodyView = undefined;
  }

  /** `w15:done`: whether the thread is resolved. Setting it creates the w15 part when the document has none. */
  get resolved(): boolean {
    return isDone(commentExOf(this.parts, this.paraId)?.done);
  }
  set resolved(value: boolean) {
    const extended = this.parts.ensureCommentsEx();
    const paraId = this.ensureParaId();
    const entries = (extended.commentEx ??= []);
    let entry = entries.find((e) => e.paraId === paraId);
    if (!entry) {
      entry = w15Factory.createCTCommentEx({ paraId });
      entries.push(entry);
      linkParents(entry, extended);
    }
    entry.done = value ? '1' : '0';
  }

  /** The `w14:paraId` of the comment's first paragraph: the key of its w15 and w16cid entries (extension). */
  get paraId(): string | undefined {
    return commentParaId(this.element);
  }

  /** The replies to this comment, in document order; a reply's own replies nest further (Word shows one level). */
  get replies(): Comment[] {
    return this.replyList;
  }

  /** The comment this one replies to (extension). */
  get parent(): Comment | undefined {
    return this.parentComment;
  }

  /**
   * A reply in the same thread, anchored to the same range, as Word writes them. `options` are
   * `insertComment`'s (CR-002 section 28.3).
   */
  async reply(content: CommentContent, options?: CommentOptions): Promise<Comment> {
    const paragraphs = paragraphsOfContent(content);
    const parts = await this.parts.ensureAll();
    const id = commentId(parts, this.body, options);
    placeAfter(this.body, this.id, id);
    const reply = addComment(parts, this.body, id, paragraphs, this, options);
    this.replyList.push(reply);
    reply.parentComment = this;
    return reply;
  }

  /** Removes the comment and its replies: the markers, the reference runs and every side-part entry. */
  async delete(): Promise<void> {
    await this.parts.ensureAll();
    this.removeThread();
  }

  /**
   * @internal `delete()` once the side parts it touches are loaded: public so that a resolution's
   * removals (`flushPendingRemovals`), which load only the parts the document has, take a comment away
   * without creating any (CR-002 section 54).
   */
  removeThread(): void {
    for (const comment of this.thread()) comment.removeSelf();
    if (this.parentComment) {
      const i = this.parentComment.replyList.indexOf(this);
      if (i >= 0) this.parentComment.replyList.splice(i, 1);
      this.parentComment = undefined;
    }
  }

  /**
   * The commented range: one `Range` from the comment's start to its end, across paragraphs where the comment
   * spans them, as Office JS's `Comment.getRange()` (CR-014 phase 1, D2; check 52: a comment from one paragraph
   * into the next reads `BODY1B\rLIST1A`). A start marker between blocks (Word writes one before a control's
   * first paragraph) starts the range at the next paragraph, and an end marker between blocks ends it at the
   * previous one's end. Undefined when the comment has no markers in this body. `getRanges()` gives the range a
   * paragraph at a time, as this method did before.
   */
  getRange(): Range | undefined {
    const markers = markersOf(this.body.container, (p) => this.body.hiddenRuns(p)).filter((m) => m.id === this.id);
    const start = markers.find((m) => m.kind === 'start');
    const end = [...markers].reverse().find((m) => m.kind === 'end');
    const from = start && this.anchor(start, 'start');
    const to = end && this.anchor(end, 'end');
    if (!from || !to) return this.getRanges()[0];
    const paragraphs = this.body.paragraphs;
    const a = paragraphs.findIndex((p) => p.element === from.paragraph.element);
    const b = paragraphs.findIndex((p) => p.element === to.paragraph.element);
    if (a < 0 || b < 0 || b < a || (a === b && to.offset < from.offset)) return this.getRanges()[0];
    return a === b ? new Range(from.paragraph, from.offset, to.offset)
      : new Range(from.paragraph, from.offset, to.offset, false, false, to.paragraph);
  }

  /** Where a range marker stands: its paragraph and offset, a block-level one's the next or the previous paragraph's edge. */
  private anchor(marker: CommentMarker, side: 'start' | 'end'): { paragraph: Paragraph; offset: number } | undefined {
    if (marker.paragraph && marker.paragraphContainer) {
      const paragraph = new Paragraph(marker.paragraph, marker.paragraphContainer, this.body);
      return { paragraph, offset: marker.offset ?? (side === 'start' ? 0 : paragraph.text.length) };
    }
    const at = marker.owner.indexOf(marker.item);
    if (at < 0) return undefined;
    const blocks = side === 'start' ? marker.owner.slice(at + 1) : marker.owner.slice(0, at).reverse();
    for (const block of blocks) {
      const found = storyLayout([block], () => '').paragraphs;
      const sp = side === 'start' ? found[0] : found[found.length - 1];
      if (!sp) continue;
      const paragraph = new Paragraph(sp.element, sp.container, this.body);
      return { paragraph, offset: side === 'start' ? 0 : paragraph.text.length };
    }
    return undefined;
  }

  /**
   * The commented range a paragraph at a time: a `Range` per paragraph the comment touches (extension; what
   * `getRange()` returned until CR-014 phase 1). Empty when the comment has no markers in this body.
   */
  getRanges(): Range[] {
    const markers = markersOf(this.body.container, (p) => this.body.hiddenRuns(p)).filter((m) => m.id === this.id && m.paragraph !== undefined);
    if (markers.length === 0) return [];
    const paragraphs = this.body.paragraphs;
    const indexOf = (marker: CommentMarker): number => paragraphs.findIndex((p) => p.element === marker.paragraph);
    const start = markers.find((m) => m.kind === 'start');
    const end = [...markers].reverse().find((m) => m.kind === 'end');
    if (!start || !end) {
      const marker = start ?? end ?? markers.find((m) => m.kind === 'reference')!;
      const i = indexOf(marker);
      if (i < 0) return [];
      const p = paragraphs[i]!;
      const at = marker.offset ?? 0;
      return [new Range(p, at, marker.kind === 'start' ? p.text.length : at)];
    }
    const from = indexOf(start);
    const to = indexOf(end);
    if (from < 0 || to < 0 || to < from) return [];
    const out: Range[] = [];
    for (let i = from; i <= to; i++) {
      const p = paragraphs[i]!;
      out.push(new Range(p, i === from ? start.offset ?? 0 : 0, i === to ? end.offset ?? p.text.length : p.text.length));
    }
    return out;
  }

  // --- internals ---

  /** This comment and, depth first, its replies. */
  private thread(): Comment[] {
    return [this as Comment, ...this.replyList.flatMap((r) => r.thread())];
  }

  private removeSelf(): void {
    removeMarkers(this.body.container, this.id);
    const comments = this.parts.comments.comment;
    if (comments) {
      const i = comments.indexOf(this.element);
      if (i >= 0) comments.splice(i, 1);
    }
    // what the comment's own content related to - a picture's image - goes on save when nothing else
    // uses it, as for a picture deleted outright (CR-002 section 50; docx4j's Removals)
    const released = new Set<string>();
    relationshipIdsIn(this.element.content, released);
    const pkg = this.body.package_ as { pendingRelationshipRemovals?: RelationshipRemovals } | undefined;
    queueRelationshipRemovals(pkg?.pendingRelationshipRemovals, this.parts.commentsPart, released);
    const paraId = this.paraId;
    if (paraId === undefined) return;
    const extended = this.parts.commentsEx?.commentEx;
    if (extended) {
      const i = extended.findIndex((e) => e.paraId === paraId);
      if (i >= 0) extended.splice(i, 1);
    }
    const ids = this.parts.commentsIds?.commentId;
    if (ids) {
      const i = ids.findIndex((e) => e.paraId === paraId);
      if (i >= 0) {
        const durableId = ids[i]!.durableId;
        ids.splice(i, 1);
        if (durableId) this.parts.extensible?.remove(durableId);
        this.parts.datesUtc.delete(paraId.toUpperCase());
      }
    }
  }

  /** The first paragraph's `w14:paraId`, assigned when the comment has none (Word always writes one). */
  ensureParaId(): string {
    const existing = this.paraId;
    if (existing !== undefined) return existing;
    const content = (this.element.content ??= []);
    let first = content.find((e) => typeNameOf(e) === 'org_docx4j_wml.P') as Element<wml.P> | undefined;
    if (!first) {
      first = commentParagraph('', true);
      content.unshift(first as never);
      linkParents(first, this.element);
    }
    const paraId = randomHexId((id) => paraIdTaken(this.parts, this.body.container, id));
    first.value.paraId = paraId;
    return paraId;
  }

  /** Used by `commentsOf` while threading. */
  addReply(reply: Comment): void {
    reply.parentComment = this;
    this.replyList.push(reply);
  }
}

/** The run Word puts first in a comment: the annotation reference in the `CommentReference` style. */
function annotationRefRun(): Element<wml.R> {
  return runOf([el.annotationRef(wmlFactory.createRAnnotationRef())], applyRunOptions({}, { style: COMMENT_REFERENCE_STYLE }));
}

/**
 * The run that anchors a comment in the body: `w:commentReference` in the `CommentReference` style,
 * over the run properties of the text it stands by when given (Word's reference run for a comment at
 * a caret carried the text's `w:sz`, CR-002 section 36).
 */
function referenceRun(id: number, textRPr?: wml.RPr): Element<wml.R> {
  const rPr = applyRunOptions(copyRPr(textRPr) ?? {}, { style: COMMENT_REFERENCE_STYLE });
  return runOf([el.commentReference(wmlFactory.createRCommentReference({ id }))], rPr);
}

/** A comment marker that takes no room: a range start or end, or a run holding only a reference. */
function isCommentMarker(item: Element | undefined): boolean {
  if (!item) return false;
  const name = item.name?.localPart;
  if (name === 'commentRangeStart' || name === 'commentRangeEnd') return true;
  if (typeNameOf(item) !== 'org_docx4j_wml.R') return false;
  const content = ((item.value as wml.R).content ?? []) as Element[];
  return content.length > 0 && content.every((c) => c.name?.localPart === 'commentReference');
}

/**
 * Where an empty range's markers go when other comments' empty ranges already stand at the caret:
 * inside them, right after their starts, as Word nests a second comment made at the same place
 * (start 0, start 1, end 1, reference 1, end 0, reference 0; CR-002 section 36). A start counts only
 * when its comment's end is at the caret too, so a comment over text that begins there is not entered.
 */
function nestedIndex(owner: Element[], index: number): number {
  let from = index;
  while (isCommentMarker(owner[from - 1])) from--;
  let to = index;
  while (isCommentMarker(owner[to])) to++;
  const endsHere = new Set(owner.slice(from, to).filter((e) => e.name?.localPart === 'commentRangeEnd').map((e) => String((e.value as { id?: unknown }).id)));
  let at = from;
  while (at < to && owner[at]!.name?.localPart === 'commentRangeStart' && endsHere.has(String((owner[at]!.value as { id?: unknown }).id))) at++;
  return at > from ? at : index;
}

/** A paragraph of a comment: the `CommentText` style, opened by the annotation reference run. */
function commentParagraph(text: string, first: boolean): Element<wml.P> {
  const runs: Element[] = [];
  if (first) runs.push(annotationRefRun());
  if (text !== '') runs.push(textRun(text));
  return paragraphOf(runs, { TYPE_NAME: 'org_docx4j_wml.PPr', pStyle: { val: COMMENT_TEXT_STYLE } });
}

function insertAt(owner: Element[], index: number, items: Element[], parent: object | undefined): void {
  owner.splice(index, 0, ...items);
  for (const item of items) linkParents(item, parent);
}

function parentOf(element: Element): object | undefined {
  return (element.value as { PARENT?: object }).PARENT;
}

/** The innermost branch of a kept `mc:AlternateContent` no reader takes that holds a value, or undefined. */
function unreadBranchOf(value: object): object | undefined {
  for (let o = (value as { PARENT?: object }).PARENT, guard = 0; o && guard < 256; o = (o as { PARENT?: object }).PARENT, guard++) {
    const tn = (o as { TYPE_NAME?: string }).TYPE_NAME;
    if (tn === undefined || !BRANCH_TYPES.has(tn)) continue;
    const ac = (o as { PARENT?: object }).PARENT as mce.AlternateContent | undefined;
    if (ac && selectedBranchOf(ac) !== (o as { any?: Element[] }).any) return o;
  }
  return undefined;
}

/**
 * Where a marker for this segment's run goes when it is not put inside the revision the run sits
 * in (`placeAround`): beside the run, but outside the `w:ins` or `w:del`, so that a comment made
 * while change tracking is on is a comment and not an insertion of its own (CR-002 section 13).
 * Accepting or rejecting the revision then leaves the comment where it is.
 */
function markerSite(seg: TextSegment, fallback: object, after: boolean): { owner: Element[]; index: number; parent: object } {
  const item = seg.revision ? seg.revision.element : seg.runOwner[seg.runIndex]!;
  const owner = seg.revision ? seg.revision.owner : seg.runOwner;
  return { owner, index: owner.indexOf(item) + (after ? 1 : 0), parent: parentOf(item) ?? fallback };
}

/** Writes the markers of a new comment around a range, splitting runs at its boundaries as `Range.font` does. */
function placeAround(range: Range, id: number): void {
  const paragraph = range.paragraph;
  paragraph.splitAt(range.start);
  if (range.end > range.start) paragraph.splitAt(range.end);
  const covered = range.end > range.start ? paragraph.segments().filter((s) => s.end > range.start && s.start < range.end) : [];
  const start = el.commentRangeStart(wmlFactory.createCommentRangeStart({ id }));
  const end = el.commentRangeEnd(wmlFactory.createCommentRangeEnd({ id }));
  if (covered.length === 0) {
    // An empty range or an empty paragraph: both range markers with nothing between them, then the
    // reference run in the text's run properties, and a later comment at the same place nested inside
    // the earlier - what Word writes (measured by Jason 2026-09-29, CR-002 section 36).
    const segments = paragraph.segments();
    const after = segments.find((s) => s.start >= range.start);
    const before = [...segments].reverse().find((s) => s.end <= range.start);
    const text = after ?? before;
    const site = after ? markerSite(after, paragraph.p, false) : before ? markerSite(before, paragraph.p, true) : undefined;
    const owner = site ? site.owner : (paragraph.p.content ??= []) as Element[];
    const index = nestedIndex(owner, site ? site.index : owner.length);
    insertAt(owner, index, [start, end, referenceRun(id, text?.run.rPr)], site ? site.parent : paragraph.p);
    return;
  }
  const reference = referenceRun(id);
  const first = covered[0]!;
  const last = covered[covered.length - 1]!;
  // A comment on part of an insertion, as Word writes it (the editor's measure-word2010.docx and
  // measure-word15.docx, 2026-09-30; CR-002 section 38): the range start inside the w:ins before the
  // first commented run, the w:ins closed at the range's end - the rest of it a second w:ins with a
  // new id - and the end marker and the plain reference run after it. So the anchor stays narrow
  // and the reference is no insertion. Before, both markers went outside the w:ins, which widened
  // the anchor to the whole insertion. A deletion or a move keeps the markers outside (a deletion's
  // form is not measured; a move's halves are not to be split).
  const inInsertion = (seg: TextSegment): boolean => seg.revision?.kind === 'ins';
  if (inInsertion(last)) {
    paragraph.splitHolderAround(last.revision!, last.runIndex + 1, [end, reference], () => paragraph.nextAnnotationId());
  } else {
    const after = markerSite(last, paragraph.p, true);
    insertAt(after.owner, after.index, [end, reference], after.parent);
  }
  if (inInsertion(first)) insertAt(first.runOwner, first.runIndex, [start], first.revision!.value);
  else {
    const before = markerSite(first, paragraph.p, false);
    insertAt(before.owner, before.index, [start], before.parent);
  }
}

/** Writes the markers of a reply next to its parent's, as Word nests them. */
function placeAfter(body: Body, parentId: number, id: number): void {
  // the parent's markers in every branch of a kept mc:AlternateContent, each branch no reader takes
  // apart, so that the reply is anchored in each as its parent is (CR-013 phase 3)
  const groups = new Map<object | undefined, CommentMarker[]>();
  for (const marker of markersOf(body.container, undefined, 'every')) {
    if (marker.id !== parentId) continue;
    const key = unreadBranchOf(marker.item.value as object);
    let group = groups.get(key);
    if (!group) groups.set(key, group = []);
    group.push(marker);
  }
  const read = groups.get(undefined) ?? [];
  if (!read.some((m) => (m.kind === 'reference' && m.run && m.runOwner) || m.kind === 'end')) {
    // refused before anything is written, so that no marker is left naming a comment never made
    throw new Docx4JException(`Comment ${parentId} has no markers in this body; cannot anchor a reply`);
  }
  for (const markers of groups.values()) placeAfterIn(markers, id);
}

/** A reply's markers after its parent's, in one branch's markers (or those outside any no reader takes). */
function placeAfterIn(markers: CommentMarker[], id: number): void {
  const start = markers.find((m) => m.kind === 'start');
  const end = [...markers].reverse().find((m) => m.kind === 'end');
  const reference = [...markers].reverse().find((m) => m.kind === 'reference');
  if (!(reference?.run && reference.runOwner) && !end) return;
  if (start) insertAt(start.owner, start.owner.indexOf(start.item) + 1, [el.commentRangeStart(wmlFactory.createCommentRangeStart({ id }))], parentOf(start.item));
  if (end) insertAt(end.owner, end.owner.indexOf(end.item) + 1, [el.commentRangeEnd(wmlFactory.createCommentRangeEnd({ id }))], parentOf(end.item));
  if (reference?.run && reference.runOwner) {
    insertAt(reference.runOwner, reference.runOwner.indexOf(reference.run) + 1, [referenceRun(id)], parentOf(reference.run));
  } else if (end) {
    insertAt(end.owner, end.owner.indexOf(end.item) + 1, [referenceRun(id)], parentOf(end.item));
  }
}

/**
 * The `w:comment` and the side-part entries of a new comment; the markers are already written. The
 * paragraphs come checked, by `paragraphsOfContent` before the markers were placed.
 */
function addComment(parts: CommentParts, body: Body, id: number, paragraphs: Element<wml.P>[], parent: Comment | undefined, options?: CommentOptions): Comment {
  const author = options?.author ?? parts.author;
  // one instant for w:date (the local wall clock, as Word writes it) and w16cex:dateUtc (UTC)
  const made = options?.date ?? new Date();
  const paraId = randomHexId((candidate) => paraIdTaken(parts, body.container, candidate));
  paragraphs[0]!.value.paraId = paraId;
  const comment = wmlFactory.createCommentsComment({
    id, author: author.name, initials: options?.initials ?? initialsOf(author),
    date: wordCalendarOf(made),
    content: paragraphs as wml.Comments.Comment['content'],
  });
  const comments = (parts.comments.comment ??= []);
  comments.push(comment);
  linkParents(comment, parts.comments);

  const extended = parts.ensureCommentsEx();
  const entry = w15Factory.createCTCommentEx({ paraId, done: '0' });
  if (parent) entry.paraIdParent = parent.ensureParaId();
  (extended.commentEx ??= []).push(entry);
  linkParents(entry, extended);

  if (parts.commentsIds) {
    const durableId = randomHexId((candidate) => (parts.commentsIds!.commentId ?? []).some((e) => e.durableId === candidate));
    const commentId = w16cidFactory.createCTCommentId({ paraId, durableId });
    (parts.commentsIds.commentId ??= []).push(commentId);
    linkParents(commentId, parts.commentsIds);
    if (parts.extensible) {
      parts.extensible.set(durableId, made);
      parts.datesUtc.set(paraId.toUpperCase(), made);
    }
  }
  addPerson(parts, author.name, author.email);
  return new Comment(comment, parts, body);
}

/** Word records every comment author in `w:people`; the presence info carries the email. */
function addPerson(parts: CommentParts, name: string, email: string | undefined): void {
  const people = parts.people;
  if (!people) return;
  const list = (people.person ??= []);
  if (list.some((p) => p.author === name)) return;
  // w15:contact is declared required but Word omits it; the factory's init is partial, so it is
  // simply left out (see the CR's section 9)
  const person = w15Factory.createCTPerson({
    author: name,
    presenceInfo: w15Factory.createCTPresenceInfo({ providerId: 'None', userId: email ?? name }),
  });
  list.push(person);
  linkParents(person, people);
}

/**
 * `Range.insertComment` (CR-002 section 3.8): the markers, the comment and the side-part entries.
 *
 * `options` (CR-002 section 28.3) name the author, the initials, the date and the `w:id`: in a
 * co-editing room the author is the peer who made the comment and the date is when they made it,
 * neither of which is a property of whoever is saving. A given `id` is checked unused and a clash
 * throws.
 */
export async function insertComment(range: Range, content: CommentContent, options?: CommentOptions): Promise<Comment> {
  // Everything that can refuse the call runs before anything is written: the content, then the id
  // once the parts can say which are taken. A refusal after the markers were placed left them naming
  // a comment that does not exist (CR-002 section 31).
  const paragraphs = paragraphsOfContent(content);
  const body = range.paragraph.parentBody;
  const parts = await commentPartsAccess().create(body);
  const id = commentId(parts, body, options);
  placeAround(range, id);
  return addComment(parts, body, id, paragraphs, undefined, options);
}

/**
 * A comment's entry for markers the caller placed (CR-002 section 36, asked by the editor for its
 * E4.e, which places a new comment's markers itself: a selection may span paragraphs, where
 * `Range.insertComment` anchors in one). Everything `insertComment` writes but the markers: the
 * `w:comment` with the given id, author, initials, date and paragraphs (the first with a
 * `w14:paraId`); its `w15:commentEx` (`w15:done="0"`, and `w15:paraIdParent` for a reply); its
 * `w16cid:commentId` where that part exists, and its `w16cex:commentExtensible` where that one does;
 * and the author's `w15:person`. The comment parts are created as `insertComment` creates them, and
 * the date follows section 29's rule (`w:date` on the local wall clock, `w16cex:dateUtc` the date
 * given). The id must be unused in the comments part - markers naming it are expected, being the
 * caller's - and is not checked against them; a comment with no markers is still a comment, which
 * Word shows unanchored.
 */
export async function addCommentEntry(body: Body, entry: CommentEntry): Promise<Comment> {
  const paragraphs = paragraphsOfContent(entry.content);
  const parts = await commentPartsAccess().create(body);
  const id = entry.id;
  if (!Number.isInteger(id) || id < 0) throw new Docx4JException(`A comment id must be a non-negative integer, not ${id}`);
  if ((parts.comments.comment ?? []).some((c) => c.id === id)) {
    throw new Docx4JException(`The comment id ${id} is already in use in this document's comments part`);
  }
  return addComment(parts, body, id, paragraphs, entry.parent, entry);
}

/**
 * The `w:id` for a new comment: the caller's when given, checked unused across the comments part and
 * the markers in the story, else the next free one. A clash throws rather than renumbering, as a
 * chosen list `numId` does: the caller is asserting something about a room's state (section 28.3).
 */
function commentId(parts: CommentParts, body: Body, options?: CommentOptions): number {
  if (options?.id === undefined) return nextCommentId(parts.comments, body.container);
  const id = options.id;
  if (!Number.isInteger(id) || id < 0) throw new Docx4JException(`A comment id must be a non-negative integer, not ${id}`);
  if ((parts.comments.comment ?? []).some((c) => c.id === id)) {
    throw new Docx4JException(`The comment id ${id} is already in use in this document's comments part`);
  }
  if (markersOf(body.container).some((m) => m.id === id)) {
    throw new Docx4JException(`The comment id ${id} is already named by a marker in this story`);
  }
  return id;
}

/** The paragraphs of a comment's content: a string splits on newlines, elements are taken as they are. */
function paragraphsOfContent(content: CommentContent): Element<wml.P>[] {
  if (typeof content === 'string') return content.split('\n').map((line, i) => commentParagraph(line, i === 0));
  const paragraphs = [...content];
  if (paragraphs.length === 0) throw new Docx4JException('A comment needs at least one paragraph');
  return paragraphs;
}

/**
 * `getComments()` on a body, a paragraph or a range: the comments anchored there in document
 * order, replies nested under their parent (a reply is in the array only when its parent is not).
 */
export async function commentsOf(scope: Body | Paragraph | Range): Promise<Comment[]> {
  const paragraph = scope instanceof Range ? scope.paragraph : scope instanceof Paragraph ? scope : undefined;
  const body = paragraph ? paragraph.parentBody : (scope as Body);
  const parts = await commentPartsAccess().load(body);
  if (!parts) return [];
  await flushPendingRemovals(body, parts);
  const views = viewsOf(parts, body);
  const ids = scope instanceof Range ? idsInRange(scope) : paragraph ? unique(markersOfParagraph(paragraph.element, paragraph.container).map((m) => m.id)) : commentIdsOf(body.container);
  const inScope = new Set(ids);
  const out: Comment[] = [];
  for (const id of ids) {
    const view = views.get(id);
    if (!view) continue;
    if (view.parent && inScope.has(view.parent.id)) continue;
    out.push(view);
  }
  return out;
}

/**
 * Comments whose anchored text went with a rejected insertion or an accepted deletion (check 32;
 * `TrackedChange`'s `removeRevision`) had their markers removed then; their entries go here, with
 * their replies, when the comments are next read or the package is saved (`WordprocessingMLPackage.saveTo`).
 * With the last comment, the comment parts go too (`flushPendingRemovals`).
 */
export async function flushPendingCommentRemovals(body: Body): Promise<void> {
  const pkg = body.package_ as { pendingCommentRemovals?: Set<number> } | undefined;
  if (!pkg?.pendingCommentRemovals?.size) return;
  const parts = await commentPartsAccess().load(body);
  if (!parts) { pkg.pendingCommentRemovals.clear(); return; }
  await flushPendingRemovals(body, parts);
}

/**
 * The entries of the comments a resolution took away, replies included. The side parts the document
 * has are loaded for it and none is created: before section 54 a removal created the commentsExtended,
 * commentsIds and people parts a document lacked, as `delete()` does.
 *
 * With the last comment a resolution takes away, the comment parts go, as Word 365 removed them
 * (docx4j CR-043 section 15; docx4j-python's artefact 8 rejected, docx4j's U22, core-ts check 32): the
 * comments part and its commentsExtended, commentsIds and commentsExtensible parts, with their
 * relationships; and the people part where no tracked change is left to name an author (U22: Word kept
 * it while a deletion was still tracked). docx4j decides when its resolution ends; this, when the
 * entries go, so tracked changes made or resolved in between count.
 */
async function flushPendingRemovals(body: Body, parts: CommentParts): Promise<void> {
  const pkg = body.package_ as { pendingCommentRemovals?: Set<number>; getAllTrackedChanges?: () => Promise<unknown[]> } | undefined;
  const pending = pkg?.pendingCommentRemovals;
  if (!pending?.size) return;
  await parts.loadExisting();
  const views = viewsOf(parts, body);
  let removed = false;
  for (const id of [...pending]) {
    const view = views.get(id);
    if (view) { view.removeThread(); removed = true; }
    pending.delete(id);
  }
  if (!removed || (parts.comments.comment ?? []).length > 0) return;
  const people = parts.people !== undefined && pkg?.getAllTrackedChanges !== undefined && (await pkg.getAllTrackedChanges()).length === 0;
  parts.drop(people);
}

/** A `Comment` view per entry in the parts, replies linked to their parents. */
function viewsOf(parts: CommentParts, body: Body): Map<number, Comment> {
  const views = new Map<number, Comment>();
  for (const comment of parts.comments.comment ?? []) {
    if (typeof comment.id === 'number' && !views.has(comment.id)) views.set(comment.id, new Comment(comment, parts, body));
  }
  const byParaId = new Map<string, Comment>();
  for (const view of views.values()) {
    const paraId = view.paraId;
    if (paraId !== undefined) byParaId.set(paraId, view);
  }
  for (const view of views.values()) {
    const parentParaId = commentExOf(parts, view.paraId)?.paraIdParent;
    const parent = parentParaId === undefined ? undefined : byParaId.get(parentParaId);
    if (parent && parent !== view) parent.addReply(view);
  }
  return views;
}

/** The comments a range touches: a comment whose markers overlap [start, end]. */
function idsInRange(range: Range): number[] {
  const spans = new Map<number, { from: number; to: number }>();
  const order: number[] = [];
  for (const marker of markersOfParagraph(range.paragraph.element, range.paragraph.container, range.paragraph.hiddenRuns())) {
    const at = marker.offset ?? 0;
    const span = spans.get(marker.id);
    if (span) { span.from = Math.min(span.from, at); span.to = Math.max(span.to, at); } else { spans.set(marker.id, { from: at, to: at }); order.push(marker.id); }
  }
  return order.filter((id) => { const s = spans.get(id)!; return s.from <= range.end && s.to >= range.start; });
}

function unique(ids: number[]): number[] {
  const seen = new Set<number>();
  return ids.filter((id) => (seen.has(id) ? false : (seen.add(id), true)));
}

setCommentApi({ commentsOf, insertComment });

export type { CommentParts, Author } from './comments.mjs';