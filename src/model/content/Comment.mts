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
import { runOf, paragraphOf, typeNameOf, type Element } from './tree.mjs';
import {
  COMMENT_REFERENCE_STYLE, COMMENT_TEXT_STYLE, calendarToDate, commentExOf, commentIdsOf, commentPartsAccess, commentParaId,
  dateToCalendar, initialsOf, isDone, markersOf, markersOfParagraph, nextCommentId, paraIdTaken, randomHexId, removeMarkers,
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

  /** `w:date`; undefined when the comment carries none. */
  get creationDate(): Date | undefined {
    return calendarToDate(this.element.date);
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
    return this.commentBody.text;
  }
  set content(text: string) {
    const paraId = commentParaId(this.element);
    const paragraphs = text.split('\n').map((line, i) => commentParagraph(line, i === 0));
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

  /** A reply in the same thread, anchored to the same range, as Word writes them. */
  async reply(text: string): Promise<Comment> {
    const parts = await this.parts.ensureAll();
    const id = nextCommentId(parts.comments, this.body.container);
    placeAfter(this.body, this.id, id);
    const reply = addComment(parts, this.body, id, text, this);
    this.replyList.push(reply);
    reply.parentComment = this;
    return reply;
  }

  /** Removes the comment and its replies: the markers, the reference runs and every side-part entry. */
  async delete(): Promise<void> {
    await this.parts.ensureAll();
    for (const comment of this.thread()) comment.removeSelf();
    if (this.parentComment) {
      const i = this.parentComment.replyList.indexOf(this);
      if (i >= 0) this.parentComment.replyList.splice(i, 1);
      this.parentComment = undefined;
    }
  }

  /**
   * The commented range: a `Range` per paragraph the comment touches (extension; Office JS
   * returns one `Range`). Empty when the comment has no markers in this body.
   */
  getRange(): Range[] {
    const markers = markersOf(this.body.container).filter((m) => m.id === this.id && m.paragraph !== undefined);
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

/** The run that anchors a comment in the body: `w:commentReference` in the `CommentReference` style. */
function referenceRun(id: number): Element<wml.R> {
  return runOf([el.commentReference(wmlFactory.createRCommentReference({ id }))], applyRunOptions({}, { style: COMMENT_REFERENCE_STYLE }));
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

/** Writes the markers of a new comment around a range, splitting runs at its boundaries as `Range.font` does. */
function placeAround(range: Range, id: number): void {
  const paragraph = range.paragraph;
  paragraph.splitAt(range.start);
  if (range.end > range.start) paragraph.splitAt(range.end);
  const covered = range.end > range.start ? paragraph.segments().filter((s) => s.end > range.start && s.start < range.end) : [];
  const start = el.commentRangeStart(wmlFactory.createCommentRangeStart({ id }));
  const end = el.commentRangeEnd(wmlFactory.createCommentRangeEnd({ id }));
  const reference = referenceRun(id);
  if (covered.length === 0) {
    // an empty range or an empty paragraph: the reference run only, which is all Word requires
    const segments = paragraph.segments();
    const after = segments.find((s) => s.start >= range.start);
    const before = [...segments].reverse().find((s) => s.end <= range.start);
    if (after) insertAt(after.runOwner, after.runIndex, [reference], parentOf(after.runOwner[after.runIndex]!) ?? paragraph.p);
    else if (before) insertAt(before.runOwner, before.runIndex + 1, [reference], parentOf(before.runOwner[before.runIndex]!) ?? paragraph.p);
    else insertAt((paragraph.p.content ??= []) as Element[], paragraph.p.content!.length, [reference], paragraph.p);
    return;
  }
  const first = covered[0]!;
  const last = covered[covered.length - 1]!;
  const firstRun = first.runOwner[first.runIndex]!;
  const lastRun = last.runOwner[last.runIndex]!;
  insertAt(last.runOwner, last.runOwner.indexOf(lastRun) + 1, [end, reference], parentOf(lastRun) ?? paragraph.p);
  insertAt(first.runOwner, first.runOwner.indexOf(firstRun), [start], parentOf(firstRun) ?? paragraph.p);
}

/** Writes the markers of a reply next to its parent's, as Word nests them. */
function placeAfter(body: Body, parentId: number, id: number): void {
  const markers = markersOf(body.container).filter((m) => m.id === parentId);
  const start = markers.find((m) => m.kind === 'start');
  const end = [...markers].reverse().find((m) => m.kind === 'end');
  const reference = [...markers].reverse().find((m) => m.kind === 'reference');
  if (start) insertAt(start.owner, start.owner.indexOf(start.item) + 1, [el.commentRangeStart(wmlFactory.createCommentRangeStart({ id }))], parentOf(start.item));
  if (end) insertAt(end.owner, end.owner.indexOf(end.item) + 1, [el.commentRangeEnd(wmlFactory.createCommentRangeEnd({ id }))], parentOf(end.item));
  if (reference?.run && reference.runOwner) {
    insertAt(reference.runOwner, reference.runOwner.indexOf(reference.run) + 1, [referenceRun(id)], parentOf(reference.run));
  } else if (end) {
    insertAt(end.owner, end.owner.indexOf(end.item) + 1, [referenceRun(id)], parentOf(end.item));
  } else {
    throw new Docx4JException(`Comment ${parentId} has no markers in this body; cannot anchor a reply`);
  }
}

/** The `w:comment` and the side-part entries of a new comment; the markers are already written. */
function addComment(parts: CommentParts, body: Body, id: number, text: string, parent: Comment | undefined): Comment {
  const author = parts.author;
  const paragraphs = text.split('\n').map((line, i) => commentParagraph(line, i === 0));
  const paraId = randomHexId((candidate) => paraIdTaken(parts, body.container, candidate));
  paragraphs[0]!.value.paraId = paraId;
  const comment = wmlFactory.createCommentsComment({
    id, author: author.name, initials: initialsOf(author), date: dateToCalendar(new Date()),
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
    parts.extensible?.set(durableId, new Date());
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

/** `Range.insertComment` (CR-002 section 3.8): the markers, the comment and the side-part entries. */
export async function insertComment(range: Range, text: string): Promise<Comment> {
  const body = range.paragraph.parentBody;
  const parts = await commentPartsAccess().create(body);
  const id = nextCommentId(parts.comments, body.container);
  placeAround(range, id);
  return addComment(parts, body, id, text, undefined);
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

/** The comments a range touches: a comment whose markers overlap [start, end]. */
function idsInRange(range: Range): number[] {
  const spans = new Map<number, { from: number; to: number }>();
  const order: number[] = [];
  for (const marker of markersOfParagraph(range.paragraph.element, range.paragraph.container)) {
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