// The comment plumbing of CR-002 section 3.8 (phase G), split from the `Comment` view so that
// nothing here imports a part at runtime: the part classes and the four comment parts are
// reached through `CommentPartsAccess`, which `src/parts/wml/comments.mts` registers, and the
// views are reached through `CommentApi`, which `Comment.mts` registers. That keeps the runtime
// import graph one-directional (tree -> comments -> Body/Paragraph/Range -> Comment), as
// CR-001's "no runtime import cycles" rule asks.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as w15 from '@docx4j/generated-objects-ts/modules/org_docx4j_w15';
import type * as w16cid from '@docx4j/generated-objects-ts/modules/org_docx4j_w16cid';
import type { XmlCalendar } from '@docx4j/generated-objects-ts';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { typeNameOf, childrenOf, runItemsOf, itemTextOf, revisionKindOf, RUN_HOLDERS, type Element } from './tree.mjs';
import type { Body } from './Body.mjs';
import type { Paragraph } from './Paragraph.mjs';
import type { Range } from './Range.mjs';
import type { Comment } from './Comment.mjs';

/** The style Word gives a comment's paragraphs. */
export const COMMENT_TEXT_STYLE = 'CommentText';
/** The character style of the annotation reference run, in the comment and in the body. */
export const COMMENT_REFERENCE_STYLE = 'CommentReference';

/**
 * Who a comment (and, with phase F, a tracked change) is by. `WordprocessingMLPackage.author`;
 * there is no signed-in user here, so the package carries the identity.
 */
export interface Author {
  name: string;
  /** Word's `w:initials`; derived from the name when not given. */
  initials?: string;
  /** Written to `w:people` as the author's presence info; Office JS `Comment.authorEmail`. */
  email?: string;
}

/** The initials Word would show: the first letter of each word of the name, at most three. */
export function initialsOf(author: Author): string {
  if (author.initials) return author.initials;
  return author.name.split(/\s+/).filter((w) => w.length > 0).slice(0, 3).map((w) => w[0]!.toUpperCase()).join('');
}

/**
 * The `w16cex` part (`w16cex:commentsExtensible`, a UTC date per comment), typed since objects
 * 0.1.3 (`CommentsExtensiblePart`). Only kept in step when the document already has the part;
 * none is created.
 */
export interface CommentsExtensible {
  /** Adds or updates the entry of a durable id. */
  set(durableId: string, dateUtc: Date): void;
  remove(durableId: string): void;
}

/** The four (five with w16cex) comment parts of a document, unmarshalled. */
export interface CommentParts {
  /** The comments part, for the `Body` view over a comment's paragraphs. */
  readonly commentsPart: XmlPart<unknown>;
  /** `w:comments` (docx4j CommentsPart contents). */
  readonly comments: wml.Comments;
  /** `w15:commentsEx`: the thread links and the done flags. */
  commentsEx: w15.CTCommentsEx | undefined;
  /** `w15:people`. */
  people: w15.CTPeople | undefined;
  /** `w16cid:commentsIds`: the durable ids; loaded by `ensureAll` only, since no view reads it. */
  commentsIds: w16cid.CTCommentsIds | undefined;
  /** `w16cex:commentsExtensible`, when the document has one. */
  extensible: CommentsExtensible | undefined;
  /** The package's author identity. */
  readonly author: Author;
  /** The w15 extended part, created (with its relationship and content type) when absent. */
  ensureCommentsEx(): w15.CTCommentsEx;
  /** Loads and creates the side parts the bookkeeping needs (w15, w16cid; w16cex only if present). */
  ensureAll(): Promise<CommentParts>;
}

/** How the content API reaches the comment parts; implemented and registered by `src/parts/wml/comments.mts`. */
export interface CommentPartsAccess {
  /** The document's comment parts, or undefined when it has none. Only what a view reads is unmarshalled. */
  load(body: Body): Promise<CommentParts | undefined>;
  /** As `load`, creating every absent part and the two comment styles. */
  create(body: Body): Promise<CommentParts>;
}

let partsAccess: CommentPartsAccess | undefined;

/** Registered by `src/parts/wml/comments.mts` (imported for that by `WordprocessingMLPackage`). */
export function setCommentPartsAccess(access: CommentPartsAccess): void {
  partsAccess = access;
}

export function commentPartsAccess(): CommentPartsAccess {
  if (!partsAccess) throw new Docx4JException('Comments need the parts layer: import @docx4j/core-ts or @docx4j/core-ts/packages, which registers it');
  return partsAccess;
}

/** How `Body`, `Paragraph` and `Range` reach the `Comment` views; registered by `Comment.mts`. */
export interface CommentApi {
  commentsOf(scope: Body | Paragraph | Range): Promise<Comment[]>;
  insertComment(range: Range, text: string): Promise<Comment>;
}

let api: CommentApi | undefined;

export function setCommentApi(implementation: CommentApi): void {
  api = implementation;
}

export function commentApi(): CommentApi {
  if (!api) throw new Docx4JException('The Comment views are not loaded: import @docx4j/core-ts or @docx4j/core-ts/model');
  return api;
}

// --- the markers in the document (w:commentRangeStart, w:commentRangeEnd, w:commentReference) ---

export type CommentMarkerKind = 'start' | 'end' | 'reference';

/** One comment marker in a body, with everything an edit needs: what holds it, and where it is. */
export interface CommentMarker {
  id: number;
  kind: CommentMarkerKind;
  /** The `{ name, value }` pair: w:commentRangeStart, w:commentRangeEnd or w:commentReference. */
  item: Element;
  /** The array holding the item (a paragraph's content, a run's content, a body's content). */
  owner: Element[];
  /** For a reference: the run holding it and the array holding that run. */
  run?: Element<wml.R>;
  runOwner?: Element[];
  /** The paragraph the marker is in, and the block-level array holding that paragraph. */
  paragraph?: Element<wml.P>;
  paragraphContainer?: Element[];
  /** The marker's offset in the paragraph's text; undefined for a block-level marker. */
  offset?: number;
}

const COMMENT_RANGE_START = 'org_docx4j_wml.CommentRangeStart';
const COMMENT_RANGE_END = 'org_docx4j_wml.CommentRangeEnd';
const COMMENT_REFERENCE = 'org_docx4j_wml.R.CommentReference';

/**
 * Every comment marker under a container (a body, a cell, a paragraph), in document order.
 * Deleted content (`w:del`) is not entered, as the text model does not enter it either.
 */
export function markersOf(container: object): CommentMarker[] {
  const out: CommentMarker[] = [];
  collect(out, container);
  return out;
}

/** The markers of one paragraph, with their offsets in its text. */
export function markersOfParagraph(element: Element<wml.P>, container: Element[]): CommentMarker[] {
  const out: CommentMarker[] = [];
  collect(out, undefined, { element, container });
  return out;
}

function collect(out: CommentMarker[], container: object | undefined, paragraph?: { element: Element<wml.P>; container: Element[] }): void {
  const visitParagraph = (element: Element<wml.P>, blockContainer: Element[]): void => {
    let pos = 0;
    const visitRuns = (items: Element[] | undefined): void => {
      if (!items) return;
      for (const el of items) {
        const tn = typeNameOf(el);
        if (tn === 'org_docx4j_wml.R') {
          const content = (el.value as wml.R).content ?? [];
          for (const item of content) {
            if (typeNameOf(item) === COMMENT_REFERENCE) {
              out.push({
                id: (item.value as wml.R.CommentReference).id, kind: 'reference', item, owner: content as Element[],
                run: el as Element<wml.R>, runOwner: items, paragraph: element, paragraphContainer: blockContainer, offset: pos,
              });
            } else pos += itemTextOf(item)?.length ?? 0;
          }
        } else if (tn === COMMENT_RANGE_START || tn === COMMENT_RANGE_END) {
          out.push({
            id: (el.value as wml.CommentRangeStart).id, kind: tn === COMMENT_RANGE_START ? 'start' : 'end',
            item: el, owner: items, paragraph: element, paragraphContainer: blockContainer, offset: pos,
          });
        } else if (tn !== undefined && RUN_HOLDERS.has(tn)) {
          visitRuns(runItemsOf(el.value as object));
        } else {
          // the accepted view, which is what the offsets are in: a w:ins and a w:moveTo hold text
          // and markers, a w:del and a w:moveFrom hold neither (CR-002 phase F)
          const kind = revisionKindOf(el);
          if (kind === 'ins' || kind === 'moveTo') visitRuns(runItemsOf(el.value as object));
        }
      }
    };
    visitRuns(runItemsOf(element.value));
  };
  const visitBlocks = (items: Element[]): void => {
    for (const el of items) {
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.P') visitParagraph(el as Element<wml.P>, items);
      else if (tn === COMMENT_RANGE_START || tn === COMMENT_RANGE_END) {
        out.push({ id: (el.value as wml.CommentRangeStart).id, kind: tn === COMMENT_RANGE_START ? 'start' : 'end', item: el, owner: items });
      } else if (typeof el.value === 'object' && el.value !== null) {
        if (tn === 'org_docx4j_wml.Tbl') {
          for (const row of childrenOf(el.value) ?? []) {
            for (const cell of childrenOf(row.value as object) ?? []) visitBlocks(childrenOf(cell.value as object) ?? []);
          }
        } else {
          const children = childrenOf(el.value);
          if (children) visitBlocks(children);
        }
      }
    }
  };
  if (paragraph) visitParagraph(paragraph.element, paragraph.container);
  else if (container) visitBlocks(childrenOf(container) ?? []);
}

/** The comment ids under a container, in document order, first occurrence first. */
export function commentIdsOf(container: object): number[] {
  const seen = new Set<number>();
  const out: number[] = [];
  for (const m of markersOf(container)) if (!seen.has(m.id)) { seen.add(m.id); out.push(m.id); }
  return out;
}

/** Removes every marker of a comment from a body, and the reference run when nothing is left in it. */
export function removeMarkers(container: object, id: number): void {
  for (const marker of markersOf(container)) {
    if (marker.id !== id) continue;
    const i = marker.owner.indexOf(marker.item);
    if (i >= 0) marker.owner.splice(i, 1);
    if (marker.kind === 'reference' && marker.run && marker.runOwner) {
      const content = marker.run.value.content ?? [];
      if (content.length === 0) {
        const j = marker.runOwner.indexOf(marker.run);
        if (j >= 0) marker.runOwner.splice(j, 1);
      }
    }
  }
}

// --- ids ---

/** The next free `w:id` for a comment: above every comment and every marker in the body. */
export function nextCommentId(comments: wml.Comments, container: object): number {
  let max = -1;
  for (const c of comments.comment ?? []) if (typeof c.id === 'number' && c.id > max) max = c.id;
  for (const m of markersOf(container)) if (m.id > max) max = m.id;
  return max + 1;
}

/** An 8 hex digit id, as Word writes `w14:paraId` and `w16cid:durableId` (never 0, high bit clear). */
export function randomHexId(taken: (id: string) => boolean): string {
  for (let attempt = 0; attempt < 1000; attempt++) {
    const id = (1 + Math.floor(Math.random() * 0x7ffffffe)).toString(16).toUpperCase().padStart(8, '0');
    if (!taken(id)) return id;
  }
  throw new Docx4JException('Could not allocate a paragraph id');
}

/** The `w14:paraId` of a comment's first paragraph: the key of its w15 and w16cid entries. */
export function commentParaId(comment: wml.Comments.Comment): string | undefined {
  for (const el of comment.content ?? []) if (typeNameOf(el) === 'org_docx4j_wml.P') return (el.value as wml.P).paraId;
  return undefined;
}

/** True when a paragraph id is already used by a comment or by a paragraph of the document. */
export function paraIdTaken(parts: CommentParts, container: object, id: string): boolean {
  for (const c of parts.comments.comment ?? []) for (const el of c.content ?? []) {
    if (typeNameOf(el) === 'org_docx4j_wml.P' && (el.value as wml.P).paraId === id) return true;
  }
  let found = false;
  const visit = (items: Element[]): void => {
    for (const el of items) {
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.P') { if ((el.value as wml.P).paraId === id) found = true; }
      else if (typeof el.value === 'object' && el.value !== null) {
        if (tn === 'org_docx4j_wml.Tbl') {
          for (const row of childrenOf(el.value) ?? []) for (const cell of childrenOf(row.value as object) ?? []) visit(childrenOf(cell.value as object) ?? []);
        } else {
          const children = childrenOf(el.value);
          if (children) visit(children);
        }
      }
    }
  };
  visit(childrenOf(container) ?? []);
  return found;
}

// --- w15 entries ---

/** The `w15:commentEx` entry keyed by a comment's first paragraph id. */
export function commentExOf(parts: CommentParts, paraId: string | undefined): w15.CTCommentEx | undefined {
  if (!paraId) return undefined;
  return (parts.commentsEx?.commentEx ?? []).find((e) => e.paraId === paraId);
}

/** Word writes `w15:done="0"` and `"1"`; `w:ST_OnOff` also allows true/false/on/off. */
export function isDone(value: string | undefined): boolean {
  return value === '1' || value === 'true' || value === 'on';
}

// --- dates ---

/** A Jsonix calendar (what `w:date` unmarshals to) as a Date; undefined when there is none. */
export function calendarToDate(calendar: XmlCalendar | undefined): Date | undefined {
  if (!calendar || calendar.year === undefined || Number.isNaN(calendar.year)) return undefined;
  const num = (v: number | undefined, fallback = 0): number => (v === undefined || Number.isNaN(v) ? fallback : v);
  const ms = Date.UTC(calendar.year, num(calendar.month, 1) - 1, num(calendar.day, 1), num(calendar.hour), num(calendar.minute), num(calendar.second), Math.round(num(calendar.fractionalSecond) * 1000));
  return new Date(ms - num(calendar.timezone) * 60000);
}

/** A Date as the UTC calendar Word writes (`2026-09-15T09:00:00Z`). */
export function dateToCalendar(date: Date): XmlCalendar {
  return {
    year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate(),
    hour: date.getUTCHours(), minute: date.getUTCMinutes(), second: date.getUTCSeconds(),
    fractionalSecond: 0, timezone: 0,
  };
}

/** `2026-09-15T09:00:00Z`, the form Word writes in w16cex. */
export function toUtcString(date: Date): string {
  return `${date.toISOString().substring(0, 19)}Z`;
}
