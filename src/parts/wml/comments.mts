// The part half of CR-002 phase G: the four comment parts of a document, created and loaded for
// the `Comment` views. Registered with the content API through `setCommentPartsAccess`, so that
// the model never imports a part (CR-001's "no runtime import cycles"); `WordprocessingMLPackage`
// imports this module for that registration, as `packages/index.mts` does for the package classes.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as w15 from '@docx4j/generated-objects-ts/modules/org_docx4j_w15';
import type * as w16cid from '@docx4j/generated-objects-ts/modules/org_docx4j_w16cid';
import * as wmlFactory from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import * as w15Factory from '@docx4j/generated-objects-ts/factory/org_docx4j_w15';
import * as w16cidFactory from '@docx4j/generated-objects-ts/factory/org_docx4j_w16cid';
import { linkParents } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { parseXml } from '../../xml/dom.mjs';
import { Namespaces } from '../Namespaces.mjs';
import { MainDocumentPart, CommentsPart, CommentsExtendedPart, CommentsIdsPart, PeoplePart } from './index.mjs';
import type { Body } from '../../model/content/Body.mjs';
import {
  COMMENT_REFERENCE_STYLE, COMMENT_TEXT_STYLE, setCommentPartsAccess, dateToCalendar,
  type Author, type CommentParts, type CommentsExtensible,
} from '../../model/content/comments.mjs';
import { dateOf } from '../../model/content/tracking.mjs';
// Loaded for its side effect: the Comment views register themselves with the content API.
import '../../model/content/Comment.mjs';
import { commentStyles } from './commentStyles.mjs';

/** The comment parts of one document. */
class DocumentCommentParts implements CommentParts {
  commentsEx: w15.CTCommentsEx | undefined;
  people: w15.CTPeople | undefined;
  commentsIds: w16cid.CTCommentsIds | undefined;
  extensible: CommentsExtensible | undefined;
  readonly datesUtc = new Map<string, Date>();
  private loadedAll = false;

  constructor(
    readonly main: MainDocumentPart,
    readonly commentsPart: CommentsPart,
    readonly comments: wml.Comments,
    readonly author: Author,
  ) {}

  ensureCommentsEx(): w15.CTCommentsEx {
    if (this.commentsEx) return this.commentsEx;
    const existing = this.main.commentsExtendedPart;
    if (existing) {
      if (!existing.isUnmarshalled) throw new Docx4JException(`${existing.partName} is not unmarshalled; call getComments() first`);
      this.commentsEx = existing.contents;
      return this.commentsEx;
    }
    const part = new CommentsExtendedPart();
    part.setContents(w15Factory.createCTCommentsEx({ commentEx: [] }));
    this.main.addTargetPart(part);
    this.commentsEx = part.contents;
    return this.commentsEx;
  }

  async ensureAll(): Promise<CommentParts> {
    if (this.loadedAll) return this;
    this.ensureCommentsEx();
    this.people ??= await ensurePeople(this.main);
    this.commentsIds ??= await ensureCommentsIds(this.main);
    this.extensible ??= await loadExtensible(this.main);
    this.loadedAll = true;
    return this;
  }
}

/** The document part a body's comments belong to: the body's own part, else the package's main part. */
function mainPartOf(body: Body): MainDocumentPart {
  if (body.part instanceof MainDocumentPart) return body.part;
  const pkg = body.package_ as { mainDocumentPart?: MainDocumentPart } | undefined;
  if (pkg?.mainDocumentPart) return pkg.mainDocumentPart;
  throw new Docx4JException('Comments need a main document part; this body has none');
}

/** `WordprocessingMLPackage.author`, read without importing the package (no runtime cycle). */
function authorOf(body: Body): Author {
  const pkg = body.package_ as { author?: Author } | undefined;
  return pkg?.author ?? { name: 'docx4j' };
}

async function partsFor(main: MainDocumentPart, commentsPart: CommentsPart, author: Author): Promise<DocumentCommentParts> {
  const parts = new DocumentCommentParts(main, commentsPart, await commentsPart.getContents(), author);
  // The two side parts a view reads: the threads and done flags, and the authors' emails.
  if (main.commentsExtendedPart) parts.commentsEx = await main.commentsExtendedPart.getContents();
  if (main.peoplePart) parts.people = await main.peoplePart.getContents();
  await readDatesUtc(main, parts.datesUtc);
  return parts;
}

const W16CID_NS = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const W16CEX_NS = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';

/**
 * Each comment's `w16cex:dateUtc`, by the paragraph id `w16cid:commentId` pairs its durable id with.
 * From the parts' contents when they are unmarshalled, else from their XML, which leaves them
 * untouched: no view needs them otherwise, and an unmarshalled part is re-marshalled on save.
 */
async function readDatesUtc(main: MainDocumentPart, into: Map<string, Date>): Promise<void> {
  const idsPart = main.commentsIdsPart;
  const extensiblePart = main.commentsExtensiblePart;
  if (!idsPart || !extensiblePart) return;
  const paraIdOf = new Map<string, string>();
  if (idsPart.isUnmarshalled) {
    for (const e of idsPart.contents.commentId ?? []) if (e.durableId && e.paraId) paraIdOf.set(e.durableId.toUpperCase(), e.paraId.toUpperCase());
  } else {
    for (const e of Array.from(parseXml(await idsPart.getXml()).getElementsByTagNameNS(W16CID_NS, 'commentId'))) {
      const durableId = e.getAttributeNS(W16CID_NS, 'durableId');
      const paraId = e.getAttributeNS(W16CID_NS, 'paraId');
      if (durableId && paraId) paraIdOf.set(durableId.toUpperCase(), paraId.toUpperCase());
    }
  }
  const add = (durableId: string | undefined, date: Date | undefined): void => {
    const paraId = durableId ? paraIdOf.get(durableId.toUpperCase()) : undefined;
    if (paraId && date && !Number.isNaN(date.getTime())) into.set(paraId, date);
  };
  if (extensiblePart.isUnmarshalled) {
    for (const e of extensiblePart.contents.commentExtensible ?? []) add(e.durableId, dateOf(e.dateUtc));
  } else {
    for (const e of Array.from(parseXml(await extensiblePart.getXml()).getElementsByTagNameNS(W16CEX_NS, 'commentExtensible'))) {
      const value = e.getAttributeNS(W16CEX_NS, 'dateUtc');
      add(e.getAttributeNS(W16CEX_NS, 'durableId') ?? undefined, value ? new Date(value) : undefined);
    }
  }
}

async function ensurePeople(main: MainDocumentPart): Promise<w15.CTPeople> {
  if (main.peoplePart) return main.peoplePart.getContents();
  const part = new PeoplePart();
  part.setContents(w15Factory.createCTPeople({ person: [] }));
  main.addTargetPart(part);
  return part.contents;
}

async function ensureCommentsIds(main: MainDocumentPart): Promise<w16cid.CTCommentsIds> {
  if (main.commentsIdsPart) return main.commentsIdsPart.getContents();
  const part = new CommentsIdsPart();
  part.setContents(w16cidFactory.createCTCommentsIds({ commentId: [] }));
  main.addTargetPart(part);
  return part.contents;
}

/** The w16cex part, typed (objects 0.1.3), when the document has one; none is created. */
async function loadExtensible(main: MainDocumentPart): Promise<CommentsExtensible | undefined> {
  const part = main.commentsExtensiblePart;
  if (!part) return undefined;
  const contents = await part.getContents();
  const entries = (contents.commentExtensible ??= []);
  const indexOf = (durableId: string): number => entries.findIndex((e) => e.durableId?.toLowerCase() === durableId.toLowerCase());
  return {
    set(durableId: string, dateUtc: Date): void {
      const i = indexOf(durableId);
      if (i >= 0) entries[i]!.dateUtc = dateToCalendar(dateUtc);
      else entries.push({ durableId, dateUtc: dateToCalendar(dateUtc) });
    },
    remove(durableId: string): void {
      const i = indexOf(durableId);
      if (i >= 0) entries.splice(i, 1);
    },
  };
}

/**
 * Word's `CommentText` and `CommentReference` styles, added when the styles part lacks them.
 * The part is only unmarshalled when its XML does not already name both styles, so a document
 * that has them keeps its styles part byte for byte.
 */
async function ensureCommentStyles(main: MainDocumentPart): Promise<void> {
  const part = main.styleDefinitionsPart;
  if (!part) return;
  if (!part.isUnmarshalled) {
    const xml = await part.getXml();
    if (xml.includes(`w:styleId="${COMMENT_TEXT_STYLE}"`) && xml.includes(`w:styleId="${COMMENT_REFERENCE_STYLE}"`)) return;
  }
  const styles = await part.getContents();
  const list = (styles.style ??= []);
  // the styles part is about to change: the property resolver caches resolved styles
  const pkg = main.package as { refreshPropertyResolver?: () => void } | undefined;
  const add = (style: wml.Style): void => {
    list.push(style);
    linkParents(style, styles);
    pkg?.refreshPropertyResolver?.();
  };
  for (const style of commentStyles()) if (!list.some((s) => s.styleId === style.styleId)) add(style);
}

setCommentPartsAccess({
  async load(body: Body): Promise<CommentParts | undefined> {
    const main = mainPartOf(body);
    if (!main.commentsPart) return undefined;
    return partsFor(main, main.commentsPart, authorOf(body));
  },
  async create(body: Body): Promise<CommentParts> {
    const main = mainPartOf(body);
    let commentsPart = main.commentsPart;
    if (!commentsPart) {
      commentsPart = new CommentsPart();
      commentsPart.setContents(wmlFactory.createComments({ comment: [] }));
      main.addTargetPart(commentsPart);
    }
    await ensureCommentStyles(main);
    const parts = await partsFor(main, commentsPart, authorOf(body));
    await parts.ensureAll();
    return parts;
  },
});

