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
import { Namespaces } from '../Namespaces.mjs';
import { DefaultXmlPart } from '../DefaultXmlPart.mjs';
import { MainDocumentPart, CommentsPart, CommentsExtendedPart, CommentsIdsPart, PeoplePart } from './index.mjs';
import type { Body } from '../../model/content/Body.mjs';
import {
  COMMENT_REFERENCE_STYLE, COMMENT_TEXT_STYLE, setCommentPartsAccess, toUtcString,
  type Author, type CommentParts, type CommentsExtensible,
} from '../../model/content/comments.mjs';
// Loaded for its side effect: the Comment views register themselves with the content API.
import '../../model/content/Comment.mjs';

/** `w16cex` (Word 2018): the durable id and a UTC date per comment. Not in the object model. */
const W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';

/** The comment parts of one document. */
class DocumentCommentParts implements CommentParts {
  commentsEx: w15.CTCommentsEx | undefined;
  people: w15.CTPeople | undefined;
  commentsIds: w16cid.CTCommentsIds | undefined;
  extensible: CommentsExtensible | undefined;
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
  return parts;
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

/** The w16cex part as a DOM, when the document has one; none is created (the model does not type it). */
async function loadExtensible(main: MainDocumentPart): Promise<CommentsExtensible | undefined> {
  const rels = main.relationshipsPart;
  const rel = rels?.getRelationshipByType(Namespaces.COMMENTS_EXTENSIBLE);
  const part = rel && rels ? rels.getPart(rel) : undefined;
  if (!(part instanceof DefaultXmlPart)) return undefined;
  const document = await part.getDocument();
  const root = document.documentElement;
  if (!root) return undefined;
  const entryOf = (durableId: string): Element | undefined => {
    for (let node = root.firstChild; node; node = node.nextSibling) {
      if (node.nodeType !== 1) continue;
      const element = node as Element;
      const id = element.getAttributeNS(W16CEX, 'durableId') ?? element.getAttribute('w16cex:durableId');
      if (id && id.toLowerCase() === durableId.toLowerCase()) return element;
    }
    return undefined;
  };
  return {
    set(durableId: string, dateUtc: Date): void {
      let entry = entryOf(durableId);
      if (!entry) {
        entry = document.createElementNS(W16CEX, 'w16cex:commentExtensible');
        entry.setAttributeNS(W16CEX, 'w16cex:durableId', durableId);
        root.appendChild(entry);
      }
      entry.setAttributeNS(W16CEX, 'w16cex:dateUtc', toUtcString(dateUtc));
    },
    remove(durableId: string): void {
      const entry = entryOf(durableId);
      if (entry) root.removeChild(entry);
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
  const add = (style: wml.Style): void => {
    list.push(style);
    linkParents(style, styles);
  };
  if (!list.some((s) => s.styleId === COMMENT_TEXT_STYLE)) {
    add(wmlFactory.createStyle({
      type: 'paragraph', styleId: COMMENT_TEXT_STYLE,
      name: wmlFactory.createStyleName({ val: 'annotation text' }),
      basedOn: wmlFactory.createStyleBasedOn({ val: 'Normal' }),
      uiPriority: wmlFactory.createStyleUiPriority({ val: 99 }),
      semiHidden: wmlFactory.createBooleanDefaultTrue(),
      unhideWhenUsed: wmlFactory.createBooleanDefaultTrue(),
      rPr: wmlFactory.createRPr({ sz: wmlFactory.createHpsMeasure({ val: 20 }), szCs: wmlFactory.createHpsMeasure({ val: 20 }) }),
    }));
  }
  if (!list.some((s) => s.styleId === COMMENT_REFERENCE_STYLE)) {
    add(wmlFactory.createStyle({
      type: 'character', styleId: COMMENT_REFERENCE_STYLE,
      name: wmlFactory.createStyleName({ val: 'annotation reference' }),
      basedOn: wmlFactory.createStyleBasedOn({ val: 'DefaultParagraphFont' }),
      uiPriority: wmlFactory.createStyleUiPriority({ val: 99 }),
      semiHidden: wmlFactory.createBooleanDefaultTrue(),
      unhideWhenUsed: wmlFactory.createBooleanDefaultTrue(),
      rPr: wmlFactory.createRPr({ sz: wmlFactory.createHpsMeasure({ val: 16 }), szCs: wmlFactory.createHpsMeasure({ val: 16 }) }),
    }));
  }
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

export { W16CEX };
