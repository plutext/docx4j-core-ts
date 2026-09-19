// WordprocessingML parts (docx4j org.docx4j.openpackaging.parts.WordprocessingML).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as w15 from '@docx4j/generated-objects-ts/modules/org_docx4j_w15';
import type * as w16cid from '@docx4j/generated-objects-ts/modules/org_docx4j_w16cid';
import type * as w16cex from '@docx4j/generated-objects-ts/modules/org_docx4j_w16cex';
import type * as wne from '@docx4j/generated-objects-ts/modules/org_docx4j_com_microsoft_schemas_office_word_x2006_wordml';
import { XmlPart } from '../XmlPart.mjs';
import { Part } from '../Part.mjs';
import { PartName } from '../../opc/PartName.mjs';
import { ContentTypes } from '../../opc/ContentTypes.mjs';
import { Namespaces } from '../Namespaces.mjs';
import { DefaultXmlPart } from '../DefaultXmlPart.mjs';
import { ThemePart } from '../dml/index.mjs';
import { Body } from '../../model/content/Body.mjs';
import { Docx4JException } from '../../opc/exceptions.mjs';

const bodies = new WeakMap<object, Body>();

/**
 * Builds the package's `PropertyResolver` if it has one to build, so that the effective reads
 * of a `Body` obtained asynchronously work (CR-001 Phase B step 2). Structural, not by import:
 * a static import of `WordprocessingMLPackage` here would be a runtime cycle.
 */
async function ensurePropertyResolver(part: Part): Promise<void> {
  const pkg = part.package as { getPropertyResolver?: () => Promise<unknown> } | undefined;
  await pkg?.getPropertyResolver?.();
}

/** The Body view of a part's block-level container, one per container object. */
function bodyOf(part: XmlPart<unknown>, container: { content?: unknown[] } | undefined, prefix: string): Body {
  if (!container) throw new Docx4JException(`${part.partName} has no content yet; await getContents() first, or set contents`);
  let body = bodies.get(container);
  if (!body) {
    body = new Body(part, container as { content?: never[] }, prefix, part.package);
    bodies.set(container, body);
  }
  return body;
}

const W = Namespaces.NS_WORD12;
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';
const WNE = 'http://schemas.microsoft.com/office/word/2006/wordml';

/**
 * Shared by the main document and the glossary document (docx4j DocumentPart): the shortcuts
 * to styles, numbering, fonts, settings, theme, comments and notes.
 */
export abstract class DocumentPart<T> extends XmlPart<T> {
  styleDefinitionsPart: StyleDefinitionsPart | undefined;
  numberingDefinitionsPart: NumberingDefinitionsPart | undefined;
  fontTablePart: FontTablePart | undefined;
  themePart: ThemePart | undefined;
  documentSettingsPart: DocumentSettingsPart | undefined;
  webSettingsPart: WebSettingsPart | undefined;
  commentsPart: CommentsPart | undefined;
  commentsExtendedPart: CommentsExtendedPart | undefined;
  commentsIdsPart: CommentsIdsPart | undefined;
  commentsExtensiblePart: CommentsExtensiblePart | undefined;
  footnotesPart: FootnotesPart | undefined;
  endnotesPart: EndnotesPart | undefined;
  peoplePart: PeoplePart | undefined;

  override setPartShortcut(part: Part, relationshipType: string): boolean {
    switch (relationshipType) {
      case Namespaces.STYLES: this.styleDefinitionsPart = part as StyleDefinitionsPart; return true;
      case Namespaces.NUMBERING: this.numberingDefinitionsPart = part as NumberingDefinitionsPart; return true;
      case Namespaces.FONT_TABLE: this.fontTablePart = part as FontTablePart; return true;
      case Namespaces.THEME: this.themePart = part as ThemePart; return true;
      case Namespaces.SETTINGS: this.documentSettingsPart = part as DocumentSettingsPart; return true;
      case Namespaces.WEB_SETTINGS: this.webSettingsPart = part as WebSettingsPart; return true;
      case Namespaces.COMMENTS: this.commentsPart = part as CommentsPart; return true;
      case Namespaces.COMMENTS_EXTENDED: this.commentsExtendedPart = part as CommentsExtendedPart; return true;
      case Namespaces.COMMENTS_IDS: this.commentsIdsPart = part as CommentsIdsPart; return true;
      case Namespaces.COMMENTS_EXTENSIBLE: this.commentsExtensiblePart = part as CommentsExtensiblePart; return true;
      case Namespaces.FOOTNOTES: this.footnotesPart = part as FootnotesPart; return true;
      case Namespaces.ENDNOTES: this.endnotesPart = part as EndnotesPart; return true;
      case Namespaces.OFFICE_2011_PEOPLE: this.peoplePart = part as PeoplePart; return true;
      default: return false;
    }
  }

  /** Header parts related from this document, in relationship order. */
  get headerParts(): HeaderPart[] {
    return this.targets(Namespaces.HEADER) as HeaderPart[];
  }

  get footerParts(): FooterPart[] {
    return this.targets(Namespaces.FOOTER) as FooterPart[];
  }

  protected targets(relationshipType: string): Part[] {
    const rp = this.relationshipsPart;
    if (!rp) return [];
    return rp.getRelationshipsByType(relationshipType).map((r) => rp.getPart(r)).filter((p): p is Part => p !== undefined);
  }
}

/** `/word/document.xml` (docx4j MainDocumentPart). */
export class MainDocumentPart extends DocumentPart<wml.Document> {
  glossaryDocumentPart: GlossaryDocumentPart | undefined;

  constructor(partName: PartName | string = '/word/document.xml', contentType: string = ContentTypes.WORDPROCESSINGML_DOCUMENT) {
    super(partName, contentType, Namespaces.DOCUMENT, { namespaceURI: W, localPart: 'document' });
  }

  override setPartShortcut(part: Part, relationshipType: string): boolean {
    if (relationshipType === Namespaces.GLOSSARY_DOCUMENT) { this.glossaryDocumentPart = part as GlossaryDocumentPart; return true; }
    return super.setPartShortcut(part, relationshipType);
  }

  /** The Office JS-shaped view of `w:body`; the part must be unmarshalled (`getBody()` does that). */
  get body(): Body {
    const doc = this.contents;
    doc.body ??= { TYPE_NAME: 'org_docx4j_wml.Body' };
    return bodyOf(this, doc.body, 'body');
  }

  async getBody(): Promise<Body> {
    await this.getContents();
    await ensurePropertyResolver(this);
    return this.body;
  }
}

/** `/word/glossary/document.xml` (docx4j GlossaryDocumentPart). */
export class GlossaryDocumentPart extends DocumentPart<wml.GlossaryDocument> {
  constructor(partName: PartName | string = '/word/glossary/document.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_GLOSSARYDOCUMENT, Namespaces.GLOSSARY_DOCUMENT, { namespaceURI: W, localPart: 'glossaryDocument' });
  }
}

export class StyleDefinitionsPart extends XmlPart<wml.Styles> {
  constructor(partName: PartName | string = '/word/styles.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_STYLES, Namespaces.STYLES, { namespaceURI: W, localPart: 'styles' });
  }
}

export class NumberingDefinitionsPart extends XmlPart<wml.Numbering> {
  constructor(partName: PartName | string = '/word/numbering.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_NUMBERING, Namespaces.NUMBERING, { namespaceURI: W, localPart: 'numbering' });
  }
}

export class FontTablePart extends XmlPart<wml.Fonts> {
  constructor(partName: PartName | string = '/word/fontTable.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_FONTTABLE, Namespaces.FONT_TABLE, { namespaceURI: W, localPart: 'fonts' });
  }
}

export class DocumentSettingsPart extends XmlPart<wml.CTSettings> {
  constructor(partName: PartName | string = '/word/settings.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_SETTINGS, Namespaces.SETTINGS, { namespaceURI: W, localPart: 'settings' });
  }
}

export class WebSettingsPart extends XmlPart<wml.CTWebSettings> {
  constructor(partName: PartName | string = '/word/webSettings.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_WEBSETTINGS, Namespaces.WEB_SETTINGS, { namespaceURI: W, localPart: 'webSettings' });
  }
}

export class HeaderPart extends XmlPart<wml.Hdr> {
  constructor(partName: PartName | string = '/word/header1.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_HEADER, Namespaces.HEADER, { namespaceURI: W, localPart: 'hdr' });
  }
  /** The header's content as a Body; addresses are 'header:<relId>/n'. */
  get body(): Body {
    return bodyOf(this, this.contents, `header:${this.sourceRelationship?.id ?? this.partName.fileName}`);
  }
  async getBody(): Promise<Body> {
    await this.getContents();
    await ensurePropertyResolver(this);
    return this.body;
  }
}

export class FooterPart extends XmlPart<wml.Ftr> {
  constructor(partName: PartName | string = '/word/footer1.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_FOOTER, Namespaces.FOOTER, { namespaceURI: W, localPart: 'ftr' });
  }
  /** The footer's content as a Body; addresses are 'footer:<relId>/n'. */
  get body(): Body {
    return bodyOf(this, this.contents, `footer:${this.sourceRelationship?.id ?? this.partName.fileName}`);
  }
  async getBody(): Promise<Body> {
    await this.getContents();
    await ensurePropertyResolver(this);
    return this.body;
  }
}

export class FootnotesPart extends XmlPart<wml.CTFootnotes> {
  constructor(partName: PartName | string = '/word/footnotes.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_FOOTNOTES, Namespaces.FOOTNOTES, { namespaceURI: W, localPart: 'footnotes' });
  }
}

export class EndnotesPart extends XmlPart<wml.CTEndnotes> {
  constructor(partName: PartName | string = '/word/endnotes.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_ENDNOTES, Namespaces.ENDNOTES, { namespaceURI: W, localPart: 'endnotes' });
  }
}

export class CommentsPart extends XmlPart<wml.Comments> {
  constructor(partName: PartName | string = '/word/comments.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_COMMENTS, Namespaces.COMMENTS, { namespaceURI: W, localPart: 'comments' });
  }
}

export class CommentsExtendedPart extends XmlPart<w15.CTCommentsEx> {
  constructor(partName: PartName | string = '/word/commentsExtended.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_COMMENTS_EXTENDED, Namespaces.COMMENTS_EXTENDED, { namespaceURI: W15, localPart: 'commentsEx' });
  }
}

export class CommentsIdsPart extends XmlPart<w16cid.CTCommentsIds> {
  constructor(partName: PartName | string = '/word/commentsIds.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_COMMENTS_IDS, Namespaces.COMMENTS_IDS, { namespaceURI: W16CID, localPart: 'commentsIds' });
  }
}

/** `/word/commentsExtensible.xml` (w16cex, Word 2018): a durable id and a UTC date per comment. Typed since objects 0.1.3. */
export class CommentsExtensiblePart extends XmlPart<w16cex.CTCommentsExtensible> {
  constructor(partName: PartName | string = '/word/commentsExtensible.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_COMMENTS_EXTENSIBLE, Namespaces.COMMENTS_EXTENSIBLE, { namespaceURI: W16CEX, localPart: 'commentsExtensible' });
  }
}

export class PeoplePart extends XmlPart<w15.CTPeople> {
  constructor(partName: PartName | string = '/word/people.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_PEOPLE, Namespaces.OFFICE_2011_PEOPLE, { namespaceURI: W15, localPart: 'people' });
  }
}

export class KeyMapCustomizationsPart extends XmlPart<wne.CTTcg> {
  constructor(partName: PartName | string = '/word/customizations.xml') {
    super(partName, ContentTypes.MS_WORD_KEYMAP, Namespaces.KEYMAP, { namespaceURI: WNE, localPart: 'tcg' });
  }
}

export class VbaDataPart extends XmlPart<wne.CTVbaSuppData> {
  constructor(partName: PartName | string = '/word/vbaData.xml') {
    super(partName, ContentTypes.OFFICEDOCUMENT_VBA_DATA, Namespaces.VBA_DATA_WORD, { namespaceURI: WNE, localPart: 'vbaSuppData' });
  }
}

/** A VML drawing (`/word/vmlDrawing1.vml`, docx4j VMLPart), as a DOM: the root `<xml>` is docx4j's own wrapper, not typed here yet. */
export class VMLPart extends DefaultXmlPart {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.VML_DRAWING, Namespaces.VML);
  }
}

export { ThemePart };
