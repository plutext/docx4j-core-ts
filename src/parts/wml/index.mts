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
import type { PropertyResolver } from '../../model/properties/PropertyResolver.mjs';
import type { AbstractListDefinition, ListDefinition } from '../../model/listnumbering/definitions.mjs';
import { NumberingDefinitions } from '../../model/listnumbering/definitions.mjs';
import { Emulator } from '../../model/listnumbering/Emulator.mjs';
import type { NumberingState } from '../../model/listnumbering/state.mjs';
import { DEFAULT_NUMBERING_XML } from './defaultNumbering.mjs';
import { RunFontSelector } from '../../model/fonts/RunFontSelector.mjs';
import { IdentityPlusMapper } from '../../model/fonts/IdentityPlusMapper.mjs';
import type { Mapper } from '../../model/fonts/Mapper.mjs';
import { themeOfSettings, type FontSettings } from '../../model/fonts/defaultTheme.mjs';
import { fontsInUse, walkStories, type StyleSource } from '../../model/fonts/fontsInUse.mjs';

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

  // --- fonts (CR-001 Phase B step 4) ------------------------------------------------------

  private fontSelector: RunFontSelector | undefined;
  private fontSelectorPending: Promise<RunFontSelector> | undefined;
  private fontMapper: Mapper | undefined;

  /**
   * The package's `RunFontSelector`, built once: the property resolver, the theme part's font
   * scheme, the settings part's `w:themeFontLang` and the package's `fonts.defaultTheme`.
   *
   * Asynchronous because it reads the theme and settings parts; `runFontSelector` is the
   * synchronous accessor afterwards. Both parts are read with `readContents()`, so a document
   * whose theme and settings nobody has edited still saves byte for byte.
   */
  async getRunFontSelector(): Promise<RunFontSelector> {
    if (this.fontSelector) return this.fontSelector;
    this.fontSelectorPending ??= this.buildRunFontSelector().then((selector) => {
      this.fontSelector = selector;
      this.fontSelectorPending = undefined;
      return selector;
    });
    return this.fontSelectorPending;
  }

  private async buildRunFontSelector(): Promise<RunFontSelector> {
    const theme = this.themePart === undefined ? undefined : await readQuietly(this.themePart);
    const settings = this.documentSettingsPart === undefined
      ? undefined : await readQuietly(this.documentSettingsPart);
    const pkg = this.package as { propertyResolverOrUndefined?: PropertyResolver; fonts?: FontSettings } | undefined;
    return new RunFontSelector({
      resolver: pkg?.propertyResolverOrUndefined,
      theme,
      themeFontLang: settings?.themeFontLang,
      defaultTheme: themeOfSettings(pkg?.fonts),
    });
  }

  /** The selector, which must already have been built (`await getRunFontSelector()`). */
  get runFontSelector(): RunFontSelector {
    if (!this.fontSelector) {
      throw new Docx4JException('No RunFontSelector yet: await mainDocumentPart.getRunFontSelector()');
    }
    return this.fontSelector;
  }

  /** The selector if it has been built, without throwing (what a read with a fallback wants). */
  get runFontSelectorOrUndefined(): RunFontSelector | undefined {
    return this.fontSelector;
  }

  /**
   * The package's font mapper, populated from {@link fontsInUse} and the font table part
   * (docx4j `WordprocessingMLPackage.getFontMapper` / `setFontMapper`): an
   * `IdentityPlusMapper` over the default font registry unless one is given.
   *
   * @param mapper a mapper of the caller's own, which is then populated and kept
   */
  async getFontMapper(mapper?: Mapper): Promise<Mapper> {
    if (mapper === undefined && this.fontMapper) return this.fontMapper;
    const m = mapper ?? new IdentityPlusMapper();
    m.populate(await this.fontsInUse(),
      this.fontTablePart === undefined ? undefined : await readQuietly(this.fontTablePart));
    this.fontMapper = m;
    return m;
  }

  /** Discard the selector and the mapper: the styles, theme or settings changed under us. */
  refreshFonts(): void {
    this.fontSelector = undefined;
    this.fontSelectorPending = undefined;
    this.fontMapper = undefined;
  }

  /**
   * The document font names, as docx4j's `fontsInUse()`: the four slots of every `w:rFonts` on
   * the runs, paragraph marks and content controls of the body, headers, footers, footnotes,
   * endnotes and comments; the styles in use with their `w:basedOn` chains; the numbering
   * levels; the document defaults; and the default font. Theme references are resolved, and a
   * CJK font name is collected by its English name.
   */
  async fontsInUse(): Promise<Set<string>> {
    const selector = await this.getRunFontSelector();
    const numbering = this.numberingDefinitionsPart === undefined
      ? undefined : await readQuietly(this.numberingDefinitionsPart);
    return fontsInUse(await this.stories(), await this.styleSource(), numbering, selector);
  }

  /** The style ids used directly in the document (docx4j `getStylesInUse()`); not those that
   *  others are merely based on. */
  async getStylesInUse(): Promise<Set<string>> {
    return walkStories(await this.stories(), await this.styleSource(), undefined, false).styles;
  }

  /** The block-level content of each story, in docx4j's order: the body, then the headers and
   *  footers, the endnotes, the footnotes and the comments. */
  private async stories(): Promise<unknown[]> {
    const out: unknown[] = [(await this.getContents()).body?.content ?? []];
    const rp = this.relationshipsPart;
    for (const r of rp?.relationships.relationship ?? []) {
      const part = rp?.getPart(r);
      if (part instanceof HeaderPart) out.push((await readQuietly(part))?.content ?? []);
      else if (part instanceof FooterPart) out.push((await readQuietly(part))?.content ?? []);
    }
    if (this.endnotesPart) out.push((await readQuietly(this.endnotesPart))?.endnote ?? []);
    if (this.footnotesPart) out.push((await readQuietly(this.footnotesPart))?.footnote ?? []);
    if (this.commentsPart) out.push((await readQuietly(this.commentsPart))?.comment ?? []);
    return out;
  }

  /** The styles part and its `w:default="1"` style ids, read without costing it its round trip. */
  private async styleSource(): Promise<StyleSource> {
    const styles = this.styleDefinitionsPart === undefined
      ? undefined : await readQuietly(this.styleDefinitionsPart);
    const source: StyleSource = { styles };
    for (const style of styles?.style ?? []) {
      if (style._default !== true || style.styleId === undefined) continue;
      if (style.type === 'paragraph') source.defaultParagraphStyleId ??= style.styleId;
      else if (style.type === 'character') source.defaultCharacterStyleId ??= style.styleId;
      else if (style.type === 'table') source.defaultTableStyleId ??= style.styleId;
    }
    return source;
  }
}

/** `readContents()`, answering undefined for a part that cannot be read: a document with an
 *  unreadable theme or settings part still resolves its fonts, as docx4j's selector does. */
async function readQuietly<T>(part: XmlPart<T>): Promise<T | undefined> {
  try {
    return await part.readContents();
  } catch {
    return undefined;
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

/**
 * `/word/numbering.xml` (docx4j NumberingDefinitionsPart): the list definitions, the indent a
 * level contributes, and the `Emulator` that turns a paragraph into a label.
 *
 * The definitions are read **privately** (`readContents()`), as the `PropertyResolver` reads the
 * styles, so that a document whose labels are only read still saves byte for byte; where the
 * package already has a resolver the definitions it built are reused, since it needs them for
 * `getInd`. A change to the part is picked up by `WordprocessingMLPackage.refresh()`.
 */
export class NumberingDefinitionsPart extends XmlPart<wml.Numbering> {
  constructor(partName: PartName | string = '/word/numbering.xml') {
    super(partName, ContentTypes.WORDPROCESSINGML_NUMBERING, Namespaces.NUMBERING, { namespaceURI: W, localPart: 'numbering' });
  }

  /** The tree `readContents()` gave, kept so the definitions can be rebuilt without re-reading. */
  private privateTree: wml.Numbering | undefined;
  private ownDefinitions: NumberingDefinitions | undefined;
  private ownDefinitionsFrom: wml.Numbering | undefined;
  private emulator: Emulator | undefined;

  /** The resolver of the package this part belongs to, if it has one built. */
  private get resolver(): PropertyResolver | undefined {
    return (this.package as { propertyResolverOrUndefined?: PropertyResolver } | undefined)?.propertyResolverOrUndefined;
  }

  /** The `w:numbering` the definitions are built over: the live tree, or the private read. */
  private get tree(): wml.Numbering | undefined {
    return this.isUnmarshalled ? this.contents : this.privateTree;
  }

  /**
   * docx4j `initialiseMaps` and the two maps it fills: every `w:abstractNum` and every `w:num`,
   * with `w:numStyleLink` resolved. Synchronous, so the part's tree must have been read - by
   * {@link getDefinitions}, or by the package's `getPropertyResolver()`.
   */
  get definitions(): NumberingDefinitions {
    const live = this.tree;
    const resolver = this.resolver;
    const shared = resolver?.getNumberingDefinitions();
    // the resolver's copy, unless the part has been unmarshalled and edited since it read it
    if (shared !== undefined && (!this.isUnmarshalled || shared.numbering === live)) return shared;
    if (live === undefined) {
      throw new Docx4JException(`${this.partName} has not been read yet; await getDefinitions() first`);
    }
    if (this.ownDefinitions === undefined || this.ownDefinitionsFrom !== live) {
      this.ownDefinitions = new NumberingDefinitions(live, (styleId) => this.resolver?.getStyle(styleId));
      this.ownDefinitionsFrom = live;
    }
    return this.ownDefinitions;
  }

  /** {@link definitions}, reading the part first where nothing has. */
  async getDefinitions(): Promise<NumberingDefinitions> {
    if (!this.isUnmarshalled && this.privateTree === undefined) this.privateTree = await this.readContents();
    return this.definitions;
  }

  /** `w:abstractNumId` -> the definition. docx4j `getAbstractListDefinitions()`. */
  get abstractListDefinitions(): ReadonlyMap<string, AbstractListDefinition> {
    return this.definitions.abstractListDefinitions;
  }

  /** `w:numId` -> the definition. docx4j `getInstanceListDefinitions()`. */
  get instanceListDefinitions(): ReadonlyMap<string, ListDefinition> {
    return this.definitions.instanceListDefinitions;
  }

  /** The indent a `w:numPr`'s level contributes. docx4j `getInd(NumPr)`. */
  getInd(numPr: wml.PPrBase.NumPr): wml.PPrBase.Ind | undefined {
    return this.definitions.getInd(numPr);
  }

  /** The indent a level contributes. docx4j `getInd(String, String)`. */
  getIndOf(numId: string, ilvl: string | undefined): wml.PPrBase.Ind | undefined {
    return this.definitions.indOf(numId, ilvl);
  }

  /** The paragraph style a level is linked to (ECMA-376 17.9.24), or undefined. */
  getLinkedStyleId(numId: string, ilvl: string | undefined): string | undefined {
    return this.definitions.getLinkedStyleId(numId, ilvl);
  }

  /**
   * docx4j `getEmulator([reset])`: the numbering emulator over this part, built once. `reset`
   * gives it a fresh default {@link NumberingState} and re-reads the definitions, as docx4j's
   * does - every list starts again.
   */
  getEmulator(reset = false): Emulator {
    if (reset) {
      this.ownDefinitions = undefined;
      this.ownDefinitionsFrom = undefined;
      this.emulator = undefined;
    }
    const definitions = this.definitions;
    if (this.emulator === undefined || this.emulator.definitions !== definitions
      || this.emulator.resolver !== this.resolver) {
      this.emulator = new Emulator(definitions, this.resolver,
        this.emulator === undefined ? undefined : this.emulator.numberingState);
    }
    return this.emulator;
  }

  /** The counters the state-less `Emulator` calls use. docx4j `getNumberingState()`. */
  get numberingState(): NumberingState {
    return this.getEmulator().numberingState;
  }

  /**
   * Replaces this part's contents with docx4j's default numbering (`numbering.xml` from its
   * resources): abstractNum 0 is Word's bullet set and abstractNum 1 its decimal set, with
   * `w:num` 1 naming the decimal one and `w:num` 2 the bullet one.
   * docx4j `unmarshalDefaultNumbering()`.
   */
  async unmarshalDefaultNumbering(): Promise<wml.Numbering> {
    this.setXml(DEFAULT_NUMBERING_XML);
    this.privateTree = undefined;
    this.ownDefinitions = undefined;
    this.ownDefinitionsFrom = undefined;
    this.emulator = undefined;
    return this.getContents();
  }

  /** Forgets the definitions and the emulator: the tree has changed under us. */
  refreshDefinitions(): void {
    this.privateTree = undefined;
    this.ownDefinitions = undefined;
    this.ownDefinitionsFrom = undefined;
    this.emulator = undefined;
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
