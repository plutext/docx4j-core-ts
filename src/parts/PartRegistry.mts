import type { Relationship } from '@docx4j/generated-objects-ts/modules/org_docx4j_relationships';
import { Part } from './Part.mjs';
import { PartName } from '../opc/PartName.mjs';
import { ContentTypes, IMAGE_CONTENT_TYPES_BY_EXTENSION, isXmlContentType } from '../opc/ContentTypes.mjs';
import { Namespaces } from './Namespaces.mjs';
import { BinaryPart, ImagePart, EmbeddedPackagePart, OleObjectBinaryPart, ObfuscatedFontPart, TrueTypeFontPart, AlternativeFormatInputPart } from './BinaryPart.mjs';
import { DefaultXmlPart, CustomXmlDataStoragePart } from './DefaultXmlPart.mjs';
import * as wml from './wml/index.mjs';
import * as dml from './dml/index.mjs';
import * as pml from './pml/index.mjs';
import * as sml from './sml/index.mjs';
import * as docProps from './docProps/index.mjs';
import * as customXml from './customXml/index.mjs';

/** Creates a part for a name and content type. */
export type PartFactory = (partName: PartName, contentType: string) => Part;

/**
 * Content type to part class (docx4j ContentTypeManager.newPartForContentType), with a first
 * look at the relationship type for the generic content types (altChunk, embedded package, OLE,
 * custom XML). Unknown XML becomes `DefaultXmlPart`, anything else `BinaryPart`.
 */
export class PartRegistry {
  private readonly byContentType = new Map<string, PartFactory>();
  private readonly byRelationshipType = new Map<string, PartFactory>();

  register(contentType: string, factory: PartFactory): this {
    this.byContentType.set(contentType, factory);
    return this;
  }

  /** Consulted before the content type; for relationship types whose parts share a generic content type. */
  registerForRelationship(relationshipType: string, factory: PartFactory): this {
    this.byRelationshipType.set(relationshipType, factory);
    return this;
  }

  createPart(partName: PartName, contentType: string | undefined, rel?: Relationship): Part {
    const relType = rel?.type;
    if (relType) {
      const byRel = this.byRelationshipType.get(relType);
      if (byRel) return byRel(partName, contentType ?? ContentTypes.APPLICATION_XML);
    }
    if (contentType !== undefined) {
      const factory = this.byContentType.get(contentType);
      if (factory) return factory(partName, contentType);
      if (contentType.startsWith('image/')) return new ImagePart(partName, contentType);
      if (isXmlContentType(contentType) || partName.extension.toLowerCase() === 'xml') {
        return new DefaultXmlPart(partName, contentType, relType ?? '');
      }
      return new BinaryPart(partName, contentType, relType ?? '');
    }
    // No content type registered: Word would reject the package; keep the bytes anyway.
    if (relType === Namespaces.IMAGE) {
      return new ImagePart(partName, IMAGE_CONTENT_TYPES_BY_EXTENSION[partName.extension.toLowerCase()] ?? ContentTypes.OCTET_STREAM);
    }
    if (partName.extension.toLowerCase() === 'xml') return new DefaultXmlPart(partName, ContentTypes.APPLICATION_XML, relType ?? '');
    return new BinaryPart(partName, ContentTypes.OCTET_STREAM, relType ?? '');
  }
}

/** docx4j's table. Extend with `register` for content types of your own. */
export const defaultPartRegistry = new PartRegistry()
  .registerForRelationship(Namespaces.AF, (n, ct) => new AlternativeFormatInputPart(n, ct))
  .registerForRelationship(Namespaces.EMBEDDED_PKG, (n, ct) => new EmbeddedPackagePart(n, ct))
  .registerForRelationship(Namespaces.OLE_OBJECT, (n, ct) => new OleObjectBinaryPart(n, ct))
  .registerForRelationship(Namespaces.CUSTOM_XML_DATA_STORAGE, (n, ct) => new CustomXmlDataStoragePart(n, ct))
  .register(ContentTypes.WORDPROCESSINGML_DOCUMENT, (n, ct) => new wml.MainDocumentPart(n, ct))
  .register(ContentTypes.WORDPROCESSINGML_DOCUMENT_MACROENABLED, (n, ct) => new wml.MainDocumentPart(n, ct))
  .register(ContentTypes.WORDPROCESSINGML_TEMPLATE, (n, ct) => new wml.MainDocumentPart(n, ct))
  .register(ContentTypes.WORDPROCESSINGML_TEMPLATE_MACROENABLED, (n, ct) => new wml.MainDocumentPart(n, ct))
  .register(ContentTypes.WORDPROCESSINGML_GLOSSARYDOCUMENT, (n) => new wml.GlossaryDocumentPart(n))
  .register(ContentTypes.WORDPROCESSINGML_STYLES, (n) => new wml.StyleDefinitionsPart(n))
  .register(ContentTypes.WORDPROCESSINGML_NUMBERING, (n) => new wml.NumberingDefinitionsPart(n))
  .register(ContentTypes.WORDPROCESSINGML_FONTTABLE, (n) => new wml.FontTablePart(n))
  .register(ContentTypes.WORDPROCESSINGML_SETTINGS, (n) => new wml.DocumentSettingsPart(n))
  .register(ContentTypes.WORDPROCESSINGML_WEBSETTINGS, (n) => new wml.WebSettingsPart(n))
  .register(ContentTypes.WORDPROCESSINGML_HEADER, (n) => new wml.HeaderPart(n))
  .register(ContentTypes.WORDPROCESSINGML_FOOTER, (n) => new wml.FooterPart(n))
  .register(ContentTypes.WORDPROCESSINGML_FOOTNOTES, (n) => new wml.FootnotesPart(n))
  .register(ContentTypes.WORDPROCESSINGML_ENDNOTES, (n) => new wml.EndnotesPart(n))
  .register(ContentTypes.WORDPROCESSINGML_COMMENTS, (n) => new wml.CommentsPart(n))
  .register(ContentTypes.WORDPROCESSINGML_COMMENTS_EXTENDED, (n) => new wml.CommentsExtendedPart(n))
  .register(ContentTypes.WORDPROCESSINGML_COMMENTS_IDS, (n) => new wml.CommentsIdsPart(n))
  .register(ContentTypes.WORDPROCESSINGML_PEOPLE, (n) => new wml.PeoplePart(n))
  .register(ContentTypes.MS_WORD_KEYMAP, (n) => new wml.KeyMapCustomizationsPart(n))
  .register(ContentTypes.OFFICEDOCUMENT_VBA_DATA, (n) => new wml.VbaDataPart(n))
  .register(ContentTypes.VML_DRAWING, (n) => new wml.VMLPart(n))
  .register(ContentTypes.OFFICEDOCUMENT_THEME, (n) => new dml.ThemePart(n))
  .register(ContentTypes.OFFICEDOCUMENT_THEME_OVERRIDE, (n) => new dml.ThemeOverridePart(n))
  .register(ContentTypes.DRAWINGML_CHART, (n) => new dml.ChartPart(n))
  .register(ContentTypes.DRAWINGML_CHART_SHAPES, (n) => new dml.ChartShapePart(n))
  .register(ContentTypes.DRAWINGML_DRAWING, (n) => new dml.DrawingPart(n))
  .register(ContentTypes.DRAWINGML_DIAGRAM_DATA, (n) => new dml.DiagramDataPart(n))
  .register(ContentTypes.DRAWINGML_DIAGRAM_LAYOUT, (n) => new dml.DiagramLayoutPart(n))
  .register(ContentTypes.DRAWINGML_DIAGRAM_LAYOUT_HEADER, (n) => new dml.DiagramLayoutHeaderPart(n))
  .register(ContentTypes.DRAWINGML_DIAGRAM_STYLE, (n) => new dml.DiagramStylePart(n))
  .register(ContentTypes.DRAWINGML_DIAGRAM_COLORS, (n) => new dml.DiagramColorsPart(n))
  .register(ContentTypes.DRAWINGML_DIAGRAM_DRAWING, (n) => new dml.DiagramDrawingPart(n))
  .register(ContentTypes.CHART_STYLE, (n) => new dml.ChartStylePart(n))
  .register(ContentTypes.CHART_COLOR_STYLE, (n) => new dml.ChartColorStylePart(n))
  .register(ContentTypes.CHART_EX, (n) => new dml.ChartExSpacePart(n))
  .register(ContentTypes.PACKAGE_COREPROPERTIES, (n) => new docProps.DocPropsCorePart(n))
  .register(ContentTypes.OFFICEDOCUMENT_EXTENDEDPROPERTIES, (n) => new docProps.DocPropsExtendedPart(n))
  .register(ContentTypes.OFFICEDOCUMENT_CUSTOMPROPERTIES, (n) => new docProps.DocPropsCustomPart(n))
  .register(ContentTypes.OFFICEDOCUMENT_CUSTOMXML_DATASTORAGEPROPERTIES, (n) => new customXml.CustomXmlDataStoragePropertiesPart(n))
  .register(ContentTypes.OFFICEDOCUMENT_FONT, (n) => new ObfuscatedFontPart(n))
  .register(ContentTypes.TRUETYPE_FONT, (n) => new TrueTypeFontPart(n))
  .register(ContentTypes.OFFICEDOCUMENT_OLE_OBJECT, (n, ct) => new OleObjectBinaryPart(n, ct))
  .register(ContentTypes.OFFICEDOCUMENT_ACTIVEX_OBJECT, (n, ct) => new OleObjectBinaryPart(n, ct))
  .register(ContentTypes.PRESENTATIONML_MAIN, (n, ct) => new pml.MainPresentationPart(n, ct))
  .register(ContentTypes.PRESENTATIONML_TEMPLATE, (n, ct) => new pml.MainPresentationPart(n, ct))
  .register(ContentTypes.PRESENTATIONML_MACROENABLED, (n, ct) => new pml.MainPresentationPart(n, ct))
  .register(ContentTypes.PRESENTATIONML_TEMPLATE_MACROENABLED, (n, ct) => new pml.MainPresentationPart(n, ct))
  .register(ContentTypes.PRESENTATIONML_SLIDESHOW, (n, ct) => new pml.MainPresentationPart(n, ct))
  .register(ContentTypes.PRESENTATIONML_SLIDE, (n) => new pml.SlidePart(n))
  .register(ContentTypes.PRESENTATIONML_SLIDE_LAYOUT, (n) => new pml.SlideLayoutPart(n))
  .register(ContentTypes.PRESENTATIONML_SLIDE_MASTER, (n) => new pml.SlideMasterPart(n))
  .register(ContentTypes.PRESENTATIONML_NOTES_SLIDE, (n) => new pml.NotesSlidePart(n))
  .register(ContentTypes.PRESENTATIONML_NOTES_MASTER, (n) => new pml.NotesMasterPart(n))
  .register(ContentTypes.PRESENTATIONML_HANDOUT_MASTER, (n) => new pml.HandoutMasterPart(n))
  .register(ContentTypes.PRESENTATIONML_PRES_PROPS, (n) => new pml.PresentationPropertiesPart(n))
  .register(ContentTypes.PRESENTATIONML_VIEW_PROPS, (n) => new pml.ViewPropertiesPart(n))
  .register(ContentTypes.PRESENTATIONML_TABLE_STYLES, (n) => new pml.TableStylesPart(n))
  .register(ContentTypes.PRESENTATIONML_COMMENTS, (n) => new pml.PresentationCommentsPart(n))
  .register(ContentTypes.PRESENTATIONML_COMMENT_AUTHORS, (n) => new pml.CommentAuthorsPart(n))
  .register(ContentTypes.PRESENTATIONML_TAGS, (n) => new pml.TagsPart(n))
  .register(ContentTypes.PRESENTATIONML_FONT_DATA, (n) => new pml.FontDataPart(n))
  .register(ContentTypes.SPREADSHEETML_WORKBOOK, (n, ct) => new sml.WorkbookPart(n, ct))
  .register(ContentTypes.SPREADSHEETML_WORKBOOK_MACROENABLED, (n, ct) => new sml.WorkbookPart(n, ct))
  .register(ContentTypes.SPREADSHEETML_TEMPLATE, (n, ct) => new sml.WorkbookPart(n, ct))
  .register(ContentTypes.SPREADSHEETML_TEMPLATE_MACROENABLED, (n, ct) => new sml.WorkbookPart(n, ct))
  .register(ContentTypes.SPREADSHEETML_WORKSHEET, (n) => new sml.WorksheetPart(n))
  .register(ContentTypes.SPREADSHEETML_CHARTSHEET, (n) => new sml.ChartsheetPart(n))
  .register(ContentTypes.SPREADSHEETML_SHARED_STRINGS, (n) => new sml.SharedStringsPart(n))
  .register(ContentTypes.SPREADSHEETML_STYLES, (n) => new sml.StylesPart(n))
  .register(ContentTypes.SPREADSHEETML_CALC_CHAIN, (n) => new sml.CalcChainPart(n))
  .register(ContentTypes.SPREADSHEETML_COMMENTS, (n) => new sml.SpreadsheetCommentsPart(n))
  .register(ContentTypes.SPREADSHEETML_TABLE, (n) => new sml.TablePart(n))
  .register(ContentTypes.SPREADSHEETML_PIVOT_TABLE, (n) => new sml.PivotTablePart(n))
  .register(ContentTypes.SPREADSHEETML_PIVOT_CACHE_DEFINITION, (n) => new sml.PivotCacheDefinitionPart(n))
  .register(ContentTypes.SPREADSHEETML_PIVOT_CACHE_RECORDS, (n) => new sml.PivotCacheRecordsPart(n))
  .register(ContentTypes.SPREADSHEETML_QUERY_TABLE, (n) => new sml.QueryTablePart(n))
  .register(ContentTypes.SPREADSHEETML_CONNECTIONS, (n) => new sml.ConnectionsPart(n))
  .register(ContentTypes.SPREADSHEETML_EXTERNAL_LINK, (n) => new sml.ExternalLinkPart(n))
  .register(ContentTypes.SPREADSHEETML_PRINTER_SETTINGS, (n) => new sml.PrinterSettingsPart(n));
