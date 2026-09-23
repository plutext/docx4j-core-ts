// SpreadsheetML parts (docx4j org.docx4j.openpackaging.parts.SpreadsheetML).
import type * as sml from '@docx4j/generated-objects-ts/modules/org_xlsx4j_sml';
import { XmlPart } from '../XmlPart.mjs';
import { DefaultXmlPart } from '../DefaultXmlPart.mjs';
import { BinaryPart } from '../BinaryPart.mjs';
import { Part } from '../Part.mjs';
import { PartName } from '../../opc/PartName.mjs';
import { ContentTypes } from '../../opc/ContentTypes.mjs';
import { Namespaces } from '../Namespaces.mjs';
import { Xlsx4jException } from '../../opc/exceptions.mjs';
import { ThemePart } from '../dml/index.mjs';

const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export class WorkbookPart extends XmlPart<sml.Workbook> {
  sharedStringsPart: SharedStringsPart | undefined;
  stylesPart: StylesPart | undefined;
  themePart: ThemePart | undefined;
  calcChainPart: CalcChainPart | undefined;
  /** The workbook's survey part, if it has one (CR-004). */
  surveyPart: SurveyPart | undefined;
  /** The Power Pivot data model (`/xl/model/item.data`), if the workbook has one (CR-004). */
  dataModelPart: DataModelPart | undefined;
  constructor(partName: PartName | string = '/xl/workbook.xml', contentType: string = ContentTypes.SPREADSHEETML_WORKBOOK) {
    super(partName, contentType, Namespaces.SPREADSHEETML_WORKBOOK, { namespaceURI: S, localPart: 'workbook' });
  }
  override setPartShortcut(part: Part, relationshipType: string): boolean {
    switch (relationshipType) {
      case Namespaces.SPREADSHEETML_SHARED_STRINGS: this.sharedStringsPart = part as SharedStringsPart; return true;
      case Namespaces.SPREADSHEETML_STYLES: this.stylesPart = part as StylesPart; return true;
      case Namespaces.THEME: this.themePart = part as ThemePart; return true;
      case Namespaces.SPREADSHEETML_CALC_CHAIN: this.calcChainPart = part as CalcChainPart; return true;
      case Namespaces.SPREADSHEETML_SURVEY: this.surveyPart = part as SurveyPart; return true;
      case Namespaces.SPREADSHEETML_DATA_MODEL: this.dataModelPart = part as DataModelPart; return true;
      default: return false;
    }
  }
  /**
   * The worksheets, in `sheets` order (the order Excel shows the tabs in) once the workbook is
   * unmarshalled; before that, in relationship order.
   */
  get worksheetParts(): WorksheetPart[] {
    const rp = this.relationshipsPart;
    if (!rp) return [];
    if (this.isUnmarshalled) {
      const sheets = this.contents.sheets?.sheet;
      if (sheets) {
        const out: WorksheetPart[] = [];
        for (const sheet of sheets) {
          const part = rp.getPart(sheet.id);
          if (part instanceof WorksheetPart) out.push(part);
        }
        return out;
      }
    }
    return rp.getRelationshipsByType(Namespaces.SPREADSHEETML_WORKSHEET).map((r) => rp.getPart(r)).filter((p): p is WorksheetPart => p instanceof WorksheetPart);
  }

  /** The worksheets in `sheets` order, unmarshalling the workbook first. */
  async getWorksheetParts(): Promise<WorksheetPart[]> {
    await this.getContents();
    return this.worksheetParts;
  }

  /** The worksheet at an index of `sheets` (docx4j `WorkbookPart.getWorksheet(int)`). */
  getWorksheet(index: number): WorksheetPart {
    const sheets = this.worksheetParts;
    const sheet = sheets[index];
    if (!sheet) throw new Xlsx4jException(`No sheet at index ${index}. (There are ${sheets.length} sheets) `);
    return sheet;
  }

  /** The workbook's slicer caches (`/xl/slicerCaches/slicerCacheN.xml`), in relationship order. */
  get slicerCacheParts(): SlicerCachePart[] {
    return this.partsByRelationshipType(Namespaces.SPREADSHEETML_SLICER_CACHE, SlicerCachePart);
  }

  /** The workbook's timeline caches (`/xl/timelineCaches/timelineCacheN.xml`), in relationship order. */
  get timelineCacheParts(): TimelineCachePart[] {
    return this.partsByRelationshipType(Namespaces.SPREADSHEETML_TIMELINE_CACHE, TimelineCachePart);
  }

  /** The workbook's custom data properties parts, in relationship order; each names a `CustomDataPart`. */
  get customDataPropertiesParts(): CustomDataPropertiesPart[] {
    return this.partsByRelationshipType(Namespaces.SPREADSHEETML_CUSTOM_DATA_PROPERTIES, CustomDataPropertiesPart);
  }

  /** docx4j `WorkbookPart.isDate1904()`. */
  get isDate1904(): boolean {
    return this.contents.workbookPr?.date1904 === true;
  }
}

export class WorksheetPart extends XmlPart<sml.Worksheet> {
  /** The sheet's drawing part (docx4j's `Drawing`), set from the `drawing` relationship. */
  drawingPart: Part | undefined;
  /** The sheet's legacy comments part. */
  commentsPart: SpreadsheetCommentsPart | undefined;
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.SPREADSHEETML_WORKSHEET, Namespaces.SPREADSHEETML_WORKSHEET, { namespaceURI: S, localPart: 'worksheet' });
  }
  override setPartShortcut(part: Part, relationshipType: string): boolean {
    switch (relationshipType) {
      case Namespaces.SPREADSHEETML_DRAWING: this.drawingPart = part; return true;
      case Namespaces.SPREADSHEETML_COMMENTS: this.commentsPart = part as SpreadsheetCommentsPart; return true;
      default: return false;
    }
  }
  /** The sheet's table parts (`xl/tables/tableN.xml`), in relationship order. */
  get tableParts(): TablePart[] {
    return this.partsByRelationshipType(Namespaces.SPREADSHEETML_TABLE, TablePart);
  }
  /** The sheet's slicers parts (`/xl/slicers/slicerN.xml`), in relationship order (CR-004). */
  get slicersParts(): SlicersPart[] {
    return this.partsByRelationshipType(Namespaces.SPREADSHEETML_SLICERS, SlicersPart);
  }

  /** The sheet's timelines parts (`/xl/timelines/timelineN.xml`), in relationship order (CR-004). */
  get timelinesParts(): TimelinesPart[] {
    return this.partsByRelationshipType(Namespaces.SPREADSHEETML_TIMELINES, TimelinesPart);
  }

  /**
   * The sheet's form-control properties parts (`/xl/ctrlProps/ctrlPropN.xml`), in relationship
   * order: one per `x:control`, whose `r:id` names it from inside the sheet's
   * `mc:AlternateContent` (CR-004).
   */
  get controlPropertiesParts(): ControlPropertiesPart[] {
    return this.partsByRelationshipType(Namespaces.SPREADSHEETML_CONTROL_PROPERTIES, ControlPropertiesPart);
  }

  /** The workbook this sheet belongs to, through its source relationship. */
  get workbookPart(): WorkbookPart | undefined {
    const owner = this.owningRelationshipPart?.sourceP;
    return owner instanceof WorkbookPart ? owner : undefined;
  }
}

export class ChartsheetPart extends XmlPart<sml.CTChartsheet> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.SPREADSHEETML_CHARTSHEET, Namespaces.SPREADSHEETML_CHARTSHEET, { namespaceURI: S, localPart: 'chartsheet' });
  }
}

export class SharedStringsPart extends XmlPart<sml.CTSst> {
  constructor(partName: PartName | string = '/xl/sharedStrings.xml') {
    super(partName, ContentTypes.SPREADSHEETML_SHARED_STRINGS, Namespaces.SPREADSHEETML_SHARED_STRINGS, { namespaceURI: S, localPart: 'sst' });
  }
}

/** docx4j SpreadsheetML.Styles. */
export class StylesPart extends XmlPart<sml.CTStylesheet> {
  constructor(partName: PartName | string = '/xl/styles.xml') {
    super(partName, ContentTypes.SPREADSHEETML_STYLES, Namespaces.SPREADSHEETML_STYLES, { namespaceURI: S, localPart: 'styleSheet' });
  }
}

/** docx4j SpreadsheetML.CalcChain. */
export class CalcChainPart extends XmlPart<sml.CTCalcChain> {
  constructor(partName: PartName | string = '/xl/calcChain.xml') {
    super(partName, ContentTypes.SPREADSHEETML_CALC_CHAIN, Namespaces.SPREADSHEETML_CALC_CHAIN, { namespaceURI: S, localPart: 'calcChain' });
  }
}

/** docx4j SpreadsheetML.CommentsPart, renamed here to avoid the clash with the WordprocessingML one. */
export class SpreadsheetCommentsPart extends XmlPart<sml.CTComments> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.SPREADSHEETML_COMMENTS, Namespaces.SPREADSHEETML_COMMENTS, { namespaceURI: S, localPart: 'comments' });
  }
}

export class TablePart extends XmlPart<sml.CTTable> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.SPREADSHEETML_TABLE, Namespaces.SPREADSHEETML_TABLE, { namespaceURI: S, localPart: 'table' });
  }
}

export class PivotTablePart extends XmlPart<sml.CTPivotTableDefinition> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.SPREADSHEETML_PIVOT_TABLE, Namespaces.SPREADSHEETML_PIVOT_TABLE, { namespaceURI: S, localPart: 'pivotTableDefinition' });
  }
}

export class PivotCacheDefinitionPart extends XmlPart<sml.CTPivotCacheDefinition> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.SPREADSHEETML_PIVOT_CACHE_DEFINITION, Namespaces.SPREADSHEETML_PIVOT_CACHE_DEFINITION, { namespaceURI: S, localPart: 'pivotCacheDefinition' });
  }
}

export class PivotCacheRecordsPart extends XmlPart<sml.CTPivotCacheRecords> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.SPREADSHEETML_PIVOT_CACHE_RECORDS, Namespaces.SPREADSHEETML_PIVOT_CACHE_RECORDS, { namespaceURI: S, localPart: 'pivotCacheRecords' });
  }
}

export class QueryTablePart extends XmlPart<sml.CTQueryTable> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.SPREADSHEETML_QUERY_TABLE, Namespaces.SPREADSHEETML_QUERY_TABLE, { namespaceURI: S, localPart: 'queryTable' });
  }
}

export class ConnectionsPart extends XmlPart<sml.CTConnections> {
  constructor(partName: PartName | string = '/xl/connections.xml') {
    super(partName, ContentTypes.SPREADSHEETML_CONNECTIONS, Namespaces.SPREADSHEETML_CONNECTIONS, { namespaceURI: S, localPart: 'connections' });
  }
}

export class ExternalLinkPart extends XmlPart<sml.CTExternalLink> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.SPREADSHEETML_EXTERNAL_LINK, Namespaces.SPREADSHEETML_EXTERNAL_LINK, { namespaceURI: S, localPart: 'externalLink' });
  }
}

export class PrinterSettingsPart extends BinaryPart {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.SPREADSHEETML_PRINTER_SETTINGS, Namespaces.SPREADSHEETML_PRINTER_SETTINGS);
  }
}

// The Excel 2010 and 2013 extension parts (CR-004 phase A; docx4j
// org.docx4j.openpackaging.parts.SpreadsheetML).  The seven XML ones are `DefaultXmlPart`s for
// now: their DOM round-trips byte for byte, and nothing re-marshals them, so an `mc:Ignorable`
// they carry is untouched.  Phase B makes them `XmlPart`s over the `x14` and `x15` roots.

/** A slicer cache (`x14:slicerCacheDefinition`), from the workbook. docx4j `SlicerCachePart`. */
export class SlicerCachePart extends DefaultXmlPart {
  constructor(partName: PartName | string = '/xl/slicerCaches/slicerCache1.xml') {
    super(partName, ContentTypes.SPREADSHEETML_SLICER_CACHE, Namespaces.SPREADSHEETML_SLICER_CACHE);
  }
}

/** A sheet's slicers (`x14:slicers`). docx4j `SlicersPart`. */
export class SlicersPart extends DefaultXmlPart {
  constructor(partName: PartName | string = '/xl/slicers/slicer1.xml') {
    super(partName, ContentTypes.SPREADSHEETML_SLICERS, Namespaces.SPREADSHEETML_SLICERS);
  }
}

/** A timeline cache (`x15:timelineCacheDefinition`), from the workbook. docx4j `TimelineCachePart`. */
export class TimelineCachePart extends DefaultXmlPart {
  constructor(partName: PartName | string = '/xl/timelineCaches/timelineCache1.xml') {
    super(partName, ContentTypes.SPREADSHEETML_TIMELINE_CACHE, Namespaces.SPREADSHEETML_TIMELINE_CACHE);
  }
}

/** A sheet's timelines (`x15:timelines`). docx4j `TimelinesPart`. */
export class TimelinesPart extends DefaultXmlPart {
  constructor(partName: PartName | string = '/xl/timelines/timeline1.xml') {
    super(partName, ContentTypes.SPREADSHEETML_TIMELINES, Namespaces.SPREADSHEETML_TIMELINES);
  }
}

/** A form control's properties (`x14:formControlPr`), from a worksheet. docx4j `ControlPropertiesPart`. */
export class ControlPropertiesPart extends DefaultXmlPart {
  constructor(partName: PartName | string = '/xl/ctrlProps/ctrlProp1.xml') {
    super(partName, ContentTypes.SPREADSHEETML_CONTROL_PROPERTIES, Namespaces.SPREADSHEETML_CONTROL_PROPERTIES);
  }
}

/** The properties of a custom data part (`x14:datastoreItem`). docx4j `CustomDataPropertiesPart`. */
export class CustomDataPropertiesPart extends DefaultXmlPart {
  constructor(partName: PartName | string = '/xl/customDataProps/customDataProps1.xml') {
    super(partName, ContentTypes.SPREADSHEETML_CUSTOM_DATA_PROPERTIES, Namespaces.SPREADSHEETML_CUSTOM_DATA_PROPERTIES);
  }
}

/** A survey (`x15:survey`), from the workbook. docx4j `SurveyPart`. */
export class SurveyPart extends DefaultXmlPart {
  constructor(partName: PartName | string = '/xl/surveys/survey1.xml') {
    super(partName, ContentTypes.SPREADSHEETML_SURVEY, Namespaces.SPREADSHEETML_SURVEY);
  }
}

/**
 * Custom data, whatever an add-in put there (docx4j `CustomDataPart`).  Its content type is the
 * generic `application/binary`, so the registry chooses it by relationship type, as it does for
 * altChunk, embedded packages, OLE objects and custom XML (CR-001 section 5.4).
 */
export class CustomDataPart extends BinaryPart {
  constructor(partName: PartName | string = '/xl/customData/customData1.dat', contentType: string = ContentTypes.SPREADSHEETML_CUSTOM_DATA) {
    super(partName, contentType, Namespaces.SPREADSHEETML_CUSTOM_DATA);
  }
}

/** The Power Pivot data model (`/xl/model/item.data`, [MS-XLDM]). docx4j `DataModelPart`. */
export class DataModelPart extends BinaryPart {
  constructor(partName: PartName | string = '/xl/model/item.data') {
    super(partName, ContentTypes.SPREADSHEETML_DATA_MODEL, Namespaces.SPREADSHEETML_DATA_MODEL);
  }
}
