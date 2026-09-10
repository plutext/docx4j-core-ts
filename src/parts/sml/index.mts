// SpreadsheetML parts (docx4j org.docx4j.openpackaging.parts.SpreadsheetML).
import type * as sml from '@docx4j/generated-objects-ts/modules/org_xlsx4j_sml';
import { XmlPart } from '../XmlPart.mjs';
import { BinaryPart } from '../BinaryPart.mjs';
import { Part } from '../Part.mjs';
import { PartName } from '../../opc/PartName.mjs';
import { ContentTypes } from '../../opc/ContentTypes.mjs';
import { Namespaces } from '../Namespaces.mjs';
import { ThemePart } from '../dml/index.mjs';

const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export class WorkbookPart extends XmlPart<sml.Workbook> {
  sharedStringsPart: SharedStringsPart | undefined;
  stylesPart: StylesPart | undefined;
  themePart: ThemePart | undefined;
  calcChainPart: CalcChainPart | undefined;
  constructor(partName: PartName | string = '/xl/workbook.xml', contentType: string = ContentTypes.SPREADSHEETML_WORKBOOK) {
    super(partName, contentType, Namespaces.SPREADSHEETML_WORKBOOK, { namespaceURI: S, localPart: 'workbook' });
  }
  override setPartShortcut(part: Part, relationshipType: string): boolean {
    switch (relationshipType) {
      case Namespaces.SPREADSHEETML_SHARED_STRINGS: this.sharedStringsPart = part as SharedStringsPart; return true;
      case Namespaces.SPREADSHEETML_STYLES: this.stylesPart = part as StylesPart; return true;
      case Namespaces.THEME: this.themePart = part as ThemePart; return true;
      case Namespaces.SPREADSHEETML_CALC_CHAIN: this.calcChainPart = part as CalcChainPart; return true;
      default: return false;
    }
  }
  /** Worksheet parts in relationship order (the sheet order is in `contents.sheets`). */
  get worksheetParts(): WorksheetPart[] {
    const rp = this.relationshipsPart;
    if (!rp) return [];
    return rp.getRelationshipsByType(Namespaces.SPREADSHEETML_WORKSHEET).map((r) => rp.getPart(r)).filter((p): p is WorksheetPart => p instanceof WorksheetPart);
  }
}

export class WorksheetPart extends XmlPart<sml.Worksheet> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.SPREADSHEETML_WORKSHEET, Namespaces.SPREADSHEETML_WORKSHEET, { namespaceURI: S, localPart: 'worksheet' });
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
