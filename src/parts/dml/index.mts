// DrawingML parts (docx4j org.docx4j.openpackaging.parts.DrawingML and ThemePart).
import type * as dml from '@docx4j/generated-objects-ts/modules/org_docx4j_dml';
import type * as chart from '@docx4j/generated-objects-ts/modules/org_docx4j_dml_chart';
import type * as diagram from '@docx4j/generated-objects-ts/modules/org_docx4j_dml_diagram';
import type * as chartDrawing from '@docx4j/generated-objects-ts/modules/org_docx4j_dml_chartDrawing';
import type * as ssDrawing from '@docx4j/generated-objects-ts/modules/org_docx4j_dml_spreadsheetdrawing';
import { XmlPart } from '../XmlPart.mjs';
import { DefaultXmlPart } from '../DefaultXmlPart.mjs';
import { PartName } from '../../opc/PartName.mjs';
import { ContentTypes } from '../../opc/ContentTypes.mjs';
import { Namespaces } from '../Namespaces.mjs';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const CDR = 'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

export class ThemePart extends XmlPart<dml.Theme> {
  constructor(partName: PartName | string = '/word/theme/theme1.xml') {
    super(partName, ContentTypes.OFFICEDOCUMENT_THEME, Namespaces.THEME, { namespaceURI: A, localPart: 'theme' });
  }
}

export class ThemeOverridePart extends XmlPart<dml.CTBaseStylesOverride> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.OFFICEDOCUMENT_THEME_OVERRIDE, Namespaces.THEME_OVERRIDE, { namespaceURI: A, localPart: 'themeOverride' });
  }
}

/** A chart (docx4j org.docx4j.openpackaging.parts.DrawingML.Chart). */
export class ChartPart extends XmlPart<chart.CTChartSpace> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.DRAWINGML_CHART, Namespaces.SPREADSHEETML_CHART, { namespaceURI: C, localPart: 'chartSpace' });
  }
}

/** The user shapes of a chart (docx4j ChartShapePart). */
export class ChartShapePart extends XmlPart<chartDrawing.CTDrawing> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.DRAWINGML_CHART_SHAPES, Namespaces.CHART_USER_SHAPES, { namespaceURI: CDR, localPart: 'userShapes' });
  }
}

/** A spreadsheet drawing (docx4j org.docx4j.openpackaging.parts.DrawingML.Drawing). */
export class DrawingPart extends XmlPart<ssDrawing.CTDrawing> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.DRAWINGML_DRAWING, Namespaces.SPREADSHEETML_DRAWING, { namespaceURI: XDR, localPart: 'wsDr' });
  }
}

export class DiagramDataPart extends XmlPart<diagram.CTDataModel> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.DRAWINGML_DIAGRAM_DATA, Namespaces.DRAWINGML_DIAGRAM_DATA, { namespaceURI: DGM, localPart: 'dataModel' });
  }
}

export class DiagramLayoutPart extends XmlPart<diagram.CTDiagramDefinition> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.DRAWINGML_DIAGRAM_LAYOUT, Namespaces.DRAWINGML_DIAGRAM_LAYOUT, { namespaceURI: DGM, localPart: 'layoutDef' });
  }
}

export class DiagramLayoutHeaderPart extends XmlPart<diagram.CTDiagramDefinitionHeader> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.DRAWINGML_DIAGRAM_LAYOUT_HEADER, Namespaces.DRAWINGML_DIAGRAM_LAYOUT_HEADER, { namespaceURI: DGM, localPart: 'layoutDefHdr' });
  }
}

export class DiagramStylePart extends XmlPart<diagram.CTStyleDefinition> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.DRAWINGML_DIAGRAM_STYLE, Namespaces.DRAWINGML_DIAGRAM_STYLE, { namespaceURI: DGM, localPart: 'styleDef' });
  }
}

export class DiagramColorsPart extends XmlPart<diagram.CTColorTransform> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.DRAWINGML_DIAGRAM_COLORS, Namespaces.DRAWINGML_DIAGRAM_COLORS, { namespaceURI: DGM, localPart: 'colorsDef' });
  }
}

/** Chart style, colour style, chartEx and diagram drawing parts: DOM until typed (later CR). */
export class ChartStylePart extends DefaultXmlPart {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.CHART_STYLE, Namespaces.CHART_STYLE);
  }
}

export class ChartColorStylePart extends DefaultXmlPart {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.CHART_COLOR_STYLE, Namespaces.CHART_COLOR_STYLE);
  }
}

export class ChartExSpacePart extends DefaultXmlPart {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.CHART_EX, Namespaces.CHART_EX);
  }
}

export class DiagramDrawingPart extends DefaultXmlPart {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.DRAWINGML_DIAGRAM_DRAWING, Namespaces.DRAWINGML_DIAGRAM_DRAWING);
  }
}
