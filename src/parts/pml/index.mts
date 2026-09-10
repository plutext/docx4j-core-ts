// PresentationML parts (docx4j org.docx4j.openpackaging.parts.PresentationML).
import type * as pml from '@docx4j/generated-objects-ts/modules/org_pptx4j_pml';
import type * as dml from '@docx4j/generated-objects-ts/modules/org_docx4j_dml';
import { XmlPart } from '../XmlPart.mjs';
import { BinaryPart } from '../BinaryPart.mjs';
import { Part } from '../Part.mjs';
import { PartName } from '../../opc/PartName.mjs';
import { ContentTypes } from '../../opc/ContentTypes.mjs';
import { Namespaces } from '../Namespaces.mjs';
import { ThemePart } from '../dml/index.mjs';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

export class MainPresentationPart extends XmlPart<pml.Presentation> {
  themePart: ThemePart | undefined;
  constructor(partName: PartName | string = '/ppt/presentation.xml', contentType: string = ContentTypes.PRESENTATIONML_MAIN) {
    super(partName, contentType, Namespaces.PRESENTATIONML_MAIN, { namespaceURI: P, localPart: 'presentation' });
  }
  override setPartShortcut(part: Part, relationshipType: string): boolean {
    if (relationshipType === Namespaces.THEME) { this.themePart = part as ThemePart; return true; }
    return false;
  }
  get slideParts(): SlidePart[] {
    const rp = this.relationshipsPart;
    if (!rp) return [];
    return rp.getRelationshipsByType(Namespaces.PRESENTATIONML_SLIDE).map((r) => rp.getPart(r)).filter((p): p is SlidePart => p instanceof SlidePart);
  }
}

export class SlidePart extends XmlPart<pml.Sld> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.PRESENTATIONML_SLIDE, Namespaces.PRESENTATIONML_SLIDE, { namespaceURI: P, localPart: 'sld' });
  }
}

export class SlideLayoutPart extends XmlPart<pml.SldLayout> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.PRESENTATIONML_SLIDE_LAYOUT, Namespaces.PRESENTATIONML_SLIDE_LAYOUT, { namespaceURI: P, localPart: 'sldLayout' });
  }
}

export class SlideMasterPart extends XmlPart<pml.SldMaster> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.PRESENTATIONML_SLIDE_MASTER, Namespaces.PRESENTATIONML_SLIDE_MASTER, { namespaceURI: P, localPart: 'sldMaster' });
  }
}

export class NotesSlidePart extends XmlPart<pml.Notes> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.PRESENTATIONML_NOTES_SLIDE, Namespaces.PRESENTATIONML_NOTES_SLIDE, { namespaceURI: P, localPart: 'notes' });
  }
}

export class NotesMasterPart extends XmlPart<pml.NotesMaster> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.PRESENTATIONML_NOTES_MASTER, Namespaces.PRESENTATIONML_NOTES_MASTER, { namespaceURI: P, localPart: 'notesMaster' });
  }
}

export class HandoutMasterPart extends XmlPart<pml.HandoutMaster> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.PRESENTATIONML_HANDOUT_MASTER, Namespaces.PRESENTATIONML_HANDOUT_MASTER, { namespaceURI: P, localPart: 'handoutMaster' });
  }
}

export class PresentationPropertiesPart extends XmlPart<pml.PresentationPr> {
  constructor(partName: PartName | string = '/ppt/presProps.xml') {
    super(partName, ContentTypes.PRESENTATIONML_PRES_PROPS, Namespaces.PRESENTATIONML_PRES_PROPS, { namespaceURI: P, localPart: 'presentationPr' });
  }
}

export class ViewPropertiesPart extends XmlPart<pml.ViewPr> {
  constructor(partName: PartName | string = '/ppt/viewProps.xml') {
    super(partName, ContentTypes.PRESENTATIONML_VIEW_PROPS, Namespaces.PRESENTATIONML_VIEW_PROPS, { namespaceURI: P, localPart: 'viewPr' });
  }
}

export class TableStylesPart extends XmlPart<dml.CTTableStyleList> {
  constructor(partName: PartName | string = '/ppt/tableStyles.xml') {
    super(partName, ContentTypes.PRESENTATIONML_TABLE_STYLES, Namespaces.PRESENTATIONML_TABLE_STYLES, { namespaceURI: A, localPart: 'tblStyleLst' });
  }
}

/** docx4j PresentationML.CommentsPart, renamed here to avoid the clash with the WordprocessingML one. */
export class PresentationCommentsPart extends XmlPart<pml.CTCommentList> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.PRESENTATIONML_COMMENTS, Namespaces.PRESENTATIONML_COMMENTS, { namespaceURI: P, localPart: 'cmLst' });
  }
}

export class CommentAuthorsPart extends XmlPart<pml.CTCommentAuthorList> {
  constructor(partName: PartName | string = '/ppt/commentAuthors.xml') {
    super(partName, ContentTypes.PRESENTATIONML_COMMENT_AUTHORS, Namespaces.PRESENTATIONML_COMMENT_AUTHORS, { namespaceURI: P, localPart: 'cmAuthorLst' });
  }
}

export class TagsPart extends XmlPart<pml.TagLst> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.PRESENTATIONML_TAGS, Namespaces.PRESENTATIONML_TAGS, { namespaceURI: P, localPart: 'tagLst' });
  }
}

export class FontDataPart extends BinaryPart {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.PRESENTATIONML_FONT_DATA, Namespaces.PRESENTATIONML_FONT_DATA);
  }
}
