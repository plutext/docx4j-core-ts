// PresentationML parts (docx4j org.docx4j.openpackaging.parts.PresentationML).
import type * as pml from '@docx4j/generated-objects-ts/modules/org_pptx4j_pml';
import type * as dml from '@docx4j/generated-objects-ts/modules/org_docx4j_dml';
import {
  createPresentationSldIdLstSldId, createPresentationSldMasterIdLstSldMasterId,
  createSlideLayoutIdListSldLayoutId,
} from '@docx4j/generated-objects-ts/factory/org_pptx4j_pml';
import { XmlPart } from '../XmlPart.mjs';
import { BinaryPart } from '../BinaryPart.mjs';
import { Part } from '../Part.mjs';
import { PartName } from '../../opc/PartName.mjs';
import { ContentTypes } from '../../opc/ContentTypes.mjs';
import { Namespaces } from '../Namespaces.mjs';
import { Pptx4jException } from '../../opc/exceptions.mjs';
import type { AddPartBehaviour, RelationshipsPart } from '../RelationshipsPart.mjs';
import { ThemePart } from '../dml/index.mjs';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** The parts an id list's `r:id`s point at, in list order, skipping any that is not of the class. */
function partsByRelId<T extends Part>(entries: readonly { rid: string }[], rp: RelationshipsPart, cls: new (...args: never[]) => T): T[] {
  const out: T[] = [];
  for (const entry of entries) {
    const part = rp.getPart(entry.rid);
    if (part instanceof cls) out.push(part);
  }
  return out;
}

/**
 * A new `p:sldId/@id` (docx4j `JaxbPmlPart.getSlideId()`, ECMA-376 4.8.17 ST_SlideId: 256 to
 * 2147483647). Random, as docx4j's is: nothing in the package numbers them.
 */
export function nextSlideId(): number {
  return Math.floor(Math.random() * 2147483392) + 256;
}

/**
 * A new `p:sldLayoutId/@id` or `p:sldMasterId/@id` (docx4j
 * `JaxbPmlPart.getSlideLayoutOrMasterId()`, ECMA-376 4.8.18 and 4.8.20: above 2147483648).
 */
export function nextSlideLayoutOrMasterId(): number {
  return Math.floor(Math.random() * 2147483647) + 2147483648;
}

export class MainPresentationPart extends XmlPart<pml.Presentation> {
  themePart: ThemePart | undefined;
  notesMasterPart: NotesMasterPart | undefined;
  commentAuthorsPart: CommentAuthorsPart | undefined;
  presentationPropertiesPart: PresentationPropertiesPart | undefined;
  viewPropertiesPart: ViewPropertiesPart | undefined;
  tableStylesPart: TableStylesPart | undefined;
  constructor(partName: PartName | string = '/ppt/presentation.xml', contentType: string = ContentTypes.PRESENTATIONML_MAIN) {
    super(partName, contentType, Namespaces.PRESENTATIONML_MAIN, { namespaceURI: P, localPart: 'presentation' });
  }

  override setPartShortcut(part: Part, relationshipType: string): boolean {
    switch (relationshipType) {
      case Namespaces.THEME: this.themePart = part as ThemePart; return true;
      case Namespaces.PRESENTATIONML_NOTES_MASTER: this.notesMasterPart = part as NotesMasterPart; return true;
      case Namespaces.PRESENTATIONML_COMMENT_AUTHORS: this.commentAuthorsPart = part as CommentAuthorsPart; return true;
      case Namespaces.PRESENTATIONML_PRES_PROPS: this.presentationPropertiesPart = part as PresentationPropertiesPart; return true;
      case Namespaces.PRESENTATIONML_VIEW_PROPS: this.viewPropertiesPart = part as ViewPropertiesPart; return true;
      case Namespaces.PRESENTATIONML_TABLE_STYLES: this.tableStylesPart = part as TableStylesPart; return true;
      default: return false;
    }
  }

  /**
   * The slides, in `p:sldIdLst` order (docx4j `getSlideParts()`), once the part is unmarshalled;
   * before that, in relationship order, which is the order they were added in.
   */
  get slideParts(): SlidePart[] {
    const rp = this.relationshipsPart;
    if (!rp) return [];
    if (this.isUnmarshalled) {
      const entries = this.contents.sldIdLst?.sldId;
      if (entries) return partsByRelId(entries, rp, SlidePart);
    }
    return rp.getRelationshipsByType(Namespaces.PRESENTATIONML_SLIDE).map((r) => rp.getPart(r)).filter((p): p is SlidePart => p instanceof SlidePart);
  }

  /** The slides in `p:sldIdLst` order, unmarshalling the part first (docx4j `getSlideParts()`). */
  async getSlideParts(): Promise<SlidePart[]> {
    await this.getContents();
    return this.slideParts;
  }

  /** The slide masters, in `p:sldMasterIdLst` order once the part is unmarshalled. */
  get slideMasterParts(): SlideMasterPart[] {
    const rp = this.relationshipsPart;
    if (!rp) return [];
    if (this.isUnmarshalled) {
      const entries = this.contents.sldMasterIdLst?.sldMasterId;
      if (entries) return partsByRelId(entries, rp, SlideMasterPart);
    }
    return rp.getRelationshipsByType(Namespaces.PRESENTATIONML_SLIDE_MASTER).map((r) => rp.getPart(r)).filter((p): p is SlideMasterPart => p instanceof SlideMasterPart);
  }

  /** The slide at an index of `p:sldIdLst` (docx4j `getSlide(int)`). */
  getSlide(index: number): SlidePart {
    const slides = this.slideParts;
    const slide = slides[index];
    if (!slide) throw new Pptx4jException(`No slide at index ${index}. (There are ${slides.length} slides)`);
    return slide;
  }

  /** docx4j `getSlideCount()`. */
  get slideCount(): number {
    return this.contents.sldIdLst?.sldId?.length ?? 0;
  }

  /**
   * Adds a slide part as a target and appends a `p:sldId` for it (docx4j
   * `addSlideIdListEntry(SlidePart)`); the part must be unmarshalled.
   */
  addSlideIdListEntry(slidePart: SlidePart, mode: AddPartBehaviour = 'OVERWRITE_IF_NAME_EXISTS'): pml.Presentation.SldIdLst.SldId {
    const rel = this.addTargetPart(slidePart, mode);
    const entry = createPresentationSldIdLstSldId({ id: nextSlideId(), rid: rel.id });
    this.sldIdLst().sldId!.push(entry);
    return entry;
  }

  /**
   * docx4j `addSlide(SlidePart)` / `addSlide(int, SlidePart)`: appends the slide, or inserts it
   * at `index`, renaming the part if the name is taken (PowerPoint cannot open a package in which
   * two relationships target the same slide, so a slide is never added twice).
   *
   * `layoutPart`, when given, becomes the slide's layout relationship, as a slide needs one.
   */
  addSlide(slidePart: SlidePart, index?: number, layoutPart?: SlideLayoutPart): pml.Presentation.SldIdLst.SldId {
    const sldId = this.sldIdLst().sldId!;
    if (index !== undefined && (index < 0 || index > sldId.length)) {
      throw new Pptx4jException(`Can't add slide at index ${index}. (There are ${sldId.length} slides) `);
    }
    const rel = this.addTargetPart(slidePart, 'RENAME_IF_NAME_EXISTS');
    const entry = createPresentationSldIdLstSldId({ id: nextSlideId(), rid: rel.id });
    if (index === undefined) sldId.push(entry); else sldId.splice(index, 0, entry);
    if (layoutPart) slidePart.addTargetPart(layoutPart);
    return entry;
  }

  /** docx4j `addSlideMasterIdListEntry(SlideMasterPart)`. */
  addSlideMasterIdListEntry(masterPart: SlideMasterPart): pml.Presentation.SldMasterIdLst.SldMasterId {
    const rel = this.addTargetPart(masterPart);
    const entry = createPresentationSldMasterIdLstSldMasterId({ id: nextSlideLayoutOrMasterId(), rid: rel.id });
    const lst = (this.contents.sldMasterIdLst ??= { sldMasterId: [] });
    (lst.sldMasterId ??= []).push(entry);
    return entry;
  }

  /** docx4j `removeSlide(int)`: the `p:sldId`, the relationship and the part. */
  removeSlide(index: number): SlidePart {
    const sldId = this.sldIdLst().sldId!;
    const entry = sldId[index];
    if (!entry) throw new Pptx4jException(`No slide at index ${index}. (There are ${sldId.length} slides) `);
    sldId.splice(index, 1);
    const rp = this.relationshipsPart!;
    const part = rp.getPart(entry.rid) as SlidePart;
    rp.removePart(part.partName);
    return part;
  }

  private sldIdLst(): pml.Presentation.SldIdLst {
    const contents = this.contents;
    const lst = (contents.sldIdLst ??= { sldId: [] });
    lst.sldId ??= [];
    return lst;
  }
}

export class SlidePart extends XmlPart<pml.Sld> {
  slideLayoutPart: SlideLayoutPart | undefined;
  notesSlidePart: NotesSlidePart | undefined;
  commentsPart: PresentationCommentsPart | undefined;
  constructor(partName: PartName | string = '/ppt/slides/slide1.xml') {
    super(partName, ContentTypes.PRESENTATIONML_SLIDE, Namespaces.PRESENTATIONML_SLIDE, { namespaceURI: P, localPart: 'sld' });
  }
  override setPartShortcut(part: Part, relationshipType: string): boolean {
    switch (relationshipType) {
      case Namespaces.PRESENTATIONML_SLIDE_LAYOUT: this.slideLayoutPart = part as SlideLayoutPart; return true;
      case Namespaces.PRESENTATIONML_NOTES_SLIDE: this.notesSlidePart = part as NotesSlidePart; return true;
      case Namespaces.PRESENTATIONML_COMMENTS: this.commentsPart = part as PresentationCommentsPart; return true;
      default: return false;
    }
  }
}

export class SlideLayoutPart extends XmlPart<pml.SldLayout> {
  /** docx4j `SlideLayoutPart.getSlideMasterPart()`. */
  slideMasterPart: SlideMasterPart | undefined;
  constructor(partName: PartName | string = '/ppt/slideLayouts/slideLayout1.xml') {
    super(partName, ContentTypes.PRESENTATIONML_SLIDE_LAYOUT, Namespaces.PRESENTATIONML_SLIDE_LAYOUT, { namespaceURI: P, localPart: 'sldLayout' });
  }
  override setPartShortcut(part: Part, relationshipType: string): boolean {
    if (relationshipType === Namespaces.PRESENTATIONML_SLIDE_MASTER) { this.slideMasterPart = part as SlideMasterPart; return true; }
    return false;
  }
}

export class SlideMasterPart extends XmlPart<pml.SldMaster> {
  themePart: ThemePart | undefined;
  constructor(partName: PartName | string = '/ppt/slideMasters/slideMaster1.xml') {
    super(partName, ContentTypes.PRESENTATIONML_SLIDE_MASTER, Namespaces.PRESENTATIONML_SLIDE_MASTER, { namespaceURI: P, localPart: 'sldMaster' });
  }
  override setPartShortcut(part: Part, relationshipType: string): boolean {
    if (relationshipType === Namespaces.THEME) { this.themePart = part as ThemePart; return true; }
    return false;
  }

  /**
   * Adds a layout part as a target and appends a `p:sldLayoutId` for it (docx4j
   * `addSlideLayoutIdListEntry`); the master must be unmarshalled.
   */
  addSlideLayoutIdListEntry(slideLayoutPart: SlideLayoutPart): pml.SlideLayoutIdList.SldLayoutId {
    const rel = this.addTargetPart(slideLayoutPart);
    const entry = createSlideLayoutIdListSldLayoutId({ id: nextSlideLayoutOrMasterId(), rid: rel.id });
    const lst = (this.contents.sldLayoutIdLst ??= { sldLayoutId: [] });
    (lst.sldLayoutId ??= []).push(entry);
    return entry;
  }

  /** This master's layouts, in `p:sldLayoutIdLst` order once the master is unmarshalled. */
  get slideLayoutParts(): SlideLayoutPart[] {
    const rp = this.relationshipsPart;
    if (!rp) return [];
    if (this.isUnmarshalled) {
      const entries = this.contents.sldLayoutIdLst?.sldLayoutId;
      if (entries) return partsByRelId(entries, rp, SlideLayoutPart);
    }
    return rp.getRelationshipsByType(Namespaces.PRESENTATIONML_SLIDE_LAYOUT).map((r) => rp.getPart(r)).filter((p): p is SlideLayoutPart => p instanceof SlideLayoutPart);
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
