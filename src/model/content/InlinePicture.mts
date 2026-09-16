// CR-002 phase C: inline pictures. The counterpart of docx4j's
// BinaryPartAbstractImage.createImageInline (and of its ImageInfo, whose job the header readers
// below do): an ImagePart, a relationship from the part the body belongs to, and a
// w:drawing/wp:inline sized from the image's pixel dimensions and resolution.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as dml from '@docx4j/generated-objects-ts/modules/org_docx4j_dml';
import type * as wp from '@docx4j/generated-objects-ts/modules/org_docx4j_dml_wordprocessingDrawing';
import type * as pic from '@docx4j/generated-objects-ts/modules/org_docx4j_dml_picture';
import * as elWml from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import * as elPic from '@docx4j/generated-objects-ts/el/org_docx4j_dml_picture';
import { createInline, createCTEffectExtent } from '@docx4j/generated-objects-ts/factory/org_docx4j_dml_wordprocessingDrawing';
import { createPic, createCTPictureNonVisual } from '@docx4j/generated-objects-ts/factory/org_docx4j_dml_picture';
import {
  createCTPositiveSize2D, createCTNonVisualDrawingProps, createCTNonVisualGraphicFrameProperties, createCTGraphicalObjectFrameLocking,
  createGraphic, createGraphicData, createCTBlip, createCTBlipFillProperties, createCTStretchInfoProperties, createCTRelativeRect,
  createCTShapeProperties, createCTTransform2D, createCTPoint2D, createCTPresetGeometry2D, createCTGeomGuideList, createCTNonVisualPictureProperties,
} from '@docx4j/generated-objects-ts/factory/org_docx4j_dml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { ContentTypes, IMAGE_CONTENT_TYPES_BY_EXTENSION } from '../../opc/ContentTypes.mjs';
import { base64Decode, base64Encode } from '../../xml/dom.mjs';
import { ImagePart } from '../../parts/BinaryPart.mjs';
import { PartName } from '../../opc/PartName.mjs';
import type { Part } from '../../parts/Part.mjs';
import { type Element, find, runOf } from './tree.mjs';
import type { Paragraph } from './Paragraph.mjs';

/** English Metric Units per inch, and per twip (docx4j UnitsOfMeasurement). */
const EMU_PER_INCH = 914400;
const EMU_PER_TWIP = 635;
/** docx4j's ImageInfo default when a format carries no resolution. */
const DEFAULT_DPI = 96;
const EMU_PER_POINT = 12700;

const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

/** Office JS `Word.ImageFormat`, the subset Word writes into a docx. */
export type ImageFormat = 'Unsupported' | 'Bmp' | 'Emf' | 'Gif' | 'Jpeg' | 'Png' | 'Svg' | 'Tiff' | 'Wmf';

/** What a header reader answers: pixels and the resolution the file declares. */
export interface ImageInfo {
  contentType: string;
  /** The file extension a part of this kind is given ('png'). */
  extension: string;
  widthPx: number;
  heightPx: number;
  /** Horizontal and vertical resolution; 96 when the format carries none (docx4j's default). */
  dpiX: number;
  dpiY: number;
}

/**
 * The size and kind of an image, from its header: PNG (IHDR and pHYs), JPEG (SOFn and the JFIF
 * APP0 densities), GIF (the logical screen descriptor) and BMP (the DIB header, with its pixels
 * per metre). docx4j reads these through XML Graphics Commons' ImageInfo; this is the same
 * information, and the same 96 dpi default when a file declares none.
 */
export function imageInfoOf(bytes: Uint8Array): ImageInfo {
  const png = pngInfo(bytes);
  if (png) return png;
  const jpeg = jpegInfo(bytes);
  if (jpeg) return jpeg;
  const gif = gifInfo(bytes);
  if (gif) return gif;
  const bmp = bmpInfo(bytes);
  if (bmp) return bmp;
  const head = Array.from(bytes.subarray(0, 8)).map((b) => b.toString(16).padStart(2, '0')).join(' ');
  throw new Docx4JException(`Unsupported image format (first bytes ${head}); PNG, JPEG, GIF and BMP are read here`);
}

function be16(b: Uint8Array, at: number): number { return (b[at]! << 8) | b[at + 1]!; }
function be32(b: Uint8Array, at: number): number { return ((b[at]! << 24) >>> 0) + (b[at + 1]! << 16) + (b[at + 2]! << 8) + b[at + 3]!; }
function le16(b: Uint8Array, at: number): number { return b[at]! | (b[at + 1]! << 8); }
function le32(b: Uint8Array, at: number): number { return (b[at]! | (b[at + 1]! << 8) | (b[at + 2]! << 16) | (b[at + 3]! << 24)) >>> 0; }
function sle32(b: Uint8Array, at: number): number { return b[at]! | (b[at + 1]! << 8) | (b[at + 2]! << 16) | (b[at + 3]! << 24); }

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

function pngInfo(b: Uint8Array): ImageInfo | undefined {
  if (b.length < 24 || PNG_SIGNATURE.some((v, i) => b[i] !== v)) return undefined;
  const info: ImageInfo = {
    contentType: ContentTypes.IMAGE_PNG, extension: 'png',
    widthPx: be32(b, 16), heightPx: be32(b, 20), dpiX: DEFAULT_DPI, dpiY: DEFAULT_DPI,
  };
  // pHYs: pixels per unit, unit 1 = metre.
  for (let at = 8; at + 8 <= b.length;) {
    const length = be32(b, at);
    const type = String.fromCharCode(b[at + 4]!, b[at + 5]!, b[at + 6]!, b[at + 7]!);
    if (type === 'pHYs' && at + 8 + 9 <= b.length) {
      if (b[at + 8 + 8] === 1) {
        info.dpiX = Math.round(be32(b, at + 8) * 0.0254) || DEFAULT_DPI;
        info.dpiY = Math.round(be32(b, at + 12) * 0.0254) || DEFAULT_DPI;
      }
      break;
    }
    if (type === 'IDAT' || type === 'IEND') break;
    at += 12 + length;
  }
  return info;
}

function jpegInfo(b: Uint8Array): ImageInfo | undefined {
  if (b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) return undefined;
  const info: ImageInfo = {
    contentType: ContentTypes.IMAGE_JPEG, extension: 'jpeg',
    widthPx: 0, heightPx: 0, dpiX: DEFAULT_DPI, dpiY: DEFAULT_DPI,
  };
  for (let at = 2; at + 4 <= b.length;) {
    if (b[at] !== 0xff) { at++; continue; }
    const marker = b[at + 1]!;
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { at += 2; continue; }
    if (marker === 0xd9 || marker === 0xda) break;
    const length = be16(b, at + 2);
    const data = at + 4;
    if (marker === 0xe0 && b.length >= data + 12 && String.fromCharCode(b[data]!, b[data + 1]!, b[data + 2]!, b[data + 3]!) === 'JFIF') {
      const units = b[data + 7];
      const x = be16(b, data + 8);
      const y = be16(b, data + 10);
      if (units === 1 && x && y) { info.dpiX = x; info.dpiY = y; }
      else if (units === 2 && x && y) { info.dpiX = Math.round(x * 2.54); info.dpiY = Math.round(y * 2.54); }
    }
    // SOF0..SOF15, less the marker-only and non-frame ones
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      info.heightPx = be16(b, data + 1);
      info.widthPx = be16(b, data + 3);
      if (info.widthPx && info.heightPx) return info;
    }
    at += 2 + length;
  }
  return info.widthPx && info.heightPx ? info : undefined;
}

function gifInfo(b: Uint8Array): ImageInfo | undefined {
  if (b.length < 10 || b[0] !== 0x47 || b[1] !== 0x49 || b[2] !== 0x46) return undefined;
  return { contentType: ContentTypes.IMAGE_GIF, extension: 'gif', widthPx: le16(b, 6), heightPx: le16(b, 8), dpiX: DEFAULT_DPI, dpiY: DEFAULT_DPI };
}

function bmpInfo(b: Uint8Array): ImageInfo | undefined {
  if (b.length < 26 || b[0] !== 0x42 || b[1] !== 0x4d) return undefined;
  const headerSize = le32(b, 14);
  const info: ImageInfo = { contentType: ContentTypes.IMAGE_BMP, extension: 'bmp', widthPx: 0, heightPx: 0, dpiX: DEFAULT_DPI, dpiY: DEFAULT_DPI };
  if (headerSize === 12) {
    info.widthPx = le16(b, 18);
    info.heightPx = le16(b, 20);
    return info;
  }
  info.widthPx = sle32(b, 18);
  info.heightPx = Math.abs(sle32(b, 22));
  if (b.length >= 46) {
    const x = sle32(b, 38);
    const y = sle32(b, 42);
    if (x > 0) info.dpiX = Math.round(x * 0.0254) || DEFAULT_DPI;
    if (y > 0) info.dpiY = Math.round(y * 0.0254) || DEFAULT_DPI;
  }
  return info;
}

/** The natural size of an image in EMU: pixels over the resolution it declares. */
export function naturalSizeEmu(info: ImageInfo): { cx: number; cy: number } {
  return {
    cx: Math.max(1, Math.round((info.widthPx / (info.dpiX || DEFAULT_DPI)) * EMU_PER_INCH)),
    cy: Math.max(1, Math.round((info.heightPx / (info.dpiY || DEFAULT_DPI)) * EMU_PER_INCH)),
  };
}

/** Options for `insertInlinePicture`; Office JS has none, this is the extension surface. */
export interface InlinePictureOptions {
  /** wp:docPr/@name, the filename hint docx4j asks for. */
  name?: string;
  /** wp:docPr/@descr, Office JS altTextDescription. */
  altTextDescription?: string;
  /** The width in points; the height follows the aspect ratio unless it is given too. */
  width?: number;
  height?: number;
}

/** A w:drawing holding one wp:inline for an image part, as docx4j createImageInline builds it. */
export function drawingFor(relId: string, cx: number, cy: number, id: number, name: string, altText: string): Element<wml.Drawing> {
  const extent = (): dml.CTPositiveSize2D => createCTPositiveSize2D({ cx, cy });
  const picture: pic.Pic = createPic({
    nvPicPr: createCTPictureNonVisual({
      cNvPr: createCTNonVisualDrawingProps({ id, name, descr: altText }),
      cNvPicPr: createCTNonVisualPictureProperties({}),
    }),
    blipFill: createCTBlipFillProperties({
      blip: createCTBlip({ embed: relId }),
      stretch: createCTStretchInfoProperties({ fillRect: createCTRelativeRect({}) }),
    }),
    spPr: createCTShapeProperties({
      xfrm: createCTTransform2D({ off: createCTPoint2D({ x: 0, y: 0 }), ext: extent() }),
      prstGeom: createCTPresetGeometry2D({ prst: 'rect', avLst: createCTGeomGuideList({}) }),
    }),
  });
  const inline: wp.Inline = createInline({
    distT: 0, distB: 0, distL: 0, distR: 0,
    extent: extent(),
    effectExtent: createCTEffectExtent({ l: 0, t: 0, r: 0, b: 0 }),
    docPr: createCTNonVisualDrawingProps({ id, name, descr: altText }),
    cNvGraphicFramePr: createCTNonVisualGraphicFrameProperties({ graphicFrameLocks: createCTGraphicalObjectFrameLocking({ noChangeAspect: true }) }),
    graphic: createGraphic({ graphicData: createGraphicData({ uri: PIC_NS, any: [elPic.pic(picture)] }) }),
  });
  return elWml.drawing({ anchorOrInline: [inline] }) as Element<wml.Drawing>;
}

/** The run holding a new picture, plus the image part it added (the caller inserts the run). */
export interface NewPicture {
  run: Element<wml.R>;
  drawing: Element<wml.Drawing>;
  imagePart: ImagePart;
  relId: string;
}

/**
 * Adds the image as a part of `source` (the main document part, or the header or footer the body
 * belongs to) with a relationship, and builds the run that shows it. docx4j:
 * `BinaryPartAbstractImage.createImagePart` then `createImageInline`.
 */
export function addImage(source: Part, base64: string, scope: object, maxCx: number | undefined, options: InlinePictureOptions = {}): NewPicture {
  const pkg = source.package;
  if (!pkg) throw new Docx4JException('This part is not in a package yet; add it before adding an image');
  const bytes = base64Decode(base64);
  const info = imageInfoOf(bytes);
  const imagePart = new ImagePart(freeImageName(pkg, info.extension), info.contentType);
  imagePart.setBytes(bytes);
  const rel = source.addTargetPart(imagePart);
  let { cx, cy } = naturalSizeEmu(info);
  if (options.width !== undefined) {
    const wanted = Math.round(options.width * EMU_PER_POINT);
    cy = options.height !== undefined ? Math.round(options.height * EMU_PER_POINT) : Math.max(1, Math.round((cy * wanted) / cx));
    cx = wanted;
  } else if (options.height !== undefined) {
    const wanted = Math.round(options.height * EMU_PER_POINT);
    cx = Math.max(1, Math.round((cx * wanted) / cy));
    cy = wanted;
  } else if (maxCx !== undefined && cx > maxCx) {
    // docx4j CxCy.scale: an image wider than the text area is scaled down, keeping the ratio.
    cy = Math.max(1, Math.round((cy * maxCx) / cx));
    cx = maxCx;
  }
  const id = nextDrawingId(scope);
  const name = options.name ?? `Picture ${id}`;
  const drawing = drawingFor(rel.id, cx, cy, id, name, options.altTextDescription ?? '');
  return { run: runOf([drawing]), drawing, imagePart, relId: rel.id };
}

/** `/word/media/imageN.<ext>`, the first N free in the package (docx4j getNewPartName). */
function freeImageName(pkg: { getPart(name: PartName | string): unknown }, extension: string): PartName {
  for (let i = 1; ; i++) {
    const name = PartName.of(`/word/media/image${i}.${extension}`);
    if (!pkg.getPart(name)) return name;
  }
}

/** An id no other drawing in this part uses: wp:docPr and pic:cNvPr are the same type. */
function nextDrawingId(scope: object): number {
  let max = 0;
  for (const props of find<dml.CTNonVisualDrawingProps>(scope, 'org_docx4j_dml.CTNonVisualDrawingProps')) {
    if (typeof props.id === 'number' && props.id > max) max = props.id;
  }
  return max + 1;
}

/** The writable width of a section in EMU (page width less the margins), for scaling. */
export function writableWidthEmu(container: object): number | undefined {
  const sectPr = (container as { sectPr?: wml.SectPr }).sectPr;
  const pgSz = sectPr?.pgSz;
  if (!pgSz?.w) return undefined;
  const left = sectPr?.pgMar?.left ?? 0;
  const right = sectPr?.pgMar?.right ?? 0;
  const twips = pgSz.w - left - right;
  return twips > 0 ? twips * EMU_PER_TWIP : undefined;
}

const FORMAT_BY_CONTENT_TYPE: Readonly<Record<string, ImageFormat>> = {
  [ContentTypes.IMAGE_PNG]: 'Png', [ContentTypes.IMAGE_X_PNG]: 'Png', [ContentTypes.IMAGE_JPEG]: 'Jpeg',
  [ContentTypes.IMAGE_GIF]: 'Gif', [ContentTypes.IMAGE_BMP]: 'Bmp', [ContentTypes.IMAGE_TIFF]: 'Tiff',
  [ContentTypes.IMAGE_SVG]: 'Svg', [ContentTypes.IMAGE_EMF]: 'Emf', [ContentTypes.IMAGE_EMF2]: 'Emf',
  [ContentTypes.IMAGE_WMF]: 'Wmf',
};

/**
 * A subset of Office JS `Word.InlinePicture` over a `w:drawing` holding a `wp:inline`. A view:
 * the drawing, the run it is in and the paragraph. Sizes are points, as Office JS reports them.
 */
export class InlinePicture {
  constructor(
    /** The `w:drawing` element pair. */
    readonly element: Element<wml.Drawing>,
    /** The run holding the drawing, and the array holding the run. */
    readonly run: Element<wml.R>,
    readonly paragraph: Paragraph,
  ) {}

  /** The wp:inline; a floating picture (wp:anchor) is not an InlinePicture, and throws. */
  get inline(): wp.Inline {
    const first = this.element.value.anchorOrInline?.[0];
    if (!first || (first as { TYPE_NAME?: string }).TYPE_NAME !== 'org_docx4j_dml_wordprocessingDrawing.Inline') {
      throw new Docx4JException('This drawing is not an inline picture (wp:anchor is a floating shape)');
    }
    return first as wp.Inline;
  }

  /** Width in points, as Office JS. */
  get width(): number {
    return this.inline.extent.cx / EMU_PER_POINT;
  }
  set width(points: number) {
    this.resize(Math.round(points * EMU_PER_POINT), undefined);
  }

  get height(): number {
    return this.inline.extent.cy / EMU_PER_POINT;
  }
  set height(points: number) {
    this.resize(undefined, Math.round(points * EMU_PER_POINT));
  }

  /** wp:docPr/@descr (Word's alt text). */
  get altTextDescription(): string {
    return this.inline.docPr.descr ?? '';
  }
  set altTextDescription(text: string) {
    this.inline.docPr.descr = text;
    const p = this.pic();
    if (p) p.nvPicPr.cNvPr.descr = text;
  }

  /** wp:docPr/@title (typed since objects 0.1.3; before that the attribute was not marshalled, CR-002 section 8). */
  get altTextTitle(): string {
    return this.inline.docPr.title ?? '';
  }
  set altTextTitle(text: string) {
    this.inline.docPr.title = text;
  }

  /** The kind of image, from the content type of the part the picture embeds. */
  get imageFormat(): ImageFormat {
    const part = this.imagePart();
    return (part && FORMAT_BY_CONTENT_TYPE[part.contentType]) ?? 'Unsupported';
  }

  /** The relationship id of the embedded image (r:embed), if any (extension). */
  get relId(): string | undefined {
    return this.pic()?.blipFill.blip?.embed;
  }

  /** The image's bytes as base64 (Office JS getBase64ImageSrc). */
  async getBase64(): Promise<string> {
    const part = this.imagePart();
    if (!part) throw new Docx4JException('This picture has no image part (a linked or missing image)');
    return base64Encode(await part.getBytes());
  }

  /** Office JS's name for `getBase64()`. */
  getBase64ImageSrc(): Promise<string> {
    return this.getBase64();
  }

  /** The image part this picture embeds (extension). */
  imagePart(): ImagePart | undefined {
    const relId = this.relId;
    const source = this.paragraph.parentBody.part;
    if (!relId || !source) return undefined;
    const part = source.relationshipsPart?.getPart(relId);
    return part instanceof ImagePart ? part : undefined;
  }

  /** Removes the picture; the run goes too when the picture was all it held. */
  delete(): void {
    const content = this.run.value.content as Element[] | undefined;
    const at = content?.indexOf(this.element) ?? -1;
    if (content && at >= 0) content.splice(at, 1);
    if (!content || content.length === 0) {
      const holder = runHolderOf(this.paragraph.p, this.run);
      const i = holder ? holder.indexOf(this.run as Element) : -1;
      if (holder && i >= 0) holder.splice(i, 1);
    }
  }

  private pic(): pic.Pic | undefined {
    const any = this.inline.graphic.graphicData.any as Element<pic.Pic>[] | undefined;
    const first = any?.[0];
    return first && typeof first === 'object' && 'value' in first ? first.value : undefined;
  }

  private resize(cx: number | undefined, cy: number | undefined): void {
    const extent = this.inline.extent;
    const ratio = extent.cy / extent.cx;
    const w = cx ?? Math.max(1, Math.round((cy as number) / ratio));
    const h = cy ?? Math.max(1, Math.round((cx as number) * ratio));
    extent.cx = w;
    extent.cy = h;
    const ext = this.pic()?.spPr.xfrm?.ext;
    if (ext) { ext.cx = w; ext.cy = h; }
  }
}

/** The array holding a run inside a paragraph (the paragraph's content, or a holder's). */
function runHolderOf(p: object, run: Element<wml.R>): Element[] | undefined {
  const seek = (items: Element[] | undefined): Element[] | undefined => {
    if (!items) return undefined;
    if (items.includes(run as Element)) return items;
    for (const item of items) {
      const v = item.value as { content?: Element[]; customXmlOrSmartTagOrSdt?: Element[]; sdtContent?: { content?: Element[] } } | null;
      if (typeof v !== 'object' || v === null) continue;
      const nested = seek(v.content ?? v.customXmlOrSmartTagOrSdt ?? v.sdtContent?.content);
      if (nested) return nested;
    }
    return undefined;
  };
  return seek((p as { content?: Element[] }).content);
}

export { IMAGE_CONTENT_TYPES_BY_EXTENSION };
