// CR-002 phase C: inline pictures. The counterpart of docx4j's
// BinaryPartAbstractImage.createImageInline (and of its ImageInfo, whose job the header readers
// below do): an ImagePart, a relationship from the part the body belongs to, and a
// w:drawing/wp:inline sized from the image's pixel dimensions and resolution.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as dml from '@docx4j/generated-objects-ts/modules/org_docx4j_dml';
import type * as wp from '@docx4j/generated-objects-ts/modules/org_docx4j_dml_wordprocessingDrawing';
import type * as pic from '@docx4j/generated-objects-ts/modules/org_docx4j_dml_picture';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import { inlinePicture } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { ContentTypes, IMAGE_CONTENT_TYPES_BY_EXTENSION } from '../../opc/ContentTypes.mjs';
import { base64Decode, base64Encode, decodeXmlText } from '../../xml/dom.mjs';
import { ImagePart } from '../../parts/BinaryPart.mjs';
import { PartName } from '../../opc/PartName.mjs';
import type { Part } from '../../parts/Part.mjs';
import { type Element, type RevisionHolder, find, runOf, linkParents, revisionKindOf, isElement, typeNameOf, isAlternateContent, selectedBranchOf, assertOutsideBranch } from './tree.mjs';
import { type ChangeTracker, copyRPr } from './tracking.mjs';
import { type RelationshipRemovals, relationshipIdsIn, queueRelationshipRemovals } from './relationships.mjs';
import type { Paragraph } from './Paragraph.mjs';

/** English Metric Units per inch, and per twip (docx4j UnitsOfMeasurement). */
const EMU_PER_INCH = 914400;
const EMU_PER_TWIP = 635;
/** docx4j's ImageInfo default when a format carries no resolution. */
const DEFAULT_DPI = 96;
const EMU_PER_POINT = 12700;

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
 * information, and the same 96 dpi default when a file declares none. An SVG document is told by
 * its root element and has no pixel size (both 0): it can be the part behind an existing picture
 * (CR-005 section 9, a binding that keeps the drawing) but not be placed as a new one.
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
  const svg = svgInfo(bytes);
  if (svg) return svg;
  const head = Array.from(bytes.subarray(0, 8)).map((b) => b.toString(16).padStart(2, '0')).join(' ');
  throw new Docx4JException(`Unsupported image format (first bytes ${head}); PNG, JPEG, GIF, BMP and SVG are read here`);
}

/** An SVG document: an XML declaration, comments and a doctype allowed before the `svg` root. */
function svgInfo(b: Uint8Array): ImageInfo | undefined {
  if (b.length < 5) return undefined;
  let head: string;
  try { head = decodeXmlText(b.subarray(0, Math.min(b.length, 4096))); } catch { return undefined; }
  const body = head.replace(/^\uFEFF/, '').replace(/<\?xml[^>]*\?>/, '').replace(/<!--[\s\S]*?-->/g, '').replace(/<!DOCTYPE[^>]*>/i, '').trimStart();
  if (!/^<(\w+:)?svg[\s>]/.test(body)) return undefined;
  return { contentType: ContentTypes.IMAGE_SVG, extension: 'svg', widthPx: 0, heightPx: 0, dpiX: DEFAULT_DPI, dpiY: DEFAULT_DPI };
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
  /** wp:docPr/@title, Office JS altTextTitle; not written when absent. */
  altTextTitle?: string;
  /** The width in points; the height follows the aspect ratio unless it is given too. */
  width?: number;
  height?: number;
}

/** The run holding a new picture, plus the image part it added (the caller inserts the run). */
export interface NewPicture {
  run: Element<wml.R>;
  drawing: Element<wml.Drawing>;
  imagePart: ImagePart;
  relId: string;
}

/** A new image part, its relationship from the part it was added to, and what its header says. */
export interface NewImagePart {
  imagePart: ImagePart;
  relId: string;
  info: ImageInfo;
}

/**
 * Adds the image as a part of `source` (the main document part, or the header or footer the body
 * belongs to) with a relationship, and nothing else: docx4j `BinaryPartAbstractImage.createImagePart`
 * (`BindingTraverserXSLT.createImagePartReturnRelId`, where a picture binding replaces an `r:embed`
 * and keeps the drawing, CR-005 section 9). Throws, and adds nothing, when the bytes are not an
 * image this package reads.
 */
export function addImagePart(source: Part, base64: string): NewImagePart {
  const pkg = source.package;
  if (!pkg) throw new Docx4JException('This part is not in a package yet; add it before adding an image');
  const bytes = base64Decode(base64);
  const info = imageInfoOf(bytes);
  const imagePart = new ImagePart(freeImageName(pkg, info.extension), info.contentType);
  imagePart.setBytes(bytes);
  const rel = source.addTargetPart(imagePart);
  return { imagePart, relId: rel.id, info };
}

/**
 * Adds the image as a part of `source` with a relationship (`addImagePart`), and builds the run
 * that shows it. docx4j: `BinaryPartAbstractImage.createImagePart` then `createImageInline`.
 */
export function addImage(source: Part, base64: string, scope: object, maxCx: number | undefined, options: InlinePictureOptions = {}): NewPicture {
  const bytes = base64Decode(base64);
  const info = imageInfoOf(bytes);
  if (info.widthPx === 0 || info.heightPx === 0) {
    throw new Docx4JException(`A ${info.extension.toUpperCase()} image has no pixel size to place it at; PNG, JPEG, GIF and BMP can be placed`);
  }
  const { imagePart, relId } = addImagePart(source, base64);
  const rel = { id: relId };
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
  const drawing = inlinePicture(rel.id, {
    cx, cy, id, name,
    descr: options.altTextDescription ?? '',
    ...(options.altTextTitle === undefined ? {} : { title: options.altTextTitle }),
  }) as Element<wml.Drawing>;
  return { run: runOf([drawing]), drawing, imagePart, relId: rel.id };
}

/** `/word/media/imageN.<ext>`, the first N free in the package (docx4j getNewPartName). */
function freeImageName(pkg: { getPart(name: PartName | string): unknown }, extension: string): PartName {
  for (let i = 1; ; i++) {
    const name = PartName.of(`/word/media/image${i}.${extension}`);
    if (!pkg.getPart(name)) return name;
  }
}

/**
 * The run showing an image already related from the part as `relId`, an inline picture `cx` by `cy` EMU, its drawing id
 * the next free one in `scope` and named `Picture <id>`, as `addImage` names one.
 */
export function imageRunOf(relId: string, scope: object, cx: number, cy: number): Element<wml.R> {
  const id = nextDrawingId(scope);
  return runOf([inlinePicture(relId, { cx, cy, id, name: `Picture ${id}`, descr: '' }) as Element]);
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

/** The TYPE_NAME of `asvg:svgBlip`, Office 2016's SVG picture extension of an `a:blip`. */
const SVG_BLIP = 'org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main.CTSVGBlip';
/** The `a:ext` URI that `asvg:svgBlip` stands under, and its namespace. */
const SVG_BLIP_EXT_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
const SVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';

/** The `asvg:svgBlip` in a blip's `a:extLst`, if any: an SVG picture's SVG (CR-002 section 52). */
export function svgBlipOf(blip: dml.CTBlip | undefined): { embed?: string } | undefined {
  for (const ext of blip?.extLst?.ext ?? []) {
    const any: unknown = ext.any;
    if (isElement(any) && typeNameOf(any) === SVG_BLIP) return any.value as { embed?: string };
  }
  return undefined;
}

/**
 * Points a blip at a new image the way Word's own picture binding does (test/README.md checks 46 and
 * 47; CR-002 section 52). The blip's effects and other extensions stay.
 * - A raster image goes in `r:embed`, and an `asvg:svgBlip` is taken away. Left in place, Word would go
 *   on drawing its old SVG, and would replace the new image with a PNG of that SVG when it saves.
 * - An SVG goes under the `asvg:svgBlip`, and `r:embed` is taken away. Word puts a PNG of the SVG it
 *   makes itself in `r:embed`, which this package cannot make, and Word 365 draws the SVG without one.
 * `r:link` is left as it was.
 */
export function setBlipImage(blip: dml.CTBlip, relId: string, svg: boolean): void {
  if (!svg) {
    blip.embed = relId;
    const list = blip.extLst;
    if (list?.ext) {
      list.ext = list.ext.filter((ext) => !(isElement(ext.any) && typeNameOf(ext.any) === SVG_BLIP));
      if (list.ext.length === 0) delete blip.extLst;
    }
    return;
  }
  delete blip.embed;
  const existing = svgBlipOf(blip);
  if (existing) {
    existing.embed = relId;
    return;
  }
  const list = (blip.extLst ??= { TYPE_NAME: 'org_docx4j_dml.CTOfficeArtExtensionList', ext: [] } as dml.CTOfficeArtExtensionList);
  const ext = {
    TYPE_NAME: 'org_docx4j_dml.CTOfficeArtExtension',
    uri: SVG_BLIP_EXT_URI,
    any: { name: { namespaceURI: SVG_NS, localPart: 'svgBlip' }, value: { TYPE_NAME: SVG_BLIP, embed: relId } },
  } as unknown as dml.CTOfficeArtExtension;
  (list.ext ??= []).push(ext);
  linkParents(list, blip);
}

const FORMAT_BY_CONTENT_TYPE: Readonly<Record<string, ImageFormat>> = {
  [ContentTypes.IMAGE_PNG]: 'Png', [ContentTypes.IMAGE_X_PNG]: 'Png', [ContentTypes.IMAGE_JPEG]: 'Jpeg',
  [ContentTypes.IMAGE_GIF]: 'Gif', [ContentTypes.IMAGE_BMP]: 'Bmp', [ContentTypes.IMAGE_TIFF]: 'Tiff',
  [ContentTypes.IMAGE_SVG]: 'Svg', [ContentTypes.IMAGE_EMF]: 'Emf', [ContentTypes.IMAGE_EMF2]: 'Emf',
  [ContentTypes.IMAGE_WMF]: 'Wmf',
};

/**
 * The inline pictures a run's item holds: the item itself when it is a `w:drawing` whose `wp:inline`
 * holds a `pic:pic`, and those in a kept `mc:AlternateContent`'s selected branch (`selectedBranchOf`,
 * CR-010 phase 3). A drawing of anything else, such as a chart, a text box, a group or SmartArt, is no
 * inline picture: Office JS lists none for an inline text box (CR-010 section 12.6, test/README.md
 * check 42).
 */
export function picturesIn(item: Element): Element<wml.Drawing>[] {
  if (isAlternateContent(item)) {
    return (selectedBranchOf(item.value as mce.AlternateContent) ?? []).flatMap((child) => (isElement(child) ? picturesIn(child) : []));
  }
  return item.name.localPart === 'drawing' && holdsInlinePicture(item.value as wml.Drawing) ? [item as Element<wml.Drawing>] : [];
}

/** True for a drawing whose `wp:inline` holds a `pic:pic`. */
function holdsInlinePicture(drawing: wml.Drawing): boolean {
  const inline = drawing.anchorOrInline?.[0] as wp.Inline | undefined;
  if (!inline || inline.TYPE_NAME !== 'org_docx4j_dml_wordprocessingDrawing.Inline') return false;
  const first = inline.graphic?.graphicData?.any?.[0];
  return isElement(first) && typeNameOf(first) === 'org_docx4j_dml_picture.Pic';
}

/**
 * A subset of Office JS `Word.InlinePicture` over a `w:drawing` holding a `wp:inline` with a `pic:pic`.
 * A view: the drawing, the run it is in and the paragraph. Sizes are points, as Office JS reports them.
 *
 * A picture in a kept `mc:AlternateContent`'s selected branch has the element as its `holder`. Reads and
 * the setters are the selected branch's drawing; `delete` removes the element whole, every branch with
 * it (CR-010 phase 3).
 */
export class InlinePicture {
  constructor(
    /** The `w:drawing` element pair. */
    readonly element: Element<wml.Drawing>,
    /** The run holding the drawing, and the array holding the run. */
    readonly run: Element<wml.R>,
    readonly paragraph: Paragraph,
    /** The run's item the drawing stands in, where that is not the drawing: a kept `mc:AlternateContent`. */
    readonly holder?: Element,
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
    this.assertWritable();
    this.inline.docPr.descr = text;
    const p = this.pic();
    if (p) p.nvPicPr.cNvPr.descr = text;
  }

  /** wp:docPr/@title (typed since objects 0.1.3; before that the attribute was not marshalled, CR-002 section 8). */
  get altTextTitle(): string {
    return this.inline.docPr.title ?? '';
  }
  set altTextTitle(text: string) {
    this.assertWritable();
    this.inline.docPr.title = text;
  }

  /**
   * The kind of image, from the content type of the part `getBase64` reads: `Svg` for an SVG picture
   * (`svgImagePart`), else the `a:blip`'s part's. Office JS gives `null` for an SVG picture, though its
   * typings declare the property non-null and `Word.ImageFormat` has `Svg` (test/README.md check 45);
   * this package gives `Svg`, a departure (CR-002 section 52).
   */
  get imageFormat(): ImageFormat {
    const part = this.svgImagePart() ?? this.imagePart();
    return (part && FORMAT_BY_CONTENT_TYPE[part.contentType]) ?? 'Unsupported';
  }

  /** The relationship id of the embedded image (r:embed), if any (extension). */
  get relId(): string | undefined {
    return this.pic()?.blipFill.blip?.embed;
  }

  /**
   * The relationship id of the SVG the picture embeds beside its image, if any (extension): the
   * `r:embed` of an `asvg:svgBlip` in the `a:blip`'s `a:extLst`, as Office 2016 and later write an SVG
   * picture, with the `a:blip`'s own image a PNG of it for readers that don't understand SVG.
   */
  get svgRelId(): string | undefined {
    const embed = svgBlipOf(this.pic()?.blipFill.blip)?.embed;
    return embed === undefined || embed === '' ? undefined : embed;
  }

  /**
   * The image's bytes as base64 (Office JS getBase64ImageSrc). For an SVG picture, the SVG
   * (`svgImagePart`), as Office JS gives it (test/README.md check 43); else the `a:blip`'s image.
   */
  async getBase64(): Promise<string> {
    const part = this.svgImagePart() ?? this.imagePart();
    if (!part) throw new Docx4JException('This picture has no image part (a linked or missing image)');
    return base64Encode(await part.getBytes());
  }

  /** Office JS's name for `getBase64()`. */
  getBase64ImageSrc(): Promise<string> {
    return this.getBase64();
  }

  /** The image part this picture embeds: the `a:blip`'s, which is an SVG picture's PNG (extension). */
  imagePart(): ImagePart | undefined {
    return this.partFor(this.relId);
  }

  /** The SVG part an SVG picture embeds beside its image (`svgRelId`), if any (extension). */
  svgImagePart(): ImagePart | undefined {
    return this.partFor(this.svgRelId);
  }

  private partFor(relId: string | undefined): ImagePart | undefined {
    const source = this.paragraph.parentBody.part;
    if (!relId || !source) return undefined;
    const part = source.relationshipsPart?.getPart(relId);
    return part instanceof ImagePart ? part : undefined;
  }

  /**
   * Removes the picture; the run goes too when the picture was all it held. A picture in a kept
   * `mc:AlternateContent` takes the element with it, every branch (`holder`).
   *
   * While the package tracks changes it is marked deleted instead: its run goes into a `w:del`, and
   * when the run held more than the picture, the picture is first given a run of its own with the
   * same properties, so the deletion holds it alone. Accepting the deletion then takes it, and its
   * image part with it (CR-002 section 46). A picture its author inserted under tracking is taken
   * back, as `Paragraph.delete` takes back its own text, and one inside a deletion is refused.
   *
   * A picture removed outright (tracking off, or taken back) queues its image relationship on the
   * package. On save the relationship goes if nothing then references it, and the image part with it
   * when nothing else targets it. Waiting for the save lets a picture cut and pasted back keep its
   * image.
   */
  delete(): void {
    this.assertWritable();
    const place = placeOf(this.paragraph.p, this.run);
    if (!place) return;
    const revision = place.revision;
    const tracker = this.paragraph.changeTracker;
    if (tracker && (revision?.kind === 'del' || revision?.kind === 'moveFrom')) {
      throw new Docx4JException(`This picture is inside a w:${revision.kind} (deleted by ${revision.value.author}); reject the deletion first`);
    }
    if (tracker && !tracker.ownInsertion(revision)) {
      this.markDeleted(tracker, place);
      return;
    }
    const item = this.holder ?? this.element;
    const content = this.run.value.content as Element[] | undefined;
    const at = content?.indexOf(item) ?? -1;
    if (content && at >= 0) content.splice(at, 1);
    if (!content || content.length === 0) {
      const i = place.items.indexOf(this.run as Element);
      if (i >= 0) place.items.splice(i, 1);
      if (revision && revision.items.length === 0) {                     // an insertion left empty goes with it
        const j = revision.owner.indexOf(revision.element);
        if (j >= 0) revision.owner.splice(j, 1);
      }
    }
    const ids = new Set<string>();
    relationshipIdsIn(item, ids);
    const pkg = this.paragraph.parentBody.package_ as { pendingRelationshipRemovals?: RelationshipRemovals } | undefined;
    queueRelationshipRemovals(pkg?.pendingRelationshipRemovals, this.paragraph.parentBody.part, ids);
  }

  /** The picture's run in a `w:del`, the picture split into a run of its own when the run held more. */
  private markDeleted(tracker: ChangeTracker, place: Place): void {
    const run = this.run;
    const at = place.items.indexOf(run as Element);
    const content = (run.value.content ?? []) as Element[];
    const item = this.holder ?? this.element as Element;
    const i = content.indexOf(item);
    if (at < 0 || i < 0) return;
    const before = content.slice(0, i);
    const after = content.slice(i + 1);
    const pieces: Element[] = [];
    let alone: Element<wml.R> = run;
    if (before.length > 0 || after.length > 0) {
      alone = runOf([item], copyRPr(run.value.rPr));
      if (before.length > 0) {
        run.value.content = before as wml.R['content'];
        pieces.push(run as Element);
      }
    }
    pieces.push(tracker.del([alone as Element]));
    if (after.length > 0) pieces.push(runOf(after, copyRPr(run.value.rPr)) as Element);
    place.items.splice(at, 1, ...pieces);
    linkParents(pieces, place.parent);
  }

  private pic(): pic.Pic | undefined {
    const any = this.inline.graphic.graphicData.any as Element<pic.Pic>[] | undefined;
    const first = any?.[0];
    return first && typeof first === 'object' && 'value' in first ? first.value : undefined;
  }

  /**
   * Refuses a write to a picture in a paragraph in a branch of a kept block-, row- or cell-level
   * `mc:AlternateContent` (CR-013 D4). A picture in a run-level element's branch is the run's own, and
   * deleting it removes the element (CR-010 phase 3).
   */
  private assertWritable(): void {
    assertOutsideBranch(this.paragraph.p, 'InlinePicture', 'this picture');
  }

  private resize(cx: number | undefined, cy: number | undefined): void {
    this.assertWritable();
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
/** Where a run stands: the array holding it, what that array belongs to, and the revision it is in, if any. */
interface Place { items: Element[]; parent: object; revision?: RevisionHolder }

function placeOf(p: wml.P, run: Element<wml.R>): Place | undefined {
  const seek = (items: Element[] | undefined, parent: object, revision: RevisionHolder | undefined): Place | undefined => {
    if (!items) return undefined;
    if (items.includes(run as Element)) return revision ? { items, parent, revision } : { items, parent };
    for (const item of items) {
      const v = item.value as { content?: Element[]; customXmlOrSmartTagOrSdt?: Element[]; sdtContent?: { content?: Element[] } } | null;
      if (typeof v !== 'object' || v === null) continue;
      const holder = v.sdtContent ?? v;
      const nested = v.content ?? v.customXmlOrSmartTagOrSdt ?? v.sdtContent?.content;
      const kind = revisionKindOf(item);
      const inner = kind && nested ? { kind, element: item, value: v as unknown as wml.CTTrackChange, items: nested, owner: items } : revision;
      const found = seek(nested, holder, inner);
      if (found) return found;
    }
    return undefined;
  };
  return seek(p.content as Element[] | undefined, p, undefined);
}

export { IMAGE_CONTENT_TYPES_BY_EXTENSION };
