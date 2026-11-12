// SPDX-License-Identifier: Apache-2.0
//
// Picture bindings (CR-005 section 9; the specification's section 9.3, and Word's own picture
// content control). Two shapes bind a control to a node whose value is base64 image data:
//
//   (a) Word's picture content control, `w:sdtPr/w:picture` with a `w:dataBinding`: Word fills the
//       control's picture from the node when it opens the document;
//   (b) `od:Handler=picture` on the tag of a rich text control, bound through the tag's `od:xpath`
//       entry and carrying no `w:dataBinding` (Word has no floating picture control, and a rich text
//       control bound to the node would be filled with the base64 as text).
//
// docx4j's `bind.xslt`: modes `picture3` and `picture3richtext` replace the first `a:blip`'s
// `r:embed` with a relationship to a new image part and keep everything else about the drawing - its
// anchor, wrapping, position and extent (`BindingTraverserXSLT.xpathInjectImageRelId`; REQ-061). The
// blip itself is changed as Word changes it, not as docx4j does (`setBlipImage`, CR-002 section 52):
// its effects stay, which docx4j's bare `a:blip` drops, and an SVG goes under `asvg:svgBlip`.
// Word's own picture control, mapped (a), is the one exception to the kept extent: Word fits the node's picture inside
// the drawing's box, keeping its aspect ratio, whenever it fills one, flagged or not (CR-005 section 10.2, E1, the
// portfolio's bridge rounds core-ts-picture-fit and -2), so the bind writes the extent Word would show. A tag-bound
// control (b) keeps docx4j's extent; Word never fills one.
// With a `width` parameter on the tag the content is replaced by a new inline picture
// (`xpathInjectImage`; REQ-062): at its natural size, scaled down to fit `width=N` twips or the
// text width, whichever is smaller; `width=auto` the text width alone. A control in a table cell
// fits the cell's width less its margins rather than the page's. An SVG value can replace a picture
// (no size is needed) but not be placed as a new one. The reverse direction writes a mapped picture
// control's image back to its node as base64, as Word 15 does (check 38); a tag-bound control is
// never written back, as docx4j's is not.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as dml from '@docx4j/generated-objects-ts/modules/org_docx4j_dml';
import type * as wp from '@docx4j/generated-objects-ts/modules/org_docx4j_dml_wordprocessingDrawing';
import type * as pic from '@docx4j/generated-objects-ts/modules/org_docx4j_dml_picture';
import { find, isElement, typeNameOf, type Element } from '../content/tree.mjs';
import type { ContentControl } from '../content/ContentControl.mjs';
import { addImage, addImagePart, writableWidthEmu, setBlipImage, svgBlipOf, type ImageInfo } from '../content/InlinePicture.mjs';
import { ContentTypes } from '../../opc/ContentTypes.mjs';
import { ImagePart } from '../../parts/BinaryPart.mjs';
import { base64Decode, base64Encode } from '../../xml/dom.mjs';
import { enclosingCellOf } from '../properties/table.mjs';
import type { CustomXmlPartLookup } from './XmlMapping.mjs';
import { tagParamsOf, selectEntryNode, PICTURE_HANDLER } from './opendope.mjs';
import type { BindingResult } from './bindings.mjs';
import { showPicturePlaceholder } from './picturePlaceholder.mjs';

const EMU_PER_TWIP = 635;
const BLIP = 'org_docx4j_dml.CTBlip';
const INLINE = 'org_docx4j_dml_wordprocessingDrawing.Inline';
const ANCHOR = 'org_docx4j_dml_wordprocessingDrawing.Anchor';
const PIC = 'org_docx4j_dml_picture.Pic';
/** docx4j's ImageInfo default when a format carries no resolution, as `naturalSizeEmu` takes it. */
const DEFAULT_DPI = 96;

/**
 * What the tag's `width` says: absent keeps the control's drawing and replaces its image; `auto`
 * a new inline picture at its natural size, scaled down to the text width; `N` the same, scaled
 * down to N twips when that is narrower than the text width (a value that is neither is `auto`,
 * which is what docx4j does with one it cannot parse).
 */
export type PictureWidth = { kind: 'keep' } | { kind: 'auto' } | { kind: 'twips'; twips: number };

export interface PictureOutcome {
  done: boolean;
  /** Why the control was left as it was. */
  note?: string;
}

/** Whether a control is bound to a picture through its tag (`od:Handler=picture` with an `od:xpath`). */
export function isPictureHandlerBound(control: ContentControl): boolean {
  const params = tagParamsOf(control.tag);
  return params.get('od:Handler') === PICTURE_HANDLER && params.has('od:xpath');
}

/** The tag's `width` parameter, read as `PictureWidth` says. */
export function pictureWidthOf(params: Map<string, string>): PictureWidth {
  const width = params.get('width');
  if (width === undefined) return { kind: 'keep' };
  const twips = /^\d+$/.test(width.trim()) ? Number(width) : 0;
  return twips > 0 ? { kind: 'twips', twips } : { kind: 'auto' };
}

/** The node's value as base64: a data URL's prefix dropped, whitespace left to the lenient decoder. */
function base64Of(value: string): string {
  return value.replace(/^\s*data:[^,]*,/, '').trim();
}

function messageOf(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

/** The first `a:blip` anywhere in the control's content, whose `r:embed` is the image. */
function firstBlipOf(control: ContentControl): dml.CTBlip | undefined {
  for (const blip of find<dml.CTBlip>(control.sdt.sdtContent ?? {}, BLIP)) return blip;
  return undefined;
}

/**
 * The `wp:inline` or `wp:anchor` whose `pic:pic` the blip fills, and that picture: undefined where the blip is anywhere
 * else (a shape's fill, say), whose box this package doesn't refit.
 */
function frameOf(control: ContentControl, blip: dml.CTBlip): { frame: wp.Inline | wp.Anchor; picture: pic.Pic } | undefined {
  const content = control.sdt.sdtContent ?? {};
  for (const frame of [...find<wp.Inline>(content, INLINE), ...find<wp.Anchor>(content, ANCHOR)]) {
    const first = frame.graphic?.graphicData?.any?.[0];
    if (!isElement(first) || typeNameOf(first) !== PIC) continue;
    const picture = first.value as pic.Pic;
    if (picture.blipFill?.blip === blip) return { frame, picture };
  }
  return undefined;
}

/**
 * Word's fit of a picture swapped into a drawing (CR-005 section 10.2, E1): the image, at the aspect ratio its pixels and
 * resolution give, as large as fits inside the box `cx` by `cy`, one side the box's. Undefined for an image with no pixel
 * size (an SVG) or a box with no area: the box is kept.
 */
function fitInsideBox(cx: number, cy: number, info: ImageInfo): { cx: number; cy: number } | undefined {
  if (!(cx > 0 && cy > 0 && info.widthPx > 0 && info.heightPx > 0)) return undefined;
  const w = info.widthPx / (info.dpiX || DEFAULT_DPI);
  const h = info.heightPx / (info.dpiY || DEFAULT_DPI);
  // the box's width binds where cx / w <= cy / h
  return cx * h <= cy * w
    ? { cx, cy: Math.max(1, Math.round((cx * h) / w)) }
    : { cx: Math.max(1, Math.round((cy * w) / h)), cy };
}

/** Whether Word fills the control from its node on open: its own picture control, mapped by a `w:dataBinding`. */
function filledByWord(control: ContentControl): boolean {
  return control.type === 'Picture' && control.xmlMapping.dataBinding !== undefined;
}

/** A width stated in twips (`dxa`, or no type), else undefined. */
function twipsOf(width: wml.TblWidth | undefined): number | undefined {
  return width && typeof width.w === 'number' && width.w >= 0 && (width.type === undefined || width.type === 'dxa') ? width.w : undefined;
}

/**
 * Word's default cell margins, 108 twips left and right (the `Normal Table` style's
 * `w:tblCellMar`, which every table inherits), taken where neither the cell nor the table states
 * one; a table style stating other margins is not read.
 */
const DEFAULT_CELL_MARGIN_TWIPS = 108;

/** The cell's left and right margins together, in twips: `w:tcMar`, else the table's `w:tblCellMar`, else Word's defaults. */
function cellMarginsTwips(tc: wml.Tc, tbl: wml.Tbl): number {
  const own = tc.tcPr?.tcMar;
  const table = tbl.tblPr?.tblCellMar;
  const left = twipsOf(own?.left) ?? twipsOf(table?.left ?? table?.start) ?? DEFAULT_CELL_MARGIN_TWIPS;
  const right = twipsOf(own?.right) ?? twipsOf(table?.right ?? table?.end) ?? DEFAULT_CELL_MARGIN_TWIPS;
  return left + right;
}

/**
 * The width a new picture must fit, in EMU: the cell's `w:tcW` less its margins when the control
 * is in a table cell and the width is stated in twips (a percentage or `auto` falls to the page),
 * else the page's text width; narrowed to `width=N`. Undefined where nothing states a width (a
 * header's body has no `w:sectPr`): the natural size then.
 */
function maxWidthEmu(control: ContentControl, width: PictureWidth): number | undefined {
  const cell = enclosingCellOf(control.element.value as unknown as wml.P);
  const tcW = cell ? twipsOf(cell.tc.tcPr?.tcW) : undefined;
  const inner = cell && tcW !== undefined ? tcW - cellMarginsTwips(cell.tc, cell.tbl) : undefined;
  const cellEmu = inner !== undefined && inner > 0 ? inner * EMU_PER_TWIP : undefined;
  const text = cellEmu ?? writableWidthEmu(control.parentBody.container);
  if (width.kind === 'twips') {
    const wanted = width.twips * EMU_PER_TWIP;
    return text === undefined ? wanted : Math.min(wanted, text);
  }
  return text;
}

/**
 * Binds one control to a value: base64 image data (a data URL accepted). With `keep`, the first
 * `a:blip` is pointed at a new image part and the drawing is left as it is, as Word binds a picture
 * control (`setBlipImage`: a raster image in `r:embed`, the `asvg:svgBlip` taken away; an SVG under
 * the `asvg:svgBlip`, `r:embed` taken away; the effects kept; CR-002 section 52). Where Word fills the control, a mapped
 * picture control, the drawing's extent (`wp:extent`, and `a:ext` where the picture has one) becomes the image fitted
 * inside it, keeping its aspect ratio, as Word's is on open (`fitInsideBox`; CR-005 section 10.2, E1); a picture
 * control holding no picture gets a new inline one at its natural size instead (docx4j's pre-v3
 * fallback), a tag-bound control holding none is left and noted. Otherwise the control's content
 * becomes one run holding a new inline picture, sized as `PictureWidth` says, in the first
 * paragraph of a block-level control (its properties kept) or as a run-level control's content.
 * An empty value gives a picture control Word's form, the flag and a placeholder picture (`picturePlaceholderRun`;
 * CR-005 section 10, D1 (a2)), unless it shows one already; a tag-bound control it leaves, with a note. A value that is
 * not an image this package reads (PNG, JPEG, GIF, BMP; SVG where the drawing is kept, since it has no pixel size to
 * place a new picture at), and a body with no part each leave the control as it was, with a note. The image part the control showed before
 * stays in the package, as docx4j leaves it.
 */
export function bindPicture(control: ContentControl, value: string, width: PictureWidth, label: string): PictureOutcome {
  const base64 = base64Of(value);
  const source = control.parentBody.part;
  if (base64 === '') {
    // a picture control's empty node: Word's form, the flag and a placeholder picture (CR-005 section 10, D1 (a2));
    // one already showing a placeholder picture is left as it is, a tag-bound control left as before
    if (control.type !== 'Picture') return { done: false, note: `${label}: the node is empty; the control is left as it is` };
    if (control.isShowingPlaceholder && firstBlipOf(control)) return { done: false };
    if (!showPicturePlaceholder(control)) return { done: false, note: `${label}: the control's body has no part to hold an image` };
    return { done: true };
  }
  if (!source) return { done: false, note: `${label}: the control's body has no part to hold an image` };
  if (width.kind === 'keep') {
    const blip = firstBlipOf(control);
    if (blip) {
      let info: ImageInfo;
      try {
        const added = addImagePart(source, base64);
        setBlipImage(blip, added.relId, added.info.contentType === ContentTypes.IMAGE_SVG);
        info = added.info;
      } catch (e) {
        return { done: false, note: `${label}: not an image this package reads (${messageOf(e)})` };
      }
      if (filledByWord(control)) fitDrawing(control, blip, info);
      control.isShowingPlaceholder = false;
      return { done: true };
    }
    if (control.type !== 'Picture') return { done: false, note: `${label}: the control holds no a:blip to point at the image (od:Handler=picture without width keeps the drawing)` };
    width = { kind: 'auto' };
  }
  let run: Element;
  try {
    run = addImage(source, base64, control.parentBody.container, maxWidthEmu(control, width)).run as Element;
  } catch (e) {
    return { done: false, note: `${label}: not an image this package reads (${messageOf(e)})` };
  }
  control.setBoundContent([run]);
  return { done: true };
}

/** The drawing the blip fills given the image's extent fitted inside its own (E1); nothing where it has no box to fit. */
function fitDrawing(control: ContentControl, blip: dml.CTBlip, info: ImageInfo): void {
  const held = frameOf(control, blip);
  const extent = held?.frame.extent;
  const fitted = extent && fitInsideBox(extent.cx, extent.cy, info);
  if (!held || !extent || !fitted) return;
  extent.cx = fitted.cx;
  extent.cy = fitted.cy;
  const ext = held.picture.spPr?.xfrm?.ext;
  if (ext) { ext.cx = fitted.cx; ext.cy = fitted.cy; }
}

/** Binds one tag-bound control (shape b): the entry's node, the tag's width. */
export function applyPictureHandlerTo(control: ContentControl, parts: CustomXmlPartLookup): PictureOutcome {
  const params = tagParamsOf(control.tag);
  const entryId = params.get('od:xpath') ?? '';
  const selected = selectEntryNode(entryId, parts);
  if ('note' in selected) return { done: false, note: selected.note };
  const outcome = bindPicture(control, selected.node.text, pictureWidthOf(params), `od:xpath=${entryId}`);
  // a stale storeItemID resolved as Word would (CR-020): said where the outcome has nothing else to say
  return selected.repair && outcome.done && !outcome.note ? { ...outcome, note: selected.repair } : outcome;
}

/** The picture pass over the tag-bound controls, after the text bindings: counted into the result, what was left said in its notes. */
export function applyPictureHandlersTo(controls: ContentControl[], parts: CustomXmlPartLookup, result: BindingResult): void {
  for (const control of controls) {
    if (!isPictureHandlerBound(control)) continue;
    result.bound++;
    const outcome = applyPictureHandlerTo(control, parts);
    if (outcome.done) result.updated++; else result.skipped++;
    if (outcome.note) (result.notes ??= []).push(outcome.note);
  }
}

/**
 * The reverse direction for a mapped picture control (docx4j's `UpdateXmlFromDocumentSurface`
 * skips pictures; Word 15 writes the picture back, check 38): the bytes of the image part the
 * first `a:blip` shows, as base64, into the node: its `asvg:svgBlip`'s SVG where it has one, as
 * `InlinePicture.getBase64` gives (CR-002 section 52), else its `r:embed`'s image. Nothing is written when the node already holds
 * those bytes (compared decoded, so Word's line-wrapped base64 counts as the same), when the
 * control shows no picture, or when the blip's part is not an image part. Asynchronous, since a
 * part's bytes may still be in the container.
 */
export async function updateFromPictureControl(control: ContentControl): Promise<boolean> {
  const mapping = control.xmlMapping;
  // a control showing its placeholder picture holds no picture of its node's (CR-005 section 10)
  if (!mapping.dataBinding || control.type !== 'Picture' || control.isShowingPlaceholder) return false;
  const node = mapping.customXmlNode;
  if (!node) return false;
  const blip = firstBlipOf(control);
  const source = control.parentBody.part;
  const relId = svgBlipOf(blip)?.embed || blip?.embed;
  if (!relId || !source) return false;
  const part = source.relationshipsPart?.getPart(relId);
  if (!(part instanceof ImagePart)) return false;
  const bytes = await part.getBytes();
  const held = base64Decode(node.text);
  if (held.length === bytes.length && held.every((b, i) => b === bytes[i])) return false;
  node.text = base64Encode(bytes);
  return true;
}

/** `updateFromContentControls`' picture pass: the mapped picture controls, counted into the result. */
export async function updateFromPictureControls(controls: ContentControl[], result: BindingResult): Promise<void> {
  for (const control of controls) {
    if (control.type !== 'Picture' || !control.xmlMapping.dataBinding) continue;
    result.bound++;
    if (await updateFromPictureControl(control)) result.updated++; else result.skipped++;
  }
}
