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
// anchor, wrapping, position and extent (`BindingTraverserXSLT.xpathInjectImageRelId`; REQ-061).
// With a `width` parameter on the tag the content is replaced by a new inline picture
// (`xpathInjectImage`; REQ-062): at its natural size, scaled down to fit `width=N` twips or the
// text width, whichever is smaller; `width=auto` the text width alone. A control in a table cell
// fits the cell's width less its margins rather than the page's. An SVG value can replace a picture
// (no size is needed) but not be placed as a new one. The reverse direction writes a mapped picture
// control's image back to its node as base64, as Word 15 does (check 38); a tag-bound control is
// never written back, as docx4j's is not.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as dml from '@docx4j/generated-objects-ts/modules/org_docx4j_dml';
import { find, type Element } from '../content/tree.mjs';
import type { ContentControl } from '../content/ContentControl.mjs';
import { addImage, addImagePart, writableWidthEmu } from '../content/InlinePicture.mjs';
import { ImagePart } from '../../parts/BinaryPart.mjs';
import { base64Decode, base64Encode } from '../../xml/dom.mjs';
import { enclosingCellOf } from '../properties/table.mjs';
import type { CustomXmlPartLookup } from './XmlMapping.mjs';
import { tagParamsOf, selectEntryNode, PICTURE_HANDLER } from './opendope.mjs';
import type { BindingResult } from './bindings.mjs';

const EMU_PER_TWIP = 635;
const BLIP = 'org_docx4j_dml.CTBlip';

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
 * `a:blip`'s `r:embed` is pointed at a new image part and the drawing is left as it is; a picture
 * control holding no picture gets a new inline one at its natural size instead (docx4j's pre-v3
 * fallback), a tag-bound control holding none is left and noted. Otherwise the control's content
 * becomes one run holding a new inline picture, sized as `PictureWidth` says, in the first
 * paragraph of a block-level control (its properties kept) or as a run-level control's content.
 * An empty value, a value that is not an image this package reads (PNG, JPEG, GIF, BMP; SVG where
 * the drawing is kept, since it has no pixel size to place a new picture at), and a body with no
 * part each leave the control as it was, with a note. The image part the control showed before
 * stays in the package, as docx4j leaves it.
 */
export function bindPicture(control: ContentControl, value: string, width: PictureWidth, label: string): PictureOutcome {
  const base64 = base64Of(value);
  if (base64 === '') return { done: false, note: `${label}: the node is empty; the control is left as it is` };
  const source = control.parentBody.part;
  if (!source) return { done: false, note: `${label}: the control's body has no part to hold an image` };
  if (width.kind === 'keep') {
    const blip = firstBlipOf(control);
    if (blip) {
      try {
        blip.embed = addImagePart(source, base64).relId;
      } catch (e) {
        return { done: false, note: `${label}: not an image this package reads (${messageOf(e)})` };
      }
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

/** Binds one tag-bound control (shape b): the entry's node, the tag's width. */
export function applyPictureHandlerTo(control: ContentControl, parts: CustomXmlPartLookup): PictureOutcome {
  const params = tagParamsOf(control.tag);
  const entryId = params.get('od:xpath') ?? '';
  const selected = selectEntryNode(entryId, parts);
  if ('note' in selected) return { done: false, note: selected.note };
  return bindPicture(control, selected.node.text, pictureWidthOf(params), `od:xpath=${entryId}`);
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
 * first `a:blip` embeds, as base64, into the node. Nothing is written when the node already holds
 * those bytes (compared decoded, so Word's line-wrapped base64 counts as the same), when the
 * control shows no picture, or when the blip's part is not an image part. Asynchronous, since a
 * part's bytes may still be in the container.
 */
export async function updateFromPictureControl(control: ContentControl): Promise<boolean> {
  const mapping = control.xmlMapping;
  if (!mapping.isMapped || control.type !== 'Picture') return false;
  const node = mapping.customXmlNode;
  if (!node) return false;
  const blip = firstBlipOf(control);
  const source = control.parentBody.part;
  if (!blip?.embed || !source) return false;
  const part = source.relationshipsPart?.getPart(blip.embed);
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
    if (control.type !== 'Picture' || !control.xmlMapping.isMapped) continue;
    result.bound++;
    if (await updateFromPictureControl(control)) result.updated++; else result.skipped++;
  }
}
