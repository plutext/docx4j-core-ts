// SPDX-License-Identifier: Apache-2.0
//
// The picture a picture control shows while it has none: a mapped one whose node is empty, and one just made (CR-005
// section 10, D1 (a2) and D2). Word gives such a
// control its own form on opening, whatever the file holds (docx4j's check 71): the flag, `w:showingPlcHdr`, and one run
// holding an inline drawing, 1270000 by 1270000 EMU, of a placeholder picture. Word's picture is its own artwork, so this
// package's is docx4j's: the same pale blue, 200 by 200 pixels, with a grey, dash-edged inner square (docx4j's
// `org/docx4j/model/datastorage/picture-placeholder.png`, 676 bytes, SHA-1 e9e4ee21363e5c8d0f3200fd920b03d5e66e4ab7;
// Jason's choice, 2026-10-25). Word replaces it with its own on opening; elsewhere it is what shows.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Part } from '../../parts/Part.mjs';
import { ImagePart } from '../../parts/BinaryPart.mjs';
import type { Element } from '../content/tree.mjs';
import type { ContentControl } from '../content/ContentControl.mjs';
import { addImagePart, imageRunOf } from '../content/InlinePicture.mjs';

/** docx4j's placeholder picture, a PNG, as base64. */
export const PICTURE_PLACEHOLDER_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAIAAAAiOjnJAAAACXBIWXMAABYlAAAWJQFJUiTwAAACVklEQVR42u3cwU3jYBCGYRtx50gDNOUOckwDOaWBHNOBm6IAuEEBSByQIkv5B1njgGLP8xx3vasVeuVh82m3f/v46uDWHnwJEBbCQlggLISFsEBYCAthgbAQFsICYSEshAXCQlgU87jkF59Px+indvuDx9b+2PTH/y+sX/6gXdd9vr/O+U08ds+PnU/HdFtOIX+iT/8rncsbaxgGX8eNGcdx4TX0xsLfChEWwgJhISwq65f8pyAzP39jpZ6eX7yxcAoRFuTcYIQ26WyPSQenEGGBsBAWwgJhISxqMEITMkLjFCIsSDNC02CExilEWCAshIWwQFgIixqM0ISM0DiFCAvSjNA0GKFxChEWCAthISwQFsKiBiM0ISM0TiHCgjQjNA1GaJxChAXCQlgIC4SFsKjBCE3ICI1TiLAgzQhNgxEapxBhgbAQFsICYSEsajBCEzJC4xQiLEgzQtNghMYpRFggLISFsEBYCIsajNCEjNA4hQgL0ozQNBihcQoRFggLYSEsEBbCogYjNCEjNE4hwoI0IzQNRmicQoQFwkJYCAuEhbCowQhNyAiNU4iwIM0ITYMRGqcQYYGwEBbCAmEhLGowQhMyQuMUIixIM0LTYITGKURYICyEhbBAWAiLGozQhIzQOIUIC9KM0DQYoXEKERYIC2EhLBAWwqIGIzQhIzROIcKCNCM0DUZoNvfGuq57avoaaz7gsdU95o3Fmt9Yu/3h8m3WtZkfgXjsnh9Lf4PVLfyAFJxChIWwQFgIC2GBsBAWwgJhISyEBcJCWAgLhIWwEBb8+AbsaRgh6xNNrAAAAABJRU5ErkJggg==';

/** The drawing's extent, Word's: 1270000 EMU (100 points) square. */
export const PICTURE_PLACEHOLDER_EMU = 1270000;

/** The relationship each source part's placeholder image was added under, so that every control shares one part. */
const added = new WeakMap<Part, string>();

/** The run holding the placeholder drawing in `source` (`scope` its drawing ids'), the image part added once per part. */
export function picturePlaceholderRun(source: Part, scope: object): Element<wml.R> {
  let relId = added.get(source);
  if (relId === undefined || !(source.relationshipsPart?.getPart(relId) instanceof ImagePart)) {
    relId = addImagePart(source, PICTURE_PLACEHOLDER_BASE64).relId;
    added.set(source, relId);
  }
  return imageRunOf(relId, scope, PICTURE_PLACEHOLDER_EMU, PICTURE_PLACEHOLDER_EMU);
}

/**
 * Word's form for a picture control (`w:picture`): the flag, and its content one run holding the placeholder drawing, in
 * a paragraph where the control is block-level. What a mapped control whose node is empty shows (CR-005 section 10, D1
 * (a2)) and what `insertContentControl('Picture')` makes (D2), as Word's do. False, and nothing changed, where the
 * control's body has no part to hold the image.
 */
export function showPicturePlaceholder(control: ContentControl): boolean {
  const source = control.parentBody.part;
  if (!source) return false;
  control.setBoundContent([picturePlaceholderRun(source, control.parentBody.container) as Element]);
  control.isShowingPlaceholder = true;
  return true;
}
