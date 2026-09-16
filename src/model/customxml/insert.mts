// CR-002 phase E, section 3.5: `insertContentControl(kind?)` on Body, Paragraph and Range.
// The `w:sdt` itself is built by the objects package's `sdt` / `sdtPr` / `nextSdtId` (objects
// CR-003 section 3.1, in 0.1.4); what is left here is what needs a part or a view.
import type { SdtKind } from '@docx4j/generated-objects-ts/builders/wml';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import { type Element, typeNameOf } from '../content/tree.mjs';
import type { ContentControlType } from '../content/ContentControl.mjs';

/**
 * The kind the builders' `sdt` takes: Office JS's `Unknown` has no kind element, as `RichText`
 * has none (Word reports an untyped control as rich text).
 */
export function sdtKindFor(kind: ContentControlType | undefined): SdtKind | undefined {
  return kind === undefined || kind === 'Unknown' ? undefined : kind;
}

/**
 * The tree a new control's `w:id` must be free in (what `nextSdtId` scans): the part's whole
 * contents when the body has a part, so that a control added inside another control, a cell or a
 * header still gets an id free in the part; the container itself otherwise.
 */
export function controlIdScope(body: { part?: XmlPart<unknown> | undefined; container: object }): object {
  const part = body.part;
  return part?.isUnmarshalled ? (part.contents as object) : body.container;
}

/** True for a run-level item (what a Range wraps). */
export function isRunLevel(element: Element): boolean {
  const tn = typeNameOf(element);
  return tn === 'org_docx4j_wml.R' || tn === 'org_docx4j_wml.SdtRun' || tn === 'org_docx4j_wml.P.Hyperlink';
}
