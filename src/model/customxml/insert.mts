// CR-002 phase E, section 3.5: `insertContentControl(kind?)` on Body, Paragraph and Range.
// The `w:sdt` itself is built by the objects package's `sdt` / `sdtPr` / `nextSdtId` (objects
// CR-003 section 3.1, in 0.1.4); what is left here is what needs a part or a view.
import type { SdtKind } from '@docx4j/generated-objects-ts/builders/wml';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import { type Element, typeNameOf } from '../content/tree.mjs';
import type { ContentControlType } from '../content/ContentControl.mjs';
import { declareIgnorable } from '../content/tracking.mjs';

const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

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

/**
 * Lists `w15` in the `mc:Ignorable` of the body's part when the new control's `w:sdtPr` holds a
 * `w15` element (a repeating section or its item): Word 2010 knows no `w15` and opens a part with
 * one only when the prefix is ignorable (CR-002 section 40, `test/README.md` check 35).
 */
export function declareSdtExtensions(body: { part?: XmlPart<unknown> | undefined }, sdt: { value?: unknown }): void {
  const items = (sdt.value as { sdtPr?: { rPrOrAliasOrLock?: Array<{ name?: { namespaceURI?: string } }> } } | undefined)?.sdtPr?.rPrOrAliasOrLock;
  if (items?.some((item) => item.name?.namespaceURI === W15_NS)) declareIgnorable(body.part, 'w15');
}

/** True for a run-level item (what a Range wraps). */
export function isRunLevel(element: Element): boolean {
  const tn = typeNameOf(element);
  return tn === 'org_docx4j_wml.R' || tn === 'org_docx4j_wml.SdtRun' || tn === 'org_docx4j_wml.P.Hyperlink';
}
