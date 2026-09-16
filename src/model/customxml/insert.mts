// CR-002 phase E, section 3.5: `insertContentControl(kind?)` on Body, Paragraph and Range.
// The element half lives here so that the three views share one builder and one id allocator; the
// views themselves only decide what to wrap.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import * as w14el from '@docx4j/generated-objects-ts/el/org_docx4j_w14';
import * as w15el from '@docx4j/generated-objects-ts/el/org_docx4j_w15';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { type Element, typeNameOf, find } from '../content/tree.mjs';
import type { ContentControlType } from '../content/ContentControl.mjs';

/** The `w:sdtPr` child that types a control, by the kind Word reports. */
function kindElement(kind: ContentControlType): Element | undefined {
  switch (kind) {
    case 'PlainText': return el.text({}) as Element;
    case 'RichText': return el.richText({}) as Element;
    case 'Picture': return el.picture({}) as Element;
    case 'BuildingBlockGallery': return el.docPartObj({}) as Element;
    case 'CheckBox': return w14el.checkbox({ checked: { val: '0' }, checkedState: { val: '2612', font: 'MS Gothic' }, uncheckedState: { val: '2610', font: 'MS Gothic' } }) as Element;
    case 'ComboBox': return el.comboBox({}) as Element;
    case 'DropDownList': return el.dropDownList({}) as Element;
    case 'DatePicker': return el.date({ dateFormat: { val: 'd/MM/yyyy' }, storeMappedDataAs: { val: 'dateTime' } }) as Element;
    case 'RepeatingSection': return w15el.repeatingSection({}) as Element;
    case 'RepeatingSectionItem': return w15el.repeatingSectionItem({}) as Element;
    case 'Group': return el.group({}) as Element;
    case 'Citation': return el.citation({}) as Element;
    case 'Bibliography': return el.bibliography({}) as Element;
    case 'Equation': return el.equation({}) as Element;
    default: return undefined;   // 'Unknown' and the untyped case: Word reports a rich-text control
  }
}

/** A free `w:id` for a new control: unique among the ids in the tree, as Word's are. */
export function nextControlId(root: object): number {
  const used = new Set<number>();
  for (const sdtPr of find<wml.SdtPr>(root, 'org_docx4j_wml.SdtPr')) {
    for (const item of sdtPr.rPrOrAliasOrLock ?? []) {
      if (item.name.localPart !== 'id') continue;
      const val = (item.value as { val?: number }).val;
      if (typeof val === 'number') used.add(val);
    }
  }
  let id = 0;
  do {
    id = 1 + Math.floor(Math.random() * 0x3fffffff);
  } while (used.has(id));
  return id;
}

/** The `w:sdtPr` of a new control: the id, and the element that types it (none for rich text). */
export function sdtPrFor(kind: ContentControlType | undefined, id: number, options: { tag?: string; title?: string } = {}): wml.SdtPr {
  const items: Element[] = [];
  if (options.title !== undefined) items.push(el.alias({ val: options.title }) as Element);
  if (options.tag !== undefined) items.push(el.tag({ val: options.tag }) as Element);
  items.push(el.id({ val: id }) as Element);
  const typed = kind === undefined || kind === 'RichText' ? undefined : kindElement(kind);
  if (typed) items.push(typed);
  return { TYPE_NAME: 'org_docx4j_wml.SdtPr', rPrOrAliasOrLock: items as wml.SdtPr['rPrOrAliasOrLock'] };
}

/** A block-level `w:sdt` (SdtBlock) wrapping these block-level elements. */
export function sdtBlockFor(content: Element[], kind: ContentControlType | undefined, id: number): Element<wml.SdtBlock> {
  const value: wml.SdtBlock = {
    TYPE_NAME: 'org_docx4j_wml.SdtBlock',
    sdtPr: sdtPrFor(kind, id),
    sdtContent: { TYPE_NAME: 'org_docx4j_wml.SdtContentBlock', content: content as wml.SdtContentBlock['content'] },
  };
  return { name: { namespaceURI: W_NS, localPart: 'sdt' }, value };
}

/** A run-level `w:sdt` (SdtRun) wrapping these run-level items. */
export function sdtRunFor(content: Element[], kind: ContentControlType | undefined, id: number): Element<wml.SdtRun> {
  const value: wml.SdtRun = {
    TYPE_NAME: 'org_docx4j_wml.SdtRun',
    sdtPr: sdtPrFor(kind, id),
    sdtContent: { TYPE_NAME: 'org_docx4j_wml.CTSdtContentRun', content: content as wml.CTSdtContentRun['content'] },
  };
  return { name: { namespaceURI: W_NS, localPart: 'sdt' }, value };
}

/** Rejects a kind that only makes sense at run level in a block control, and the reverse. */
export function checkKind(kind: ContentControlType | undefined, form: 'Block' | 'Run'): void {
  if (kind === 'RepeatingSection' && form === 'Run') {
    throw new Docx4JException('A repeating section is a block-level control; insert it on a body or a paragraph');
  }
}

/** True for a run-level item (what a Range wraps). */
export function isRunLevel(element: Element): boolean {
  const tn = typeNameOf(element);
  return tn === 'org_docx4j_wml.R' || tn === 'org_docx4j_wml.SdtRun' || tn === 'org_docx4j_wml.P.Hyperlink';
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
