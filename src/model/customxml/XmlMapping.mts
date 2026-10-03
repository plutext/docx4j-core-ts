// CR-002 phase E, section 3.5: Word.XmlMapping over w:dataBinding (CTDataBinding), the link
// between a content control and a node of a custom XML part. A view, as everything else here is:
// the w:sdtPr is the state.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import * as w15el from '@docx4j/generated-objects-ts/el/org_docx4j_w15';
import { sdtProperty } from '@docx4j/generated-objects-ts/builders/wml';
import type { Element } from '../content/tree.mjs';
import type { ContentControl } from '../content/ContentControl.mjs';
import type { CustomXmlPart, CustomXmlNode } from './CustomXmlPart.mjs';
import { canonicalXPathOf } from './xpath.mjs';
import { applyBindingTo } from './bindings.mjs';

/** The Word 2012 (w15) namespace; not imported from `ContentControl`, which imports this module. */
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** What `XmlMapping` needs of `pkg.customXmlParts`, so that this module does not import the package. */
export interface CustomXmlPartLookup {
  readonly items: CustomXmlPart[];
  getItem(id: string): CustomXmlPart | undefined;
}

/**
 * A subset of Office JS `Word.XmlMapping`: the content control's data binding. `isMapped` is
 * whether the control has a `w:dataBinding` (or, on a rich text control or a repeating section,
 * the `w15:dataBinding` Word writes there, which is what `setMapping` writes on one: CR-002
 * section 40); `customXmlNode` evaluates the XPath now, so it needs
 * the XPath engine to be warm (`await pkg.customXmlParts.load()` once, CR-002 section 3.6).
 */
export class XmlMapping {
  constructor(
    /** The control this mapping belongs to. */
    readonly contentControl: ContentControl,
    private readonly lookup: CustomXmlPartLookup | undefined,
  ) {}

  /** The `w:dataBinding` value, or undefined when the control has none (extension). */
  get dataBinding(): wml.CTDataBinding | undefined {
    const pr = this.contentControl.sdt.sdtPr;
    // Word 2013 writes `w15:dataBinding` on a repeating section and on a rich-text control bound
    // to a container; everywhere else the binding is `w:dataBinding`. Both are CT_DataBinding.
    const item = sdtProperty(pr, 'dataBinding') ?? sdtProperty(pr, 'dataBinding', W15_NS);
    return item?.value as wml.CTDataBinding | undefined;
  }

  get isMapped(): boolean {
    return this.dataBinding !== undefined;
  }

  get xpath(): string {
    return this.dataBinding?.xpath ?? '';
  }

  get prefixMappings(): string {
    return this.dataBinding?.prefixMappings ?? '';
  }

  /** `w:storeItemID`, the itemID of the custom XML part (extension: Office JS reports the part). */
  get storeItemID(): string {
    return this.dataBinding?.storeItemID ?? '';
  }

  /** The part the binding names, when the package has it. */
  get customXmlPart(): CustomXmlPart | undefined {
    const id = this.storeItemID;
    return id === '' ? undefined : this.lookup?.getItem(id);
  }

  /** The node the XPath selects now, or undefined when it selects nothing. */
  get customXmlNode(): CustomXmlNode | undefined {
    const part = this.customXmlPart;
    const binding = this.dataBinding;
    if (!part || !binding) return undefined;
    return part.selectSingleNode(binding.xpath, binding.prefixMappings);
  }

  /**
   * Binds the control to what the XPath selects, as Word's own mapping does: the part is the one
   * given, else the one the control is already bound to, else the first custom XML part (the
   * built-in property stores last) in which the XPath selects a node. Returns false and changes
   * nothing when nothing is selected, which is what Word reports. **The node's value then goes into
   * the control**, whatever text it held: Office JS's `setMapping` on a plain text control holding
   * "quick brown", on an empty one and on one already holding the value, and Word's XML Mapping
   * Pane mapping an existing control, each showed the node's "Ann" and left the part unchanged
   * (`test/README.md` check 34, 2026-10-03; CR-002 section 39). Before, only the binding was
   * written and the values stayed apart until `applyBindings`.
   */
  setMapping(xpath: string, prefixMappings?: string, part?: CustomXmlPart): boolean {
    const candidates = part ? [part] : this.candidates();
    for (const candidate of candidates) {
      const node = candidate.selectSingleNode(xpath, prefixMappings);
      if (!node) continue;
      this.write(xpath, prefixMappings ?? '', candidate.id);
      applyBindingTo(this.contentControl);
      return true;
    }
    return false;
  }

  /** Binds to a node: its canonical XPath and the prefixes that path needs, as Word writes them; the node's value goes into the control (check 34). */
  setMappingByNode(node: CustomXmlNode): boolean {
    const { xpath, prefixMappings } = canonicalXPathOf(node.node);
    this.write(xpath, prefixMappings, node.ownerPart.id);
    applyBindingTo(this.contentControl);
    return true;
  }

  /** Removes the binding (`w:dataBinding` and `w15:dataBinding`); the control keeps the content it shows. */
  delete(): void {
    this.contentControl.removeProperty('dataBinding', W_NS);
    this.contentControl.removeProperty('dataBinding', W15_NS);
  }

  /**
   * Whether the control's binding is `w15:dataBinding`, as Word 2013 and later write it: on a rich
   * text control (one with `w:richText`, or with no kind element at all) and on a repeating
   * section. `w:dataBinding` on a rich text control makes it a plain text control to Word 15 and
   * Word 2010 alike - shown from its node on open, given `<w:text/>` on saving, an edit written
   * back as text (`test/README.md` check 35, 1f; CR-002 section 40).
   */
  get usesW15(): boolean {
    const type = this.contentControl.type;
    return type === 'RichText' || type === 'RepeatingSection';
  }

  /** Writes the binding in the element the control's kind takes, replacing one in the other namespace where it stood. */
  private write(xpath: string, prefixMappings: string, storeItemID: string): void {
    const binding: wml.CTDataBinding = { TYPE_NAME: 'org_docx4j_wml.CTDataBinding', xpath, storeItemID };
    if (prefixMappings !== '') binding.prefixMappings = prefixMappings;
    const w15 = this.usesW15;
    const element = (w15 ? w15el.dataBinding(binding) : el.dataBinding(binding)) as Element;
    const items = this.contentControl.sdt.sdtPr?.rPrOrAliasOrLock as Element[] | undefined;
    const other = w15 ? W_NS : W15_NS;
    const at = items?.findIndex((i) => i.name.localPart === 'dataBinding' && i.name.namespaceURI === other) ?? -1;
    if (items && at >= 0) {
      if (items.some((i) => i.name.localPart === 'dataBinding' && i.name.namespaceURI === element.name.namespaceURI)) items.splice(at, 1);
      else items[at] = element;
    }
    // putProperty also lists w15 in the part's mc:Ignorable when the element is a w15 one
    this.contentControl.putProperty(element);
  }

  /** The parts to try, the customer's data before Word's property stores. */
  private candidates(): CustomXmlPart[] {
    const current = this.customXmlPart;
    const all = this.lookup?.items ?? [];
    const rest = all.filter((p) => p !== current);
    return [...(current ? [current] : []), ...rest.filter((p) => !p.builtIn), ...rest.filter((p) => p.builtIn)];
  }
}
