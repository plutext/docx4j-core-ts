// CR-002 phase E, section 3.5: Word.XmlMapping over w:dataBinding (CTDataBinding), the link
// between a content control and a node of a custom XML part. A view, as everything else here is:
// the w:sdtPr is the state.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { sdtProperty } from '@docx4j/generated-objects-ts/builders/wml';
import type { Element } from '../content/tree.mjs';
import type { ContentControl } from '../content/ContentControl.mjs';
import type { CustomXmlPart, CustomXmlNode } from './CustomXmlPart.mjs';
import { canonicalXPathOf } from './xpath.mjs';

/** The Word 2012 (w15) namespace; not imported from `ContentControl`, which imports this module. */
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

/** What `XmlMapping` needs of `pkg.customXmlParts`, so that this module does not import the package. */
export interface CustomXmlPartLookup {
  readonly items: CustomXmlPart[];
  getItem(id: string): CustomXmlPart | undefined;
}

/**
 * A subset of Office JS `Word.XmlMapping`: the content control's data binding. `isMapped` is
 * whether the control has a `w:dataBinding`; `customXmlNode` evaluates the XPath now, so it needs
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
   * nothing when nothing is selected, which is what Word reports.
   */
  setMapping(xpath: string, prefixMappings?: string, part?: CustomXmlPart): boolean {
    const candidates = part ? [part] : this.candidates();
    for (const candidate of candidates) {
      const node = candidate.selectSingleNode(xpath, prefixMappings);
      if (!node) continue;
      this.write(xpath, prefixMappings ?? '', candidate.id);
      return true;
    }
    return false;
  }

  /** Binds to a node: its canonical XPath and the prefixes that path needs, as Word writes them. */
  setMappingByNode(node: CustomXmlNode): boolean {
    const { xpath, prefixMappings } = canonicalXPathOf(node.node);
    this.write(xpath, prefixMappings, node.ownerPart.id);
    return true;
  }

  /** Removes the binding (`w:dataBinding`); the control keeps the content it shows. */
  delete(): void {
    this.contentControl.removeProperty('dataBinding');
  }

  private write(xpath: string, prefixMappings: string, storeItemID: string): void {
    const binding: wml.CTDataBinding = { TYPE_NAME: 'org_docx4j_wml.CTDataBinding', xpath, storeItemID };
    if (prefixMappings !== '') binding.prefixMappings = prefixMappings;
    this.contentControl.putProperty(el.dataBinding(binding) as Element);
  }

  /** The parts to try, the customer's data before Word's property stores. */
  private candidates(): CustomXmlPart[] {
    const current = this.customXmlPart;
    const all = this.lookup?.items ?? [];
    const rest = all.filter((p) => p !== current);
    return [...(current ? [current] : []), ...rest.filter((p) => !p.builtIn), ...rest.filter((p) => p.builtIn)];
  }
}
