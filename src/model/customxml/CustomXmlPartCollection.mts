// CR-002 phase E, section 3.5: `pkg.customXmlParts`, Office JS's Word.CustomXmlPartCollection over
// CR-001's `pkg.customXmlDataStorageParts`, plus docx4j's two binding directions.
//
// Asynchronous only where something is parsed or unmarshalled (section 3's rule): `load()` parses
// every custom XML part's DOM and warms the XPath engine, and `applyBindings` /
// `updateFromContentControls` unmarshal the parts whose controls they read. Everything else is
// synchronous, as Office JS is.
import { Docx4JException } from '../../opc/exceptions.mjs';
import { Namespaces } from '../../parts/Namespaces.mjs';
import { PartName } from '../../opc/PartName.mjs';
import type { Part } from '../../parts/Part.mjs';
import { CustomXmlDataStoragePart } from '../../parts/DefaultXmlPart.mjs';
import { CustomXmlDataStoragePropertiesPart, DS_NS } from '../../parts/customXml/index.mjs';
import { parseXml, decodeXmlText } from '../../xml/dom.mjs';
import { createDatastoreItem } from '@docx4j/generated-objects-ts/factory/org_docx4j_customXmlProperties';
import type { Body } from '../content/Body.mjs';
import type { ContentControl } from '../content/ContentControl.mjs';
import { CustomXmlPart, type CustomXmlPartOwner } from './CustomXmlPart.mjs';
import type { CustomXmlPartLookup } from './XmlMapping.mjs';
import type { XPathEngine } from './xpath.mjs';
import { applyBindingsTo, updateFromControls, type BindingResult } from './bindings.mjs';

/**
 * What the collection needs of its package. A structural interface, so that this module does not
 * import `WordprocessingMLPackage` (which imports it): the package hands itself in.
 */
export interface CustomXmlHost {
  readonly xpathEngine: XPathEngine;
  readonly customXmlDataStorageParts: Map<string, CustomXmlDataStoragePart>;
  /** The bodies whose content controls carry bindings; unmarshals the parts (main, headers, footers). */
  getBoundBodies(): Promise<Body[]>;
  /** The same, of the parts that are already unmarshalled: nothing is read from the container. */
  boundBodies(): Body[];
  /** Where a new custom XML part's relationship goes: the main document part (Word drops parts related from the package alone). */
  customXmlRelationshipSource(): Part;
}

/** A subset of Office JS `Word.CustomXmlPartCollection`. */
export class CustomXmlPartCollection implements CustomXmlPartOwner, CustomXmlPartLookup {
  private readonly views = new Map<CustomXmlDataStoragePart, CustomXmlPart>();

  constructor(private readonly host: CustomXmlHost) {}

  get xpathEngine(): XPathEngine {
    return this.host.xpathEngine;
  }

  /**
   * Parses every custom XML part's DOM and its properties part, and warms the XPath engine. One
   * await, and the collection and the node views are synchronous from then on (section 3.6).
   */
  async load(): Promise<CustomXmlPart[]> {
    for (const part of this.host.customXmlDataStorageParts.values()) {
      await part.parseDocument();
      if (!this.views.has(part)) this.views.set(part, await this.viewFor(part));
    }
    await this.host.xpathEngine.ready();
    return this.items;
  }

  /** The parts, in the order they were registered. Call `load()` once first: a DOM has to be parsed. */
  get items(): CustomXmlPart[] {
    const out: CustomXmlPart[] = [];
    for (const part of this.host.customXmlDataStorageParts.values()) {
      const view = this.views.get(part);
      if (!view) {
        throw new Docx4JException(`Custom XML part ${part.partName} is not parsed yet: await pkg.customXmlParts.load() once`);
      }
      out.push(view);
    }
    return out;
  }

  /** The parts whose document element is in that namespace (Office JS getByNamespace). */
  getByNamespace(namespaceUri: string): CustomXmlPart[] {
    return this.items.filter((part) => part.namespaceUri === namespaceUri);
  }

  /** The part with that itemID; the comparison ignores case and accepts the id with or without braces. */
  getItem(id: string): CustomXmlPart | undefined {
    const key = normaliseId(id);
    return this.items.find((part) => normaliseId(part.id) === key);
  }

  /**
   * A new custom XML part holding this XML, with its properties part and a fresh itemID, related
   * from the main document part (docx4j's `AbstractMigrator.addPropertiesPart` recipe: Word
   * silently drops a custom XML part that only the package relates to).
   */
  add(xml: string, options: { schemaRefs?: string[] } = {}): CustomXmlPart {
    const document = parseXml(xml);
    const root = document.documentElement;
    if (!root) throw new Docx4JException('A custom XML part needs a document element');

    const dataPart = new CustomXmlDataStoragePart('/customXml/item1.xml');
    dataPart.setDocument(document);
    this.host.customXmlRelationshipSource().addTargetPart(dataPart, 'RENAME_IF_NAME_EXISTS');

    const itemID = `{${uuid().toUpperCase()}}`;
    const propsName = PartName.of(`/customXml/itemProps${suffixOf(dataPart.partName)}.xml`);
    const props = new CustomXmlDataStoragePropertiesPart(propsName);
    const refs = options.schemaRefs ?? (root.namespaceURI ? [root.namespaceURI] : []);
    props.setContents(createDatastoreItem({
      itemID,
      ...(refs.length > 0 ? { schemaRefs: { schemaRef: refs.map((uri) => ({ uri })) } } : {}),
    }));
    dataPart.addTargetPart(props, 'RENAME_IF_NAME_EXISTS');

    dataPart.itemId = itemID.toLowerCase();
    this.host.customXmlDataStorageParts.set(dataPart.itemId, dataPart);
    const view = new CustomXmlPart(dataPart, this, itemID, refs);
    this.views.set(dataPart, view);
    return view;
  }

  /**
   * Removes a part, its properties part and the relationships (Office JS CustomXmlPart.delete).
   * Bindings to it are removed from the controls of the parts that are already unmarshalled; a
   * part that is not loaded keeps its `w:dataBinding`, which then names a part that is gone.
   */
  removePart(view: CustomXmlPart): void {
    const part = view.part;
    const source = this.host.customXmlRelationshipSource();
    source.relationshipsPart?.removePart(part.partName);
    part.package?.parts.remove(part.partName);
    if (part.itemId) this.host.customXmlDataStorageParts.delete(part.itemId);
    this.views.delete(part);
    for (const control of this.controlsOf(this.host.boundBodies())) {
      if (normaliseId(control.xmlMapping.storeItemID) === normaliseId(view.id)) control.xmlMapping.delete();
    }
  }

  /**
   * docx4j `BindingHandler.applyBindings`: pushes the custom XML values into the bound controls,
   * which is what Word does when it opens the document. Unmarshals the main document part (and the
   * headers and footers it has) and parses the custom XML parts.
   */
  async applyBindings(): Promise<BindingResult> {
    await this.load();
    return applyBindingsTo(this.controlsOf(await this.host.getBoundBodies()));
  }

  /**
   * docx4j `UpdateXmlFromDocumentSurface.updateCustomXmlParts`: writes what the controls now show
   * back into the custom XML. The reverse of `applyBindings`, and what makes an edit to a bound
   * control show in Word (CR-002 section 12).
   */
  async updateFromContentControls(): Promise<BindingResult> {
    await this.load();
    return updateFromControls(this.controlsOf(await this.host.getBoundBodies()));
  }

  private controlsOf(bodies: Body[]): ContentControl[] {
    return bodies.flatMap((body) => body.contentControls);
  }

  /** The view of a part: its itemID and schema references, read from the properties part's DOM. */
  private async viewFor(part: CustomXmlDataStoragePart): Promise<CustomXmlPart> {
    const rels = part.relationshipsPart;
    const rel = rels?.getRelationshipByType(Namespaces.CUSTOM_XML_DATA_STORAGE_PROPERTIES);
    const props = rel && rels ? rels.getPart(rel) : undefined;
    let itemID = part.itemId ?? '';
    const schemaRefs: string[] = [];
    if (props) {
      // Read the DOM rather than unmarshalling, so the properties part stays byte for byte (as Load does).
      const document = parseXml(decodeXmlText(await props.getBytes()));
      const root = document.documentElement;
      if (root) {
        itemID = root.getAttributeNS(DS_NS, 'itemID') || root.getAttribute('ds:itemID') || itemID;
        const refs = root.getElementsByTagNameNS(DS_NS, 'schemaRef');
        for (let i = 0; i < refs.length; i++) {
          const uri = refs[i]!.getAttributeNS(DS_NS, 'uri') ?? refs[i]!.getAttribute('ds:uri');
          if (uri) schemaRefs.push(uri);
        }
      }
    }
    return new CustomXmlPart(part, this, itemID, schemaRefs);
  }
}

/** An itemID without braces, lower-cased: how docx4j keys the map and how Office JS compares ids. */
function normaliseId(id: string): string {
  return id.replace(/[{}]/g, '').toLowerCase();
}

/** `itemProps<N>.xml` for `item<N>.xml`, so the two parts of a store keep the same number. */
function suffixOf(partName: PartName): string {
  const m = /item(\d*)\.xml$/.exec(partName.name);
  return m?.[1] ?? '';
}

/** A version 4 UUID, as docx4j's `UUID.randomUUID()` gives (crypto when the runtime has it). */
function uuid(): string {
  const c = (globalThis as { crypto?: { randomUUID?: () => string } }).crypto;
  if (c?.randomUUID) return c.randomUUID();
  const hex: string[] = [];
  for (let i = 0; i < 36; i++) {
    if (i === 8 || i === 13 || i === 18 || i === 23) { hex.push('-'); continue; }
    if (i === 14) { hex.push('4'); continue; }
    const r = Math.floor(Math.random() * 16);
    hex.push((i === 19 ? (r & 0x3) | 0x8 : r).toString(16));
  }
  return hex.join('');
}

