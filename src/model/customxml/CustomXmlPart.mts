// CR-002 phase E, section 3.5: Word.CustomXmlPart and Word.CustomXmlNode as views over the
// custom XML data storage part's DOM (CR-001's CustomXmlDataStoragePart) and its properties part.
// Views, as Body and Paragraph are: the DOM is the state, nothing is cached but the prefix
// manager. Reading never marks the part for re-marshalling; every mutation does (markModified).
import { Docx4JException } from '../../opc/exceptions.mjs';
import type { CustomXmlDataStoragePart } from '../../parts/DefaultXmlPart.mjs';
import { parseXml, serializeXml, childElements } from '../../xml/dom.mjs';
import { type XPathEngine, canonicalXPathOf, parsePrefixMappings, formatPrefixMappings, xmlOf } from './xpath.mjs';

/** The namespaces of the property stores Word reports as built in (core, extended, cover page). */
const BUILT_IN_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
  'http://schemas.microsoft.com/office/2006/coverPageProps',
  'http://purl.org/dc/elements/1.1/',
]);

/** Office JS `Word.CustomXmlNodeType`, in the CR's spelling. */
export type CustomXmlNodeType = 'Element' | 'Attribute' | 'Text' | 'CData' | 'ProcessingInstruction' | 'Comment' | 'Document' | 'Other';

const NODE_TYPES: Readonly<Record<number, CustomXmlNodeType>> = {
  1: 'Element', 2: 'Attribute', 3: 'Text', 4: 'CData', 7: 'ProcessingInstruction', 8: 'Comment', 9: 'Document',
};

/** One prefix mapping of a part's namespace manager. */
export interface CustomXmlPrefixMapping {
  prefix: string;
  namespaceUri: string;
}

/**
 * Office JS `Word.CustomXmlPrefixMappingCollection`: the prefixes an XPath over this part may use.
 * Seeded from the namespace declarations of the document element, so that a part Word wrote can be
 * queried without declaring anything.
 */
export class CustomXmlPrefixMappingCollection {
  private readonly map = new Map<string, string>();

  constructor(seed: Record<string, string> = {}) {
    for (const [prefix, uri] of Object.entries(seed)) if (prefix !== '') this.map.set(prefix, uri);
  }

  get items(): CustomXmlPrefixMapping[] {
    return [...this.map].map(([prefix, namespaceUri]) => ({ prefix, namespaceUri }));
  }

  addNamespace(prefix: string, namespaceUri: string): void {
    this.map.set(prefix, namespaceUri);
  }

  /** The URI a prefix stands for, '' when it stands for none (Office JS lookupNamespace). */
  lookupNamespace(prefix: string): string {
    return this.map.get(prefix) ?? '';
  }

  /** The first prefix bound to a URI, '' when none is (Office JS lookupPrefix). */
  lookupPrefix(namespaceUri: string): string {
    for (const [prefix, uri] of this.map) if (uri === namespaceUri) return prefix;
    return '';
  }

  /** The mappings as an `xmlns:ns0='...'` string and as a map, for the engine. */
  toMap(): Record<string, string> {
    return Object.fromEntries(this.map);
  }
}

/** What a part view needs of its collection: the engine, and the removal of a part. */
export interface CustomXmlPartOwner {
  readonly xpathEngine: XPathEngine;
  removePart(part: CustomXmlPart): void;
}

/**
 * A subset of Office JS `Word.CustomXmlPart` over a `CustomXmlDataStoragePart` (the customer's XML,
 * a DOM) and its properties part (the `ds:datastoreItem` with the itemID and the schema
 * references). The properties part is read from its DOM, never unmarshalled, so it stays byte for
 * byte (as `Load` does for the itemID).
 */
export class CustomXmlPart {
  private prefixes: CustomXmlPrefixMappingCollection | undefined;

  constructor(
    /** The data part (extension: docx4j's part is the object here). */
    readonly part: CustomXmlDataStoragePart,
    private readonly owner: CustomXmlPartOwner,
    /** The itemID as the properties part spells it, '{...}'. */
    readonly itemID: string,
    /** The `ds:schemaRef` URIs of the properties part. */
    private readonly schemaRefs: string[] = [],
  ) {}

  /** The itemID, as Office JS reports it ('{...}'). */
  get id(): string {
    return this.itemID;
  }

  /** The DOM; throws when the part has not been parsed yet (`await pkg.customXmlParts.load()`). */
  get document(): Document {
    const document = this.part.parsedDocument;
    if (!document) {
      throw new Docx4JException(`Custom XML part ${this.part.partName} is not parsed yet: await pkg.customXmlParts.load() once`);
    }
    return document;
  }

  /** The document element as a node view. */
  get documentElement(): CustomXmlNode {
    const root = this.document.documentElement;
    if (!root) throw new Docx4JException(`Custom XML part ${this.part.partName} has no document element`);
    return new CustomXmlNode(root, this);
  }

  /** The namespace of the document element ('' when it is in none). */
  get namespaceUri(): string {
    return this.document.documentElement?.namespaceURI ?? '';
  }

  /** True for Word's own property stores (core, extended, custom and cover-page properties). */
  get builtIn(): boolean {
    return BUILT_IN_NAMESPACES.has(this.namespaceUri);
  }

  /** The prefixes an XPath over this part may use; seeded from the document element's declarations. */
  get namespaceManager(): CustomXmlPrefixMappingCollection {
    return (this.prefixes ??= new CustomXmlPrefixMappingCollection(declaredNamespaces(this.document.documentElement)));
  }

  /** The `ds:schemaRef` URIs of the properties part (Office JS schemaCollection). */
  get schemaCollection(): string[] {
    return [...this.schemaRefs];
  }

  /** The part's XML. Synchronous: a DOM part does not go through the marshaller. */
  getXml(): string {
    return serializeXml(this.document.documentElement ?? this.document);
  }

  /** Replaces the part's XML; the part is re-marshalled on save. */
  setXml(xml: string): void {
    this.part.setDocument(parseXml(xml));
    this.prefixes = undefined;
  }

  /** The nodes an XPath selects; `namespaceMappings` is Word's `xmlns:ns0='...'` string, else the namespace manager's. */
  selectNodes(xpath: string, namespaceMappings?: string): CustomXmlNode[] {
    return this.selectFrom(this.document, xpath, namespaceMappings);
  }

  selectSingleNode(xpath: string, namespaceMappings?: string): CustomXmlNode | undefined {
    return this.selectNodes(xpath, namespaceMappings)[0];
  }

  /**
   * Inserts `xml` as a child of the node the XPath selects, at `index` (appended when omitted).
   * Departure from Office JS, which puts `namespaceMappings` second: here it is last and optional,
   * as CR-002 section 3.5 specifies (section 12).
   */
  insertElement(xpath: string, xml: string, namespaceMappings?: string, index?: number): CustomXmlNode {
    const parent = this.requireNode(xpath, namespaceMappings);
    return parent.appendChildNode(xml, undefined, undefined, undefined, index);
  }

  /** Replaces the element the XPath selects with `xml`. */
  updateElement(xpath: string, xml: string, namespaceMappings?: string): void {
    const node = this.requireNode(xpath, namespaceMappings);
    const parent = node.node.parentNode;
    if (!parent) throw new Docx4JException(`${xpath} selects the document element; use setXml to replace the whole part`);
    parent.replaceChild(importedFragment(this.document, xml), node.node);
    this.touch();
  }

  /** Removes the element the XPath selects. */
  deleteElement(xpath: string, namespaceMappings?: string): void {
    const node = this.requireNode(xpath, namespaceMappings);
    node.delete();
  }

  /** Adds an attribute to the element the XPath selects. */
  insertAttribute(xpath: string, name: string, value: string, namespaceMappings?: string): void {
    this.updateAttribute(xpath, name, value, namespaceMappings);
  }

  /** Sets an attribute of the element the XPath selects (`prefix:name` is written in that prefix's namespace). */
  updateAttribute(xpath: string, name: string, value: string, namespaceMappings?: string): void {
    const node = this.requireNode(xpath, namespaceMappings);
    const element = node.node as Element;
    if (element.nodeType !== 1) throw new Docx4JException(`${xpath} does not select an element`);
    const prefix = name.includes(':') ? name.substring(0, name.indexOf(':')) : '';
    const uri = prefix ? (this.mappingsFor(namespaceMappings)[prefix] ?? this.namespaceManager.lookupNamespace(prefix)) : '';
    if (prefix && uri) element.setAttributeNS(uri, name, value);
    else element.setAttribute(name, value);
    this.touch();
  }

  /** Removes an attribute of the element the XPath selects. */
  deleteAttribute(xpath: string, name: string, namespaceMappings?: string): void {
    const node = this.requireNode(xpath, namespaceMappings);
    const element = node.node as Element;
    if (element.nodeType !== 1) throw new Docx4JException(`${xpath} does not select an element`);
    const prefix = name.includes(':') ? name.substring(0, name.indexOf(':')) : '';
    const uri = prefix ? (this.mappingsFor(namespaceMappings)[prefix] ?? this.namespaceManager.lookupNamespace(prefix)) : '';
    if (prefix && uri) element.removeAttributeNS(uri, name.substring(name.indexOf(':') + 1));
    else element.removeAttribute(name);
    this.touch();
  }

  /** Removes the part, its properties part and the relationships; bound controls lose their mapping. */
  delete(): void {
    this.owner.removePart(this);
  }

  // --- internals, shared with CustomXmlNode ---

  /** Marks the part for re-marshalling: a DOM read for a query alone leaves it byte for byte. */
  touch(): void {
    this.part.markModified();
  }

  get engine(): XPathEngine {
    return this.owner.xpathEngine;
  }

  /** The prefix map to use: the given string, else the namespace manager's. */
  mappingsFor(namespaceMappings: string | undefined): Record<string, string> {
    return namespaceMappings === undefined ? this.namespaceManager.toMap() : parsePrefixMappings(namespaceMappings);
  }

  selectFrom(context: Node, xpath: string, namespaceMappings?: string): CustomXmlNode[] {
    return this.engine.select(xpath, context, this.mappingsFor(namespaceMappings)).map((node) => new CustomXmlNode(node, this));
  }

  private requireNode(xpath: string, namespaceMappings?: string): CustomXmlNode {
    const node = this.selectSingleNode(xpath, namespaceMappings);
    if (!node) throw new Docx4JException(`${xpath} selects nothing in custom XML part ${this.part.partName}`);
    return node;
  }
}

/**
 * A subset of Office JS `Word.CustomXmlNode`: a view over one DOM node of a custom XML part.
 * `xpath` is the canonical `/ns0:a[1]/ns0:b[2]` form Word writes, which is what
 * `XmlMapping.setMappingByNode` stores.
 */
export class CustomXmlNode {
  constructor(
    /** The DOM node (extension). */
    readonly node: Node,
    /** The part the node belongs to. */
    readonly ownerPart: CustomXmlPart,
  ) {}

  /** The local name; for an attribute, its name. */
  get baseName(): string {
    if (this.node.nodeType === 2) return (this.node as Attr).localName ?? (this.node as Attr).name;
    return (this.node as Element).localName ?? this.node.nodeName;
  }

  get namespaceUri(): string {
    return (this.node as Element).namespaceURI ?? '';
  }

  get nodeType(): CustomXmlNodeType {
    return NODE_TYPES[this.node.nodeType] ?? 'Other';
  }

  /** The node's value: an attribute's value, a text node's data, an element's text. */
  get nodeValue(): string {
    return this.node.nodeType === 1 ? this.text : (this.node.nodeValue ?? '');
  }
  set nodeValue(value: string) {
    if (this.node.nodeType === 1) { this.text = value; return; }
    this.node.nodeValue = value;
    this.ownerPart.touch();
  }

  /** The node's text content; setting replaces an element's children with one text node. */
  get text(): string {
    return this.node.textContent ?? '';
  }
  set text(value: string) {
    this.node.textContent = value;
    this.ownerPart.touch();
  }

  /** The node's XML. */
  get xml(): string {
    return xmlOf(this.node);
  }

  /** The canonical XPath of this node, as Word writes it into `w:dataBinding`. */
  get xpath(): string {
    return canonicalXPathOf(this.node).xpath;
  }

  /** The prefix mappings that XPath needs (empty when the node is in no namespace). */
  get prefixMappings(): string {
    return canonicalXPathOf(this.node).prefixMappings;
  }

  get attributes(): CustomXmlNode[] {
    const attributes = (this.node as Element).attributes;
    const out: CustomXmlNode[] = [];
    for (let i = 0; attributes && i < attributes.length; i++) {
      const attr = attributes.item(i);
      if (attr && attr.name !== 'xmlns' && !attr.name.startsWith('xmlns:')) out.push(new CustomXmlNode(attr, this.ownerPart));
    }
    return out;
  }

  get childNodes(): CustomXmlNode[] {
    const out: CustomXmlNode[] = [];
    for (let child = this.node.firstChild; child; child = child.nextSibling) out.push(new CustomXmlNode(child, this.ownerPart));
    return out;
  }

  /** The element children only (extension: what a reader of a data store usually wants). */
  get childElements(): CustomXmlNode[] {
    return childElements(this.node).map((el) => new CustomXmlNode(el, this.ownerPart));
  }

  get parentNode(): CustomXmlNode | undefined {
    const parent = this.node.nodeType === 2 ? (this.node as Attr).ownerElement : this.node.parentNode;
    return parent && parent.nodeType !== 9 ? new CustomXmlNode(parent, this.ownerPart) : undefined;
  }

  get firstChild(): CustomXmlNode | undefined {
    return this.node.firstChild ? new CustomXmlNode(this.node.firstChild, this.ownerPart) : undefined;
  }

  get lastChild(): CustomXmlNode | undefined {
    return this.node.lastChild ? new CustomXmlNode(this.node.lastChild, this.ownerPart) : undefined;
  }

  get nextSibling(): CustomXmlNode | undefined {
    return this.node.nextSibling ? new CustomXmlNode(this.node.nextSibling, this.ownerPart) : undefined;
  }

  get previousSibling(): CustomXmlNode | undefined {
    return this.node.previousSibling ? new CustomXmlNode(this.node.previousSibling, this.ownerPart) : undefined;
  }

  hasChildNodes(): boolean {
    return this.node.firstChild !== null;
  }

  selectNodes(xpath: string, namespaceMappings?: string): CustomXmlNode[] {
    return this.ownerPart.selectFrom(this.node, xpath, namespaceMappings);
  }

  selectSingleNode(xpath: string, namespaceMappings?: string): CustomXmlNode | undefined {
    return this.selectNodes(xpath, namespaceMappings)[0];
  }

  /**
   * Appends a child. Both forms of CR-002 open question 8 are taken: the XML text form
   * `appendChildNode('<price>$20</price>')`, which is what an agent produces, and Office JS's
   * `appendChildNode(name, namespaceUri, nodeType, nodeValue)`.
   */
  appendChildNode(xml?: string, namespaceUri?: string, nodeType?: CustomXmlNodeType, nodeValue?: string, index?: number): CustomXmlNode {
    const created = this.create(xml, namespaceUri, nodeType, nodeValue);
    if (created.nodeType === 2) {
      (this.node as Element).setAttributeNode(created as Attr);
      this.ownerPart.touch();
      return new CustomXmlNode(created, this.ownerPart);
    }
    const children = childElements(this.node);
    const before = index === undefined ? null : children[index] ?? null;
    this.node.insertBefore(created, before);
    this.ownerPart.touch();
    return new CustomXmlNode(created, this.ownerPart);
  }

  /** Inserts a node before a sibling (appended when the sibling is omitted), as Office JS. */
  insertNodeBefore(xml: string, nextSibling?: CustomXmlNode): CustomXmlNode {
    const created = this.create(xml);
    this.node.insertBefore(created, nextSibling ? nextSibling.node : null);
    this.ownerPart.touch();
    return new CustomXmlNode(created, this.ownerPart);
  }

  removeChild(child: CustomXmlNode): void {
    if (child.node.nodeType === 2) {
      (this.node as Element).removeAttributeNode(child.node as Attr);
    } else {
      this.node.removeChild(child.node);
    }
    this.ownerPart.touch();
  }

  replaceChildNode(oldNode: CustomXmlNode, xml: string): CustomXmlNode {
    const created = this.create(xml);
    this.node.replaceChild(created, oldNode.node);
    this.ownerPart.touch();
    return new CustomXmlNode(created, this.ownerPart);
  }

  /** Removes this node from its parent. */
  delete(): void {
    const parent = this.parentNode;
    if (!parent) throw new Docx4JException('The document element cannot be deleted; delete the part instead');
    parent.removeChild(this);
  }

  private create(xml?: string, namespaceUri?: string, nodeType?: CustomXmlNodeType, nodeValue?: string): Node {
    const document = this.ownerPart.document;
    if (nodeType !== undefined || namespaceUri !== undefined) {
      const name = xml ?? '';
      if (name === '') throw new Docx4JException('A node needs a name');
      if (nodeType === 'Attribute') {
        const attr = namespaceUri ? document.createAttributeNS(namespaceUri, name) : document.createAttribute(name);
        attr.value = nodeValue ?? '';
        return attr;
      }
      if (nodeType === 'Text') return document.createTextNode(nodeValue ?? '');
      const element = namespaceUri ? document.createElementNS(namespaceUri, name) : document.createElement(name);
      if (nodeValue !== undefined && nodeValue !== '') element.appendChild(document.createTextNode(nodeValue));
      return element;
    }
    if (xml === undefined) throw new Docx4JException('Pass the XML of the node, or its name, namespace, type and value');
    return importedFragment(document, xml);
  }
}

/** Parses an XML fragment into a node of this document. */
function importedFragment(document: Document, xml: string): Node {
  const parsed = parseXml(xml);
  const root = parsed.documentElement;
  if (!root) throw new Docx4JException(`Not a well-formed XML element: ${xml}`);
  const imported = document.importNode ? document.importNode(root, true) : root;
  return imported;
}

/** The namespace declarations of an element, as a prefix map. */
function declaredNamespaces(element: Element | null): Record<string, string> {
  const out: Record<string, string> = {};
  if (!element) return out;
  const attributes = element.attributes;
  for (let i = 0; attributes && i < attributes.length; i++) {
    const attr = attributes.item(i);
    if (!attr) continue;
    if (attr.name === 'xmlns') out[''] = attr.value;
    else if (attr.name.startsWith('xmlns:')) out[attr.name.substring(6)] = attr.value;
  }
  return out;
}

export { formatPrefixMappings, parsePrefixMappings };
