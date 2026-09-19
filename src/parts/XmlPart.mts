import { unmarshalNode, marshalNode, NAMESPACE_PREFIXES, type Jsonix } from '@docx4j/generated-objects-ts';
import { Part } from './Part.mjs';
import { PartName } from '../opc/PartName.mjs';
import { Namespaces } from './Namespaces.mjs';
import { Docx4JException } from '../opc/exceptions.mjs';
import { parseXml, serializeXml, encodeText, decodeXmlText, XML_DECLARATION } from '../xml/dom.mjs';
import { log } from '../model/properties/log.mjs';

const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
const MC_NS = Namespaces.MARKUP_COMPATIBILITY;

/** The objects package's prefix table the other way round, prefix to URI, built once. */
let prefixToUri: Map<string, string> | undefined;
function knownUriOf(prefix: string): string | undefined {
  if (prefixToUri === undefined) {
    prefixToUri = new Map();
    for (const [uri, p] of Object.entries(NAMESPACE_PREFIXES)) if (p !== '' && !prefixToUri.has(p)) prefixToUri.set(p, uri);
  }
  return prefixToUri.get(prefix);
}

const warnedPrefixes = new Set<string>();

/** A qualified name for a part's root element. */
export interface RootName {
  namespaceURI: string;
  localPart: string;
}

/** A hook run over the DOM before unmarshalling (the MCE preprocessor); set by the package on load. */
export type DomPreprocessor = (doc: Document, part?: XmlPart<unknown>) => void;

/**
 * A part whose content is a typed element of the object model (docx4j JaxbXmlPart<E>). `T` is the
 * root element's value type (`Document`, `Styles`, ...). Contents are unmarshalled on first
 * `getContents()`; once unmarshalled, `contents` is a plain property. A part that was never
 * unmarshalled is written back from its source bytes on save.
 */
export class XmlPart<T = unknown> extends Part {
  private element: Jsonix.TypedNamedValue<T> | undefined;
  private bytes: Uint8Array | undefined;
  /** Root element name used when contents are set on a new part; subclasses set it. */
  rootName: RootName | undefined;
  /** Runs over the DOM before unmarshalling; the package sets its MCE preprocessor here. */
  preprocessor: DomPreprocessor | undefined;
  /**
   * The namespace declarations the source root element carried, prefix to URI, recorded when the
   * part was parsed. They are what `mc:Ignorable` prefixes the object model does not bind are
   * re-declared from on marshal; see `declareIgnorablePrefixes`.
   */
  private sourceDeclarations: Map<string, string> | undefined;

  constructor(partName: PartName | string, contentType: string, relationshipType: string, rootName?: RootName) {
    super(partName, contentType, relationshipType);
    this.rootName = rootName;
  }

  /** True once contents were unmarshalled or set; the part is then re-marshalled on save. */
  get isUnmarshalled(): boolean {
    return this.element !== undefined;
  }

  get isLoaded(): boolean {
    return this.element !== undefined || this.bytes !== undefined;
  }

  /** The unmarshalled contents (docx4j getJaxbElement()); throws when not yet unmarshalled. */
  get contents(): T {
    if (this.element === undefined) throw new Docx4JException(`Part ${this.partName} is not unmarshalled yet; await getContents() first`);
    return this.element.value;
  }

  /** The root element as [name, value]; undefined until unmarshalled. */
  get typedElement(): Jsonix.TypedNamedValue<T> | undefined {
    return this.element;
  }

  /** Unmarshals on first call (docx4j getContents()). */
  async getContents(): Promise<T> {
    if (this.element === undefined) {
      const bytes = this.bytes ?? (await this.loadSourceBytes());
      if (bytes === undefined) throw new Docx4JException(`Part ${this.partName} has no content: set contents or load it from a container`);
      const doc = parseXml(decodeXmlText(bytes));
      this.recordDeclarations(doc);
      this.preprocessor?.(doc, this as XmlPart<unknown>);
      this.element = await unmarshalNode<Jsonix.TypedNamedValue<T>>(doc);
      this.rootName ??= { namespaceURI: this.element.name.namespaceURI, localPart: this.element.name.localPart };
      this.bytes = undefined;
    }
    return this.element.value;
  }

  /**
   * The contents as a **private tree**, without marking the part unmarshalled - so a part read
   * this way is still written back from its source bytes, byte for byte. The live tree is
   * returned when the part is already unmarshalled, since then the part is re-marshalled
   * anyway and a second copy would only go stale.
   *
   * This is what a reader that must not cost a part its round trip uses: `PropertyResolver`
   * reads the styles and numbering parts of a document nobody has otherwise touched.
   */
  async readContents(): Promise<T> {
    if (this.element !== undefined) return this.element.value;
    const bytes = this.bytes ?? (await this.loadSourceBytes());
    if (bytes === undefined) throw new Docx4JException(`Part ${this.partName} has no content: set contents or load it from a container`);
    const doc = parseXml(decodeXmlText(bytes));
    this.recordDeclarations(doc);
    this.preprocessor?.(doc, this as XmlPart<unknown>);
    return (await unmarshalNode<Jsonix.TypedNamedValue<T>>(doc)).value;
  }

  /** Replaces the contents (docx4j setJaxbElement / setContents). The root name defaults to the part's. */
  setContents(value: T, rootName?: RootName): void {
    const name = rootName ?? this.element?.name ?? this.rootName;
    if (!name) throw new Docx4JException(`Part ${this.partName}: no root element name; pass one`);
    this.element = { name: { namespaceURI: name.namespaceURI, localPart: name.localPart }, value };
    this.bytes = undefined;
  }

  /** Replaces the contents with raw XML bytes; they are unmarshalled on the next getContents(). */
  setBytes(bytes: Uint8Array): void {
    this.bytes = bytes;
    this.element = undefined;
  }

  /** Replaces the contents with XML text (docx4j setContents from a string). */
  setXml(xml: string): void {
    this.setBytes(encodeText(xml));
  }

  /** The XML as it would be saved (docx4j XmlUtils.marshaltoString on the part). */
  async getXml(): Promise<string> {
    return decodeXmlText(await this.getBytes());
  }

  /** Marshalled bytes when unmarshalled or set, else the source bytes untouched. */
  async getBytes(): Promise<Uint8Array> {
    if (this.element !== undefined) return encodeText(XML_DECLARATION + serializeXml(await this.marshalToNode()));
    if (this.bytes !== undefined) return this.bytes;
    const loaded = await this.loadSourceBytes();
    if (loaded === undefined) throw new Docx4JException(`Part ${this.partName} has no content and no source container`);
    return loaded;
  }

  /** The root element as a DOM element (docx4j XmlUtils.marshaltoW3CDomDocument). */
  async marshalToNode(): Promise<Element> {
    if (this.element === undefined) await this.getContents();
    const root = await marshalNode(this.element as Jsonix.TypedNamedValue);
    this.declareIgnorablePrefixes(root);
    return root;
  }

  /** The source root element's namespace declarations, prefix to URI (`xmlns=` is the empty prefix). */
  private recordDeclarations(doc: Document): void {
    const root = doc.documentElement;
    if (!root) return;
    const map = new Map<string, string>();
    for (const attr of Array.from(root.attributes)) {
      if (attr.name === 'xmlns') map.set('', attr.value);
      else if (attr.namespaceURI === XMLNS_NS || attr.name.startsWith('xmlns:')) map.set(attr.localName ?? attr.name.slice(6), attr.value);
    }
    this.sourceDeclarations = map;
  }

  /**
   * Every prefix the marshalled root's `mc:Ignorable` names must be declared on that root
   * (ECMA-376 Part 3 10.1.2; Word and Excel treat an undeclared one as a corrupt part and offer
   * to repair the file). The object model binds only the namespaces it has modules for, so an
   * `mc:Ignorable` prefix whose content was dropped on unmarshal - `xr6`, `xr10` and `xr2` on
   * Excel's `xl/workbook.xml`, for instance - loses its declaration: the runtime declares one
   * prefix per entry of the objects package's table and the facade can only *remove* unused
   * declarations, never add one it never made (CR-001 section 17).
   *
   * So the part re-declares them itself: from what the source root declared, else from the
   * objects package's `NAMESPACE_PREFIXES`. A prefix neither knows is dropped from
   * `mc:Ignorable` - an undeclared prefix is the one thing that must not be written - and warned
   * about once.
   */
  private declareIgnorablePrefixes(root: Element): void {
    const attr = Array.from(root.attributes).find((a) => a.namespaceURI === MC_NS && (a.localName ?? a.name) === 'Ignorable');
    if (!attr) return;
    const prefixes = attr.value.split(/\s+/).filter((p) => p !== '');
    if (prefixes.length === 0) return;
    const kept: string[] = [];
    let dropped = false;
    for (const prefix of prefixes) {
      if (root.getAttributeNS(XMLNS_NS, prefix) || root.getAttribute(`xmlns:${prefix}`)) { kept.push(prefix); continue; }
      const uri = this.sourceDeclarations?.get(prefix) ?? knownUriOf(prefix);
      if (uri === undefined) {
        dropped = true;
        if (!warnedPrefixes.has(prefix)) {
          warnedPrefixes.add(prefix);
          log.warn(`${this.partName}: mc:Ignorable names the prefix "${prefix}", which nothing declares; dropping it (an undeclared prefix makes Word and Excel offer to repair the file)`);
        }
        continue;
      }
      root.setAttributeNS(XMLNS_NS, `xmlns:${prefix}`, uri);
      kept.push(prefix);
    }
    if (!dropped) return;
    if (kept.length === 0) root.removeAttributeNode(attr);
    else root.setAttributeNS(MC_NS, attr.name, kept.join(' '));
  }
}
