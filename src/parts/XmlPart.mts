import { unmarshalNode, marshalNode, NAMESPACE_PREFIXES, type Jsonix, type MarshalOptions, type WordReading } from '@docx4j/generated-objects-ts';
import { Part, type PartBytesOptions } from './Part.mjs';
import { PartName } from '../opc/PartName.mjs';
import { Docx4JException, PartUnmarshalException } from '../opc/exceptions.mjs';
import { parseXml, serializeXml, encodeText, decodeXmlText, XML_DECLARATION, assertXmlCharacters } from '../xml/dom.mjs';
import { bindSnapshot, boundsOf, type Bound } from '../model/content/binder.mjs';
import type { PartWarning, ValueKeptWarning } from '../opc/mce/preprocessor.mjs';

/**
 * The element or attribute a located runtime error names, out of its message: jsonix writes
 * `w:sz/@w:val (line 7): ...` when it could place the fault and `w:bogus (line 3): ...` for an
 * element, so the location is everything before the position or the colon - and `w:sz` must not be
 * cut at its own prefix separator, which a naive split on `:` does.
 */
/** The runtime's located message: `w:sz/@w:val (line 12): reason`, the position optional. */
const LOCATION_PREFIX = /^(.+?)(?: \(line \d+(?:, column \d+)?\))?: /;

function locationOf(message: string, located: boolean): string | undefined {
  if (!located) return undefined;
  return LOCATION_PREFIX.exec(message)?.[1];
}

/** ` (line 12, column 9)`, or as much of it as is known. */
function positionOf(line: number | undefined, column: number | undefined): string {
  if (line === undefined) return '';
  return column === undefined ? ` (line ${line})` : ` (line ${line}, column ${column})`;
}

/** A value the objects facade read by Word's rule, as a warning on the part (CR-015 D8). */
function valueKept(reading: WordReading, partName: string): ValueKeptWarning {
  const node = reading.element as Element & { lineNumber?: number; columnNumber?: number };
  const name = reading.attribute;
  const attribute = node.getAttributeNodeNS?.(name.namespaceURI || null, name.localPart)?.name ?? name.localPart;
  const value = reading.value;
  const read = value === true ? 'on' : value === undefined ? 'absent' : JSON.stringify(value);
  // a form Word refuses a document for is dropped (the objects package's CR-015; CR-017 P1)
  const outcome = reading.kind === 'value.refused' ? 'dropped: Word 365 refuses a document with this value' : `read as ${read}, as Word 365 reads it`;
  return {
    code: reading.kind, partName,
    ...(typeof node.lineNumber === 'number' ? { line: node.lineNumber } : {}),
    ...(typeof node.columnNumber === 'number' ? { column: node.columnNumber } : {}),
    element: node.nodeName, attribute, text: reading.text, value,
    message: `${partName}: ${node.nodeName}/@${attribute}=${JSON.stringify(reading.text)}${positionOf(node.lineNumber, node.columnNumber)} is not a value its type allows; ${outcome}`,
  };
}

/** What `selectObjects` needs: the engine to evaluate with, and any prefixes beyond the defaults. */
export interface SelectOptions {
  /** The engine; `pkg.selectObjects` and `Body.select` pass the package's. It is readied on the snapshot. */
  engine?: {
    ready(context?: Node): Promise<void>;
    select(expression: string, context: Node, namespaces?: Record<string, string>): Node[];
  } | undefined;
  /** Prefixes added to (or overriding) the objects package's own table. */
  namespaces?: Record<string, string> | undefined;
}

/** The objects package's prefix table as XPath wants it, prefix to URI; built once. */
const NAMESPACE_PREFIXES_BY_PREFIX: Record<string, string> = (() => {
  const out: Record<string, string> = {};
  for (const [uri, prefix] of Object.entries(NAMESPACE_PREFIXES)) if (prefix !== '' && !(prefix in out)) out[prefix] = uri;
  return out;
})();

/** A qualified name for a part's root element. */
export interface RootName {
  namespaceURI: string;
  localPart: string;
}

/**
 * A hook run over the DOM before unmarshalling (the MCE preprocessor); set by the package on load.
 * What it returns becomes the part's `warnings`.
 */
export type DomPreprocessor = (doc: Document, part?: XmlPart<unknown>) => void | readonly PartWarning[];

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
  private parseWarnings: readonly PartWarning[] = [];

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

  /**
   * True while the part holds nothing but its source: never unmarshalled, no contents or bytes set.
   * Such a part is written back byte for byte, and what a private read finds in it cannot change,
   * which is what lets a reader remember it (CR-002 section 31).
   */
  get isUntouched(): boolean {
    return this.element === undefined && this.bytes === undefined;
  }

  /**
   * What the last parse of this part reported: an mc:AlternateContent whose branches' text differs
   * (`mce.branches_differ`, CR-010 phase 1), and each attribute value its type refuses that was read as
   * Word 365 reads it (`value.kept.exact`, `value.kept.invalid`: CR-015 D8). Empty until the part is
   * parsed, by `getContents()` or a private read, and again once its contents are set.
   */
  get warnings(): readonly PartWarning[] {
    return this.parseWarnings;
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
      const doc = this.parse(bytes);
      this.preprocess(doc);
      this.element = await this.unmarshal(doc);
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
    const doc = this.parse(bytes);
    this.preprocess(doc);
    return (await this.unmarshal(doc)).value;
  }

  private preprocess(doc: Document): void {
    const warnings = this.preprocessor?.(doc, this as XmlPart<unknown>);
    this.parseWarnings = warnings ? [...warnings] : [];
  }

  /** Replaces the contents (docx4j setJaxbElement / setContents). The root name defaults to the part's. */
  setContents(value: T, rootName?: RootName): void {
    const name = rootName ?? this.element?.name ?? this.rootName;
    if (!name) throw new Docx4JException(`Part ${this.partName}: no root element name; pass one`);
    this.element = { name: { namespaceURI: name.namespaceURI, localPart: name.localPart }, value };
    this.bytes = undefined;
    this.parseWarnings = [];
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
  async getBytes(options?: PartBytesOptions): Promise<Uint8Array> {
    if (this.element !== undefined) return encodeText(assertXmlCharacters(XML_DECLARATION + serializeXml(await this.marshalToNode(options)), `Part ${this.partName}`));
    if (this.bytes !== undefined) return this.bytes;
    const loaded = await this.loadSourceBytes();
    if (loaded === undefined) throw new Docx4JException(`Part ${this.partName} has no content and no source container`);
    return loaded;
  }

  /**
   * The objects an XPath selects in this part, through a DOM marshalled for the purpose
   * (CR-006; docx4j `JaxbXmlPart.getJAXBNodesViaXPath`, which does it with a JAXB `Binder`).
   *
   * Each result carries the node the expression selected, the tree object behind it, the attribute
   * name when an attribute was selected, and the node's path from the root. The DOM is a **snapshot**:
   * changing it changes nothing, and a select leaves the part's own bytes alone - the marshal is
   * thrown away. Every call marshals afresh, since the tree has no change log a cached snapshot could
   * be checked against (CR-006 section 2).
   *
   * The prefixes default to the ones the objects package writes, so `//w:p[w:pPr/w:pStyle/@w:val='Heading1']`
   * needs no mapping; `options.namespaces` adds to or overrides them. The engine is the caller's
   * (`pkg.xpathEngine`: XPath 1.0 by default, 3.1 with `FontoXPathEngine`), and is readied here on
   * the snapshot itself: whether an engine can evaluate a tree is a property of the document that
   * tree is in (CR-002 section 21), and readying it on a marshal of its own would cost a second one
   * (CR-002 section 31).
   */
  async selectObjects(xpath: string, options: SelectOptions = {}): Promise<Bound[]> {
    const element = this.element ?? (await this.getContents(), this.element);
    if (element === undefined) throw new Docx4JException(`Part ${this.partName} has no contents to select over`);
    const engine = options.engine;
    if (engine === undefined) throw new Docx4JException(`selectObjects on ${this.partName} needs an XPath engine; pkg.selectObjects passes the package's`);
    const snapshot = await bindSnapshot(element);
    await engine.ready(snapshot.root);
    const namespaces = { ...NAMESPACE_PREFIXES_BY_PREFIX, ...options.namespaces };
    return boundsOf(engine.select(xpath, snapshot.root, namespaces), snapshot);
  }

  /**
   * The root element as a DOM element (docx4j XmlUtils.marshaltoW3CDomDocument). Run properties are written in
   * their normal form unless the package's `normaliseOnSave` (else the objects package's `setNormaliseOnSave`) is
   * false (CR-015 D2, the objects facade's CR-013 phase C): the tree keeps what it read.
   */
  async marshalToNode(options?: PartBytesOptions): Promise<Element> {
    if (this.element === undefined) await this.getContents();
    const marshal: MarshalOptions = {};
    const normalise = this.package?.normaliseOnSave;
    if (normalise !== undefined) marshal.normaliseOnSave = normalise;
    if (options?.onNormalised !== undefined) marshal.onNormalised = options.onNormalised;
    return marshalNode(this.element as Jsonix.TypedNamedValue, marshal);
  }

  /**
   * Parses the part's bytes, and names the part when the XML is not well-formed. xmldom's error
   * (Node) carries the line and column of the fault; a browser's says them in its text, which
   * `parseXml` passes on.
   */
  private parse(bytes: Uint8Array): Document {
    try {
      return parseXml(decodeXmlText(bytes));
    } catch (cause) {
      const at = (cause as { locator?: { lineNumber?: unknown; columnNumber?: unknown } } | undefined)?.locator;
      const line = typeof at?.lineNumber === 'number' ? at.lineNumber : undefined;
      const column = typeof at?.columnNumber === 'number' ? at.columnNumber : undefined;
      const reason = (cause instanceof Error ? cause.message : String(cause)).replace(/^Malformed XML: /, '');
      throw new PartUnmarshalException(`${this.partName}: Malformed XML${positionOf(line, column)}: ${reason}`, this.partName.name, { cause, line, column });
    }
  }

  /**
   * Unmarshals, and names the part when the runtime refuses. The runtime's own message says what is
   * wrong and, since `@docx4j/jsonix` 3.4.0, where - `w:sz/@w:val (line 12): Argument [NaN] must be
   * an integer` - but not which part it was reading, which is the first thing a person editing a
   * part's XML needs (CR-002 section 27), nor the value it refused. The message names the part, and
   * for an attribute gives the value as written and the column too:
   * `/word/styles.xml: w:sz/@w:val="big" (line 12, column 9): Argument [NaN] must be an integer, ...`.
   * The location it found is lifted onto the exception so a caller does not parse the message; `line`
   * is undefined where the parser supplies no position, which is every browser (only xmldom carries a
   * locator).
   */
  private async unmarshal(doc: Document): Promise<Jsonix.TypedNamedValue<T>> {
    const kept: ValueKeptWarning[] = [];
    try {
      const element = await unmarshalNode<Jsonix.TypedNamedValue<T>>(doc, {
        onWordReading: (reading) => kept.push(valueKept(reading, this.partName.name)),
      });
      if (kept.length > 0) this.parseWarnings = [...this.parseWarnings, ...kept];
      return element;
    } catch (cause) {
      type LocatedNode = { lineNumber?: number; columnNumber?: number; getAttributeNS?(ns: string | null, local: string): string | null };
      const located = cause as { message?: unknown; node?: LocatedNode | undefined; attributeName?: { namespaceURI?: string; localPart?: string } | undefined; jsonixLocated?: boolean };
      const message = typeof located?.message === 'string' ? located.message : String(cause);
      const location = locationOf(message, located?.jsonixLocated === true);
      const line = typeof located?.node?.lineNumber === 'number' ? located.node.lineNumber : undefined;
      const column = typeof located?.node?.columnNumber === 'number' ? located.node.columnNumber : undefined;
      let detail = message;
      if (location !== undefined) {
        const reason = message.slice(LOCATION_PREFIX.exec(message)?.[0].length ?? 0);
        const name = located?.attributeName;
        const value = name?.localPart ? located?.node?.getAttributeNS?.(name.namespaceURI || null, name.localPart) : undefined;
        const shown = typeof value === 'string' ? `=${JSON.stringify(value.length > 80 ? `${value.slice(0, 80)}...` : value)}` : '';
        detail = `${location}${shown}${positionOf(line, column)}: ${reason}`;
      }
      throw new PartUnmarshalException(`${this.partName}: ${detail}`, this.partName.name, { cause, location, line, column });
    }
  }
}
