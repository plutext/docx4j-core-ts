import { Part } from './Part.mjs';
import { PartName } from '../opc/PartName.mjs';
import { ContentTypes } from '../opc/ContentTypes.mjs';
import { Namespaces } from './Namespaces.mjs';
import { Docx4JException } from '../opc/exceptions.mjs';
import { parseXml, serializeXmlPart, encodeText, decodeXmlText } from '../xml/dom.mjs';

/**
 * An XML part the object model does not type, held as a DOM (docx4j DefaultXmlPart / XmlPart).
 * The document is parsed on first `getDocument()`; a part whose document was never requested is
 * written back from its source bytes.
 */
export class DefaultXmlPart extends Part {
  private document: Document | undefined;
  private bytes: Uint8Array | undefined;
  /**
   * A DOM parsed for reading only (CR-002 phase E): the part still writes its source bytes on
   * save, so reading a custom XML part does not cost its byte-for-byte round trip. `markModified`
   * adopts it.
   */
  private readDocument: Document | undefined;

  constructor(partName: PartName | string, contentType: string = ContentTypes.APPLICATION_XML, relationshipType: string = '') {
    super(partName, contentType, relationshipType);
  }

  get isLoaded(): boolean {
    return this.document !== undefined || this.bytes !== undefined;
  }

  /** Whether a DOM is parsed (for reading or as the part's content). */
  get isParsed(): boolean {
    return this.parsedDocument !== undefined;
  }

  /** The DOM as the part's own content: from here on it is re-marshalled on save (docx4j getDocument). */
  async getDocument(): Promise<Document> {
    const document = await this.parseDocument();
    this.markModified();
    return document;
  }

  /**
   * The DOM, parsed on first call, **without** adopting it: an untouched part still writes its
   * source bytes (CR-001's rule). Call `markModified()` after editing the DOM in place. Extension
   * of docx4j's DefaultXmlPart, for the custom XML views of CR-002 phase E.
   */
  async parseDocument(): Promise<Document> {
    if (this.document !== undefined) return this.document;
    if (this.readDocument === undefined) {
      const bytes = this.bytes ?? (await this.loadSourceBytes());
      if (bytes === undefined) throw new Docx4JException(`Part ${this.partName} has no content: set a document or load it from a container`);
      this.readDocument = parseXml(decodeXmlText(bytes));
    }
    return this.readDocument;
  }

  /** The DOM if it is already parsed, else undefined: what a synchronous view reads (extension). */
  get parsedDocument(): Document | undefined {
    return this.document ?? this.readDocument;
  }

  /** Adopts the parsed DOM as the part's content, so that save re-marshals it (extension). */
  markModified(): void {
    if (this.document !== undefined) return;
    if (this.readDocument === undefined) throw new Docx4JException(`Part ${this.partName} has no parsed document to mark as modified`);
    this.document = this.readDocument;
    this.bytes = undefined;
  }

  setDocument(document: Document): void {
    this.document = document;
    this.readDocument = undefined;
    this.bytes = undefined;
  }

  setXml(xml: string): void {
    this.bytes = encodeText(xml);
    this.document = undefined;
    this.readDocument = undefined;
  }

  async getXml(): Promise<string> {
    return decodeXmlText(await this.getBytes());
  }

  async getBytes(): Promise<Uint8Array> {
    if (this.document !== undefined) return encodeText(serializeXmlPart(this.document));
    if (this.bytes !== undefined) return this.bytes;
    const loaded = await this.loadSourceBytes();
    if (loaded === undefined) throw new Docx4JException(`Part ${this.partName} has no content and no source container`);
    return loaded;
  }
}

/**
 * A custom XML data storage part (docx4j CustomXmlDataStoragePart): the customer's own XML,
 * held as a DOM. Its `itemId` comes from the sibling properties part and is the key in
 * `OpcPackage.customXmlDataStorageParts`.
 */
export class CustomXmlDataStoragePart extends DefaultXmlPart {
  /** The `ds:itemID` of the sibling properties part, lower-cased, once registered. */
  itemId: string | undefined;

  constructor(partName: PartName | string, contentType: string = ContentTypes.OFFICEDOCUMENT_CUSTOMXML_DATASTORAGE) {
    super(partName, contentType, Namespaces.CUSTOM_XML_DATA_STORAGE);
  }
}
