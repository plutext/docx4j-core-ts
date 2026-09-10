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

  constructor(partName: PartName | string, contentType: string = ContentTypes.APPLICATION_XML, relationshipType: string = '') {
    super(partName, contentType, relationshipType);
  }

  get isLoaded(): boolean {
    return this.document !== undefined || this.bytes !== undefined;
  }

  get isParsed(): boolean {
    return this.document !== undefined;
  }

  async getDocument(): Promise<Document> {
    if (this.document === undefined) {
      const bytes = this.bytes ?? (await this.loadSourceBytes());
      if (bytes === undefined) throw new Docx4JException(`Part ${this.partName} has no content: set a document or load it from a container`);
      this.document = parseXml(decodeXmlText(bytes));
      this.bytes = undefined;
    }
    return this.document;
  }

  setDocument(document: Document): void {
    this.document = document;
    this.bytes = undefined;
  }

  setXml(xml: string): void {
    this.bytes = encodeText(xml);
    this.document = undefined;
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
