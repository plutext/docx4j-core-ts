import { marshalNode, type PackageElement } from '@docx4j/generated-objects-ts';
import type { PartStore, PartSink, PutOptions } from './PartStore.mjs';
import { ContentTypeManager } from './ContentTypeManager.mjs';
import { ContentTypes, isXmlContentType } from './ContentTypes.mjs';
import { CONTENT_TYPES_NAME } from './PartName.mjs';
import { InvalidFormatException } from './exceptions.mjs';
import { parseXml, serializeXml, childElements, encodeText, decodeXmlText, stripXmlDeclaration, base64Decode, base64Encode, XML_DECLARATION } from '../xml/dom.mjs';

export const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';

interface FlatPart {
  name: string;
  contentType: string;
  /** Element child of pkg:xmlData, or a typed element from the objects facade. */
  xml?: Element | { name: unknown; value: unknown };
  /** Base64 text of pkg:binaryData. */
  binary?: string;
}

/**
 * A flat OPC package (`pkg:package`, what Office JS `getOoxml()` returns) as a container
 * (docx4j org.docx4j.openpackaging.io3.stores.FlatOpcPartStore is not a class there, but the
 * same job is done by `FlatOpcXmlImporter`). The content types are carried per part, so
 * `[Content_Types].xml` is synthesised. XML parts are serialised back to bytes on `load`; the
 * document loaded from them is therefore not byte-identical to what Office would have zipped.
 */
export class FlatOpcPartStore implements PartStore {
  private readonly parts = new Map<string, FlatPart>();
  private readonly order: string[] = [];
  private readonly cache = new Map<string, Uint8Array>();
  private contentTypesXml?: Uint8Array;

  /** From the `pkg:package` string. */
  constructor(source: string | PackageElement) {
    if (typeof source === 'string') this.readString(source);
    else this.readPackageElement(source);
    this.order.unshift(CONTENT_TYPES_NAME);
  }

  private readString(text: string): void {
    const doc = parseXml(text);
    const root = doc.documentElement;
    if (root.namespaceURI !== PKG_NS || root.localName !== 'package') {
      throw new InvalidFormatException(`Not a flat OPC package: root is ${root.nodeName}`);
    }
    for (const part of childElements(root)) {
      if (part.namespaceURI !== PKG_NS || part.localName !== 'part') continue;
      const name = part.getAttributeNS(PKG_NS, 'name') || part.getAttribute('pkg:name');
      const contentType = part.getAttributeNS(PKG_NS, 'contentType') || part.getAttribute('pkg:contentType');
      if (!name || !contentType) throw new InvalidFormatException('pkg:part without pkg:name or pkg:contentType');
      const entry: FlatPart = { name: name.replace(/^\//, ''), contentType };
      for (const child of childElements(part)) {
        if (child.namespaceURI !== PKG_NS) continue;
        if (child.localName === 'xmlData') entry.xml = childElements(child)[0];
        else if (child.localName === 'binaryData') entry.binary = child.textContent ?? '';
      }
      this.add(entry);
    }
  }

  private readPackageElement(pkg: PackageElement): void {
    for (const part of pkg.value.part ?? []) {
      const name = (part.name ?? '').replace(/^\//, '');
      const contentType = part.contentType ?? '';
      const entry: FlatPart = { name, contentType };
      const any = part.xmlData?.any;
      if (any !== undefined && any !== null) entry.xml = any as FlatPart['xml'];
      if (part.binaryData !== undefined && part.binaryData !== null) entry.binary = String(part.binaryData);
      this.add(entry);
    }
  }

  private add(entry: FlatPart): void {
    this.parts.set(entry.name, entry);
    this.order.push(entry.name);
  }

  /** The content type of a part as carried by the flat package. */
  contentTypeOf(partName: string): string | undefined {
    return this.parts.get(partName)?.contentType;
  }

  partNames(): Iterable<string> {
    return this.order;
  }

  has(partName: string): boolean {
    return partName === CONTENT_TYPES_NAME || this.parts.has(partName);
  }

  async load(partName: string): Promise<Uint8Array> {
    if (partName === CONTENT_TYPES_NAME) return (this.contentTypesXml ??= encodeText(this.buildContentTypes().toXml()));
    const cached = this.cache.get(partName);
    if (cached) return cached;
    const part = this.parts.get(partName);
    if (!part) throw new Error(`No part ${partName} in flat OPC package`);
    let bytes: Uint8Array;
    if (part.binary !== undefined) {
      bytes = base64Decode(part.binary);
    } else if (part.xml !== undefined) {
      const el = isNode(part.xml) ? part.xml : await marshalNode(part.xml as Parameters<typeof marshalNode>[0]);
      bytes = encodeText(XML_DECLARATION + serializeXml(el));
    } else {
      bytes = new Uint8Array(0);
    }
    this.cache.set(partName, bytes);
    return bytes;
  }

  private buildContentTypes(): ContentTypeManager {
    const ctm = ContentTypeManager.createDefault();
    for (const part of this.parts.values()) {
      if (part.contentType === ContentTypes.RELATIONSHIPS_PART) continue;
      ctm.addContentType('/' + part.name, part.contentType);
    }
    return ctm;
  }
}

function isNode(value: unknown): value is Element {
  return typeof value === 'object' && value !== null && typeof (value as Node).nodeType === 'number';
}

export interface FlatOpcSinkOptions {
  /** The `mso-application` processing instruction's progid, e.g. 'Word.Document'; omitted when undefined. */
  progId?: string;
}

/**
 * Produces the `pkg:package` string for Office JS `insertOoxml()`. XML parts are embedded as
 * XML (their declaration removed), everything else as base64 with the padding Word writes.
 */
export class FlatOpcPartSink implements PartSink<string> {
  private readonly parts: string[] = [];
  private readonly progId: string | undefined;

  constructor(options?: FlatOpcSinkOptions) {
    this.progId = options?.progId;
  }

  put(partName: string, bytes: Uint8Array, options?: PutOptions): void {
    const name = partName.startsWith('/') ? partName : '/' + partName;
    if (name === '/' + CONTENT_TYPES_NAME) return; // content types travel on each part
    const contentType = options?.contentType ?? (name.endsWith('.rels') ? ContentTypes.RELATIONSHIPS_PART : ContentTypes.OCTET_STREAM);
    const xml = isXmlContentType(contentType) || name.endsWith('.xml') || name.endsWith('.rels');
    if (xml) {
      const text = stripXmlDeclaration(decodeXmlText(bytes));
      this.parts.push(`<pkg:part pkg:name="${escapeAttr(name)}" pkg:contentType="${escapeAttr(contentType)}"><pkg:xmlData>${text}</pkg:xmlData></pkg:part>`);
    } else {
      const b64 = base64Encode(bytes).replace(/.{76}/g, '$&\r\n');
      this.parts.push(`<pkg:part pkg:name="${escapeAttr(name)}" pkg:contentType="${escapeAttr(contentType)}" pkg:compression="store"><pkg:binaryData>${b64}</pkg:binaryData></pkg:part>`);
    }
  }

  async finish(): Promise<string> {
    const pi = this.progId ? `<?mso-application progid="${escapeAttr(this.progId)}"?>` : '';
    return `<?xml version="1.0" standalone="yes"?>${pi}<pkg:package xmlns:pkg="${PKG_NS}">${this.parts.join('')}</pkg:package>`;
  }
}

function escapeAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
}
