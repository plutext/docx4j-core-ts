// DOM, text and base64 helpers over the runtime's DOM implementation. In Node the runtime brings
// @xmldom/xmldom; in browsers it uses the native DOM. Jsonix.DOM is typed since @docx4j/jsonix 3.2.1.
import { Jsonix } from '@docx4j/generated-objects-ts';

const DOM = Jsonix.DOM;

/** The declaration docx4j writes on every XML part. */
export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Parses XML text to a DOM document; throws on malformed input. */
export function parseXml(text: string): Document {
  return DOM.parse(text);
}

/** Serialises a node (document or element) without an XML declaration. */
export function serializeXml(node: Node): string {
  return DOM.serialize(node);
}

export function createDocument(): Document {
  return DOM.createDocument();
}

/** XML declaration plus the serialised document element, as one string. */
export function serializeXmlPart(node: Node): string {
  const el = node.nodeType === 9 ? (node as Document).documentElement : node;
  return XML_DECLARATION + serializeXml(el);
}

const utf8Encoder = new TextEncoder();

/** UTF-8 bytes of a string. */
export function encodeText(text: string): Uint8Array {
  return utf8Encoder.encode(text);
}

/**
 * Text of an XML part's bytes: honours a UTF-8 or UTF-16 byte order mark, else assumes UTF-8
 * (the only encoding Office writes without a BOM). The BOM is not part of the result.
 */
export function decodeXmlText(bytes: Uint8Array): string {
  if (bytes.length >= 2) {
    if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
    if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  }
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return new TextDecoder('utf-8').decode(bytes.subarray(3));
  }
  return new TextDecoder('utf-8').decode(bytes);
}

/** Removes an XML declaration (and a leading BOM) from XML text. */
export function stripXmlDeclaration(text: string): string {
  return text.replace(/^\uFEFF?\s*<\?xml[^>]*\?>\s*/, '');
}

const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
const B64_LOOKUP = new Uint8Array(128);
for (let i = 0; i < B64.length; i++) B64_LOOKUP[B64.charCodeAt(i)] = i;

export function base64Encode(bytes: Uint8Array): string {
  let out = '';
  let i = 0;
  for (; i + 2 < bytes.length; i += 3) {
    const n = (bytes[i]! << 16) | (bytes[i + 1]! << 8) | bytes[i + 2]!;
    out += B64[n >> 18] + B64[(n >> 12) & 63] + B64[(n >> 6) & 63] + B64[n & 63];
  }
  if (i < bytes.length) {
    const n = (bytes[i]! << 16) | ((i + 1 < bytes.length ? bytes[i + 1]! : 0) << 8);
    out += B64[n >> 18] + B64[(n >> 12) & 63] + (i + 1 < bytes.length ? B64[(n >> 6) & 63] : '=') + '=';
  }
  return out;
}

export function base64Decode(text: string): Uint8Array {
  const clean = text.replace(/[^A-Za-z0-9+/]/g, '');
  const out = new Uint8Array(Math.floor((clean.length * 3) / 4));
  let o = 0;
  for (let i = 0; i < clean.length; i += 4) {
    const a = B64_LOOKUP[clean.charCodeAt(i)]!;
    const b = i + 1 < clean.length ? B64_LOOKUP[clean.charCodeAt(i + 1)]! : 0;
    const c = i + 2 < clean.length ? B64_LOOKUP[clean.charCodeAt(i + 2)]! : 0;
    const d = i + 3 < clean.length ? B64_LOOKUP[clean.charCodeAt(i + 3)]! : 0;
    const n = (a << 18) | (b << 12) | (c << 6) | d;
    if (o < out.length) out[o++] = (n >> 16) & 255;
    if (i + 2 < clean.length && o < out.length) out[o++] = (n >> 8) & 255;
    if (i + 3 < clean.length && o < out.length) out[o++] = n & 255;
  }
  return out.subarray(0, o);
}

/** Element children of a node, in order. */
export function childElements(node: Node): Element[] {
  const out: Element[] = [];
  for (let c = node.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) out.push(c as Element);
  return out;
}
