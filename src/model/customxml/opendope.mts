// SPDX-License-Identifier: Apache-2.0
//
// The OpenDoPE conventions the bind step reads (CR-005 section 8.2): a control's `w:tag` as the
// specification's parameter string (`od:xpath=x1&od:ContentType=application/xhtml+xml`), and the
// XPaths part (`http://opendope.org/xpaths`), whose entries name a data part, an XPath and its
// prefix mappings, as a `w:dataBinding` would. Phase B (section 3) will read the rest of the parts.
import type { CustomXmlPartLookup } from './XmlMapping.mjs';
import type { CustomXmlNode } from './CustomXmlPart.mjs';

export const OPENDOPE_XPATHS_NS = 'http://opendope.org/xpaths';
export const XHTML_CONTENT_TYPE = 'application/xhtml+xml';
/** The tag's `od:Handler` value for a picture (the specification's section 9.3; CR-005 section 9). */
export const PICTURE_HANDLER = 'picture';

/** The tag's parameters, `key=value` pairs separated by `&`, in order; a value is as written (the content type keeps its `/` and `+`). */
export function tagParamsOf(tag: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const pair of tag.split('&')) {
    const eq = pair.indexOf('=');
    if (eq <= 0) continue;
    const key = pair.slice(0, eq).trim();
    if (!out.has(key)) out.set(key, pair.slice(eq + 1).trim());
  }
  return out;
}

/**
 * An XPaths part entry: `<xpath id="x1"><dataBinding storeItemID xpath prefixMappings/></xpath>`.
 * The attributes are unqualified (`xpaths.xsd`: `attributeFormDefault="unqualified"`; docx4j's
 * `org.opendope.xpaths.Xpaths` and its invoice.docx write them so); a `w:`-qualified form, which
 * this package's own check 37 and 38 files were made with, is read as well.
 */
export interface XPathsEntry {
  id: string;
  storeItemID: string;
  xpath: string;
  prefixMappings: string;
}

/** The XPaths part's entries by id, from the first custom XML part in the OpenDoPE XPaths namespace; empty without one. */
export function xpathsEntriesOf(parts: CustomXmlPartLookup): Map<string, XPathsEntry> {
  const out = new Map<string, XPathsEntry>();
  const part = parts.items.find((p) => p.namespaceUri === OPENDOPE_XPATHS_NS);
  if (!part) return out;
  const root = part.document.documentElement;
  if (!root) return out;
  const entries = root.getElementsByTagNameNS(OPENDOPE_XPATHS_NS, 'xpath');
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i]!;
    const id = entry.getAttribute('id');
    const binding = entry.getElementsByTagNameNS(OPENDOPE_XPATHS_NS, 'dataBinding')[0];
    if (!id || !binding) continue;
    const read = (name: string): string => binding.getAttribute(name) || binding.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', name) || binding.getAttribute(`w:${name}`) || '';
    out.set(id, { id, storeItemID: read('storeItemID'), xpath: read('xpath'), prefixMappings: read('prefixMappings') });
  }
  return out;
}

/**
 * The node an `od:xpath` entry selects: the entry's XPath over the part its `storeItemID` names
 * (the id compared without braces or case, as `getItem` does). What a tag-bound control's pass
 * starts from; the note says why there is no node.
 */
export function selectEntryNode(entryId: string, parts: CustomXmlPartLookup): { node: CustomXmlNode } | { note: string } {
  const entry = xpathsEntriesOf(parts).get(entryId);
  if (!entry) return { note: `od:xpath=${entryId}: no such entry in the XPaths part` };
  const part = parts.getItem(entry.storeItemID) ?? parts.items.find((p) => p.id.replace(/[{}]/g, '').toLowerCase() === entry.storeItemID.replace(/[{}]/g, '').toLowerCase());
  if (!part) return { note: `od:xpath=${entryId}: no custom XML part ${entry.storeItemID}` };
  const node = part.selectSingleNode(entry.xpath, entry.prefixMappings);
  if (!node) return { note: `od:xpath=${entryId}: ${entry.xpath} selects nothing` };
  return { node };
}
