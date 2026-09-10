import { PartName } from './PartName.mjs';
import { ContentTypes } from './ContentTypes.mjs';
import { InvalidFormatException } from './exceptions.mjs';
import { XML_DECLARATION, parseXml, childElements } from '../xml/dom.mjs';

export const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

/**
 * `[Content_Types].xml` (docx4j org.docx4j.openpackaging.contenttype.ContentTypeManager):
 * `Default` maps an extension to a content type, `Override` a part name. Lookups are
 * case-insensitive; the names are written back as given.
 */
export class ContentTypeManager {
  /** Extension (lower case) to content type. */
  private readonly defaults = new Map<string, string>();
  /** Lower-cased part name to [part name as written, content type]. */
  private readonly overrides = new Map<string, [string, string]>();

  /** A manager with the two defaults every package needs: rels and xml. */
  static createDefault(): ContentTypeManager {
    const ctm = new ContentTypeManager();
    ctm.addDefaultContentType('rels', ContentTypes.RELATIONSHIPS_PART);
    ctm.addDefaultContentType('xml', ContentTypes.APPLICATION_XML);
    return ctm;
  }

  /** Parses `[Content_Types].xml`. */
  static parse(xml: string): ContentTypeManager {
    const ctm = new ContentTypeManager();
    const doc = parseXml(xml);
    const root = doc.documentElement;
    if (root.localName !== 'Types') throw new InvalidFormatException(`[Content_Types].xml: expected Types, found ${root.nodeName}`);
    for (const el of childElements(root)) {
      if (el.localName === 'Default') {
        const ext = el.getAttribute('Extension');
        const ct = el.getAttribute('ContentType');
        if (ext !== null && ct !== null) ctm.addDefaultContentType(ext, ct);
      } else if (el.localName === 'Override') {
        const name = el.getAttribute('PartName');
        const ct = el.getAttribute('ContentType');
        if (name !== null && ct !== null) ctm.addOverrideContentType(name, ct);
      }
    }
    return ctm;
  }

  /** The XML of `[Content_Types].xml`, defaults first as Word writes it. */
  toXml(): string {
    let out = `${XML_DECLARATION}<Types xmlns="${CONTENT_TYPES_NS}">`;
    for (const [ext, ct] of this.defaults) out += `<Default Extension="${escapeAttr(ext)}" ContentType="${escapeAttr(ct)}"/>`;
    for (const [, [name, ct]] of this.overrides) out += `<Override PartName="${escapeAttr(name)}" ContentType="${escapeAttr(ct)}"/>`;
    return out + '</Types>';
  }

  /** The content type of a part: its override if any, else the default for its extension. */
  getContentType(partName: PartName | string): string | undefined {
    const pn = PartName.of(partName);
    const override = this.overrides.get(pn.key);
    if (override) return override[1];
    return this.defaults.get(pn.extension.toLowerCase());
  }

  addDefaultContentType(extension: string, contentType: string): void {
    this.defaults.set(extension.toLowerCase(), contentType);
  }

  addOverrideContentType(partName: PartName | string, contentType: string): void {
    const pn = PartName.of(partName);
    this.overrides.set(pn.key, [pn.name, contentType]);
  }

  /** Registers a part's content type: as a default when one already covers the extension, else as an override (docx4j RelationshipsPart.addPart). */
  addContentType(partName: PartName | string, contentType: string): void {
    const pn = PartName.of(partName);
    if (this.defaults.get(pn.extension.toLowerCase()) === contentType) return;
    this.addOverrideContentType(pn, contentType);
  }

  /** Removes the override for a part; when there is none, the default for its extension (docx4j semantics). */
  removeContentType(partName: PartName | string): void {
    const pn = PartName.of(partName);
    if (this.overrides.delete(pn.key)) return;
    this.defaults.delete(pn.extension.toLowerCase());
  }

  removeOverrideContentType(partName: PartName | string): void {
    this.overrides.delete(PartName.of(partName).key);
  }

  removeDefaultContentType(extension: string): void {
    this.defaults.delete(extension.toLowerCase());
  }

  getDefaultContentType(extension: string): string | undefined {
    return this.defaults.get(extension.toLowerCase());
  }

  getOverrideContentType(partName: PartName | string): string | undefined {
    return this.overrides.get(PartName.of(partName).key)?.[1];
  }

  /** Whether any default or override names this content type. */
  isContentTypeRegistered(contentType: string): boolean {
    for (const ct of this.defaults.values()) if (ct === contentType) return true;
    for (const [, ct] of this.overrides.values()) if (ct === contentType) return true;
    return false;
  }

  /** Part names overridden to a content type (docx4j getPartNameOverridenByContentType, all of them). */
  getPartNamesForContentType(contentType: string): PartName[] {
    const out: PartName[] = [];
    for (const [name, ct] of this.overrides.values()) if (ct === contentType) out.push(PartName.of(name));
    return out;
  }

  get defaultEntries(): ReadonlyMap<string, string> {
    return this.defaults;
  }

  get overrideEntries(): [PartName, string][] {
    return [...this.overrides.values()].map(([name, ct]) => [PartName.of(name), ct]);
  }
}

function escapeAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
}
