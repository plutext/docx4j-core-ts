// CR-002 phase C: insertOoxml. Word's `insertOoxml` takes a flat OPC `pkg:package` string; so
// does this one. The incoming package's body content is taken, the parts that content references
// through relationships (images, embedded objects, charts) are copied into the target package
// under fresh part names and relationship ids, the references in the content are rewritten, and
// the content is inserted. Styles and numbering are not merged (CR-002 open question 5: that is
// what docx4j's MergeDocx does). A bare `w:p` / `w:tbl` fragment is accepted too, which is what
// `insertXml` takes; the two share this module.
import { wml as parseFragment, walkAll } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { PartName } from '../../opc/PartName.mjs';
import { ContentTypes } from '../../opc/ContentTypes.mjs';
import { BinaryPart, ImagePart } from '../../parts/BinaryPart.mjs';
import { isExternal } from '../../parts/RelationshipsPart.mjs';
import type { RelationshipsPart } from '../../parts/RelationshipsPart.mjs';
import type { Part } from '../../parts/Part.mjs';
import type { Element } from './tree.mjs';

/** The flat OPC namespace; its presence is how a `pkg:package` is told from a fragment. */
const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';

/** True for a flat OPC package (what Word's `insertOoxml` is given), as opposed to a bare fragment. */
export function isFlatOpc(ooxml: string): boolean {
  return ooxml.includes(PKG_NS);
}

/**
 * Attributes that carry a relationship id. A value is rewritten only when the source package
 * really has a relationship of that id, so numeric ids (w:bookmarkStart/@w:id, wp:docPr/@id)
 * are never touched.
 */
const REL_ATTRIBUTES = ['embed', 'link', 'id', 'href', 'pict', 'dm', 'lo', 'qs', 'cs', 'topLeft', 'topRight', 'bottomLeft', 'bottomRight'];

/** The relationships namespace, for the same attributes on DOM nodes kept as `any` content. */
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export interface OoxmlOptions {
  /** The DOM preprocessor (the package's MCE one), for both routes. */
  preprocess?: (doc: Document) => void;
  /** The part new relationships are added to: the main document part, or the header or footer. */
  target?: Part;
}

/**
 * The block-level content of `ooxml`, ready to insert: a flat OPC package's body content with
 * its parts copied and its references rewritten, or a fragment's elements.
 */
export async function contentOf(ooxml: string, options: OoxmlOptions = {}): Promise<Element[]> {
  if (!isFlatOpc(ooxml)) {
    return parseFragment(ooxml, { wrapper: 'body', preprocess: options.preprocess });
  }
  const target = options.target;
  if (!target) throw new Docx4JException('A pkg:package can only be inserted into a body that has a part');
  // Imported here so that the content API does not import the packages statically: parts/wml
  // builds Body views, and a static import back would be a cycle (CLAUDE.md).
  const { WordprocessingMLPackage } = await import('../../packages/WordprocessingMLPackage.mjs');
  const source = await WordprocessingMLPackage.load(ooxml, options.preprocess ? { preprocessor: options.preprocess } : {});
  const main = source.getMainDocumentPart();
  const document = await main.getContents();
  const content = (document.body?.content ?? []) as Element[];
  if (content.length === 0) return [];
  const rewrites = await copyReferencedParts(content, main, target);
  if (rewrites.size > 0) rewriteRelationshipIds(content, rewrites);
  return content;
}

/**
 * Copies every part the content references into the target's package and returns the mapping
 * from the incoming relationship ids to the new ones.
 */
async function copyReferencedParts(content: Element[], sourcePart: Part, target: Part): Promise<Map<string, string>> {
  const sourceRels = sourcePart.relationshipsPart;
  const map = new Map<string, string>();
  if (!sourceRels) return map;
  const referenced = new Set<string>();
  collectRelationshipIds(content, sourceRels, referenced);
  const copied = new Map<string, Part>();
  for (const id of referenced) {
    const rel = sourceRels.getRelationshipById(id);
    if (!rel) continue;
    if (isExternal(rel)) {
      map.set(id, target.getRelationshipsPart(true)!.addExternalRelationship(rel.type, rel.target).id);
      continue;
    }
    const part = sourceRels.getPart(rel);
    if (!part) continue;
    const newId = await copyPart(part, target, copied);
    if (newId) map.set(id, newId);
  }
  return map;
}

/** Copies a part (and, recursively, everything its own relationships target) as a target of `owner`. */
async function copyPart(source: Part, owner: Part, copied: Map<string, Part>): Promise<string | undefined> {
  const pkg = owner.package;
  if (!pkg) throw new Docx4JException('The target part is not in a package');
  const already = copied.get(source.partName.key);
  if (already) {
    const rel = owner.getRelationshipsPart(true)!.getRel(already.partName) ?? owner.addTargetPart(already, 'REUSE_EXISTING');
    return rel.id;
  }
  const bytes = await source.getBytes();
  const name = freeName(pkg, source);
  const copy = source.contentType.startsWith('image/')
    ? new ImagePart(name, source.contentType)
    : new BinaryPart(name, source.contentType, source.relationshipType);
  copy.setBytes(bytes);
  copied.set(source.partName.key, copy);
  const rel = owner.addTargetPart(copy);
  const childRels = source.relationshipsPart;
  if (childRels && childRels.size > 0) {
    const copyRels = copy.getRelationshipsPart(true)!;
    for (const child of childRels.list) {
      if (isExternal(child)) { copyRels.addExternalRelationship(child.type, child.target, child.id); continue; }
      const childPart = childRels.getPart(child);
      if (!childPart) continue;
      // The child keeps its id, so nothing inside the copied part has to be rewritten.
      await copyChild(childPart, copy, copyRels, child.id, copied);
    }
  }
  return rel.id;
}

async function copyChild(source: Part, owner: Part, rels: RelationshipsPart, relId: string, copied: Map<string, Part>): Promise<void> {
  const already = copied.get(source.partName.key);
  if (already) { rels.addPart(already, 'REUSE_EXISTING', relId); return; }
  const pkg = owner.package!;
  const bytes = await source.getBytes();
  const copy = source.contentType.startsWith('image/')
    ? new ImagePart(freeName(pkg, source), source.contentType)
    : new BinaryPart(freeName(pkg, source), source.contentType, source.relationshipType);
  copy.setBytes(bytes);
  copied.set(source.partName.key, copy);
  rels.addPart(copy, 'OVERWRITE_IF_NAME_EXISTS', relId);
  const childRels = source.relationshipsPart;
  if (childRels && childRels.size > 0) {
    const copyRels = copy.getRelationshipsPart(true)!;
    for (const child of childRels.list) {
      if (isExternal(child)) { copyRels.addExternalRelationship(child.type, child.target, child.id); continue; }
      const childPart = childRels.getPart(child);
      if (childPart) await copyChild(childPart, copy, copyRels, child.id, copied);
    }
  }
}

/** A part name free in the package: `/word/media/imageN.ext` for images, else the source name with a suffix. */
function freeName(pkg: { getPart(name: PartName | string): unknown }, source: Part): PartName {
  if (source.contentType.startsWith('image/')) {
    const ext = source.partName.extension || extensionFor(source.contentType);
    for (let i = 1; ; i++) {
      const candidate = PartName.of(`/word/media/image${i}.${ext}`);
      if (!pkg.getPart(candidate)) return candidate;
    }
  }
  if (!pkg.getPart(source.partName)) return source.partName;
  const name = source.partName.name;
  const dot = name.lastIndexOf('.');
  const stem = dot > name.lastIndexOf('/') ? name.substring(0, dot) : name;
  const suffix = dot > name.lastIndexOf('/') ? name.substring(dot) : '';
  for (let i = 2; ; i++) {
    const candidate = PartName.of(`${stem}${i}${suffix}`);
    if (!pkg.getPart(candidate)) return candidate;
  }
}

function extensionFor(contentType: string): string {
  const found = Object.entries({ png: ContentTypes.IMAGE_PNG, jpeg: ContentTypes.IMAGE_JPEG, gif: ContentTypes.IMAGE_GIF, bmp: ContentTypes.IMAGE_BMP, tiff: ContentTypes.IMAGE_TIFF, svg: ContentTypes.IMAGE_SVG, emf: ContentTypes.IMAGE_EMF, wmf: ContentTypes.IMAGE_WMF })
    .find(([, ct]) => ct === contentType);
  return found ? found[0] : 'bin';
}

/** Every relationship id the content refers to (typed objects and the DOM nodes kept as `any`). */
function collectRelationshipIds(value: unknown, rels: RelationshipsPart, out: Set<string>): void {
  visitReferences(value, (id) => {
    if (rels.getRelationshipById(id)) out.add(id);
    return undefined;
  });
}

/** Rewrites the relationship ids the content refers to, in place. */
export function rewriteRelationshipIds(value: unknown, map: Map<string, string>): void {
  visitReferences(value, (id) => map.get(id));
}

/**
 * Offers every value that could be a relationship id to `rewrite`, in the typed tree and in the
 * DOM an `xs:any` property holds (the builders' `walkAll`, objects CR-003 section 3.6); a string
 * it returns replaces the value. Typed objects carry the id in one of `REL_ATTRIBUTES`, DOM
 * elements in an attribute of the relationships namespace.
 */
function visitReferences(value: unknown, rewrite: (id: string) => string | undefined): void {
  walkAll(value, (object) => {
    const record = object as Record<string, unknown>;
    for (const key of REL_ATTRIBUTES) {
      const item = record[key];
      if (typeof item !== 'string') continue;
      const next = rewrite(item);
      if (next !== undefined) record[key] = next;
    }
  }, (node) => {
    const attributes = node.attributes;
    for (let i = 0; i < attributes.length; i++) {
      const attr = attributes[i]!;
      if (attr.namespaceURI === R_NS && typeof attr.value === 'string') {
        const next = rewrite(attr.value);
        if (next !== undefined) attr.value = next;
      }
    }
  });
}
