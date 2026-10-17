/**
 * The relationships content references, and the removal of those nothing references any more
 * (CR-002 section 46, docx4j CR-043's D7): shared by the tracked-change resolution
 * (`TrackedChange.mts`) and the pictures (`InlinePicture.delete`, section 47).
 */
import { getContextSync } from '@docx4j/generated-objects-ts';
import type { Part } from '../../parts/Part.mjs';

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Per type, the properties that are attributes in the relationships namespace (`r:id`, `r:embed`, `r:link`, ...). */
const relationshipAttributes = new Map<string, readonly string[]>();

function relationshipAttributesOf(typeName: string): readonly string[] {
  let names = relationshipAttributes.get(typeName);
  if (names) return names;
  const found: string[] = [];
  type Property = { name?: string; attributeName?: { namespaceURI?: string } | null };
  type Info = { properties?: Property[]; baseTypeInfo?: Info };
  let info: Info | undefined;
  try { info = getContextSync().getTypeInfoByName(typeName) as unknown as Info | undefined; } catch { info = undefined; }
  for (let ti = info; ti; ti = ti.baseTypeInfo) {
    for (const property of ti.properties ?? []) {
      if (property.attributeName?.namespaceURI === R_NS && property.name) found.push(property.name);
    }
  }
  relationshipAttributes.set(typeName, names = found);
  return names;
}

/**
 * Adds to `into` every relationship id a tree references: the `r:` attributes of its typed objects (read
 * off the model's mapping, as docx4j reads its fields' `@XmlAttribute`), those kept in an
 * `otherAttributes` map, and those of any DOM it holds.
 */
export function relationshipIdsIn(tree: unknown, into: Set<string>): void {
  const seen = new Set<object>();
  const visitDom = (node: Node): void => {
    if (node.nodeType !== 1) return;
    const attributes = (node as unknown as { attributes: ArrayLike<Attr> }).attributes;
    for (let i = 0; i < attributes.length; i++) {
      const a = attributes[i]!;
      if (a.namespaceURI === R_NS && a.value) into.add(a.value);
    }
    for (let c = node.firstChild; c; c = c.nextSibling) visitDom(c);
  };
  const visit = (o: unknown): void => {
    if (typeof o !== 'object' || o === null || seen.has(o)) return;
    seen.add(o);
    if (Array.isArray(o)) { for (const item of o) visit(item); return; }
    if (typeof (o as { nodeType?: unknown }).nodeType === 'number') { visitDom(o as Node); return; }
    const record = o as Record<string, unknown>;
    const typeName = record.TYPE_NAME;
    if (typeof typeName === 'string') {
      for (const name of relationshipAttributesOf(typeName)) {
        const value = record[name];
        if (typeof value === 'string' && value) into.add(value);
      }
    }
    const other = record.otherAttributes;
    if (other && typeof other === 'object') {
      for (const [key, value] of Object.entries(other)) if (key.startsWith(`{${R_NS}}`) && typeof value === 'string') into.add(value);
    }
    for (const key in record) if (key !== 'PARENT' && key !== 'otherAttributes') visit(record[key]);
  };
  visit(tree);
}

/**
 * The relationships of `part` among `candidates` that nothing left in the part references are removed,
 * and the part each targeted when no relationship in the package targets it any more, with what that
 * part alone related to (`RelationshipsPart.removeParts`). A part not unmarshalled can't be read for
 * what still references, so nothing is removed from it.
 */
export function dropRelationshipsUnused(part: Part, candidates: Set<string>): void {
  const rels = part.relationshipsPart;
  const contents = (part as { isUnmarshalled?: boolean; contents?: unknown }).isUnmarshalled
    ? (part as { contents?: unknown }).contents : undefined;
  if (!rels || contents === undefined) return;
  const used = new Set<string>();
  relationshipIdsIn(contents, used);
  for (const id of candidates) {
    if (used.has(id)) continue;
    const rel = rels.getRelationshipById(id);
    if (!rel) continue;
    const target = rels.getPart(rel);
    rels.removeRelationship(rel);
    if (target && target.sourceRelationships.length === 0) {
      target.relationshipsPart?.removeParts();
      target.package?.parts.remove(target.partName);
    }
  }
}

/**
 * Relationship ids, by the name of the part they belong to, that content removed outside a
 * resolution referenced: a picture deleted with tracking off, or taken back from its author's own
 * insertion. They are settled when the package is saved rather than at once, so that content put
 * back in the meantime (a cut and a paste of the same drawing) keeps its relationship.
 */
export type RelationshipRemovals = Map<string, Set<string>>;

/** Queues `ids` of `part` for `flushPendingRelationshipRemovals`. */
export function queueRelationshipRemovals(pending: RelationshipRemovals | undefined, part: Part | undefined, ids: Iterable<string>): void {
  if (!pending || !part) return;
  let set = pending.get(part.partName.name);
  for (const id of ids) {
    if (!set) pending.set(part.partName.name, set = new Set());
    set.add(id);
  }
}

/** What `flushPendingRelationshipRemovals` needs of the package. */
interface PartsHost { getPart(name: string): Part | undefined }

/** Settles the queued ids against each part as it now stands (`dropRelationshipsUnused`); then `pending` is empty. */
export function flushPendingRelationshipRemovals(pkg: PartsHost, pending: RelationshipRemovals): void {
  for (const [name, ids] of pending) {
    const part = pkg.getPart(name);
    if (part) dropRelationshipsUnused(part, ids);
  }
  pending.clear();
}
