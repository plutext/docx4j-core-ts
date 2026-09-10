import type { Relationships, Relationship } from '@docx4j/generated-objects-ts/modules/org_docx4j_relationships';
import { XmlPart } from './XmlPart.mjs';
import { Part, setRelationshipsPartFactory, type RelationshipSource } from './Part.mjs';
import { PartName } from '../opc/PartName.mjs';
import { ContentTypes } from '../opc/ContentTypes.mjs';
import { Namespaces } from './Namespaces.mjs';
import { InvalidFormatException } from '../opc/exceptions.mjs';

/** What `addPart` does when a part of the same name is already in the package (docx4j AddPartBehaviour). */
export type AddPartBehaviour = 'OVERWRITE_IF_NAME_EXISTS' | 'REUSE_EXISTING' | 'RENAME_IF_NAME_EXISTS';

export const RELATIONSHIPS_ROOT = { namespaceURI: Namespaces.RELATIONSHIPS, localPart: 'Relationships' };
const RELATIONSHIP_TYPE_NAME = 'org_docx4j_relationships.Relationship';

/**
 * A relationships part (docx4j org.docx4j.openpackaging.parts.relationships.RelationshipsPart):
 * `/_rels/.rels` for the package, `/word/_rels/document.xml.rels` for a part. Always unmarshalled
 * (load walks it), hence always re-marshalled on save.
 */
export class RelationshipsPart extends XmlPart<Relationships> {
  /** The part (or package) whose relationships these are. */
  sourceP: RelationshipSource | undefined;
  private nextId = 1;

  constructor(source?: RelationshipSource, partName?: PartName | string) {
    super(partName ?? (source ? PartName.relsFor(source.partName) : '/_rels/.rels'), ContentTypes.RELATIONSHIPS_PART, '', RELATIONSHIPS_ROOT);
    this.sourceP = source;
    if (source) this.package = source.package;
  }

  /** The package relationships part, `/_rels/.rels`. */
  static createPackageRels(): RelationshipsPart {
    const rp = new RelationshipsPart(undefined, '/_rels/.rels');
    rp.setContents({ relationship: [] });
    return rp;
  }

  /** A part's own, empty, relationships part. */
  static createRelationshipsPartForPart(source: RelationshipSource): RelationshipsPart {
    const rp = new RelationshipsPart(source);
    rp.setContents({ relationship: [] });
    return rp;
  }

  /** The name of the part these relationships belong to; '/' for the package. */
  get sourceName(): PartName {
    return this.sourceP?.partName ?? PartName.ROOT;
  }

  get isPackageRelationshipPart(): boolean {
    return this.sourceName.name === '/';
  }

  /** The relationships (docx4j getRelationships()); unmarshalled parts only. */
  get relationships(): Relationships {
    const r = this.contents;
    r.relationship ??= [];
    return r;
  }

  /** The list of relationships, the live array. */
  get list(): Relationship[] {
    return this.relationships.relationship!;
  }

  get size(): number {
    return this.list.length;
  }

  getRelationshipById(id: string): Relationship | undefined {
    return this.list.find((r) => r.id === id);
  }

  getRelationshipByType(type: string): Relationship | undefined {
    return this.list.find((r) => r.type === type);
  }

  getRelationshipsByType(type: string): Relationship[] {
    return this.list.filter((r) => r.type === type);
  }

  /** The part a relationship targets; undefined for external targets and parts not in the package. */
  getPart(rel: Relationship | string): Part | undefined {
    const r = typeof rel === 'string' ? this.getRelationshipById(rel) : rel;
    if (!r || isExternal(r)) return undefined;
    return this.package?.getPart(this.resolveTarget(r));
  }

  /** The part name a relationship's target resolves to (internal targets only). */
  resolveTarget(rel: Relationship): PartName {
    return PartName.resolve(this.sourceName, rel.target);
  }

  /** The (first) relationship targeting a part. */
  getRel(partName: PartName | string): Relationship | undefined {
    const pn = PartName.of(partName);
    return this.list.find((r) => !isExternal(r) && this.resolveTarget(r).equals(pn));
  }

  isATarget(partName: PartName | string): boolean {
    return this.getRel(partName) !== undefined;
  }

  /** The next free `rIdN`. */
  getNextId(): string {
    const used = new Set(this.list.map((r) => r.id));
    let id: string;
    do {
      id = `rId${this.nextId++}`;
    } while (used.has(id));
    return id;
  }

  /** Adds a relationship; allocates an id when it has none or its id is taken. Returns the id. */
  addRelationship(rel: Relationship): string {
    if (!rel.id || this.getRelationshipById(rel.id)) rel.id = this.getNextId();
    if (!rel.TYPE_NAME) (rel as { TYPE_NAME?: string }).TYPE_NAME = RELATIONSHIP_TYPE_NAME;
    this.list.push(rel);
    return rel.id;
  }

  /** An external relationship (a hyperlink, a linked image). */
  addExternalRelationship(type: string, target: string, relId?: string): Relationship {
    const rel: Relationship = { TYPE_NAME: RELATIONSHIP_TYPE_NAME, id: relId ?? '', type, target, targetMode: 'External' };
    this.addRelationship(rel);
    return rel;
  }

  /**
   * Adds a part as a target (docx4j addPart): registers it in the package and its content type,
   * writes a relationship whose target is relative to the source part, and returns that
   * relationship. `mode` says what to do when the name is taken (default OVERWRITE_IF_NAME_EXISTS).
   */
  addPart(part: Part, mode: AddPartBehaviour = 'OVERWRITE_IF_NAME_EXISTS', relId?: string): Relationship {
    const pkg = this.package;
    if (!pkg) throw new InvalidFormatException('Relationships part has no package');
    let newPartName = part.partName;
    const existing = pkg.getPart(newPartName);
    if (existing) {
      if (mode === 'REUSE_EXISTING') part = existing;
      if (mode === 'RENAME_IF_NAME_EXISTS') {
        newPartName = this.newPartName(part.partName);
        part.partName = newPartName;
      }
    }
    const target = PartName.relativize(this.sourceName, part.partName);
    const existsAlready = this.list.find((r) => r.target === target);
    if (existsAlready && mode === 'REUSE_EXISTING') return existsAlready;
    if (existsAlready && mode === 'RENAME_IF_NAME_EXISTS') {
      throw new InvalidFormatException('Found existing rel, and yet constructed part name should be globally unique');
    }
    if (existsAlready && existing && mode === 'OVERWRITE_IF_NAME_EXISTS') {
      existsAlready.type = part.relationshipType;
      this.loadPart(part, existsAlready);
      return existsAlready;
    }
    const rel: Relationship = { TYPE_NAME: RELATIONSHIP_TYPE_NAME, id: relId ?? '', type: part.relationshipType, target };
    this.addRelationship(rel);
    pkg.contentTypeManager.addContentType(part.partName, part.contentType);
    this.loadPart(part, rel);
    return rel;
  }

  /** Links a part into the package as the target of a relationship of this part (docx4j loadPart). */
  loadPart(part: Part, sourceRelationship: Relationship): void {
    const pkg = this.package;
    if (!pkg) throw new InvalidFormatException('Relationships part has no package');
    part.owningRelationshipPart = this;
    part.sourceRelationships.push(sourceRelationship);
    part.package = pkg;
    pkg.parts.put(part);
  }

  /** Removes the relationship to a part, the part, and (recursively) the parts only it related to. Returns what was removed. */
  removePart(partName: PartName | string): PartName[] {
    const pn = PartName.of(partName);
    const removed: PartName[] = [];
    const pkg = this.package;
    const part = pkg?.getPart(pn);
    if (!part || !pkg) return removed;
    this.removeRelationship(pn);
    if (part.relationshipsPart) removed.push(...part.relationshipsPart.removeParts());
    pkg.parts.remove(pn);
    removed.push(pn);
    return removed;
  }

  /** Removes every part these relationships target (docx4j removeParts), recursively. */
  removeParts(): PartName[] {
    const removed: PartName[] = [];
    for (const rel of [...this.list]) {
      if (isExternal(rel)) continue;
      const pn = this.resolveTarget(rel);
      const part = this.package?.getPart(pn);
      // A part targeted by another relationship elsewhere stays; only its link from here goes.
      if (part && part.sourceRelationships.length <= 1) removed.push(...this.removePart(pn));
      else this.removeRelationship(rel);
    }
    return removed;
  }

  /** Removes a relationship, by object or by the part it targets. */
  removeRelationship(rel: Relationship | PartName | string): void {
    const r = typeof rel === 'object' && !(rel instanceof PartName) ? rel : this.getRel(rel as PartName | string);
    if (!r) return;
    const i = this.list.indexOf(r);
    if (i >= 0) this.list.splice(i, 1);
    const target = isExternal(r) ? undefined : this.package?.getPart(this.resolveTarget(r));
    if (target) {
      const j = target.sourceRelationships.indexOf(r);
      if (j >= 0) target.sourceRelationships.splice(j, 1);
    }
  }

  removeRelationshipsByType(type: string): void {
    for (const r of this.getRelationshipsByType(type)) this.removeRelationship(r);
  }

  private newPartName(proposed: PartName): PartName {
    const pkg = this.package!;
    const name = proposed.name;
    const dot = name.lastIndexOf('.');
    const prefix = dot > name.lastIndexOf('/') ? name.substring(0, dot) : name;
    const suffix = dot > name.lastIndexOf('/') ? name.substring(dot) : '';
    for (let i = 2; ; i++) {
      const candidate = PartName.of(`${prefix}${i}${suffix}`);
      if (!pkg.getPart(candidate)) return candidate;
    }
  }
}

export function isExternal(rel: Relationship): boolean {
  return rel.targetMode === 'External';
}

setRelationshipsPartFactory((source) => RelationshipsPart.createRelationshipsPartForPart(source));
