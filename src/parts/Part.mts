import type { Relationship } from '@docx4j/generated-objects-ts/modules/org_docx4j_relationships';
import { PartName } from '../opc/PartName.mjs';
import { InvalidFormatException } from '../opc/exceptions.mjs';
import type { RelationshipsPart, AddPartBehaviour } from './RelationshipsPart.mjs';
import type { OpcPackage } from '../packages/OpcPackage.mjs';

/** Something that owns a relationships part: a part or the package (docx4j Base). */
export interface RelationshipSource {
  readonly partName: PartName;
  readonly package: OpcPackage | undefined;
  relationshipsPart: RelationshipsPart | undefined;
  setPartShortcut(part: Part, relationshipType: string): boolean;
}

/** Set by the RelationshipsPart module: avoids a runtime import cycle (Part <- XmlPart <- RelationshipsPart). */
let relationshipsPartFactory: ((source: RelationshipSource) => RelationshipsPart) | undefined;
export function setRelationshipsPartFactory(factory: (source: RelationshipSource) => RelationshipsPart): void {
  relationshipsPartFactory = factory;
}
export function createRelationshipsPartFor(source: RelationshipSource): RelationshipsPart {
  if (!relationshipsPartFactory) throw new Error('RelationshipsPart not initialised');
  return relationshipsPartFactory(source);
}

/**
 * A part (docx4j org.docx4j.openpackaging.parts.Part): a name, a content type, the relationship
 * type it is the target of, and its own relationships part if it has one.
 */
export abstract class Part implements RelationshipSource {
  partName: PartName;
  contentType: string;
  relationshipType: string;
  /** The package this part belongs to; set when the part is added. */
  package: OpcPackage | undefined;
  /** This part's own relationships (`/word/_rels/document.xml.rels`), if any. */
  relationshipsPart: RelationshipsPart | undefined;
  /** The relationships part that first loaded this part. */
  owningRelationshipPart: RelationshipsPart | undefined;
  /** Every relationship that targets this part. */
  readonly sourceRelationships: Relationship[] = [];

  constructor(partName: PartName | string, contentType: string, relationshipType: string) {
    this.partName = PartName.of(partName);
    this.contentType = contentType;
    this.relationshipType = relationshipType;
  }

  /** Whether the zip entry should be deflated (false for already-compressed media). */
  get compress(): boolean {
    return true;
  }

  /** The first relationship that targets this part. */
  get sourceRelationship(): Relationship | undefined {
    return this.sourceRelationships[0];
  }

  /** This part's relationships part, created on demand when `createIfAbsent`. */
  getRelationshipsPart(createIfAbsent = false): RelationshipsPart | undefined {
    if (!this.relationshipsPart && createIfAbsent) this.relationshipsPart = createRelationshipsPartFor(this);
    return this.relationshipsPart;
  }

  /**
   * Adds a part as a target of this one (docx4j Base.addTargetPart): registers it in the package,
   * its content type in the content type manager, and a relationship in this part's relationships.
   */
  addTargetPart(target: Part, mode?: AddPartBehaviour, proposedRelId?: string): Relationship {
    if (!this.package) throw new InvalidFormatException('Package not set; if you are adding part2 to part1, make sure part1 is added first.');
    const rel = this.getRelationshipsPart(true)!.addPart(target, mode, proposedRelId);
    this.setPartShortcut(target, target.relationshipType);
    return rel;
  }

  /** Subclasses keep typed references to well-known targets (styles, numbering, ...); returns true when one was set. */
  setPartShortcut(_part: Part, _relationshipType: string): boolean {
    return false;
  }

  /** The bytes this part was loaded from, if it came from a container. Undefined for new parts. */
  protected async loadSourceBytes(): Promise<Uint8Array | undefined> {
    const store = this.package?.sourcePartStore;
    if (!store) return undefined;
    const name = this.partName.storeName;
    if (store.has(name)) return store.load(name);
    const decoded = decodeURIComponent(name);
    if (decoded !== name && store.has(decoded)) return store.load(decoded);
    return undefined;
  }

  /** The bytes to write on save. Untouched parts come back from the source container unchanged. */
  abstract getBytes(): Promise<Uint8Array>;

  /** Whether the part carries content of its own (set or unmarshalled) rather than relying on the source container. */
  abstract get isLoaded(): boolean;
}
