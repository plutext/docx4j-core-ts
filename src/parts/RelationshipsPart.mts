import { DoctypeException } from '../opc/exceptions.mjs';
import { doctypeWarning, presenceOf } from '../opc/doctypeParts.mjs';
import type { Relationships, Relationship } from '@docx4j/generated-objects-ts/modules/org_docx4j_relationships';
import { XmlPart } from './XmlPart.mjs';
import { Part, setRelationshipsPartFactory, type RelationshipSource } from './Part.mjs';
import { PartName } from '../opc/PartName.mjs';
import { ContentTypes } from '../opc/ContentTypes.mjs';
import { Namespaces } from './Namespaces.mjs';
import { InvalidFormatException, RelationshipsNamespaceException, type PartUnmarshalException, type RelationshipsNotReadWarning } from '../opc/exceptions.mjs';
import { relationshipsBeforeTheBreak } from '../opc/repair/relationshipsBeforeTheBreak.mjs';
import { parseXml } from '../xml/dom.mjs';
import type { PartWarning } from '../opc/mce/preprocessor.mjs';

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

  /**
   * Set by the load under `LoadOptions.repair` `'word'` or `'all'` (CR-031 section 2.3): a part in a wrong namespace is read
   * as Word's repair reads it, not refused.
   */
  readsByLocalName = false;

  /**
   * The parse error, where this part isn't well formed and was read short (CR-032); undefined otherwise. The load refuses
   * `/_rels/.rels` with it where the main part's relationship wasn't among those kept.
   */
  readFailure: PartUnmarshalException | undefined;
  private readShort: RelationshipsNotReadWarning | undefined;

  /**
   * A relationships part that isn't well formed, at every value of `repair` (CR-032; WD 0.76, docx4j CR-075 D5 and D6): the
   * relationships before the break, each whole `Relationship` by its local name with its `Id`, `Type` and `Target`, read as a
   * part in the right namespace, as Word's repair reads it. `readRoot` then says so, `part.not_read` with the count kept.
   */
  protected override readMalformed(text: string, failure: PartUnmarshalException): Document {
    const { kept, xml } = relationshipsBeforeTheBreak(text, parseXml, failure);
    const partName = this.partName.name;
    const reason = failure.message.startsWith(`${partName}: `) ? failure.message.slice(partName.length + 2) : failure.message;
    this.readFailure = failure;
    this.readShort = {
      code: 'part.not_read',
      partName,
      count: kept.length,
      reason,
      line: failure.line,
      column: failure.column,
      message: `${partName} isn't well formed (${reason}); the ${kept.length} relationship${kept.length === 1 ? '' : 's'} before the break read, `
        + "as Word's repair reads them, and a part only a lost one named is absent (CR-032). A save would drop what the others named: "
        + "it is refused unless the load's repair is 'word' or 'all', or checkLossOnSave is false",
    };
    return parseXml(xml);
  }

  /** Part 2 §6.2.5 b) forbids a DOCTYPE in a relationships part (CR-033). */
  protected override get forbidsDoctype(): boolean {
    return true;
  }

  /**
   * A relationships part holding a DOCTYPE (CR-033; the DTD handling policy's section 6.1): refused, naming the part. Under
   * `repair: 'word'` or `'all'` a part's relationships part is read as Word 365 repairs it, every relationship lost
   * (dtd-18-document-rels-bare): `part.not_read` with none kept, so a part only it named is absent and the save is CR-032's, and
   * `xml.doctype` with `repaired`. `/_rels/.rels` is refused at every value, as Word refuses it after its repair prompt.
   */
  protected override readForbiddenDoctype(_text: string): Document {
    const partName = this.partName.name;
    if (!this.readsByLocalName) throw new DoctypeException(partName, "the package is refused, as Word 365 refuses it unless its repair is followed (repair: 'word')");
    if (this.sourceP === undefined || partName === '/_rels/.rels') throw new DoctypeException(partName, 'the package is refused at every value of repair, as Word 365 refuses it after its repair prompt');
    const reason = "a DOCTYPE, which ECMA-376 Part 2 §6.2.5 forbids in a relationships part";
    this.readShort = {
      code: 'part.not_read',
      partName,
      count: 0,
      reason,
      line: undefined,
      column: undefined,
      message: `${partName} holds ${reason}; read as Word 365 repairs it, every relationship lost, and a part only it named is absent (CR-033)`,
    };
    this.doctypeRepaired = doctypeWarning(partName, this.contentType, presenceOf(), { repaired: true });
    return parseXml(`<Relationships xmlns="${Namespaces.RELATIONSHIPS}"/>`);
  }

  /** Where this part wasn't well formed and was read short, the warning saying so (CR-032); a part given contents since has none. */
  get notRead(): RelationshipsNotReadWarning | undefined {
    for (const w of this.warnings) if (w.code === 'part.not_read') return w;
    return undefined;
  }

  /** The read short, said first (CR-032), then the namespace (CR-025 section 8, CR-031 section 2.3). */
  protected override readRoot(doc: Document): readonly PartWarning[] {
    const short = this.readShort;
    this.readShort = undefined;
    if (!short) this.readFailure = undefined;
    const namespace = this.readNamespace(doc);
    return short ? [short, ...namespace] : namespace;
  }

  /**
   * Refuses a part in another namespace, or in none, on its root or on any `Relationship` (CR-025 section 8, D10), before the
   * runtime drops what it can't bind. Under `readsByLocalName` (`repair: 'word'`, CR-031 section 2.3) the root and every
   * `Relationship` are read by their local names instead, as Word's repair reads them, so the runtime binds them all and a
   * save writes the right namespace, with one `package.relationships_namespace` warning. A root or a child with another
   * local name is left to the runtime, as before: Word is unmeasured on both, and a child is then a `part.skipped` on this
   * part, which `pkg.warnings` gathers (D12).
   */
  private readNamespace(doc: Document): readonly PartWarning[] {
    const root = doc.documentElement;
    if (!root || root.localName !== 'Relationships') return [];
    const isRelationship = (c: Node): c is Element => c.nodeType === 1 && (c as Element).localName === 'Relationship';
    let wrong: string | null | undefined = root.namespaceURI !== Namespaces.RELATIONSHIPS ? root.namespaceURI ?? null : undefined;
    for (let c = root.firstChild; c && wrong === undefined; c = c.nextSibling) {
      if (isRelationship(c) && c.namespaceURI !== Namespaces.RELATIONSHIPS) wrong = c.namespaceURI ?? null;
    }
    if (wrong === undefined) return [];
    if (!this.readsByLocalName) throw new RelationshipsNamespaceException(this.partName.name, wrong);
    const fixed = inRelationshipsNamespace(doc, root);
    let count = 0;
    while (root.firstChild) {
      const c = root.firstChild;
      root.removeChild(c);
      if (isRelationship(c)) count++;
      fixed.appendChild(isRelationship(c) && c.namespaceURI !== Namespaces.RELATIONSHIPS ? inRelationshipsNamespace(doc, c) : c);
    }
    doc.replaceChild(fixed, root);
    return [{
      code: 'package.relationships_namespace',
      partName: this.partName.name,
      namespace: wrong,
      count,
      message: `${this.partName.name}: a relationships part in ${wrong === null ? 'no namespace' : `another namespace, ${wrong}`}, `
        + `read as Word's repair reads it, its ${count} relationship${count === 1 ? '' : 's'} by local name (REQ-207)`,
    }];
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

/**
 * `el` again, in the relationships namespace under its local name, with its attributes but not its namespace declarations,
 * and with what it holds moved across (CR-031 section 2.3).
 */
function inRelationshipsNamespace(doc: Document, el: Element): Element {
  const out = doc.createElementNS(Namespaces.RELATIONSHIPS, el.localName ?? el.nodeName);
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes.item(i)!;
    if (a.name === 'xmlns' || a.name.startsWith('xmlns:')) continue;
    out.setAttributeNS(a.namespaceURI, a.name, a.value);
  }
  if (el.localName === 'Relationship') while (el.firstChild) out.appendChild(el.firstChild);
  return out;
}

export function isExternal(rel: Relationship): boolean {
  return rel.targetMode === 'External';
}

setRelationshipsPartFactory((source) => RelationshipsPart.createRelationshipsPartForPart(source));
