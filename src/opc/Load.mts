import type { Relationship } from '@docx4j/generated-objects-ts/modules/org_docx4j_relationships';
import type { PartStore } from './PartStore.mjs';
import { PartName, CONTENT_TYPES_NAME } from './PartName.mjs';
import { ContentTypeManager } from './ContentTypeManager.mjs';
import { Docx4JException, DuplicateEntryException, InvalidFormatException, MultipleMainPartsException, type DuplicateEntryWarning, type MultipleMainPartsWarning } from './exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import { RelationshipsPart, isExternal } from '../parts/RelationshipsPart.mjs';
import { XmlPart, type DomPreprocessor } from '../parts/XmlPart.mjs';
import { CustomXmlDataStoragePart } from '../parts/DefaultXmlPart.mjs';
import { CustomXmlDataStoragePropertiesPart, DS_NS } from '../parts/customXml/index.mjs';
import { PartRegistry, defaultPartRegistry } from '../parts/PartRegistry.mjs';
import type { Part, RelationshipSource } from '../parts/Part.mjs';
import type { OpcPackage } from '../packages/OpcPackage.mjs';
import { decodeXmlText, parseXml } from '../xml/dom.mjs';

export interface LoadOptions {
  /** Part classes by content type; defaults to docx4j's table. */
  registry?: PartRegistry;
  /**
   * Resolve `mc:AlternateContent` before unmarshalling, as Word does on open. Default false since
   * CR-010 phase 4: every element is kept with all its branches, as docx4j keeps it since its CR-021;
   * readers take the branch resolving would (`selectedBranchOf`), mutators every branch, and a save
   * writes them all. Either way a part's `warnings` report where the branches' text differs. True
   * resolves, but keeps the element under a `KEPT_HOSTS` host and one nothing can be selected from
   * (CR-010 phase 1).
   */
  mcePreprocess?: boolean;
  /** Runs over every XML part's DOM before it is unmarshalled; set by the package from `mcePreprocess` (`mceKeep` or `mcePreprocess`). */
  preprocessor?: DomPreprocessor;
  /**
   * A container holding two entries of one name, OPC part names compared without case (WD REQ-206's `duplicate_entries`;
   * CR-024): `'refuse'`, the default, throws `DuplicateEntryException` naming each before anything is read; `'first'` reads
   * the first entry of each name in the container's order (a zip's central directory), as Word reads it after its repair,
   * ignores the rest, and records a `DuplicateEntryWarning` per name on `pkg.warnings`. A save then writes only the copy read.
   *
   * The same switch decides a package with more than one officeDocument relationship (REQ-206's statement, WD 0.100; CR-024
   * section 9): `'refuse'` throws `MultipleMainPartsException` naming the targets, and `'first'` reads the **last**
   * relationship's target, as a fresh Word does, records one `MultipleMainPartsWarning`, and drops the other relationships,
   * so a save writes one main part. The value says `first` because it's the WD's one switch for both.
   */
  duplicateEntries?: 'refuse' | 'first';
}

/** Relationship types whose target is the main part (docx4j PackageRelsUtil.getNameOfMainPart). */
const MAIN_PART_RELATIONSHIPS = [
  Namespaces.DOCUMENT,
  'http://schemas.microsoft.com/office/2006/relationships/graphicFrameDoc',
  Namespaces.DRAWINGML_DIAGRAM_LAYOUT,
  Namespaces.DOCUMENT_STRICT,
];

/** What REQ-206's statement counts (docx4j PackageRelsUtil.isMainPart): a transitional and a strict relationship are two. */
const OFFICE_DOCUMENT_RELATIONSHIPS: readonly string[] = [Namespaces.DOCUMENT, Namespaces.DOCUMENT_STRICT];

/**
 * Loads a package from a container (docx4j io3.Load3): content types, then the package
 * relationships, then every part reachable through relationships, recursively. Parts are not
 * unmarshalled. `createPackage` picks the package class from the main part's content type.
 */
export async function loadPackage<P extends OpcPackage>(
  store: PartStore,
  createPackage: (mainPartContentType: string | undefined) => P,
  options: LoadOptions = {},
): Promise<P> {
  const registry = options.registry ?? defaultPartRegistry;
  const duplicateEntries = duplicateEntriesOf(store, options.duplicateEntries);
  if (!store.has(CONTENT_TYPES_NAME)) throw new InvalidFormatException(`${CONTENT_TYPES_NAME} is missing from this package`);
  const ctm = ContentTypeManager.parse(decodeXmlText(await store.load(CONTENT_TYPES_NAME)));

  const packageRels = RelationshipsPart.createPackageRels();
  if (!store.has('_rels/.rels')) throw new InvalidFormatException('_rels/.rels appears to be missing from this package');
  packageRels.setBytes(await store.load('_rels/.rels'));
  await packageRels.getContents();
  const mainParts = oneMainPart(packageRels, options.duplicateEntries);

  const mainRel = packageRels.list.find((r) => MAIN_PART_RELATIONSHIPS.includes(r.type));
  if (!mainRel) throw new InvalidFormatException('No relationship of type officeDocument');
  const mainPartName = PartName.resolve(PartName.ROOT, mainRel.target);
  const pkg = createPackage(ctm.getContentType(mainPartName));
  pkg.contentTypeManager = ctm;
  pkg.sourcePartStore = store;
  pkg.relationshipsPart = packageRels;
  pkg.duplicateEntries = duplicateEntries;
  pkg.multipleMainParts = mainParts?.warning;
  packageRels.sourceP = pkg;
  packageRels.package = pkg;

  const loader = new Loader(pkg, store, registry, options.preprocessor);
  await loader.addPartsFromRelationships(pkg, packageRels);
  await loader.registerCustomXmlDataStorageParts();
  // after the walk, so a target the kept part reaches by a relationship of its own, or a second relationship names, keeps its type
  for (const name of mainParts?.dropped ?? []) if (!pkg.getPart(name)) ctm.removeOverrideContentType(name);
  return pkg;
}

/**
 * More than one officeDocument relationship (WD REQ-206's statement; CR-024 section 9): refused under `'refuse'`, naming the
 * targets. Under `'first'` the last stands and the others go here, before anything is read, so that the package's class, the
 * walk, `getMainPart()` and `getMainDocumentPart()` all find one main part, and a save writes one. `dropped` is the part
 * names they gave, whose Overrides `loadPackage` settles after the walk.
 */
function oneMainPart(rels: RelationshipsPart, mode: LoadOptions['duplicateEntries']): { warning: MultipleMainPartsWarning; dropped: PartName[] } | undefined {
  const mains = rels.list.filter((r) => OFFICE_DOCUMENT_RELATIONSHIPS.includes(r.type));
  if (mains.length < 2) return undefined;
  // `PartName.resolve` would make a part name of a URL
  const named = (r: Relationship) => (isExternal(r) ? r.target : PartName.resolve(PartName.ROOT, r.target).name);
  if (mode !== 'first') throw new MultipleMainPartsException(mains.map(named));
  const others = mains.slice(0, -1);
  for (const r of others) rels.removeRelationship(r);
  const partName = named(mains[mains.length - 1]!);
  return {
    warning: {
      code: 'package.multiple_main_parts',
      partName,
      count: mains.length,
      message: `${mains.length} officeDocument relationships in this package (REQ-206); the last's target read, ${partName}`,
    },
    dropped: others.filter((r) => !isExternal(r)).map((r) => PartName.resolve(PartName.ROOT, r.target)),
  };
}

/**
 * The store's duplicated names (CR-024, WD REQ-206), refused under `'refuse'` before anything is read, else one warning each.
 * A store without `duplicates()` is not checked.
 */
function duplicateEntriesOf(store: PartStore, mode: LoadOptions['duplicateEntries']): DuplicateEntryWarning[] {
  if (mode !== undefined && mode !== 'refuse' && mode !== 'first') {
    throw new Docx4JException(`duplicateEntries: '${String(mode)}' is not 'refuse' or 'first' (WD REQ-206)`);
  }
  const duplicates = store.duplicates?.();
  if (!duplicates || duplicates.size === 0) return [];
  if (mode !== 'first') throw new DuplicateEntryException(Object.fromEntries(duplicates));
  return [...duplicates].map(([entry, count]) => ({
    code: 'package.duplicate_entry' as const,
    entry,
    count,
    message: `${entry}: ${count} entries of one name in this package; the first read, the rest ignored (REQ-206)`,
  }));
}

class Loader {
  private readonly handled = new Set<string>();

  constructor(
    private readonly pkg: OpcPackage,
    private readonly store: PartStore,
    private readonly registry: PartRegistry,
    private readonly preprocessor: DomPreprocessor | undefined,
  ) {}

  async addPartsFromRelationships(source: RelationshipSource, rp: RelationshipsPart): Promise<void> {
    for (const r of rp.list) {
      try {
        await this.getPart(source, rp, r);
      } catch (e) {
        if (e instanceof Docx4JException) throw e;
        throw new Docx4JException(`Failed to add parts from relationships of ${rp.partName}`, { cause: e });
      }
    }
  }

  private async getPart(source: RelationshipSource, rp: RelationshipsPart, r: Relationship): Promise<void> {
    if (r.type === Namespaces.HYPERLINK || isExternal(r)) return;
    const partName = rp.resolveTarget(r);
    const existing = this.handled.has(partName.key) ? this.pkg.getPart(partName) : undefined;
    if (existing) {
      // A part can be the target of several relationships (a header used by two sections, an image twice).
      source.setPartShortcut(existing, r.type);
      existing.sourceRelationships.push(r);
      return;
    }
    const part = this.createPart(partName, r, rp);
    source.setPartShortcut(part, r.type);
    if (part.relationshipType === '') part.relationshipType = r.type;
    if (part instanceof XmlPart && this.preprocessor) part.preprocessor = this.preprocessor;
    rp.loadPart(part, r);
    this.handled.add(partName.key);

    const relsName = PartName.relsFor(partName).storeName;
    if (this.store.has(relsName)) {
      const rrp = new RelationshipsPart(part);
      rrp.setBytes(await this.store.load(relsName));
      await rrp.getContents();
      part.relationshipsPart = rrp;
      await this.addPartsFromRelationships(part, rrp);
    }
  }

  private createPart(partName: PartName, r: Relationship, rp: RelationshipsPart): Part {
    const storeName = partName.storeName;
    if (!this.store.has(storeName) && !this.store.has(decodeURIComponent(storeName))) {
      throw new Docx4JException(`For source ${rp.sourceName}, cannot find part ${partName} from rel ${r.id}=${r.target}`);
    }
    return this.registry.createPart(partName, this.pkg.contentTypeManager.getContentType(partName), r);
  }

  /** Indexes custom XML parts by the itemID of their properties part (docx4j Load.registerCustomXmlDataStorageParts). */
  async registerCustomXmlDataStorageParts(): Promise<void> {
    for (const part of this.pkg.parts) {
      if (!(part instanceof CustomXmlDataStoragePart)) continue;
      const rel = part.relationshipsPart?.getRelationshipByType(Namespaces.CUSTOM_XML_DATA_STORAGE_PROPERTIES);
      const props = rel ? part.relationshipsPart?.getPart(rel) : undefined;
      if (!(props instanceof CustomXmlDataStoragePropertiesPart)) continue;
      // Read the itemID from the DOM rather than unmarshalling, so the properties part stays untouched.
      const doc = parseXml(decodeXmlText(await props.getBytes()));
      const itemId = doc.documentElement.getAttributeNS(DS_NS, 'itemID') || doc.documentElement.getAttribute('ds:itemID');
      if (!itemId) continue;
      part.itemId = itemId.toLowerCase();
      this.pkg.customXmlDataStorageParts.set(part.itemId, part);
    }
  }
}
