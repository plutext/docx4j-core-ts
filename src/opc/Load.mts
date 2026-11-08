import type { Relationship } from '@docx4j/generated-objects-ts/modules/org_docx4j_relationships';
import type { PartStore } from './PartStore.mjs';
import { PartName, CONTENT_TYPES_NAME } from './PartName.mjs';
import { ContentTypeManager } from './ContentTypeManager.mjs';
import { Docx4JException, DuplicateEntryException, ExternalMainPartException, InvalidFormatException, MultipleMainPartsException, type DuplicateEntryWarning, type ExternalMainPartWarning, type MultipleMainPartsWarning } from './exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import { RelationshipsPart, isExternal } from '../parts/RelationshipsPart.mjs';
import { XmlPart, type DomPreprocessor } from '../parts/XmlPart.mjs';
import { CustomXmlDataStoragePart } from '../parts/DefaultXmlPart.mjs';
import { CustomXmlDataStoragePropertiesPart, DS_NS } from '../parts/customXml/index.mjs';
import { PartRegistry, defaultPartRegistry } from '../parts/PartRegistry.mjs';
import type { Part, RelationshipSource } from '../parts/Part.mjs';
import type { OpcPackage } from '../packages/OpcPackage.mjs';
import { decodeXmlText, parseXml } from '../xml/dom.mjs';

/**
 * WD REQ-207's option (CR-031): how a load treats a package Word reads only through a repair, or won't open. `'refuse'`, the
 * default, keeps every refusal; `'word'` follows Word's repair where the WD states one, with that repair's warning; `'all'`
 * does that and makes this package's own repairs, each documented and warned (README.md, "Repairs").
 */
export type RepairMode = 'refuse' | 'word' | 'all';

/** The values of `LoadOptions.repair`, in the WD's order. */
export const REPAIR_MODES: readonly RepairMode[] = Object.freeze(['refuse', 'word', 'all']);

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
   * How a package Word reads only through a repair, or won't open, is loaded (WD REQ-207 and REQ-208; CR-031). It acts at the
   * load, so it holds for every operation on the package. Any other value is refused.
   *
   * - `'refuse'`, the default: every refusal stands. A container holding two entries of one name, OPC part names compared
   *   without case, throws `DuplicateEntryException` (REQ-206, CR-024); more than one officeDocument relationship,
   *   `MultipleMainPartsException` (CR-024 section 9); a relationships part in another namespace or none,
   *   `RelationshipsNamespaceException` (CR-025 section 8); an `External` officeDocument relationship beside an internal
   *   one, `ExternalMainPartException` (REQ-206's NOTE).
   * - `'word'`: Word's repair, each with its warning. The first entry of each duplicated name is read, in the container's
   *   order (a zip's central directory), the rest ignored, a `DuplicateEntryWarning` each; of several officeDocument
   *   relationships the **last** one's target, as a fresh Word reads it, the others dropped, one `MultipleMainPartsWarning`;
   *   a relationships part in a wrong namespace read by local name, a `RelationshipsNamespaceWarning`. A save writes the
   *   repaired package: one copy of each entry, one main part, the right namespace. `External` beside an internal
   *   relationship is still refused, since Word refuses it even through its repair.
   * - `'all'`: `'word'`, and this package's own repairs. One today: `External` beside an internal relationship, the internal
   *   target read and every `External` officeDocument relationship dropped, its URL never fetched, an
   *   `ExternalMainPartWarning`.
   *
   * At every value an `External` officeDocument relationship with no internal one is refused, as a package with no main
   * document part (REQ-321), and the save's own gate is left alone: `pkg.checkLossOnSave` decides whether a save may lose
   * what a read skipped (CR-031 D6).
   */
  repair?: RepairMode;
}

/** Relationship types whose target is the main part (docx4j PackageRelsUtil.getNameOfMainPart). */
const MAIN_PART_RELATIONSHIPS = [
  Namespaces.DOCUMENT,
  'http://schemas.microsoft.com/office/2006/relationships/graphicFrameDoc',
  Namespaces.DRAWINGML_DIAGRAM_LAYOUT,
  Namespaces.DOCUMENT_STRICT,
];

/**
 * What REQ-206's statement counts (docx4j PackageRelsUtil.isMainPart): a transitional and a strict relationship are two. An
 * External one isn't counted (CR-024 section 9.2): it names no part, so it can't be what the package reads.
 */
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
  const repair = repairOf(options.repair);
  const followsWord = repair !== 'refuse';
  const duplicateEntries = duplicateEntriesOf(store, followsWord);
  if (!store.has(CONTENT_TYPES_NAME)) throw new InvalidFormatException(`${CONTENT_TYPES_NAME} is missing from this package`);
  const ctm = ContentTypeManager.parse(decodeXmlText(await store.load(CONTENT_TYPES_NAME)));

  const packageRels = RelationshipsPart.createPackageRels();
  if (!store.has('_rels/.rels')) throw new InvalidFormatException('_rels/.rels appears to be missing from this package');
  packageRels.readsByLocalName = followsWord;
  packageRels.setBytes(await store.load('_rels/.rels'));
  await packageRels.getContents();
  // External first, as docx4j has it: refused at refuse and word whatever else the relationships hold, and dropped at all
  // before Word's last-relationship repair
  const external = externalMainPart(packageRels, repair);
  const mainParts = oneMainPart(packageRels, followsWord);

  // an External relationship names no part, so it gives way to one that does
  const mainRels = packageRels.list.filter((r) => MAIN_PART_RELATIONSHIPS.includes(r.type));
  const mainRel = mainRels.find((r) => !isExternal(r)) ?? mainRels[0];
  if (!mainRel) throw new InvalidFormatException('No relationship of type officeDocument');
  const mainPartName = PartName.resolve(PartName.ROOT, mainRel.target);
  const pkg = createPackage(ctm.getContentType(mainPartName));
  pkg.contentTypeManager = ctm;
  pkg.sourcePartStore = store;
  pkg.relationshipsPart = packageRels;
  pkg.duplicateEntries = duplicateEntries;
  pkg.multipleMainParts = mainParts?.warning;
  pkg.externalMainPart = external;
  packageRels.sourceP = pkg;
  packageRels.package = pkg;

  const loader = new Loader(pkg, store, registry, options.preprocessor, followsWord);
  await loader.addPartsFromRelationships(pkg, packageRels);
  await loader.registerCustomXmlDataStorageParts();
  // after the walk, so a target the kept part reaches by a relationship of its own, or a second relationship names, keeps its type
  for (const name of mainParts?.dropped ?? []) if (!pkg.getPart(name)) ctm.removeOverrideContentType(name);
  return pkg;
}

/**
 * More than one officeDocument relationship (WD REQ-206's statement; CR-024 section 9): refused under `repair: 'refuse'`, naming
 * the targets. Under `'word'` and `'all'` the last stands and the others go here, before anything is read, so that the package's class, the
 * walk, `getMainPart()` and `getMainDocumentPart()` all find one main part, and a save writes one. `dropped` is the part
 * names they gave, whose Overrides `loadPackage` settles after the walk.
 */
function oneMainPart(rels: RelationshipsPart, followsWord: boolean): { warning: MultipleMainPartsWarning; dropped: PartName[] } | undefined {
  const mains = rels.list.filter((r) => OFFICE_DOCUMENT_RELATIONSHIPS.includes(r.type) && !isExternal(r));
  if (mains.length < 2) return undefined;
  const names = mains.map((r) => PartName.resolve(PartName.ROOT, r.target));
  if (!followsWord) throw new MultipleMainPartsException(names.map((name) => name.name));
  for (const r of mains.slice(0, -1)) rels.removeRelationship(r);
  const partName = names[names.length - 1]!.name;
  return {
    warning: {
      code: 'package.multiple_main_parts',
      partName,
      count: mains.length,
      message: `${mains.length} officeDocument relationships in this package (REQ-206); the last's target read, ${partName}`,
    },
    dropped: names.slice(0, -1),
  };
}

/** `LoadOptions.repair`, `'refuse'` when not given; any other value refused (WD REQ-207). */
function repairOf(repair: unknown): RepairMode {
  if (repair === undefined) return 'refuse';
  if (!(REPAIR_MODES as readonly unknown[]).includes(repair)) {
    throw new Docx4JException(`repair: '${String(repair)}' is not ${REPAIR_MODES.map((m) => `'${m}'`).join(', ')} (WD REQ-207)`);
  }
  return repair as RepairMode;
}

/**
 * An `External` officeDocument relationship (WD REQ-206's NOTE; CR-031 section 2.2). With no internal one beside it the package
 * has no main document part, refused at every value (REQ-321), where the URL used to be taken for a part name. Beside an
 * internal one it's refused at `'refuse'` and `'word'`, first or last alike, since Word refuses it even through its repair;
 * at `'all'` every such relationship goes here, before anything is read, so a save writes none and the URL is never
 * fetched, and the warning names the part read: the last internal relationship's target, as Word's repair of several reads.
 */
function externalMainPart(rels: RelationshipsPart, repair: RepairMode): ExternalMainPartWarning | undefined {
  const mains = rels.list.filter((r) => OFFICE_DOCUMENT_RELATIONSHIPS.includes(r.type));
  const externals = mains.filter(isExternal);
  if (externals.length === 0) return undefined;
  const urls = externals.map((r) => r.target);
  const internals = mains.filter((r) => !isExternal(r)).map((r) => PartName.resolve(PartName.ROOT, r.target).name);
  if (internals.length === 0) {
    throw new InvalidFormatException(`This package's only officeDocument relationship is External, to ${urls.join(', ')}, so it has no `
      + 'main document part (REQ-321); refused at every value of repair. The URL is not fetched');
  }
  if (repair !== 'all') throw new ExternalMainPartException(urls, internals);
  for (const r of externals) rels.removeRelationship(r);
  const partName = internals[internals.length - 1]!;
  return {
    code: 'package.external_main_part',
    partName,
    urls,
    message: `${urls.length === 1 ? 'An External officeDocument relationship' : `${urls.length} External officeDocument relationships`}, `
      + `to ${urls.join(', ')}, beside the internal one dropped, ${partName} read, the URL not fetched: a repair of this package's own (REQ-207)`,
  };
}

/**
 * The store's duplicated names (CR-024, WD REQ-206), refused under `repair: 'refuse'` before anything is read, else one warning
 * each. A store without `duplicates()` is not checked.
 */
function duplicateEntriesOf(store: PartStore, followsWord: boolean): DuplicateEntryWarning[] {
  const duplicates = store.duplicates?.();
  if (!duplicates || duplicates.size === 0) return [];
  if (!followsWord) throw new DuplicateEntryException(Object.fromEntries(duplicates));
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
    private readonly followsWord: boolean,
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
      rrp.readsByLocalName = this.followsWord;
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
