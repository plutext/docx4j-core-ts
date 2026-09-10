import type { Relationship } from '@docx4j/generated-objects-ts/modules/org_docx4j_relationships';
import type { PartStore } from './PartStore.mjs';
import { PartName, CONTENT_TYPES_NAME } from './PartName.mjs';
import { ContentTypeManager } from './ContentTypeManager.mjs';
import { Docx4JException, InvalidFormatException } from './exceptions.mjs';
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
  /** Resolve `mc:AlternateContent` before unmarshalling (default true, as docx4j and Word). */
  mcePreprocess?: boolean;
  /** Runs over every XML part's DOM before it is unmarshalled; set by the package from `mcePreprocess`. */
  preprocessor?: DomPreprocessor;
}

/** Relationship types whose target is the main part (docx4j PackageRelsUtil.getNameOfMainPart). */
const MAIN_PART_RELATIONSHIPS = [
  Namespaces.DOCUMENT,
  'http://schemas.microsoft.com/office/2006/relationships/graphicFrameDoc',
  Namespaces.DRAWINGML_DIAGRAM_LAYOUT,
  Namespaces.DOCUMENT_STRICT,
];

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
  if (!store.has(CONTENT_TYPES_NAME)) throw new InvalidFormatException(`${CONTENT_TYPES_NAME} is missing from this package`);
  const ctm = ContentTypeManager.parse(decodeXmlText(await store.load(CONTENT_TYPES_NAME)));

  const packageRels = RelationshipsPart.createPackageRels();
  if (!store.has('_rels/.rels')) throw new InvalidFormatException('_rels/.rels appears to be missing from this package');
  packageRels.setBytes(await store.load('_rels/.rels'));
  await packageRels.getContents();

  const mainRel = packageRels.list.find((r) => MAIN_PART_RELATIONSHIPS.includes(r.type));
  if (!mainRel) throw new InvalidFormatException('No relationship of type officeDocument');
  const mainPartName = PartName.resolve(PartName.ROOT, mainRel.target);
  const pkg = createPackage(ctm.getContentType(mainPartName));
  pkg.contentTypeManager = ctm;
  pkg.sourcePartStore = store;
  pkg.relationshipsPart = packageRels;
  packageRels.sourceP = pkg;
  packageRels.package = pkg;

  const loader = new Loader(pkg, store, registry, options.preprocessor);
  await loader.addPartsFromRelationships(pkg, packageRels);
  await loader.registerCustomXmlDataStorageParts();
  return pkg;
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
