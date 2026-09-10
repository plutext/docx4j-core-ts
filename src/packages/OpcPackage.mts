import type { Relationship } from '@docx4j/generated-objects-ts/modules/org_docx4j_relationships';
import { PartName } from '../opc/PartName.mjs';
import { ContentTypeManager } from '../opc/ContentTypeManager.mjs';
import type { PartStore, PartSink } from '../opc/PartStore.mjs';
import { ZipPartStore, ZipPartSink } from '../opc/ZipPartStore.mjs';
import { FlatOpcPartStore, FlatOpcPartSink, type FlatOpcSinkOptions } from '../opc/FlatOpcPartStore.mjs';
import { loadPackage, type LoadOptions } from '../opc/Load.mjs';
import { savePackage } from '../opc/Save.mjs';
import { Docx4JException } from '../opc/exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import { Part, type RelationshipSource } from '../parts/Part.mjs';
import { Parts } from '../parts/Parts.mjs';
import { XmlPart } from '../parts/XmlPart.mjs';
import { RelationshipsPart, type AddPartBehaviour } from '../parts/RelationshipsPart.mjs';
import type { CustomXmlDataStoragePart } from '../parts/DefaultXmlPart.mjs';
import { DocPropsCorePart, DocPropsExtendedPart, DocPropsCustomPart } from '../parts/docProps/index.mjs';
import { mcePreprocess } from '../opc/mce/preprocessor.mjs';
import { createPackageForContentType, registerGenericPackageClass } from './registry.mjs';

/** Anything `OpcPackage.load` accepts: zip bytes, a flat OPC string, or a container. */
export type PackageSource = Uint8Array | ArrayBuffer | string | PartStore;

/**
 * An Open Packaging Conventions package (docx4j org.docx4j.openpackaging.packages.OpcPackage):
 * the content types, the package relationships and the parts reachable from them.
 */
export class OpcPackage implements RelationshipSource {
  /** The pseudo part name of the package itself: the source of the package relationships. */
  readonly partName = PartName.ROOT;
  contentTypeManager: ContentTypeManager = ContentTypeManager.createDefault();
  relationshipsPart: RelationshipsPart;
  readonly parts = new Parts();
  /** Where untouched parts are copied from on save. Undefined for a new package. */
  sourcePartStore: PartStore | undefined;
  /** Custom XML parts by the (lower-cased) itemID of their properties part. */
  readonly customXmlDataStorageParts = new Map<string, CustomXmlDataStoragePart>();
  /** The options this package was loaded with. */
  loadOptions: LoadOptions = {};
  docPropsCorePart: DocPropsCorePart | undefined;
  docPropsExtendedPart: DocPropsExtendedPart | undefined;
  docPropsCustomPart: DocPropsCustomPart | undefined;

  constructor() {
    this.relationshipsPart = RelationshipsPart.createPackageRels();
    this.relationshipsPart.sourceP = this;
    this.relationshipsPart.package = this;
  }

  get package(): OpcPackage {
    return this;
  }

  /**
   * Loads a package, picking the class (WordprocessingMLPackage, ...) from the main part's
   * content type. A string is a flat OPC package; bytes are a zip; a PartStore is used as is.
   */
  static async load(source: PackageSource, options: LoadOptions = {}): Promise<OpcPackage> {
    const store = toStore(source);
    const opts: LoadOptions = { ...options };
    if (opts.preprocessor === undefined && opts.mcePreprocess !== false) opts.preprocessor = mcePreprocess;
    const pkg = await loadPackage(store, createPackageForContentType, opts);
    pkg.loadOptions = opts;
    return pkg;
  }

  getPart(partName: PartName | string): Part | undefined {
    return this.parts.get(partName);
  }

  getRelationshipsPart(): RelationshipsPart {
    return this.relationshipsPart;
  }

  /** Adds a part as a target of the package (docProps, the main part). */
  addTargetPart(target: Part, mode?: AddPartBehaviour, proposedRelId?: string): Relationship {
    const rel = this.relationshipsPart.addPart(target, mode, proposedRelId);
    this.setPartShortcut(target, target.relationshipType);
    return rel;
  }

  setPartShortcut(part: Part, relationshipType: string): boolean {
    switch (relationshipType) {
      case Namespaces.PROPERTIES_CORE: this.docPropsCorePart = part as DocPropsCorePart; return true;
      case Namespaces.PROPERTIES_EXTENDED: this.docPropsExtendedPart = part as DocPropsExtendedPart; return true;
      case Namespaces.PROPERTIES_CUSTOM: this.docPropsCustomPart = part as DocPropsCustomPart; return true;
      default: return false;
    }
  }

  /** The main part: the target of the officeDocument relationship. */
  getMainPart(): Part | undefined {
    const rel = this.relationshipsPart.getRelationshipByType(Namespaces.DOCUMENT) ?? this.relationshipsPart.getRelationshipByType(Namespaces.DOCUMENT_STRICT);
    return rel ? this.relationshipsPart.getPart(rel) : undefined;
  }

  /** Unmarshals every XML part, so that `contents` is available synchronously everywhere. */
  async unmarshalAll(): Promise<void> {
    for (const part of this.parts) if (part instanceof XmlPart) await part.getContents();
  }

  /** Adds empty core properties if absent (docx4j addDocPropsCorePart). */
  addDocPropsCorePart(): DocPropsCorePart {
    if (!this.docPropsCorePart) {
      const part = new DocPropsCorePart();
      part.setContents({});
      this.addTargetPart(part);
    }
    return this.docPropsCorePart!;
  }

  addDocPropsExtendedPart(): DocPropsExtendedPart {
    if (!this.docPropsExtendedPart) {
      const part = new DocPropsExtendedPart();
      part.setContents({});
      this.addTargetPart(part);
    }
    return this.docPropsExtendedPart!;
  }

  /** The package as a zip. */
  save(options?: { level?: number }): Promise<Uint8Array> {
    return this.saveTo(new ZipPartSink(options));
  }

  /** The package as a flat OPC `pkg:package` string, for Office JS `insertOoxml()`. */
  saveFlatOpc(options?: FlatOpcSinkOptions): Promise<string> {
    return this.saveTo(new FlatOpcPartSink({ progId: this.progId, ...options }));
  }

  /** Writes the package to any sink. */
  saveTo<R>(sink: PartSink<R>): Promise<R> {
    return savePackage(this, sink);
  }

  /** The `mso-application` progid written to flat OPC; subclasses know theirs. */
  protected get progId(): string | undefined {
    return undefined;
  }
}

function toStore(source: PackageSource): PartStore {
  if (typeof source === 'string') return new FlatOpcPartStore(source);
  if (source instanceof Uint8Array || source instanceof ArrayBuffer) return new ZipPartStore(source);
  if (isPartStore(source)) return source;
  throw new Docx4JException('Unsupported package source');
}

function isPartStore(value: unknown): value is PartStore {
  return typeof value === 'object' && value !== null && typeof (value as PartStore).load === 'function' && typeof (value as PartStore).has === 'function';
}

registerGenericPackageClass(OpcPackage);
