import type { Relationship } from '@docx4j/generated-objects-ts/modules/org_docx4j_relationships';
import { PartName } from '../opc/PartName.mjs';
import { ContentTypeManager } from '../opc/ContentTypeManager.mjs';
import { MemoryPartSink, type PartStore, type PartSink } from '../opc/PartStore.mjs';
import { ZipPartStore, ZipPartSink } from '../opc/ZipPartStore.mjs';
import { FlatOpcPartStore, FlatOpcPartSink, type FlatOpcSinkOptions } from '../opc/FlatOpcPartStore.mjs';
import { loadPackage, type LoadOptions } from '../opc/Load.mjs';
import { savePackage, type SaveReport } from '../opc/Save.mjs';
import { Docx4JException } from '../opc/exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import { Part, type RelationshipSource } from '../parts/Part.mjs';
import { Parts } from '../parts/Parts.mjs';
import { XmlPart } from '../parts/XmlPart.mjs';
import { RelationshipsPart, type AddPartBehaviour } from '../parts/RelationshipsPart.mjs';
import type { CustomXmlDataStoragePart } from '../parts/DefaultXmlPart.mjs';
import { DocPropsCorePart, DocPropsExtendedPart, DocPropsCustomPart } from '../parts/docProps/index.mjs';
import { assertTimeZone, runtimeTimeZone } from '../model/dateZone.mjs';
import { mceKeep, mcePreprocess, type PartWarning } from '../opc/mce/preprocessor.mjs';
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
  /**
   * Whether a save writes run properties in their normal form (CR-015 D2; docx4j CR-061's
   * `setNormaliseOnSave`): a colour Word accepts as upper-case `RRGGBB`, one it rejects left out (a text or
   * underline colour) or `auto` (a fill's, a border's), a `w:highlight`, `w:vertAlign` or `w:em` Word repairs away
   * left out; a revision's former properties too. Only what is written changes: the tree keeps what it read, and a
   * part nobody unmarshalled is still copied byte for byte. Undefined defers to the objects package's process-wide
   * `setNormaliseOnSave` (on unless set off); a package's own setting wins over it. False writes values as read.
   * On/off values and `w:u`'s are settled when the part is read, so they are written the same either way.
   */
  normaliseOnSave: boolean | undefined;
  /** What the last save did beyond writing the package: each value it normalised (CR-015 D8), each it wrote past what Word stores (CR-017 P3). Undefined until one. */
  lastSave: SaveReport | undefined;
  private dateTimeZone_: string | undefined;
  /**
   * The client's time zone for this package, an IANA id (`Australia/Sydney`), or undefined for the runtime's (CR-018; the
   * portfolio's date rule of 2026-10-28, docx4j's `OpcPackage.setDateTimeZone`). Every instant the package holds is shown
   * in it, and a `w:date` without `dateUtc` read in it: a new revision's or comment's stamp (Word's local wall-clock
   * `w:date`), `TrackedChange.date` and `Comment.creationDate`, and date fields when they are evaluated (CR-012 phase 5).
   * It is the client's for this package, not the document's, and reaches no saved file as such: a server reusing a
   * loaded package sets it per request. Under Office JS the runtime's zone is already the client's. A bound date has no
   * zone (CR-002 section 56) unless `boundDatesConvertedAsWord` asks for Word's conversion in this one. A zone the
   * runtime does not know is refused with a `RangeError` at the set.
   */
  get dateTimeZone(): string | undefined {
    return this.dateTimeZone_;
  }
  set dateTimeZone(id: string | undefined) {
    this.dateTimeZone_ = id === undefined ? undefined : assertTimeZone(id);
  }
  /** The zone in force: `dateTimeZone`, else the runtime's. */
  effectiveDateTimeZone(): string {
    return this.dateTimeZone_ ?? runtimeTimeZone();
  }
  /**
   * Show bound dates as Word displays them in the package's zone (CR-018, the date rule's point 3; docx4j's
   * `setBoundDatesConvertedAsWord`): off by default, a bound date being the day its fields say wherever it is read
   * (CR-002 section 56). On, a bound value carrying a `Z` or an offset is read as an instant and shown in
   * `effectiveDateTimeZone()`, as Word reads the `Z` as UTC and shows the machine's local time (a day early west of
   * UTC); a value with no marker, or a date alone, stays nominal. Setting `dateTimeZone` does not turn this on.
   */
  boundDatesConvertedAsWord = false;
  /**
   * Write the save's time into the core properties (CR-018, the date rule's point 6; docx4j's `docx4j.dcterms.write`):
   * off by default, a save leaving `dcterms:created` and `dcterms:modified` as they are. On, each save sets
   * `dcterms:modified` to now, UTC to the second (`2026-10-28T03:04:05Z`, the W3CDTF form Office writes), and
   * `dcterms:created` too for a package created here that has none; a loaded package with no core properties part is
   * left without one, as docx4j leaves it.
   */
  writeTimesOnSave = false;

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
    if (opts.preprocessor === undefined) opts.preprocessor = opts.mcePreprocess === true ? mcePreprocess : mceKeep;
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

  /**
   * The warnings of every part parsed so far (each XML part's `warnings`, CR-010 phase 1). A part
   * is parsed when something reads it, so a reader that wants a whole package's warnings reads its
   * parts first; `unmarshalAll()` does, at the cost of re-marshalling every part on save.
   */
  get warnings(): PartWarning[] {
    const out: PartWarning[] = [];
    for (const part of this.parts) if (part instanceof XmlPart) out.push(...part.warnings);
    return out;
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

  /**
   * A deep copy of the package, of the same class (docx4j `OpcPackage.clone()`).
   *
   * Made the way docx4j makes it, **through save and reload** — into a `MemoryPartStore` rather
   * than a zip, so nothing is deflated — because that is the copy whose cost matches what was
   * touched: a part nobody unmarshalled is copied as the bytes it was loaded with, byte for byte,
   * and never parsed; an unmarshalled part is marshalled once and comes back in the clone as its
   * own tree, unmarshalled again on demand, sharing nothing with this package's. (The alternative,
   * a part-by-part copy with `deepCopy` on every unmarshalled tree, would have to unmarshal
   * lazily anyway and would not be simpler.)
   *
   * The clone's `sourcePartStore` is that memory store, so it no longer depends on this package's
   * container; the load options are this package's. Settings that are not in the package
   * (`author`, `fonts`, tracking) are carried across by the subclasses.
   */
  async clone(): Promise<this> {
    const store = await this.saveTo(new MemoryPartSink());
    const copy = await OpcPackage.load(store, this.loadOptions);
    if (copy.constructor !== this.constructor) {
      throw new Docx4JException(`Cloning ${this.constructor.name} produced ${copy.constructor.name}`);
    }
    copy.normaliseOnSave = this.normaliseOnSave;
    copy.dateTimeZone = this.dateTimeZone;
    copy.boundDatesConvertedAsWord = this.boundDatesConvertedAsWord;
    copy.writeTimesOnSave = this.writeTimesOnSave;
    this.copyPackageSettingsTo(copy);
    return copy as this;
  }

  /** Carries the settings that are not in the package itself into a clone; subclasses add theirs. */
  protected copyPackageSettingsTo(_target: OpcPackage): void {
    // OpcPackage itself has none
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
