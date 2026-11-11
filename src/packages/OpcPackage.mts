import type { Relationship } from '@docx4j/generated-objects-ts/modules/org_docx4j_relationships';
import { PartName } from '../opc/PartName.mjs';
import { ContentTypeManager } from '../opc/ContentTypeManager.mjs';
import { MemoryPartSink, type PartStore, type PartSink } from '../opc/PartStore.mjs';
import { ZipPartStore, ZipPartSink } from '../opc/ZipPartStore.mjs';
import { FlatOpcPartStore, FlatOpcPartSink, type FlatOpcSinkOptions } from '../opc/FlatOpcPartStore.mjs';
import { loadPackage, type LoadOptions } from '../opc/Load.mjs';
import { savePackage, type SaveReport } from '../opc/Save.mjs';
import { Docx4JException, type DuplicateEntryWarning, type ExternalMainPartWarning, type MultipleMainPartsWarning, type RelationshipsNamespaceWarning, type RelationshipsNotReadWarning } from '../opc/exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import { Part, type RelationshipSource } from '../parts/Part.mjs';
import { Parts } from '../parts/Parts.mjs';
import { XmlPart } from '../parts/XmlPart.mjs';
import { RelationshipsPart, isExternal, type AddPartBehaviour } from '../parts/RelationshipsPart.mjs';
import type { CustomXmlDataStoragePart } from '../parts/DefaultXmlPart.mjs';
import { DocPropsCorePart, DocPropsExtendedPart, DocPropsCustomPart } from '../parts/docProps/index.mjs';
import { assertTimeZone, runtimeTimeZone } from '../model/dateZone.mjs';
import { mceKeep, mcePreprocess, type PartWarning, type MisplacedContentWarning } from '../opc/mce/preprocessor.mjs';
import { createPackageForContentType, registerGenericPackageClass } from './registry.mjs';

/** A warning on `pkg.warnings`: the load's (a duplicated entry, CR-024; an `External` main relationship, CR-031) or a part's. */
export type PackageWarning = DuplicateEntryWarning | MultipleMainPartsWarning | ExternalMainPartWarning | PartWarning;

/**
 * What a load repaired (WD REQ-207's `repaired`, CR-031): `'own'` where a repair of this package's own was made, `'word'` where
 * only Word's were followed, else `'none'`.
 */
export type Repaired = 'none' | 'word' | 'own';

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
  /**
   * The names the container held more than one entry of, each read from its first entry under
   * `LoadOptions.repair: 'word'` (WD REQ-206, CR-024); empty for a package loaded without one, or made here.
   */
  duplicateEntries: readonly DuplicateEntryWarning[] = [];
  /**
   * Where the package relationships held more than one officeDocument relationship, read from the last under
   * `LoadOptions.repair: 'word'` (WD REQ-206's statement, CR-024 section 9); undefined for a package with one, or made here.
   */
  multipleMainParts: MultipleMainPartsWarning | undefined;
  /**
   * Where the package relationships held an `External` officeDocument relationship beside an internal one, dropped under
   * `LoadOptions.repair: 'all'` (WD REQ-206's NOTE, CR-031 section 2.2); undefined otherwise.
   */
  externalMainPart: ExternalMainPartWarning | undefined;
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
  /** What the last save did beyond writing the package: each value it normalised (CR-015 D8), each it wrote past what Word stores (CR-017 P3), each misplaced child it wrote under `checkContentOnSave = false` (CR-016). Undefined until one. */
  lastSave: SaveReport | undefined;
  /**
   * Whether a save refuses a part holding a child its container never admits (CR-016 D2, docx4j CR-064's rules in the
   * objects facade, its CR-017): a run straight in the body, a table in a paragraph, a cell beside a table's rows, a
   * string in a run, a name and value that disagree. On by default: the save throws `MisplacedContentException` with
   * every such child in the part, its path and docx4j's message naming the fix, and writes nothing, since Word 365
   * refuses such a file or repairs it by dropping content (check 72). Off, the part is written as held and each is on
   * `lastSave.misplaced`. Only a part being marshalled is checked (one unmarshalled, or given contents); a part copied
   * from the source is never refused, so a document read and saved untouched never is. `checkContent()` is the same
   * check on request. A misplaced child read from a file is a `target.misplaced` warning on its part (D7).
   */
  checkContentOnSave = true;
  /**
   * Whether a save refuses to lose what a read dropped (CR-025 sections 2.3 and 2.4, D8): a part holding `part.skipped` content
   * that was changed throws `SaveLossException` and nothing is written. Off, the part is written without those items, each on
   * `lastSave.lost`. A part read and not changed is copied either way, so nothing is lost there. Word's own save drops such
   * content with no message (check 76), so off is Word's behaviour.
   */
  checkLossOnSave = true;
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

  /** The main part: the target of the officeDocument relationship, an External one naming none. */
  getMainPart(): Part | undefined {
    const rels = this.relationshipsPart;
    const rel = [...rels.getRelationshipsByType(Namespaces.DOCUMENT), ...rels.getRelationshipsByType(Namespaces.DOCUMENT_STRICT)].find((r) => !isExternal(r));
    return rel ? rels.getPart(rel) : undefined;
  }

  /**
   * The warnings of every part parsed so far (each XML part's `warnings`, CR-010 phase 1), after the
   * load's own: a name the container holds twice, and more than one officeDocument relationship, read under
   * `repair: 'word'` (CR-024), and an `External` main relationship dropped under `repair: 'all'` (CR-031). A part
   * is parsed when something reads it, so a reader that wants a whole package's warnings reads its
   * parts first; `unmarshalAll()` does, at the cost of re-marshalling every part on save.
   */
  get warnings(): PackageWarning[] {
    const out: PackageWarning[] = [...this.duplicateEntries];
    if (this.multipleMainParts) out.push(this.multipleMainParts);
    if (this.externalMainPart) out.push(this.externalMainPart);
    for (const part of this.parts) if (part instanceof XmlPart) out.push(...part.warnings);
    // a relationships part's own, what its read dropped: a child of another local name, an attribute the model doesn't map
    // (CR-025 section 8, D12); after the parts', and outside the save's gate, since none names a part
    for (const rels of [this.relationshipsPart, ...[...this.parts].map((p) => p.relationshipsPart)]) if (rels) out.push(...rels.warnings);
    return out;
  }

  /**
   * Each relationships part a load under `repair: 'word'` read by local name, being in a wrong namespace (CR-031 section
   * 2.3): the package's and its parts', in that order.
   */
  get relationshipsNamespaces(): RelationshipsNamespaceWarning[] {
    const out: RelationshipsNamespaceWarning[] = [];
    for (const rels of [this.relationshipsPart, ...[...this.parts].map((p) => p.relationshipsPart)]) {
      for (const w of rels?.warnings ?? []) if (w.code === 'package.relationships_namespace') out.push(w);
    }
    return out;
  }

  /**
   * Each relationships part that wasn't well formed and was read short, the relationships before the break kept (CR-032): the
   * package's and its parts', in that order. A save is refused while one is here unless the load's `repair` is `'word'` or
   * `'all'`, or `checkLossOnSave` is false. Not counted in `repaired`, as docx4j doesn't count it: it is read so at every value.
   */
  get relationshipsNotRead(): RelationshipsNotReadWarning[] {
    const out: RelationshipsNotReadWarning[] = [];
    for (const rels of [this.relationshipsPart, ...[...this.parts].map((p) => p.relationshipsPart)]) {
      for (const w of rels?.warnings ?? []) if (w.code === 'part.not_read') out.push(w);
    }
    return out;
  }

  /**
   * What the load repaired (WD REQ-207's `repaired`, CR-031 section 2.4), read from what it recorded, so it reports what the
   * load did whatever the option: `'own'` for a repair of this package's own (an `External` main relationship dropped),
   * `'word'` where only Word's were followed (a duplicated entry, several main relationships, a relationships part read by
   * local name), else `'none'`.
   */
  get repaired(): Repaired {
    if (this.externalMainPart) return 'own';
    return this.duplicateEntries.length > 0 || this.multipleMainParts || this.relationshipsNamespaces.length > 0 ? 'word' : 'none';
  }

  /**
   * Every child a part holds that its container never admits (CR-016 D6, the check on request): what a save would
   * refuse for, over every typed XML part, each finding with its part, its path, the rule and docx4j's message naming
   * the fix. Nothing is changed, and an untouched part stays untouched (read privately). `[]` for a package whose
   * content is all in its place.
   */
  async checkContent(): Promise<MisplacedContentWarning[]> {
    const out: MisplacedContentWarning[] = [];
    for (const part of this.parts) if (part instanceof XmlPart) out.push(...(await part.checkContent()));
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
    copy.checkContentOnSave = this.checkContentOnSave;
    copy.checkLossOnSave = this.checkLossOnSave;
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
