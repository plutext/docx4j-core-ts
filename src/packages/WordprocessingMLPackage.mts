import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { createCTCompat, createCTCompatSetting } from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { OpcPackage, type PackageSource } from './OpcPackage.mjs';
import { registerPackageClass } from './registry.mjs';
import type { LoadOptions } from '../opc/Load.mjs';
import { ContentTypes } from '../opc/ContentTypes.mjs';
import { Docx4JException } from '../opc/exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import type { Part } from '../parts/Part.mjs';
import { MainDocumentPart, StyleDefinitionsPart, DocumentSettingsPart, type NumberingDefinitionsPart } from '../parts/wml/index.mjs';
import { DocPropsCorePart, DocPropsExtendedPart } from '../parts/docProps/index.mjs';
import { DEFAULT_STYLES_XML } from '../parts/wml/defaultStyles.mjs';
import { ThemePart } from '../parts/dml/index.mjs';
import {
  newFontSettings, themeOfSettings, type FontSettings, type DefaultThemeValue,
} from '../model/fonts/defaultTheme.mjs';
import { HeaderPart, FooterPart } from '../parts/wml/index.mjs';
import type { Body, Address, Outline, OutlineParagraph, OutlineTable } from '../model/content/Body.mjs';
import type { Paragraph } from '../model/content/Paragraph.mjs';
import type { Author } from '../model/content/comments.mjs';
import { CustomXmlPartCollection } from '../model/customxml/CustomXmlPartCollection.mjs';
import { DefaultXPathEngine, type XPathEngine } from '../model/customxml/xpath.mjs';
import { ChangeTracker, type ChangeTrackingMode, type TrackingHost } from '../model/content/tracking.mjs';
import { PropertyResolver } from '../model/properties/PropertyResolver.mjs';
import type { Emulator } from '../model/listnumbering/Emulator.mjs';
import { PropertyResolverNotCreatedException } from '../opc/exceptions.mjs';
import type { TrackedChange } from '../model/content/TrackedChange.mjs';
import { XmlPart } from '../parts/XmlPart.mjs';
import type { PartSink } from '../opc/PartStore.mjs';
// Registers the comment parts with the content API (CR-002 phase G); no cycle: the model half
// never imports a part, as packages/registry.mts does for the package classes.
import '../parts/wml/comments.mjs';

/** The outline of a document: the body's paragraphs and tables, then each header's and footer's (loaded ones only). */
export interface DocumentOutline extends Outline {
  headers: { relId: string; partName: string; paragraphs: OutlineParagraph[]; tables: OutlineTable[] }[];
  footers: { relId: string; partName: string; paragraphs: OutlineParagraph[]; tables: OutlineTable[] }[];
}

const MAIN_CONTENT_TYPES = [
  ContentTypes.WORDPROCESSINGML_DOCUMENT,
  ContentTypes.WORDPROCESSINGML_DOCUMENT_MACROENABLED,
  ContentTypes.WORDPROCESSINGML_TEMPLATE,
  ContentTypes.WORDPROCESSINGML_TEMPLATE_MACROENABLED,
];

/** docx4j PageSizePaper. */
export type PageSizePaper = 'LETTER' | 'LEGAL' | 'A3' | 'A4' | 'A5' | 'B4JIS';

/** Portrait width and height in twips, and the `w:code`, as docx4j PageDimensions. */
const PAGE_SIZES: Record<PageSizePaper, [w: number, h: number, code: number]> = {
  LETTER: [12240, 15840, 1],
  LEGAL: [12240, 20160, 5],
  A3: [16839, 23814, 8],
  A4: [11907, 16839, 9],
  A5: [8391, 11907, 11],
  B4JIS: [14572, 20639, 12],
};

export interface CreatePackageOptions {
  pageSize?: PageSizePaper;
  landscape?: boolean;
  /** Which Office theme the new package's theme part carries, and which its font references
   *  resolve against: `'2023'` (the default, Word 365's Aptos Display / Aptos), `'2013'`
   *  (Calibri Light / Calibri) or `'2007'` (Cambria / Calibri). Sets `pkg.fonts.defaultTheme`. */
  defaultTheme?: DefaultThemeValue;
  /** The `w:compat/w:compatSetting` value for `compatibilityMode`: `'15'` (the default) is what
   *  Word 2013 and later write, and is what keeps Word 365 out of compatibility mode. `''`
   *  leaves the setting out, which is docx4j's own `createPackage` behaviour. */
  compatibilityMode?: string;
}

/** The URI of Word's own compatibility settings (docx4j `DocumentSettingsPart.setWordCompatSetting`). */
const WORD_COMPAT_URI = 'http://schemas.microsoft.com/office/word';

/**
 * The `w:compat/w:compatSetting`s a created document carries, in the order Word writes them.
 * `compatibilityMode` 15 is what Word 2013 and later write and is what keeps Word 365 out of
 * compatibility mode; `overrideTableStyleFontSizeAndJustification` is docx4j's own
 * (`DocumentSettingsPart.setOverrideTableStyleFontSizeAndJustification`). Word 365 writes four
 * more - `enableOpenTypeFeatures`, `doNotFlipMirrorIndents`, `differentiateMultirowTableHeaders`
 * and `useWord2013TrackBottomHyphenation`, all 1 - plus a locale-dependent `w:themeFontLang`,
 * which are deliberately left out for now (CR-001 section 17); adding them is one edit here.
 */
const DEFAULT_COMPAT_SETTINGS: ReadonlyArray<{ name: string; val: string }> = [
  // The six settings Word 365 writes for a new document, in its order; docx4j's
  // DocumentSettingsPart.setCompatSettingsAsWord365 (9de10aac9) writes the same six.
  { name: 'compatibilityMode', val: '15' },
  { name: 'overrideTableStyleFontSizeAndJustification', val: '1' },
  { name: 'enableOpenTypeFeatures', val: '1' },
  { name: 'doNotFlipMirrorIndents', val: '1' },
  { name: 'differentiateMultirowTableHeaders', val: '1' },
  { name: 'useWord2013TrackBottomHyphenation', val: '1' },
];

/** A docx (docx4j WordprocessingMLPackage). */
export class WordprocessingMLPackage extends OpcPackage implements TrackingHost {
  mainDocumentPart: MainDocumentPart | undefined;

  /**
   * Who this package's comments (CR-002 phase G) and its tracked changes (phase F) are by.
   * There is no signed-in user here, so the package carries the identity; the initials default to
   * the first letter of each word of the name and the email, when given, is written to `w:people`.
   */
  author: Author = { name: 'docx4j' };
  /** A fixed date for new revisions; undefined means the moment each one is written (CR-002 phase F). */
  trackedChangeDate: Date | undefined;
  private trackingMode: ChangeTrackingMode | undefined;
  private tracker: ChangeTracker | undefined;
  /** Set when `changeTrackingMode` was written but the settings part was not unmarshalled yet. */
  private trackingPending = false;

  /**
   * The package's font settings (CR-001 Phase B step 4), docx4j's `docx4j.fonts.*` properties
   * scoped to one package: `defaultTheme` says which Office theme a font reference resolves
   * against where the package has **no theme part**, and which theme part `createPackage()`
   * gives a new one. `'2023'` (Word 365's Aptos Display / Aptos) by default; `'2013'` is
   * Calibri Light / Calibri and `'2007'` Cambria / Calibri.
   *
   * Set it before the first `getRunFontSelector()` or `createPackage()`; afterwards call
   * `mainDocumentPart.refreshFonts()`.
   */
  readonly fonts: FontSettings = newFontSettings();

  static override async load(source: PackageSource, options?: LoadOptions): Promise<WordprocessingMLPackage> {
    const pkg = await OpcPackage.load(source, options);
    if (!(pkg instanceof WordprocessingMLPackage)) {
      throw new Docx4JException(`Not a WordprocessingML package: main part is ${pkg.getMainPart()?.contentType ?? 'missing'}`);
    }
    return pkg;
  }

  /**
   * A new document with a main document part (one section, the page size and 2.54 cm margins),
   * docx4j's default styles, settings, and core and extended properties (docx4j createPackage).
   */
  static async createPackage(options: CreatePackageOptions = {}): Promise<WordprocessingMLPackage> {
    const pkg = new WordprocessingMLPackage();
    if (options.defaultTheme !== undefined) pkg.fonts.defaultTheme = options.defaultTheme;
    const [pw, ph, code] = PAGE_SIZES[options.pageSize ?? 'A4'];
    const landscape = options.landscape === true;
    const pgSz: wml.SectPr.PgSz = { w: landscape ? ph : pw, h: landscape ? pw : ph, code };
    if (landscape) pgSz.orient = 'landscape';
    const pgMar: wml.SectPr.PgMar = { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 708, footer: 708, gutter: 0 };
    const document: wml.Document = { body: { content: [], sectPr: { pgSz, pgMar } } };
    const main = new MainDocumentPart();
    main.setContents(document);
    pkg.addTargetPart(main);

    const styles = new StyleDefinitionsPart();
    styles.setXml(DEFAULT_STYLES_XML);
    main.addTargetPart(styles);

    /* The settings part carries `w:compat/w:compatSetting compatibilityMode` 15, which Word 2013
     * and later write into every document they create: without it Word 365 opens the document in
     * compatibility mode ("Compatibility Mode" in the title bar, the pre-2013 layout rules).
     * docx4j writes no mode at all (its `getCompatibilityMode()` answers 12 when the setting is
     * absent, and says so), so this is a deliberate departure from it (CR-001 section 17);
     * `{ compatibilityMode: '' }` leaves the setting out and gets docx4j's behaviour back. */
    const settings = new DocumentSettingsPart();
    const compatibilityMode = options.compatibilityMode ?? '15';
    const compatSetting = DEFAULT_COMPAT_SETTINGS
      .filter((s) => s.name !== 'compatibilityMode' || compatibilityMode !== '')
      .map((s) => createCTCompatSetting({
        name: s.name, uri: WORD_COMPAT_URI, val: s.name === 'compatibilityMode' ? compatibilityMode : s.val,
      }));
    settings.setContents({ compat: createCTCompat({ compatSetting }) });
    main.addTargetPart(settings);

    const core = new DocPropsCorePart();
    core.setContents({});
    pkg.addTargetPart(core);
    const app = new DocPropsExtendedPart();
    app.setContents({});
    pkg.addTargetPart(app);

    /* The theme part (/word/theme/theme1.xml), as Word puts one in every document it creates
     * and as docx4j's createPackage has done since 17.1.1 (CR-001 section 14.6): without it,
     * the default styles' docDefaults - which reference the theme fonts
     * (w:rFonts w:asciiTheme="minorHAnsi" ...) - had nothing to resolve against, and every
     * consumer had to guess the Office theme's faces for itself. Which theme it is, is
     * `pkg.fonts.defaultTheme`; the same setting answers a themeless package's references. */
    const theme = new ThemePart();
    theme.setXml(themeOfSettings(pkg.fonts).xml);
    main.addTargetPart(theme);

    // so that `pkg.body`'s effective reads work without a further await: the default styles are
    // XML bytes, so this is the one place the resolver has to be built before the caller asks
    await pkg.getPropertyResolver();
    return pkg;
  }

  override setPartShortcut(part: Part, relationshipType: string): boolean {
    if (relationshipType === Namespaces.DOCUMENT || relationshipType === Namespaces.DOCUMENT_STRICT) {
      this.mainDocumentPart = part as MainDocumentPart;
      return true;
    }
    return super.setPartShortcut(part, relationshipType);
  }

  /** The main document part; throws if the package has none. */
  getMainDocumentPart(): MainDocumentPart {
    if (!this.mainDocumentPart) throw new Docx4JException('No main document part');
    return this.mainDocumentPart;
  }

  /** Office JS `document.body`: the main document part's body; the part must be unmarshalled (`getBody()` does that). */
  get body(): Body {
    return this.getMainDocumentPart().body;
  }

  async getBody(): Promise<Body> {
    return this.getMainDocumentPart().getBody();
  }

  // --- property resolution (CR-001 Phase B step 2) ----------------------------------------

  private resolver: PropertyResolver | undefined;
  private resolverPending: Promise<PropertyResolver> | undefined;

  /**
   * docx4j `MainDocumentPart.getPropertyResolver()`: the resolver of effective paragraph, run,
   * paragraph-mark and table properties, built once and kept for the life of the package.
   *
   * Asynchronous because it unmarshals the styles and numbering parts; `propertyResolver` is
   * the synchronous accessor afterwards. A style *added* to the styles part later is found; a
   * style *modified* or *removed* needs `refresh()`.
   */
  async getPropertyResolver(): Promise<PropertyResolver> {
    if (this.resolver) return this.resolver;
    this.resolverPending ??= PropertyResolver.create(this).then((resolver) => {
      this.resolver = resolver;
      this.resolverPending = undefined;
      return resolver;
    });
    return this.resolverPending;
  }

  /**
   * The resolver, which must already have been created: building it unmarshals parts, which
   * cannot be done synchronously. Throws `PropertyResolverNotCreatedException` naming
   * `getPropertyResolver()` otherwise.
   */
  get propertyResolver(): PropertyResolver {
    if (!this.resolver) throw new PropertyResolverNotCreatedException();
    return this.resolver;
  }

  /** The resolver if it has been created, without throwing (what a read with a fallback wants). */
  get propertyResolverOrUndefined(): PropertyResolver | undefined {
    return this.resolver;
  }

  /** Tells the resolver, when there is one, that the styles part has changed. */
  refreshPropertyResolver(): void {
    this.resolver?.refresh();
    this.mainDocumentPart?.numberingDefinitionsPart?.refreshDefinitions();
  }

  /**
   * Builds the resolver and the numbering definitions again, over the parts the package holds
   * **now**: what docx4j's `getPropertyResolver(true)` does, and what a styles or numbering part
   * *added* after the first resolver was built needs (`refreshPropertyResolver()` re-reads the
   * parts the resolver already knows about, but cannot find new ones).
   */
  async refresh(): Promise<PropertyResolver> {
    this.resolver = undefined;
    this.resolverPending = undefined;
    this.mainDocumentPart?.numberingDefinitionsPart?.refreshDefinitions();
    return this.getPropertyResolver();
  }

  // --- list numbering (CR-001 Phase B step 3) --------------------------------------------

  /** `/word/numbering.xml`, or undefined. docx4j `MainDocumentPart.getNumberingDefinitionsPart()`. */
  get numberingDefinitionsPart(): NumberingDefinitionsPart | undefined {
    return this.mainDocumentPart?.numberingDefinitionsPart;
  }

  /**
   * The numbering emulator over this package's numbering part, with its definitions read
   * (docx4j `NumberingDefinitionsPart.getEmulator()`), or undefined where there is no numbering
   * part. Asynchronous because reading the part is; `numberingDefinitionsPart.getEmulator()` is
   * the synchronous accessor afterwards.
   */
  async getNumberingEmulator(): Promise<Emulator | undefined> {
    const part = this.numberingDefinitionsPart;
    if (part === undefined) return undefined;
    await this.getPropertyResolver();
    await part.getDefinitions();
    return part.getEmulator();
  }

  // --- change tracking (CR-002 phase F, section 3.7) --------------------------------------

  /**
   * Office JS `document.changeTrackingMode`, backed by `w:trackRevisions` in the settings part.
   * While it is not `Off`, every mutation of the content API writes Word's revision markup with
   * `author` and `trackedChangeDate` instead of editing in place.
   *
   * Reading is synchronous, so it reports `Off` for a loaded package whose settings part has not
   * been unmarshalled; `getChangeTrackingMode()` unmarshals it and is the one to await. Writing
   * is synchronous too and is written through to the settings part on save when the part is not
   * unmarshalled yet (`setChangeTrackingMode()` does it there and then).
   */
  get changeTrackingMode(): ChangeTrackingMode {
    if (this.trackingMode !== undefined) return this.trackingMode;
    const settings = this.mainDocumentPart?.documentSettingsPart;
    if (settings?.isUnmarshalled) return (this.trackingMode = modeOf(settings.contents));
    return 'Off';
  }

  set changeTrackingMode(mode: ChangeTrackingMode) {
    this.trackingMode = mode;
    this.tracker = undefined;
    const settings = this.mainDocumentPart ? this.settingsPart() : undefined;
    if (settings?.isUnmarshalled) { writeTrackRevisions(settings.contents, mode); this.trackingPending = false; }
    else this.trackingPending = true;
  }

  /** The mode as the settings part has it, unmarshalling the part (docx4j's async counterpart of the getter). */
  async getChangeTrackingMode(): Promise<ChangeTrackingMode> {
    const settings = this.mainDocumentPart?.documentSettingsPart;
    if (!settings) return this.trackingMode ?? 'Off';
    return (this.trackingMode = modeOf(await settings.getContents()));
  }

  /** Writes `w:trackRevisions`, unmarshalling the settings part, or creating it when there is none. */
  async setChangeTrackingMode(mode: ChangeTrackingMode): Promise<void> {
    this.trackingMode = mode;
    this.tracker = undefined;
    this.trackingPending = false;
    writeTrackRevisions(await this.settingsPart().getContents(), mode);
  }

  /** The settings part, created (with its relationship and content type) when the document has none. */
  private settingsPart(): DocumentSettingsPart {
    const main = this.getMainDocumentPart();
    let settings = main.documentSettingsPart;
    if (!settings) {
      settings = new DocumentSettingsPart();
      settings.setContents({});
      main.addTargetPart(settings);
    }
    return settings;
  }

  /**
   * Office JS `document.getTrackedChanges()`: every tracked change in the main document part,
   * in document order (headers and footers have their own `Body`). The part must be
   * unmarshalled, as `body` needs it to be.
   */
  getTrackedChanges(): TrackedChange[] {
    return this.body.getTrackedChanges();
  }

  /** The tracker every content-API mutation asks for; undefined while the mode is `Off`. */
  get changeTracker(): ChangeTracker | undefined {
    const mode = this.changeTrackingMode;
    if (mode === 'Off') return undefined;
    if (!this.tracker || this.tracker.mode !== mode) this.tracker = new ChangeTracker(this, mode);
    return this.tracker;
  }

  /** TrackingHost: the trees a new revision id must be above, which is every part already unmarshalled. */
  markupRoots(): object[] {
    const out: object[] = [];
    for (const part of this.parts) {
      if (part instanceof XmlPart && part.isUnmarshalled) {
        const contents = part.contents;
        if (typeof contents === 'object' && contents !== null) out.push(contents);
      }
    }
    return out;
  }

  /** A clone carries the identity, the tracking date, the theme setting and the XPath engine. */
  protected override copyPackageSettingsTo(target: OpcPackage): void {
    if (!(target instanceof WordprocessingMLPackage)) return;
    target.author = { ...this.author };
    target.trackedChangeDate = this.trackedChangeDate;
    target.fonts.defaultTheme = this.fonts.defaultTheme;
    target.xpathEngine = this.xpathEngine;
  }

  override async saveTo<R>(sink: PartSink<R>): Promise<R> {
    if (this.trackingPending) await this.setChangeTrackingMode(this.trackingMode ?? 'Off');
    return super.saveTo(sink);
  }

  /** The paragraph at an address anywhere in the document: 'body/3', 'header:rId5/0', a paraId, or a text match. */
  async paragraphAt(address: Address): Promise<Paragraph | undefined> {
    for (const body of await this.bodies(address)) {
      const p = body.paragraphAt(address);
      if (p) return p;
    }
    return undefined;
  }

  /** Every paragraph and table with its address, style and text (CR-002 section 3.3). Unmarshals the main part, headers and footers. */
  async outline(): Promise<DocumentOutline> {
    const main = await this.getBody();
    const out: DocumentOutline = { ...main.outline(), headers: [], footers: [] };
    for (const part of this.getMainDocumentPart().headerParts) {
      const body = await part.getBody();
      out.headers.push({ relId: part.sourceRelationship?.id ?? '', partName: part.partName.name, ...body.outline() });
    }
    for (const part of this.getMainDocumentPart().footerParts) {
      const body = await part.getBody();
      out.footers.push({ relId: part.sourceRelationship?.id ?? '', partName: part.partName.name, ...body.outline() });
    }
    return out;
  }

  private async bodies(address: Address): Promise<Body[]> {
    const main = this.getMainDocumentPart();
    if (typeof address === 'string' && !address.startsWith('w14:')) {
      const prefix = address.split('/')[0]!;
      if (prefix === 'body') return [await main.getBody()];
      const [kind, relId] = prefix.split(':');
      const parts = kind === 'header' ? main.headerParts : kind === 'footer' ? main.footerParts : [];
      const part = parts.find((p) => p.sourceRelationship?.id === relId || p.partName.fileName === relId);
      return part instanceof HeaderPart || part instanceof FooterPart ? [await part.getBody()] : [];
    }
    const all: Body[] = [await main.getBody()];
    for (const part of [...main.headerParts, ...main.footerParts]) all.push(await part.getBody());
    return all;
  }

  // --- custom XML (CR-002 phase E) ---

  /**
   * The XPath engine the custom XML views use (CR-002 section 3.6): `document.evaluate` in a
   * browser or an add-in, the optional `xpath` package in Node. Settable, so a consumer with
   * another DOM (or one who wants no dynamic import) plugs in its own. Warmed once by
   * `customXmlParts.load()`; every node call is synchronous afterwards.
   */
  xpathEngine: XPathEngine = new DefaultXPathEngine();

  /**
   * Office JS `document.customXmlParts`: the custom XML data storage parts as views, with
   * `add`, and docx4j's `applyBindings()` / `updateFromContentControls()` (CR-002 section 3.5).
   */
  get customXmlParts(): CustomXmlPartCollection {
    const pkg = this;
    return (this.customXmlPartCollection ??= new CustomXmlPartCollection({
      // a getter, so that setting pkg.xpathEngine later is seen by the collection
      get xpathEngine(): XPathEngine { return pkg.xpathEngine; },
      customXmlDataStorageParts: this.customXmlDataStorageParts,
      getBoundBodies: async () => {
        const main = this.getMainDocumentPart();
        const bodies = [await main.getBody()];
        for (const part of [...main.headerParts, ...main.footerParts]) bodies.push(await part.getBody());
        return bodies;
      },
      boundBodies: () => {
        const main = this.mainDocumentPart;
        if (!main?.isUnmarshalled) return [];
        const bodies = [main.body];
        for (const part of [...main.headerParts, ...main.footerParts]) if (part.isUnmarshalled) bodies.push(part.body);
        return bodies;
      },
      customXmlRelationshipSource: () => this.getMainDocumentPart(),
    }));
  }
  private customXmlPartCollection: CustomXmlPartCollection | undefined;

  protected override get progId(): string {
    return 'Word.Document';
  }
}

/**
 * `w:trackRevisions` is a flag, so a file cannot distinguish `TrackAll` from `TrackMineOnly`;
 * a document that has it on reads as `TrackAll`, and either value writes it.
 */
function modeOf(settings: wml.CTSettings): ChangeTrackingMode {
  const flag = settings.trackRevisions;
  return flag === undefined || flag.val === false ? 'Off' : 'TrackAll';
}

function writeTrackRevisions(settings: wml.CTSettings, mode: ChangeTrackingMode): void {
  if (mode === 'Off') delete settings.trackRevisions; else settings.trackRevisions = {};
}

registerPackageClass(MAIN_CONTENT_TYPES, WordprocessingMLPackage);
