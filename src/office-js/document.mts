// `context.document` and its properties (CR-002 section 3.10): Office JS's Document over a
// WordprocessingMLPackage. The class names are Office JS's, since `Word.run` callbacks are what
// this serves; the parts behind them keep docx4j's (DocPropsCorePart, DocumentSettingsPart).

import type * as core from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_core';
import type * as dc from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_core_dc_elements';
import type * as extended from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_extended';
import * as elDc from '@docx4j/generated-objects-ts/el/org_docx4j_docProps_core_dc_elements';
import { Docx4JException } from '../opc/exceptions.mjs';
import type { WordprocessingMLPackage } from '../packages/WordprocessingMLPackage.mjs';
import type { Body } from '../model/content/Body.mjs';
import type { Range } from '../model/content/Range.mjs';
import type { TrackedChange } from '../model/content/TrackedChange.mjs';
import type { ChangeTrackingModeValue } from './enums.mjs';
import { NotSupportedError } from './errors.mjs';

/** What `Word.run` was given: the selection to hand to `getSelection()`, and the preparation to do. */
export interface RunOptions {
  /** The range `context.document.getSelection()` returns; there is no host selection in Node. */
  selection?: Range;
  /**
   * Unmarshal the core and extended properties and the settings parts before the callback, so
   * that `document.properties` and `document.changeTrackingMode` read synchronously as Office JS
   * does (default true). False leaves those parts untouched, and reads of them then report
   * nothing; a write throws.
   */
  unmarshalSideParts?: boolean;
}

function literalText(value: dc.SimpleLiteral | undefined): string {
  return value?.content?.join('') ?? '';
}

function literal(text: string): dc.SimpleLiteral {
  return { content: [text] };
}

/** dcterms:W3CDTF, as docx4j writes dcterms:created and dcterms:modified. */
function w3cdtf(date: Date): dc.SimpleLiteral {
  return { TYPE_NAME: 'org_docx4j_docProps_core_dc_terms.W3CDTF', content: [date.toISOString().replace(/\.\d{3}Z$/, 'Z')] };
}

function dateOf(value: dc.SimpleLiteral | undefined): Date | undefined {
  const text = literalText(value);
  if (text === '') return undefined;
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? undefined : date;
}

function dateOfCalendar(value: core.XmlCalendar | undefined): Date | undefined {
  if (!value || value.year === undefined) return undefined;
  const ms = Date.UTC(value.year, (value.month ?? 1) - 1, value.day ?? 1, value.hour ?? 0, value.minute ?? 0, value.second ?? 0);
  return new Date(ms - (value.timezone ?? 0) * 60000);
}

function calendarOf(date: Date): core.XmlCalendar {
  return {
    year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate(),
    hour: date.getUTCHours(), minute: date.getUTCMinutes(), second: date.getUTCSeconds(), timezone: 0,
  };
}

/**
 * A subset of Office JS `Word.DocumentProperties` over `docProps/core.xml` (docx4j
 * DocPropsCorePart) and `docProps/app.xml` (DocPropsExtendedPart). Reads report '' or undefined
 * when a property is absent; a write creates the part when the package has none.
 */
export class DocumentProperties {
  constructor(private readonly pkg: WordprocessingMLPackage) {}

  private coreProps(create: boolean): core.CoreProperties | undefined {
    const part = create ? this.pkg.addDocPropsCorePart() : this.pkg.docPropsCorePart;
    if (!part) return undefined;
    if (!part.isUnmarshalled) {
      if (!create) return undefined;
      throw new Docx4JException(`${part.partName} is not unmarshalled; await its getContents() first (Word.run does that unless unmarshalSideParts is false)`);
    }
    return part.contents;
  }

  private extendedProps(create: boolean): extended.Properties | undefined {
    const part = create ? this.pkg.addDocPropsExtendedPart() : this.pkg.docPropsExtendedPart;
    if (!part) return undefined;
    if (!part.isUnmarshalled) {
      if (!create) return undefined;
      throw new Docx4JException(`${part.partName} is not unmarshalled; await its getContents() first (Word.run does that unless unmarshalSideParts is false)`);
    }
    return part.contents;
  }

  /** dc:title. */
  get title(): string {
    return literalText(this.coreProps(false)?.title?.value);
  }
  set title(value: string) {
    this.coreProps(true)!.title = elDc.title(literal(value));
  }

  /** dc:subject. */
  get subject(): string {
    return literalText(this.coreProps(false)?.subject);
  }
  set subject(value: string) {
    this.coreProps(true)!.subject = literal(value);
  }

  /** dc:creator (Office JS calls it the author). */
  get author(): string {
    return literalText(this.coreProps(false)?.creator);
  }
  set author(value: string) {
    this.coreProps(true)!.creator = literal(value);
  }

  /** dc:description (Office JS calls it comments). */
  get comments(): string {
    return literalText(this.coreProps(false)?.description?.value);
  }
  set comments(value: string) {
    this.coreProps(true)!.description = elDc.description(literal(value));
  }

  /** cp:keywords. */
  get keywords(): string {
    return this.coreProps(false)?.keywords ?? '';
  }
  set keywords(value: string) {
    this.coreProps(true)!.keywords = value;
  }

  /** cp:category. */
  get category(): string {
    return this.coreProps(false)?.category ?? '';
  }
  set category(value: string) {
    this.coreProps(true)!.category = value;
  }

  /** cp:lastModifiedBy. */
  get lastAuthor(): string {
    return this.coreProps(false)?.lastModifiedBy ?? '';
  }
  set lastAuthor(value: string) {
    this.coreProps(true)!.lastModifiedBy = value;
  }

  /** cp:revision. */
  get revisionNumber(): string {
    return this.coreProps(false)?.revision ?? '';
  }
  set revisionNumber(value: string) {
    this.coreProps(true)!.revision = value;
  }

  /** dcterms:created. */
  get creationDate(): Date | undefined {
    return dateOf(this.coreProps(false)?.created);
  }
  set creationDate(value: Date | undefined) {
    const props = this.coreProps(true)!;
    if (value === undefined) delete props.created; else props.created = w3cdtf(value);
  }

  /** dcterms:modified. */
  get lastSaveTime(): Date | undefined {
    return dateOf(this.coreProps(false)?.modified);
  }
  set lastSaveTime(value: Date | undefined) {
    const props = this.coreProps(true)!;
    if (value === undefined) delete props.modified; else props.modified = w3cdtf(value);
  }

  /** cp:lastPrinted. */
  get lastPrintDate(): Date | undefined {
    return dateOfCalendar(this.coreProps(false)?.lastPrinted);
  }
  set lastPrintDate(value: Date | undefined) {
    const props = this.coreProps(true)!;
    if (value === undefined) delete props.lastPrinted; else props.lastPrinted = calendarOf(value);
  }

  /** Manager (extended properties). */
  get manager(): string {
    return this.extendedProps(false)?.manager ?? '';
  }
  set manager(value: string) {
    this.extendedProps(true)!.manager = value;
  }

  /** Company (extended properties). */
  get company(): string {
    return this.extendedProps(false)?.company ?? '';
  }
  set company(value: string) {
    this.extendedProps(true)!.company = value;
  }

  /** Application (extended properties). */
  get applicationName(): string {
    return this.extendedProps(false)?.application ?? '';
  }
  set applicationName(value: string) {
    this.extendedProps(true)!.application = value;
  }

  /** Template (extended properties). */
  get template(): string {
    return this.extendedProps(false)?.template ?? '';
  }
  set template(value: string) {
    this.extendedProps(true)!.template = value;
  }

  /** DocSecurity (extended properties); read only, as Office JS. */
  get security(): number {
    return this.extendedProps(false)?.docSecurity ?? 0;
  }
}

/**
 * A subset of Office JS `Word.Document` over a package. `body` is the main document part's body;
 * everything else is the document-level surface add-in code reaches for first.
 */
export class Document {
  readonly properties: DocumentProperties;

  constructor(
    /** The package this document is (extension: Office JS has no package). */
    readonly package_: WordprocessingMLPackage,
    private readonly options: RunOptions = {},
  ) {
    this.properties = new DocumentProperties(package_);
  }

  /** The main document part's body; `Word.run` unmarshalled the part before the callback. */
  get body(): Body {
    return this.package_.body;
  }

  /** The caller's selection (`Word.run(pkg, fn, { selection })`); there is none in Node. */
  getSelection(): Range {
    const selection = this.options.selection;
    if (!selection) {
      throw new NotSupportedError('Document', 'getSelection', 'there is no host selection; pass one as Word.run(pkg, fn, { selection })');
    }
    return selection;
  }

  /**
   * Office JS `document.changeTrackingMode`: the package's, which is `w:trackRevisions` in the
   * settings part and, while it is on, makes every edit through the content API write revision
   * markup (CR-002 phase F). The shim adds nothing of its own; `Word.run` has unmarshalled the
   * settings part, so both directions are synchronous here.
   */
  get changeTrackingMode(): ChangeTrackingModeValue {
    return this.package_.changeTrackingMode;
  }

  set changeTrackingMode(mode: ChangeTrackingModeValue) {
    const settings = this.package_.getMainDocumentPart().documentSettingsPart;
    if (settings && !settings.isUnmarshalled) {
      throw new Docx4JException(`${settings.partName} is not unmarshalled; await its getContents() first (Word.run does that unless unmarshalSideParts is false)`);
    }
    this.package_.changeTrackingMode = mode;
  }

  /** Office JS `document.getTrackedChanges()`: the body's, in document order (CR-002 phase F). */
  getTrackedChanges(): TrackedChange[] {
    return this.body.getTrackedChanges();
  }

  /**
   * The comments of the document. CR-002 phase G puts `getComments()` on `Body`; until it lands
   * this reports none, so add-in code that lists comments runs either way.
   */
  getComments(): unknown[] | Promise<unknown[]> {
    const body = this.body as Body & { getComments?: () => unknown[] | Promise<unknown[]> };
    return typeof body.getComments === 'function' ? body.getComments() : [];
  }

  /** The whole document as a flat OPC `pkg:package` string, as Office JS's `body.getOoxml()`. */
  getOoxml(): Promise<string> {
    return this.package_.saveFlatOpc();
  }

  /** The document as a docx (extension; Office JS's `save()` saves in the host). */
  save(): Promise<Uint8Array> {
    return this.package_.save();
  }
}
