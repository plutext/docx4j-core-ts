import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import type { OpcPackage } from '../../packages/OpcPackage.mjs';
import { type Element, type TextViewOptions, typeNameOf, childrenOf, linkParents, BLOCK_LEVEL_TYPES, textOf, textOfView, rowsOf, cellsOf } from './tree.mjs';
import { r as textRun, br as breakItem, tbl as tableOf } from '@docx4j/generated-objects-ts/builders/wml';
import { runOf, paragraphOf } from './tree.mjs';
import { Paragraph } from './Paragraph.mjs';
import { Range } from './Range.mjs';
import { Table } from './Table.mjs';
import { ContentControl, collectControls, type ContentControlType } from './ContentControl.mjs';
import { sdtBlockFor, nextControlId, checkKind } from '../customxml/insert.mjs';
import { InlinePicture, addImage, writableWidthEmu, type InlinePictureOptions } from './InlinePicture.mjs';
import { contentOf } from './ooxml.mjs';
import type { SearchOptions } from './search.mjs';
import { commentApi } from './comments.mjs';
import type { Comment } from './Comment.mjs';
import { type ChangeTracker, trackerOf, trackInsertedParagraph, trackInsertedTable } from './tracking.mjs';
import { type TrackedChange, trackedChangesOfRow } from './TrackedChange.mjs';

/** Where a paragraph or table is, for agents and across tool calls (CR-002 section 3.3). */
export type Address = string | { contains: string } | { paraId: string };

/** A view over a block-level element that is not a paragraph (a table, a content control), for insert targets and outlines. */
export interface BlockElement {
  element: Element;
  container: Element[];
}

export interface OutlineParagraph {
  address: string;
  paraId?: string;
  style: string;
  text: string;
}

export interface OutlineTable {
  address: string;
  rows: number;
  cols: number;
}

export interface Outline {
  paragraphs: OutlineParagraph[];
  tables: OutlineTable[];
}

/**
 * A subset of Office JS `Word.Body` over any container with block-level content: the main
 * document's body, a header, a footer, a footnote, a cell or a content control. A view: the
 * tree is the state; `content` is the live array.
 */
export class Body {
  constructor(
    /** The part whose tree this is; marks it for re-marshalling by being unmarshalled. */
    readonly part: XmlPart<unknown> | undefined,
    /** The container value: a Body, Hdr, Ftr, Tc, SdtContentBlock, ... */
    readonly container: { content?: Element[] },
    /** The address prefix: 'body', 'header:rId3', 'footer:rId4', 'footnote:2'. */
    readonly prefix: string = 'body',
    readonly package_?: OpcPackage,
  ) {}

  /** The live content array (docx4j getContent()); block-level elements only, sectPr excluded. */
  get content(): Element[] {
    return (this.container.content ??= []);
  }

  /** Every paragraph in document order, descending into tables and content controls. */
  get paragraphs(): Paragraph[] {
    const out: Paragraph[] = [];
    const visit = (items: Element[]): void => {
      for (const el of items) {
        const tn = typeNameOf(el);
        if (tn === 'org_docx4j_wml.P') out.push(new Paragraph(el as Element<wml.P>, items, this));
        else {
          const v = el.value;
          if (typeof v !== 'object' || v === null) continue;
          if (tn === 'org_docx4j_wml.Tbl') {
            for (const row of rowsOf(v)) for (const cell of cellsOf(row.element.value)) visit(childrenOf(cell.element.value) ?? []);
          } else {
            const children = childrenOf(v);
            if (children) visit(children);
          }
        }
      }
    };
    visit(this.content);
    return out;
  }

  /** The tables directly in this body, as Office JS reports them (nested tables are a cell's). */
  get tables(): Table[] {
    return this.content.filter((el) => typeNameOf(el) === 'org_docx4j_wml.Tbl').map((element) => new Table(element as Element<wml.Tbl>, this.content, this));
  }

  /** Every content control in this body, in document order, nested ones included. */
  get contentControls(): ContentControl[] {
    const out: ContentControl[] = [];
    collectControls(this.content, this, out);
    return out;
  }

  /** Every inline picture in this body, in document order (Office JS Body.inlinePictures). */
  get inlinePictures(): InlinePicture[] {
    return this.paragraphs.flatMap((p) => p.inlinePictures);
  }

  /**
   * A view over a container nested in this body (a cell, a content control), sharing this body's
   * part and package. The address prefix is the container's own address, so that the addresses a
   * nested body reports carry on from this one.
   */
  sub(container: { content?: Element[] }, prefix: string = this.prefix): Body {
    return new Body(this.part, container, prefix, this.package_);
  }

  /** The view of a `w:p` somewhere in this body's tree; its container comes from PARENT. */
  paragraphFor(p: wml.P): Paragraph | undefined {
    const container = childrenOf(((p as { PARENT?: object }).PARENT ?? {}) as object);
    const element = container?.find((e) => e.value === p) as Element<wml.P> | undefined;
    return element && container ? new Paragraph(element, container, this) : undefined;
  }

  /** The text, a paragraph per line. */
  get text(): string {
    return this.paragraphs.map((p) => p.text).join('\n');
  }

  /**
   * The text in one of the two views (extension; `text` is the accepted one).
   * `{ view: 'original' }` reads the document as it was before the tracked changes.
   */
  getText(options?: TextViewOptions): string {
    if (options?.view !== 'original') return this.text;
    return this.paragraphs.map((p) => textOfView(p.p, options)).join('\n');
  }

  /** The package's change tracker while `changeTrackingMode` is on; undefined when it is off (CR-002 phase F). */
  get changeTracker(): ChangeTracker | undefined {
    return trackerOf(this.package_);
  }

  insertParagraph(text: string, location: 'Start' | 'End'): Paragraph {
    const el = paragraphOf(text === '' ? [] : [textRun(text)]);
    return this.insertElement(el, location) as Paragraph;
  }

  /** Text at the start of the first paragraph, the end of the last, or replacing everything with one paragraph. */
  insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range {
    const paragraphs = this.paragraphs;
    if (location === 'Replace' || paragraphs.length === 0) {
      this.clear();
      const p = this.insertParagraph(text, 'End');
      return p.getRange('Content');
    }
    if (location === 'Start') return paragraphs[0]!.insertText(text, 'Start');
    return paragraphs[paragraphs.length - 1]!.insertText(text, 'End');
  }

  /**
   * A table of `rowCount` by `columnCount` cells at the start or end (Office JS insertTable);
   * `values` fills it row by row. The columns are equal over the section's text width.
   */
  insertTable(rowCount: number, columnCount: number, location: 'Start' | 'End', values?: string[][]): Table {
    if (rowCount < 1 || columnCount < 1) throw new Docx4JException(`A table needs at least one row and one column; got ${rowCount} by ${columnCount}`);
    const rows = Array.from({ length: rowCount }, (_, r) => Array.from({ length: columnCount }, (_, c) => values?.[r]?.[c] ?? ''));
    const width = writableWidthTwips(this.container as { sectPr?: wml.SectPr });
    const element = tableOf(rows, width ? { width } : {}) as Element<wml.Tbl>;
    this.insertElement(element as Element, location);
    return new Table(element, this.content, this);
  }

  /**
   * Adds the image as a part of this body's part, with a relationship, and shows it in a new
   * paragraph at the start or the end (Office JS insertInlinePictureFromBase64). An image wider
   * than the text area is scaled down, as docx4j's CxCy.scale does.
   */
  insertInlinePictureFromBase64(base64: string, location: 'Start' | 'End', options?: InlinePictureOptions): InlinePicture {
    if (!this.part) throw new Docx4JException('This body has no part, so an image cannot be related to it');
    const { run, drawing } = addImage(this.part, base64, this.container, writableWidthEmu(this.container), options);
    const paragraph = this.insertParagraph('', location);
    paragraph.p.content!.push(run as never);
    linkParents(run, paragraph.p);
    return new InlinePicture(drawing, run, paragraph);
  }

  insertBreak(type: 'Page' | 'Line', location: 'Start' | 'End'): void {
    const p = this.insertParagraph('', location);
    const run = runOf([breakItem(type === 'Page' ? 'page' : undefined)]);
    const tracker = this.changeTracker;
    const item = tracker ? tracker.ins([run]) : run;
    p.p.content!.push(item as never);
    linkParents(item, p.p);
  }

  /**
   * Inserts block-level elements (docx4j addObject): at the start or end of this body, or
   * before or after a target paragraph, element or address. Links PARENT. Rejects anything the
   * body cannot hold, naming what was passed. Returns the view of the first inserted element.
   */
  insertElement(element: Element | Element[], location: 'Start' | 'End' | 'Before' | 'After', target?: Paragraph | BlockElement | Address): Paragraph | BlockElement {
    const elements = Array.isArray(element) ? element : [element];
    if (elements.length === 0) throw new Docx4JException('Nothing to insert');
    for (const el of elements) {
      const tn = typeNameOf(el);
      if (tn === undefined || !BLOCK_LEVEL_TYPES.has(tn)) {
        throw new Docx4JException(`A body cannot hold ${describe(el)}; block-level elements only (w:p, w:tbl, w:sdt, ...)`);
      }
    }
    let container = this.content;
    let index: number;
    if (location === 'Start') index = 0;
    else if (location === 'End') index = container.length;
    else {
      const t = this.resolveTarget(target);
      container = t.container;
      index = container.indexOf(t.element) + (location === 'After' ? 1 : 0);
    }
    container.splice(index, 0, ...elements);
    const owner = this.ownerOf(container);
    const tracker = this.changeTracker;
    for (const el of elements) {
      linkParents(el, owner);
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.P') {
        this.assignParaId(el.value as wml.P);
        if (tracker) trackInsertedParagraph(tracker, el.value as wml.P);
      } else if (tn === 'org_docx4j_wml.Tbl' && tracker) {
        trackInsertedTable(tracker, el.value as wml.Tbl);
      }
    }
    const first = elements[0]!;
    return typeNameOf(first) === 'org_docx4j_wml.P' ? new Paragraph(first as Element<wml.P>, container, this) : { element: first, container };
  }

  /** Inserts a WML fragment (one or more w:p / w:tbl / ... as inside document.xml); the standard prefixes are declared for it. */
  async insertXml(xml: string, location: 'Start' | 'End' | 'Before' | 'After', target?: Paragraph | BlockElement | Address): Promise<(Paragraph | BlockElement)[]> {
    return this.insertContent(await contentOf(xml, { preprocess: this.preprocessor() }), location, target);
  }

  /**
   * Word's `insertOoxml`: a flat OPC `pkg:package` string, whose body content is inserted with
   * the parts it references (images, embedded objects) copied into this package under fresh
   * names and relationship ids; styles and numbering are not merged (CR-002 open question 5).
   * A bare `w:p` / `w:tbl` fragment is accepted too, as `insertXml` takes. Returns the views of
   * what was inserted, as `insertXml` does, since a package may bring several blocks.
   */
  async insertOoxml(ooxml: string, location: 'Start' | 'End' | 'Replace' | 'Before' | 'After', target?: Paragraph | BlockElement | Address): Promise<(Paragraph | BlockElement)[]> {
    const elements = await contentOf(ooxml, { preprocess: this.preprocessor(), target: this.part });
    if (location === 'Replace') {
      this.clear();
      return this.insertContent(elements, 'End');
    }
    return this.insertContent(elements, location, target);
  }

  /** The package's DOM preprocessor (the MCE one), for fragments and incoming packages. */
  private preprocessor(): ((doc: Document) => void) | undefined {
    const preprocess = this.package_?.loadOptions.preprocessor;
    return preprocess ? (doc: Document) => preprocess(doc) : undefined;
  }

  private insertContent(elements: Element[], location: 'Start' | 'End' | 'Before' | 'After', target?: Paragraph | BlockElement | Address): (Paragraph | BlockElement)[] {
    if (elements.length === 0) return [];
    this.insertElement(elements, location, target);
    return elements.map((el) => typeNameOf(el) === 'org_docx4j_wml.P'
      ? new Paragraph(el as Element<wml.P>, this.containerOf(el), this)
      : { element: el, container: this.containerOf(el) });
  }

  /** Matches over every paragraph, in document order. */
  search(text: string, options?: SearchOptions): Range[] {
    return this.paragraphs.flatMap((p) => p.search(text, options));
  }

  /**
   * Replaces every match of `find` with `replace`, last match first so that the offsets of the
   * ones still to do stay valid (CR-002 section 3.7). Returns the number replaced; tracked when
   * the package's `changeTrackingMode` is on.
   */
  replaceText(find: string, replace: string, options?: SearchOptions): number {
    const matches = this.search(find, options);
    for (let i = matches.length - 1; i >= 0; i--) matches[i]!.insertText(replace, 'Replace');
    return matches.length;
  }

  /**
   * Removes everything (section properties stay). While the package tracks changes, every
   * paragraph is marked deleted and every row's `w:trPr` takes a `w:del` instead.
   */
  clear(): void {
    const tracker = this.changeTracker;
    if (!tracker) { this.content.length = 0; return; }
    for (const row of this.rowElements()) tracker.markRowDeleted(row.value as wml.Tr);
    for (const p of this.paragraphs) p.delete();
  }

  // --- change tracking (CR-002 phase F) ---------------------------------------------------

  /** Every tracked change in this body, in document order. */
  getTrackedChanges(): TrackedChange[] {
    const out: TrackedChange[] = [];
    const visit = (items: Element[]): void => {
      for (const el of items) {
        const tn = typeNameOf(el);
        if (tn === 'org_docx4j_wml.P') { out.push(...new Paragraph(el as Element<wml.P>, items, this).getTrackedChanges()); continue; }
        const v = el.value;
        if (typeof v !== 'object' || v === null) continue;
        if (tn === 'org_docx4j_wml.Tbl') {
          for (const row of rowsOf(v)) {
            out.push(...trackedChangesOfRow(row.element, row.container));
            for (const cell of cellsOf(row.element.value)) visit(childrenOf(cell.element.value) ?? []);
          }
          continue;
        }
        const children = childrenOf(v);
        if (children) visit(children);
      }
    };
    visit(this.content);
    return out;
  }

  /**
   * Accepts every tracked change, last first so that a paragraph join never disturbs one still
   * to do (docx4j `AcceptTrackedChanges` is the reference). Returns how many were accepted.
   */
  acceptAll(): number {
    const changes = this.getTrackedChanges();
    for (let i = changes.length - 1; i >= 0; i--) changes[i]!.accept();
    return changes.length;
  }

  /** Rejects every tracked change, last first. Returns how many were rejected. */
  rejectAll(): number {
    const changes = this.getTrackedChanges();
    for (let i = changes.length - 1; i >= 0; i--) changes[i]!.reject();
    return changes.length;
  }

  /** Every `w:tr` in this body's tables, in document order (rows of nested tables included). */
  private rowElements(): Element<wml.Tr>[] {
    const out: Element<wml.Tr>[] = [];
    const visit = (items: Element[]): void => {
      for (const el of items) {
        const v = el.value;
        if (typeof v !== 'object' || v === null) continue;
        if (typeNameOf(el) === 'org_docx4j_wml.Tbl') {
          for (const row of rowsOf(v)) {
            out.push(row.element);
            for (const cell of cellsOf(row.element.value)) visit(childrenOf(cell.element.value) ?? []);
          }
          continue;
        }
        const children = childrenOf(v);
        if (children) visit(children);
      }
    };
    visit(this.content);
    return out;
  }

  /** The part's XML (extension; Office JS getOoxml wraps it in a package). */
  async getXml(): Promise<string> {
    if (!this.part) throw new Docx4JException('This body has no part');
    return this.part.getXml();
  }

  // --- addresses (CR-002 section 3.3) ---

  /** The element at an address: 'body/3' (fourth block), 'body/3/1/0/2' (table row 1, cell 0, block 2), a paraId, or a text match. */
  elementAt(address: Address): BlockElement | undefined {
    if (typeof address === 'string' && address.startsWith('w14:')) address = { paraId: address.substring(4) };
    if (typeof address === 'object') {
      const p = 'paraId' in address
        ? this.paragraphs.find((q) => q.paraId?.toLowerCase() === address.paraId.toLowerCase())
        : this.paragraphs.find((q) => q.text.includes(address.contains));
      return p ? { element: p.element, container: p.container } : undefined;
    }
    // The prefix may itself hold slashes: a cell's body is addressed 'body/4/0/1'.
    if (!address.startsWith(this.prefix)) return undefined;
    const tail = address.substring(this.prefix.length);
    if (tail !== '' && !tail.startsWith('/')) return undefined;
    let current: object = this.container;
    let container: Element[] | undefined;
    let element: Element | undefined;
    for (const seg of tail === '' ? [] : tail.substring(1).split('/')) {
      container = childrenOf(current);
      const i = Number(seg);
      if (!container || !Number.isInteger(i) || i < 0 || i >= container.length) return undefined;
      element = container[i]!;
      current = element.value as object;
    }
    return element && container ? { element, container } : undefined;
  }

  paragraphAt(address: Address): Paragraph | undefined {
    const found = this.elementAt(address);
    if (!found || typeNameOf(found.element) !== 'org_docx4j_wml.P') return undefined;
    return new Paragraph(found.element as Element<wml.P>, found.container, this);
  }

  /** The ordinal address of a paragraph or element in this body, from PARENT links; undefined when it is not here. */
  addressOf(target: Paragraph | BlockElement): string | undefined {
    const path = this.pathOf(target.element.value as object);
    return path ? [this.prefix, ...path].join('/') : undefined;
  }

  /** What an agent asks for first: every paragraph and table with its address, style and text. */
  outline(): Outline {
    const paragraphs: OutlineParagraph[] = [];
    const tables: OutlineTable[] = [];
    const visit = (items: Element[], path: number[]): void => {
      items.forEach((el, i) => {
        const tn = typeNameOf(el);
        const here = [...path, i];
        const address = [this.prefix, ...here].join('/');
        if (tn === 'org_docx4j_wml.P') {
          const p = el.value as wml.P;
          const entry: OutlineParagraph = { address, style: p.pPr?.pStyle?.val ?? 'Normal', text: textOf(p) };
          if (p.paraId) entry.paraId = p.paraId;
          paragraphs.push(entry);
        } else if (tn === 'org_docx4j_wml.Tbl') {
          const rows = childrenOf(el.value as object) ?? [];
          const cols = rows.length ? (childrenOf(rows[0]!.value as object) ?? []).filter((c) => typeNameOf(c) === 'org_docx4j_wml.Tc').length : 0;
          tables.push({ address, rows: rows.filter((r) => typeNameOf(r) === 'org_docx4j_wml.Tr').length, cols });
          rows.forEach((row, ri) => (childrenOf(row.value as object) ?? []).forEach((cell, ci) => visit(childrenOf(cell.value as object) ?? [], [...here, ri, ci])));
        } else if (typeof el.value === 'object' && el.value !== null) {
          const children = childrenOf(el.value);
          if (children) visit(children, here);
        }
      });
    };
    visit(this.content, []);
    return { paragraphs, tables };
  }

  // --- docx4j names ---

  /** docx4j MainDocumentPart.addParagraphOfText. */
  addParagraphOfText(text: string): wml.P {
    return this.insertParagraph(text, 'End').p;
  }

  /** docx4j MainDocumentPart.addStyledParagraphOfText. */
  addStyledParagraphOfText(styleId: string, text: string): wml.P {
    const p = this.insertParagraph(text, 'End');
    p.style = styleId;
    return p.p;
  }

  /** docx4j addObject: appends a block-level element. */
  addObject(element: Element): void {
    this.insertElement(element, 'End');
  }

  /** docx4j getContent(). */
  getContent(): Element[] {
    return this.content;
  }

  // --- internals ---

  private resolveTarget(target: Paragraph | BlockElement | Address | undefined): BlockElement {
    if (target === undefined) throw new Docx4JException("'Before' and 'After' need a target");
    if (target instanceof Paragraph) return { element: target.element, container: target.container };
    if (typeof target === 'object' && 'element' in target) return target;
    const found = this.elementAt(target);
    if (!found) throw new Docx4JException(`No element at ${typeof target === 'string' ? target : JSON.stringify(target)}`);
    return found;
  }

  /** The typed object whose content array this is, for PARENT links. */
  private ownerOf(container: Element[]): object {
    if (container === this.container.content) return this.container;
    let owner: object = this.container;
    const seek = (value: object): boolean => {
      const children = childrenOf(value);
      if (!children) return false;
      if (children === container) { owner = value; return true; }
      for (const el of children) {
        const v = el.value;
        if (typeof v !== 'object' || v === null) continue;
        if (typeNameOf(el) === 'org_docx4j_wml.Tbl') {
          for (const row of rowsOf(v)) for (const cell of cellsOf(row.element.value)) if (seek(cell.element.value)) return true;
        } else if (seek(v)) return true;
      }
      return false;
    };
    seek(this.container);
    // SdtBlock: the content array belongs to its sdtContent
    const o = owner as { TYPE_NAME?: string; sdtContent?: object };
    return o.TYPE_NAME === 'org_docx4j_wml.SdtBlock' && o.sdtContent ? o.sdtContent : owner;
  }

  private containerOf(element: Element): Element[] {
    let found: Element[] = this.content;
    const seek = (value: object): boolean => {
      const children = childrenOf(value);
      if (!children) return false;
      if (children.includes(element)) { found = children; return true; }
      for (const el of children) {
        const v = el.value;
        if (typeof v !== 'object' || v === null) continue;
        if (typeNameOf(el) === 'org_docx4j_wml.Tbl') {
          for (const row of rowsOf(v)) for (const cell of cellsOf(row.element.value)) if (seek(cell.element.value)) return true;
        } else if (seek(v)) return true;
      }
      return false;
    };
    seek(this.container);
    return found;
  }

  private pathOf(value: object): number[] | undefined {
    const path: number[] = [];
    let child: object = value;
    for (let guard = 0; guard < 64; guard++) {
      let parent = (child as { PARENT?: object }).PARENT;
      if (!parent) return undefined;
      // sdtContent is not a level of its own, at block, row, cell or run level
      if (/SdtContent(Block|Row|Cell|Run)$/.test((parent as { TYPE_NAME?: string }).TYPE_NAME ?? '')) parent = (parent as { PARENT?: object }).PARENT;
      if (!parent) return undefined;
      const children = childrenOf(parent);
      if (!children) return undefined;
      const i = children.findIndex((el) => el.value === child);
      if (i < 0) return undefined;
      path.unshift(i);
      if (parent === this.container) return path;
      child = parent;
    }
    return undefined;
  }

  /** Gives a new paragraph a w14:paraId when the document already uses them. */
  private assignParaId(p: wml.P): void {
    if (p.paraId) return;
    if (!this.usesParaIds()) return;
    let id: string;
    do {
      id = Math.floor(Math.random() * 0x7fffffff).toString(16).toUpperCase().padStart(8, '0');
    } while (this.paragraphs.some((q) => q.paraId === id));
    p.paraId = id;
  }

  private usesParaIds(): boolean {
    return this.paragraphs.some((q) => q.paraId !== undefined);
  }

  // --- comments (CR-002 phase G) ---

  /**
   * The comments anchored in this body, in document order, with replies nested under their
   * parent (only top-level comments are in the array). Asynchronous because it unmarshals the
   * comment parts.
   */
  async getComments(): Promise<Comment[]> {
    return commentApi().commentsOf(this);
  }

  // --- content controls (CR-002 phase E) ---

  /**
   * Wraps everything this body holds in a new content control (Office JS
   * `Body.insertContentControl`), typed for the kind when one is given; the `w:id` is free in this
   * body. Returns the control.
   */
  insertContentControl(kind?: ContentControlType): ContentControl {
    checkKind(kind, 'Block');
    const content = [...this.content];
    const sdt = sdtBlockFor(content, kind, nextControlId(this.container));
    this.content.length = 0;
    this.content.push(sdt as Element);
    linkParents(sdt, this.container);
    return new ContentControl(sdt, this.content, this);
  }
}

/** The width of the text area in twips (page width less the margins), for a new table. */
function writableWidthTwips(container: { sectPr?: wml.SectPr }): number | undefined {
  const pgSz = container.sectPr?.pgSz;
  if (!pgSz?.w) return undefined;
  const width = pgSz.w - (container.sectPr?.pgMar?.left ?? 0) - (container.sectPr?.pgMar?.right ?? 0);
  return width > 0 ? width : undefined;
}

function describe(el: Element): string {
  const tn = typeNameOf(el);
  return `${el.name?.localPart ?? '?'} (${tn ?? 'no TYPE_NAME'})`;
}
