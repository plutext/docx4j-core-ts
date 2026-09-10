import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import type { OpcPackage } from '../../packages/OpcPackage.mjs';
import { type Element, typeNameOf, childrenOf, linkParents, BLOCK_LEVEL_TYPES, textOf } from './tree.mjs';
import { wml as parseFragment, r as textRun, br as breakItem } from '@docx4j/generated-objects-ts/builders/wml';
import { runOf, paragraphOf } from './tree.mjs';
import { Paragraph } from './Paragraph.mjs';
import { Range } from './Range.mjs';
import type { SearchOptions } from './search.mjs';

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
            for (const row of childrenOf(v) ?? []) for (const cell of childrenOf(row.value as object) ?? []) visit(childrenOf(cell.value as object) ?? []);
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

  /** The tables directly in this body (as elements; `Table` views are CR-002 Phase C). */
  get tables(): BlockElement[] {
    return this.content.filter((el) => typeNameOf(el) === 'org_docx4j_wml.Tbl').map((element) => ({ element, container: this.content }));
  }

  /** The text, a paragraph per line. */
  get text(): string {
    return this.paragraphs.map((p) => p.text).join('\n');
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

  insertBreak(type: 'Page' | 'Line', location: 'Start' | 'End'): void {
    const p = this.insertParagraph('', location);
    p.p.content!.push(runOf([breakItem(type === 'Page' ? 'page' : undefined)]) as never);
    linkParents(p.p.content, p.p);
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
    for (const el of elements) {
      linkParents(el, owner);
      if (typeNameOf(el) === 'org_docx4j_wml.P') this.assignParaId(el.value as wml.P);
    }
    const first = elements[0]!;
    return typeNameOf(first) === 'org_docx4j_wml.P' ? new Paragraph(first as Element<wml.P>, container, this) : { element: first, container };
  }

  /** Inserts a WML fragment (one or more w:p / w:tbl / ... as inside document.xml); the standard prefixes are declared for it. */
  async insertXml(xml: string, location: 'Start' | 'End' | 'Before' | 'After', target?: Paragraph | BlockElement | Address): Promise<(Paragraph | BlockElement)[]> {
    const preprocess = this.package_?.loadOptions.preprocessor;
    const elements = await parseFragment(xml, { wrapper: 'body', preprocess: preprocess ? (doc) => preprocess(doc) : undefined });
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

  /** Removes everything (section properties stay). */
  clear(): void {
    this.content.length = 0;
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
    const parts = address.split('/');
    const prefix = parts[0]!;
    if (prefix !== this.prefix && !(prefix === 'body' && this.prefix === 'body')) return undefined;
    let current: object = this.container;
    let container: Element[] | undefined;
    let element: Element | undefined;
    for (const seg of parts.slice(1)) {
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
          for (const row of childrenOf(v) ?? []) for (const cell of childrenOf(row.value as object) ?? []) if (seek(cell.value as object)) return true;
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
          for (const row of childrenOf(v) ?? []) for (const cell of childrenOf(row.value as object) ?? []) if (seek(cell.value as object)) return true;
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
      // sdtContent is not a level of its own
      if ((parent as { TYPE_NAME?: string }).TYPE_NAME?.endsWith('SdtContentBlock')) parent = (parent as { PARENT?: object }).PARENT;
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
}

function describe(el: Element): string {
  const tn = typeNameOf(el);
  return `${el.name?.localPart ?? '?'} (${tn ?? 'no TYPE_NAME'})`;
}
