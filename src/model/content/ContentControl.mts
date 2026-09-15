// CR-002 phase C: ContentControl, a subset of Office JS's Word.ContentControl over w:sdt in its
// four forms (block, run, row and cell). The typed kinds (checkbox, date, drop-down), the XML
// mapping and insertContentControl are phase E; what is here is what a reader and a text editor
// need. Note (objects CR-002): a run-level w:sdt parsed at body level comes back as SdtBlock, so
// the form is decided by what the content holds as well as by the type.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { marshalString } from '@docx4j/generated-objects-ts';
import { type Element, typeNameOf, childrenOf, textOf, runItemsOf, segmentsOf, linkParents, SDT_TYPES } from './tree.mjs';
import type { Body } from './Body.mjs';
import type { Paragraph } from './Paragraph.mjs';
import { Range } from './Range.mjs';
import { Table } from './Table.mjs';
import type { SearchOptions } from './search.mjs';

/** Office JS `Word.ContentControlType`, as Word reports it from w:sdtPr. */
export type ContentControlType =
  | 'Unknown' | 'RichText' | 'PlainText' | 'Picture' | 'BuildingBlockGallery' | 'CheckBox' | 'ComboBox'
  | 'DropDownList' | 'DatePicker' | 'RepeatingSection' | 'RepeatingSectionItem' | 'Group' | 'Citation' | 'Bibliography' | 'Equation';

/** The kind element in w:sdtPr to the type Word reports. */
const TYPE_BY_ELEMENT: Readonly<Record<string, ContentControlType>> = {
  text: 'PlainText', richText: 'RichText', picture: 'Picture', docPartObj: 'BuildingBlockGallery', docPartList: 'BuildingBlockGallery',
  checkbox: 'CheckBox', comboBox: 'ComboBox', dropDownList: 'DropDownList', date: 'DatePicker',
  repeatingSection: 'RepeatingSection', repeatingSectionItem: 'RepeatingSectionItem', group: 'Group',
  citation: 'Citation', bibliography: 'Bibliography', equation: 'Equation',
};

/** Which of the four w:sdt forms a control is. */
export type ContentControlForm = 'Block' | 'Run' | 'Row' | 'Cell';

/** A subset of Office JS `Word.ContentControl` over a `w:sdt`. A view, as Paragraph is. */
export class ContentControl {
  constructor(
    /** The `w:sdt` element pair (SdtBlock, SdtRun, CTSdtRow or CTSdtCell). */
    readonly element: Element<wml.SdtBlock | wml.SdtRun | wml.CTSdtRow | wml.CTSdtCell>,
    /** The array holding the element. */
    readonly container: Element[],
    /** The body the control is in (the part's). */
    readonly parentBody: Body,
  ) {}

  /** The sdt value. */
  get sdt(): { sdtPr?: wml.SdtPr; sdtContent?: { content?: Element[] } } {
    return this.element.value as { sdtPr?: wml.SdtPr; sdtContent?: { content?: Element[] } };
  }

  /** The control's content array (`w:sdtContent`); empty and detached when the control has none (reads never add one). */
  get content(): Element[] {
    return this.sdt.sdtContent?.content ?? [];
  }

  /** Which of the four forms this is: block, run, row or cell content. */
  get form(): ContentControlForm {
    switch (typeNameOf(this.element)) {
      case 'org_docx4j_wml.CTSdtRow': return 'Row';
      case 'org_docx4j_wml.CTSdtCell': return 'Cell';
      case 'org_docx4j_wml.SdtRun': return 'Run';
      default: {
        // objects CR-002: a run-level control parsed at body level is typed SdtBlock; what it
        // holds decides.
        const first = this.content.find((e) => typeNameOf(e) !== undefined);
        const tn = first ? typeNameOf(first) : undefined;
        return tn === 'org_docx4j_wml.R' ? 'Run' : 'Block';
      }
    }
  }

  /** The type Word reports (w:sdtPr's kind element); RichText when none is given, as Word does. */
  get type(): ContentControlType {
    for (const item of this.sdt.sdtPr?.rPrOrAliasOrLock ?? []) {
      const type = TYPE_BY_ELEMENT[item.name.localPart];
      if (type) return type;
    }
    return 'RichText';
  }

  /** w:tag. */
  get tag(): string {
    return this.property('tag');
  }
  set tag(value: string) {
    this.setProperty('tag', value, (v) => el.tag({ val: v }) as Element);
  }

  /** w:alias, which is what Word's user interface calls the title. */
  get title(): string {
    return this.property('alias');
  }
  set title(value: string) {
    this.setProperty('alias', value, (v) => el.alias({ val: v }) as Element);
  }

  /** w:id, the number Word gives a control; 0 when it has none. */
  get id(): number {
    const item = this.sdt.sdtPr?.rPrOrAliasOrLock?.find((i) => i.name.localPart === 'id');
    const val = (item?.value as { val?: number } | undefined)?.val;
    return typeof val === 'number' ? val : 0;
  }

  /** The control's text, a line per paragraph (docx4j TextUtils). */
  get text(): string {
    return textOf(this.sdt.sdtContent ?? {});
  }

  /** Every paragraph in the control; for a run-level control, the paragraph it is in. */
  get paragraphs(): Paragraph[] {
    if (this.form === 'Run') {
      const p = this.parentParagraph();
      return p ? [p] : [];
    }
    return this.body().paragraphs;
  }

  /** The tables directly in the control; for a row or cell control, the table it is part of. */
  get tables(): Table[] {
    if (this.form === 'Block') return this.body().tables;
    if (this.form === 'Row' || this.form === 'Cell') {
      const table = this.ancestorTable();
      return table ? [table] : [];
    }
    return [];
  }

  /** The controls nested in this one, in document order. */
  get contentControls(): ContentControl[] {
    const out: ContentControl[] = [];
    collectControls(this.content, this.parentBody, out);
    return out;
  }

  /** Text at the start, the end, or replacing the control's text. */
  insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range {
    const form = this.form;
    if ((form === 'Row' || form === 'Cell') && location === 'Replace') {
      throw new Docx4JException(`A ${form.toLowerCase()}-level content control holds ${form === 'Row' ? 'rows' : 'cells'}; replace the text of its cells instead`);
    }
    if (form !== 'Run') return this.body().insertText(text, location);
    const span = this.runSpan();
    if (!span) throw new Docx4JException('This content control is not in a paragraph');
    const { paragraph, start, end } = span;
    if (location === 'Start') return paragraph.splice(start, start, text);
    if (location === 'End') return paragraph.splice(end, end, text);
    return paragraph.splice(start, end, text);
  }

  /** A paragraph at the start or end of a block control; before or after the paragraph of a run control. */
  insertParagraph(text: string, location: 'Start' | 'End' | 'Before' | 'After'): Paragraph {
    const form = this.form;
    if ((form === 'Row' || form === 'Cell') && (location === 'Start' || location === 'End')) {
      throw new Docx4JException(`A ${form.toLowerCase()}-level content control holds ${form === 'Row' ? 'rows' : 'cells'}, not paragraphs; insert into a cell's body`);
    }
    if (form === 'Run' || location === 'Before' || location === 'After') {
      const p = this.parentParagraph();
      if (p) return p.insertParagraph(text, location === 'Before' ? 'Before' : 'After');
      return this.parentBody.insertElement(paragraphElement(text), location === 'Before' ? 'Before' : 'After', { element: this.element as Element, container: this.container }) as Paragraph;
    }
    return this.body().insertParagraph(text, location);
  }

  /** Matches within the control. */
  search(text: string, options?: SearchOptions): Range[] {
    if (this.form !== 'Run') return this.body().search(text, options);
    const range = this.getRange();
    return range ? range.search(text, options) : [];
  }

  /**
   * The control's range. For a run-level control, exactly the text it holds; for the other forms,
   * the first paragraph (departure: a Range here is within one paragraph, CR-002 section 7).
   */
  getRange(location: 'Whole' | 'Start' | 'End' | 'Content' = 'Whole'): Range {
    if (this.form === 'Run') {
      const span = this.runSpan();
      if (!span) throw new Docx4JException('This content control is not in a paragraph');
      const { paragraph, start, end } = span;
      if (location === 'Start') return new Range(paragraph, start, start);
      if (location === 'End') return new Range(paragraph, end, end);
      return new Range(paragraph, start, end);
    }
    const first = this.paragraphs[0];
    if (!first) throw new Docx4JException('This content control holds no paragraph');
    return first.getRange(location);
  }

  /** Removes the control; `keepContent` puts what it held in its place, as Office JS. */
  delete(keepContent = false): void {
    const at = this.container.indexOf(this.element as Element);
    if (at < 0) return;
    const content = keepContent ? [...this.content] : [];
    this.container.splice(at, 1, ...content);
    if (keepContent) {
      const owner = ownerOf(this.container, this.parentBody);
      for (const item of content) linkParents(item, owner);
    }
  }

  /** The XML of the control (extension). */
  getXml(): Promise<string> {
    return marshalString(this.element as Element);
  }

  // --- internals ---

  /** A Body view over the control's content, for the block form; the `w:sdtContent` is created if the control has none. */
  private body(): Body {
    const sdt = this.sdt;
    sdt.sdtContent ??= {};
    const prefix = this.parentBody.addressOf({ element: this.element as Element, container: this.container }) ?? this.parentBody.prefix;
    return this.parentBody.sub(sdt.sdtContent as { content?: Element[] }, prefix);
  }

  private property(localPart: string): string {
    const item = this.sdt.sdtPr?.rPrOrAliasOrLock?.find((i) => i.name.localPart === localPart);
    return (item?.value as { val?: string } | undefined)?.val ?? '';
  }

  private setProperty(localPart: string, value: string, make: (value: string) => Element): void {
    const sdt = this.sdt as { sdtPr?: wml.SdtPr };
    sdt.sdtPr ??= {};
    const items = (sdt.sdtPr.rPrOrAliasOrLock ??= []);
    const at = items.findIndex((i) => i.name.localPart === localPart);
    if (value === '') { if (at >= 0) items.splice(at, 1); return; }
    const element = make(value) as never;
    if (at >= 0) items[at] = element; else items.push(element);
  }

  /** The paragraph a run-level control is in. */
  private parentParagraph(): Paragraph | undefined {
    let current: object | undefined = this.element.value as object;
    for (let guard = 0; current && guard < 64; guard++) {
      if (typeNameOf(current) === 'org_docx4j_wml.P') return this.parentBody.paragraphFor(current as wml.P);
      current = (current as { PARENT?: object }).PARENT;
    }
    return undefined;
  }

  /** The offsets a run-level control covers in its paragraph. */
  private runSpan(): { paragraph: Paragraph; start: number; end: number } | undefined {
    const paragraph = this.parentParagraph();
    if (!paragraph) return undefined;
    const inside = (run: object): boolean => {
      let current: object | undefined = run;
      for (let guard = 0; current && guard < 32; guard++) {
        if (current === (this.element.value as object)) return true;
        current = (current as { PARENT?: object }).PARENT;
      }
      return false;
    };
    let start: number | undefined;
    let end = 0;
    for (const seg of segmentsOf(paragraph.p)) {
      if (!inside(seg.run)) continue;
      if (start === undefined) start = seg.start;
      end = seg.end;
    }
    if (start === undefined) {
      // an empty control: its position is where its runs would be
      return { paragraph, start: 0, end: 0 };
    }
    return { paragraph, start, end };
  }

  /** The table a row-level or cell-level control belongs to. */
  private ancestorTable(): Table | undefined {
    let current: object | undefined = (this.element.value as { PARENT?: object }).PARENT;
    for (let guard = 0; current && guard < 8; guard++) {
      if (typeNameOf(current) === 'org_docx4j_wml.Tbl') {
        const tbl = current as wml.Tbl;
        const container = childrenOf(((tbl as { PARENT?: object }).PARENT ?? {}) as object);
        const element = container?.find((e) => e.value === tbl) as Element<wml.Tbl> | undefined;
        return element && container ? new Table(element, container, this.parentBody) : undefined;
      }
      current = (current as { PARENT?: object }).PARENT;
    }
    return undefined;
  }
}

/** Every content control under a list of block-level items, in document order, nested ones included. */
export function collectControls(items: Element[], body: Body, out: ContentControl[]): void {
  for (const element of items) {
    const tn = typeNameOf(element);
    if (tn === undefined) continue;
    if (SDT_TYPES.has(tn)) {
      const control = new ContentControl(element as Element<wml.SdtBlock>, items, body);
      out.push(control);
      collectControls(control.content, body, out);
      continue;
    }
    if (tn === 'org_docx4j_wml.P') {
      collectRunControls(element.value as wml.P, body, out);
      continue;
    }
    if (tn === 'org_docx4j_wml.Tbl') {
      collectInRows(childrenOf(element.value as object) ?? [], body, out);
      continue;
    }
    const children = childrenOf(element.value as object);
    if (children) collectControls(children, body, out);
  }
}

/** Row-level controls and the controls inside the rows of a table. */
function collectInRows(items: Element[], body: Body, out: ContentControl[]): void {
  for (const element of items) {
    const tn = typeNameOf(element);
    if (tn === undefined) continue;
    if (SDT_TYPES.has(tn)) {
      const control = new ContentControl(element as Element<wml.CTSdtRow>, items, body);
      out.push(control);
      collectInRows(control.content, body, out);
      continue;
    }
    if (tn === 'org_docx4j_wml.Tr') collectInCells(childrenOf(element.value as object) ?? [], body, out);
  }
}

/** Cell-level controls and the controls inside the cells of a row. */
function collectInCells(items: Element[], body: Body, out: ContentControl[]): void {
  for (const element of items) {
    const tn = typeNameOf(element);
    if (tn === undefined) continue;
    if (SDT_TYPES.has(tn)) {
      const control = new ContentControl(element as Element<wml.CTSdtCell>, items, body);
      out.push(control);
      collectInCells(control.content, body, out);
      continue;
    }
    if (tn === 'org_docx4j_wml.Tc') collectControls(childrenOf(element.value as object) ?? [], body, out);
  }
}

/** The run-level controls in a paragraph (or any run holder), in order, nested ones included. */
export function collectRunControls(holder: object, body: Body, out: ContentControl[]): void {
  const items = runItemsOf(holder);
  if (!items) return;
  for (const element of items) {
    const tn = typeNameOf(element);
    if (tn === undefined) continue;
    if (SDT_TYPES.has(tn)) {
      const control = new ContentControl(element as Element<wml.SdtRun>, items, body);
      out.push(control);
      collectRunControls(element.value as object, body, out);
      continue;
    }
    if (tn !== 'org_docx4j_wml.R') collectRunControls(element.value as object, body, out);
  }
}

function paragraphElement(text: string): Element {
  return el.p({ content: text === '' ? [] : [el.r({ content: [el.t({ value: text })] })] }) as Element;
}

function ownerOf(container: Element[], body: Body): object {
  // the container is a live array of the body's tree; PARENT is set from whatever holds it
  const seek = (value: object): object | undefined => {
    const children = childrenOf(value);
    if (children === container) return value;
    if (!children) return undefined;
    for (const child of children) {
      const v = child.value;
      if (typeof v !== 'object' || v === null) continue;
      const found = seek(v);
      if (found) return found;
    }
    return undefined;
  };
  return seek(body.container as object) ?? (body.container as object);
}

