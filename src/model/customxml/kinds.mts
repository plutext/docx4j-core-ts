// CR-002 phase E, section 3.5: the kind-specific views Office JS hangs off a content control
// (checkboxContentControl, datePickerContentControl, dropDownListContentControl,
// comboBoxContentControl, pictureContentControl, repeatingSectionContentControl,
// groupContentControl). Each is a view over the matching w:sdtPr child, as ContentControl is a
// view over the w:sdt; `undefined` for a control of another kind, as Office JS reports it.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as w14 from '@docx4j/generated-objects-ts/modules/org_docx4j_w14';
import type * as w15 from '@docx4j/generated-objects-ts/modules/org_docx4j_w15';
import { deepCopy } from '@docx4j/generated-objects-ts';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import type { Element } from '../content/tree.mjs';
import type { ContentControl } from '../content/ContentControl.mjs';
import type { InlinePicture } from '../content/InlinePicture.mjs';

/** Word's default checkbox glyphs (docx4j BindingTraverserXSLT uses the same two). */
export const CHECKED_SYMBOL = '☒';
export const UNCHECKED_SYMBOL = '☐';
/** The font Word writes with them. */
export const CHECKBOX_FONT = 'MS Gothic';

function isOn(value: string | undefined): boolean {
  return value === '1' || value === 'true' || value === 'on';
}

/** A hex character code ('2612') as its character; the string itself when it is not one. */
function symbolOf(state: { val?: string } | undefined, fallback: string): string {
  const val = state?.val;
  if (!val) return fallback;
  return /^[0-9A-Fa-f]{4,5}$/.test(val) ? String.fromCodePoint(parseInt(val, 16)) : val;
}

function codeOf(symbol: string): string {
  const point = symbol.codePointAt(0);
  return point === undefined ? '2610' : point.toString(16).toUpperCase().padStart(4, '0');
}

/** Office JS `Word.CheckboxContentControl` over `w14:checkbox`. */
export class CheckboxContentControl {
  constructor(private readonly control: ContentControl, private readonly checkbox: w14.CTSdtCheckbox) {}

  get isChecked(): boolean {
    return isOn(this.checkbox.checked?.val);
  }
  set isChecked(value: boolean) {
    (this.checkbox.checked ??= {}).val = value ? '1' : '0';
    this.control.setCheckboxGlyph(value ? this.checkedSymbol : this.uncheckedSymbol);
  }

  /** The glyph shown when checked (`w14:checkedState/@w14:val`, a hex character code). */
  get checkedSymbol(): string {
    return symbolOf(this.checkbox.checkedState, CHECKED_SYMBOL);
  }
  set checkedSymbol(symbol: string) {
    (this.checkbox.checkedState ??= {}).val = codeOf(symbol);
    this.checkbox.checkedState.font ??= CHECKBOX_FONT;
  }

  get uncheckedSymbol(): string {
    return symbolOf(this.checkbox.uncheckedState, UNCHECKED_SYMBOL);
  }
  set uncheckedSymbol(symbol: string) {
    (this.checkbox.uncheckedState ??= {}).val = codeOf(symbol);
    this.checkbox.uncheckedState.font ??= CHECKBOX_FONT;
  }

  /** The font of the two glyphs (extension; Word writes MS Gothic). */
  get font(): string {
    return this.checkbox.checkedState?.font ?? this.checkbox.uncheckedState?.font ?? CHECKBOX_FONT;
  }
}

/** Office JS `Word.DatePickerContentControl` over `w:date`. */
export class DatePickerContentControl {
  constructor(private readonly date: wml.CTSdtDate) {}

  /** `w:dateFormat`, Word's .NET date pattern ('d MMMM yyyy'). Office JS calls it dateDisplayFormat. */
  get dateFormat(): string {
    return this.date.dateFormat?.val ?? '';
  }
  set dateFormat(format: string) {
    if (format === '') { delete this.date.dateFormat; return; }
    this.date.dateFormat = { val: format };
  }

  /** Office JS's name for the same property. */
  get dateDisplayFormat(): string {
    return this.dateFormat;
  }
  set dateDisplayFormat(format: string) {
    this.dateFormat = format;
  }

  /** `w:lid`, the language the date is shown in ('en-AU'). */
  get dateDisplayLocale(): string {
    return this.date.lid?.val ?? '';
  }
  set dateDisplayLocale(locale: string) {
    if (locale === '') { delete this.date.lid; return; }
    this.date.lid = { val: locale };
  }

  /** `w:calendar` ('gregorian', 'hijri', ...). */
  get dateCalendarType(): string {
    return this.date.calendar?.val ?? 'gregorian';
  }
  set dateCalendarType(calendar: string) {
    this.date.calendar = { val: calendar as wml.CTCalendarType['val'] };
  }

  /** `w:storeMappedDataAs`: how a bound date is written into the custom XML ('dateTime', 'date', 'text'). */
  get dateStorageFormat(): string {
    return this.date.storeMappedDataAs?.val ?? 'text';
  }
  set dateStorageFormat(format: string) {
    this.date.storeMappedDataAs = { val: format as wml.CTSdtDateMappingType['val'] };
  }

  /** `w:fullDate`, the date itself when Word has one (extension). */
  get fullDate(): Date | undefined {
    const c = this.date.fullDate;
    if (!c || c.year === undefined || Number.isNaN(c.year)) return undefined;
    return new Date(Date.UTC(c.year, (c.month ?? 1) - 1, c.day ?? 1, c.hour ?? 0, c.minute ?? 0, Math.floor(c.second ?? 0)));
  }
  set fullDate(value: Date | undefined) {
    if (!value) { delete this.date.fullDate; return; }
    this.date.fullDate = {
      year: value.getUTCFullYear(), month: value.getUTCMonth() + 1, day: value.getUTCDate(),
      hour: value.getUTCHours(), minute: value.getUTCMinutes(), second: value.getUTCSeconds(), timezone: 0,
    };
  }
}

/** One entry of a drop-down list or combo box (Office JS `Word.ContentControlListItem`). */
export class ContentControlListItem {
  constructor(private readonly items: wml.CTSdtListItem[], private readonly item: wml.CTSdtListItem) {}

  get displayText(): string {
    return this.item.displayText ?? '';
  }
  set displayText(text: string) {
    this.item.displayText = text;
  }

  get value(): string {
    return this.item.value ?? '';
  }
  set value(value: string) {
    this.item.value = value;
  }

  get index(): number {
    return this.items.indexOf(this.item);
  }

  delete(): void {
    const at = this.index;
    if (at >= 0) this.items.splice(at, 1);
  }
}

/** Office JS `Word.DropDownListContentControl` / `Word.ComboBoxContentControl`: the same shape. */
export class ListContentControl {
  constructor(private readonly list: wml.CTSdtDropDownList | wml.CTSdtComboBox) {}

  get listItems(): ContentControlListItem[] {
    const items = (this.list.listItem ??= []);
    return items.map((item) => new ContentControlListItem(items, item));
  }

  /** Adds an entry; `value` defaults to the display text, `index` appends when omitted. */
  addListItem(displayText: string, value?: string, index?: number): ContentControlListItem {
    const items = (this.list.listItem ??= []);
    const item: wml.CTSdtListItem = { TYPE_NAME: 'org_docx4j_wml.CTSdtListItem', displayText, value: value ?? displayText };
    if (index === undefined || index >= items.length) items.push(item); else items.splice(Math.max(0, index), 0, item);
    return new ContentControlListItem(items, item);
  }

  deleteAllListItems(): void {
    this.list.listItem = [];
  }

  /** `w:lastValue`, what Word last showed (extension). */
  get lastValue(): string {
    return this.list.lastValue ?? '';
  }
  set lastValue(value: string) {
    this.list.lastValue = value;
  }
}

/** Office JS `Word.PictureContentControl`. */
export class PictureContentControl {
  constructor(private readonly control: ContentControl) {}

  /** The picture the control holds, if any. */
  get inlinePicture(): InlinePicture | undefined {
    for (const paragraph of this.control.paragraphs) {
      const picture = paragraph.inlinePictures[0];
      if (picture) return picture;
    }
    return undefined;
  }
}

/** Office JS `Word.RepeatingSectionContentControl` over `w15:repeatingSection`. */
export class RepeatingSectionContentControl {
  constructor(private readonly control: ContentControl, private readonly section: w15.CTSdtRepeatedSection) {}

  /** The repeated items: the `w15:repeatingSectionItem` controls directly inside. */
  get items(): ContentControl[] {
    return this.control.contentControls.filter((c) => c.type === 'RepeatingSectionItem');
  }

  get sectionTitle(): string {
    return this.section.sectionTitle?.val ?? '';
  }
  set sectionTitle(title: string) {
    if (title === '') { delete this.section.sectionTitle; return; }
    this.section.sectionTitle = { val: title };
  }

  /** Whether Word offers the + and x buttons (`w15:doNotAllowInsertDeleteSection`, inverted). */
  get allowInsertDeleteSection(): boolean {
    const flag = this.section.doNotAllowInsertDeleteSection;
    return !(flag !== undefined && flag.val !== false);
  }
  set allowInsertDeleteSection(allow: boolean) {
    if (allow) delete this.section.doNotAllowInsertDeleteSection;
    else this.section.doNotAllowInsertDeleteSection = { val: true };
  }

  /**
   * Adds a copy of the item at `index` after it, as Word's + button does, and returns the new
   * item's control. The copy is a deep copy with its PARENT links rebuilt by the container.
   */
  insertItemAfter(index: number): ContentControl {
    const items = this.items;
    const item = items[index];
    if (!item) throw new Docx4JException(`This repeating section has ${items.length} item(s); there is none at ${index}`);
    return item.insertCopyAfter();
  }
}

/** Office JS `Word.GroupContentControl`: no members of its own. */
export class GroupContentControl {}

/** The `w:sdtContent` glyph run of a checkbox, as Word writes it (docx4j's binding does the same). */
export function checkboxRun(symbol: string, font: string, rPr?: wml.RPr): Element {
  const properties: wml.RPr = rPr ? deepCopy(rPr) : {};
  properties.rFonts = { ascii: font, hAnsi: font, eastAsia: font, hint: 'eastAsia' };
  return el.r({ rPr: properties, content: [el.t({ value: symbol })] }) as Element;
}
