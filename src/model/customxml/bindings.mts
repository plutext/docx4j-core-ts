// CR-002 phase E, section 3.5: docx4j's two directions, under docx4j's names.
//
//   applyBindings()             = BindingHandler.applyBindings: push the custom XML values into the
//                                 bound controls, which is what Word does when it opens a document.
//   updateFromContentControls() = UpdateXmlFromDocumentSurface.updateCustomXmlParts: write what the
//                                 controls now show back into the custom XML.
//
// The Java is the reference for the shapes written (the run properties come from w:sdtPr/w:rPr, the
// value is trimmed, an empty value gives Word's placeholder run and w:showingPlcHdr, a multiline
// text control turns newlines into w:br); the departures are recorded in CR-002 section 12.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy } from '@docx4j/generated-objects-ts';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { sdtProperty } from '@docx4j/generated-objects-ts/builders/wml';
import type { Element } from '../content/tree.mjs';
import type { ContentControl } from '../content/ContentControl.mjs';

/** The style Word gives placeholder text, and the text docx4j's placeholder.xml carries. */
export const PLACEHOLDER_STYLE = 'PlaceholderText';
export const PLACEHOLDER_TEXT = 'Click here to enter text.';

/** What `applyBindings` and `updateFromContentControls` report. */
export interface BindingResult {
  /** Controls carrying a `w:dataBinding`. */
  bound: number;
  /** Controls whose value was written (into the document, or into the custom XML). */
  updated: number;
  /** Controls left alone: the part is missing, the XPath selects nothing, or the kind is not handled. */
  skipped: number;
}

function isTrue(value: string): boolean {
  const v = value.trim().toLowerCase();
  return v === 'true' || v === '1' || v === 'on' || v === 'yes';
}

/**
 * Kinds that hold structure rather than a value. Word maps a repeating section to a node set and
 * repeats what is inside it; writing a bound value into one would flatten the repeat, so these are
 * left alone in both directions (CR-002 section 12) — as is any control that holds a table or
 * another control.
 */
const CONTAINER_KINDS = new Set(['RepeatingSection', 'RepeatingSectionItem', 'Group', 'BuildingBlockGallery']);

function isContainer(control: ContentControl): boolean {
  if (CONTAINER_KINDS.has(control.type)) return true;
  if (control.contentControls.length > 0) return true;
  return control.form === 'Block' && control.tables.length > 0;
}

/** docx4j ValueInserterPlainTextImpl: the runs for a value, with the w:sdtPr run properties. */
export function runsForValue(value: string, rPr: wml.RPr | undefined, multiLine: boolean): Element[] {
  if (value === '') return [placeholderRun(rPr)];
  const lines = multiLine ? value.split(/[\n\r\f]+/).filter((line) => line !== '') : [value.replace(/[\n\r\f]+/g, '')];
  const out: Element[] = [];
  lines.forEach((line, i) => {
    if (i > 0) out.push(el.r({ ...(rPr ? { rPr: deepCopy(rPr) } : {}), content: [el.br({})] }) as Element);
    const text: wml.Text = { TYPE_NAME: 'org_docx4j_wml.Text', value: line };
    if (line.startsWith(' ') || line.endsWith(' ')) text.space = 'preserve';
    out.push(el.r({ ...(rPr ? { rPr: deepCopy(rPr) } : {}), content: [el.t(text)] }) as Element);
  });
  return out;
}

/** docx4j's placeholder.xml: a PlaceholderText-styled run, which goes with w:showingPlcHdr. */
function placeholderRun(rPr: wml.RPr | undefined): Element {
  const properties: wml.RPr = rPr ? deepCopy(rPr) : {};
  properties.rStyle = { val: PLACEHOLDER_STYLE };
  return el.r({ rPr: properties, content: [el.t({ value: PLACEHOLDER_TEXT })] }) as Element;
}

/**
 * A date in Word's `w:dateFormat` pattern (the .NET forms Word writes: `d MMMM yyyy`,
 * `dddd, d MMMM yyyy`, `d/MM/yyyy`), in the control's `w:lid` locale. Month and day names come
 * from `Intl`; text inside single or double quotes is literal. docx4j uses Java's
 * `SimpleDateFormat` with `dddd` mapped to `EEEE`; this is the same vocabulary.
 */
export function formatDate(pattern: string, date: Date, locale = 'en-US'): string {
  const pad = (n: number, width: number): string => String(n).padStart(width, '0');
  const name = (options: Intl.DateTimeFormatOptions): string => {
    try {
      return new Intl.DateTimeFormat(locale, options).format(date);
    } catch {
      return new Intl.DateTimeFormat('en-US', options).format(date);
    }
  };
  const hours12 = date.getHours() % 12 === 0 ? 12 : date.getHours() % 12;
  return pattern.replace(/d{1,4}|M{1,4}|y{2,4}|H{1,2}|h{1,2}|m{1,2}|s{1,2}|tt|'[^']*'|"[^"]*"/g, (token) => {
    if (token.startsWith("'") || token.startsWith('"')) return token.slice(1, -1);
    switch (token) {
      case 'd': return String(date.getDate());
      case 'dd': return pad(date.getDate(), 2);
      case 'ddd': return name({ weekday: 'short' });
      case 'dddd': return name({ weekday: 'long' });
      case 'M': return String(date.getMonth() + 1);
      case 'MM': return pad(date.getMonth() + 1, 2);
      case 'MMM': return name({ month: 'short' });
      case 'MMMM': return name({ month: 'long' });
      case 'yy': return pad(date.getFullYear() % 100, 2);
      case 'yyy': case 'yyyy': return String(date.getFullYear());
      case 'H': return String(date.getHours());
      case 'HH': return pad(date.getHours(), 2);
      case 'h': return String(hours12);
      case 'hh': return pad(hours12, 2);
      case 'm': return String(date.getMinutes());
      case 'mm': return pad(date.getMinutes(), 2);
      case 's': return String(date.getSeconds());
      case 'ss': return pad(date.getSeconds(), 2);
      case 'tt': return date.getHours() < 12 ? 'AM' : 'PM';
      default: return token;
    }
  });
}

/** Word's stored date form, `2015-01-29T00:00:00` (local, as `w:fullDate` without a zone). */
function isoOf(date: Date): string {
  const pad = (n: number): string => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

/** The date a bound value holds, when it is one (`2015-01-29T00:00:00`, `2015-01-29`, or with a Z). */
function dateOf(value: string): Date | undefined {
  if (!/^\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}(:\d{2})?)?/.test(value)) return undefined;
  const parsed = new Date(value.replace(/Z$/, '').replace(' ', 'T'));
  return Number.isNaN(parsed.getTime()) ? undefined : parsed;
}

/**
 * Pushes the custom XML value into one control (docx4j BindingHandler's per-sdt work). Returns
 * false when nothing was written: no binding, no part, an XPath that selects nothing, or a kind
 * this phase does not handle (pictures, explicit rich text).
 */
export function applyBindingTo(control: ContentControl): boolean {
  const mapping = control.xmlMapping;
  if (!mapping.isMapped) return false;
  const node = mapping.customXmlNode;
  if (!node) return false;
  const value = node.text.trim();

  // A picture binding carries base64 image data; docx4j replaces the a:blip embed. Deferred (section 12).
  if (control.type === 'Picture') return false;
  // Explicit rich text is bound from flat OPC or XHTML in docx4j; deferred, as it is there for this route.
  if (sdtProperty(control.sdt.sdtPr, 'richText') !== undefined) return false;
  if (isContainer(control)) return false;

  if (control.type === 'CheckBox') {
    const checkbox = control.checkboxContentControl;
    if (!checkbox) return false;
    checkbox.isChecked = isTrue(value);
    return true;
  }

  let display = value;
  if (control.type === 'DropDownList' || control.type === 'ComboBox') {
    const list = control.dropDownListContentControl ?? control.comboBoxContentControl;
    if (list) {
      const match = list.listItems.find((item) => item.value === value);
      if (match) display = match.displayText;
      list.lastValue = value;
    }
  } else if (control.type === 'DatePicker') {
    const date = control.datePickerContentControl;
    const parsed = dateOf(value);
    if (date && parsed) {
      date.fullDate = parsed;
      // Word shows the date in w:dateFormat, not the stored form (docx4j formats it the same way)
      if (date.dateFormat !== '') display = formatDate(date.dateFormat, parsed, date.dateDisplayLocale || undefined);
    }
  }

  control.setBoundContent(runsForValue(display, control.runProperties, control.isMultiLine), display === '');
  return true;
}

/** `applyBindings` over a list of controls. */
export function applyBindingsTo(controls: ContentControl[]): BindingResult {
  const result: BindingResult = { bound: 0, updated: 0, skipped: 0 };
  for (const control of controls) {
    if (!control.xmlMapping.isMapped) continue;
    result.bound++;
    if (applyBindingTo(control)) result.updated++; else result.skipped++;
  }
  return result;
}

/**
 * Writes what one control shows back into the node it is bound to (docx4j
 * UpdateXmlFromDocumentSurface). A checkbox writes 'true' or 'false' and a list writes the entry's
 * value rather than its display text, which is what the data meant in the first place (docx4j
 * writes the rendered glyph; section 12).
 */
export function updateFromControl(control: ContentControl): boolean {
  const mapping = control.xmlMapping;
  if (!mapping.isMapped) return false;
  const node = mapping.customXmlNode;
  if (!node) return false;
  if (control.type === 'Picture') return false;
  if (sdtProperty(control.sdt.sdtPr, 'richText') !== undefined) return false;
  if (isContainer(control)) return false;
  // A control showing its placeholder holds no value.
  if (control.isShowingPlaceholder) return false;

  let value = control.text;
  if (control.type === 'CheckBox') {
    const checkbox = control.checkboxContentControl;
    if (!checkbox) return false;
    value = checkbox.isChecked ? 'true' : 'false';
  } else if (control.type === 'DropDownList' || control.type === 'ComboBox') {
    const list = control.dropDownListContentControl ?? control.comboBoxContentControl;
    const match = list?.listItems.find((item) => item.displayText === value);
    if (match) value = match.value;
  } else if (control.type === 'DatePicker') {
    // the control shows the formatted date; the data keeps Word's stored form (docx4j skips dates)
    const full = control.datePickerContentControl?.fullDate;
    if (full) value = isoOf(full);
  }
  if (node.text === value) return false;
  node.text = value;
  return true;
}

/** `updateFromContentControls` over a list of controls. */
export function updateFromControls(controls: ContentControl[]): BindingResult {
  const result: BindingResult = { bound: 0, updated: 0, skipped: 0 };
  for (const control of controls) {
    if (!control.xmlMapping.isMapped) continue;
    result.bound++;
    if (updateFromControl(control)) result.updated++; else result.skipped++;
  }
  return result;
}
