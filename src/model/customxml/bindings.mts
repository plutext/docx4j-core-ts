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
import { sdtProperty, textOf } from '@docx4j/generated-objects-ts/builders/wml';
import type { Element } from '../content/tree.mjs';
import type { ContentControl } from '../content/ContentControl.mjs';
import { bindPicture } from './pictures.mjs';

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
  /** What the XHTML and picture passes left or dropped, one line each (CR-005 sections 8.2 and 9; REQ-076). */
  notes?: string[];
}

/**
 * Whether a node's value is a Flat OPC package: what Word writes into the node of a mapped rich
 * text control when its content is edited, a whole `pkg:package` as escaped text (`test/README.md`
 * check 35). It is content, never a control's text - Word 2010 shows it as raw XML in a plain text
 * control bound to the same node - so neither direction moves it as text (CR-002 section 40).
 */
export function holdsFlatOpc(value: string): boolean {
  return /^\s*(<\?[^>]*\?>\s*)*<([\w.-]+:)?package[\s>]/.test(value)
    && value.includes('http://schemas.microsoft.com/office/2006/xmlPackage');
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

/**
 * A control's text as its node takes it (CR-002 section 60 D4, F1, F2): `textOf`'s characters, as before CR-012 phase 4
 * moved `ContentControl.text` to Office JS's, its paragraphs and breaks a line feed each, a non-breaking hyphen `-` as
 * Word writes it (check 55b), an optional hyphen U+00AD, a note reference nothing. Office JS's U+000B and U+001E are no
 * XML characters.
 */
export function boundValueOf(control: ContentControl): string {
  return textOf((control.sdt.sdtContent ?? {}) as object).replace(/\u2011/g, '-');
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

/**
 * docx4j's placeholder.xml: a PlaceholderText-styled run over the control's own run properties, which goes with
 * w:showingPlcHdr; its text docx4j's, or a glossary entry's (CR-002 section 63, D8).
 */
export function placeholderRun(rPr: wml.RPr | undefined, text = PLACEHOLDER_TEXT): Element {
  const properties: wml.RPr = rPr ? deepCopy(rPr) : {};
  properties.rStyle = { val: PLACEHOLDER_STYLE };
  const t: wml.Text = { TYPE_NAME: 'org_docx4j_wml.Text', value: text };
  if (text.startsWith(' ') || text.endsWith(' ')) t.space = 'preserve';
  return el.r({ rPr: properties, content: [el.t(t)] }) as Element;
}

/**
 * A date in Word's `w:dateFormat` pattern (the .NET forms Word writes: `d MMMM yyyy`,
 * `dddd, d MMMM yyyy`, `d/MM/yyyy`), in the control's `w:lid` locale. Month and day names come
 * from `Intl`; text inside single or double quotes is literal. docx4j uses Java's
 * `SimpleDateFormat` with `dddd` mapped to `EEEE`; this is the same vocabulary.
 *
 * The date's **UTC fields** are formatted: a bound date is its nominal date, the fields as written,
 * which `w:fullDate` keeps with a `Z` and `DatePickerContentControl.fullDate` returns as a `Date` in
 * UTC, so the day shown is the same whatever the machine's time zone (CR-002 section 56).
 */
export function formatDate(pattern: string, date: Date, locale = 'en-US'): string {
  const pad = (n: number, width: number): string => String(n).padStart(width, '0');
  const name = (options: Intl.DateTimeFormatOptions): string => {
    try {
      return new Intl.DateTimeFormat(locale, { ...options, timeZone: 'UTC' }).format(date);
    } catch {
      return new Intl.DateTimeFormat('en-US', { ...options, timeZone: 'UTC' }).format(date);
    }
  };
  const hours = date.getUTCHours();
  const hours12 = hours % 12 === 0 ? 12 : hours % 12;
  return pattern.replace(/d{1,4}|M{1,4}|y{2,4}|H{1,2}|h{1,2}|m{1,2}|s{1,2}|tt|'[^']*'|"[^"]*"/g, (token) => {
    if (token.startsWith("'") || token.startsWith('"')) return token.slice(1, -1);
    switch (token) {
      case 'd': return String(date.getUTCDate());
      case 'dd': return pad(date.getUTCDate(), 2);
      case 'ddd': return name({ weekday: 'short' });
      case 'dddd': return name({ weekday: 'long' });
      case 'M': return String(date.getUTCMonth() + 1);
      case 'MM': return pad(date.getUTCMonth() + 1, 2);
      case 'MMM': return name({ month: 'short' });
      case 'MMMM': return name({ month: 'long' });
      case 'yy': return pad(date.getUTCFullYear() % 100, 2);
      case 'yyy': case 'yyyy': return String(date.getUTCFullYear());
      case 'H': return String(hours);
      case 'HH': return pad(hours, 2);
      case 'h': return String(hours12);
      case 'hh': return pad(hours12, 2);
      case 'm': return String(date.getUTCMinutes());
      case 'mm': return pad(date.getUTCMinutes(), 2);
      case 's': return String(date.getUTCSeconds());
      case 'ss': return pad(date.getUTCSeconds(), 2);
      case 'tt': return hours < 12 ? 'AM' : 'PM';
      default: return token;
    }
  });
}

/** Word's stored date form, `2015-01-29T00:00:00`: a nominal date's fields, with no zone (CR-002 section 56). */
function isoOf(date: Date): string {
  const pad = (n: number): string => String(n).padStart(2, '0');
  return `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`
    + `T${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}:${pad(date.getUTCSeconds())}`;
}

/**
 * The date a bound value holds, when it is one (`2015-01-29`, `2015-01-29T00:00:00`, with seconds' fractions,
 * a `Z` or an offset), as its nominal date: the fields as written, in UTC, the zone marker ignored, so that
 * `2026-10-22T00:00:00Z` and `2026-10-22T00:00:00+05:00` are 22 October in any time zone (docx4j f62314d3b's
 * rule, as Word shows them; CR-002 section 56). A date that doesn't exist (`2026-02-30`) is not one.
 */
function dateOf(value: string): Date | undefined {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?(?:Z|[+-]\d{2}:?\d{2})?$/.exec(value);
  if (!m) return undefined;
  const [year, month, day, hour, minute, second] = m.slice(1).map((f) => Number(f ?? 0)) as [number, number, number, number, number, number];
  const date = new Date(Date.UTC(year, month - 1, day, hour, minute, second));
  // Date.UTC rolls an impossible field over (30 February is 2 March): such a value is not a date
  const exists = date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day
    && date.getUTCHours() === hour && date.getUTCMinutes() === minute && date.getUTCSeconds() === second;
  return exists ? date : undefined;
}

/**
 * Pushes the custom XML value into one control (docx4j BindingHandler's per-sdt work). Returns
 * false when nothing was written: no binding, no part, an XPath that selects nothing, a kind
 * this phase does not handle (explicit rich text), a node holding a Flat OPC package, or a
 * picture control whose node holds no image (`notes` says which; CR-005 section 9).
 */
export function applyBindingTo(control: ContentControl, notes?: string[]): boolean {
  const mapping = control.xmlMapping;
  if (!mapping.isMapped) return false;
  const node = mapping.customXmlNode;
  if (!node) return false;
  const value = node.text.trim();

  // A picture control's node holds base64 image data: the picture's a:blip is pointed at a new
  // image part and the drawing kept, as docx4j's bind.xslt mode picture3 does (CR-005 section 9).
  if (control.type === 'Picture') {
    const outcome = bindPicture(control, value, { kind: 'keep' }, control.title || control.tag || `picture control ${control.id}`);
    if (outcome.note) notes?.push(outcome.note);
    return outcome.done;
  }
  // Explicit rich text is bound from flat OPC or XHTML in docx4j; deferred, as it is there for this route. A block
  // control bound by w:dataBinding is plain text to Word, w:richText or not (CR-002 section 61 H4), and is bound as text.
  if (sdtProperty(control.sdt.sdtPr, 'richText') !== undefined && control.type === 'RichText') return false;
  if (isContainer(control)) return false;
  // A Flat OPC package is a rich text control's content, not text: no control is filled with it.
  if (holdsFlatOpc(value)) return false;

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

  // an empty value shows the control's placeholder in Word's form: the entry its w:placeholder names, else Word's default
  // entry for a text control, written into the glossary on save (CR-002 section 63, D8), where docx4j writes the bare
  // form, which shows its text only until the control is first written into. Another kind naming no entry keeps docx4j's run.
  if (display === '' && (control.placeholderEntry !== undefined || control.type === 'RichText' || control.type === 'PlainText')) {
    control.showPlaceholder();
    return true;
  }
  control.setBoundContent(runsForValue(display, control.runProperties, control.isMultiLine), display === '');
  return true;
}

/** `applyBindings` over a list of controls. */
export function applyBindingsTo(controls: ContentControl[]): BindingResult {
  const result: BindingResult = { bound: 0, updated: 0, skipped: 0 };
  const notes: string[] = [];
  for (const control of controls) {
    if (!control.xmlMapping.isMapped) continue;
    result.bound++;
    if (applyBindingTo(control, notes)) result.updated++; else result.skipped++;
  }
  if (notes.length > 0) result.notes = notes;
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
  // A picture control is the asynchronous pass's (updateFromPictureControl, CR-005 section 9): its
  // image's bytes go back to the node as base64, as Word 15 writes them (check 38).
  if (control.type === 'Picture') return false;
  // a block control bound by w:dataBinding is plain text to Word, w:richText or not (H4); its text goes back as text
  if (sdtProperty(control.sdt.sdtPr, 'richText') !== undefined && control.type === 'RichText') return false;
  if (isContainer(control)) return false;
  // A control showing its placeholder holds no value.
  if (control.isShowingPlaceholder) return false;
  // The node holds a rich text control's content as Flat OPC: text is not written over it.
  if (holdsFlatOpc(node.text)) return false;

  let value = boundValueOf(control);
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
    // counted by the picture pass, which reads the image part (CR-005 section 9)
    if (control.type === 'Picture') continue;
    result.bound++;
    if (updateFromControl(control)) result.updated++; else result.skipped++;
  }
  return result;
}
