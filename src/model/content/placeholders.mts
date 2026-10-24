// CR-002 section 63: a content control's placeholder in Word's form (D8, D9). Word keeps a placeholder's text in the
// glossary document part, as a building block the control names in `w:placeholder` (check 64); the control's runs show
// it while `w:showingPlcHdr` is on. A control with no `w:placeholder` is in the bare form, what docx4j's binding writes:
// Word writes into it as into Word's form (check 63's run 2), but has no entry to show once it is emptied, nor for
// `placeholderText` to read, and fills an emptied one with five spaces. The form is the reference's: a control naming any entry,
// present in the glossary or not, is in Word's form while the file is open, though Word's save drops a reference whose
// entry no glossary holds (check 65), so where this package writes a reference it writes the entry too.
//
// The entries are read privately, so an untouched glossary is still saved byte for byte; those written here wait for the
// save, as `requireStyles` does, and go into the glossary then, which is made where the file has none: the glossary
// document part and a styles part holding Placeholder Text, related from it (D9 (a), check 66). Word completes the rest.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import * as f from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { sdtProperty, sdtKindOf, textOf } from '@docx4j/generated-objects-ts/builders/wml';
import { deepCopy } from '@docx4j/generated-objects-ts';
import { linkParents, type Element } from './tree.mjs';
import { uuid } from '../customxml/CustomXmlPartCollection.mjs';
import { PLACEHOLDER_STYLE, placeholderRun } from '../customxml/bindings.mjs';
import { placeholderStyle } from '../../parts/wml/placeholderStyle.mjs';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import type { Part } from '../../parts/Part.mjs';

/** Word's default placeholder entry, which a control Word makes names (checks 64 to 66). */
export const DEFAULT_PLACEHOLDER_ENTRY = 'DefaultPlaceholder_-1854013440';
/** The default entry's text, as Word 365 writes it. */
export const DEFAULT_PLACEHOLDER_TEXT = 'Click or tap here to enter text.';

/** A default placeholder entry: its name and its text. */
export interface DefaultEntry {
  name: string;
  text: string;
}

/**
 * Word's default placeholder entry for each kind of control that has one, as Word 365 makes them from the Developer tab
 * (Jason's `placeholders.docx`, 2026-10-25; CR-002 section 63.22, as docx4j's CR-063 D8 has it): each in the glossary
 * with the default entry's `w:docPartPr`, its body one run in the Placeholder Text style. A picture, a check box and a
 * repeating section have none; the group, citation, equation, bibliography and `w:docPartObj` kinds are unmeasured.
 */
export const DEFAULT_ENTRIES = {
  text: { name: DEFAULT_PLACEHOLDER_ENTRY, text: DEFAULT_PLACEHOLDER_TEXT },
  date: { name: 'DefaultPlaceholder_-1854013437', text: 'Click or tap to enter a date.' },
  list: { name: 'DefaultPlaceholder_-1854013438', text: 'Choose an item.' },
  buildingBlock: { name: 'DefaultPlaceholder_-1854013436', text: 'Choose a building block.' },
  repeatingSectionItem: {
    name: 'DefaultPlaceholder_-1854013435',
    text: 'Enter any content that you want to repeat, including other content controls. You can also insert this control around table rows in order to repeat parts of a table.',
  },
} as const satisfies Record<string, DefaultEntry>;

/** The default entry Word gives a control of this kind (`DEFAULT_ENTRIES`); undefined for a kind it gives none. */
export function defaultEntryFor(sdtPr: wml.SdtPr | undefined): DefaultEntry | undefined {
  switch (sdtKindOf(sdtPr)) {
    case 'RichText': case 'PlainText': return DEFAULT_ENTRIES.text;
    case 'DatePicker': return DEFAULT_ENTRIES.date;
    case 'DropDownList': case 'ComboBox': return DEFAULT_ENTRIES.list;
    case 'BuildingBlockGallery': return sdtProperty(sdtPr, 'docPartList') ? DEFAULT_ENTRIES.buildingBlock : undefined;
    case 'RepeatingSectionItem': return DEFAULT_ENTRIES.repeatingSectionItem;
    default: return undefined;
  }
}

/** The glossary entry a control's `w:placeholder` names; undefined where it names none (the bare form). */
export function placeholderEntryOf(sdtPr: wml.SdtPr | undefined): string | undefined {
  const value = sdtProperty(sdtPr, 'placeholder')?.value as wml.CTPlaceholder | undefined;
  return value === undefined ? undefined : value.docPart?.val ?? '';
}

/**
 * Names `entry` in the control's `w:placeholder`, where Word writes it: after `w:id`, before `w:showingPlcHdr` and the
 * binding (check 64's controls, as Word saved them).
 */
export function namePlaceholderEntry(sdtPr: wml.SdtPr, entry: string): void {
  const items = (sdtPr.rPrOrAliasOrLock ??= []);
  const element = el.placeholder(f.createCTPlaceholder({ docPart: f.createCTPlaceholderDocPart({ val: entry }) })) as never as (typeof items)[number];
  const at = items.findIndex((i) => i.name.localPart === 'placeholder');
  if (at >= 0) { items[at] = element; linkParents(element, sdtPr); return; }
  const after = new Set(['showingPlcHdr', 'dataBinding', 'temporary']);
  const before = items.findIndex((i) => after.has(i.name.localPart));
  items.splice(before < 0 ? items.length : before, 0, element);
  linkParents(element, sdtPr);
}

/** A new entry's name, as Word names one: 32 upper-case hex digits, a new GUID's (check 65). */
export function newEntryName(): string {
  return uuid().replace(/-/g, '').toUpperCase();
}

/** What a package's glossary gives: an entry's text by name. */
export interface GlossaryHost {
  getMainDocumentPart(): { glossaryDocumentPart: XmlPart<wml.GlossaryDocument> | undefined; addTargetPart(part: Part): unknown };
}

/**
 * The placeholder entries of a package's glossary: those its glossary document part holds, read privately (the part
 * stays untouched), and those written since, which go into the glossary when the package is saved (`flush`).
 */
export class GlossaryEntries {
  /** Entries by name, as the glossary part was last read; undefined until `load`. */
  private read: Map<string, GlossaryEntry> | undefined;
  /** The part they were read from. */
  private source: XmlPart<wml.GlossaryDocument> | undefined;
  /** Entries to write on save: name, text. */
  private readonly pending = new Map<string, string>();

  /** Reads the glossary's entries, privately; again only where the part changed. */
  async load(glossary: XmlPart<wml.GlossaryDocument> | undefined): Promise<void> {
    if (glossary === this.source && this.read !== undefined) return;
    this.source = glossary;
    this.read = glossary ? entriesOf(await glossary.readContents().catch(() => undefined)) : new Map();
  }

  /**
   * An entry's text, `null` where neither the glossary nor this session holds it, as Office JS reads it (check 65). An
   * entry's paragraphs are a `\r` apart, as a control's are.
   */
  textOf(name: string): string | null {
    return this.entry(name)?.text ?? null;
  }

  /**
   * An entry's body, the blocks of its `w:docPartBody` (not copies), `null` where neither the glossary nor this session
   * holds it; an entry written this session is one paragraph, one run of its text in the Placeholder Text style.
   */
  bodyOf(name: string): Element[] | null {
    return this.entry(name)?.body ?? null;
  }

  /** True once the glossary has been read (`load`): until then an entry it holds reads as missing. */
  get loaded(): boolean {
    return this.read !== undefined || this.source?.isUnmarshalled === true;
  }

  private entry(name: string): GlossaryEntry | undefined {
    const pending = this.pending.get(name);
    if (pending !== undefined) return { text: pending, body: [el.p({ content: [placeholderRun(undefined, pending) as never] }) as Element] };
    const part = this.source;
    if (part?.isUnmarshalled) return entriesOf(part.contents).get(name);
    return this.read?.get(name);
  }

  /** Whether the glossary holds an entry of this name, or one is waiting to be written. */
  has(name: string): boolean {
    return this.textOf(name) !== null;
  }

  /** An entry to write on save, where the glossary doesn't hold one of that name. */
  require(name: string, text: string): void {
    if (!this.has(name)) this.pending.set(name, text);
  }

  /** True while entries wait for the save. */
  get waiting(): boolean {
    return this.pending.size > 0;
  }

  /** Entries waiting for the save, for a clone. */
  copyTo(target: GlossaryEntries): void {
    for (const [name, text] of this.pending) target.pending.set(name, text);
  }

  /**
   * Writes the waiting entries into the glossary, making it where the file has none (D9 (a): the glossary document
   * part and a styles part holding Placeholder Text, related from it), and puts Placeholder Text into a glossary's
   * styles part that lacks it: an entry whose style its glossary lacks loses it in Word's save (check 65's 65b).
   */
  async flush(host: GlossaryHost, make: { glossary: () => XmlPart<wml.GlossaryDocument>; styles: () => XmlPart<wml.Styles> }): Promise<void> {
    if (this.pending.size === 0) return;
    const main = host.getMainDocumentPart();
    let glossary = main.glossaryDocumentPart;
    if (!glossary) {
      glossary = make.glossary();
      glossary.setContents(f.createGlossaryDocument({ docParts: f.createCTDocParts({ docPart: [] }) }));
      main.addTargetPart(glossary);
    }
    const doc = await glossary.getContents();
    const docParts = (doc.docParts ??= f.createCTDocParts({ docPart: [] }));
    linkParents(docParts, doc);
    const present = entriesOf(doc);
    for (const [name, text] of this.pending) {
      if (present.has(name)) continue;
      const part = docPartOf(name, text);
      docParts.docPart.push(part);
      linkParents(part, docParts);
    }
    this.pending.clear();
    this.source = glossary;
    await ensureGlossaryStyles(glossary, make.styles);
  }
}

/** An entry: its text, its body's paragraphs a `\r` apart, and the body's blocks. */
interface GlossaryEntry {
  text: string;
  body: Element[];
}

/** A glossary's entries by name. */
function entriesOf(doc: wml.GlossaryDocument | undefined): Map<string, GlossaryEntry> {
  const out = new Map<string, GlossaryEntry>();
  for (const part of doc?.docParts?.docPart ?? []) {
    const name = part.docPartPr?.nameOrStyleOrCategory.find((i) => i.TYPE_NAME === 'org_docx4j_wml.CTDocPartName') as wml.CTDocPartName | undefined;
    if (!name?.val) continue;
    const body = (part.docPartBody?.content ?? []) as Element[];
    const paragraphs = body.filter((e) => (e.value as { TYPE_NAME?: string }).TYPE_NAME === 'org_docx4j_wml.P');
    out.set(name.val, { text: paragraphs.map((p) => textOf(p.value as object)).join('\r'), body });
  }
  return out;
}

/**
 * What a control shows of an entry (CR-002 section 63.18, as revised by 63.21): a copy of each of its body's paragraphs'
 * run-level content, every run as the glossary holds it. Word brought an entry back so, the control's own run properties
 * not over it (check 68, cases 21 to 24), and doesn't restyle it (check 65's 65b, case 9). The paragraphs' properties stay
 * in the glossary, whose styles they name.
 */
export function entryContentOf(body: Element[]): Element[][] {
  const out: Element[][] = [];
  for (const block of body) {
    if ((block.value as { TYPE_NAME?: string }).TYPE_NAME !== 'org_docx4j_wml.P') continue;
    const items = ((block.value as wml.P).content ?? []).filter((e) => !['proofErr', 'bookmarkStart', 'bookmarkEnd'].includes((e as Element).name?.localPart ?? '')) as Element[];
    out.push(items.map((item) => ({ name: item.name, value: deepCopy(item.value) }) as Element));
  }
  return out;
}

/**
 * A placeholder entry as Word writes one (check 65): the default entry's `w:docPartPr` (category General, gallery
 * placeholder, type `bbPlcHdr`, behaviour content) with a `w:guid` of its own, and a body of one paragraph, one run in
 * the Placeholder Text style.
 */
function docPartOf(name: string, text: string): wml.CTDocPart {
  const pr = f.createCTDocPartPr({
    nameOrStyleOrCategory: [
      f.createCTDocPartName({ val: name }),
      f.createCTDocPartCategory({ name: f.createCTDocPartCategoryName({ val: 'General' }), gallery: f.createCTDocPartGallery({ val: 'placeholder' }) }),
      f.createCTDocPartTypes({ type: [f.createCTDocPartType({ val: 'bbPlcHdr' })] }),
      f.createCTDocPartBehaviors({ behavior: [f.createCTDocPartBehavior({ val: 'content' })] }),
      f.createCTGuid({ val: `{${uuid().toUpperCase()}}` }),
    ],
  });
  const body = f.createBody({ content: [el.p({ content: [placeholderRun(undefined, text) as never] }) as never] });
  return f.createCTDocPart({ docPartPr: pr, docPartBody: body });
}

/** Default Paragraph Font as Word defines it, which Placeholder Text is based on (check 34's glossary). */
function defaultParagraphFont(): wml.Style {
  return f.createStyle({
    type: 'character', _default: true, styleId: 'DefaultParagraphFont',
    name: f.createStyleName({ val: 'Default Paragraph Font' }),
    uiPriority: f.createStyleUiPriority({ val: 1 }),
    semiHidden: f.createBooleanDefaultTrue(),
    unhideWhenUsed: f.createBooleanDefaultTrue(),
  });
}

/** The glossary's styles part holds Placeholder Text and the style it is based on, made where the glossary has none. */
async function ensureGlossaryStyles(glossary: XmlPart<wml.GlossaryDocument>, make: () => XmlPart<wml.Styles>): Promise<void> {
  const holder = glossary as unknown as { styleDefinitionsPart?: XmlPart<wml.Styles>; addTargetPart(part: Part): unknown };
  let part = holder.styleDefinitionsPart;
  if (!part) {
    part = make();
    part.setContents(f.createStyles({ style: [] }));
    holder.addTargetPart(part);
  } else {
    // a styles part that has them already stays untouched
    const read = await part.readContents().catch(() => undefined);
    const ids = new Set((read?.style ?? []).map((s) => s.styleId));
    if (ids.has('DefaultParagraphFont') && ids.has(PLACEHOLDER_STYLE)) return;
  }
  const styles = await part.getContents();
  const list = (styles.style ??= []);
  const ids = new Set(list.map((s) => s.styleId));
  for (const style of [defaultParagraphFont(), placeholderStyle()]) {
    if (ids.has(style.styleId)) continue;
    list.push(style);
    linkParents(style, styles);
  }
}
