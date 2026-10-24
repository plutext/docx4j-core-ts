// CR-002 phase C: ContentControl, a subset of Office JS's Word.ContentControl over w:sdt in its
// four forms (block, run, row and cell). The typed kinds (checkbox, date, drop-down), the XML
// mapping and insertContentControl are phase E; what is here is what a reader and a text editor
// need. Note (objects CR-002): a run-level w:sdt parsed at body level comes back as SdtBlock, so
// the form is decided by what the content holds as well as by the type.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as w14 from '@docx4j/generated-objects-ts/modules/org_docx4j_w14';
import type * as w15 from '@docx4j/generated-objects-ts/modules/org_docx4j_w15';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import * as w15el from '@docx4j/generated-objects-ts/el/org_docx4j_w15';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { NotSupportedError } from '../../office-js/errors.mjs';
import { GeneralException, InvalidArgument } from './generalException.mjs';
import { marshalString, deepCopy } from '@docx4j/generated-objects-ts';
// CR-002 phase E: the mapping and the typed kinds live in src/model/customxml/; they import
// nothing from here at run time (types only), so this direction is the only one.
import { XmlMapping, type CustomXmlPartLookup } from '../customxml/XmlMapping.mjs';
import {
  CheckboxContentControl, DatePickerContentControl, ListContentControl, PictureContentControl,
  RepeatingSectionContentControl, GroupContentControl, checkboxRun, CHECKBOX_FONT,
} from '../customxml/kinds.mjs';
import { updateFromControl, holdsFlatOpc, placeholderRun, PLACEHOLDER_STYLE } from '../customxml/bindings.mjs';
import { dropExcluded, forWriter, officeItems, type LineBreaks } from './officeText.mjs';
import { declareIgnorable, markHolderDeleted, type ChangeTracker } from './tracking.mjs';
import { placeholderEntryOf, namePlaceholderEntry, newEntryName, entryContentOf, DEFAULT_PLACEHOLDER_ENTRY, DEFAULT_PLACEHOLDER_TEXT, type GlossaryEntries } from './placeholders.mjs';
import { relationshipIdsIn, queueRelationshipRemovals, type RelationshipRemovals } from './relationships.mjs';
import { commentReferencesIn } from './TrackedChange.mjs';
import { wrappingCellOf, paragraphEndOf, endsWithParagraph, endWithParagraph } from './tableRules.mjs';
import { sdtProperty, sdtKindOf, type SdtKind } from '@docx4j/generated-objects-ts/builders/wml';
import { type Element, type BranchMode, typeNameOf, childrenOf, runItemsOf, runsOf, linkParents, runOf, SDT_TYPES, ALTERNATE_CONTENT, BLOCK_LEVEL_TYPES, selectedBranchOf, branchListsOf, inBranch, assertOutsideBranch } from './tree.mjs';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import type { Body } from './Body.mjs';
import type { Paragraph } from './Paragraph.mjs';
import { Range } from './Range.mjs';
import { Table } from './Table.mjs';
import type { SearchOptions } from './search.mjs';

/**
 * Office JS `Word.ContentControlType`, as Word reports it from w:sdtPr: the builders' `SdtKind`
 * (objects CR-003 section 3.1) and Office JS's `Unknown`, which no `w:sdtPr` produces.
 */
export type ContentControlType = SdtKind | 'Unknown';

/** Which of the four w:sdt forms a control is. */
export type ContentControlForm = 'Block' | 'Run' | 'Row' | 'Cell';

/** Office JS `Word.ContentControlAppearance` (w15:appearance). */
export type ContentControlAppearance = 'BoundingBox' | 'Tags' | 'Hidden';

/** The namespaces of the Word 2010 and 2012 content-control properties. */
export const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

/** The markers of a control inserted or deleted under tracking (`markHolderInserted`, `markHolderDeleted`). */
const TRACKED_HOLDER_MARKER = /^customXml(Ins|Del)Range(Start|End)$/;

/** What a block, row or cell control's content ends with, for its range's end: a paragraph, a table or a row. */
const BLOCK_ITEMS = new Set(['org_docx4j_wml.P', 'org_docx4j_wml.Tbl', 'org_docx4j_wml.Tr', 'org_docx4j_wml.Tc']);

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
    const kind = sdtKindOf(this.sdt.sdtPr);
    // Word reads a block control bound by w:dataBinding, untyped or w:richText, as single-line plain text, and saves it
    // so (checks 35 and 56; CR-002 section 61 H4); the file is left as it is
    return kind === 'RichText' && isBoundBlockControl(this) ? 'PlainText' : kind;
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

  /**
   * The control's text as Office JS gives it (CR-012 phase 4): a run-level control's span of
   * `Paragraph.text`; a block-level one's content as `Body.text` reads it, its paragraphs a mark `\r` apart
   * (check 51); a row control's cells a `\t` apart and a cell control's paragraphs `\r` apart (check 53).
   */
  get text(): string {
    if (this.form === 'Run') {
      const span = this.runSpan();
      return span ? span.paragraph.text.substring(span.start, span.end) : '';
    }
    // a row control's cells a \t apart (check 53: ROW15A\tROW15B), a cell control's paragraphs a \r apart; a
    // control's content ending with a row (a row control, or a table last) leaves the row's end out (check 53)
    const text = this.body().text;
    return text.endsWith('\r\n') ? text.slice(0, -2) : text;
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

  /**
   * Office JS `getContentControls` (WordApi 1.5): the content controls `contentControls` lists, those of `options.types`
   * where given. Word's `contentControls` lists rich-text controls only; this package's lists every kind, a recorded
   * departure (CR-002 section 60 G1, check 55b), so this is how add-in code that asks for plain-text controls finds them.
   */
  getContentControls(options?: ContentControlOptions): ContentControl[] {
    return controlsOfTypes(this.contentControls, options);
  }

  /**
   * Text at the start, the end, or replacing the control's text. On a **bound** control the value
   * is written through to the custom XML node as well when the mapping resolves (CR-002 phase E):
   * Word refreshes a bound control from the custom XML part when it opens the document, so an edit
   * that only touched `w:sdtContent` would not show. When the XPath engine is not warm yet the
   * write is left to `pkg.customXmlParts.updateFromContentControls()`.
   */
  insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range {
    this.assertWritable();
    // a control showing its placeholder takes the text in the placeholder's place, wherever it is put (check 58 cases 20
    // to 22, CR-012 section 18; CR-002 section 63)
    if (this.showsPlaceholderContent) return this.writePlaceholder(text);
    const range = this.insertTextInto(text, location);
    this.syncBoundNode();
    return range;
  }

  // --- CR-002 section 63: writing into a control showing its placeholder ---

  /**
   * The glossary entry the control's `w:placeholder` names, which makes it Word's form; undefined where it names none,
   * the bare form, which has no entry to show once emptied, nor for `placeholderText` to read (CR-002 section 63; an
   * extension). Writes are the same in both forms (check 63's run 2).
   */
  get placeholderEntry(): string | undefined {
    return placeholderEntryOf(this.sdt.sdtPr);
  }

  /** @internal True for a block or run-level control showing its placeholder, which a text write replaces (section 63). */
  get showsPlaceholderContent(): boolean {
    return this.isShowingPlaceholder && (this.form === 'Block' || this.form === 'Run');
  }

  /** The package's glossary entries, where the control is in a package. */
  private get glossary(): GlossaryEntries | undefined {
    return (this.parentBody.package_ as { glossaryEntries?: GlossaryEntries } | undefined)?.glossaryEntries;
  }

  /**
   * @internal Text written into the content of this control while it shows its placeholder, by any route: its own
   * `insertText`, its 'Content' range's, a block control's paragraph's (CR-002 section 63, D1, D2 (b), D5, D11). The
   * placeholder is no content, in either form (checks 64 and 67, check 63's run 2): its runs go, untracked, the text takes
   * their place in one paragraph, and the flag is cleared; under tracking only the insertion is recorded. The text takes
   * the control's own run properties (`w:sdtPr/w:rPr`) and nothing of the placeholder's (check 68, cases 1 to 20).
   */
  writePlaceholder(text: string, options: { lineBreaks?: LineBreaks } = {}): Range {
    this.assertWritable();
    const form = this.form;
    // refused before anything changes: a line break into a run-level control (CR-002 section 60 F4)
    if (form === 'Run') textForRunControl(this, dropExcluded(text));
    this.takePlaceholderRuns(form);
    this.isShowingPlaceholder = false;
    let range: Range;
    if (form === 'Run') {
      const paragraph = this.parentParagraph();
      if (!paragraph) throw new Docx4JException('This content control is not in a paragraph');
      range = this.insertIntoEmpty(paragraph, text);
    } else {
      range = this.firstParagraph().splice(0, 0, text, 'before', undefined, options.lineBreaks);
    }
    // every run now in the control is the write's: each takes a copy of the control's run properties (D11)
    const rPr = this.runProperties;
    if (rPr) {
      const holder = form === 'Run' ? (this.sdt.sdtContent as object) : undefined;
      const runs = holder ? runsOf(holder) : this.body().paragraphs.flatMap((p) => runsOf(p.p));
      for (const run of runs) {
        run.value.rPr = deepCopy(rPr);
        linkParents(run.value.rPr, run.value);
      }
    }
    this.syncBoundNode();
    return range;
  }

  /** The control's first paragraph, made where its content has none. */
  private firstParagraph(): Paragraph {
    const body = this.body();
    const first = body.paragraphs[0];
    if (first) return first;
    const p = el.p({ content: [] }) as Element;
    const content = (this.sdt.sdtContent!.content ??= []);
    content.push(p);
    linkParents(p, this.sdt.sdtContent);
    return body.paragraphs[0]!;
  }

  /**
   * Takes the placeholder's runs out of the content: a run-level control's, or a block control's paragraphs', the first
   * paragraph kept, emptied, and any other going. What held them.
   */
  private takePlaceholderRuns(form: ContentControlForm): { runs: Element<wml.R>[]; holder: Element[]; owner: object } {
    const runs: Element<wml.R>[] = [];
    const take = (items: Element[]): void => {
      for (let i = 0; i < items.length; i++) {
        if (typeNameOf(items[i]!) !== 'org_docx4j_wml.R') continue;
        runs.push(items.splice(i, 1)[0] as Element<wml.R>);
        i--;
      }
    };
    if (form === 'Run') {
      const owner = (this.sdt.sdtContent ??= {});
      const holder = (owner.content ??= []);
      take(holder);
      return { runs, holder, owner };
    }
    const content = this.sdt.sdtContent?.content ?? [];
    const paragraphs = content.filter((e) => typeNameOf(e) === 'org_docx4j_wml.P') as Element<wml.P>[];
    for (const p of paragraphs) take((p.value.content ??= []) as Element[]);
    for (const p of paragraphs.slice(1)) content.splice(content.indexOf(p), 1);
    const first = paragraphs[0] ?? (this.firstParagraph().element as Element<wml.P>);
    return { runs, holder: (first.value.content ??= []) as Element[], owner: first.value };
  }

  /**
   * Shows Word's filler for a control emptied in the bare form, where no glossary entry gives it a placeholder: the flag
   * on over a run of five spaces, in the control's own run properties (CR-002 section 63, D10; check 63's run 2, cases 6,
   * 17, 32 and 43, and check 52's case 13). Nothing is marked under tracking. The spaces read as no text while the flag is
   * on (CR-012 section 18).
   */
  private showFiller(): void {
    const { holder, owner } = this.takePlaceholderRuns(this.form);
    const rPr = this.runProperties;
    const run = el.r({ ...(rPr ? { rPr: deepCopy(rPr) } : {}), content: [el.t({ TYPE_NAME: 'org_docx4j_wml.Text', value: '     ', space: 'preserve' })] }) as Element;
    pushBeforeMarker(holder, run);
    linkParents(run, owner);
    this.isShowingPlaceholder = true;
  }

  /**
   * @internal A write through the control's 'Content' or 'Whole' range that the control takes (D3, D4): text into a
   * control showing its placeholder, or into an empty run-level control, goes inside it; a 'Whole' range's Replace over
   * a run-level control showing its placeholder takes the control, the text in its place. Undefined where the range's
   * own write applies.
   */
  writeThrough(part: 'Content' | 'Whole', text: string, location: string): Range | undefined {
    if (part === 'Content') {
      if (this.showsPlaceholderContent) return this.writePlaceholder(text);
      if (this.form !== 'Run' || !this.runSpan()?.empty) return undefined;
      this.assertWritable();
      const range = this.insertTextInto(text, 'Replace');
      this.syncBoundNode();
      return range;
    }
    if (location !== 'Replace' || this.form !== 'Run' || !this.isShowingPlaceholder) return undefined;
    return this.takeWhole(text);
  }

  /**
   * @internal A delete through the control's 'Content' or 'Whole' range that the control takes; false where the range's
   * own delete applies. A 'Content' delete of a control showing its placeholder keeps the control and the flag: in the
   * bare form over Word's filler, five spaces (D10; check 63's run 2, cases 6, 17, 32 and 43, nothing marked under
   * tracking); in Word's form nothing changes, the placeholder being no content. A 'Whole' delete takes a run-level
   * control showing its placeholder (D3).
   */
  deleteThrough(part: 'Content' | 'Whole'): boolean {
    if (part === 'Whole') {
      if (this.form !== 'Run' || !this.isShowingPlaceholder) return false;
      this.takeWhole('');
      return true;
    }
    if (!this.showsPlaceholderContent) return false;
    this.assertWritable();
    if (this.placeholderEntry === undefined) this.showFiller();
    return true;
  }

  /**
   * @internal After a 'Content' delete, untracked, that left the control holding no text: in Word's form it shows its
   * entry again (check 66 case 3); in the bare form Word's filler (D10; check 52's case 13), as for a control of another
   * kind than text naming an entry the glossary lacks.
   */
  showEmptied(): void {
    if (this.isShowingPlaceholder || this.text !== '' || (this.form !== 'Block' && this.form !== 'Run')) return;
    const name = this.placeholderEntry;
    if (name !== undefined && (this.isTextControl || this.glossary?.has(name) === true)) this.showPlaceholder(); else this.showFiller();
  }

  /** @internal A text control (untyped, rich text or plain text), whose default placeholder is Word's default entry. */
  get isTextControl(): boolean {
    return this.type === 'RichText' || this.type === 'PlainText';
  }

  /**
   * @internal Whether an empty control shows a placeholder in Word's form (D8; docx4j CR-063's D3): a text control, or one
   * of another kind naming an entry the glossary holds.
   */
  get showsWordsFormPlaceholder(): boolean {
    if (this.isTextControl) return true;
    const name = this.placeholderEntry;
    return name !== undefined && name !== '' && this.glossary?.has(name) === true;
  }

  /**
   * @internal Takes the control whole (D3): a Replace or delete over a control showing its placeholder, as Word takes one
   * (check 63 cases 12, 18 to 20, 26, 38, 44 to 46 and 52). Untracked it goes, a block control leaving its paragraph,
   * emptied, in its place; tracked it is marked deleted in Word's form (`markHolderDeleted`: `w:customXmlDelRange`
   * markers around it and inside its content), its placeholder gone unmarked and the flag cleared (check 63's run 2,
   * cases 38, 44 to 46 and 52). `text` goes where a run-level control stood, before it under tracking, as Word puts a
   * Replace's insertion first.
   */
  takeWhole(text = ''): Range {
    this.assertWritable();
    const form = this.form;
    const tracker = this.parentBody.changeTracker;
    if (form === 'Run') {
      const span = this.runSpan();
      if (!span) throw new Docx4JException('This content control is not in a paragraph');
      const { paragraph, start } = span;
      if (!tracker) {
        const at = this.container.indexOf(this.element as Element);
        if (at >= 0) this.container.splice(at, 1);
        return text === '' ? new Range(paragraph, start, start) : paragraph.splice(start, start, text);
      }
      const inserted = text === '' ? new Range(paragraph, start, start) : paragraph.splice(start, start, text);
      this.markWholeDeleted(tracker, form);
      return inserted;
    }
    const first = this.firstParagraph();
    if (!tracker) {
      const at = this.container.indexOf(this.element as Element);
      this.takePlaceholderRuns(form);
      const paragraph = first.element as Element;
      const owner = (this.element.value as { PARENT?: object }).PARENT ?? ownerOf(this.container, this.parentBody);
      this.container.splice(at, 1, paragraph);
      linkParents(paragraph, owner);
      return new Range(this.parentBody.paragraphFor(paragraph.value as wml.P) ?? first, 0, 0);
    }
    this.markWholeDeleted(tracker, form);
    return new Range(first, 0, 0);
  }

  /** The control marked deleted with Word's markers, its placeholder gone, unmarked, and the flag cleared (`takeWhole`). */
  private markWholeDeleted(tracker: ChangeTracker, form: ContentControlForm): void {
    markHolderDeleted(tracker, this.element as Element, this.container, () => { this.takePlaceholderRuns(form); });
    this.isShowingPlaceholder = false;
  }

  /**
   * Shows the control's placeholder (CR-002 section 63, D8; section 63.18): a copy of the content of the entry its
   * `w:placeholder` names, as the glossary holds it, the control's own run properties not over it, as Word shows it
   * (check 68, section 63.21), with `w:showingPlcHdr`. A control naming no entry,
   * or a text control naming one the glossary lacks, is made to name Word's default entry, which is written into the
   * glossary on save where it lacks it (D9), as docx4j's CR-063 has it: Word's save drops a reference no entry holds (check
   * 65). A control of another kind keeps the name it has (the binding pass gives such a control docx4j's run instead), and
   * so does any control until the glossary is read (`getBody`, `Word.run`); each shows Word's default text. An extension.
   */
  showPlaceholder(): void {
    this.assertWritable();
    const sdtPr = (this.sdt.sdtPr ??= {} as wml.SdtPr);
    const glossary = this.glossary;
    let name = this.placeholderEntry ?? '';
    // Word's default entry is its text controls' (docx4j CR-063's D3)
    const missing = name === '' || (glossary?.loaded === true && !glossary.has(name));
    if (missing && (this.isTextControl || name === '')) {
      name = DEFAULT_PLACEHOLDER_ENTRY;
      namePlaceholderEntry(sdtPr, name);
    }
    if (name === DEFAULT_PLACEHOLDER_ENTRY) glossary?.require(name, DEFAULT_PLACEHOLDER_TEXT);
    const body = glossary?.bodyOf(name);
    const paragraphs = body ? entryContentOf(body) : [[placeholderRun(undefined, DEFAULT_PLACEHOLDER_TEXT)]];
    this.setBoundContent(paragraphs[0] ?? [], true);
    // a block control shows an entry's further paragraphs too, where it may hold more than one (CR-002 section 61)
    if (this.form !== 'Block' || paragraphs.length < 2 || isBoundBlockControl(this)) return;
    const content = this.sdt.sdtContent!.content ??= [];
    const first = content.find((e) => typeNameOf(e) === 'org_docx4j_wml.P') as Element<wml.P> | undefined;
    let at = first ? content.indexOf(first) + 1 : content.length;
    // the further paragraphs take a copy of the control's own paragraph's properties, as docx4j's binding gives them
    const pPr = first?.value.pPr;
    for (const runs of paragraphs.slice(1)) {
      const p = el.p({ ...(pPr ? { pPr: deepCopy(pPr) } : {}), content: runs as never }) as Element;
      content.splice(at++, 0, p);
      linkParents(p, this.sdt.sdtContent);
    }
  }

  private insertTextInto(text: string, location: 'Start' | 'End' | 'Replace'): Range {
    const form = this.form;
    if ((form === 'Row' || form === 'Cell') && location === 'Replace') {
      throw new Docx4JException(`A ${form.toLowerCase()}-level content control holds ${form === 'Row' ? 'rows' : 'cells'}; replace the text of its cells instead`);
    }
    if (form !== 'Run') return this.body().insertText(text, location);
    text = textForRunControl(this, dropExcluded(text));   // CR-002 section 60: F4, F3
    const span = this.runSpan();
    if (!span) throw new Docx4JException('This content control is not in a paragraph');
    const { paragraph, start, end } = span;
    if (span.empty) return this.insertIntoEmpty(paragraph, text);
    if (location === 'Start') return paragraph.splice(start, start, text);
    if (location === 'End') return paragraph.splice(end, end, text);
    return paragraph.splice(start, end, text);
  }

  /** A paragraph at the start or end of a block control; before or after the paragraph of a run control. */
  insertParagraph(text: string, location: 'Start' | 'End' | 'Before' | 'After'): Paragraph {
    this.assertWritable();
    const form = this.form;
    if ((location === 'Start' || location === 'End') && form === 'Block' && this.isShowingPlaceholder) {
      // CR-002 section 63, D1: the placeholder goes, its paragraph left empty, the flag cleared, and the new paragraph
      // follows, in either form (check 64 cases 6 and 27, check 63's run 2 cases 3 and 29)
      this.takePlaceholderRuns(form);
      this.isShowingPlaceholder = false;
    }
    if ((location === 'Start' || location === 'End') && isBoundBlockControl(this)) {
      // Word holds a bound block control to one paragraph (CR-002 section 61): into a single-line one the text joins
      // its paragraph, as Word's insertParagraph writes it; into a multi-line one Word's own second paragraph makes a
      // file Word can't open, so it is refused (check 56, H2)
      if (this.isMultiLine) {
        throw new NotSupportedError('ContentControl', 'insertParagraph', 'this control is bound by w:dataBinding, and Word holds a bound '
          + "block control to one paragraph: its own second paragraph in a multi-line one makes a file Word can't open (CR-002 section 61, check 56); "
          + 'insertText with a line feed writes a line break');
      }
      const paragraphs = this.paragraphs;
      const p = location === 'Start' ? paragraphs[0] : paragraphs[paragraphs.length - 1];
      if (p) {
        p.insertText(text, location);
        this.syncBoundNode();
        return p;
      }
    }
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
   * The control's range. For a run-level control, exactly the text it holds; for the other forms, its
   * paragraphs: one `Range` from the first one's start to the last one's end (CR-014 phase 1), its 'Whole'
   * ending with the last one's mark, as Office JS's does (check 51: `CCP1\rCCP2\rCCP3\r`, its 'Content'
   * without the last `\r`).
   */
  getRange(location: 'Whole' | 'Start' | 'End' | 'Content' = 'Whole'): Range {
    const range = this.rangeOf(location);
    // its own 'Content' and 'Whole' ranges know the control, for a write into one with no width (CR-002 section 63, D4)
    if ((location === 'Content' || location === 'Whole') && (this.form === 'Run' || this.form === 'Block')) range.ofControl = { control: this, part: location };
    return range;
  }

  private rangeOf(location: 'Whole' | 'Start' | 'End' | 'Content'): Range {
    if (this.form === 'Run') {
      const span = this.runSpan();
      if (!span) throw new Docx4JException('This content control is not in a paragraph');
      const { paragraph, start, end } = span;
      if (location === 'Start') return new Range(paragraph, start, start);
      if (location === 'End') return new Range(paragraph, end, end);
      return new Range(paragraph, start, end);
    }
    const paragraphs = this.paragraphs;
    const first = paragraphs[0];
    if (!first) throw new Docx4JException('This content control holds no paragraph');
    const last = paragraphs[paragraphs.length - 1]!;
    if (location === 'End') return last.getRange('End');
    if (location === 'Start') return first.getRange('Start');
    // a row-level control, or one whose content ends with a table: 'Content' ends with the last cell's text, as
    // the control's text does, and 'Whole' with \t\t, as Word 365 gives them (check 53)
    const lastItem = [...this.content].reverse().find((item) => BLOCK_ITEMS.has(typeNameOf(item) ?? ''));
    if (this.form === 'Row' || (lastItem !== undefined && typeNameOf(lastItem) === 'org_docx4j_wml.Tbl')) {
      return location === 'Whole' ? first.rangeTo(last, true, '\t\t') : first.rangeTo(last, false);
    }
    if (paragraphs.length === 1) return first.getRange(location);
    return new Range(first, 0, last.text.length, true, location === 'Whole', last);
  }

  /**
   * Removes the control; `keepContent` puts what it held in its place, as Office JS. Where the control held the
   * paragraph a table cell ends with, that paragraph stays in its place, emptied, its properties kept, outside
   * any control, as Word keeps a cell's last paragraph (check 52, case 95; CR-002 section 55.6).
   */
  delete(keepContent = false): void {
    this.assertWritable();
    const at = this.container.indexOf(this.element as Element);
    if (at < 0) return;
    const content = keepContent ? [...this.content] : [];
    const owner = ownerOf(this.container, this.parentBody);
    const cell = this.form === 'Block' ? wrappingCellOf(owner) : undefined;
    const end = cell ? paragraphEndOf(cell.content as Element[] | undefined) : undefined;
    this.container.splice(at, 1, ...content);
    if (keepContent) for (const item of content) linkParents(item, owner);
    if (cell && !endsWithParagraph(cell.content as Element[] | undefined)) {
      if (end) {
        end.element.value.content = [];
        this.container.splice(at, 0, end.element as Element);
        linkParents(end.element, owner);
      } else endWithParagraph(cell);
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
    this.assertWritable();
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
  private runSpan(): { paragraph: Paragraph; start: number; end: number; empty?: true } | undefined {
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
    // where the walk passes the control: a control holding no text stands there, after the text before it
    // in the paragraph (before, it was reported at 0, and text put into it went to the paragraph's start)
    let at = 0;
    const segs = paragraph.segments({ visit: (el, offset) => { if (el === (this.element as Element)) at = offset; } });
    for (const seg of segs) {
      if (!inside(seg.run)) continue;
      if (start === undefined) start = seg.start;
      end = seg.end;
    }
    if (start === undefined) return { paragraph, start: at, end: at, empty: true };
    return { paragraph, start, end };
  }

  /**
   * Text into a run control that holds no run: a new run in its content, in a `w:ins` under
   * tracking (Word: check 30, new-control-typed), placed before the control's trailing insertion
   * marker when it is itself an insertion.
   */
  private insertIntoEmpty(paragraph: Paragraph, text: string): Range {
    text = textForRunControl(this, dropExcluded(text));
    const sdt = this.sdt;
    const items = ((sdt.sdtContent ??= {} as { content?: Element[] }).content ??= []);
    const run = runOf(officeItems(text));
    const tracker = paragraph.changeTracker;
    const item: Element = tracker ? tracker.ins([run]) : run;
    const at = items.findIndex((el) => el.name?.localPart === 'customXmlInsRangeStart');
    items.splice(at < 0 ? items.length : at, 0, item);
    linkParents(item, sdt.sdtContent);
    const span = this.runSpan()!;
    return new Range(paragraph, span.start, span.end);
  }

  // --- CR-002 phase E: the XML mapping, the typed kinds and the w:sdtPr properties ---

  /** The control's data binding (Office JS `xmlMapping`), backed by `w:dataBinding`. */
  get xmlMapping(): XmlMapping {
    const pkg = this.parentBody.package_ as { customXmlParts?: CustomXmlPartLookup } | undefined;
    return new XmlMapping(this, pkg?.customXmlParts);
  }

  /**
   * The placeholder Word shows for an empty control: the text of the glossary entry the control's `w:placeholder` names,
   * as Office JS reads it (CR-002 section 63, D8; check 65), `null` where the control names none (the bare form) or the
   * glossary lacks the entry. The glossary is read when the body is (`getBody`, `Word.run`). Setting it names Word's
   * default entry for Word's default text, else a new entry holding the text, named as Word names one (32 hex digits),
   * which is written into the glossary on save, the entry named before left there; a control showing its placeholder,
   * or holding nothing, shows the new text. A control's content is left alone. Until section 63 the text was written as
   * the control's content in the bare form, and a control holding content refused it.
   */
  get placeholderText(): string | null {
    const name = this.placeholderEntry;
    if (name === undefined) return null;
    return this.glossary?.textOf(name) ?? null;
  }
  set placeholderText(text: string) {
    this.assertWritable();
    const own = text !== '' && text !== DEFAULT_PLACEHOLDER_TEXT;
    const name = own ? newEntryName() : DEFAULT_PLACEHOLDER_ENTRY;
    const shown = own ? text : DEFAULT_PLACEHOLDER_TEXT;
    namePlaceholderEntry((this.sdt.sdtPr ??= {} as wml.SdtPr), name);
    this.glossary?.require(name, shown);
    // the entry's run, as the glossary will hold it: the control's own run properties aren't over it (check 68)
    if (this.isShowingPlaceholder || this.text === '') this.setBoundContent([placeholderRun(undefined, shown)], true);
  }

  /**
   * Blocks taken out of the control's content: the relationships only they used (a picture's image) and the comments
   * whose references they held go when the package is saved, as a deletion's do.
   */
  private removeBlocks(content: Element[], blocks: Element[]): void {
    const pkg = this.parentBody.package_ as { pendingRelationshipRemovals?: RelationshipRemovals; pendingCommentRemovals?: Set<number> } | undefined;
    const ids = new Set<string>();
    relationshipIdsIn(blocks, ids);
    if (ids.size > 0) queueRelationshipRemovals(pkg?.pendingRelationshipRemovals, this.parentBody.part, ids);
    if (pkg?.pendingCommentRemovals) commentReferencesIn(blocks, pkg.pendingCommentRemovals);
    for (const block of blocks) {
      const at = content.indexOf(block);
      if (at >= 0) content.splice(at, 1);
    }
  }

  /** Office JS `appearance`: how Word draws the control (`w15:appearance`). */
  get appearance(): ContentControlAppearance {
    const val = (sdtProperty(this.sdt.sdtPr, 'appearance', W15_NS)?.value as { val?: string } | undefined)?.val;
    return val === 'tags' ? 'Tags' : val === 'hidden' ? 'Hidden' : 'BoundingBox';
  }
  set appearance(value: ContentControlAppearance) {
    const val = value === 'Tags' ? 'tags' : value === 'Hidden' ? 'hidden' : 'boundingBox';
    this.putProperty(w15el.appearance({ val }) as Element);
  }

  /** Office JS `color`: the control's colour as '#RRGGBB' (`w15:color`). */
  get color(): string {
    const val = (sdtProperty(this.sdt.sdtPr, 'color', W15_NS)?.value as { val?: string } | undefined)?.val;
    return val === undefined || val === 'auto' ? '' : `#${val.replace(/^#/, '')}`;
  }
  set color(value: string) {
    if (value === '') { this.removeProperty('color'); return; }
    this.putProperty(w15el.color({ val: value.replace(/^#/, '').toUpperCase() }) as Element);
  }

  /** Office JS `cannotDelete`: `w:lock` sdtLocked or sdtContentLocked. */
  get cannotDelete(): boolean {
    const val = this.lockValue();
    return val === 'sdtLocked' || val === 'sdtContentLocked';
  }
  set cannotDelete(value: boolean) {
    this.setLock(value, this.cannotEdit);
  }

  /** Office JS `cannotEdit`: `w:lock` contentLocked or sdtContentLocked. */
  get cannotEdit(): boolean {
    const val = this.lockValue();
    return val === 'contentLocked' || val === 'sdtContentLocked';
  }
  set cannotEdit(value: boolean) {
    this.setLock(this.cannotDelete, value);
  }

  /** Office JS `removeWhenEdited`: `w:temporary`, the control Word drops once it is filled in. */
  get removeWhenEdited(): boolean {
    const item = sdtProperty(this.sdt.sdtPr, 'temporary');
    return item !== undefined && (item.value as wml.BooleanDefaultTrue | undefined)?.val !== false;
  }
  set removeWhenEdited(value: boolean) {
    if (!value) { this.removeProperty('temporary'); return; }
    this.putProperty(el.temporary({ val: true }) as Element);
  }

  /** The checkbox view, for a `w14:checkbox` control; undefined for every other kind, as Office JS. */
  get checkboxContentControl(): CheckboxContentControl | undefined {
    const value = sdtProperty(this.sdt.sdtPr, 'checkbox', W14_NS)?.value as w14.CTSdtCheckbox | undefined;
    return value ? new CheckboxContentControl(this, value) : undefined;
  }

  /** The date view, for a `w:date` control. */
  get datePickerContentControl(): DatePickerContentControl | undefined {
    const value = sdtProperty(this.sdt.sdtPr, 'date')?.value as wml.CTSdtDate | undefined;
    return value ? new DatePickerContentControl(value, () => this.assertWritable()) : undefined;
  }

  /** The list view, for a `w:dropDownList` control. */
  get dropDownListContentControl(): ListContentControl | undefined {
    const value = sdtProperty(this.sdt.sdtPr, 'dropDownList')?.value as wml.CTSdtDropDownList | undefined;
    return value ? new ListContentControl(value, () => this.assertWritable()) : undefined;
  }

  /** The list view, for a `w:comboBox` control. */
  get comboBoxContentControl(): ListContentControl | undefined {
    const value = sdtProperty(this.sdt.sdtPr, 'comboBox')?.value as wml.CTSdtComboBox | undefined;
    return value ? new ListContentControl(value, () => this.assertWritable()) : undefined;
  }

  /** The picture view, for a `w:picture` control. */
  get pictureContentControl(): PictureContentControl | undefined {
    return sdtProperty(this.sdt.sdtPr, 'picture') ? new PictureContentControl(this) : undefined;
  }

  /** The repeating-section view, for a `w15:repeatingSection` control. */
  get repeatingSectionContentControl(): RepeatingSectionContentControl | undefined {
    const value = sdtProperty(this.sdt.sdtPr, 'repeatingSection', W15_NS)?.value as w15.CTSdtRepeatedSection | undefined;
    return value ? new RepeatingSectionContentControl(this, value) : undefined;
  }

  /** The group view, for a `w:group` control (Office JS: an object with no members of its own). */
  get groupContentControl(): GroupContentControl | undefined {
    return sdtProperty(this.sdt.sdtPr, 'group') ? new GroupContentControl() : undefined;
  }

  /**
   * Adds or replaces a `w:sdtPr` child (extension). A `w15` element (`w15:appearance`, `w15:color`,
   * `w15:dataBinding`, the repeating section) puts `w15` in the part's `mc:Ignorable`, as Word's own
   * files have it: Word 2010 opens a part holding one only then (CR-002 section 40).
   */
  putProperty(element: Element): void {
    this.assertWritable();
    const sdt = this.sdt as { sdtPr?: wml.SdtPr };
    sdt.sdtPr ??= {};
    const items = (sdt.sdtPr.rPrOrAliasOrLock ??= []);
    const at = items.findIndex((i) => i.name.localPart === element.name.localPart && i.name.namespaceURI === element.name.namespaceURI);
    if (at >= 0) items[at] = element as never; else items.push(element as never);
    if (element.name.namespaceURI === W15_NS) declareIgnorable(this.parentBody.part, 'w15');
  }

  /** Removes a `w:sdtPr` child (extension). */
  removeProperty(localPart: string, namespaceURI?: string): void {
    this.assertWritable();
    const items = this.sdt.sdtPr?.rPrOrAliasOrLock;
    if (!items) return;
    const at = items.findIndex((i) => i.name.localPart === localPart && (namespaceURI === undefined || i.name.namespaceURI === namespaceURI));
    if (at >= 0) items.splice(at, 1);
  }

  /** The run properties `w:sdtPr/w:rPr` a bound value is written with (docx4j does the same). */
  get runProperties(): wml.RPr | undefined {
    return sdtProperty(this.sdt.sdtPr, 'rPr')?.value as wml.RPr | undefined;
  }

  /** Whether a plain-text control accepts several lines (`w:text/@w:multiLine`). */
  get isMultiLine(): boolean {
    return (sdtProperty(this.sdt.sdtPr, 'text')?.value as wml.CTSdtText | undefined)?.multiLine === true;
  }

  /** Office JS's name for `isShowingPlaceholder` (`showingPlaceholderText`, WordApiDesktop 1.3, read-only there). */
  get showingPlaceholderText(): boolean {
    return this.isShowingPlaceholder;
  }

  /** Whether the control is showing its placeholder rather than a value (`w:showingPlcHdr`). */
  get isShowingPlaceholder(): boolean {
    const item = sdtProperty(this.sdt.sdtPr, 'showingPlcHdr');
    return item !== undefined && (item.value as wml.BooleanDefaultTrue | undefined)?.val !== false;
  }
  set isShowingPlaceholder(value: boolean) {
    if (!value) { this.removeProperty('showingPlcHdr'); return; }
    this.putProperty(el.showingPlcHdr({}) as Element);   // <w:showingPlcHdr/>, as Word writes it
  }

  /**
   * Replaces what the control shows with these runs (docx4j BindingHandler's applyBoundContent):
   * into the first paragraph of a block, row or cell control, keeping its `w:pPr`, and in place of
   * the runs of a run-level one.
   */
  setBoundContent(runs: Element[], showingPlaceholder = false): void {
    this.assertWritable();
    const sdt = this.sdt;
    sdt.sdtContent ??= {};
    const content = ((sdt.sdtContent as { content?: Element[] }).content ??= []);
    const first = content.find((e) => typeNameOf(e) !== undefined);
    if (this.form === 'Run' || (first && typeNameOf(first) === 'org_docx4j_wml.R')) {
      // a tracked control's own markers stay where they are, its inner end first and its inner start last (check 30)
      const ends = content.filter((e) => TRACKED_HOLDER_MARKER.test(e.name?.localPart ?? '') && e.name.localPart.endsWith('End'));
      const starts = content.filter((e) => TRACKED_HOLDER_MARKER.test(e.name?.localPart ?? '') && e.name.localPart.endsWith('Start'));
      content.length = 0;
      content.push(...ends, ...runs, ...starts);
      linkParents(content, sdt.sdtContent as object);
    } else if (!isBoundBlockControl(this)) {
      // a cell or row control, or a block control section 61's rule doesn't cover (a w15:dataBinding, Word's rich-text
      // binding, or none): the value goes into its first paragraph, the rest as it is
      const paragraph = this.paragraphs[0];
      if (paragraph) {
        paragraph.p.content = runs as wml.P['content'];
        linkParents(paragraph.p.content, paragraph.p);
      } else {
        const p = paragraphElementOf(runs);
        content.push(p);
        linkParents(p, sdt.sdtContent as object);
      }
    } else {
      // A bound block control holds one paragraph (CR-002 section 61): the value goes into its first paragraph and any
      // other block it holds goes, as docx4j's bind.xslt writes one paragraph (section 61.6, docx4j-python's review).
      // Before, the first paragraph took the value and the rest stayed, so a control holding two kept them, and Word
      // still couldn't open the file.
      const blocks = content.filter((e) => BLOCK_LEVEL_TYPES.has(typeNameOf(e) ?? ''));
      const paragraph = blocks.find((e) => typeNameOf(e) === 'org_docx4j_wml.P') as Element<wml.P> | undefined;
      if (paragraph) {
        paragraph.value.content = runs as wml.P['content'];
        linkParents(paragraph.value.content, paragraph.value);
      } else {
        const p = paragraphElementOf(runs);
        content.push(p);
        linkParents(p, sdt.sdtContent as object);
      }
      const others = blocks.filter((e) => e !== paragraph);
      if (others.length > 0) this.removeBlocks(content, others);
    }
    this.isShowingPlaceholder = showingPlaceholder;
    // The placeholder run names the PlaceholderText style, which a document made here lacks: the
    // package splices the definition in when it is saved, as it does a new hyperlink's style. Word
    // 2010 drops a w:rStyle naming a style the part does not define (CR-002 section 40).
    if (showingPlaceholder) {
      (this.parentBody.package_ as { requireStyles?: (ids: string[]) => void } | undefined)?.requireStyles?.([PLACEHOLDER_STYLE]);
    }
  }

  /** Shows a checkbox's glyph, as Word writes it: one run in the checkbox font (docx4j checkboxRun). */
  setCheckboxGlyph(symbol: string): void {
    const checkbox = this.checkboxContentControl;
    this.setBoundContent([checkboxRun(symbol, checkbox?.font ?? CHECKBOX_FONT, this.runProperties)]);
  }

  /** A copy of this control right after it, as Word's repeating-section + button does (extension). */
  insertCopyAfter(): ContentControl {
    this.assertWritable();
    const at = this.container.indexOf(this.element as Element);
    if (at < 0) throw new Docx4JException('This content control is not in its container any more');
    const copy: Element = { name: this.element.name, value: deepCopy(this.element.value) as never };
    this.container.splice(at + 1, 0, copy);
    linkParents(copy, ownerOf(this.container, this.parentBody));
    return new ContentControl(copy as Element<wml.SdtBlock>, this.container, this.parentBody);
  }

  /**
   * @internal Writes the control's text into the node it is bound to, when there is one and the XPath engine
   * is warm (CR-002 section 63, D7: every text write into a bound control). Silent when it is not: reading a custom XML part is `pkg.customXmlParts.load()`'s job,
   * and `updateFromContentControls()` is the explicit form of this write.
   */
  syncBoundNode(): void {
    try {
      if (this.xmlMapping.isMapped) updateFromControl(this);
    } catch {
      // the part is not parsed, or the engine is not ready: leave it to updateFromContentControls()
    }
  }

  private lockValue(): string | undefined {
    return (sdtProperty(this.sdt.sdtPr, 'lock')?.value as wml.CTLock | undefined)?.val;
  }

  private setLock(cannotDelete: boolean, cannotEdit: boolean): void {
    const val: wml.STLock | undefined = cannotDelete && cannotEdit ? 'sdtContentLocked'
      : cannotDelete ? 'sdtLocked'
      : cannotEdit ? 'contentLocked'
      : undefined;
    if (val === undefined) { this.removeProperty('lock'); return; }
    this.putProperty(el.lock({ val }) as Element);
  }

  /**
   * True for a control in the selected branch of a kept `mc:AlternateContent`, at any level, which the
   * collectors list as Word reads it (CR-013 phase 2). Such a view is read-only (D4).
   */
  get inBranch(): boolean {
    return inBranch(this.element.value as object);
  }

  /** Refuses a write to a control in a branch (CR-013 D4). Internal: the control's mutators, and its kinds', call it first. */
  assertWritable(): void {
    assertOutsideBranch(this.element.value as object, 'ContentControl');
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

/**
 * Every content control under a list of block-level items, in document order, nested ones included. In a
 * kept `mc:AlternateContent` a reader's (`'read'`, the default) takes the selected branch's controls (CR-013
 * phase 2), the binding passes (`'every'`) every branch's, each other branch's where the element stands
 * (phase 3).
 */
export function collectControls(items: Element[], body: Body, out: ContentControl[], mode: BranchMode = 'read'): void {
  for (const element of items) {
    const tn = typeNameOf(element);
    if (tn === undefined) continue;
    if (SDT_TYPES.has(tn)) {
      const control = new ContentControl(element as Element<wml.SdtBlock>, items, body);
      out.push(control);
      collectControls(control.content, body, out, mode);
      continue;
    }
    if (tn === 'org_docx4j_wml.P') {
      collectRunControls(element.value as wml.P, body, out, mode);
      continue;
    }
    if (tn === 'org_docx4j_wml.Tbl') {
      collectInRows(childrenOf(element.value as object) ?? [], body, out, mode);
      continue;
    }
    if (tn === ALTERNATE_CONTENT) {
      for (const branch of branchListsOf(element.value as mce.AlternateContent, mode)) collectControls(branch, body, out, mode);
      continue;
    }
    const children = childrenOf(element.value as object);
    if (children) collectControls(children, body, out, mode);
  }
}

/** Row-level controls and the controls inside the rows of a table. */
function collectInRows(items: Element[], body: Body, out: ContentControl[], mode: BranchMode): void {
  for (const element of items) {
    const tn = typeNameOf(element);
    if (tn === undefined) continue;
    if (SDT_TYPES.has(tn)) {
      const control = new ContentControl(element as Element<wml.CTSdtRow>, items, body);
      out.push(control);
      collectInRows(control.content, body, out, mode);
      continue;
    }
    if (tn === 'org_docx4j_wml.Tr') collectInCells(childrenOf(element.value as object) ?? [], body, out, mode);
    else if (tn === ALTERNATE_CONTENT) for (const branch of branchListsOf(element.value as mce.AlternateContent, mode)) collectInRows(branch, body, out, mode);
  }
}

/** Cell-level controls and the controls inside the cells of a row. */
function collectInCells(items: Element[], body: Body, out: ContentControl[], mode: BranchMode): void {
  for (const element of items) {
    const tn = typeNameOf(element);
    if (tn === undefined) continue;
    if (SDT_TYPES.has(tn)) {
      const control = new ContentControl(element as Element<wml.CTSdtCell>, items, body);
      out.push(control);
      collectInCells(control.content, body, out, mode);
      continue;
    }
    if (tn === 'org_docx4j_wml.Tc') collectControls(childrenOf(element.value as object) ?? [], body, out, mode);
    else if (tn === ALTERNATE_CONTENT) for (const branch of branchListsOf(element.value as mce.AlternateContent, mode)) collectInCells(branch, body, out, mode);
  }
}

/** The run-level controls in a paragraph (or any run holder), in order, nested ones included. */
export function collectRunControls(holder: object, body: Body, out: ContentControl[], mode: BranchMode = 'read'): void {
  // a kept mc:AlternateContent among runs: its selected branch's (CR-013 phase 2), or every branch's (phase 3)
  if ((holder as { TYPE_NAME?: string }).TYPE_NAME === ALTERNATE_CONTENT) {
    for (const branch of branchListsOf(holder as mce.AlternateContent, mode)) collectRunItems(branch, body, out, mode);
    return;
  }
  const items = runItemsOf(holder);
  if (items) collectRunItems(items, body, out, mode);
}

function collectRunItems(items: Element[], body: Body, out: ContentControl[], mode: BranchMode): void {
  for (const element of items) {
    const tn = typeNameOf(element);
    if (tn === undefined) continue;
    if (SDT_TYPES.has(tn)) {
      const control = new ContentControl(element as Element<wml.SdtRun>, items, body);
      out.push(control);
      collectRunControls(element.value as object, body, out, mode);
      continue;
    }
    if (tn !== 'org_docx4j_wml.R') collectRunControls(element.value as object, body, out, mode);
  }
}

function paragraphElement(text: string): Element {
  const items = officeItems(forWriter(dropExcluded(text), 'one-paragraph'));
  return el.p({ content: items.length === 0 ? [] : [el.r({ content: items as never[] })] }) as Element;
}

/**
 * A view of the content control whose `w:sdt` value this is, found in its owner's list (a paragraph's runs, a body's
 * blocks, a holder's items).
 */
export function controlViewOf(sdt: object, body: Body): ContentControl {
  const owner = (sdt as { PARENT?: Record<string, unknown> }).PARENT;
  for (const value of Object.values(owner ?? {})) {
    if (!Array.isArray(value)) continue;
    const element = (value as Element[]).find((e) => e?.value === sdt);
    if (element) return new ContentControl(element as Element<wml.SdtBlock>, value as Element[], body);
  }
  throw new Docx4JException('This content control is not in the document any more');
}

/**
 * A block-level control bound by `w:dataBinding`, untyped, `w:text` or `w:richText`, whose node doesn't hold flat OPC
 * (CR-002 section 61 H1): Word holds its content to one paragraph, and can't open a file where it holds more (check 56,
 * docx4j-python's file 3). A node not resolved (the custom XML parts not loaded) counts as not flat OPC. `w15:dataBinding`
 * is Word's rich-text binding, and isn't this.
 */
export function isBoundBlockControl(control: ContentControl): boolean {
  if (control.form !== 'Block') return false;
  const pr = control.sdt.sdtPr;
  if (sdtProperty(pr, 'dataBinding') === undefined) return false;
  const kind = sdtKindOf(pr);
  if (kind !== 'RichText' && kind !== 'PlainText') return false;
  let node: { text: string } | undefined;
  try {
    node = control.xmlMapping.customXmlNode;
  } catch {
    node = undefined;   // the parts aren't parsed: the rule applies
  }
  return !(node && holdsFlatOpc(node.text));
}

/**
 * How a control holds lines, as Word reads it: `multi` or `single` for a plain-text control (`w:text` and its
 * `w:multiLine`), `single` for a bound block control Word reads as plain text (section 61 H4); undefined for any other.
 */
export function plainTextLinesOf(control: ContentControl): 'multi' | 'single' | undefined {
  if (sdtProperty(control.sdt.sdtPr, 'text') !== undefined) return control.isMultiLine ? 'multi' : 'single';
  return isBoundBlockControl(control) ? 'single' : undefined;
}

/**
 * Text for a run-level control: a line feed or carriage return there is refused before anything changes (CR-002
 * section 60 F4: Word, given one, moved a bound plain-text control to block level), and a plain-text control's breaks
 * go as its kind takes them (F3: dropped in a single-line one).
 */
export function textForRunControl(control: ContentControl, text: string): string {
  if (/[\r\n]/.test(text)) {
    throw new NotSupportedError('ContentControl', 'insertText', 'a line feed or carriage return in a run-level content control would make a '
      + 'paragraph break inside it, which Word answers by moving the control to block level (CR-002 section 60, F4); write a line break (U+000B) instead');
  }
  const lines = plainTextLinesOf(control);
  return lines === 'single' ? forWriter(text, 'single-line') : text;
}

/** An item put last in a control's content, before the inner start marker a control tracked as inserted ends with (check 30). */
function pushBeforeMarker(items: Element[], item: Element): void {
  const last = items[items.length - 1];
  const marker = last !== undefined && TRACKED_HOLDER_MARKER.test(last.name?.localPart ?? '') && last.name.localPart.endsWith('Start');
  items.splice(marker ? items.length - 1 : items.length, 0, item);
}

/** A `w:p` holding these run-level items (CR-002 phase E: the bound content of an empty control). */
function paragraphElementOf(runs: Element[]): Element {
  return el.p({ content: runs as never }) as Element;
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

/** Word's error for a block write into a bound block control, where measured (CR-002 section 61 H3; check 56). */
export type BlockRefusal = 'InvalidArgument' | 'GeneralException';

/**
 * Refuses, before anything changes, putting blocks into a block-level control bound by `w:dataBinding` (untyped,
 * `w:text` or `w:richText`, its node not flat OPC) that would leave it holding anything but one paragraph: Word can't
 * open such a file, and its repair loses the extra blocks (CR-002 section 61, check 56). Word's own error where
 * measured (`insertTable` `InvalidArgument`, `insertOoxml` `GeneralException`), else `NotSupportedError` (H3).
 */
export function assertControlHolds(body: Body, holder: object, container: Element[], elements: readonly Element[], refusal?: BlockRefusal,
  member: [string, string] = ['Body', 'insertElement']): void {
  const owner = holder as { TYPE_NAME?: string; PARENT?: { TYPE_NAME?: string } } | undefined;
  if (owner?.TYPE_NAME !== 'org_docx4j_wml.SdtContentBlock' || owner.PARENT?.TYPE_NAME !== 'org_docx4j_wml.SdtBlock') return;
  const control = controlViewOf(owner.PARENT, body);
  if (!isBoundBlockControl(control)) return;
  const blocks = [...container, ...elements].filter((e) => BLOCK_LEVEL_TYPES.has(typeNameOf(e) ?? ''));
  if (blocks.length === 1 && typeNameOf(blocks[0]!) === 'org_docx4j_wml.P') return;
  const message = `this content control is bound by w:dataBinding, and Word holds a bound block control to one paragraph: it can't open `
    + `a file where one holds more, and its repair loses them (CR-002 section 61, check 56)`;
  if (refusal === 'InvalidArgument') throw new InvalidArgument(`Table: ${message}`);
  if (refusal === 'GeneralException') throw new GeneralException(`insertOoxml: ${message}`);
  throw new NotSupportedError(member[0], member[1], message);
}

/** Office JS `Word.ContentControlOptions`: the types `getContentControls` lists. */
export interface ContentControlOptions {
  /** `ContentControlType` values; Office JS takes the enum's strings. */
  types?: readonly string[];
}

/** The controls of `options.types`, or all of them. */
export function controlsOfTypes(controls: ContentControl[], options?: ContentControlOptions): ContentControl[] {
  const types = options?.types;
  return types && types.length > 0 ? controls.filter((c) => types.includes(c.type)) : controls;
}

/**
 * Whether a control of this kind made on an empty range shows Word's default placeholder, as Word's does (CR-002 section
 * 63, D8; check 64): a rich text or plain text control. Word's other kinds have entries of their own, unmeasured.
 */
export function showsPlaceholderWhenMade(kind: ContentControlType | undefined): boolean {
  return kind === undefined || kind === 'RichText' || kind === 'PlainText';
}

/** A stand-in for a block-level control about to be made, for `assertControlHolds` before `sdtOf` takes its content. */
export const BLOCK_CONTROL: Element = Object.freeze({
  name: { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localPart: 'sdt' },
  value: { TYPE_NAME: 'org_docx4j_wml.SdtBlock' },
}) as unknown as Element;

/**
 * Whether a block-level control holds one paragraph and nothing else at block level, or nothing yet: what a bound
 * block control may hold (CR-002 section 61).
 */
export function holdsOneParagraph(control: ContentControl): boolean {
  const content = (control.sdt.sdtContent as { content?: Element[] } | undefined)?.content ?? [];
  const blocks = content.filter((e) => BLOCK_LEVEL_TYPES.has(typeNameOf(e) ?? ''));
  return blocks.length === 0 || (blocks.length === 1 && typeNameOf(blocks[0]!) === 'org_docx4j_wml.P');
}

