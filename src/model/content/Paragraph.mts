import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { PropertyResolver } from '../properties/PropertyResolver.mjs';
import type { CellContext } from '../properties/table.mjs';
import { deepCopy, marshalString } from '@docx4j/generated-objects-ts';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { NotSupportedError } from '../../office-js/errors.mjs';
import { t as textItem, br as breakItem } from '@docx4j/generated-objects-ts/builders/wml';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import { type Element, segmentsOf, runsOf, runItemsOf, inShowingPlaceholder, linkParents, typeNameOf, W_NS, runOf, paragraphOf, revisionKindOf, objectInSpan, WHOLE, isAlternateContent, selectedBranchOf, inBranch, assertOutsideBranch, type HiddenRuns, type AlternateLocation, type SpanEdges, type SegmentOptions, type TextSegment, type TextViewOptions, type RevisionHolder } from './tree.mjs';
import { relationshipIdsIn, queueRelationshipRemovals, type RelationshipRemovals } from './relationships.mjs';

function withRPr(run: Element<wml.R>, rPr: wml.RPr | undefined): Element<wml.R> {
  if (rPr) run.value.rPr = rPr;
  return run;
}
import { Font, markHolderOf, type FontTracking, type RPrHolder } from './Font.mjs';
import { runFontSelectorOf } from '../fonts/lookup.mjs';
import { Range } from './Range.mjs';
import { officeMatches, expandReplacement, paragraphMarkPieces, refuseMarkReplacement, refuseWildcardMarkReplacement, type SearchOptions } from './search.mjs';
import type { Body, BlockElement } from './Body.mjs';
import { builtInOf, idOfBuiltIn, styleNameOf, styleIdOf } from './styles.mjs';
import { cellOf, type TableCell } from './Table.mjs';
import { ContentControl, collectRunControls, controlViewOf, plainTextLinesOf, textForRunControl, assertControlHolds, controlsOfTypes, holdsOneParagraph, showsPlaceholderWhenMade, BLOCK_CONTROL, type ContentControlType, type ContentControlOptions } from './ContentControl.mjs';
import { breakParagraphAt, BREAKABLE } from './spanEdits.mjs';
import { endsCell, endsWithParagraph, wrappingCellOf } from './tableRules.mjs';
import { dropExcluded, forWriter, linesOf, officeItems, oneParagraphRun, holdsElementCharacters, type LineBreaks } from './officeText.mjs';
import { sdt as sdtOf, nextSdtId, sdtProperty } from '@docx4j/generated-objects-ts/builders/wml';
import { controlIdScope, sdtKindFor, declareSdtExtensions } from '../customxml/insert.mjs';
import { InlinePicture, picturesIn, addImage, writableWidthEmu, type InlinePictureOptions } from './InlinePicture.mjs';
import { contentOf } from './ooxml.mjs';
import { commentApi, type CommentContent, type CommentOptions } from './comments.mjs';
import type { Comment } from './Comment.mjs';
import { ChangeTracker, copyRPr, markDeleted, toDeletedText, highestAnnotationId, canonical, restorePPr, wrapInsertedRuns, markHolderInserted, pruneParagraphProperties, markEquationDeleted, equationRevisionIn, type TrackingHost } from './tracking.mjs';
import { TrackedChange, trackedChangesOfParagraph } from './TrackedChange.mjs';
import {
  type List, type ListItem, type StartListOptions,
  isListItem, listOf, listOrNullObjectOf, listItemOf, listItemOrNullObjectOf,
  startNewList, attachToList, detachFromList,
} from './List.mjs';

/**
 * Office JS `Word.Alignment`, as an *effective* read reports it: an absent `w:jc` resolves to
 * `'Left'`, which is what Word does and what Office JS answers. `'Unknown'` is what a direct
 * read (`formatting({ direct: true })`) says for "no `w:jc` here", and is accepted by the
 * setter, where it removes the element (CR-001 Phase B step 2 decision 4).
 */
export type Alignment = 'Left' | 'Centered' | 'Right' | 'Justified';

/** {@link Alignment} plus the direct reads' "nothing stated" and the setter's "remove it". */
export type AlignmentOrUnknown = Alignment | 'Unknown';

/** Whether a read reports direct formatting instead of the effective value. */
export interface FormattingOptions {
  /**
   * `true` reads the direct formatting only - the `w:pPr` or `w:rPr` on the element itself -
   * as this package did before CR-001 Phase B. The default is the effective value, resolved
   * through the document defaults and the style chain, which is what Office JS reports.
   */
  direct?: boolean;
}

/** Office JS's paragraph formatting properties, read in one call (extension). */
export interface ParagraphFormatting {
  alignment: AlignmentOrUnknown;
  leftIndent: number;
  rightIndent: number;
  firstLineIndent: number;
  spaceBefore: number;
  spaceAfter: number;
  lineSpacing: number;
  outlineLevel: number;
}

const TWIPS_PER_POINT = 20;

/** Office JS's paragraph formatting read from a `w:pPr`, resolved (`effective`) or direct. */
function formattingOf(pPr: wml.PPr | undefined, effective: boolean): ParagraphFormatting {
  return {
    alignment: alignmentOf(pPr, effective),
    leftIndent: (pPr?.ind?.left ?? 0) / TWIPS_PER_POINT,
    rightIndent: (pPr?.ind?.right ?? 0) / TWIPS_PER_POINT,
    firstLineIndent: firstLineIndentOf(pPr),
    spaceBefore: (pPr?.spacing?.before ?? 0) / TWIPS_PER_POINT,
    spaceAfter: (pPr?.spacing?.after ?? 0) / TWIPS_PER_POINT,
    lineSpacing: lineSpacingOf(pPr),
    outlineLevel: outlineLevelOf(pPr),
  };
}

/** The `w:pPr` members `formattingOf` reads; everything else is compared as it is written. */
const FORMATTING_MEMBERS = new Set(['jc', 'ind', 'spacing', 'outlineLvl']);

/**
 * A subset of Office JS `Word.Paragraph` over a `w:p` in a body, cell, header, footer or
 * content control. A light view: it holds the element and the array that contains it; nothing
 * is cached, so views may be created and discarded freely. Every write mutates the tree in
 * place and links PARENT on what it adds.
 */
export class Paragraph {
  constructor(
    /** The `w:p` element pair. */
    readonly element: Element<wml.P>,
    /** The content array holding the element (a body's, a cell's, a content control's). */
    readonly container: Element[],
    /** The nearest body view (the part's), for addresses and paragraph ids. */
    readonly parentBody: Body,
  ) {}

  /** The paragraph value (docx4j: the P). */
  get p(): wml.P {
    return this.element.value;
  }

  /** @internal A view of another paragraph of this one's body, for a `Range` across paragraphs (CR-014). */
  viewOf(element: Element<wml.P>, container: Element[]): Paragraph {
    return new Paragraph(element, container, this.parentBody);
  }

  /**
   * @internal A range from this paragraph's start to the end of `last`, a later paragraph of the same story, with
   * `last`'s mark where `mark` is set, or `endMark` in its place (CR-014 phase 1).
   */
  rangeTo(last: Paragraph, mark: boolean, endMark?: string): Range {
    return new Range(this, 0, last.text.length, true, mark, last.element === this.element ? undefined : last, endMark);
  }

  /** How many tables deep the paragraph is: 0 outside a table, 1 in a cell, 2 in a nested table's (Office JS `tableNestingLevel`). */
  get tableNestingLevel(): number {
    let level = 0;
    for (let cell = this.parentTableCell; cell; cell = cell.parentTable.parentTableCell) level++;
    return level;
  }

  get index(): number {
    return this.container.indexOf(this.element);
  }

  /**
   * The paragraph's text as Office JS gives it in Word's default view (CR-012 phase 4, section 6.2):
   * the accepted text, hidden text left out, in Word's characters (check 48: a note reference U+0002,
   * a line break U+000B, a page break U+000C, a column break U+000E, a `w:cr` U+000D, the hyphens
   * U+001E and U+001F, a symbol `(`, an equation as Word draws it). A kept `mc:AlternateContent`
   * reads as its selected branch (check 49). Every offset a `Range` or `search` takes counts in it.
   */
  get text(): string {
    return this.segments().map((s) => s.text).join('');
  }

  /** Replaces the content with one run of the text, keeping the first run's formatting. */
  set text(value: string) {
    this.assertWritable();
    // one paragraph, its line breaks w:br and w:cr (CR-002 section 60 E2), or as a plain-text control takes them (F3)
    const lines = this.blockControl ? plainTextLinesOf(this.blockControl) : undefined;
    const mode: LineBreaks = lines === 'multi' ? 'multi-line' : lines === 'single' ? 'single-line' : 'one-paragraph';
    // in a block control showing its placeholder, the text in the placeholder's place (CR-002 section 63)
    const placeholder = this.placeholderControl;
    if (placeholder) { placeholder.writePlaceholder(value, { lineBreaks: mode }); return; }
    if (this.changeTracker) { this.splice(0, this.text.length, value, 'before', WHOLE, mode); return; }
    const rPr = runsOf(this.p)[0]?.value.rPr;
    this.releaseRelationships(this.p.content);
    this.p.content = [oneParagraphRun(value, rPr ? deepCopy(rPr) : undefined, mode)] as wml.P['content'];
    linkParents(this.p.content, this.p);
  }

  /**
   * The text in one of the two views (extension; `text` is the accepted one). `{ view: 'original' }`
   * reads the document as it was before the tracked changes: `w:del` included, `w:ins` excluded.
   */
  getText(options?: TextViewOptions): string {
    return options?.view === 'original' ? this.segments({ view: 'original' }).map((s) => s.text).join('') : this.text;
  }

  /** The style id (w:pStyle; docx4j's name for it); 'Normal' when none is set. Extension: Office JS has `style` and `styleBuiltIn` only. */
  get styleId(): string {
    return this.p.pPr?.pStyle?.val ?? 'Normal';
  }
  set styleId(id: string) {
    // through pPr(), so that a tracked write records w:pPrChange first (CR-002 phase F)
    const pPr = this.pPr();
    if (id === '' || id === 'Normal') delete pPr.pStyle;
    else pPr.pStyle = { val: id };
    this.settleFormatting();
  }

  /**
   * The style's display name ('Heading 1', 'My Style'), as Office JS: from the styles part's
   * `w:name` when that part is unmarshalled, else derived from the id. Setting accepts a display
   * name, a stored name, or an id.
   */
  get style(): string {
    return styleNameOf(this.parentBody.package_, this.styleId);
  }
  set style(name: string) {
    this.styleId = styleIdOf(this.parentBody.package_, name);
  }

  /** The `Word.Style` value ('Heading1'), or 'Other' when the style is not a built-in one, as Office JS. */
  get styleBuiltIn(): string {
    return builtInOf(this.styleId);
  }
  set styleBuiltIn(value: string) {
    this.styleId = idOfBuiltIn(value);
  }

  /**
   * The paragraph properties which actually apply: the document defaults, the style chain and
   * the numbering level, then this paragraph's own `w:pPr` (docx4j
   * `PropertyResolver.getEffectivePPr`). A live, possibly cached object: copy it before
   * changing it. Needs the package's resolver; see `Body.propertyResolver`.
   */
  get effectivePPr(): wml.PPr {
    return this.parentBody.propertyResolver.getEffectivePPr(this.p.pPr, this.cellContext);
  }

  /** The paragraph mark's effective run properties (docx4j `getEffectiveParagraphMarkRPr`). */
  get effectiveParagraphMarkRPr(): wml.RPr {
    return this.parentBody.propertyResolver.getEffectiveParagraphMarkRPr(this.p.pPr, this.cellContext);
  }

  /**
   * Where this paragraph sits in a table, as far as its formatting goes (docx4j
   * `PropertyResolver.cellContextOf`; CR-007): the table style its table resolves to and the
   * conditional formats it is under. Undefined outside a table, and in a text box, note or comment
   * reached from a cell. Every effective read of the paragraph and of its runs passes it, so a run
   * in a header row the table style makes bold reads bold. Each read walks the table's rows; code
   * resolving every paragraph of a large table should walk with a `TableContextTracker` instead.
   * Needs the package's resolver, as the effective reads do.
   */
  get cellContext(): CellContext | undefined {
    return this.parentBody.propertyResolver.cellContextOf(this.p);
  }

  /** The `w:pPr` a formatting read works from: the resolved one, or this paragraph's own. */
  private readPPr(options?: FormattingOptions): wml.PPr | undefined {
    return options?.direct === true ? this.p.pPr : this.effectivePPr;
  }

  /**
   * Office JS's paragraph formatting in one call, and the way to ask for direct formatting:
   * `formatting({ direct: true })` is what the individual getters reported before CR-001
   * Phase B step 2 (an absent property reads 0, 10 or `'Unknown'`).
   */
  formatting(options?: FormattingOptions): ParagraphFormatting {
    return formattingOf(this.readPPr(options), options?.direct !== true);
  }

  /** The effective alignment; an absent `w:jc` after resolution is `'Left'`, as Office JS. */
  get alignment(): Alignment {
    return alignmentOf(this.effectivePPr, true) as Alignment;
  }
  set alignment(v: AlignmentOrUnknown) {
    const pPr = this.pPr();
    const val: wml.JcEnumeration | undefined = v === 'Left' ? 'left' : v === 'Centered' ? 'center' : v === 'Right' ? 'right' : v === 'Justified' ? 'both' : undefined;
    if (val === undefined) delete pPr.jc;
    else if (this.inheritedAlignment() === v) delete pPr.jc;
    else pPr.jc = { val };
    this.settleFormatting();
    pruneParagraphProperties(this.p);                                  // no empty w:pPr left (the editor's finding on 0.3.1)
  }

  /**
   * The alignment the paragraph resolves to with no `w:jc` of its own: its style's, else the
   * defaults'. Word's Ctrl+L on a centred paragraph whose style is left-aligned removes the `w:jc`
   * rather than writing `w:jc="left"`, with tracking on (check 26) and off (the editor's
   * `measure-word2010.docx` and `measure-word15.docx`, 2026-09-30; CR-002 section 38), so a value the
   * paragraph already inherits is not written. Measured for left over a left-aligned style; taken to
   * hold for any value. Undefined without a resolver, where the value is written as asked.
   */
  private inheritedAlignment(): Alignment | undefined {
    const resolver = (this.parentBody.package_ as { propertyResolverOrUndefined?: PropertyResolver } | undefined)?.propertyResolverOrUndefined;
    if (!resolver) return undefined;
    const { jc: _jc, pPrChange: _change, PARENT: _parent, ...rest } = this.p.pPr as wml.PPr & { PARENT?: object };
    // in a table cell what is inherited includes the table style's (CR-007)
    return alignmentOf(resolver.getEffectivePPr(rest as wml.PPr, resolver.cellContextOf(this.p)), true) as Alignment;
  }

  /** Indents and spacing in points, as Office JS; the effective values (CR-001 Phase B step 2). */
  get leftIndent(): number { return (this.effectivePPr.ind?.left ?? 0) / TWIPS_PER_POINT; }
  set leftIndent(pt: number) { this.ind().left = Math.round(pt * TWIPS_PER_POINT); this.settleFormatting(); }
  get rightIndent(): number { return (this.effectivePPr.ind?.right ?? 0) / TWIPS_PER_POINT; }
  set rightIndent(pt: number) { this.ind().right = Math.round(pt * TWIPS_PER_POINT); this.settleFormatting(); }
  /** First-line indent; negative for a hanging indent, as Office JS. */
  get firstLineIndent(): number {
    return firstLineIndentOf(this.effectivePPr);
  }
  set firstLineIndent(pt: number) {
    const ind = this.ind();
    if (pt < 0) { ind.hanging = Math.round(-pt * TWIPS_PER_POINT); delete ind.firstLine; } else { ind.firstLine = Math.round(pt * TWIPS_PER_POINT); delete ind.hanging; }
    this.settleFormatting();
  }
  get spaceBefore(): number { return (this.effectivePPr.spacing?.before ?? 0) / TWIPS_PER_POINT; }
  set spaceBefore(pt: number) { this.spacing().before = Math.round(pt * TWIPS_PER_POINT); this.settleFormatting(); }
  get spaceAfter(): number { return (this.effectivePPr.spacing?.after ?? 0) / TWIPS_PER_POINT; }
  set spaceAfter(pt: number) { this.spacing().after = Math.round(pt * TWIPS_PER_POINT); this.settleFormatting(); }
  /** Line spacing in points: w:line/240 lines of 12pt when the rule is auto, else w:line twips; 0 when not set. */
  get lineSpacing(): number {
    return lineSpacingOf(this.effectivePPr);
  }
  set lineSpacing(pt: number) { const s = this.spacing(); s.line = Math.round(pt * TWIPS_PER_POINT); s.lineRule = 'exact'; this.settleFormatting(); }
  /** w:outlineLvl + 1 (1 to 9); 10 for body text, as Office JS. A heading style supplies it. */
  get outlineLevel(): number { return outlineLevelOf(this.effectivePPr); }
  set outlineLevel(level: number) { const pPr = this.pPr(); if (level >= 10 || level < 1) delete pPr.outlineLvl; else pPr.outlineLvl = { val: level - 1 }; this.settleFormatting(); }

  /**
   * The runs' formatting: reads the *effective* properties of the runs holding the paragraph's text, `null` (or `''`,
   * `Mixed`) where they differ, and the mark's where it has no text, as Office JS reads them (CR-002 section 59); writes
   * direct formatting to every run (extension: `Font` is over runs, not the paragraph mark).
   */
  get font(): Font {
    return this.getFont();
  }

  /** `font`, with `{ direct: true }` for the runs' direct formatting (extension). */
  getFont(options?: FormattingOptions): Font {
    const holders = (write: boolean): wml.R[] => {
      if (write) { this.assertWritable(); return runsOf(this.p).map((r) => r.value); }
      // a read takes the runs of Paragraph.text, Word's default view, as a range over it does; with none, a placeholder's
      // runs, which Word's font reads though its text doesn't (check 52 case 13, after Accept All: Word's bold filler)
      const runs = this.getRange('Content').runs.map((r) => r.value);
      return runs.length > 0 ? runs : runsOf(this.p).map((r) => r.value).filter((r) => inShowingPlaceholder(r));
    };
    const mark = (): RPrHolder => markHolderOf(this.p);
    if (options?.direct === true) return new Font(holders, () => this.fontTracking(), undefined, undefined, mark);
    return new Font(holders, () => this.fontTracking(),
      (rPr) => this.parentBody.propertyResolver.getEffectiveRPr(rPr, this.p.pPr, this.cellContext),
      (rPr) => runFontSelectorOf(this.parentBody.package_)?.asciiFontName(rPr), mark,
      () => this.parentBody.propertyResolver.getThemeColors());
  }

  /** w14:paraId, the stable address Word gives paragraphs. */
  get paraId(): string | undefined {
    return this.p.paraId;
  }
  set paraId(id: string | undefined) {
    this.assertWritable();
    if (id === undefined) delete this.p.paraId; else this.p.paraId = id;
  }

  /** The runs, direct and nested (hyperlinks, content controls, insertions), in order. */
  get runs(): Element<wml.R>[] {
    return runsOf(this.p);
  }

  /** The cell this paragraph is in, or undefined when it is not in a table (Office JS). */
  get parentTableCell(): TableCell | undefined {
    return cellOf(this.p, this.parentBody);
  }

  /** The run-level content controls in this paragraph, in order, nested ones included. */
  get contentControls(): ContentControl[] {
    const out: ContentControl[] = [];
    collectRunControls(this.p, this.parentBody, out);
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
   * The inline pictures in this paragraph, in order (Office JS Paragraph.inlinePictures): each
   * `w:drawing` whose `wp:inline` holds a `pic:pic`, in a run or in the selected branch of a kept
   * `mc:AlternateContent` (`picturesIn`). An inline chart, text box or group is not one (CR-010
   * section 12.6).
   */
  get inlinePictures(): InlinePicture[] {
    const out: InlinePicture[] = [];
    for (const run of runsOf(this.p)) {
      for (const item of (run.value.content ?? []) as Element[]) {
        for (const drawing of picturesIn(item)) out.push(new InlinePicture(drawing, run, this, drawing === item ? undefined : item));
      }
    }
    return out;
  }

  /** Text inserted at the start, the end, or replacing the whole paragraph's text. */
  insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range {
    // in a block control showing its placeholder the text takes the placeholder's place wherever it is put (CR-002
    // section 63; check 63's run 2, cases 7 to 9 and 33 to 35)
    const placeholder = this.placeholderControl;
    if (placeholder) return placeholder.writePlaceholder(text);
    const length = this.text.length;
    if (location === 'Start') return this.splice(0, 0, text);
    if (location === 'End') return this.writeAtEnd(text);
    // tracked, the insertion goes first, before the deletion, as a range's Replace writes it and Word wrote this one
    // (check 63 cases 45 and 52; CR-002 section 63 D3)
    if (this.changeTracker && length > 0) return this.getRange('Content').insertText(text, 'Replace');
    return this.splice(0, length, text, 'before', WHOLE);
  }

  /**
   * A new paragraph of the text before or after this one, with this paragraph's properties (as Word). Beside the only
   * paragraph of a block control showing its placeholder it goes outside the control, as Word's went, and the control
   * stays at block level (CR-002 section 63, D6 (b); check 63 cases 10, 11, 36 and 37).
   */
  insertParagraph(text: string, location: 'Before' | 'After'): Paragraph {
    this.assertWritable();
    const pPr = this.p.pPr ? deepCopy(this.p.pPr) : undefined;
    if (pPr) { delete pPr.rPr; delete (pPr as { sectPr?: unknown }).sectPr; delete pPr.pPrChange; }
    const run = oneParagraphRun(text);   // E2
    const el = paragraphOf(run.value.content?.length ? [run] : [], pPr);
    const control = this.blockControl;
    if (control?.isShowingPlaceholder && holdsOneParagraph(control)) {
      // the body holding the control: this paragraph's is the control's content
      const owner = (control.element.value as { PARENT?: { content?: Element[] } }).PARENT ?? this.parentBody.container;
      return this.parentBody.sub(owner).insertElement(el, location, { element: control.element as Element, container: control.container }) as Paragraph;
    }
    return this.parentBody.insertElement(el, location, this) as Paragraph;
  }

  /** The block-level control holding this paragraph directly while it shows its placeholder (CR-002 section 63). */
  private get placeholderControl(): ContentControl | undefined {
    const control = this.blockControl;
    return control?.showsPlaceholderContent ? control : undefined;
  }

  insertBreak(type: 'Page' | 'Line', location: 'Before' | 'After' | 'Start' | 'End'): void {
    this.assertWritable();
    const brk = breakItem(type === 'Page' ? 'page' : undefined);
    if (location === 'Before' || location === 'After') {
      const p = this.insertParagraph('', location);
      const run = runOf([brk]);
      const tracker = p.changeTracker;
      const item = tracker ? tracker.ins([run]) : run;
      p.p.content!.push(item as never);
      linkParents(item, p.p);
      return;
    }
    const run = runOf([brk]);
    const tracker = this.changeTracker;
    const item = tracker ? tracker.ins([run]) : run;
    const content = (this.p.content ??= []);
    if (location === 'Start') content.unshift(item as never); else content.push(item as never);
    linkParents(item, this.p);
  }

  /**
   * Adds the image as a part of this paragraph's part, with a relationship, and shows it at the
   * start or the end of this paragraph, or in place of its content (Office JS).
   */
  insertInlinePictureFromBase64(base64: string, location: 'Start' | 'End' | 'Replace', options?: InlinePictureOptions): InlinePicture {
    this.assertWritable();
    const part = this.parentBody.part;
    if (!part) throw new Docx4JException('This body has no part, so an image cannot be related to it');
    const tracker = this.changeTracker;
    // refused before the image part is added, not after (CR-012 section 12.4)
    if (location === 'Replace' && tracker) this.assertPlaceable(0, this.text.length, 'replace');
    const { run, drawing } = addImage(part, base64, this.parentBody.container, writableWidthEmu(this.parentBody.container), options);
    // under tracking the picture is an insertion, and what it replaces is deleted as insertText's is
    if (location === 'Replace' && tracker) this.splice(0, this.text.length, '', 'before', WHOLE);
    const content = (this.p.content ??= []) as Element[];
    if (location === 'Replace' && !tracker) { this.releaseRelationships(content); content.length = 0; }
    const item = tracker ? tracker.ins([run]) : run;
    if (location === 'Start') content.unshift(item as Element); else content.push(item as Element);
    linkParents(item, this.p);
    return new InlinePicture(drawing, run, this);
  }

  /**
   * Word's `insertOoxml` at paragraph level: a flat OPC `pkg:package` (or a bare fragment). One
   * incoming paragraph is merged into this one at 'Start' and 'End', as Word's paste does;
   * anything else is inserted as blocks before or after. 'Replace' puts the content where this
   * paragraph was and removes it.
   */
  async insertOoxml(ooxml: string, location: 'Before' | 'After' | 'Start' | 'End' | 'Replace'): Promise<(Paragraph | BlockElement)[]> {
    this.assertWritable();
    const preprocess = this.parentBody.package_?.loadOptions.preprocessor;
    const elements = await contentOf(ooxml, { preprocess: preprocess ? (doc) => preprocess(doc) : undefined, target: this.parentBody.part });
    if (elements.length === 0) return [];
    const only = elements.length === 1 && typeNameOf(elements[0]!) === 'org_docx4j_wml.P' ? (elements[0]!.value as wml.P) : undefined;
    if (only && (location === 'Start' || location === 'End')) {
      this.insertItemsAt(location === 'Start' ? 0 : this.text.length, (only.content ?? []) as Element[]);
      return [this];
    }
    if (location === 'Replace') {
      const inserted = this.parentBody.insertElement(elements, 'Before', this, 'GeneralException', this.element);
      void inserted;
      this.delete();
    } else {
      this.parentBody.insertElement(elements, location === 'Start' ? 'Before' : location === 'End' ? 'After' : location, this, 'GeneralException');
    }
    return elements.map((el) => typeNameOf(el) === 'org_docx4j_wml.P'
      ? new Paragraph(el as Element<wml.P>, this.container, this.parentBody)
      : { element: el, container: this.container });
  }

  /** Matches within this paragraph. */
  search(text: string, options?: SearchOptions): Range[] {
    // Word's paragraph mark, ^p, is searched in the paragraph's whole range, which holds its mark (CR-014 phase 3), and so
    // is a wildcard query, whose ? or class matches the mark too (check 59; CR-002 section 32, K7)
    if (options?.matchWildcards || (!options?.matchRegExp && paragraphMarkPieces(text))) return this.getRange('Whole').search(text, options);
    return this.searchMatches(text, options).map((m) => m.range);
  }

  /**
   * `search`, keeping each match's capture groups, which is what a replacement's `$1` or Word's `\1`
   * expands from (CR-002 section 26). Not an Office JS member; `replaceText` is its only caller here
   * and `Body.replaceText` the other.
   */
  searchMatches(text: string, options?: SearchOptions): { range: Range; match: RegExpExecArray }[] {
    return officeMatches(this.segments({ objects: true }), text, options)
      .map(({ start, end, match }) => ({ range: new Range(this, start, end), match }));
  }

  /**
   * Replaces every match of `find` with `replace`, last match first so that the offsets of the
   * ones still to do stay valid (CR-002 section 3.7). Returns the number replaced; tracked when
   * the package's `changeTrackingMode` is on.
   */
  replaceText(find: string, replace: string, options?: SearchOptions): number {
    refuseMarkReplacement('Paragraph', find, options);
    if (options?.matchWildcards) refuseWildcardMarkReplacement('Paragraph', this.search(find, options));
    const matches = this.searchMatches(find, options);
    for (let i = matches.length - 1; i >= 0; i--) {
      const { range, match } = matches[i]!;
      range.insertText(expandReplacement(replace, match, options), 'Replace');
    }
    return matches.length;
  }

  /** The tracked changes in this paragraph, in document order (CR-002 phase F). */
  getTrackedChanges(): TrackedChange[] {
    return trackedChangesOfParagraph(this);
  }

  /**
   * The paragraph's range: 'Whole' and 'Content' hold its whole content, the pictures at its start and
   * end among them, so deleting it deletes them (CR-002 section 48); 'Start' and 'End' are collapsed.
   * 'Whole' holds the paragraph's mark too, so its `text` ends with `markText`, as Office JS's does
   * (check 48); 'Content' ends with the paragraph's text.
   */
  getRange(location: 'Whole' | 'Start' | 'End' | 'Content' = 'Whole'): Range {
    const length = this.text.length;
    if (location === 'Start') return new Range(this, 0, 0);
    if (location === 'End') return new Range(this, length, length).atEndOf('End');
    const range = new Range(this, 0, length, true, location === 'Whole');
    return location === 'Content' ? range.atEndOf('Content') : range;
  }

  /**
   * The character Office JS gives for this paragraph's mark at the end of its 'Whole' range (check 48):
   * `\t` for the last paragraph of a table cell, the cell's end; nothing for a mark tracked as deleted, which
   * Word's default view doesn't show (CR-012 section 19, check 58); and `\r` for any other.
   */
  get markText(): string {
    if (lastInCell(this.p)) return '\t';
    return this.p.pPr?.rPr?.del !== undefined ? '' : '\r';
  }

  /**
   * Removes the paragraph from its container. While the package tracks changes, the content
   * becomes a `w:del` and the mark is marked deleted (`w:pPr/w:rPr/w:del`) instead, unless the
   * whole paragraph was this author's own insertion, which Word simply takes back.
   * The paragraph a table cell ends with, as Word reads it, stays where removing it would leave the
   * cell without one, emptied, its properties kept; tracked, its mark is never marked deleted, as Word
   * keeps a cell's last paragraph (check 52, cases 88 to 91; CR-002 section 55.6).
   */
  delete(): void {
    this.remove(true);
  }

  /**
   * `delete()`; with `keepCellEnd` false a cell's last paragraph goes as any other, for `Body.clearContent`'s
   * callers, which put their own paragraph in next.
   * @internal
   */
  remove(keepCellEnd: boolean): void {
    this.assertWritable();
    const cellEnd = keepCellEnd && endsCell(this.p);
    const tracker = this.changeTracker;
    if (tracker) { this.trackedDelete(tracker, cellEnd); return; }
    const i = this.index;
    if (i < 0) return;
    this.container.splice(i, 1);
    this.releaseRelationships(this.element);
    if (cellEnd && !this.cellStillEnds()) {
      this.p.content = [];
      this.container.splice(i, 0, this.element as Element);
    }
  }

  /** Whether the cell this paragraph was in ends with a paragraph without it, as Word reads it (CR-002 section 55). */
  private cellStillEnds(): boolean {
    return endsWithParagraph(wrappingCellOf((this.p as { PARENT?: object }).PARENT)?.content as Element[] | undefined);
  }

  private trackedDelete(tracker: ChangeTracker, cellEnd: boolean): void {
    if (markDeleted(this.p)) throw new Docx4JException('This paragraph is already marked deleted');
    // the whole content, its pictures too: before, a picture's run, having no text, stayed out of the
    // w:del, and accepting joined it into the next paragraph (CR-002 section 48)
    this.deleteText(tracker, 0, this.text.length, WHOLE);
    const rPr = this.p.pPr?.rPr;
    const ownMark = rPr?.ins !== undefined && rPr.ins.author === tracker.author;
    if (ownMark && (this.p.content ?? []).length === 0) {
      const i = this.index;
      if (i >= 0) {
        this.container.splice(i, 1);
        if (!cellEnd || this.cellStillEnds()) return;
        this.container.splice(i, 0, this.element as Element);
      }
      return;
    }
    // a cell's last mark: Word deletes the paragraph's text and never marks the mark (check 52, cases 90 and 91)
    if (cellEnd) return;
    tracker.markParagraphDeleted(this.p);
    // Word, deleting a paragraph mark ALONE under tracking, gives the next paragraph this one's
    // properties as a w:pPrChange (CR-002 section 38 item 2, the editor's measure-word15.docx), so
    // that the joined line reads in them. This method deletes the whole paragraph, text and mark - the
    // engine's only mark deletion, a Range being within one paragraph - and what Word writes for that
    // is not measured (test/README.md check 33). 0.3.1 carried the properties here too, which turned
    // an agent's new heading, put in after a paragraph it then deleted, into body text once accepted
    // (the editor's finding on 0.3.1); until the check, nothing is carried.
  }

  /**
   * The paragraph's content and its mark marked deleted, the paragraph element itself kept:
   * what the cells of a deleted table row need (a `w:tc` must keep a block-level child, and Word
   * writes `w:del` around every run and `w:pPr/w:rPr/w:del` on every mark inside a deleted row).
   * The same primitives `delete()` uses, without its "take back my own insertion" step.
   * A mark already deleted is left as it is, and the text is deleted still: a paragraph another author's
   * deletion ended keeps text a range's deletion across it must take (CR-014 phase 2).
   */
  markDeletedInPlace(tracker: ChangeTracker): void {
    this.assertWritable();
    this.deleteText(tracker, 0, this.text.length, WHOLE);
    if (!markDeleted(this.p)) tracker.markParagraphDeleted(this.p);
  }

  /**
   * Queues the relationships that content removed outright referenced (a picture's image, a link's
   * target), to be settled when the package is saved: each goes then if nothing references it, and
   * its part with it when nothing else targets that (CR-002 sections 47 and 48).
   */
  releaseRelationships(removed: unknown): void {
    const ids = new Set<string>();
    relationshipIdsIn(removed, ids);
    if (ids.size === 0) return;
    const pkg = this.parentBody.package_ as { pendingRelationshipRemovals?: RelationshipRemovals } | undefined;
    queueRelationshipRemovals(pkg?.pendingRelationshipRemovals, this.parentBody.part, ids);
  }

  /** The paragraph as XML (extension; Office JS getOoxml wraps it in a package). */
  getXml(): Promise<string> {
    return marshalString(this.element as Element);
  }

  // --- docx4j names ---

  /** docx4j: the P's content list. */
  getContent(): Element[] {
    return (this.p.content ??= []) as Element[];
  }

  // --- editing primitives, shared with Range ---

  /**
   * The paragraph's properties, created when absent. Every property setter goes through this
   * one accessor, so this is where a tracked write records `w:pPrChange` with the properties as
   * they stand before it.
   *
   * @internal public so that the list views (CR-002 phase H) write `w:numPr` the same way.
   */
  pPr(): wml.PPr {
    this.assertWritable();
    const pPr = (this.p.pPr ??= { TYPE_NAME: 'org_docx4j_wml.PPr' });
    linkParents(pPr, this.p);
    this.changeTracker?.recordPPrChange(pPr);
    return pPr;
  }

  /**
   * After a tracked write of the paragraph's properties: when the paragraph is as it was, its
   * recorded properties go back as they were and the `w:pPrChange` goes, since Word keeps no
   * formatting change that changes nothing, whoever changed it back (`test/README.md` check 26:
   * centred then left again, by the same author and by another; Increase Indent then Decrease
   * Indent; each left no `w:pPr`; CR-002 section 29). "As it was" is Office JS's paragraph
   * formatting (alignment, indents, spacing, outline level) read resolved from both, the rest of the
   * properties - the style, the list - written the same: Word's Ctrl+L and Decrease Indent removed
   * the direct values where this package's setters write `w:jc="left"` and `w:ind w:left="0"`.
   * Without the package's `PropertyResolver` only properties written the same count.
   *
   * @internal public so that the list views (CR-002 phase H) settle their `w:numPr` writes.
   */
  settleFormatting(): void {
    const pPr = this.p.pPr;
    const recorded = pPr?.pPrChange?.pPr;
    if (!pPr || !pPr.pPrChange || !this.changeTracker) return;
    const current: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(pPr)) if (key !== 'pPrChange' && key !== 'rPr' && key !== 'sectPr') current[key] = value;
    const before = { ...(recorded ?? {}) } as Record<string, unknown>;
    let same = canonical(current) === canonical(before);
    if (!same) {
      const rest = (props: Record<string, unknown>): string => canonical(Object.fromEntries(Object.entries(props).filter(([key]) => !FORMATTING_MEMBERS.has(key))));
      const resolver = (this.parentBody.package_ as { propertyResolverOrUndefined?: PropertyResolver } | undefined)?.propertyResolverOrUndefined;
      const cell = resolver?.cellContextOf(this.p);
      same = resolver !== undefined && rest(current) === rest(before)
        && canonical(formattingOf(resolver.getEffectivePPr(current as wml.PPr, cell), true)) === canonical(formattingOf(resolver.getEffectivePPr(before as wml.PPr, cell), true));
    }
    if (same) restorePPr(this.p, recorded);
  }

  /**
   * The paragraph's `w:numPr`, created when absent, through {@link pPr} so that a tracked
   * write records `w:pPrChange` first.
   *
   * @internal what the list views write through (CR-002 phase H).
   */
  numPr(): wml.PPrBase.NumPr {
    const pPr = this.pPr();
    const numPr = (pPr.numPr ??= { TYPE_NAME: 'org_docx4j_wml.PPrBase.NumPr' });
    linkParents(numPr, pPr);
    return numPr;
  }
  private ind(): wml.PPrBase.Ind {
    return (this.pPr().ind ??= {});
  }
  private spacing(): wml.PPrBase.Spacing {
    return (this.pPr().spacing ??= {});
  }

  /** The package's change tracker while `changeTrackingMode` is on; undefined when it is off. */
  get changeTracker(): ChangeTracker | undefined {
    return this.parentBody.changeTracker;
  }

  /** What `Font` needs to write `w:rPrChange`: the tracker, and the runs that are our own insertion (`Range` shares it). */
  fontTracking(): FontTracking | undefined {
    const tracker = this.changeTracker;
    if (!tracker) return undefined;
    const ownInsertions = new Set<object>();
    for (const seg of this.segments()) if (tracker.ownInsertion(seg.revision)) ownInsertions.add(seg.run);
    return { tracker, ownInsertions };
  }

  /**
   * The text segments (w:t, w:tab, ...) with their offsets in `text`; `{ objects: true }` adds the
   * pictures and the hidden runs' items.
   */
  segments(options?: SegmentOptions): TextSegment[] {
    return segmentsOf(this.p, { hidden: this.hiddenRuns(), ...options });
  }

  /**
   * True for a paragraph in a text box (`w:txbxContent`), at a `textbox:` address (CR-012 D4). Such a view
   * reads as any paragraph does and is read-only (D5): every write throws `NotSupportedError`.
   */
  get inTextBox(): boolean {
    let o = (this.p as { PARENT?: object }).PARENT;
    for (let guard = 0; o && guard < 256; guard++, o = (o as { PARENT?: object }).PARENT) {
      if ((o as { TYPE_NAME?: string }).TYPE_NAME === 'org_docx4j_wml.CTTxbxContent') return true;
    }
    return false;
  }

  /**
   * Refuses a write to a paragraph in a text box (CR-012 D5): a box is in each branch of its
   * `mc:AlternateContent`, `wps:txbx` in the Choice and `v:textbox` in the Fallback, and nothing maps the
   * one's paragraphs to the other's, so an edit could not reach every branch as CR-010's D2 has mutators do.
   * And a write to a paragraph in a branch of a kept `mc:AlternateContent` at block, row or cell level
   * (CR-013 D4), for the same reason. Internal: the content API's mutators call it first.
   */
  assertWritable(): void {
    if (this.inTextBox) {
      throw new NotSupportedError('Paragraph', 'edit', 'this paragraph is in a text box, whose paragraphs are read-only (CR-012 D5): '
        + "writes into a text box would have to reach its other branch's box too, which nothing maps yet");
    }
    assertOutsideBranch(this.p, 'Paragraph');
  }

  /**
   * True for a paragraph in the selected branch of a kept `mc:AlternateContent` at block, row or cell
   * level, which readers see as Word does (CR-013 phase 2). Such a view is read-only (D4).
   */
  get inBranch(): boolean {
    return inBranch(this.p);
  }

  /** Which of this paragraph's runs are hidden: a run's own `w:vanish`, else its styles' (CR-012 phase 4). */
  hiddenRuns(): HiddenRuns {
    return this.parentBody.hiddenRuns(this.p);
  }

  /**
   * Replaces the text in [start, end) with `text` and returns the range of the inserted text.
   * Runs are edited in place: a w:t loses or gains characters, a tab or break inside the range
   * is removed, a run left empty is removed. Text is inserted into the w:t at `start`,
   * else into a new run with the formatting of the nearest run.
   *
   * The pictures in the span go with its text: those strictly inside it, and those at an edge `edges`
   * says the span holds (`WHOLE` for a paragraph's whole content), so a span of no text can still
   * delete a picture (CR-002 section 48).
   */
  splice(start: number, end: number, text: string, prefer: 'before' | 'after' = 'before', edges?: SpanEdges, lineBreaks?: LineBreaks): Range {
    this.assertWritable();
    // text into a block control showing its placeholder takes the placeholder's place (CR-002 section 63)
    const placeholder = text === '' ? undefined : this.placeholderControl;
    if (placeholder) return placeholder.writePlaceholder(text, lineBreaks ? { lineBreaks } : {});
    // every text write into a bound control writes its node (CR-002 section 63, D7), as ContentControl.insertText did
    const bound = this.boundControlsAt(start, end);
    const range = this.spliceWrite(start, end, text, prefer, edges, lineBreaks);
    for (const control of bound) control.syncBoundNode();
    return range;
  }

  /**
   * The controls bound by `w:dataBinding` that [start, end) is in: those holding the paragraph, and the run-level ones
   * holding text there.
   */
  private boundControlsAt(start: number, end: number, ancestorsOnly = false): ContentControl[] {
    const out: ContentControl[] = [];
    const seen = new Set<object>();
    const add = (o: object): void => {
      const tn = (o as { TYPE_NAME?: string }).TYPE_NAME;
      if (tn === undefined || !SDT_VALUE_TYPES.has(tn) || seen.has(o)) return;
      seen.add(o);
      if (sdtProperty((o as { sdtPr?: wml.SdtPr }).sdtPr, 'dataBinding') === undefined) return;
      try {
        out.push(controlViewOf(o, this.parentBody));
      } catch {
        // not where its PARENT says: nothing to write
      }
    };
    for (let o = (this.p as { PARENT?: object }).PARENT; o; o = (o as { PARENT?: object }).PARENT) add(o);
    if (ancestorsOnly || !holdsBoundRunControl(this.p)) return out;
    for (const seg of this.segments({ objects: true })) {
      if (seg.end < start || seg.start > end) continue;
      for (let o = (seg.run as { PARENT?: object }).PARENT; o && o !== this.p; o = (o as { PARENT?: object }).PARENT) add(o);
    }
    return out;
  }

  /**
   * @internal Text written at the paragraph's End by Office JS's routes to it: `insertText('End')`, its End range at any
   * location, its Content range's End and After, and so a body's or a cell's `insertText('End')` (CR-002 sections 63.24
   * to 63.27; checks 69 and 70). As Word writes it (D12): after everything the paragraph holds, into none of it - a
   * run-level control, a link, a revision, a field, a bookmark's or a comment's markers, a text box - in a run of its
   * own, or the last run extended. The text takes the last item's run properties where that item is a run holding text,
   * so after a hidden run it is hidden and its range reads nothing (D13); otherwise it has none, and with `plain` (the End
   * range's After, D14) none either. Word's proofing marks, `w:proofErr`, are no item (a choice: Word redoes them).
   * Tracked, the run is a `w:ins`, or joins this author's own `w:ins` where one ends the paragraph. A line break makes
   * paragraphs, as `splice` does (CR-002 section 60 E1); text so broken after a hidden run is written visible (a choice,
   * unmeasured).
   */
  writeAtEnd(text: string, plain = false): Range {
    this.assertWritable();
    const placeholder = this.placeholderControl;
    if (placeholder) return placeholder.writePlaceholder(text);
    const length = this.text.length;
    text = dropExcluded(text);
    if (text === '') return new Range(this, length, length);
    let mode: LineBreaks = 'paragraphs';
    const lines = this.blockControl ? plainTextLinesOf(this.blockControl) : undefined;
    if (lines) mode = lines === 'multi' ? 'multi-line' : 'single-line';   // section 61 item 3, F3
    text = forWriter(text, mode);
    const broken = mode === 'paragraphs' && /[\r\n]/.test(text) ? linesOf(text) : undefined;
    const joined = broken ? broken.join('') : text;
    if (joined === '' && !broken) return new Range(this, length, length);
    const bound = this.boundControlsAt(length, length, true);
    const content = (this.p.content ??= []) as Element[];
    let at = content.length;
    while (at > 0 && typeNameOf(content[at - 1]!) === 'org_docx4j_wml.ProofErr') at--;
    const last = at > 0 ? content[at - 1] : undefined;
    const textRun = (item: Element | undefined): wml.R | undefined => {
      if (!item || typeNameOf(item) !== 'org_docx4j_wml.R') return undefined;
      const run = item.value as wml.R;
      return ((run.content ?? []) as Element[]).some((c) => c.name?.localPart === 't') ? run : undefined;
    };
    const hidden = this.hiddenRuns();
    // the run whose properties the text takes, and which it may extend: a run holding text, unless `plain`
    let source = plain ? undefined : textRun(last);
    if (source && broken && hidden(source)) source = undefined;
    const items = officeItems(joined);
    const tracker = this.changeTracker;
    let written: wml.R | undefined;
    const own = tracker && last && typeNameOf(last) === 'org_docx4j_wml.RunIns' && tracker.ownInsertion({ kind: 'ins', value: last.value } as RevisionHolder);
    if (items.length === 0) {
      // nothing but line breaks: only the paragraphs they make
    } else if (own) {
      // this author's own insertion ends the paragraph: the text joins it, after its last run, in that run's look
      const value = last!.value as { customXmlOrSmartTagOrSdt?: Element[] };
      const holder = (runItemsOf(value) as Element[] | undefined) ?? (value.customXmlOrSmartTagOrSdt = []);
      const inner = plain ? undefined : textRun(holder[holder.length - 1]);
      written = this.appendTo(holder, inner, items, value);
    } else if (tracker) {
      const ins = tracker.ins([runOf(items, copyRPr(source?.rPr))]);
      content.splice(at, 0, ins as Element);
      linkParents(ins, this.p);
      written = runsOf(ins.value as object)[0]!.value;
    } else {
      written = this.appendTo(content, source, items, this.p, at);
    }
    for (const control of bound) control.syncBoundNode();
    if (broken) {
      const offsets: number[] = [];
      let end = length;
      for (const line of broken.slice(0, -1)) offsets.push(end += line.length);
      const { first, last: tail } = breakParagraphAt(this, offsets);
      return new Range(first, length, broken[broken.length - 1]!.length, false, false, tail);
    }
    return written && hidden(written) ? new Range(this, length, length) : new Range(this, length, length + joined.length);
  }

  /**
   * `items` put at `at` in `list` (its end by default): onto the end of `run`, which is in `list` right before `at`,
   * merging into its last `w:t`, or as a new run in its properties where `run` is undefined. Returns the run written.
   */
  private appendTo(list: Element[], run: wml.R | undefined, items: Element[], parent: object, at = list.length): wml.R {
    if (run) {
      const own = (run.content ??= []) as Element[];
      const tail = own[own.length - 1];
      const head = items[0];
      if (tail?.name?.localPart === 't' && head?.name?.localPart === 't') {
        setText(tail.value as wml.Text, ((tail.value as wml.Text).value ?? '') + ((head.value as wml.Text).value ?? ''));
        items = items.slice(1);
      }
      own.push(...items);
      linkParents(items, run);
      return run;
    }
    const added = runOf(items);
    list.splice(at, 0, added as Element);
    linkParents(added, parent);
    return added.value;
  }

  /** `splice`'s write. */
  private spliceWrite(start: number, end: number, text: string, prefer: 'before' | 'after', edges: SpanEdges | undefined, lineBreaks: LineBreaks | undefined): Range {
    // Office JS's characters written back (CR-002 section 60): the characters XML 1.0 excludes go (E3), a line feed or
    // carriage return is a paragraph break (E1) unless the writer or the control takes it otherwise, and every
    // character that is an element is written as it (D1)
    text = dropExcluded(text);
    if (text === '') return this.spliceText(start, end, text, prefer, edges);
    const from = Math.max(0, Math.min(start, this.text.length));
    const landing = this.landingSegment(from, prefer, end > from);
    const runControl = this.runControlOf(landing);
    let mode: LineBreaks = lineBreaks ?? 'paragraphs';
    if (runControl) {
      text = textForRunControl(runControl, text);   // F4 and F3
    } else if (lineBreaks === undefined) {
      const lines = this.blockControl ? plainTextLinesOf(this.blockControl) : undefined;
      if (lines) mode = lines === 'multi' ? 'multi-line' : 'single-line';   // section 61 item 3, F3
    }
    text = forWriter(text, mode);
    if (mode === 'paragraphs' && /[\r\n]/.test(text)) return this.spliceLines(start, end, linesOf(text), prefer, edges, landing);
    const range = this.spliceText(start, end, text, prefer, edges);
    this.writeElements(range.start, range.start + text.length);
    return range;
  }

  /**
   * Text of several lines put in place of [start, end) as Word's insertText puts it (CR-002 section 60 E1): the lines'
   * text, then a paragraph break after each line but the last, the paragraphs keeping this one's properties; tracked,
   * each new mark inserted. Refused before anything changes where a paragraph can't be split there.
   */
  private spliceLines(start: number, end: number, lines: string[], prefer: 'before' | 'after', edges: SpanEdges | undefined, landing: TextSegment | undefined): Range {
    this.assertBreakable(landing);
    const joined = lines.join('');
    const range = this.spliceText(start, end, joined, prefer, edges);
    const from = range.start;
    this.writeElements(from, from + joined.length);
    const offsets: number[] = [];
    let at = from;
    for (const line of lines.slice(0, -1)) offsets.push(at += line.length);
    const { first, last } = breakParagraphAt(this, offsets);
    return new Range(first, from, lines[lines.length - 1]!.length, false, false, last);
  }

  /** @internal Refuses, before anything changes, a paragraph break at `at` where none can be made (E1; `Range`'s Replace across paragraphs). */
  assertBreakableAt(at: number, prefer: 'before' | 'after'): void {
    this.assertBreakable(this.landingSegment(at, prefer));
  }

  /**
   * The segment text put at `at` goes into, as `spliceText` chooses it: a replacement's, the one holding the first
   * replaced character; else the one containing `at`, else the nearer neighbour.
   */
  private landingSegment(at: number, prefer: 'before' | 'after', replacing = false): TextSegment | undefined {
    const segs = this.segments();
    if (replacing) {
      const first = segs.find((s) => s.start <= at && at < s.end);
      if (first) return first;
    }
    const containing = segs.find((s) => s.editable && s.start < at && at < s.end);
    const ending = segs.find((s) => s.editable && s.end === at);
    const starting = segs.find((s) => s.editable && s.start === at);
    return containing ?? (prefer === 'after' ? starting ?? ending : ending ?? starting);
  }

  /** The run-level content control a segment's run is in, the innermost, as a view; undefined for none. */
  private runControlOf(seg: TextSegment | undefined): ContentControl | undefined {
    for (let o = (seg?.run as { PARENT?: object } | undefined)?.PARENT; o && o !== this.p; o = (o as { PARENT?: object }).PARENT) {
      if ((o as { TYPE_NAME?: string }).TYPE_NAME === 'org_docx4j_wml.SdtRun') return controlViewOf(o, this.parentBody);
    }
    return undefined;
  }

  /** The block-level content control holding this paragraph directly, as a view; undefined for none. */
  get blockControl(): ContentControl | undefined {
    const content = (this.p as { PARENT?: object }).PARENT as { TYPE_NAME?: string; PARENT?: object } | undefined;
    const sdt = content?.TYPE_NAME === 'org_docx4j_wml.SdtContentBlock' ? content.PARENT : undefined;
    return (sdt as { TYPE_NAME?: string } | undefined)?.TYPE_NAME === 'org_docx4j_wml.SdtBlock' ? controlViewOf(sdt!, this.parentBody) : undefined;
  }

  /**
   * Refuses, before anything changes, a paragraph break where the text lands inside something a paragraph can't be
   * split in: a content control, an equation, a kept mc:AlternateContent (CR-002 section 60 E1, as `splitParagraph`
   * refuses one). A field's result is split: a complex field's runs across the paragraphs, a `w:fldSimple` made a
   * complex field first (section 62.2 D1).
   */
  private assertBreakable(seg: TextSegment | undefined): void {
    if (!seg) return;
    const refuse = (what: string): never => {
      throw new NotSupportedError('Paragraph', 'insertText', `a line feed or carriage return is a paragraph break (CR-002 section 60, E1), `
        + `and this text goes inside ${what}, which can't be cut in two`);
    };
    if (seg.alternate || seg.math) refuse('an equation or an mc:AlternateContent');
    for (let o = (seg.run as { PARENT?: object }).PARENT; o && o !== this.p; o = (o as { PARENT?: object }).PARENT) {
      const tn = (o as { TYPE_NAME?: string }).TYPE_NAME ?? '';
      if (!BREAKABLE.has(tn)) refuse(`a ${tn.replace('org_docx4j_wml.', '')}`);
    }
  }

  /**
   * The characters of [from, to) that are elements written as the elements they read as (CR-002 section 60 D1), each
   * in its run where it stood: a w:t holding one is cut around it. Characters outside the span are left as they are.
   */
  private writeElements(from: number, to: number): void {
    if (to <= from) return;
    for (const seg of this.segments().reverse()) {
      if (!seg.editable || seg.end <= from || seg.start >= to) continue;
      const value = seg.text;
      const a = Math.max(0, from - seg.start);
      const b = Math.min(value.length, to - seg.start);
      const inside = value.substring(a, b);
      if (!holdsElementCharacters(inside)) continue;
      const items: Element[] = [
        ...(a > 0 ? [textItem(value.substring(0, a)) as Element] : []),
        ...officeItems(inside),
        ...(b < value.length ? [textItem(value.substring(b)) as Element] : []),
      ];
      const parent = (seg.item.value as { PARENT?: object }).PARENT ?? seg.run;
      seg.owner.splice(seg.index, 1, ...items);
      linkParents(items, parent);
    }
  }

  /**
   * The first run's properties, for text with no run beside it to take its look from; none where that run is hidden,
   * so that text put into a paragraph whose every run is hidden can be seen (CR-002 section 60: the text was hidden
   * with it, so read as nothing, and its characters escaped the map).
   */
  private visibleFirstRunRPr(): wml.RPr | undefined {
    const first = runsOf(this.p)[0]?.value;
    return first && !this.hiddenRuns()(first) ? first.rPr : undefined;
  }

  /** `splice` before section 60: the text put in as it is, as one stretch of a w:t's text. */
  private spliceText(start: number, end: number, text: string, prefer: 'before' | 'after' = 'before', edges?: SpanEdges): Range {
    const length = this.text.length;
    start = Math.max(0, Math.min(start, length));
    end = Math.max(start, Math.min(end, length));
    const tracker = this.changeTracker;
    this.assertPlaceable(start, end, tracker === undefined ? false : text.length > 0 ? 'replace' : 'delete');
    // a run-level control showing its placeholder that the span holds whole goes with it, as Word takes one; tracked,
    // `deleteText` marks it deleted (CR-002 section 63, D3; check 63 cases 19, 20 and 26)
    if (!tracker && (end > start || edges)) for (const control of this.placeholderControlsIn(start, end, edges)) control.takeWhole();
    if (tracker) {
      // a replacement is the w:del first and the w:ins after it, as Word writes one
      const deletions = end > start || edges ? this.deleteText(tracker, start, end, edges) : [];
      if (text.length > 0) this.insertTracked(tracker, start, text, deletions[deletions.length - 1], prefer);
      return new Range(this, start, start + text.length);
    }
    let segs = this.segments();
    let inserted = false;
    const objects = this.segments({ objects: true }).filter((s) => s.object && objectInSpan(s.start, start, end, edges));
    // a comment's reference in the span goes with the text, and its comment with it (CR-014 section 12.6)
    const references = end > start || edges ? this.commentReferencesIn(start, end, edges) : [];
    // A replacement that begins inside a revision is a deletion and then the new text typed there, so
    // the new text is plain and splits the revision, as typed text is (test/README.md check 21), rather
    // than joining the revision it replaced part of.
    if (end > start && text.length > 0 && segs.find((s) => s.start <= start && start < s.end)?.revision) {
      this.spliceText(start, end, '', prefer, edges);
      return this.spliceText(start, start, text);
    }
    if (end > start || objects.length > 0 || references.length > 0) {
      // Replace: the new text takes the place (and formatting) of the first replaced character, as Word does.
      const first = segs.find((s) => s.start <= start && start < s.end);
      const toRemove: TextSegment[] = [];
      for (const seg of segs) {
        if (seg.end <= start || seg.start >= end) continue;
        const from = Math.max(0, start - seg.start);
        const to = Math.min(seg.text.length, end - seg.start);
        if (seg.editable) {
          const t = seg.item.value as wml.Text;
          const middle = seg === first && !inserted ? text : '';
          if (seg === first) inserted = true;
          const value = seg.text.substring(0, from) + middle + seg.text.substring(to);
          if (value === '') toRemove.push(seg); else setText(t, value);
        } else toRemove.push(seg);
      }
      for (const seg of [...toRemove.reverse(), ...objects]) {
        const i = seg.owner.indexOf(seg.item);
        if (i >= 0) seg.owner.splice(i, 1);
      }
      removeEmptiedAlternates([...toRemove, ...objects]);
      for (const { run, owner, id } of references) {
        const i = owner.indexOf(run);
        if (i >= 0) owner.splice(i, 1);
        this.queueCommentRemoval(id);
      }
      this.releaseRelationships(objects.map((seg) => seg.item));
      if (inserted && text.length > 0) {
        this.removeEmptyRuns();
        return new Range(this, start, start + text.length);
      }
      this.removeEmptyRuns();
      segs = this.segments();
    }
    if (text.length > 0) {
      // Text typed with tracking off is plain wherever it lands, as Word writes it (test/README.md
      // check 21): inside a revision it splits the revision, and next to one it goes beside it, not
      // in it. Otherwise it extends the w:t at `start` (the one containing it, else the one ending
      // there, else the one starting there).
      const containing = segs.find((s) => s.editable && s.start < start && start < s.end);
      if (containing?.revision && this.insertBetweenHalves(start, runOf([textItem(text)], copyRPr(containing.run.rPr)), () => this.nextAnnotationId())) {
        return new Range(this, start, start + text.length);
      }
      const ending = segs.find((s) => s.editable && s.end === start);
      const starting = segs.find((s) => s.editable && s.start === start);
      // `prefer` says which run takes text put exactly between two: the one ending there (typed text,
      // `Paragraph.insertText`) or the one starting there (`Range.insertText('Before')`, which Office
      // JS puts inside the link the range is on: check 32)
      const [near, far] = prefer === 'after' ? [starting, ending] : [ending, starting];
      const target = (containing && !containing.revision ? containing : undefined)
        ?? (near && !near.revision ? near : undefined) ?? (far && !far.revision ? far : undefined)
        ?? containing;                                   // a revision that could not be split: extended, as before
      const edge = target ? undefined : near ?? far;
      if (edge?.revision) {
        const run = runOf([textItem(text)], copyRPr(edge.run.rPr));
        this.insertBesideRevision(edge.revision, run, edge === ending ? 'after' : 'before');
        return new Range(this, start, start + text.length);
      }
      if (target) {
        const t = target.item.value as wml.Text;
        const at = start - target.start;
        setText(t, target.text.substring(0, at) + text + target.text.substring(at));
      } else {
        // no w:t to extend: a new run next to the segment at `start`, with that run's formatting
        const before = [...segs].reverse().find((s) => s.end <= start);
        const after = segs.find((s) => s.start >= start);
        const neighbour = prefer === 'after' ? after ?? before : before ?? after;
        const firstRun = this.visibleFirstRunRPr();
        const rPr = neighbour?.run.rPr ? deepCopy(neighbour.run.rPr) : firstRun ? deepCopy(firstRun) : undefined;
        const run = runOf([textItem(text)], rPr);
        if (neighbour === before && before) before.runOwner.splice(before.runIndex + 1, 0, run);
        else if (after) after.runOwner.splice(after.runIndex, 0, run);
        else {
          const content = (this.p.content ??= []);
          if (start === 0) content.unshift(run as never); else content.push(run as never);
        }
        linkParents(run, this.p);
      }
    }
    return new Range(this, start, start + text.length);
  }

  /**
   * Inserts run-level items (runs, hyperlinks, content controls) at a text offset, splitting the
   * run there when the offset falls inside one. PARENT is linked on what is inserted.
   */
  insertItemsAt(offset: number, items: Element[]): void {
    this.assertWritable();
    if (items.length === 0) return;
    this.assertPlaceable(offset, offset, false);
    const tracker = this.changeTracker;
    if (tracker) {
      // With tracking on, what goes in is an insertion: its runs in w:ins, in a link, a control or a
      // field too, placed as typed text is - beside another's revision, or between its halves, never
      // in it. Before, `insertOoxml` into a paragraph went in untracked (CR-002 section 37).
      wrapInsertedRuns(tracker, items, this.p);
      this.splitAt(offset);
      const segs = this.segments();
      const before = [...segs].reverse().find((s) => s.end <= offset);
      const after = segs.find((s) => s.start >= offset);
      tracker.assertEditable((before ?? after)?.revision);
      this.placeInsertion(offset, items, before, after, () => tracker.nextId());
      return;
    }
    this.splitAt(offset);
    const segs = this.segments();
    const after = segs.find((s) => s.start >= offset);
    const before = [...segs].reverse().find((s) => s.end <= offset);
    const neighbour = after ?? before;
    const list = neighbour ? neighbour.runOwner : ((this.p.content ??= []) as Element[]);
    const at = neighbour ? (after ? neighbour.runIndex : neighbour.runIndex + 1) : (offset === 0 ? 0 : list.length);
    list.splice(at, 0, ...items);
    const owner = (neighbour?.run as { PARENT?: object } | undefined)?.PARENT ?? this.p;
    for (const item of items) linkParents(item, owner);
  }

  /** Splits the run at a text offset so that [offset, ...) begins a run; returns nothing when the offset is already a boundary. */
  splitAt(offset: number): void {
    this.assertWritable();
    this.assertPlaceable(offset, offset, false);
    const seg = this.segments().find((s) => s.editable && s.start < offset && offset < s.end);
    if (!seg) return;
    const at = offset - seg.start;
    const t = seg.item.value as wml.Text;
    const head = seg.text.substring(0, at);
    const tail = textItem(seg.text.substring(at));
    setText(t, head);
    seg.owner.splice(seg.index + 1, 0, tail);
    linkParents(tail, (t as { PARENT?: object }).PARENT ?? seg.run);
    this.splitRunAtItem(seg, seg.owner, seg.index + 1);
  }

  /**
   * Splits `seg`'s run so that the items of `owner` from `index` on begin a run of their own, behind it.
   * `owner` is the run's content, or the selected branch of a run-level `mc:AlternateContent` in it
   * (CR-012 section 12.2): a run holds the element whole, so the branch splits with it, as the editor's
   * ED-009 section 3.4 has it. The first piece keeps the element, its other branches as they were; the
   * second gets an element of its own, the same branches in the same order with the same `Requires`, the
   * selected one holding the rest of its items and the others empty, so that a reader taking any branch
   * takes the same one in both. The run's items after the element go with the second piece.
   */
  private splitRunAtItem(seg: TextSegment, owner: Element[], index: number): void {
    const alternate = seg.alternate?.level === 'run' && owner === seg.alternate.branch ? seg.alternate : undefined;
    let moved: Element[];
    if (alternate) {
      const piece = alternatePiece(alternate, owner.splice(index));
      const at = alternate.owner.indexOf(alternate.element);
      moved = [piece, ...alternate.owner.splice(at + 1)];
    } else {
      moved = owner.splice(index);
    }
    if (moved.length === 0) return;
    const second = runOf(moved, seg.run.rPr ? deepCopy(seg.run.rPr) : undefined);
    seg.runOwner.splice(seg.runIndex + 1, 0, second);
    linkParents(second, (seg.run as { PARENT?: object }).PARENT ?? this.p);
  }

  // --- lists (CR-002 phase H) ---

  /**
   * Office JS `paragraph.isListItem`: whether Word paints a label in front of this paragraph.
   * The effective `w:numPr` decides it (docx4j `Emulator.numRefFor`), so a "List Number" style
   * counts; a `w:numId` of 0, a `w:numId` naming no `w:num`, and a level linked to another
   * paragraph style do not.
   */
  get isListItem(): boolean {
    return isListItem(this);
  }

  /** Office JS `paragraph.list`: the list this paragraph is in; throws when it is not in one. */
  get list(): List {
    return listOf(this);
  }

  /** Office JS `paragraph.listOrNullObject`: the list, or a null object (`isNullObject` true). */
  get listOrNullObject(): List {
    return listOrNullObjectOf(this);
  }

  /** Office JS `paragraph.listItem`: this paragraph's place in its list; throws when it has none. */
  get listItem(): ListItem {
    return listItemOf(this);
  }

  /** Office JS `paragraph.listItemOrNullObject`. */
  get listItemOrNullObject(): ListItem {
    return listItemOrNullObjectOf(this);
  }

  /**
   * Office JS `paragraph.startNewList()`: a new list definition (docx4j's default decimal set,
   * or its bullet set with `{ bullet: true }`) with this paragraph as its first item, adding
   * the numbering part and its relationship when the document has none.
   *
   * Asynchronous where Office JS's is synchronous, since the default definitions are XML to be
   * unmarshalled (CR-002 section 17).
   */
  startNewList(options?: StartListOptions): Promise<List> {
    this.assertWritable();
    return startNewList(this, options);
  }

  /** Office JS `paragraph.attachToList(listId, level)`: this paragraph joins that list. */
  attachToList(listId: number | string, level = 0): List {
    this.assertWritable();
    return attachToList(this, listId, level);
  }

  /** Office JS `paragraph.detachFromList()`: this paragraph leaves its list. */
  detachFromList(): void {
    this.assertWritable();
    detachFromList(this);
  }

  /**
   * The same list, starting again at this paragraph: a new `w:num` over the same
   * `w:abstractNum` with a `w:startOverride` (extension; `List.restart()` is the verb).
   */
  restartList(): List {
    this.assertWritable();
    const list = this.list;
    const level = this.listItem.level;
    const restarted = list.restart();
    this.attachToList(restarted.id, level);
    return restarted;
  }

  // --- comments (CR-002 phase G) ---

  /** The comments anchored in this paragraph, replies nested under their parent. */
  async getComments(): Promise<Comment[]> {
    return commentApi().commentsOf(this);
  }

  /** Comments the whole paragraph (sugar over `getRange().insertComment`). */
  async insertComment(content: CommentContent, options?: CommentOptions): Promise<Comment> {
    this.assertWritable();
    return commentApi().insertComment(this.getRange(), content, options);
  }

  // --- content controls (CR-002 phase E) ---

  /**
   * Wraps this paragraph in a new block-level content control (Office JS
   * `Paragraph.insertContentControl`), typed for the kind when one is given. The `w:id` is free in
   * the body. Returns the control.
   */
  insertContentControl(kind?: ContentControlType): ContentControl {
    this.assertWritable();
    const at = this.index;
    if (at < 0) throw new Docx4JException('This paragraph is not in its container any more');
    // the paragraph's parent read before the builder makes the control its parent: read after, it was the
    // control's own content, and the control's PARENT pointed into itself (found in CR-012 phase 4)
    const parent = (this.p as { PARENT?: object }).PARENT ?? this.parentBody.container;
    // a bound block control holds one paragraph (CR-002 section 61): checked before sdtOf, which takes the paragraph
    assertControlHolds(this.parentBody, parent, this.container.filter((e) => e !== this.element), [BLOCK_CONTROL], undefined, ['Paragraph', 'insertContentControl']);
    const empty = this.segments({ objects: true }).length === 0;
    const sdt = sdtOf([this.element as Element], { kind: sdtKindFor(kind), id: nextSdtId(controlIdScope(this.parentBody)), form: 'block' });
    this.container.splice(at, 1, sdt as Element);
    linkParents(sdt, parent);
    declareSdtExtensions(this.parentBody, sdt);
    // under tracking the control itself is the insertion, the paragraph what it was (check 30)
    const tracker = this.changeTracker;
    if (tracker) markHolderInserted(tracker, sdt as Element, this.container, () => {});
    const control = new ContentControl(sdt as Element<wml.SdtBlock>, this.container, this.parentBody);
    // around an empty paragraph, a control showing Word's default placeholder, as Word's (CR-002 section 63, D8; check 64)
    if (empty && showsPlaceholderWhenMade(kind)) control.showPlaceholder();
    return control;
  }

  private removeEmptyRuns(): void {
    const prune = (items: Element[]): void => {
      for (let i = items.length - 1; i >= 0; i--) {
        const el = items[i]!;
        const v = el.value as { TYPE_NAME?: string; content?: Element[]; sdtContent?: { content?: Element[] } };
        if (v.TYPE_NAME === 'org_docx4j_wml.R') {
          if (!v.content || v.content.length === 0) items.splice(i, 1);
        } else if (isAlternateContent(el)) {
          // among the runs: a selected branch an edit has emptied takes the element, every branch with
          // it, as a deletion of all of a run-level element's text does (CR-012 section 12.2 item 4)
          const branch = selectedBranchOf(el.value as mce.AlternateContent);
          if (!branch || branch.length === 0) continue;
          prune(branch);
          if (branch.length === 0) items.splice(i, 1);
        } else if (v.TYPE_NAME !== 'org_docx4j_wml.P') {
          const nested = runItemsOf(v);
          if (nested) {
            prune(nested);
            // a revision left with nothing in it goes too
            if (nested.length === 0 && revisionKindOf(el) !== undefined) items.splice(i, 1);
          }
        }
      }
    };
    prune(this.p.content ?? []);
  }

  // --- tracked editing (CR-002 phase F) ---------------------------------------------------

  /**
   * The text in [start, end) becomes a deletion: the runs holding it are isolated, their `w:t`
   * turned into `w:delText`, and each run of consecutive ones moved into one `w:del`. Text this
   * author had inserted is simply removed, as Word does, rather than nested in a `w:del`.
   * The pictures in the span (`splice`'s `edges`) are deleted with the text, each in a run of its own
   * where its run held more, as `InlinePicture.delete` gives it one (CR-002 section 48). An equation is deleted where
   * it stands, in Word's form, a `w:del` in each math run and each structure's `m:ctrlPr` (`markEquationDeleted`;
   * section 62.2 D2), the runs either side in `w:del`s of their own; one holding a revision of its own goes into the
   * `w:del` whole, as every equation did before.
   * Returns the `w:del` elements it made, in document order.
   */
  private deleteText(tracker: ChangeTracker, start: number, end: number, edges?: SpanEdges): Anchor[] {
    // a run-level control showing its placeholder that the span holds whole is marked deleted whole, in Word's markers,
    // its placeholder gone unmarked (CR-002 section 63, D3; check 63's run 2, cases 45, 46 and 52)
    for (const control of this.placeholderControlsIn(start, end, edges)) control.takeWhole();
    const targets = this.isolate(start, end, edges);
    for (const seg of targets) tracker.assertEditable(seg.revision);
    type Target = { run: Element; owner: Element[]; own: boolean; revision?: RevisionHolderOf };
    const list: Target[] = [];
    const equations: object[] = [];
    for (const seg of targets) {
      if (seg.math && !tracker.ownInsertion(seg.revision) && equationRevisionIn(seg.item.value as object) === undefined) {
        equations.push(seg.item.value as object);
        continue;
      }
      const run = seg.runOwner[seg.runIndex]!;
      if (list.length > 0 && list[list.length - 1]!.run === run) continue;
      const target: Target = { run, owner: seg.runOwner, own: tracker.ownInsertion(seg.revision) };
      if (seg.revision) target.revision = seg.revision;
      list.push(target);
    }
    // a comment's reference in the span goes into the w:del with the text, where it stands among the runs, as Word
    // puts it (check 54 case 28; CR-014 section 12.6); accepting the deletion then takes the comment
    const references = this.commentReferencesIn(start, end, edges);
    if (references.length > 0) {
      const order = runOrderOf(this.p);
      const revisions = new Map(targets.filter((seg) => seg.revision).map((seg) => [seg.runOwner, seg.revision!]));
      for (const { run, owner } of references) {
        const revision = revisions.get(owner) ?? revisionHolding(run);
        const target: Target = { run, owner, own: tracker.ownInsertion(revision) };
        if (revision) target.revision = revision;
        const key = order.get(run) ?? Number.NaN;
        const at = list.findIndex((t) => (order.get(t.run) ?? Number.NaN) > key);
        list.splice(at < 0 ? list.length : at, 0, target);
      }
    }
    // groups of consecutive runs in the same array, so that one w:del holds them all
    const groups: Target[][] = [];
    for (const t of list) {
      const last = groups[groups.length - 1];
      const previous = last?.[last.length - 1];
      if (last && previous && previous.own === t.own && previous.owner === t.owner
        && t.owner.indexOf(t.run) === t.owner.indexOf(previous.run) + 1) last.push(t);
      else groups.push([t]);
    }
    for (const equation of equations) markEquationDeleted(tracker, equation);
    const made: Anchor[] = [];
    for (let g = groups.length - 1; g >= 0; g--) {
      const group = groups[g]!;
      const owner = group[0]!.owner;
      const at = owner.indexOf(group[0]!.run);
      const runs = owner.splice(at, group.length);
      if (group[0]!.own) {
        // taking back our own insertion: the runs go, and an emptied w:ins with them; a picture's
        // image waits for the save, as `InlinePicture.delete` leaves it (section 47); a comment whose
        // reference goes with them goes too
        this.releaseRelationships(runs);
        for (const id of commentReferenceIds(runs)) this.queueCommentRemoval(id);
        const revision = group[0]!.revision;
        if (revision && revision.items.length === 0) {
          const i = revision.owner.indexOf(revision.element);
          if (i >= 0) revision.owner.splice(i, 1);
        }
        continue;
      }
      for (const run of runs) toDeletedText(run.value as wml.R);
      const del = tracker.del(runs);
      owner.splice(at, 0, del);
      linkParents(del, (group[0]!.revision?.value ?? (group[0]!.run.value as { PARENT?: object }).PARENT ?? this.p) as object);
      const revision = group[0]!.revision;
      const anchor: Anchor = { owner, element: del, parent: (revision?.value ?? (del.value as { PARENT?: object }).PARENT ?? this.p) as object };
      if (revision && revision.items === owner) anchor.revision = revision;
      made.unshift(anchor);
    }
    return made;
  }

  /**
   * The run-level controls showing their placeholder that [start, end) holds whole: where the walk passes one strictly
   * inside the span, or at an edge `edges` says it holds. A placeholder is no text, so such a control has no width
   * (CR-002 section 63, D3).
   */
  private placeholderControlsIn(start: number, end: number, edges?: SpanEdges): ContentControl[] {
    const found: object[] = [];
    this.segments({ visit: (item, at) => {
      if (typeNameOf(item) !== 'org_docx4j_wml.SdtRun' || !objectInSpan(at, start, end, edges)) return;
      const sdtPr = (item.value as wml.SdtRun).sdtPr;
      const flag = sdtProperty(sdtPr, 'showingPlcHdr');
      if (flag !== undefined && (flag.value as wml.BooleanDefaultTrue | undefined)?.val !== false) found.push(item.value as object);
    } });
    return found.map((sdt) => controlViewOf(sdt, this.parentBody));
  }

  /**
   * The runs of comments' references in [start, end): each a run holding nothing but a reference, where the walk
   * passes it, strictly inside the span or at an edge `edges` says it holds. A reference has no text, so a
   * deletion's segments never held it, and a comment whose reference a deletion took stayed. Word's deletion takes
   * it with the text (check 54 cases 27 and 28; CR-014 section 12.6).
   */
  private commentReferencesIn(start: number, end: number, edges?: SpanEdges): { run: Element; owner: Element[]; id: number }[] {
    const out: { run: Element; owner: Element[]; id: number }[] = [];
    this.segments({ visit: (item, at, _owner, run) => {
      if (item.name?.localPart !== 'commentReference' || !run || !objectInSpan(at, start, end, edges)) return;
      if (out.some((r) => r.run === run.element)) return;
      if (!((run.element.value as wml.R).content ?? []).every((c) => (c as Element).name?.localPart === 'commentReference')) return;
      out.push({ run: run.element as Element, owner: run.owner, id: Number((item.value as { id?: unknown }).id) });
    } });
    return out;
  }

  /** A comment queued to go when the package is saved, its markers, replies and entries with it, as a resolution queues one. */
  private queueCommentRemoval(id: number): void {
    if (!Number.isInteger(id)) return;
    (this.parentBody.package_ as { pendingCommentRemovals?: Set<number> } | undefined)?.pendingCommentRemovals?.add(id);
  }

  /**
   * Splits the revision that holds the text at `at` into two - the second half a copy of its
   * attributes under a new `w:id` - and puts `item` between them. Word splits a revision so when
   * text is typed inside it: around the typist's `w:ins` with tracking on, around plain text with it
   * off (`test/README.md` checks 20 and 21). False, with nothing changed but a run split at `at`,
   * where the text there is not directly in the revision (in a hyperlink inside it, say): the caller
   * then does what it did before.
   */
  private insertBetweenHalves(at: number, item: Element | Element[], nextId: () => number): boolean {
    this.splitAt(at);
    const after = this.segments().find((s) => s.start >= at && s.revision);
    const holder = after?.revision;
    const index = holder && after ? holder.items.indexOf(after.runOwner[after.runIndex]!) : -1;
    if (!holder || index <= 0) return false;
    return this.splitHolderAround(holder, index, item, nextId);
  }

  /**
   * Moves a revision's items from `index` on into a copy of it - its attributes, a new `w:id` - and
   * puts `item` between the two; with nothing from `index` on, `item` just goes after the revision.
   * False, with nothing changed, where the revision's items cannot be told apart from its attributes.
   * Internal: `Range.insertContentControl` uses it to put a new control between the halves of the
   * insertion it wraps part of (check 30).
   */
  splitHolderAround(holder: RevisionHolderOf, index: number, item: Element | Element[], nextId: () => number): boolean {
    const value = holder.value as unknown as Record<string, unknown>;
    const itemsKey = Object.keys(value).find((key) => value[key] === holder.items);
    if (!itemsKey) return false;
    const parent = (holder.value as { PARENT?: object }).PARENT ?? this.p;
    const at = holder.owner.indexOf(holder.element) + 1;
    const items = Array.isArray(item) ? item : [item];
    if (index >= holder.items.length) {
      holder.owner.splice(at, 0, ...items);
      linkParents(items, parent);
      return true;
    }
    const copy: Record<string, unknown> = {};
    for (const [key, v] of Object.entries(value)) if (key !== itemsKey) copy[key] = v;
    copy.id = nextId();
    copy[itemsKey] = holder.items.splice(index);
    const second = { name: holder.element.name, value: copy } as Element;
    holder.owner.splice(at, 0, ...items, second);
    linkParents(copy[itemsKey], copy);
    linkParents([...items, second], parent);
    return true;
  }

  /**
   * Puts `item` beside a revision rather than in it: after it, or before it - and then before the
   * range start of a move whose range opens just there, so that text typed at the start of moved
   * text is not inside the move (`test/README.md` check 20 case 03).
   */
  private insertBesideRevision(holder: RevisionHolder, item: Element | Element[], side: 'before' | 'after'): void {
    let index = holder.owner.indexOf(holder.element) + (side === 'after' ? 1 : 0);
    if (side === 'before' && (holder.kind === 'moveTo' || holder.kind === 'moveFrom')) {
      const marker = holder.kind === 'moveTo' ? 'moveToRangeStart' : 'moveFromRangeStart';
      while (index > 0 && holder.owner[index - 1]?.name?.localPart === marker) index--;
    }
    const items = Array.isArray(item) ? item : [item];
    holder.owner.splice(index, 0, ...items);
    linkParents(items, (holder.value as { PARENT?: object }).PARENT ?? this.p);
  }

  /**
   * A fresh annotation id for markup this paragraph writes itself: the package's counter when there
   * is a package that keeps one, since a document has one id space; else one above every id in the
   * story.
   */
  /** @internal `Comment` uses it to split an insertion at a comment's end. */
  nextAnnotationId(): number {
    const pkg = this.parentBody.package_ as Partial<TrackingHost> | undefined;
    if (pkg && typeof pkg.markupRoots === 'function' && 'annotationIdFloor' in pkg) {
      return new ChangeTracker(pkg as TrackingHost, 'TrackAll').nextId();
    }
    let root: object = this.p;
    for (let o: object | undefined = this.p; o !== undefined; o = (o as { PARENT?: object }).PARENT) root = o;
    return highestAnnotationId(root) + 1;
  }

  /**
   * Text inserted at `at` as a `w:ins`. A run of this author's own that is already inside a
   * `w:ins` is extended rather than nested in another one, as Word does. `anchor` is the `w:del`
   * of a replacement, which the insertion must follow.
   */
  private insertTracked(tracker: ChangeTracker, at: number, text: string, anchor?: Anchor, prefer: 'before' | 'after' = 'before'): void {
    if (anchor) {
      const rPr = copyRPr(runsOf(anchor.element.value as object, { view: 'original' })[0]?.value.rPr);
      const ins = tracker.ins([runOf([textItem(text)], rPr)]);
      // A replacement inside another's revision - an insertion, moved text - leaves its w:del in
      // that revision; the new text is not put in it too, but the revision split after the deletion
      // and the text between the halves, as for anything typed in it (check 21; the editor's finding
      // on 0.3.0, CR-002 section 37).
      const holder = anchor.revision;
      if (holder && this.splitHolderAround(holder, holder.items.indexOf(anchor.element) + 1, ins as Element, () => tracker.nextId())) return;
      anchor.owner.splice(anchor.owner.indexOf(anchor.element) + 1, 0, ins);
      linkParents(ins, anchor.parent);
      return;
    }
    this.splitAt(at);
    const segs = this.segments();
    const before = [...segs].reverse().find((s) => s.end <= at);
    const after = segs.find((s) => s.start >= at);
    const sides = [[before, 'after'], [after, 'before']] as const;
    for (const [seg, side] of prefer === 'after' ? [...sides].reverse() : sides) {
      if (!seg || !tracker.ownInsertion(seg.revision)) continue;
      if (seg.editable) {
        const t = seg.item.value as wml.Text;
        setText(t, side === 'after' ? seg.text + text : text + seg.text);
      } else {
        const run = runOf([textItem(text)], copyRPr(seg.run.rPr));
        seg.runOwner.splice(side === 'after' ? seg.runIndex + 1 : seg.runIndex, 0, run);
        linkParents(run, seg.revision!.value);
      }
      return;
    }
    const neighbour = prefer === 'after' ? after ?? before : before ?? after;
    tracker.assertEditable(neighbour?.revision);
    const rPr = copyRPr(neighbour?.run.rPr ?? this.visibleFirstRunRPr());
    const ins = tracker.ins([runOf([textItem(text)], rPr)]);
    this.placeInsertion(at, [ins as Element], before, after, () => tracker.nextId(), prefer);
  }

  /**
   * An insertion put at `at`, `before` and `after` being the segments either side of it once the runs
   * are split there: never inside another's revision.
   */
  private placeInsertion(at: number, items: Element[], before: TextSegment | undefined, after: TextSegment | undefined, nextId: () => number, prefer: 'before' | 'after' = 'before'): void {
    // Typed inside another's revision - a w:ins, a w:moveTo - the revision is split and the typist's
    // w:ins goes between its halves, as Word writes it (test/README.md check 20 case 01, check 21's
    // first run); before, it went to the revision's edge, putting the text in the wrong place.
    if (before?.revision && after?.revision && before.revision.element === after.revision.element
      && this.insertBetweenHalves(at, items, nextId)) {
      return;
    }
    // at the start of a revision, beside it - and outside a move's range that opens there (case 03)
    if (!before && after?.revision) {
      this.insertBesideRevision(after.revision, items, 'before');
      return;
    }
    const neighbour = prefer === 'after' ? after ?? before : before ?? after;
    if (neighbour) {
      const owner = neighbour.revision ? neighbour.revision.owner : neighbour.runOwner;
      const item = neighbour.revision ? neighbour.revision.element : neighbour.runOwner[neighbour.runIndex]!;
      const parent = (neighbour.revision ? neighbour.revision.value : neighbour.run) as { PARENT?: object };
      owner.splice(owner.indexOf(item) + (neighbour === before ? 1 : 0), 0, ...items);
      linkParents(items, parent.PARENT ?? this.p);
    } else {
      const content = (this.p.content ??= []) as Element[];
      if (at === 0) content.unshift(...items); else content.push(...items);
      linkParents(items, this.p);
    }
  }

  /**
   * Splits runs so that every run holding text in [start, end) holds nothing else (internal: what a
   * `Range.font` write formats). A kept `mc:AlternateContent`'s branch splits with its run (section 12.2).
   */
  splitRunsAround(start: number, end: number): void {
    this.isolate(start, end);
  }

  /**
   * Splits runs so that every run holding text in [start, end) holds nothing else, and returns
   * the segments inside the span. `splitAt` does the `w:t` boundaries; this also moves the items
   * of a run that straddles a boundary (a tab, a break, a drawing) out into runs of their own.
   * The pictures in the span (`splice`'s `edges`) are among the segments returned, and a run
   * holding one holds nothing outside the span either.
   */
  private isolate(start: number, end: number, edges?: SpanEdges): TextSegment[] {
    this.splitAt(start);
    this.splitAt(end);
    const inSpan = (s: TextSegment): boolean => s.object ? objectInSpan(s.start, start, end, edges) : s.start >= start && s.end <= end && s.text.length > 0;
    for (let guard = 0; guard < 10_000; guard++) {
      const inside = this.segments({ objects: true }).filter(inSpan);
      let split = false;
      const seen = new Set<wml.R>();
      for (const seg of inside) {
        if (seen.has(seg.run)) continue;
        seen.add(seg.run);
        if (seg.math) continue;                                    // in no run: taken whole or refused
        const ofRun = inside.filter((s) => s.run === seg.run);
        const first = ofRun[0]!;
        const last = ofRun[ofRun.length - 1]!;
        if (this.splitRunBefore(first)) { split = true; break; }
        if (this.splitRunAfter(last)) { split = true; break; }
      }
      if (!split) return inside;
    }
    return this.segments({ objects: true }).filter(inSpan);
  }

  /**
   * Refuses an edit that would cut an equation (CR-012 phase 4). Its text is one segment, Word's drawing of
   * its math runs (`mathTextOf`), so an edit may take it whole but not part of it, nor insert into it. Under change
   * tracking a deletion takes it whole in Word's form (`deleteText`, CR-002 section 62.2 D2, check 57), and a
   * replacement doesn't take it at all: where Word puts the new text beside a deleted equation isn't measured. Since
   * phase 4 every other offset is a place in the runs: `text` is the segments' text, where before it was the objects
   * package's `textOf` and section 12.4 refused an offset past a kept `mc:AlternateContent` among the runs.
   */
  private assertPlaceable(start: number, end: number, tracked: false | 'delete' | 'replace'): void {
    for (const seg of this.segments()) {
      if (!seg.math) continue;
      const cuts = (seg.start < start && start < seg.end) || (seg.start < end && end < seg.end);
      const takes = start < seg.end && seg.start < end;
      if (!cuts && !(tracked === 'replace' && takes)) continue;
      throw new NotSupportedError('Paragraph', 'edit', `[${start}, ${end}) ${cuts ? 'cuts' : 'takes'} the equation at [${seg.start}, ${seg.end})`
        + (cuts ? ': an edit takes an equation whole or not at all (CR-012 phase 4)'
          : ' under change tracking, with text in its place: a tracked replacement of an equation is not measured (CR-002 section 62.2)'));
    }
  }

  /** Moves the items before `seg` into a run of their own, in front of the rest; false where there are none. */
  private splitRunBefore(seg: TextSegment): boolean {
    if (seg.index > 0) { this.splitRunAtItem(seg, seg.owner, seg.index); return true; }
    const alternate = seg.alternate?.level === 'run' && seg.owner === seg.alternate.branch ? seg.alternate : undefined;
    if (!alternate) return false;
    const at = alternate.owner.indexOf(alternate.element);
    if (at <= 0) return false;
    this.splitRunAtItem(seg, alternate.owner, at);
    return true;
  }

  /** Moves the items after `seg` into a run of their own, behind it; false where there are none. */
  private splitRunAfter(seg: TextSegment): boolean {
    if (seg.index < seg.owner.length - 1) { this.splitRunAtItem(seg, seg.owner, seg.index + 1); return true; }
    const alternate = seg.alternate?.level === 'run' && seg.owner === seg.alternate.branch ? seg.alternate : undefined;
    if (!alternate) return false;
    const at = alternate.owner.indexOf(alternate.element);
    if (at < 0 || at >= alternate.owner.length - 1) return false;
    this.splitRunAtItem(seg, alternate.owner, at + 1);
    return true;
  }
}

type RevisionHolderOf = NonNullable<TextSegment['revision']>;

/** Where a `w:ins` goes when it must follow the `w:del` of a replacement. */
interface Anchor {
  owner: Element[];
  element: Element;
  parent: object;
  /** The revision the `w:del` sits directly in - another author's insertion or moved text - if any. */
  revision?: RevisionHolderOf;
}

export { W_NS };

// --- reading one w:pPr, direct or resolved -----------------------------------------------

function alignmentOf(pPr: wml.PPr | undefined, effective: boolean): AlignmentOrUnknown {
  switch (pPr?.jc?.val) {
    case 'left': case 'start': return 'Left';
    case 'center': return 'Centered';
    case 'right': case 'end': return 'Right';
    case 'both': case 'distribute': return 'Justified';
    // Word lays a paragraph with no w:jc out left-aligned, so that is the effective value
    default: return effective ? 'Left' : 'Unknown';
  }
}

function firstLineIndentOf(pPr: wml.PPr | undefined): number {
  const ind = pPr?.ind;
  if (ind?.hanging !== undefined) return -ind.hanging / TWIPS_PER_POINT;
  return (ind?.firstLine ?? 0) / TWIPS_PER_POINT;
}

function lineSpacingOf(pPr: wml.PPr | undefined): number {
  const spacing = pPr?.spacing;
  if (spacing?.line === undefined) return 0;
  return spacing.lineRule === undefined || spacing.lineRule === 'auto'
    ? (spacing.line / 240) * 12 : spacing.line / TWIPS_PER_POINT;
}

function outlineLevelOf(pPr: wml.PPr | undefined): number {
  const level = pPr?.outlineLvl?.val;
  return level === undefined ? 10 : level + 1;
}

/**
 * Sets a `w:t`'s text with `xml:space="preserve"` exactly when the text needs it - whitespace at
 * either end, or two whitespace characters together - which is how the builders' `t()` and Word
 * write it. Setting the attribute and never taking it off left a text shortened from "Font styles "
 * to "Font X" still preserving, the same text but not the same bytes as a `w:t` made afresh
 * (CR-002 section 32 item 2).
 */
function setText(t: wml.Text, value: string): void {
  t.value = value;
  if (/^\s|\s$|\s\s/.test(value)) t.space = 'preserve';
  else delete t.space;
}


/**
 * The element a run-level `mc:AlternateContent`'s later piece is written as (CR-012 section 12.2; the
 * editor's ED-009 section 3.4, P2): its branches in order with their `Requires`, the selected one holding
 * `items` and every other one empty.
 */
function alternatePiece(alternate: AlternateLocation, items: Element[]): Element {
  const ac = alternate.element.value as mce.AlternateContent & { choice?: { any?: unknown[] }[]; fallback?: { any?: unknown[] } };
  const emptied = (branch: { any?: unknown[] } | undefined): { any?: unknown[] } | undefined => {
    if (branch === undefined) return undefined;
    const copy: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(branch)) if (key !== 'any' && key !== 'PARENT') copy[key] = value;
    copy.any = branch.any === alternate.branch ? items : [];
    return copy as { any?: unknown[] };
  };
  const value: Record<string, unknown> = {};
  for (const [key, v] of Object.entries(ac)) {
    if (key === 'PARENT') continue;
    if (key === 'choice') value.choice = (v as { any?: unknown[] }[]).map(emptied);
    else if (key === 'fallback') value.fallback = emptied(v as { any?: unknown[] });
    else value[key] = v;
  }
  return { name: alternate.element.name, value } as Element;
}

/**
 * After an untracked deletion: a run-level `mc:AlternateContent` whose selected branch the deletion
 * emptied goes, every branch with it (CR-012 section 12.2 item 4; ED-009 section 3.5). A branch that
 * was empty before is left as it is.
 */
function removeEmptiedAlternates(removed: TextSegment[]): void {
  const seen = new Set<Element>();
  for (const seg of removed) {
    const alternate = seg.alternate;
    if (alternate?.level !== 'run' || seen.has(alternate.element)) continue;
    seen.add(alternate.element);
    if (alternate.branch.length > 0) continue;
    const i = alternate.owner.indexOf(alternate.element);
    if (i >= 0) alternate.owner.splice(i, 1);
  }
}

/** Block holders a cell's paragraph may stand in: a block content control, its content, custom XML. */
const BLOCK_HOLDERS = new Set(['org_docx4j_wml.SdtBlock', 'org_docx4j_wml.SdtContentBlock', 'org_docx4j_wml.CTSdtContentBlock', 'org_docx4j_wml.CTCustomXmlBlock']);

/** True for the last paragraph of a table cell, through the block holders between them. */
function lastInCell(p: wml.P): boolean {
  let holder = (p as { PARENT?: object }).PARENT;
  for (let guard = 0; holder && guard < 16 && BLOCK_HOLDERS.has((holder as { TYPE_NAME?: string }).TYPE_NAME ?? ''); guard++) {
    holder = (holder as { PARENT?: object }).PARENT;
  }
  if ((holder as { TYPE_NAME?: string } | undefined)?.TYPE_NAME !== 'org_docx4j_wml.Tc') return false;
  let last: wml.P | undefined;
  const visit = (items: Element[] | undefined): void => {
    for (const el of items ?? []) {
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.P') last = el.value as wml.P;
      else if (tn !== 'org_docx4j_wml.Tbl' && typeof el.value === 'object' && el.value !== null) {
        const v = el.value as { content?: Element[]; sdtContent?: { content?: Element[] } };
        visit(v.sdtContent ? v.sdtContent.content : v.content);
      }
    }
  };
  visit((holder as wml.Tc).content as Element[] | undefined);
  return last === p;
}

/** Every element among a paragraph's runs, in document order: where a deletion places a comment's reference among its runs. */
function runOrderOf(p: wml.P): Map<object, number> {
  const order = new Map<object, number>();
  const visit = (items: Element[] | undefined): void => {
    for (const el of items ?? []) {
      order.set(el, order.size);
      const v = el.value;
      if (typeof v === 'object' && v !== null && typeNameOf(el) !== 'org_docx4j_wml.R') visit(runItemsOf(v) as Element[] | undefined);
    }
  };
  visit(runItemsOf(p) as Element[] | undefined);
  return order;
}

/** The `w:ins` (or other revision) a run stands directly in, read from its `PARENT`, if any. */
function revisionHolding(run: Element): RevisionHolder | undefined {
  const holder = (run.value as { PARENT?: object }).PARENT as (wml.CTTrackChange & { PARENT?: object; customXmlOrSmartTagOrSdt?: Element[] }) | undefined;
  const items = holder?.customXmlOrSmartTagOrSdt;
  const outer = holder?.PARENT ? runItemsOf(holder.PARENT) as Element[] | undefined : undefined;
  const element = outer?.find((e) => e.value === holder);
  const kind = element ? revisionKindOf(element) : undefined;
  if (!holder || !items || !outer || !element || !kind) return undefined;
  return { kind, element, value: holder, items, owner: outer };
}

/** The ids of the comment references among runs. */
function commentReferenceIds(runs: Element[]): number[] {
  return runs.flatMap((run) => (((run.value as wml.R).content ?? []) as Element[])
    .filter((c) => c.name?.localPart === 'commentReference').map((c) => Number((c.value as { id?: unknown }).id)));
}

/** The `w:sdt` value types, block, run, row and cell. */
const SDT_VALUE_TYPES: ReadonlySet<string> = new Set(['org_docx4j_wml.SdtBlock', 'org_docx4j_wml.SdtRun', 'org_docx4j_wml.CTSdtRow', 'org_docx4j_wml.CTSdtCell']);

/** Whether a paragraph holds a run-level control bound by `w:dataBinding`, at any depth among its runs. */
function holdsBoundRunControl(holder: object): boolean {
  for (const item of runItemsOf(holder) ?? []) {
    const value = item.value as { TYPE_NAME?: string; sdtPr?: wml.SdtPr } | undefined;
    if (!value || typeof value !== 'object') continue;
    if (value.TYPE_NAME === 'org_docx4j_wml.SdtRun' && sdtProperty(value.sdtPr, 'dataBinding') !== undefined) return true;
    if (value.TYPE_NAME !== 'org_docx4j_wml.R' && holdsBoundRunControl(value)) return true;
  }
  return false;
}
