import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import { Docx4JException, PropertyResolverNotCreatedException } from '../../opc/exceptions.mjs';
import { NotSupportedError } from '../../office-js/errors.mjs';
import type { PropertyResolver } from '../properties/PropertyResolver.mjs';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import type { OpcPackage } from '../../packages/OpcPackage.mjs';
import { type Element, type Located, type TextViewOptions, type TextView, type HiddenRuns, type VanishResolver, hiddenRunsOf, typeNameOf, childrenOf, linkParents, BLOCK_LEVEL_TYPES, rowsOf, cellsOf, selectedBranchOf, branchesOf, branchListsOf, ALTERNATE_CONTENT, BRANCH_TYPES, inUnreadBranch, assertOutsideBranch, writingEveryBranch, type BranchMode } from './tree.mjs';
import { br as breakItem, tbl as tableOf } from '@docx4j/generated-objects-ts/builders/wml';
import { oneParagraphRun } from './officeText.mjs';
import { runOf, paragraphOf, runsOf } from './tree.mjs';
import { textBoxesOf } from './textBoxes.mjs';
import { outlineReader } from './outline.mjs';
import { Paragraph } from './Paragraph.mjs';
import { Range } from './Range.mjs';
import { Table, type TableRow, type TableCell } from './Table.mjs';
import { ContentControl, collectControls, assertControlHolds, controlsOfTypes, type BlockRefusal, type ContentControlType, type ContentControlOptions } from './ContentControl.mjs';
import { sdt as sdtOf, nextSdtId } from '@docx4j/generated-objects-ts/builders/wml';
import { controlIdScope, sdtKindFor, declareSdtExtensions } from '../customxml/insert.mjs';
import { InlinePicture, addImage, writableWidthEmu, type InlinePictureOptions } from './InlinePicture.mjs';
import { contentOf } from './ooxml.mjs';
import { relationshipIdsIn, queueRelationshipRemovals, type RelationshipRemovals } from './relationships.mjs';
import { type CellEnd, wrappingCellOf, paragraphEndOf, endsWithParagraph, endWithParagraph, endCellsWithParagraphs } from './tableRules.mjs';
import { keepCellMarkLast } from './spanEdits.mjs';
import { expandReplacement, paragraphMarkPieces, refuseMarkReplacement, refuseWildcardMarkReplacement, type SearchOptions } from './search.mjs';
import type { Bound } from './binder.mjs';

/**
 * What `Body.select` returns for one hit: the content-API view of the selected element where there
 * is one, and the `Bound` itself otherwise (CR-006 section 2).
 */
export type SelectResult = Paragraph | Table | TableRow | TableCell | ContentControl | Range | Bound;
import { commentApi } from './comments.mjs';
import { storyLayout } from './story.mjs';
import type { Comment } from './Comment.mjs';
import { type ChangeTracker, trackerOf, trackInsertedBlock, markHolderInserted } from './tracking.mjs';
import { type TrackedChange, type TrackedChangeToken, type Listing, trackedChangesOfRow, trackedChangeTokensOfParagraph, groupTouching, BREAK, tablePropertiesChangeOf, cellChangeOf, bodySectionChangeOf, resolveUnlisted, holderChangeAt, resolving, resolvePieces, alternateTokens, listedListing, pairTwins } from './TrackedChange.mjs';
import { pieceMatcher, type TrackedChangeSelection } from './trackedChangeFilter.mjs';
import { type List, type ListLabel, listsOf, listLabelsOf } from './List.mjs';

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
  /** The text model's text of the paragraph (CR-012 D3): hidden text left out, a deleted mark joined, in the options' view. */
  text: string;
  /** True for a paragraph in a text box, listed after its anchor at a `textbox:` address (CR-012 D3, D4). */
  textbox?: true;
  /**
   * The addresses of the paragraphs the view joins to this one, in order (a deleted mark in the `accepted`
   * view, an inserted one in `original`): their text is in `text`, and they have no entry of their own
   * (CR-012 section 17.13).
   */
  joined?: string[];
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

/** Which of a body's blocks `Body.outline` reads (CR-012 section 17.13): its top-level blocks, by index or view. */
export interface OutlineSelection {
  blocks: readonly (number | Paragraph | BlockElement)[];
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

  /**
   * The package's `PropertyResolver`, which every effective-formatting read goes through
   * (CR-001 Phase B step 2). Throws `PropertyResolverNotCreatedException` when this body has
   * no package, or the package has not built one yet: building it unmarshals the styles part,
   * which cannot be done synchronously, so `await pkg.getPropertyResolver()` (or `getBody()`,
   * which does it) first - or read direct formatting with `{ direct: true }`.
   */
  get propertyResolver(): PropertyResolver {
    const pkg = this.package_ as { propertyResolverOrUndefined?: PropertyResolver } | undefined;
    const resolver = pkg?.propertyResolverOrUndefined;
    if (!resolver) throw new PropertyResolverNotCreatedException();
    return resolver;
  }

  /** Which of a paragraph's runs are hidden: a run's own `w:vanish`, else its styles' (`hiddenRunsOf`; CR-012 phase 4). */
  hiddenRuns(p: wml.P): HiddenRuns {
    const pkg = this.package_ as { propertyResolverOrUndefined?: VanishResolver } | undefined;
    return hiddenRunsOf(pkg?.propertyResolverOrUndefined, p);
  }

  /** The live content array (docx4j getContent()); block-level elements only, sectPr excluded. */
  get content(): Element[] {
    return (this.container.content ??= []);
  }

  /**
   * Every paragraph in document order, descending into tables and content controls, and into the
   * selected branch of a kept `mc:AlternateContent` at block, row and cell level, as Word reads it
   * (CR-013 phase 2). A paragraph in a branch is read-only (D4).
   */
  get paragraphs(): Paragraph[] {
    return this.paragraphsIn('read');
  }

  /**
   * The paragraphs a walk in this mode takes: a reader's, `paragraphs`; a bulk edit's (`'every'`), every
   * branch's too, the selected branch's where a reader sees them and each other branch's after them,
   * where the element stands (CR-013 phase 3).
   */
  private paragraphsIn(mode: BranchMode): Paragraph[] {
    const out: Paragraph[] = [];
    const visit = (items: Element[]): void => {
      for (const el of items) {
        const tn = typeNameOf(el);
        if (tn === 'org_docx4j_wml.P') out.push(new Paragraph(el as Element<wml.P>, items, this));
        else {
          const v = el.value;
          if (typeof v !== 'object' || v === null) continue;
          if (tn === 'org_docx4j_wml.Tbl') {
            for (const row of rowsOf(v, mode)) for (const cell of cellsOf(row.element.value, mode)) visit(childrenOf(cell.element.value) ?? []);
          } else if (tn === ALTERNATE_CONTENT) {
            for (const branch of branchListsOf(v as mce.AlternateContent, mode)) visit(branch);
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

  /**
   * The content-API views an XPath selects (CR-006): a `ContentControl`, `Table`, `TableRow`,
   * `TableCell` or `Paragraph` where the selected element has one, at any depth - an inline content
   * control is a `ContentControl`, a nested table a `Table`; a `Range` over the whole paragraph for a
   * `w:r` or anything in one (`w:t`, `w:tab`), wherever the run sits (a hyperlink, an insertion, an
   * inline control); and the `Bound` itself for anything else - an attribute, a text node, a `w:pPr`,
   * a `w:hyperlink`, a wildcard child. An attribute or a text node is always the `Bound`, never its
   * element's view, so that what matched is not lost. A caller that wants the tree objects rather
   * than views uses `pkg.selectObjects`.
   *
   * The expression is evaluated over a snapshot of the part this body belongs to, so it sees the
   * whole part - `//w:p` from a header's body selects that header's paragraphs. Every call marshals
   * the part afresh (CR-006 section 2).
   */
  async select(xpath: string, options: { namespaces?: Record<string, string> } = {}): Promise<SelectResult[]> {
    const pkg = this.package_ as { selectObjects?: (x: string, o?: unknown) => Promise<Bound[]> } | undefined;
    if (pkg?.selectObjects === undefined) throw new Docx4JException('select needs the package that owns this body');
    const bounds = await pkg.selectObjects(xpath, { part: this.part, namespaces: options.namespaces });
    const views = this.viewsByObject();
    return bounds.map((bound) => (bound.node.nodeType === 1 ? views.get(bound.object as object) : undefined) ?? bound);
  }

  /**
   * Every view this body can make, keyed by the tree object behind it, so that a selected object
   * becomes the view a caller can act on. Built per select: the views are cheap wrappers and nothing
   * is cached but the `Body` itself (CR-002's "views, not a model"). The first view put for an object
   * is the one it keeps, so the most specific kinds go first: a content control before the
   * paragraph that holds an inline one, a table at any depth before the paragraphs in its cells, and
   * a paragraph before the runs in it. A `w:r` and what is in it map to a `Range` over the paragraph,
   * which is the nearest thing the content API has - a run is not addressable on its own in Office
   * JS. A run in a text box is not one of the paragraph's, and stays a `Bound`.
   */
  private viewsByObject(): Map<object, SelectResult> {
    const out = new Map<object, SelectResult>();
    const put = (value: unknown, view: SelectResult): void => {
      if (typeof value === 'object' && value !== null && !out.has(value)) out.set(value, view);
    };
    const controls = this.contentControls;
    for (const control of controls) put(control.element.value, control);
    const putTables = (tables: Table[]): void => {
      for (const table of tables) {
        put(table.element.value, table);
        for (const row of table.rows) {
          put(row.element.value, row);
          for (const cell of row.cells) {
            put(cell.element.value, cell);
            putTables(cell.tables);
          }
        }
      }
    };
    putTables(this.tables);
    for (const control of controls) if (control.form === 'Block') putTables(control.tables);
    for (const paragraph of this.paragraphs) {
      put(paragraph.element.value, paragraph);
      const range = paragraph.getRange();
      // both views, so that a run in a w:del is found as well as one in a w:ins
      for (const view of ['accepted', 'original'] as const) {
        for (const run of runsOf(paragraph.element.value, { view })) {
          put(run.value, range);
          for (const item of run.value.content ?? []) put(item.value, range);
        }
      }
    }
    return out;
  }

  /**
   * The tables directly in this body, as Office JS reports them (nested tables are a cell's), a table in
   * the selected branch of a kept block-level `mc:AlternateContent` among them (CR-013 phase 2).
   */
  get tables(): Table[] {
    const out: Table[] = [];
    const visit = (items: Element[]): void => {
      for (const element of items) {
        const tn = typeNameOf(element);
        if (tn === 'org_docx4j_wml.Tbl') out.push(new Table(element as Element<wml.Tbl>, items, this));
        else if (tn === ALTERNATE_CONTENT) visit(selectedBranchOf(element.value as mce.AlternateContent) ?? []);
      }
    };
    visit(this.content);
    return out;
  }

  /** Every content control in this body, in document order, nested ones included. */
  get contentControls(): ContentControl[] {
    return this.contentControlsIn('read');
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
   * The content controls a walk in this mode takes: a reader's, `contentControls`; the binding passes'
   * (`'every'`), every branch's of a kept `mc:AlternateContent` too, each other branch's where the element
   * stands (CR-013 phase 3). An extension.
   */
  contentControlsIn(mode: BranchMode): ContentControl[] {
    const out: ContentControl[] = [];
    collectControls(this.content, this, out, mode);
    return out;
  }

  /** Every inline picture in this body, in document order (Office JS Body.inlinePictures). */
  get inlinePictures(): InlinePicture[] {
    return this.paragraphs.flatMap((p) => p.inlinePictures);
  }

  /**
   * Office JS `body.lists`: every `w:num` a paragraph of this body names, in the order they are
   * first used (CR-002 phase H). Empty where the document has no numbering part, or where
   * nothing has read it yet (`await pkg.getBody()` builds the resolver, which does).
   */
  get lists(): List[] {
    return listsOf(this);
  }

  /**
   * Every numbered paragraph of this body's **story** with its label, level, list and sibling
   * index, in one walk (extension). `listItem.listString` walks the story per read, which is
   * what Word's counting needs; this is the way to ask for all of them at once.
   *
   * Keyed by the `w:p` itself: `labels.get(paragraph.p)?.listString`.
   */
  listLabels(): Map<wml.P, ListLabel> {
    return listLabelsOf(this);
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

  /**
   * The text as Office JS's `body.text` gives it in Word's default view (CR-012 phase 4; checks 30, 31, 48
   * and 51): each paragraph's text (`Paragraph.text`) followed by its mark, `\r`, the last paragraph's left
   * out; in a table a row's cells a `\t` apart, each cell's paragraphs a mark apart (its last mark not
   * written), and the row ended by `\r\n`. Text boxes are not the body's (check 42).
   */
  get text(): string {
    return this.textIn('accepted');
  }

  /**
   * The text in one of the two views (extension; `text` is the accepted one).
   * `{ view: 'original' }` reads the document as it was before the tracked changes.
   */
  getText(options?: TextViewOptions): string {
    return this.textIn(options?.view ?? 'accepted');
  }

  /**
   * The body's range (Office JS `Body.getRange`; CR-014 phase 1): 'Whole' from its first paragraph's start to
   * its last one's end, with that paragraph's mark (check 51: a cell's `CELLP1\rCELLP2\rCELLP3\t`, a note's
   * `… NOTEP2\r`), 'Content' the same without the mark, 'Start' and 'End' collapsed at its ends.
   */
  getRange(location: 'Whole' | 'Start' | 'End' | 'Content' = 'Whole'): Range {
    const paragraphs = this.paragraphs;
    const first = paragraphs[0];
    if (!first) throw new Docx4JException('This body holds no paragraph');
    const last = paragraphs[paragraphs.length - 1]!;
    if (location === 'Start') return first.getRange('Start');
    if (location === 'End') return last.getRange('End');
    if (paragraphs.length === 1) return first.getRange(location);
    return new Range(first, 0, last.text.length, true, location === 'Whole', last);
  }

  private textIn(view: TextView): string {
    // the original view's marks are the ones Reject All keeps: a deleted mark written, an inserted one not
    const { text } = storyLayout(this.content, (el, items) => {
      const p = new Paragraph(el, items, this);
      return view === 'original' ? p.getText({ view }) : p.text;
    }, view === 'original' ? (p) => p.pPr?.rPr?.ins === undefined : undefined);
    // the last paragraph's mark is not written
    return text.endsWith('\r') ? text.slice(0, -1) : text;
  }

  /** The package's change tracker while `changeTrackingMode` is on; undefined when it is off (CR-002 phase F). */
  get changeTracker(): ChangeTracker | undefined {
    return trackerOf(this.package_)?.forPart(this.part);
  }

  /** Refuses a write to a body in a branch of a kept `mc:AlternateContent`: a cell or a control a reader found there (CR-013 D4). */
  private assertWritable(): void {
    assertOutsideBranch(this.container, 'Body');
  }

  insertParagraph(text: string, location: 'Start' | 'End'): Paragraph {
    const run = oneParagraphRun(text);   // CR-002 section 60 E2: one paragraph, its line breaks w:br and w:cr
    const el = paragraphOf(run.value.content?.length ? [run] : []);
    return this.insertElement(el, location) as Paragraph;
  }

  /** Text at the start of the first paragraph, the end of the last, or replacing everything with one paragraph. */
  insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range {
    this.assertWritable();
    const paragraphs = this.paragraphs;
    if (location === 'Replace' || paragraphs.length === 0) {
      this.clearContent(false);   // its own paragraph ends a cell, so the cell's last one goes too
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
    // Word's own default (Insert Table, fixed column width Auto; the editor's check 42, Word 365): an
    // automatic table width, no w:tblLayout, the cells in twips sharing the text width - the grid and
    // cells the builder wrote, the table's width made automatic (CR-002 section 42.5 item 7).
    const tblPr = (element.value.tblPr ??= {});
    tblPr.tblW = { w: 0, type: 'auto' };
    this.insertElement(element as Element, location, undefined, 'InvalidArgument');
    return new Table(element, this.content, this);
  }

  /**
   * Adds the image as a part of this body's part, with a relationship, and shows it in a new
   * paragraph at the start or the end (Office JS insertInlinePictureFromBase64). An image wider
   * than the text area is scaled down, as docx4j's CxCy.scale does.
   */
  insertInlinePictureFromBase64(base64: string, location: 'Start' | 'End', options?: InlinePictureOptions): InlinePicture {
    this.assertWritable();
    if (!this.part) throw new Docx4JException('This body has no part, so an image cannot be related to it');
    const { run, drawing } = addImage(this.part, base64, this.container, writableWidthEmu(this.container), options);
    const paragraph = this.insertParagraph('', location);
    // under tracking the picture is an insertion, as its paragraph's mark is (insertBreak's form)
    const tracker = paragraph.changeTracker;
    const item = tracker ? tracker.ins([run]) : run;
    paragraph.p.content!.push(item as never);
    linkParents(item, paragraph.p);
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
  insertElement(element: Element | Element[], location: 'Start' | 'End' | 'Before' | 'After', target?: Paragraph | BlockElement | Address, refusal?: BlockRefusal, replacing?: Element): Paragraph | BlockElement {
    this.assertWritable();
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
      assertOutsideBranch(t.element.value as object, 'Body', 'the block to insert beside');   // beside it is in the branch too
      container = t.container;
      index = container.indexOf(t.element) + (location === 'After' ? 1 : 0);
    }
    // a bound block control holds one paragraph (CR-002 section 61): `replacing` is the block the write takes away after
    assertControlHolds(this, this.ownerOf(container), replacing ? container.filter((e) => e !== replacing) : container, elements, refusal);
    // a cell ends with a paragraph, as Word reads it: the cells of what goes in (D3), and the cell it goes in (CR-002 section 55)
    endCellsWithParagraphs(elements);
    const owner = this.ownerOf(container);
    const cell = wrappingCellOf(owner);
    const end = cell ? paragraphEndOf(cell.content as Element[] | undefined) : undefined;
    container.splice(index, 0, ...elements);
    const tracker = this.changeTracker;
    for (const el of elements) {
      linkParents(el, owner);
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.P') this.assignParaId(el.value as wml.P);
      // everything it holds is an insertion: the runs in its links and inline controls, the rows and
      // nested tables of a table, a block control's content and the control itself (CR-002 section 37)
      if (tracker) trackInsertedBlock(tracker, el, container);
    }
    if (cell) this.keepCellEnd(cell, end, container);
    const first = elements[0]!;
    return typeNameOf(first) === 'org_docx4j_wml.P' ? new Paragraph(first as Element<wml.P>, container, this) : { element: first, container };
  }

  /** Inserts a WML fragment (one or more w:p / w:tbl / ... as inside document.xml); the standard prefixes are declared for it. */
  async insertXml(xml: string, location: 'Start' | 'End' | 'Before' | 'After', target?: Paragraph | BlockElement | Address): Promise<(Paragraph | BlockElement)[]> {
    this.assertWritable();
    return this.insertContent(await contentOf(xml, { preprocess: this.preprocessor() }), location, target, 'GeneralException');
  }

  /**
   * Word's `insertOoxml`: a flat OPC `pkg:package` string, whose body content is inserted with
   * the parts it references (images, embedded objects) copied into this package under fresh
   * names and relationship ids; styles and numbering are not merged (CR-002 open question 5).
   * A bare `w:p` / `w:tbl` fragment is accepted too, as `insertXml` takes. Returns the views of
   * what was inserted, as `insertXml` does, since a package may bring several blocks.
   */
  async insertOoxml(ooxml: string, location: 'Start' | 'End' | 'Replace' | 'Before' | 'After', target?: Paragraph | BlockElement | Address): Promise<(Paragraph | BlockElement)[]> {
    this.assertWritable();
    const elements = await contentOf(ooxml, { preprocess: this.preprocessor(), target: this.part });
    if (location === 'Replace') {
      assertControlHolds(this, this.container, [], elements, 'GeneralException');   // before the clear (CR-002 section 61)
      // In a cell, what ends with a table goes in before the cell's last paragraph, which stays, emptied, as Word
      // keeps it (check 52, case 94; CR-002 section 55.6); what ends with a paragraph replaces it, as before
      const cell = endsWithParagraph(elements) ? undefined : wrappingCellOf(this.container);
      this.clearContent(cell !== undefined);
      const end = cell ? paragraphEndOf(cell.content as Element[] | undefined) : undefined;
      if (end && end.container === this.content) return this.insertContent(elements, 'Before', { element: end.element as Element, container: end.container }, 'GeneralException');
      return this.insertContent(elements, 'End', undefined, 'GeneralException');
    }
    return this.insertContent(elements, location, target, 'GeneralException');
  }

  /**
   * After an insertion into a cell: the cell still ends with a paragraph, as Word reads it (CR-002 section 55). Where
   * a table went in after the paragraph the cell ended with, in the cell's own content, Word's form (check 52, cases 86
   * and 87, `keepCellMarkLast`): that paragraph split at its end, the cell's mark staying last after the table, the
   * first half's mark the inserted one under tracking. Anywhere else docx4j's repair, an empty paragraph after any wrapper.
   */
  private keepCellEnd(cell: wml.Tc, end: CellEnd | undefined, container: Element[]): void {
    const content = (cell.content ??= []) as Element[];
    if (endsWithParagraph(content)) return;
    if (end && end.container === content && container === content) {
      keepCellMarkLast(new Paragraph(end.element, content, this), content);
      return;
    }
    const appended = endWithParagraph(cell);
    if (appended) this.assignParaId(appended);
  }

  /** The package's DOM preprocessor (the MCE one), for fragments and incoming packages. */
  private preprocessor(): ((doc: Document) => void) | undefined {
    const preprocess = this.package_?.loadOptions.preprocessor;
    return preprocess ? (doc: Document) => preprocess(doc) : undefined;
  }

  private insertContent(elements: Element[], location: 'Start' | 'End' | 'Before' | 'After', target?: Paragraph | BlockElement | Address, refusal?: BlockRefusal): (Paragraph | BlockElement)[] {
    if (elements.length === 0) return [];
    this.insertElement(elements, location, target, refusal);
    return elements.map((el) => typeNameOf(el) === 'org_docx4j_wml.P'
      ? new Paragraph(el as Element<wml.P>, this.containerOf(el), this)
      : { element: el, container: this.containerOf(el) });
  }

  /** Matches over every paragraph, in document order. */
  search(text: string, options?: SearchOptions): Range[] {
    // Word's paragraph mark, ^p, is searched across the body's whole range (CR-014 phase 3); a wildcard query across its
    // content, which Word's wildcards read without the last paragraph's mark (check 59 case 5, a control's; K7)
    if (options?.matchWildcards && this.paragraphs.length > 0) return this.getRange('Content').search(text, options);
    if (!options?.matchWildcards && !options?.matchRegExp && paragraphMarkPieces(text) && this.paragraphs.length > 0) return this.getRange('Whole').search(text, options);
    return this.paragraphs.flatMap((p) => p.search(text, options));
  }

  /**
   * Replaces every match of `find` with `replace`, last match first so that the offsets of the
   * ones still to do stay valid (CR-002 section 3.7). Returns the number replaced; tracked when
   * the package's `changeTrackingMode` is on.
   */
  replaceText(find: string, replace: string, options?: SearchOptions): number {
    refuseMarkReplacement('Body', find, options);
    if (options?.matchWildcards) refuseWildcardMarkReplacement('Body', this.search(find, options));
    this.assertWritable();
    // every branch of a kept mc:AlternateContent, each paragraph searched on its own; counted where a
    // reader sees the match, as resolve-all counts the listed changes (CR-013 phase 3)
    const matches = this.paragraphsIn('every').flatMap((p) => p.searchMatches(find, options).map((m) => ({ ...m, unread: inUnreadBranch(p.p) })));
    writingEveryBranch(() => {
      for (let i = matches.length - 1; i >= 0; i--) {
        const { range, match } = matches[i]!;
        range.insertText(expandReplacement(replace, match, options), 'Replace');
      }
    });
    return matches.filter((m) => !m.unread).length;
  }

  /**
   * Removes everything (section properties stay). While the package tracks changes, every
   * paragraph is marked deleted and every row's `w:trPr` takes a `w:del` instead. Removed outright,
   * the relationships the content held (its pictures' images) are settled on save (CR-002 section 47).
   * In a table cell the paragraph the cell ends with stays, emptied, its properties kept, and tracked
   * only its text is deleted, as Word keeps a cell's last paragraph (check 52, cases 92 and 93; CR-002 section 55.6).
   */
  clear(): void {
    this.clearContent(true);
  }

  /**
   * `clear()`; with `keepCellEnd` false a cell's last paragraph goes as well, its mark marked deleted under
   * tracking, for a caller that puts its own paragraph in next (`insertText(…, 'Replace')`, `TableCell.value`).
   * @internal
   */
  clearContent(keepCellEnd: boolean): void {
    this.assertWritable();
    const tracker = this.changeTracker;
    if (!tracker) {
      const cell = keepCellEnd ? wrappingCellOf(this.container) : undefined;
      const end = cell ? paragraphEndOf(cell.content as Element[] | undefined) : undefined;
      const ids = new Set<string>();
      relationshipIdsIn(this.content, ids);
      queueRelationshipRemovals((this.package_ as { pendingRelationshipRemovals?: RelationshipRemovals } | undefined)?.pendingRelationshipRemovals, this.part, ids);
      this.content.length = 0;
      if (cell && !endsWithParagraph(cell.content as Element[] | undefined)) {
        if (end) {
          end.element.value.content = [];
          this.content.push(end.element as Element);
          linkParents(end.element, this.container);
        } else {
          const appended = endWithParagraph(cell);
          if (appended) this.assignParaId(appended);
        }
      }
      return;
    }
    // every branch of a kept mc:AlternateContent (CR-013 phase 3)
    writingEveryBranch(() => {
      for (const row of this.rowElements()) tracker.markRowDeleted(row.value as wml.Tr);
      for (const p of this.paragraphsIn('every')) p.remove(keepCellEnd);
    });
  }

  /**
   * Marks everything in this body deleted in place, in every branch of a kept `mc:AlternateContent`: each
   * paragraph's text and mark, each row of its tables (a tracked row or table deletion's cells; CR-013
   * phase 3). Internal: what `Table.delete` and `deleteRows` write while changes are tracked.
   */
  markContentDeleted(tracker: ChangeTracker): void {
    // the paragraphs, then the rows last first: the order, and so the ids, a row's deletion has always written;
    // a row already deleted is left as it is, as a paragraph already deleted is (a range's deletion, CR-014 phase 2)
    writingEveryBranch(() => {
      for (const p of this.paragraphsIn('every')) p.markDeletedInPlace(tracker);
      for (const row of this.rowElements().reverse()) if (!(row.value as wml.Tr).trPr?.del) tracker.markRowDeleted(row.value as wml.Tr);
    });
  }

  // --- change tracking (CR-002 phase F) ---------------------------------------------------

  /**
   * Every tracked change in this body, in document order. Touching insertions or deletions by one
   * author - a deleted paragraph's text and mark, a deletion across a mark - are one change, as
   * Office JS lists them (CR-002 section 29, checks 22 and 23).
   *
   * In a kept `mc:AlternateContent` at block, paragraph, row or cell level the selected branch's changes
   * are listed, as Word reads it (CR-013 phase 2); each other branch's are their twins, paired by kind and
   * position, which resolving a listed change resolves too (phase 3, docx4j's twins).
   */
  getTrackedChanges(): TrackedChange[] {
    return this.trackedChanges(listedListing());
  }

  /** The listing in either mode (`Listing`): resolve-all's puts every other branch's changes after the story's. */
  private trackedChanges(listing: Listing): TrackedChange[] {
    const out: TrackedChangeToken[] = [];
    const visit = (items: Element[], into: TrackedChangeToken[]): void => {
      for (let i = 0; i < items.length; i++) {
        const el = items[i]!;
        const tn = typeNameOf(el);
        if (tn === 'org_docx4j_wml.P') { into.push(...trackedChangeTokensOfParagraph(new Paragraph(el as Element<wml.P>, items, this), listing)); continue; }
        // a block control or custom XML element that is itself an insertion (check 30): one piece,
        // grouping with the insertions around and inside it, as Word lists them
        const inserted = holderChangeAt(items, i, undefined, this);
        if (inserted) {
          into.push(inserted);
          visit(childrenOf(items[i + 1]!.value as object) ?? [], into);
          i += 2;
          continue;
        }
        const v = el.value;
        if (typeof v !== 'object' || v === null) continue;
        if (tn === 'org_docx4j_wml.Tbl') {
          // A table is no break of its own: an inserted or deleted row groups with the rows and the
          // inserted text touching it, as Office JS lists them (check 31: a three-row table between two
          // inserted paragraphs is one change; rows added at a table's end with a paragraph typed after
          // it, one; a paragraph whose mark is plain before rows added at the start, two). Its property
          // records are one change of their own where the table stands (CR-002 section 35), kept apart.
          const tableChange = tablePropertiesChangeOf(el as Element<wml.Tbl>, this);
          if (tableChange) into.push(BREAK, tableChange, BREAK);
          visitRows(childrenOf(v) ?? [], { element: el as Element<wml.Tbl>, container: items }, into);
          continue;
        }
        // a kept mc:AlternateContent is its selected branch, as Word reads it, and keeps nothing apart
        // (CR-013 phase 2): Word's save has no element there. Its other branches' are twins (phase 3)
        if (tn === ALTERNATE_CONTENT) { alternateTokens(v as mce.AlternateContent, visit, into, listing); continue; }
        // a block-level content control or custom XML element keeps what is inside it apart; a
        // marker between paragraphs (a bookmark or range end) keeps nothing apart
        const children = childrenOf(v);
        if (children) { into.push(BREAK); visit(children, into); into.push(BREAK); }
      }
    };
    // a table's rows, through row-level controls and custom XML and a row-level mc:AlternateContent
    const visitRows = (items: Element[], table: Located<wml.Tbl>, into: TrackedChangeToken[]): void => {
      for (const el of items) {
        const tn = typeNameOf(el);
        const v = el.value;
        if (typeof v !== 'object' || v === null) continue;
        if (tn === 'org_docx4j_wml.Tr') {
          // A row that is itself a revision reports as one change, as Office JS does: the markup its
          // cells carry (Word marks every run and mark of a deleted or inserted row) belongs to that
          // change, which accepts or rejects it along with the row.
          const tokens: TrackedChangeToken[] = [];
          visitCells(childrenOf(v) ?? [], el as Element<wml.Tr>, tokens);
          const inner = groupTouching(tokens);
          const rowChanges = trackedChangesOfRow(el as Element<wml.Tr>, items, inner, table, this);
          if (rowChanges.length > 0) into.push(...rowChanges);
          else { into.push(BREAK); for (const change of inner) into.push(change, BREAK); }
        } else if (tn === ALTERNATE_CONTENT) {
          alternateTokens(v as mce.AlternateContent, (branch, tokens) => visitRows(branch, table, tokens), into, listing);
        } else if (tn !== 'org_docx4j_wml.Tc') {
          const children = childrenOf(v);
          if (children) visitRows(children, table, into);
        }
      }
    };
    // a row's cells, through cell-level controls and custom XML and a cell-level mc:AlternateContent
    const visitCells = (items: Element[], row: Element<wml.Tr>, tokens: TrackedChangeToken[]): void => {
      for (const el of items) {
        const tn = typeNameOf(el);
        const v = el.value;
        if (typeof v !== 'object' || v === null) continue;
        if (tn === 'org_docx4j_wml.Tc') {
          const cellChange = cellChangeOf(el as Element<wml.Tc>, items, row.value, this);
          if (cellChange) tokens.push(cellChange, BREAK);
          visit(childrenOf(v) ?? [], tokens);
          tokens.push(BREAK);
        } else if (tn === ALTERNATE_CONTENT) {
          alternateTokens(v as mce.AlternateContent, (branch, into) => visitCells(branch, row, into), tokens, listing);
        } else {
          const children = childrenOf(v);
          if (children) visitCells(children, row, tokens);
        }
      }
    };
    visit(this.content, out);
    // the body's last section, after everything in it
    const section = bodySectionChangeOf(this.container as { content?: Element[]; sectPr?: wml.SectPr }, this);
    if (section) out.push(BREAK, section);
    // resolve-all's: every other branch's changes, each kept apart, after the story's (resolved first, last first)
    if (listing.extra.length > 0) out.push(BREAK, ...listing.extra);
    return groupTouching(out);
  }

  /**
   * Accepts every tracked change, last first so that a paragraph join never disturbs one still
   * to do (docx4j `AcceptTrackedChanges` is the reference). Returns how many were listed.
   *
   * With a filter (an extension, CR-011), only the single changes it matches are accepted, so a group
   * can be resolved in part, and what no change owns is left alone; the count is of the listed changes
   * with a matching piece (docx4j's `TrackedChanges.acceptAll(pkg, filter)`).
   */
  acceptAll(filter?: TrackedChangeSelection): number {
    return this.resolveAll('accept', filter);
  }

  /** Rejects every tracked change, last first, as `acceptAll` accepts them. Returns how many were listed. */
  rejectAll(filter?: TrackedChangeSelection): number {
    return this.resolveAll('reject', filter);
  }

  /**
   * Resolving one change can leave another to be listed afresh: accepting a move carries a deletion
   * nested in its moved text out of it, and resolving a broken move dissolves the rest of it into a
   * new `w:ins` or `w:del` (CR-002 section 29), so the list read first goes stale. The list is read
   * again after each pass until nothing is left, as "all" promises (found by the editor's C5); as
   * docx4j's `TrackedChanges.resolveAll` does it (CR-011): a change whose pieces all match is resolved
   * whole, one with some matching its matching pieces alone, for up to 20 passes, and two passes in a
   * row that leave as many matching changes as the one before end it. Then, with no filter, what is
   * never listed is swept.
   */
  private resolveAll(what: 'accept' | 'reject', filter: TrackedChangeSelection | undefined): number {
    const matches = filter === undefined ? undefined : pieceMatcher(filter);
    const matching = (change: TrackedChange): readonly TrackedChange[] => (matches ? change.pieces.filter(matches) : change.pieces);
    let count = 0;
    // one resolution for the lot, so the bookmarks are read before the first and after the last
    resolving(this.package_, this, () => {
      // counted as listed; resolved over every branch of a kept mc:AlternateContent, each other branch's
      // changes on their own, as docx4j's resolveAll lists them (CR-013 phase 3)
      count = this.getTrackedChanges().filter((change) => matching(change).length > 0).length;
      let left = Infinity;
      let stalled = 0;
      for (let pass = 0; pass < 20 && stalled < 2; pass++) {
        const changes = this.trackedChanges({ mode: 'every', extra: [] });
        let matched = 0;
        for (let i = changes.length - 1; i >= 0; i--) {
          const change = changes[i]!;
          const pieces = matching(change);
          if (pieces.length === 0) continue;
          matched++;
          if (pieces.length === change.pieces.length) change[what](); else resolvePieces(pieces, what);
        }
        if (matched === 0) break;
        stalled = matched >= left ? stalled + 1 : 0;
        left = matched;
      }
      // and what is never listed: a cell merge, a numbering change, records that change nothing
      if (!matches) resolveUnlisted(this.content, what);
    });
    return count;
  }

  /** Every `w:tr` in this body's tables, in document order (rows of nested tables included), in every branch (CR-013 phase 3). */
  private rowElements(): Element<wml.Tr>[] {
    const out: Element<wml.Tr>[] = [];
    const visit = (items: Element[]): void => {
      for (const el of items) {
        const v = el.value;
        if (typeof v !== 'object' || v === null) continue;
        if (typeNameOf(el) === 'org_docx4j_wml.Tbl') {
          for (const row of rowsOf(v, 'every')) {
            out.push(row.element);
            for (const cell of cellsOf(row.element.value, 'every')) visit(childrenOf(cell.element.value) ?? []);
          }
          continue;
        }
        if (typeNameOf(el) === ALTERNATE_CONTENT) { for (const branch of branchesOf(v as mce.AlternateContent)) visit(branch); continue; }
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

  /**
   * The element at an address: 'body/3' (fourth block), 'body/3/1/0/2' (table row 1, cell 0, block 2),
   * 'body/3/textbox:0/1' (the second block of the first text box body/3 anchors, CR-012 D4), a paraId, or a
   * text match. Undefined, never an error, where the address names nothing.
   */
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
      if (seg.startsWith('textbox:')) {
        // the n-th text box the paragraph here anchors, as the text model numbers them (CR-012 D4)
        if (typeNameOf(element ?? ({} as Element)) !== 'org_docx4j_wml.P') return undefined;
        const n = Number(seg.substring('textbox:'.length));
        const box = Number.isInteger(n) && n >= 0 ? textBoxesOf(current as wml.P, this.hiddenRuns(current as wml.P))[n] : undefined;
        if (!box?.content) return undefined;                          // none, or legacy WordArt: no paragraphs
        current = box.content;
        element = undefined;
        container = undefined;
        continue;
      }
      container = childrenOf(current, 'read');
      const i = Number(seg);
      if (!container || seg === '' || !Number.isInteger(i) || i < 0 || i >= container.length) return undefined;
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

  /**
   * What an agent asks for first: every paragraph and table with its address, style and text, read by the
   * text model (CR-012 D3; `Texts.of`'s options, `TextOptionsInit`): each paragraph's text is the model's
   * (hidden text left out, a paragraph whose mark the view deletes joined to the next, in the view asked
   * for), and each text box's paragraphs follow their anchoring paragraph, or the table it is in (REQ-518),
   * at `textbox:` addresses (`body/3/textbox:0/1`), flagged `textbox`. A cell's paragraphs and tables follow
   * their table's entry.
   * Async since phase 4, as the model's reading is; calls under one `TextContext` (`Texts.contextOf`) read
   * the package's stories once. The package is not changed (REQ-204).
   *
   * With `select`, only some of this body's top-level blocks are read (CR-012 section 17.13): their entries
   * are those the whole outline gives them under the same context and options, in document order. A
   * paragraph whose text a paragraph before it holds (`joined`) gives none; a block that starts inside a
   * field open before it (a table of contents left out, D12) is read inside it. Under a context, the first
   * such call over a body walks its blocks' field characters once; each block after that costs its own
   * reading. A context is a snapshot: after an edit, make another.
   */
  async outline(options?: object, context?: object, select?: OutlineSelection): Promise<Outline> {
    return outlineReader()(this, options, context, select?.blocks.map((b) => this.blockIndexOf(b)));
  }

  /** The index in `content` of one of this body's top-level blocks. */
  private blockIndexOf(block: number | Paragraph | BlockElement): number {
    const content = this.content;
    if (typeof block === 'number') {
      if (!Number.isInteger(block) || block < 0 || block >= content.length) throw new Docx4JException(`${this.prefix} has no block ${block}`);
      return block;
    }
    const i = block.container === content ? content.indexOf(block.element) : -1;
    if (i < 0) throw new Docx4JException(`Not one of ${this.prefix}'s top-level blocks: ${describe(block.element)}`);
    return i;
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

  private pathOf(value: object): (number | string)[] | undefined {
    const path: (number | string)[] = [];
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
      if (BRANCH_TYPES.has((parent as { TYPE_NAME?: string }).TYPE_NAME ?? '')) {
        // a kept mc:AlternateContent's branch is no level of its own (D2): what is in the selected one is
        // addressed from the element, body/2/0; another branch's is not addressed, readers not seeing it
        const ac = (parent as { PARENT?: object }).PARENT as mce.AlternateContent | undefined;
        if (!ac || selectedBranchOf(ac) !== children) return undefined;
        child = ac;
        continue;
      }
      if ((parent as { TYPE_NAME?: string }).TYPE_NAME === 'org_docx4j_wml.CTTxbxContent') {
        // a text box's block: addressed from the paragraph that anchors it, `textbox:n` its place among
        // that paragraph's boxes; none for a box the text model doesn't number (another branch's, a hidden run's)
        const anchor = anchorOf(parent);
        const n = anchor ? textBoxesOf(anchor, this.hiddenRuns(anchor)).findIndex((box) => box.content === parent) : -1;
        if (!anchor || n < 0) return undefined;
        path.unshift(`textbox:${n}`);
        child = anchor;
        continue;
      }
      child = parent;
    }
    return undefined;
  }

  /**
   * Gives a new paragraph a w14:paraId when the document already uses them, free in every branch (CR-013 phase 3).
   * @internal a paragraph split (CR-014 phase 2) gives its new first half one.
   */
  assignParaId(p: wml.P): void {
    if (p.paraId) return;
    const paragraphs = this.paragraphsIn('every');
    if (!paragraphs.some((q) => q.paraId !== undefined)) return;
    let id: string;
    do {
      id = Math.floor(Math.random() * 0x7fffffff).toString(16).toUpperCase().padStart(8, '0');
    } while (paragraphs.some((q) => q.paraId === id));
    p.paraId = id;
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
    this.assertWritable();
    const content = [...this.content];
    const sdt = sdtOf(content, { kind: sdtKindFor(kind), id: nextSdtId(controlIdScope(this)), form: 'block' });
    this.content.length = 0;
    this.content.push(sdt as Element);
    linkParents(sdt, this.container);
    declareSdtExtensions(this, sdt);
    // under tracking the control itself is the insertion, its content what it was (check 30)
    const tracker = this.changeTracker;
    if (tracker) markHolderInserted(tracker, sdt as Element, this.content, () => {});
    return new ContentControl(sdt as Element<wml.SdtBlock>, this.content, this);
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

/**
 * A story's tracked changes resolved (`Body.acceptAll` or `rejectAll`, with the filter if one is given),
 * then those of each text box in it - a `w:txbxContent`, in a DrawingML shape or a VML one - which Word
 * counts as a story of its own (CR-002 section 49). A text box inside a text box is reached through its
 * own story. Returns how many changes were listed (with a filter, how many had a matching piece).
 *
 * In a kept `mc:AlternateContent` the text box of every branch is resolved, so that each says the same
 * afterwards (docx4j's `McMode.ALL`), and the changes are counted in the selected branch alone, so a
 * text box held in both counts once, as a resolved load's does (CR-010 phase 3).
 */
export function resolveStory(body: Body, what: 'accept' | 'reject', filter?: TrackedChangeSelection): number {
  let count = what === 'accept' ? body.acceptAll(filter) : body.rejectAll(filter);
  for (const { box, selected } of textBoxesIn(body.container)) {
    const resolved = resolveStory(body.sub(box), what, filter);
    if (selected) count += resolved;
  }
  return count;
}

/**
 * A story's tracked changes (`Body.getTrackedChanges`), then each text box's in it, as `resolveStory`
 * resolves them: docx4j's order (its CR-043 `Lister`, CR-011). In a kept `mc:AlternateContent` the
 * selected branch's text boxes alone are listed, so a text box held in both branches is listed once;
 * each other branch's box at the same place is its twin, whose changes are paired with the listed ones
 * by kind and position and resolved with them, as docx4j resolves a text box's `mc:Fallback` twin
 * (`pairTwins`; CR-013 phase 3).
 */
export function trackedChangesOfStory(body: Body): TrackedChange[] {
  const out = body.getTrackedChanges();
  for (const { box, selected, twins } of textBoxesIn(body.container)) {
    if (!selected) continue;
    const changes = trackedChangesOfStory(body.sub(box));
    for (const twin of twins ?? []) pairTwins(changes, trackedChangesOfStory(body.sub(twin)));
    out.push(...changes);
  }
  return out;
}

/**
 * The members `textBoxesIn` does not descend into: an element's name, and properties, which hold no text box.
 * Over every fixture the walk finds the same text boxes without them and visits 38% fewer objects (CR-012
 * section 17.13: it was most of a plan's cost).
 */
const NO_TEXT_BOXES = new Set(['PARENT', 'name', 'pPr', 'rPr', 'tblPr', 'tblGrid', 'trPr', 'tcPr', 'sectPr', 'sdtPr', 'sdtEndPr']);

/** A text box `textBoxesIn` found: its `w:txbxContent`, whether a reader takes it, and its twins. */
interface FoundTextBox {
  box: { content?: Element[] };
  /** False inside a branch of a kept `mc:AlternateContent` other than its selected one. */
  selected: boolean;
  /** For a selected box in a kept `mc:AlternateContent`'s branch: the box at the same place in each other branch. */
  twins?: { content?: Element[] }[];
}

/**
 * The `w:txbxContent` of every text box in a story's content, not descending into one found, each with
 * whether it is in a branch a reader takes: false inside a branch of a kept `mc:AlternateContent` other
 * than its selected one (`selectedBranchOf`). A selected box in a branch has as twins the boxes at its
 * place, by position, in the element's other branches (docx4j's `Lister.twins`).
 */
function textBoxesIn(container: { content?: Element[] }): FoundTextBox[] {
  const out: FoundTextBox[] = [];
  const seen = new Set<object>();
  const visit = (o: unknown, selected: boolean, into: FoundTextBox[]): void => {
    if (typeof o !== 'object' || o === null || seen.has(o)) return;
    if (typeof (o as { nodeType?: unknown }).nodeType === 'number') return;      // DOM the model does not type
    seen.add(o);
    if (Array.isArray(o)) { for (const item of o) visit(item, selected, into); return; }
    const typeName = (o as { TYPE_NAME?: string }).TYPE_NAME;
    if (typeName === 'org_docx4j_wml.CTTxbxContent') { into.push({ box: o as { content?: Element[] }, selected }); return; }
    if (typeName === ALTERNATE_CONTENT) {
      const ac = o as mce.AlternateContent;
      const read = selectedBranchOf(ac);
      let primary: FoundTextBox[] = [];
      const others: FoundTextBox[][] = [];
      for (const branch of [...(ac.choice ?? []), ...(ac.fallback ? [ac.fallback] : [])]) {
        const found: FoundTextBox[] = [];
        visit(branch.any, selected && branch.any === read, found);
        into.push(...found);
        if (branch.any === read) primary = found; else others.push(found);
      }
      for (const found of others) found.forEach(({ box }, i) => { const twin = primary[i]; if (twin) (twin.twins ??= []).push(box); });
      return;
    }
    for (const key in o) if (!NO_TEXT_BOXES.has(key)) visit((o as Record<string, unknown>)[key], selected, into);
  };
  visit(container.content, true, out);
  return out;
}

/** The paragraph a text box's content is anchored in: its nearest `w:p` ancestor. */
function anchorOf(box: object): wml.P | undefined {
  let o = (box as { PARENT?: object }).PARENT;
  for (let guard = 0; o && guard < 64; guard++, o = (o as { PARENT?: object }).PARENT) {
    if ((o as { TYPE_NAME?: string }).TYPE_NAME === 'org_docx4j_wml.P') return o as wml.P;
  }
  return undefined;
}
