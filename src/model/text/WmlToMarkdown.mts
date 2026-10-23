// CR-012 phase 3: the block model written as a Markdown tree, docx4j-markdown's org.docx4j.markdown.WmlToMarkdown
// (VERSION_17_3_4, unchanged since 198d015a3), which commonmark.mts then serializes as commonmark-java does.
import type { PropertyResolver } from '../properties/PropertyResolver.mjs';
import { log } from '../properties/log.mjs';
import { LinkTargets } from './LinkTargets.mjs';
import type { TextOptions } from './TextOptions.mjs';
import type { TextResult, TextBlock, TextRun, TextRevision, TextNumbering, TextRow, TextCell, TextComment, NoteType } from './TextResult.mjs';
import {
  MdNode, MdDocument, MdParagraph, MdHeading, MdBlockQuote, MdListBlock, MdBulletList, MdOrderedList, MdListItem,
  MdFencedCodeBlock, MdThematicBreak, MdHtmlInline, MdText, MdCode, MdEmphasis, MdStrongEmphasis, MdLink, MdImage,
  MdHardLineBreak, MdTableBlock, MdTableHead, MdTableBody, MdTableRow, MdTableCell, MdStrikethrough,
  MdFootnoteDefinition, MdFootnoteReference, MdInlineMath, MdDisplayMath, type MdAlignment,
} from './commonmark.mjs';

const TWIPS_PER_LEVEL = 720;

/** The paragraph styles docx4j's importer writes, which the export reads back (docx4j `ImportStyles`). */
const QUOTE = 'Quote';
const INTENSE_QUOTE = 'IntenseQuote';
const LIST_PARAGRAPH = 'ListParagraph';
const SOURCE_CODE = 'SourceCode';

/** Paragraph styles which are code blocks by name, whatever their font. */
const CODE_STYLES = new Set([SOURCE_CODE, 'Code', 'HTMLPreformatted', 'MacroText']);

/** An open list: it spans consecutive numbered paragraphs. */
interface OpenList {
  readonly node: MdListBlock;
  /** The list: its abstractNum where the text model gives one, else its numId. */
  readonly numId: string;
  lastItem: MdListItem | undefined;
  /** The last item's left indent, where its text starts (twips). */
  itemIndent: number;
  /** The count the next item must show for this list to go on, else a new list starts at it; undefined for a bullet list. */
  nextValue: number | undefined;
}

/** An image node waiting for its destination, which needs its part's bytes (`MarkdownExporter`). */
export interface PendingImage {
  readonly node: MdImage;
  readonly run: TextRun;
}

/**
 * Writes the text model's block model (`TextResult`) as a Markdown tree (WD REQ-703): a heading from its
 * outline level, not its style's name; a list from its numbering; bold, italic, struck-through and code text
 * from what a run adds to its paragraph's style (so a heading's inherent bold doesn't become `**markers**`);
 * links from a `w:hyperlink` or a HYPERLINK field; images with the describer's text; notes as `[^n]`; in the
 * `markup` view, CriticMarkup (`{++ ++}`, `{-- --}`, `{>> <<}`). A table of contents is the text model's to
 * leave out (`toc: false`, the default). Each node made for a block carries it as its `tag`, so that
 * `MarkdownExporter.blocks` can say where each block's Markdown begins. docx4j `WmlToMarkdown`.
 */
export class WmlToMarkdown {
  private readonly textOptions: TextOptions;
  /** The open lists, the deepest last. */
  private readonly openLists: OpenList[] = [];
  /** The count each list level last showed (`numId/ilvl`), for a parent item reopened after an interruption. */
  private readonly lastValues = new Map<string, number>();
  // the open quote: it spans consecutive Quote paragraphs
  private openQuote: MdBlockQuote | undefined;
  private openQuoteDepth = 0;
  /** Above 0 while building a GFM header row, whose bold is convention, not markup. */
  private tableHeaderDepth = 0;
  private readonly styleIndents = new Map<string, number>();
  /** The images, in the order written, for their destinations. */
  readonly images: PendingImage[] = [];
  /** Whether a link target was refused (`LinkTargets`): logged once per document. */
  private linkRefused = false;

  constructor(
    private readonly result: TextResult,
    /** For a style's own indent (a block quote's depth); undefined where there are no styles to ask. */
    private readonly resolver: PropertyResolver | undefined,
    /** The schemes a link may use beyond `LinkTargets`' defaults (`MarkdownExportOptions.linkSchemes`). */
    private readonly linkSchemes: readonly string[] = [],
  ) {
    this.textOptions = result.options;
  }

  convert(): MdDocument {
    const document = new MdDocument();
    const blocks = this.result.blocks;
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i]!;
      if (isNote(block)) {
        // a note's blocks are its definition, at the place the text model reads it (after the main story)
        this.closeLists();
        this.closeQuote();
        let last = i;
        while (last + 1 < blocks.length && sameNote(block, blocks[last + 1]!)) last++;
        const definition = tagged(new MdFootnoteDefinition(noteLabel(block.noteType!, block.noteNumber, block.noteId)), block);
        this.body(blocks.slice(i, last + 1), definition);
        this.closeLists();
        this.closeQuote();
        trimLeadingMarkerSpace(definition);
        document.appendChild(definition);
        i = last;
      } else {
        i = this.block(blocks, i, document);
      }
    }
    this.closeLists();
    this.closeQuote();
    return document;
  }

  // ---------------------------------------------------------------- blocks

  private body(blocks: readonly TextBlock[], container: MdNode): void {
    for (let i = 0; i < blocks.length; i++) i = this.block(blocks, i, container);
  }

  /** Returns the index of the last block consumed (a code block spans several). */
  private block(blocks: readonly TextBlock[], index: number, container: MdNode): number {
    const block = blocks[index]!;
    switch (block.kind) {
      case 'TABLE':
        this.closeLists();
        this.closeQuote();
        this.table(block, container);
        return index;
      case 'CHART':
        this.closeLists();
        this.closeQuote();
        this.chart(block, container);
        return index;
      default:
        return this.paragraph(blocks, index, container);
    }
  }

  private paragraph(blocks: readonly TextBlock[], index: number, container: MdNode): number {
    const block = blocks[index]!;
    const style = block.style;

    // a thematic break: an empty paragraph carrying a bottom border of its own
    if (block.bottomBorder && isEmpty(block)) {
      this.closeLists();
      this.closeQuote();
      container.appendChild(tagged(new MdThematicBreak('---'), block));   // else the renderer picks its own idiom
      return index;
    }

    // display math: a paragraph which is one m:oMathPara
    const displayMath = soleDisplayMath(block);
    if (displayMath !== undefined) {
      this.closeLists();
      this.closeQuote();
      container.appendChild(tagged(new MdDisplayMath(latex(displayMath)), block));
      return index;
    }

    // a code block: a code-styled paragraph, or one whose paragraph style's own font is mono; consecutive ones
    // merged into one fence, one line each
    if (isCodeParagraph(block)) {
      this.closeLists();
      this.closeQuote();
      let literal = '';
      let last = index;
      for (let j = index; j < blocks.length; j++) {
        const b = blocks[j]!;
        if (b.kind !== 'PARAGRAPH' || !isCodeParagraph(b) || !sameStory(block, b)) break;
        literal += `${this.codeText(b)}\n`;
        last = j;
      }
      container.appendChild(tagged(new MdFencedCodeBlock(literal), block));
      return last;
    }

    // a heading by its effective outlineLvl (0 to 5 are # to ######); before the numbering, since the built-in
    // Heading styles can carry a legacy numPr
    const outlineLevel = block.outlineLevel;
    if (outlineLevel !== undefined && outlineLevel <= 5) {
      this.closeLists();
      this.closeQuote();
      const heading = tagged(new MdHeading(outlineLevel + 1), block);
      this.inlines(block, heading, true);
      container.appendChild(heading);
      return index;
    }

    // a list paragraph
    const numbering = block.numbering;
    if (numbering !== undefined) {
      this.closeQuote();
      if (numbering.label) {
        // a level with numFmt none, or a lvlText with no %n, is a label ("NOTE", "EXAMPLE"), not a list: it
        // leads the paragraph in bold
        this.closeLists();
        const label = numbering.levelText === undefined ? '' : numbering.levelText.trim();
        const labelled = this.inlineParagraph(block, false);
        if (label.length > 0) {
          if (labelled.firstChild !== undefined) labelled.prependChild(new MdText(' '));
          const strong = new MdStrongEmphasis();
          strong.appendChild(new MdText(label));
          labelled.prependChild(strong);
        }
        if (labelled.firstChild !== undefined) container.appendChild(labelled);
        return index;
      }
      this.listParagraph(block, numbering, container);
      return index;
    }

    // an unlabelled paragraph inside a list (WD REQ-703): Markdown has no empty paragraphs, so an empty one
    // neither ends the list nor joins it; one indented at least to an open item's text (or in the
    // ListParagraph style) is a follow-on paragraph of the deepest such item; any other ends the lists, and
    // the list goes on at its own count where it resumes (listParagraph)
    if (this.openLists.length > 0 && style !== QUOTE && style !== INTENSE_QUOTE) {
      if (isEmpty(block)) return index;
      const into = this.followOnItem(block, style);
      if (into !== undefined) {
        into.node.tight = false;
        into.lastItem!.appendChild(this.inlineParagraph(block, true));
        return index;
      }
    }

    this.closeLists();

    // a block quote, nested by the indent beyond the quote style's own
    if (style === QUOTE || style === INTENSE_QUOTE) {
      const beyond = value(block.indentLeft) - this.styleIndentLeft(style);
      const depth = 1 + Math.trunc(Math.max(0, beyond) / TWIPS_PER_LEVEL);
      this.appendToQuote(depth, this.inlineParagraph(block, true), container, block);
      return index;
    }

    this.closeQuote();

    const paragraph = this.inlineParagraph(block, true);
    if (paragraph.firstChild !== undefined) container.appendChild(paragraph);   // Markdown has no empty paragraphs
    return index;
  }

  private inlineParagraph(block: TextBlock, withListLabel: boolean): MdParagraph {
    const paragraph = tagged(new MdParagraph(), block);
    this.inlines(block, paragraph, withListLabel);
    return paragraph;
  }

  /** A code paragraph's text, unescaped; in `markup`, its revisions marked. */
  private codeText(block: TextBlock): string {
    let out = '';
    const runs = block.runs;
    let i = 0;
    while (i < runs.length) {
      const revision = runs[i]!.revision;
      let j = i;
      while (j < runs.length && this.sameMarker(revision, runs[j]!.revision)) j++;
      if (revision !== undefined) out += this.open(revision);
      for (let k = i; k < j; k++) {
        const r = runs[k]!;
        if (r.kind === 'TEXT' || r.kind === 'FIELD_CODE') out += r.text;
      }
      if (revision !== undefined) out += close(revision);
      i = j;
    }
    return out;
  }

  /** A paragraph style's own left indent, in twips: 0 where it has none, or there are no styles to ask. */
  private styleIndentLeft(styleId: string): number {
    const cached = this.styleIndents.get(styleId);
    if (cached !== undefined) return cached;
    let left = 0;
    if (this.resolver) {
      try {
        const ind = this.resolver.getEffectivePPr(styleId).ind;
        const l = ind?.left ?? ind?.start;
        if (l !== undefined && l !== null && Number.isFinite(Number(l))) left = Math.trunc(Number(l));
      } catch {
        // no indent for the style
      }
    }
    this.styleIndents.set(styleId, left);
    return left;
  }

  // ---------------------------------------------------------------- quotes

  private appendToQuote(depth: number, paragraph: MdParagraph, container: MdNode, block: TextBlock): void {
    if (this.openQuote === undefined || depth !== this.openQuoteDepth) {
      this.closeQuote();
      const outer = tagged(new MdBlockQuote(), block);
      let inner = outer;
      for (let d = 1; d < depth; d++) {
        const next = new MdBlockQuote();
        inner.appendChild(next);
        inner = next;
      }
      container.appendChild(outer);
      this.openQuote = inner;
      this.openQuoteDepth = depth;
    }
    this.openQuote.appendChild(paragraph);
  }

  private closeQuote(): void {
    this.openQuote = undefined;
    this.openQuoteDepth = 0;
  }

  // ---------------------------------------------------------------- lists

  /**
   * The open list whose last item an unlabelled paragraph continues: the deepest whose item's text the
   * paragraph is indented to (its left indent at least the item's, and not at the margin), or the deepest
   * open list for a ListParagraph-styled paragraph; the lists deeper than it are closed. Undefined where
   * none: the paragraph ends the lists.
   */
  private followOnItem(block: TextBlock, style: string | undefined): OpenList | undefined {
    const indent = value(block.indentLeft);
    const listParagraph = style === LIST_PARAGRAPH;
    for (let k = this.openLists.length - 1; k >= 0; k--) {                // deepest first
      const open = this.openLists[k]!;
      if (open.lastItem !== undefined && (listParagraph || (indent > 0 && indent >= open.itemIndent))) {
        this.openLists.length = k + 1;
        return open;
      }
    }
    return undefined;
  }

  private top(): OpenList {
    return this.openLists[this.openLists.length - 1]!;
  }

  private listParagraph(block: TextBlock, numbering: TextNumbering, container: MdNode): void {
    const ilvl = numbering.level;
    // the list: two w:num on one w:abstractNum are one list, counting on from each other
    const numId = numbering.abstractNumId ?? numbering.numId;
    const count = numbering.value;
    const bullet = numbering.bullet || numbering.format === undefined;

    // close any deeper (or same-level-but-different-list) contexts
    if (this.openLists.length > ilvl + 1) this.openLists.length = ilvl + 1;
    if (this.openLists.length === ilvl + 1 && this.top().numId !== numId && ilvl === 0) {
      this.openLists.pop();                                              // a different top-level list: restart
    }
    if (this.openLists.length === ilvl + 1 && (this.top().node instanceof MdBulletList) !== bullet) {
      this.openLists.pop();                                              // a level drawn as the other kind (a w:num overriding it)
    }

    const tight = block.contextualSpacing !== false;

    // the list's items show consecutive counts from its start; an item whose count is not the next one (the
    // list resumed after a paragraph that ended it, a w:num sharing the count, a start override) begins a new
    // list at its count
    if (this.openLists.length === ilvl + 1 && count !== undefined && this.top().nextValue !== undefined
        && count !== this.top().nextValue) {
      this.openLists.pop();
    }

    // open the missing levels (a level skipped over takes the paragraph's kind of list); a parent level
    // reopened after an interruption is an empty item showing the parent's last count
    while (this.openLists.length < ilvl + 1) {
      const level = this.openLists.length;
      const listBlock = tagged(createListBlock(numbering, level === ilvl, tight,
        level === ilvl ? count : this.lastValues.get(`${numId}/${level}`)), block);
      if (this.openLists.length === 0) {
        appendList(container, listBlock);
      } else {
        const parent = this.top();
        if (parent.lastItem === undefined) {
          parent.lastItem = tagged(new MdListItem(), block);
          parent.node.appendChild(parent.lastItem);
        }
        appendList(parent.lastItem, listBlock);
      }
      this.openLists.push({ node: listBlock, numId, lastItem: undefined, itemIndent: 0, nextValue: undefined });
    }

    const target = this.top();
    if (!tight) target.node.tight = false;
    const item = tagged(new MdListItem(), block);
    item.appendChild(this.inlineParagraph(block, false));
    target.node.appendChild(item);
    target.lastItem = item;
    target.itemIndent = value(block.indentLeft);
    if (target.node instanceof MdOrderedList && count !== undefined) {
      target.nextValue = count + 1;
      this.lastValues.set(`${numId}/${ilvl}`, count);
    }
  }

  private closeLists(): void {
    this.openLists.length = 0;
  }

  // ---------------------------------------------------------------- tables

  /**
   * GFM pipes: the header row is the table's `w:tblHeader` rows, else its first (several `w:tblHeader` rows
   * are one header, their cells' texts ` / ` apart, as plain text's listed format joins them); block content
   * in cells flattens to inline with `<br>`, a nested table's paragraphs included; `w:vMerge` and
   * `w:gridSpan` degrade (top-left wins, empty cells elsewhere), GFM having no spans.
   *
   * The table style's formatting is left out on purpose (docx4j CR-030 decision 4): a header row the style
   * makes bold is not written `**bold**` (its being the GFM header says as much), and GFM has no table
   * styling for the rest.
   */
  private table(table: TextBlock, container: MdNode): void {
    const tableBlock = tagged(new MdTableBlock(), table);
    const head = new MdTableHead();
    const body = new MdTableBody();
    tableBlock.appendChild(head);

    const rows = table.rows;
    let headerRows = 0;
    while (headerRows < rows.length && rows[headerRows]!.header) headerRows++;
    if (headerRows === 0 && rows.length > 0) headerRows = 1;

    if (headerRows > 0) head.appendChild(this.headerRow(rows.slice(0, headerRows)));
    for (let r = headerRows; r < rows.length; r++) {
      const row = new MdTableRow();
      for (const cell of rows[r]!.cells) {
        row.appendChild(this.tableCell(cell, false));
        // gridSpan: empty cells after it, so that the columns stay aligned
        for (let s = 1; s < cell.gridSpan; s++) row.appendChild(new MdTableCell(false));
      }
      body.appendChild(row);
    }
    if (body.firstChild !== undefined) tableBlock.appendChild(body);
    container.appendChild(tableBlock);
  }

  /** The header: one row's cells, or several rows' cells by grid column, ` / ` apart. */
  private headerRow(rows: readonly TextRow[]): MdTableRow {
    const row = new MdTableRow();
    if (rows.length === 1) {
      for (const cell of rows[0]!.cells) {
        row.appendChild(this.tableCell(cell, true));
        for (let s = 1; s < cell.gridSpan; s++) row.appendChild(new MdTableCell(true));
      }
      return row;
    }
    // each grid column's cells, from the top: a cell is in the column it starts in
    let width = 0;
    for (const r of rows) {
      let w = 0;
      for (const c of r.cells) w += Math.max(1, c.gridSpan);
      width = Math.max(width, w);
    }
    const columns: TextCell[][] = [];
    for (let g = 0; g < width; g++) columns.push([]);
    for (const r of rows) {
      let g = 0;
      for (const c of r.cells) {
        if (!c.mergeContinue) columns[g]!.push(c);
        g += Math.max(1, c.gridSpan);
      }
    }
    for (const column of columns) {
      const cell = new MdTableCell(true);
      // the column's alignment is its topmost header cell's, as with one header row
      let alignment: MdAlignment | undefined;
      if (column.length > 0) {
        const first = paragraphsOf(column[0]!);
        if (first.length > 0) alignment = cellAlignment(first[0]!);
      }
      const sink = new InlineSink(cell);
      this.tableHeaderDepth++;
      try {
        let first = true;
        for (const c of column) {
          if (isEmptyCell(c)) continue;
          if (!first) sink.text(' / ', false, false, false);
          first = false;
          this.cellContent(c, cell, sink);
        }
      } finally {
        this.tableHeaderDepth--;
      }
      sink.flush();
      cell.alignment = alignment;
      row.appendChild(cell);
    }
    return row;
  }

  private tableCell(textCell: TextCell, header: boolean): MdTableCell {
    const cell = new MdTableCell(header);
    // a vMerge continuation: the merge's top cell already holds the content
    if (textCell.mergeContinue) return cell;
    const sink = new InlineSink(cell);
    if (header) this.tableHeaderDepth++;
    try {
      this.cellContent(textCell, cell, sink);
    } finally {
      if (header) this.tableHeaderDepth--;
    }
    sink.flush();
    return cell;
  }

  /**
   * A cell's paragraphs, a nested table's included, `<br>` apart (GFM cells are single-line); its empty
   * paragraphs left out, as plain text leaves them out (Word ends every cell with a paragraph, so a nested
   * table or a block-level control comes with an empty one after it).
   *
   * A cell's alignment is its first paragraph's effective `w:jc`, empty or not; GFM aligns a column, from
   * its header cell's (the delimiter row), so a body cell's is not seen.
   */
  private cellContent(textCell: TextCell, cell: MdTableCell, sink: InlineSink): void {
    let first = true;
    const paragraphs = paragraphsOf(textCell);
    if (cell.alignment === undefined && paragraphs.length > 0) cell.alignment = cellAlignment(paragraphs[0]!);
    for (const b of paragraphs) {
      if (isEmpty(b)) continue;
      if (!first) sink.node(new MdHtmlInline('<br>'));
      first = false;
      this.runs(b, sink);
    }
  }

  // ---------------------------------------------------------------- charts

  /**
   * A chart: its note (docx4j CR-047: `[chart: clustered column; value axis: Revenue]`), its title and its
   * data as a table, after the paragraph it is in, the categories down and the series across (CR-012 phase
   * 5c fills them in; until then the text model makes no chart block).
   */
  private chart(block: TextBlock, container: MdNode): void {
    if (block.chart !== undefined) {
      const note = tagged(new MdParagraph(), block);
      note.appendChild(new MdText(block.chart.note));
      container.appendChild(note);
    }
    if (block.chartTitle !== undefined && block.chartTitle !== '') {
      const title = tagged(new MdParagraph(), block);
      title.appendChild(new MdText(block.chartTitle));
      container.appendChild(title);
    }
    if (block.chartHeader.length === 0) return;
    const table = tagged(new MdTableBlock(), block);
    const head = new MdTableHead();
    head.appendChild(chartRow(block.chartHeader, true, block.chartHeader));
    table.appendChild(head);
    if (block.chartRows.length > 0) {
      const body = new MdTableBody();
      for (const r of block.chartRows) body.appendChild(chartRow(r, false, block.chartHeader));
      table.appendChild(body);
    }
    container.appendChild(table);
  }

  // ---------------------------------------------------------------- inlines

  private inlines(block: TextBlock, parent: MdNode, withListLabel: boolean): void {
    const sink = new InlineSink(parent);
    // under list_labels: true, the label the numbering model computes leads a paragraph that isn't written as
    // a list item (a numbered heading's "1.2")
    if (withListLabel && block.listLabel !== undefined && block.listLabel.trim() !== '') {
      sink.text(block.listLabel.replace(/\t/g, ' '), false, false, false);
    }
    this.runs(block, sink);
    sink.flush();
  }

  /** A block's runs, in `markup` each revision's between its CriticMarkup marks (REQ-703, REQ-702). */
  private runs(block: TextBlock, sink: InlineSink): void {
    const runs = block.runs;
    let i = 0;
    while (i < runs.length) {
      const revision = runs[i]!.revision;
      let j = i;
      while (j < runs.length && this.sameMarker(revision, runs[j]!.revision)) j++;
      if (revision !== undefined) this.openMarker(revision, sink);
      this.linked(runs.slice(i, j), sink);
      if (revision !== undefined) sink.node(new MdHtmlInline(close(revision)));
      i = j;
    }
  }

  /** Runs, those with the same link target in one link. */
  private linked(runs: readonly TextRun[], sink: InlineSink): void {
    let i = 0;
    while (i < runs.length) {
      let target = runs[i]!.link;
      let j = i;
      while (j < runs.length && runs[j]!.link === target) j++;
      if (target !== undefined && !LinkTargets.isSafe(target, this.linkSchemes)) {
        // docx4j CR-056: a script-scheme target; the text, with no link (CR-012 section 21)
        if (!this.linkRefused) log.warn(LinkTargets.refusedMessage(target));
        this.linkRefused = true;
        target = undefined;
      }
      if (target === undefined) {
        for (let k = i; k < j; k++) this.run(runs[k]!, sink);
      } else {
        const link = new MdLink(target, undefined);
        const linkSink = new InlineSink(link);
        for (let k = i; k < j; k++) this.run(runs[k]!, linkSink);
        linkSink.flush();
        sink.node(link);
      }
      i = j;
    }
  }

  private run(run: TextRun, sink: InlineSink): void {
    switch (run.kind) {
      case 'TEXT':
        this.text(run, sink);
        return;
      case 'FIELD_CODE':
      case 'LIST_LABEL':
        sink.text(run.text, false, false, false);
        return;
      case 'NOTE_REFERENCE':
        if (this.textOptions.notes && this.textOptions.noteReferences) {
          sink.node(new MdFootnoteReference(noteLabel(run.noteType!, run.noteNumber, run.noteId)));
        }
        return;
      case 'COMMENT_REFERENCE': {
        const comment = run.commentId === undefined ? undefined : this.result.comments.get(run.commentId);
        // a reply is written after its parent, never at its own reference (REQ-504)
        if (comment === undefined || (comment.parentId !== undefined && this.result.comments.has(comment.parentId))) return;
        this.comment(comment, sink);
        for (const reply of comment.replies) this.comment(reply, sink);
        return;
      }
      case 'IMAGE':
        this.image(run, sink);
        return;
      case 'PAGE_MARKER':
        if (this.textOptions.pageMarkers) sink.text(`[page ${run.pageNumber}]`, false, false, false);
        return;
      default:
        // a note's own mark (its definition's label carries it), an object (no text)
    }
  }

  private text(run: TextRun, sink: InlineSink): void {
    const value = run.text;
    if (value === '') return;
    if (run.math) {
      const l = latex(run);
      if (l !== '') sink.node(new MdInlineMath(l, false));
      return;
    }
    const bold = run.bold && this.tableHeaderDepth === 0;               // a header row's bold is convention, not markup
    const lines = value.split('\n');
    for (let l = 0; l < lines.length; l++) {
      if (l > 0) sink.node(new MdHardLineBreak());
      const line = lines[l]!;
      if (line === '') continue;
      if (run.monospace) {
        sink.node(new MdCode(line));
      } else {
        sink.text(line, bold, run.italic, run.strike);
      }
    }
  }

  private comment(comment: TextComment, sink: InlineSink): void {
    sink.node(new MdHtmlInline('{>>'));
    this.detailText(comment.author, comment.date, sink);
    sink.text(comment.text, false, false, false);
    sink.node(new MdHtmlInline('<<}'));
  }

  /** Runs under one marker: the same type, and the same author and date where the detail shows them. */
  private sameMarker(a: TextRevision | undefined, b: TextRevision | undefined): boolean {
    if (a === undefined || b === undefined) return a === b;
    if (a.type !== b.type) return false;
    switch (this.textOptions.markupDetail) {
      case 'author': return a.author === b.author;
      case 'author_date': return a.author === b.author && a.date === b.date;
      default: return true;
    }
  }

  /** The marker as one string, for a code block, where nothing is markup. */
  private open(revision: TextRevision): string {
    return (revision.type === 'INSERTION' ? '{++' : '{--') + this.detail(revision.author, revision.date);
  }

  /** The marker in a paragraph: its braces raw, its author and date text (docx4j's `open(revision, sink)`). */
  private openMarker(revision: TextRevision, sink: InlineSink): void {
    sink.node(new MdHtmlInline(revision.type === 'INSERTION' ? '{++' : '{--'));
    this.detailText(revision.author, revision.date, sink);
  }

  /**
   * The marker's author and date as text, which the renderer escapes as it does any document text (docx4j CR-056,
   * `30fb41024`; CR-012 section 21.1): only the braces are raw. Before, the detail went out raw with them, so an
   * author of `<script>...</script>` reached a renderer that allows raw HTML as markup.
   */
  private detailText(author: string | undefined, date: string | undefined, sink: InlineSink): void {
    const d = this.detail(author, date);
    if (d !== '') sink.text(d, false, false, false);
  }

  private detail(author: string | undefined, date: string | undefined): string {
    switch (this.textOptions.markupDetail) {
      case 'author': return `${author ?? ''}: `;
      case 'author_date': return `${author ?? ''}, ${date ?? ''}: `;
      default: return '';
    }
  }

  /**
   * A picture with the describer's text (REQ-703, REQ-514); its destination, which needs its image part's
   * bytes, is the exporter's to fill in (`images`), empty for a picture whose image can't be had.
   */
  private image(run: TextRun, sink: InlineSink): void {
    const image = new MdImage('', undefined);
    image.appendChild(new MdText(run.description === undefined ? '' : run.description.text));
    this.images.push({ node: image, run });
    sink.node(image);
  }
}

// ---------------------------------------------------------------- helpers

/** A node made for a block: it begins the block's Markdown (`MarkdownExporter.blocks`). */
function tagged<T extends MdNode>(node: T, block: TextBlock): T {
  node.tag = block;
  return node;
}

/** Accumulates same-formatted text, wrapping it on flush. */
class InlineSink {
  private buf = '';
  private bold = false;
  private italic = false;
  private strike = false;

  constructor(private readonly parent: MdNode) {}

  text(s: string, b: boolean, i: boolean, st: boolean): void {
    if ((b !== this.bold || i !== this.italic || st !== this.strike) && this.buf.length > 0) this.flush();
    this.bold = b;
    this.italic = i;
    this.strike = st;
    this.buf += s;
  }

  node(n: MdNode): void {
    this.flush();
    this.parent.appendChild(n);
  }

  flush(): void {
    if (this.buf.length === 0) return;
    let n: MdNode = new MdText(this.buf);
    // strong inside emphasis: the renderer then writes ***both***, rather than switching the inner marker to _
    if (this.bold) {
      const strong = new MdStrongEmphasis();
      strong.appendChild(n);
      n = strong;
    }
    if (this.italic) {
      const em = new MdEmphasis();
      em.appendChild(n);
      n = em;
    }
    if (this.strike) {
      const st = new MdStrikethrough('~~');
      st.appendChild(n);
      n = st;
    }
    this.parent.appendChild(n);
    this.buf = '';
  }
}

function isEmpty(block: TextBlock): boolean {
  for (const r of block.runs) {
    if (r.kind === 'TEXT' && r.text !== '') return false;
    if (r.kind === 'IMAGE' || r.kind === 'NOTE_REFERENCE') return false;
  }
  return true;
}

function isEmptyCell(cell: TextCell): boolean {
  for (const b of cell.blocks) {
    if (b.kind === 'TABLE') {
      for (const r of b.rows) for (const c of r.cells) if (!isEmptyCell(c)) return false;
    } else if (b.kind === 'PARAGRAPH' && !isEmpty(b)) {
      return false;
    }
  }
  return true;
}

/** The paragraph's one run, a display equation with no revision; else undefined. */
function soleDisplayMath(block: TextBlock): TextRun | undefined {
  let found: TextRun | undefined;
  for (const r of block.runs) {
    if (r.kind === 'TEXT' && r.text === '') continue;
    if (found !== undefined || !r.displayMath || r.revision !== undefined) return undefined;
    found = r;
  }
  return found;
}

/** An equation's LaTeX: its text without the delimiters. */
function latex(run: TextRun): string {
  const s = run.text;
  const d = run.displayMath ? 2 : 1;
  return s.length >= 2 * d ? s.substring(d, s.length - d) : s;
}

/**
 * A code paragraph: styled as one, or in a named paragraph style whose own font is mono (every run
 * monospaced, none of them beyond the style). Inline code is what a run adds to its paragraph style, so a
 * paragraph whose style is already mono would get no code marks at all; this is the paragraph-level
 * counterpart. A mono document default doesn't count (a document set in Courier New throughout is prose, not
 * one fence), which is why a named style is required.
 */
function isCodeParagraph(block: TextBlock): boolean {
  const style = block.style;
  if (style === undefined) return false;
  if (CODE_STYLES.has(style)) return true;
  if (!block.monospace) return false;
  for (const r of block.runs) if (r.kind === 'TEXT' && r.monospace) return false;
  return true;
}

function sameStory(a: TextBlock, b: TextBlock): boolean {
  return a.story === b.story && a.noteId === b.noteId;
}

function value(i: number | undefined): number {
  return i ?? 0;
}

/**
 * A list block. An ordered one starts at the count given (the item's own, or a reopened parent's last), else
 * (the paragraph's own level at the top) at the level's `w:start`; Markdown numbers the items on from there.
 */
function createListBlock(numbering: TextNumbering, own: boolean, tight: boolean, count: number | undefined): MdListBlock {
  // a level the numbering definitions don't give is a bullet list
  const bullet = numbering.bullet || numbering.format === undefined;
  let listBlock: MdListBlock;
  if (bullet) {
    listBlock = new MdBulletList();
  } else {
    const ordered = new MdOrderedList();
    if (count !== undefined) {
      if (count !== 1) ordered.markerStartNumber = count;
    } else if (own && numbering.level === 0 && numbering.start > 1) {
      ordered.markerStartNumber = numbering.start;
    }
    listBlock = ordered;
  }
  listBlock.tight = tight;
  return listBlock;
}

/**
 * Appends a list. Markdown runs two adjacent lists of one kind together, numbering the second on from the
 * first; a list that directly follows another of its kind therefore takes the other delimiter (`1.` or
 * `1)`) or bullet (`-` or `*`), which CommonMark reads as a new list.
 */
function appendList(parent: MdNode, list: MdListBlock): void {
  const previous = parent.lastChild;
  if (list instanceof MdOrderedList && previous instanceof MdOrderedList) {
    list.markerDelimiter = previous.markerDelimiter === ')' ? '.' : ')';
  } else if (list instanceof MdBulletList && previous instanceof MdBulletList) {
    list.marker = previous.marker === '*' ? '-' : '*';
  }
  parent.appendChild(list);
}

function paragraphsOf(cell: TextCell): TextBlock[] {
  const out: TextBlock[] = [];
  for (const b of cell.blocks) {
    if (b.kind === 'TABLE') {
      for (const r of b.rows) for (const c of r.cells) out.push(...paragraphsOf(c));
    } else if (b.kind === 'PARAGRAPH') {
      out.push(b);
    }
  }
  return out;
}

function cellAlignment(block: TextBlock): MdAlignment | undefined {
  switch (block.alignment) {
    case 'center': return 'CENTER';
    case 'right':
    case 'end': return 'RIGHT';
    default: return undefined;                                          // LEFT is GFM's default
  }
}

function chartRow(cells: readonly string[], header: boolean, headings: readonly string[]): MdTableRow {
  const row = new MdTableRow();
  let labels = 0;
  while (labels < headings.length && headings[labels] === '') labels++;
  for (let i = 0; i < cells.length; i++) {
    const cell = new MdTableCell(header);
    if (i >= Math.max(labels, 1)) cell.alignment = 'RIGHT';
    if (cells[i] !== '') cell.appendChild(new MdText(cells[i]!));
    row.appendChild(cell);
  }
  return row;
}

function isNote(block: TextBlock): boolean {
  return block.noteType !== undefined && (block.story === 'FOOTNOTE' || block.story === 'ENDNOTE' || block.story === 'TEXT_BOX');
}

function sameNote(a: TextBlock, b: TextBlock): boolean {
  return isNote(b) && a.noteType === b.noteType && a.noteId === b.noteId;
}

/**
 * A note's label: a footnote's number (`[^1]`), an endnote's prefixed `e` (`[^e1]`), since the two are
 * numbered apart (WD REQ-503); a note whose reference carries a custom mark has no number, so its `w:id`
 * labels it (`[^f-5]`, `[^e-5]`).
 */
function noteLabel(type: NoteType, number: number, id: string | undefined): string {
  const foot = type === 'footnote';
  if (number > 0) return foot ? String(number) : `e${number}`;
  return `${foot ? 'f-' : 'e-'}${String(id)}`;
}

/** Drops the leading space left behind by the note's own mark. */
function trimLeadingMarkerSpace(definition: MdFootnoteDefinition): void {
  const paragraph = definition.firstChild;
  if (paragraph instanceof MdParagraph && paragraph.firstChild instanceof MdText) {
    const text = paragraph.firstChild;
    if (text.literal.startsWith(' ')) text.literal = text.literal.substring(1);
  }
}

function close(revision: TextRevision): string {
  return revision.type === 'INSERTION' ? '++}' : '--}';
}
