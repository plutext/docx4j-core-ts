// CR-012 phase 2: plain text from the block model (WD section 7.1), docx4j's org.docx4j.model.text.PlainText.
import type { TextOptions } from './TextOptions.mjs';
import type { TextResult, TextBlock, TextRun, TextRow, TextCell, TextRevision, TextComment } from './TextResult.mjs';

/** The spaces for each 720-twip tab stop of indent under `layout: indent` (REQ-709). */
export const INDENT_SPACES = 4;

/** The note line before a table in the listed format (REQ-605, REQ-707). */
export const LISTED_FORMAT_NOTE = '[table: listed format]';

/**
 * Plain text from the block model (WD section 7.1): each paragraph one line, a break one line more
 * (REQ-603); the text exactly as the document holds it (REQ-602), U+000A between lines and after the last.
 *
 * Tables (REQ-605): a simple table a row a line, its cells a tab apart; any other in the listed format,
 * after the note line `[table: listed format]`, a row a line as `Header: value; Header: value`. `tables:
 * tsv` writes every table in the simple form, `tables: grid` as a pandoc grid table, its spans drawn.
 *
 * What the extractor adds follows the WD's bracket convention (REQ-707): `[n]` at a note's reference and
 * at the start of the note; `[image: ...]` (or `[image, generated: ...]`) under `alt_text: true`; `[page
 * n]` under `page_markers: true`; and under `metadata: true`, before the body, one line per kind of
 * property that has any (REQ-515). In the markup view, insertions are `{++text++}`, deletions `{--text--}`
 * and comments `{>>text<<}`, at their references with their replies after them (REQ-702, REQ-504); under
 * `markup_detail` the marker's opening is followed by `author:` or `author, date:`.
 *
 * Under `layout: indent` (REQ-709) a paragraph outside a table begins with four spaces for each tab stop
 * (720 twips) of its first line's indent, rounded to the nearest stop. docx4j `PlainText`.
 */
export const PlainText = {
  /** The plain text of a result. */
  of(result: TextResult): string {
    const w = new Writer(result);
    if (result.options.metadata) w.metadata(result.metadata);
    for (const block of result.blocks) {
      w.top(block);
      w.out += '\n';
    }
    return w.out;
  },

  /** Whether a table is written in REQ-605's simple form under `tables: auto`. */
  isSimple(table: TextBlock): boolean {
    return isSimple(table);
  },
};

/**
 * No merged or split cells, no nested table, one paragraph a cell. A cell is merged or split when it
 * crosses a column edge another row has, or is in a vertical merge; a cell spanning grid columns that
 * every row spans alike merges nothing. A cell's empty paragraphs don't count.
 */
function isSimple(table: TextBlock): boolean {
  let edgesBefore: string | undefined;
  for (const row of table.rows) {
    const own = [...edges(row)].join(',');
    if (edgesBefore !== undefined && edgesBefore !== own) return false;
    edgesBefore = own;
    for (const cell of row.cells) {
      if (cell.mergeContinue || cell.mergeRestart) return false;
      let paragraphs = 0;
      for (const b of cell.blocks) {
        if (b.kind !== 'PARAGRAPH') return false;
        if (b.runs.length > 0 || b.listLabel !== undefined) paragraphs++;
      }
      if (paragraphs > 1) return false;
    }
  }
  return true;
}

/** A row's cell edges in grid columns, by gridSpan: 0, and each cell's end. */
function edges(row: TextRow): Set<number> {
  const out = new Set<number>([0]);
  let col = 0;
  for (const c of row.cells) out.add((col += Math.max(1, c.gridSpan)));
  return out;
}

/** The grid columns each cell of a row starts at, by gridSpan. */
function starts(row: TextRow): number[] {
  const out: number[] = [];
  let col = 0;
  for (const c of row.cells) {
    out.push(col);
    col += Math.max(1, c.gridSpan);
  }
  return out;
}

/** The header rows: those marked `w:tblHeader`, else the first. */
function headerRows(rows: readonly TextRow[]): number {
  let n = 0;
  while (n < rows.length && rows[n]!.header) n++;
  return n > 0 ? n : Math.min(1, rows.length);
}

/** The grid's width: the most grid columns a row spans. */
function width(rows: readonly TextRow[]): number {
  let w = 0;
  for (const row of rows) w = Math.max(w, row.cells.reduce((n, c) => n + Math.max(1, c.gridSpan), 0));
  return w;
}

/** A row of one cell spanning every grid column: a title, or a divider such as "Part A". */
function fullWidth(row: TextRow, gridWidth: number): boolean {
  return row.cells.length === 1 && Math.max(1, row.cells[0]!.gridSpan) === gridWidth;
}

/** The first header with text over a cell's grid columns. */
function header(headers: Map<number, string>, start: number, span: number): string | undefined {
  for (let g = start; g < start + span; g++) {
    const h = headers.get(g);
    if (h !== undefined && h !== '') return h;
  }
  return undefined;
}

function label(head: string | undefined, value: string): string {
  return head === undefined || head === '' ? value : `${head}: ${value}`;
}

/** A cell placed on the grid. */
interface Placed { row: number; col: number; span: number; rowSpan: number; lines: string[] }

class Writer {
  out = '';
  private readonly options: TextOptions;

  constructor(private readonly result: TextResult) {
    this.options = result.options;
  }

  // --- the metadata block (REQ-515) ---------------------------------------------------------------

  metadata(metadata: Readonly<Record<string, unknown>>): void {
    this.group('core properties', metadata.core);
    this.group('application properties', metadata.app);
    this.group('custom properties', metadata.custom);
  }

  private group(name: string, properties: unknown): void {
    if (typeof properties !== 'object' || properties === null) return;
    const entries = Object.entries(properties);
    if (entries.length === 0) return;
    this.out += `[${name}: ${entries.map(([k, v]) => `${k}: ${String(v)}`).join('; ')}]\n`;
  }

  // --- blocks -----------------------------------------------------------------------------------

  /** A block of the story: a paragraph with its indent, a table, a chart. */
  top(block: TextBlock): void {
    if (block.kind === 'PARAGRAPH' && this.options.layout === 'indent') this.indent(block);
    this.block(block);
  }

  private indent(block: TextBlock): void {
    const twips = (block.indentLeft ?? 0) + (block.indentFirstLine ?? 0) - (block.indentHanging ?? 0);
    const stops = Math.round(twips / 720);                            // Java's Math.round: half up
    for (let i = 0; i < stops * INDENT_SPACES; i++) this.out += ' ';
  }

  private block(block: TextBlock): void {
    if (block.kind === 'TABLE') this.table(block);
    else if (block.kind === 'CHART') this.chart(block);
    else this.paragraph(block);
  }

  private chart(block: TextBlock): void {
    let any = false;
    if (block.chart) {
      this.out += block.chart.note;
      any = true;
    }
    if (block.chartTitle) {
      if (any) this.out += '\n';
      this.out += block.chartTitle;
      any = true;
    }
    const lines: string[][] = [];
    if (block.chartHeader.length > 0) lines.push(block.chartHeader);
    lines.push(...block.chartRows);
    for (const line of lines) {
      if (any) this.out += '\n';
      this.out += line.join('\t');
      any = true;
    }
  }

  private paragraph(block: TextBlock): void {
    if (block.listLabel !== undefined) this.out += block.listLabel;
    const runs = block.runs;
    let i = 0;
    while (i < runs.length) {
      const revision = runs[i]!.revision;
      let j = i;
      while (j < runs.length && this.sameMarker(revision, runs[j]!.revision)) j++;
      if (revision) {
        this.out += revision.type === 'INSERTION' ? '{++' : '{--';
        this.detail(revision.author, revision.date);
      }
      for (let k = i; k < j; k++) this.run(runs[k]!);
      if (revision) this.out += revision.type === 'INSERTION' ? '++}' : '--}';
      i = j;
    }
  }

  /** Runs under one marker: the same type, and the same author and date where the detail shows them. */
  private sameMarker(a: TextRevision | undefined, b: TextRevision | undefined): boolean {
    if (a === undefined || b === undefined) return a === b;
    if (a.type !== b.type) return false;
    switch (this.options.markupDetail) {
      case 'author': return a.author === b.author;
      case 'author_date': return a.author === b.author && a.date === b.date;
      default: return true;
    }
  }

  private detail(author: string | undefined, date: string | undefined): void {
    switch (this.options.markupDetail) {
      case 'author': this.out += `${author ?? ''}: `; break;
      case 'author_date': this.out += `${author ?? ''}, ${date ?? ''}: `; break;
      default:
    }
  }

  private run(run: TextRun): void {
    switch (run.kind) {
      case 'TEXT': case 'FIELD_CODE': case 'LIST_LABEL':
        this.out += run.text;
        break;
      case 'NOTE_REFERENCE':
        if (this.options.noteReferences && this.options.notes && run.noteNumber > 0) this.out += `[${run.noteNumber}]`;
        break;
      case 'NOTE_MARK':
        if (run.noteNumber > 0) this.out += `[${run.noteNumber}]`;
        break;
      case 'COMMENT_REFERENCE': {
        const comment = run.commentId === undefined ? undefined : this.result.comments.get(run.commentId);
        if (!comment || (comment.parentId !== undefined && this.result.comments.has(comment.parentId))) break;
        this.comment(comment);
        for (const reply of comment.replies) this.comment(reply);
        break;
      }
      case 'IMAGE': {
        const d = run.description;
        if (!this.options.altText || !d) break;
        this.out += `${d.provenance === 'GENERATED' ? '[image, generated: ' : '[image: '}${d.text}]`;
        break;
      }
      case 'PAGE_MARKER':
        if (this.options.pageMarkers) this.out += `[page ${run.pageNumber}]`;
        break;
      default:
    }
  }

  private comment(comment: TextComment): void {
    this.out += '{>>';
    this.detail(comment.author, comment.date);
    this.out += `${comment.text}<<}`;
  }

  /** A block written on its own, for a table cell's value. */
  private text(block: TextBlock): string {
    const saved = this.out;
    this.out = '';
    try {
      this.block(block);
      return this.out;
    } finally {
      this.out = saved;
    }
  }

  // --- tables (REQ-605) -------------------------------------------------------------------------

  private table(table: TextBlock): void {
    let form = this.options.tables;
    if (table.simpleForm) form = 'tsv';
    if (form === 'grid') this.grid(table);
    else if (form === 'tsv' || isSimple(table)) this.tabSeparated(table);
    else this.listed(table);
  }

  private tabSeparated(table: TextBlock): void {
    const rows = table.rows;
    for (let r = 0; r < rows.length; r++) {
      if (r > 0) this.out += '\n';
      const cells = rows[r]!.cells;
      for (let c = 0; c < cells.length; c++) {
        if (c > 0) this.out += '\t';
        this.out += this.values(cells[c]!).join(' ');
      }
    }
  }

  /**
   * The listed format (REQ-605). Every cell is labelled by the header over its first grid column, or the
   * first of its columns' headers that has text; several header rows' texts are joined ` / `; a cell under
   * a header with no text is its value alone. A row whose one cell spans every grid column is that cell's
   * text alone, and labels nothing. A cell with no value is left out; a vertical merge's value is on every
   * row it covers. Labelled cells are `; ` apart; in a table with no labels cells are a tab apart.
   */
  private listed(table: TextBlock): void {
    this.out += LISTED_FORMAT_NOTE;
    const rows = table.rows;
    const gridWidth = width(rows);
    const headerCount = headerRows(rows);
    // a table of header rows only is written unlabelled, its rows as body rows
    const headersOnly = headerCount >= rows.length;
    const headers = new Map<number, string>();
    if (!headersOnly) {
      for (let r = 0; r < headerCount; r++) {
        const row = rows[r]!;
        if (fullWidth(row, gridWidth)) {
          const text = this.value(row.cells[0]!);
          if (text !== '') this.out += `\n${text}`;
          continue;
        }
        const at = starts(row);
        for (let c = 0; c < at.length; c++) {
          const cell = row.cells[c]!;
          const text = this.value(cell);
          if (text === '') continue;
          for (let g = at[c]!; g < at[c]! + Math.max(1, cell.gridSpan); g++) {
            const h = headers.get(g);
            headers.set(g, h === undefined ? text : `${h} / ${text}`);
          }
        }
      }
    }
    const labelled = headers.size > 0;
    const body = headersOnly ? rows : rows.slice(headerCount);
    const restarts = new Map<number, string>();
    for (const row of body) {
      this.out += '\n';
      if (fullWidth(row, gridWidth)) {
        this.out += this.value(row.cells[0]!);
        continue;
      }
      const at = starts(row);
      const parts: string[] = [];
      for (let c = 0; c < at.length; c++) {
        const cell = row.cells[c]!;
        let value: string;
        if (cell.mergeContinue) {
          value = restarts.get(at[c]!) ?? '';
        } else {
          value = this.value(cell);
          if (cell.mergeRestart) restarts.set(at[c]!, value);
        }
        if (value === '') continue;
        parts.push(labelled ? label(header(headers, at[c]!, Math.max(1, cell.gridSpan)), value) : value);
      }
      this.out += parts.join(labelled ? '; ' : '\t');
    }
  }

  /** A cell's paragraphs, a nested table's included, `; ` apart, empty ones left out. */
  private value(cell: TextCell): string {
    return this.values(cell).join('; ');
  }

  private values(cell: TextCell, out: string[] = []): string[] {
    for (const b of cell.blocks) {
      if (b.kind === 'TABLE') {
        for (const r of b.rows) for (const c of r.cells) this.values(c, out);
      } else {
        const s = this.text(b);
        if (s !== '') out.push(s);
      }
    }
    return out;
  }

  // --- pandoc grid tables -----------------------------------------------------------------------

  private grid(table: TextBlock): void {
    for (const line of this.gridLines(table)) {
      if (this.out.length > 0 && !this.out.endsWith('\n')) this.out += '\n';
      this.out += line;
    }
  }

  private gridLines(table: TextBlock): string[] {
    const rows = table.rows;
    const placed: Placed[] = [];
    const at: (Placed | undefined)[][] = [];
    // the grid's columns are those some row's cell edges mark, so a span every row shares draws as one column
    const union = new Set<number>();
    for (const row of rows) for (const e of edges(row)) union.add(e);
    const edgeList = [...union].sort((a, b) => a - b);
    const columns = edgeList.length - 1;
    if (columns <= 0) return [];
    const open = new Map<number, Placed>();
    for (let r = 0; r < rows.length; r++) {
      at[r] = new Array<Placed | undefined>(columns).fill(undefined);
      const row = rows[r]!;
      const begin = starts(row);
      for (let c = 0; c < begin.length; c++) {
        const cell = row.cells[c]!;
        const first = edgeList.indexOf(begin[c]!);
        const span = edgeList.indexOf(begin[c]! + Math.max(1, cell.gridSpan)) - first;
        let p: Placed;
        const opened = open.get(first);
        if (cell.mergeContinue && opened !== undefined && opened.span === span) {
          p = opened;
          p.rowSpan++;
        } else {
          p = { row: r, col: first, span, rowSpan: 1, lines: this.cellLines(cell) };
          placed.push(p);
          if (cell.mergeRestart) open.set(first, p);
          else open.delete(first);
        }
        for (let g = first; g < first + span && g < columns; g++) at[r]![g] = p;
      }
    }
    // column widths: single-column cells first, then spans widen their last column
    const widths = new Array<number>(columns).fill(3);
    for (const p of placed) if (p.span === 1) for (const l of p.lines) widths[p.col] = Math.max(widths[p.col]!, l.length);
    for (const p of placed) {
      if (p.span === 1) continue;
      const last = Math.min(columns, p.col + p.span) - 1;
      let have = 0;
      for (let g = p.col; g <= last; g++) have += widths[g]!;
      have += 3 * (last - p.col);
      for (const l of p.lines) if (l.length > have) { widths[last] = widths[last]! + (l.length - have); have = l.length; }
    }
    // row heights: a cell's lines over the rows it spans
    const heights = new Array<number>(rows.length).fill(1);
    for (const p of placed) if (p.rowSpan === 1) heights[p.row] = Math.max(heights[p.row]!, p.lines.length);
    for (const p of placed) {
      if (p.rowSpan === 1) continue;
      let have = 0;
      for (let r = p.row; r < p.row + p.rowSpan; r++) have += heights[r]!;
      if (p.lines.length > have) heights[p.row + p.rowSpan - 1] = heights[p.row + p.rowSpan - 1]! + (p.lines.length - have);
    }
    const headerCount = rows.length > 1 ? headerRows(rows) : 0;
    const out: string[] = [border(undefined, at[0], widths, '-')];
    const written = new Map<Placed, number>();
    for (let r = 0; r < rows.length; r++) {
      for (let line = 0; line < heights[r]!; line++) {
        let l = '|';
        let g = 0;
        while (g < columns) {
          const p = at[r]![g];
          const span = p === undefined ? 1 : Math.min(p.span, columns - g);
          let w = 0;
          for (let k = g; k < g + span; k++) w += widths[k]!;
          w += 3 * (span - 1);
          let text = '';
          if (p) {
            const n = written.get(p) ?? 0;
            if (n < p.lines.length) text = p.lines[n]!;
            written.set(p, n + 1);
          }
          l += ` ${text}${' '.repeat(Math.max(0, w - text.length))} |`;
          g += span;
        }
        out.push(l);
      }
      if (r < rows.length - 1) out.push(border(at[r], at[r + 1], widths, r === headerCount - 1 ? '=' : '-'));
    }
    out.push(border(at[rows.length - 1], undefined, widths, '-'));
    return out;
  }

  /** A cell's lines: each paragraph's lines, a blank line between paragraphs, a nested table drawn as a grid. */
  private cellLines(cell: TextCell): string[] {
    const lines: string[] = [];
    for (const b of cell.blocks) {
      let own: string[];
      if (b.kind === 'TABLE') {
        own = this.gridLines(b);
      } else {
        const s = this.text(b);
        if (s === '') continue;
        own = s.split('\n');
      }
      if (lines.length > 0) lines.push('');
      lines.push(...own);
    }
    return lines;
  }
}

/** The border between two rows: open where a cell spans both. */
function border(above: (Placed | undefined)[] | undefined, below: (Placed | undefined)[] | undefined, widths: number[], fill: string): string {
  let l = '';
  for (let g = 0; g < widths.length; g++) {
    const through = above !== undefined && below !== undefined && above[g] !== undefined && above[g] === below[g];
    const joinLeft = g > 0 && through && above![g - 1] === above![g];
    l += g === 0 ? (through ? '|' : '+') : joinLeft ? ' ' : '+';
    l += (through ? ' ' : fill).repeat(widths[g]! + 2);
  }
  const lastColumn = widths.length - 1;
  const throughLast = above !== undefined && below !== undefined && above[lastColumn] !== undefined && above[lastColumn] === below[lastColumn];
  return l + (throughLast ? '|' : '+');
}
