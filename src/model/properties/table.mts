// docx4j `org.docx4j.model.table`: `TableContext`, `CellContext` and `TableContextTracker`
// (docx4j CR-030, VERSION_17_3_1 at 843ac12df; CR-007 here).
//
// A table style reaches a paragraph only through the paragraph's `CellContext`: the table style
// the table resolves to and the conditional formats (`tableStyleConditions.mts`) the paragraph
// is under.  `PropertyResolver.tableContext(tbl)` builds a `TableContext` per table;
// `forParagraph(tr, tc, pPr)` gives the context of a paragraph in one of its cells, and the
// resolver's overloads which take one put the table level where ECMA-376-1 17.7.2 puts it:
// above the document defaults, below the paragraph's own style.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { isEmptyPPr, isEmptyRPr } from './styleUtil.mjs';
import { rowsOf, cellsOf } from '../content/tree.mjs';
import {
  type Condition, type Look, PRECEDENCE, DEFAULT_LOOK, lookOf, rowBandSize, colBandSize, rowCnf,
  resolve, rowConditions, applicable, formatsText, key as conditionsKey,
} from './tableStyleConditions.mjs';

/** The conditions as bits over {@link PRECEDENCE}: an exact value for keying caches. */
function maskOf(conditions: ReadonlySet<Condition>): number {
  let mask = 0;
  PRECEDENCE.forEach((t, i) => { if (conditions.has(t)) mask |= 1 << i; });
  return mask;
}


/**
 * The table level of a paragraph's formatting, as a value: two paragraphs whose keys are equal
 * get the same contribution from their table. A structured key, never a joined string (docx4j
 * CR-030 D8: ids may contain the separator); `mask` is the conditions as a number, so a cache
 * keys on the table style id and then on it.
 */
export class CellContextKey {
  /** The conditions as bits over {@link PRECEDENCE}. */
  readonly mask: number;

  constructor(
    readonly tableStyleId: string | undefined,
    /** The conditional formats which give the paragraph text formatting, `wholeTable` included where the style has one. */
    readonly conditions: ReadonlySet<Condition>,
  ) {
    this.mask = maskOf(conditions);
  }

  equals(other: CellContextKey | undefined): boolean {
    return other !== undefined && this.tableStyleId === other.tableStyleId && this.mask === other.mask;
  }

  toString(): string {
    return `${this.tableStyleId ?? null} [${PRECEDENCE.filter((t) => this.conditions.has(t)).join(', ')}]`;
  }
}

/**
 * Where a paragraph sits in a table, as far as its formatting is concerned: the table style
 * which applies to the table, and the conditional formats the paragraph is under. Handed to the
 * `PropertyResolver` overloads which take one. Obtained from a {@link TableContext}
 * (`forParagraph`), from a {@link TableContextTracker} during a walk, or from
 * `PropertyResolver.cellContextOf(p)`. Undefined stands for "not in a table". Immutable.
 */
export class CellContext {
  /** The conditions the paragraph is under, gated by the table's `w:tblLook`. */
  readonly conditions: ReadonlySet<Condition>;
  /**
   * The table style's conditional formats which apply here and give text a `w:pPr` or `w:rPr`,
   * in the order they are applied (ECMA-376-1 17.7.6).
   */
  readonly textConditions: readonly wml.CTTblStylePr[];
  /**
   * Whether the table style gives this paragraph anything at all: its own `w:pPr` or `w:rPr`, or
   * a conditional format with one. Where it does not, resolving with this context gives the same
   * answer as resolving without one.
   */
  readonly formatsText: boolean;
  /** What the table level's composition depends on. A value, for keying caches. */
  readonly key: CellContextKey;

  constructor(
    /** The id of the table style the table resolves to, or undefined where none applies. */
    readonly tableStyleId: string | undefined,
    /** That style's `w:basedOn` chain merged (`getTableStyleChain`): shared, read it only. */
    readonly tableStyle: wml.Style,
    conditions: ReadonlySet<Condition> | undefined,
  ) {
    const c = new Set<Condition>(conditions ?? []);
    this.conditions = c;
    const text: wml.CTTblStylePr[] = [];
    const named = new Set<Condition>();
    for (const pr of applicable(tableStyle, c)) {
      if (formatsText(pr)) {
        text.push(pr);
        named.add(pr.type);
      }
    }
    this.textConditions = text;
    this.formatsText = text.length > 0 || !isEmptyPPr(tableStyle.pPr) || !isEmptyRPr(tableStyle.rPr);
    this.key = new CellContextKey(tableStyleId, named);
  }

  toString(): string {
    return `CellContext[${this.tableStyleId ?? null} ${conditionsKey(this.conditions)}]`;
  }
}

/** What a {@link TableContext} needs of the resolver; `PropertyResolver` satisfies it. */
export interface TableStyleSource {
  getTableStyleIdOf(tblPr: wml.CTTblPrBase | undefined): string | undefined;
  getTableStyleChain(styleId: string | undefined): wml.Style;
}

type TrPrItem = { name?: { localPart?: string }; value?: unknown };

function gridOf(tr: wml.Tr, name: 'gridBefore' | 'gridAfter'): number {
  for (const item of (tr.trPr?.cnfStyleOrDivIdOrGridBefore ?? []) as TrPrItem[]) {
    if (item?.name?.localPart === name) {
      const val = (item.value as { val?: number } | undefined)?.val;
      return typeof val === 'number' ? Math.max(0, Math.trunc(val)) : 0;
    }
  }
  return 0;
}

/**
 * What a table's paragraphs need to know about it to be formatted: the table style it resolves
 * to (its `w:basedOn` chain merged), the conditional formats its `w:tblLook` asks for, its band
 * sizes, and where each row and cell sits. {@link forParagraph} then gives the
 * {@link CellContext} of a paragraph in one of its cells.
 *
 * A reading of the table's *content* (rows, spans, look), so it is built fresh for each table by
 * `PropertyResolver.tableContext(tbl)` and held by the caller while it is in the table; the
 * resolver keeps none (a cached one would go stale when a row is added). The rows and cells are
 * counted with nested tables excluded, a cell's column its row's `w:gridBefore` plus the spans
 * before it.
 */
export class TableContext {
  /** The id of the table style the table resolves to, or undefined. */
  readonly tableStyleId: string | undefined;
  /** Whether the table states its own `w:tblStyle` (rather than taking the default table style). */
  readonly namesStyle: boolean;
  /** The table style's chain merged; shared, read it only. */
  readonly tableStyle: wml.Style;
  readonly look: Look;
  readonly rowBandSize: number;
  readonly colBandSize: number;
  readonly rowCount: number;
  readonly colCount: number;
  private readonly rowIndex = new Map<object, number>();
  /** [first grid column, span] per cell */
  private readonly cellColumns = new Map<object, readonly [number, number]>();

  /** Use `PropertyResolver.tableContext(tbl)`. */
  constructor(tbl: wml.Tbl, resolver: TableStyleSource) {
    const tblPr = tbl.tblPr as wml.CTTblPrBase | undefined;
    this.tableStyleId = resolver.getTableStyleIdOf(tblPr);
    this.namesStyle = tblPr?.tblStyle !== undefined;
    this.tableStyle = resolver.getTableStyleChain(this.tableStyleId);

    // the table's own tblPr decides the look and band sizes; the style's is the fallback
    const stylePr = this.tableStyle.tblPr;
    if (tblPr?.tblLook !== undefined) this.look = lookOf(tblPr.tblLook);
    else if (stylePr?.tblLook !== undefined) this.look = lookOf(stylePr.tblLook);
    else this.look = DEFAULT_LOOK;
    this.rowBandSize = tblPr?.tblStyleRowBandSize !== undefined ? rowBandSize(tblPr) : rowBandSize(stylePr);
    this.colBandSize = tblPr?.tblStyleColBandSize !== undefined ? colBandSize(tblPr) : colBandSize(stylePr);

    const rows = rowsOf(tbl);
    let cols = 0;
    rows.forEach((located, r) => {
      const tr = located.element.value;
      this.rowIndex.set(tr, r);
      let c = gridOf(tr, 'gridBefore');
      for (const cell of cellsOf(tr)) {
        const tc = cell.element.value;
        const val = tc.tcPr?.gridSpan?.val;
        const span = typeof val === 'number' ? Math.max(1, Math.trunc(val)) : 1;
        this.cellColumns.set(tc, [c, span]);
        c += span;
      }
      c += gridOf(tr, 'gridAfter');
      if (c > cols) cols = c;
    });
    this.rowCount = rows.length;
    this.colCount = cols;
  }

  /**
   * The context of a paragraph in this table: the conditions its row, its cell and its own
   * `w:cnfStyle` caches say it is under, or its position where they say nothing. A row or cell
   * this table does not hold (a nested table's, or one added since this context was built) is
   * under no condition.
   */
  forParagraph(tr: wml.Tr | undefined, tc: wml.Tc | undefined, pPr?: wml.PPr | undefined): CellContext {
    return new CellContext(this.tableStyleId, this.tableStyle, this.conditions(tr, tc, pPr));
  }

  /** The context of a cell, for its row and cell properties: {@link forParagraph} without a paragraph's own `w:cnfStyle`. */
  forCell(tr: wml.Tr | undefined, tc: wml.Tc | undefined): CellContext {
    return this.forParagraph(tr, tc, undefined);
  }

  /** The row-axis conditions a row is under, for its own properties; none for a row this table does not hold. */
  rowConditions(tr: wml.Tr | undefined): Set<Condition> {
    const r = tr === undefined ? undefined : this.rowIndex.get(tr);
    if (r === undefined) return new Set();
    return rowConditions(this.look, this.rowBandSize, r, this.rowCount, rowCnf(tr!.trPr));
  }

  /** The table style's `w:tblStylePr` entries which apply under these conditions, in order: all of them. */
  applicable(conditions: ReadonlySet<Condition>): wml.CTTblStylePr[] {
    return applicable(this.tableStyle, conditions);
  }

  /** The row's index in the table, nested tables excluded, or -1 for a row it does not hold. */
  rowIndexOf(tr: wml.Tr | undefined): number {
    const r = tr === undefined ? undefined : this.rowIndex.get(tr);
    return r === undefined ? -1 : r;
  }

  private conditions(tr: wml.Tr | undefined, tc: wml.Tc | undefined, pPr: wml.PPr | undefined): Set<Condition> {
    const r = tr === undefined ? undefined : this.rowIndex.get(tr);
    const cc = tc === undefined ? undefined : this.cellColumns.get(tc);
    if (r === undefined || cc === undefined) return new Set();
    return resolve(this.look, this.rowBandSize, this.colBandSize,
      r, this.rowCount, cc[0], cc[1], this.colCount,
      rowCnf(tr!.trPr), tc!.tcPr?.cnfStyle, pPr?.cnfStyle);
  }
}

const TBL = 'org_docx4j_wml.Tbl';
const TR = 'org_docx4j_wml.Tr';
const TC = 'org_docx4j_wml.Tc';
const P = 'org_docx4j_wml.P';
const TXBX_CONTENT = 'org_docx4j_wml.CTTxbxContent';

/**
 * The objects a walk entering which enters another story: `w:txbxContent`, and what a walk
 * reaches it through (a VML text box, a DrawingML anchor or inline). Word gives the paragraphs
 * of a text box anchored in a cell none of the table's formatting (docx4j CR-030 probe T3).
 */
const STORY_BOUNDARIES: ReadonlySet<string> = new Set([
  TXBX_CONTENT,
  'org_docx4j_vml.CTTextbox',
  'org_docx4j_dml_wordprocessingDrawing.Anchor',
  'org_docx4j_dml_wordprocessingDrawing.Inline',
]);

interface Frame {
  /** Undefined for a barrier: inside a text box, or with no resolver. */
  readonly table: TableContext | undefined;
  tr: wml.Tr | undefined;
  tc: wml.Tc | undefined;
}

/** What a tracker needs of the resolver. */
export interface TableContextSource {
  tableContext(tbl: wml.Tbl): TableContext;
}

/**
 * Knows, for each paragraph a walk reaches, the table context the paragraph is in: the way to
 * resolve every paragraph of a document in context (`PropertyResolver.cellContextOf(p)` builds a
 * {@link TableContext} per call and is for one paragraph at a time).
 *
 * docx4j's is a `TraversalUtil` callback to extend; here it is the stack on its own, driven by
 * any walker: call {@link enter} with each object before descending into it and the function it
 * returns afterwards, and {@link cellContext} for each paragraph. {@link walk} is a walker that
 * does so over a whole tree.
 *
 * The table on top of the stack is the paragraph's own (a nested table pushes its own context).
 * A text box is a story of its own, so the walk starts an empty stack inside every text box and
 * restores the cell's on the way out.
 */
export class TableContextTracker {
  /** One entry per table being walked, innermost last; a barrier (no table) inside a text box. */
  private readonly stack: Frame[] = [];

  constructor(private readonly resolver: TableContextSource | undefined) {}

  /**
   * Call before descending into an object (a `w:tbl`, `w:tr`, `w:tc`, or anything else: other
   * types change nothing unless they begin a story); call what it returns after its children.
   */
  enter(o: unknown): () => void {
    const typeName = (o as { TYPE_NAME?: string } | null | undefined)?.TYPE_NAME;
    if (typeName === undefined) return NOTHING;
    const enclosing = this.stack[this.stack.length - 1];
    if (typeName === TBL) {
      // with no resolver, nothing is tracked: every paragraph is in no table
      this.stack.push({ table: this.resolver?.tableContext(o as wml.Tbl), tr: undefined, tc: undefined });
      return () => { this.stack.pop(); };
    }
    if (STORY_BOUNDARIES.has(typeName)) {
      this.stack.push({ table: undefined, tr: undefined, tc: undefined });
      return () => { this.stack.pop(); };
    }
    if (enclosing?.table !== undefined && typeName === TR) {
      const { tr, tc } = enclosing;
      enclosing.tr = o as wml.Tr;
      enclosing.tc = undefined;
      return () => { enclosing.tr = tr; enclosing.tc = tc; };
    }
    if (enclosing?.table !== undefined && typeName === TC) {
      const tc = enclosing.tc;
      enclosing.tc = o as wml.Tc;
      return () => { enclosing.tc = tc; };
    }
    return NOTHING;
  }

  /** The context of a paragraph reached by this walk, or undefined where it is in no table (or in a text box). */
  cellContext(p: wml.P): CellContext | undefined {
    const f = this.stack[this.stack.length - 1];
    if (f?.table === undefined) return undefined;
    return f.table.forParagraph(f.tr, f.tc, p.pPr);
  }

  /** The table the walk is in, or undefined (outside any table, or in a text box). */
  tableContext(): TableContext | undefined {
    return this.stack[this.stack.length - 1]?.table;
  }

  /**
   * Walks everything under `root` (an object, an element pair or an array of them), in document
   * order, calling `visit` for each paragraph with its context. It descends through every
   * member but the `PARENT` pointers, so a text box's paragraphs are reached, in no table.
   */
  walk(root: unknown, visit: (p: wml.P, cellContext: CellContext | undefined) => void): void {
    const seen = new Set<object>();
    const descend = (o: unknown, depth: number): void => {
      if (o === null || typeof o !== 'object' || depth > 200) return;
      if (Array.isArray(o)) {
        for (const item of o) descend(item, depth + 1);
        return;
      }
      if (seen.has(o)) return;
      seen.add(o);
      const typeName = (o as { TYPE_NAME?: string }).TYPE_NAME;
      if (typeName === P) visit(o as wml.P, this.cellContext(o as wml.P));
      const leave = this.enter(o);
      for (const key of Object.keys(o)) {
        if (key === 'PARENT') continue;
        descend((o as Record<string, unknown>)[key], depth + 1);
      }
      leave();
    };
    descend(root, 0);
  }
}

const NOTHING = (): void => {};

/** The types at which a walk up the parent pointers has left the cell's story, or reached a story root. */
const STORY_ROOTS: ReadonlySet<string> = new Set([
  TXBX_CONTENT, 'org_docx4j_wml.CTFtnEdn', 'org_docx4j_wml.Comments.Comment',
  'org_docx4j_wml.Hdr', 'org_docx4j_wml.Ftr', 'org_docx4j_wml.Body', 'org_docx4j_wml.Document',
]);

/**
 * The nearest enclosing cell, row and table of a paragraph, through its `PARENT` pointers and
 * whatever lies between (`w:sdt`, `w:customXml`, `w:smartTag`); undefined where the paragraph is
 * in no table, where a story begins before a cell is reached, or where the pointers are not
 * there (content built by hand and never linked: `linkParents`).
 */
export function enclosingCellOf(p: wml.P): { tbl: wml.Tbl; tr: wml.Tr; tc: wml.Tc } | undefined {
  let tc: wml.Tc | undefined;
  let tr: wml.Tr | undefined;
  const parentOf = (o: unknown): unknown => (o as { PARENT?: unknown } | null | undefined)?.PARENT;
  let hops = 0;
  for (let o = parentOf(p); o !== undefined && o !== null && hops < 1000; o = parentOf(o), hops++) {
    const typeName = (o as { TYPE_NAME?: string }).TYPE_NAME;
    if (typeName === undefined) continue;
    if (STORY_ROOTS.has(typeName)) return undefined;
    if (typeName === TC) {
      tc ??= o as wml.Tc;
    } else if (typeName === TR) {
      if (tc !== undefined && tr === undefined) tr = o as wml.Tr;
    } else if (typeName === TBL) {
      return tr === undefined || tc === undefined ? undefined : { tbl: o as wml.Tbl, tr, tc };
    }
  }
  return undefined;
}
