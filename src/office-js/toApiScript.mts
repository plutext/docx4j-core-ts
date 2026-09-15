// `toApiScript` (CR-002 section 3.10): the content-API calls that reproduce a paragraph, a range,
// a table or a body, emitted as TypeScript against `body`. The counterpart of the objects
// package's `toSource`, which emits the factory calls that build the element; this one stays in
// the vocabulary an add-in author (or an agent writing one) reads, and falls back to `insertXml`
// with the marshalled fragment for what the verbs cannot express.

import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { marshalString } from '@docx4j/generated-objects-ts';
import { readRunOptions, type RunFormatting } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../opc/exceptions.mjs';
import { Body } from '../model/content/Body.mjs';
import { Paragraph } from '../model/content/Paragraph.mjs';
import { Range } from '../model/content/Range.mjs';
import { type Element, isElement, typeNameOf, textOf } from '../model/content/tree.mjs';
import { paragraphElementOfRange } from './fragment.mjs';
import { builtInOf, displayNameOf } from '../model/content/styles.mjs';

export interface ApiScriptOptions {
  /** The variable the script inserts into; default `'body'`. */
  variable?: string;
  /** The insert location; default `'End'`. */
  location?: 'Start' | 'End';
  /** Prefix for every line; default `''`. */
  indent?: string;
}

/** What `toApiScript` takes: a view, an element of the tree, or several. */
export type ApiScriptTarget = Body | Paragraph | Range | Element | Element[];

const TWIPS_PER_POINT = 20;

/** Paragraph properties the verbs express; anything else sends the paragraph to `insertXml`. */
const PPR_KEYS = new Set(['TYPE_NAME', 'PARENT', 'pStyle', 'jc', 'ind', 'spacing', 'outlineLvl']);
const IND_KEYS = new Set(['TYPE_NAME', 'PARENT', 'left', 'right', 'firstLine', 'hanging']);
const SPACING_KEYS = new Set(['TYPE_NAME', 'PARENT', 'before', 'after', 'line', 'lineRule']);
/** Run properties the `Font` members express (the objects package's applyRunOptions vocabulary). */
const RPR_KEYS = new Set(['TYPE_NAME', 'PARENT', 'b', 'bCs', 'i', 'iCs', 'u', 'strike', 'dstrike', 'vertAlign', 'rFonts', 'sz', 'szCs', 'color', 'highlight']);
const RFONTS_KEYS = new Set(['TYPE_NAME', 'PARENT', 'ascii', 'hAnsi']);
/** Markers that carry no content: dropped, since no verb writes them. */
const IGNORED_TYPES = new Set(['org_docx4j_wml.ProofErr', 'org_docx4j_wml.CTEmpty']);
const BOOKMARK_TYPES = new Set(['org_docx4j_wml.CTBookmark', 'org_docx4j_wml.CTMarkupRange']);
/** Table properties `insertTable` can reproduce. */
const TBLPR_KEYS = new Set(['TYPE_NAME', 'PARENT', 'tblStyle', 'tblW', 'tblLook']);

/** Marks a construct the verbs cannot express: the caller falls back to `insertXml`. */
class Unexpressible extends Error {}

interface TextPiece {
  kind: 'text' | 'break';
  text: string;
  breakType: 'Page' | 'Line';
  format: RunFormatting;
}

function quote(text: string): string {
  const escaped = text
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\t/g, '\\t');
  return `'${escaped}'`;
}

function template(xml: string): string {
  return '`' + xml.replace(/\\/g, '\\\\').replace(/`/g, '\\`').replace(/\$\{/g, '\\${') + '`';
}

function num(value: number): string {
  return String(Math.round(value * 100) / 100);
}

function keysOf(value: object): string[] {
  return Object.keys(value).filter((k) => (value as Record<string, unknown>)[k] !== undefined);
}

function within(value: object, allowed: Set<string>): boolean {
  return keysOf(value).every((k) => allowed.has(k));
}

const DEFAULT_FORMAT = readRunOptions(undefined);

/** The `Font` assignments that turn `from` into `to`; throws when a property cannot be unset. */
function fontDiff(from: RunFormatting, to: RunFormatting): [string, string][] {
  const out: [string, string][] = [];
  const flags = ['bold', 'italic', 'strikeThrough', 'doubleStrikeThrough', 'subscript', 'superscript'] as const;
  for (const flag of flags) if (from[flag] !== to[flag]) out.push([flag, String(to[flag])]);
  if (from.underline !== to.underline) out.push(['underline', quote(to.underline)]);
  if (from.name !== to.name) out.push(['name', quote(to.name)]);
  if (from.size !== to.size) {
    // Font.size writes w:sz; there is no size that removes it, so a run that drops a direct size
    // after one that set it is not expressible (CR-002 section 10).
    if (to.size === 0) throw new Unexpressible('a run drops a direct font size');
    out.push(['size', num(to.size)]);
  }
  if (from.color !== to.color) out.push(['color', quote(to.color)]);
  if (from.highlightColor !== to.highlightColor) out.push(['highlightColor', to.highlightColor === null ? 'null' : quote(to.highlightColor)]);
  if (from.style !== to.style) throw new Unexpressible('a run style (w:rStyle)');
  return out;
}

/** The emitter: one instance per call, so that variable names are stable and readable. */
class Emitter {
  private readonly lines: string[] = [];
  private readonly variable: string;
  private readonly location: string;
  private readonly indent: string;
  private paragraphs = 0;
  private ranges = 0;
  private tables = 0;

  constructor(options: ApiScriptOptions) {
    this.variable = options.variable ?? 'body';
    this.location = options.location ?? 'End';
    this.indent = options.indent ?? '';
  }

  script(): string {
    return this.lines.join('\n');
  }

  private line(text: string): void {
    this.lines.push(this.indent + text);
  }

  /** A block-level element: a paragraph, a table, or anything else through `insertXml`. */
  async element(element: Element): Promise<void> {
    const typeName = typeNameOf(element);
    if (typeName === 'org_docx4j_wml.P') await this.paragraph(element as Element<wml.P>);
    else if (typeName === 'org_docx4j_wml.Tbl') await this.table(element as Element<wml.Tbl>);
    else await this.fallback(element);
  }

  async paragraph(element: Element<wml.P>): Promise<void> {
    try {
      this.emitParagraph(element.value);
    } catch (error) {
      if (!(error instanceof Unexpressible)) throw error;
      await this.fallback(element, error.message);
    }
  }

  /** A range: its text and run formatting, as one paragraph. */
  async range(range: Range): Promise<void> {
    await this.paragraph(paragraphElementOfRange(range));
  }

  private emitParagraph(p: wml.P): void {
    const properties = this.paragraphProperties(p);
    const pieces = piecesOf(p);
    const first = pieces[0];
    const plainFirst = first !== undefined && first.kind === 'text' && sameFormat(first.format, DEFAULT_FORMAT);
    // Every font change is checked before a line is emitted, so a paragraph either comes out whole
    // or falls back whole.
    let previous = DEFAULT_FORMAT;
    const diffs: ([string, string][] | undefined)[] = [];
    for (const piece of pieces.slice(plainFirst ? 1 : 0)) {
      if (piece.kind === 'break') {
        diffs.push(undefined);
        previous = DEFAULT_FORMAT;
        continue;
      }
      diffs.push(fontDiff(previous, piece.format));
      previous = piece.format;
    }

    if (properties.length === 0 && plainFirst && pieces.length === 1) {
      this.line(`${this.variable}.insertParagraph(${quote(first.text)}, '${this.location}');`);
      return;
    }
    const name = `p${++this.paragraphs}`;
    this.line(`const ${name} = ${this.variable}.insertParagraph(${quote(plainFirst ? first.text : '')}, '${this.location}');`);
    for (const [member, value] of properties) this.line(`${name}.${member} = ${value};`);
    pieces.slice(plainFirst ? 1 : 0).forEach((piece, i) => {
      if (piece.kind === 'break') {
        this.line(`${name}.insertBreak('${piece.breakType}', 'End');`);
        return;
      }
      const diff = diffs[i] ?? [];
      if (diff.length === 0) {
        this.line(`${name}.insertText(${quote(piece.text)}, 'End');`);
        return;
      }
      const range = `r${++this.ranges}`;
      this.line(`const ${range} = ${name}.insertText(${quote(piece.text)}, 'End');`);
      for (const [member, value] of diff) this.line(`${range}.font.${member} = ${value};`);
    });
  }

  /** The paragraph members that carry `w:pPr`; throws when it holds anything else. */
  private paragraphProperties(p: wml.P): [string, string][] {
    const pPr = p.pPr;
    if (!pPr) return [];
    if (!within(pPr, PPR_KEYS)) throw new Unexpressible(`paragraph properties (${keysOf(pPr).filter((k) => !PPR_KEYS.has(k)).join(', ')})`);
    const out: [string, string][] = [];
    const style = pPr.pStyle?.val;
    if (style !== undefined && style !== 'Normal') out.push(styleAssignment(style));
    const jc = pPr.jc?.val;
    if (jc !== undefined) {
      const alignment = jc === 'left' ? 'Left' : jc === 'center' ? 'Centered' : jc === 'right' ? 'Right' : jc === 'both' ? 'Justified' : undefined;
      if (alignment === undefined) throw new Unexpressible(`w:jc ${jc}`);
      out.push(['alignment', quote(alignment)]);
    }
    const ind = pPr.ind;
    if (ind) {
      if (!within(ind, IND_KEYS)) throw new Unexpressible(`w:ind (${keysOf(ind).filter((k) => !IND_KEYS.has(k)).join(', ')})`);
      if (ind.left !== undefined) out.push(['leftIndent', num(ind.left / TWIPS_PER_POINT)]);
      if (ind.right !== undefined) out.push(['rightIndent', num(ind.right / TWIPS_PER_POINT)]);
      if (ind.hanging !== undefined) out.push(['firstLineIndent', num(-ind.hanging / TWIPS_PER_POINT)]);
      else if (ind.firstLine !== undefined) out.push(['firstLineIndent', num(ind.firstLine / TWIPS_PER_POINT)]);
    }
    const spacing = pPr.spacing;
    if (spacing) {
      if (!within(spacing, SPACING_KEYS)) throw new Unexpressible(`w:spacing (${keysOf(spacing).filter((k) => !SPACING_KEYS.has(k)).join(', ')})`);
      if (spacing.before !== undefined) out.push(['spaceBefore', num(spacing.before / TWIPS_PER_POINT)]);
      if (spacing.after !== undefined) out.push(['spaceAfter', num(spacing.after / TWIPS_PER_POINT)]);
      if (spacing.line !== undefined) {
        // Paragraph.lineSpacing writes w:lineRule="exact", so only an exact rule round-trips.
        if (spacing.lineRule !== 'exact') throw new Unexpressible(`w:spacing w:lineRule="${spacing.lineRule ?? 'auto'}"`);
        out.push(['lineSpacing', num(spacing.line / TWIPS_PER_POINT)]);
      }
    }
    const outlineLvl = pPr.outlineLvl?.val;
    if (outlineLvl !== undefined) out.push(['outlineLevel', num(outlineLvl + 1)]);
    return out;
  }

  /** A table through `insertTable` when the grid is plain text; otherwise `insertXml`. */
  private async table(element: Element<wml.Tbl>): Promise<void> {
    const grid = tableGrid(element.value);
    const hasInsertTable = typeof (Body.prototype as unknown as { insertTable?: unknown }).insertTable === 'function';
    if (!grid || !hasInsertTable) {
      await this.fallback(element, grid ? 'insertTable is CR-002 phase C' : 'the table is not a plain grid of text');
      return;
    }
    const name = `t${++this.tables}`;
    const values = `[${grid.values.map((row) => `[${row.map(quote).join(', ')}]`).join(', ')}]`;
    this.line(`const ${name} = ${this.variable}.insertTable(${grid.rows}, ${grid.columns}, '${this.location}', ${values});`);
    if (grid.style !== undefined) { const [member, value] = styleAssignment(grid.style); this.line(`${name}.${member} = ${value};`); }
  }

  /** What no verb expresses: the marshalled fragment, which `insertXml` takes back. */
  private async fallback(element: Element, reason?: string): Promise<void> {
    const xml = await marshalString(element);
    if (reason !== undefined) this.line(`// ${reason}: as XML`);
    this.line(`await ${this.variable}.insertXml(${template(xml)}, '${this.location}');`);
  }
}

function sameFormat(a: RunFormatting, b: RunFormatting): boolean {
  return (Object.keys(a) as (keyof RunFormatting)[]).every((k) => a[k] === b[k]);
}

/** The text and breaks of a paragraph's runs, in order; throws for anything a verb cannot write. */
function piecesOf(p: wml.P): TextPiece[] {
  const out: TextPiece[] = [];
  for (const item of (p.content ?? []) as Element[]) {
    const typeName = typeNameOf(item);
    if (typeName !== undefined && IGNORED_TYPES.has(typeName)) continue;
    if (typeName !== undefined && BOOKMARK_TYPES.has(typeName)) continue;
    if (typeName !== 'org_docx4j_wml.R') throw new Unexpressible(`${item.name?.localPart ?? 'content'} in the paragraph`);
    const run = item.value as wml.R;
    if (run.rPr) {
      if (!within(run.rPr, RPR_KEYS)) throw new Unexpressible(`run properties (${keysOf(run.rPr).filter((k) => !RPR_KEYS.has(k)).join(', ')})`);
      if (run.rPr.rFonts && !within(run.rPr.rFonts, RFONTS_KEYS)) throw new Unexpressible('w:rFonts beyond ascii and hAnsi');
    }
    const format = readRunOptions(run.rPr);
    for (const child of (run.content ?? []) as Element[]) {
      const local = child.name?.localPart;
      if (local === 't') {
        const text = String((child.value as wml.Text).value ?? '');
        if (text !== '') out.push({ kind: 'text', text, breakType: 'Line', format });
      } else if (local === 'tab') {
        // insertText writes a tab character, not a w:tab; the text model reads both as \t.
        out.push({ kind: 'text', text: '\t', breakType: 'Line', format });
      } else if (local === 'br') {
        const type = (child.value as wml.Br).type;
        if (type !== undefined && type !== 'page' && type !== 'textWrapping') throw new Unexpressible(`w:br type ${type}`);
        out.push({ kind: 'break', text: '', breakType: type === 'page' ? 'Page' : 'Line', format });
      } else if (local === 'lastRenderedPageBreak') {
        continue;
      } else {
        throw new Unexpressible(`w:${local ?? '?'} in a run`);
      }
    }
  }
  return out;
}

interface TableGrid {
  rows: number;
  columns: number;
  values: string[][];
  style?: string;
}

/** `styleBuiltIn` with the `Word.Style` value for a built-in style id, else `style` with the display name. */
function styleAssignment(styleId: string): [string, string] {
  const builtIn = builtInOf(styleId);
  return builtIn === 'Other' ? ['style', quote(displayNameOf(styleId))] : ['styleBuiltIn', quote(builtIn)];
}

/** A table as `insertTable` would build it, or undefined when it is more than a grid of text. */
function tableGrid(tbl: wml.Tbl): TableGrid | undefined {
  const tblPr = tbl.tblPr;
  if (tblPr && !within(tblPr, TBLPR_KEYS)) return undefined;
  const values: string[][] = [];
  for (const item of (tbl.content ?? []) as Element[]) {
    const typeName = typeNameOf(item);
    if (typeName === 'org_docx4j_wml.CTTblGridCol' || typeName === 'org_docx4j_wml.TblGrid') continue;
    if (typeName !== 'org_docx4j_wml.Tr') continue;
    const row: string[] = [];
    for (const cell of ((item.value as wml.Tr).content ?? []) as Element[]) {
      if (typeNameOf(cell) !== 'org_docx4j_wml.Tc') continue;
      const tc = cell.value as wml.Tc;
      const blocks = ((tc.content ?? []) as Element[]).filter((b) => typeNameOf(b) !== 'org_docx4j_wml.TcPr');
      if (blocks.length !== 1 || typeNameOf(blocks[0]!) !== 'org_docx4j_wml.P') return undefined;
      const paragraph = blocks[0]!.value as wml.P;
      if (paragraph.pPr) return undefined;
      for (const run of (paragraph.content ?? []) as Element[]) {
        if (typeNameOf(run) !== 'org_docx4j_wml.R') return undefined;
        if ((run.value as wml.R).rPr) return undefined;
      }
      row.push(textOf(paragraph));
    }
    values.push(row);
  }
  if (values.length === 0) return undefined;
  const columns = values[0]!.length;
  if (columns === 0 || values.some((r) => r.length !== columns)) return undefined;
  const grid: TableGrid = { rows: values.length, columns, values };
  const style = tblPr?.tblStyle?.val;
  if (style !== undefined) grid.style = style;
  return grid;
}

/**
 * The content-API script that reproduces what it is given, as TypeScript against `body`: the verbs
 * where they exist (`insertParagraph`, `style`, `alignment`, the indents and spacing, `insertText`
 * with `Range.font` per run, `insertBreak`, `insertTable`), and `insertXml` with the marshalled
 * fragment for the rest. Emitted for reveal codes, for teaching the API by example, and for an
 * agent that wants the calls rather than the tree.
 *
 * The script is a sequence of statements over `body`; it uses `await` for the `insertXml` lines,
 * so run it in an async function (`new Function('body', 'return (async () => {' + script + '})()')`).
 */
export async function toApiScript(target: ApiScriptTarget, options: ApiScriptOptions = {}): Promise<string> {
  const emitter = new Emitter(options);
  if (target instanceof Paragraph) await emitter.paragraph(target.element);
  else if (target instanceof Range) await emitter.range(target);
  else if (target instanceof Body) {
    for (const element of target.content) await emitter.element(element);
  } else if (Array.isArray(target)) {
    for (const element of target) await emitter.element(element);
  } else if (isElement(target)) {
    await emitter.element(target);
  } else {
    throw new Docx4JException('toApiScript takes a Body, a Paragraph, a Range, or w:p / w:tbl elements');
  }
  return emitter.script();
}
