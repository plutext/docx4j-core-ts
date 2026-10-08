// SPDX-License-Identifier: Apache-2.0
//
// The intermediate form to WordprocessingML (CR-005 section 8): the engine's counterpart of the
// editor's `toNodes`, written as a `w:p` / `w:tbl` fragment that `Body.insertXml` (or `contentOf`)
// parses, so that one builder serves the processor's bind step and any caller with HTML to put
// into a document. What the form holds that needs the document first - a list's definition, a
// link's relationship - is resolved by the caller through `needsOf` and the resolvers here: a list
// without a definition is a plain paragraph, a link without a relationship keeps its text.
import type { PasteBlock, PasteCell, PasteInline, PasteParagraph, PasteRun, PasteTable } from './convert.mjs';
import { highlightNameOf } from './convert.mjs';

/** Word's character style for a link's text, as `Range.hyperlink` gives it. */
export const HYPERLINK_STYLE = 'Hyperlink';

/** What the blocks need from the document before they are built: a definition per list key, a relationship per link address. */
export interface BuildNeeds {
  lists: { key: string; bullet: boolean }[];
  links: string[];
}

/** The document's answers: `w:numId` per list key, `r:id` per address; undefined drops the list or the link (its text kept). */
export interface BuildResolvers {
  numId?: (key: string) => string | undefined;
  relId?: (address: string) => string | undefined;
}

function walkBlocks(blocks: readonly PasteBlock[], onParagraph: (p: PasteParagraph) => void): void {
  for (const b of blocks) {
    if (b.kind === 'paragraph') onParagraph(b);
    else for (const row of b.rows) for (const c of row) walkBlocks(c.blocks, onParagraph);
  }
}

export function needsOf(blocks: readonly PasteBlock[]): BuildNeeds {
  const lists = new Map<string, boolean>();
  const links = new Set<string>();
  walkBlocks(blocks, (p) => {
    if (p.list && !lists.has(p.list.key)) lists.set(p.list.key, p.list.bullet);
    for (const i of p.inlines) if ('text' in i && i.link !== undefined) links.add(i.link);
  });
  return { lists: [...lists].map(([key, bullet]) => ({ key, bullet })), links: [...links] };
}

const esc = (text: string): string => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const attr = (value: string | number): string => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

function rPrXml(run: PasteRun, linked: boolean): string {
  const out: string[] = [];
  // In the schema's order: rStyle, rFonts, b, i, strike, color, sz, szCs, highlight, u, vertAlign.
  if (linked) out.push(`<w:rStyle w:val="${HYPERLINK_STYLE}"/>`);
  else if (run.style) out.push(`<w:rStyle w:val="${attr(run.style)}"/>`);
  if (run.fontName) out.push(`<w:rFonts w:ascii="${attr(run.fontName)}" w:hAnsi="${attr(run.fontName)}" w:cs="${attr(run.fontName)}"/>`);
  if (run.bold) out.push('<w:b/>');
  if (run.italic) out.push('<w:i/>');
  if (run.strike) out.push('<w:strike/>');
  // A link's colour is its style's, as `Range.hyperlink` leaves no direct colour on a link.
  if (run.color && !linked) out.push(`<w:color w:val="${attr(run.color === 'auto' ? 'auto' : run.color.replace(/^#/, ''))}"/>`);
  if (run.fontSize !== undefined) { const half = Math.round(run.fontSize * 2); out.push(`<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`); }
  if (run.highlight) { const name = highlightNameOf(run.highlight); if (name) out.push(`<w:highlight w:val="${name}"/>`); }
  if (run.underline) out.push('<w:u w:val="single"/>');
  if (run.vertAlign) out.push(`<w:vertAlign w:val="${run.vertAlign}"/>`);
  return out.length > 0 ? `<w:rPr>${out.join('')}</w:rPr>` : '';
}

function runXml(run: PasteRun, linked: boolean): string {
  if (run.text === '') return '';
  const space = /^\s|\s$|\s\s/.test(run.text) ? ' xml:space="preserve"' : '';
  return `<w:r>${rPrXml(run, linked)}<w:t${space}>${esc(run.text)}</w:t></w:r>`;
}

function inlinesXml(inlines: readonly PasteInline[], resolve: BuildResolvers): string {
  const out: string[] = [];
  let i = 0;
  while (i < inlines.length) {
    const inline = inlines[i]!;
    if ('br' in inline) { out.push('<w:r><w:br/></w:r>'); i++; continue; }
    if ('tab' in inline) { out.push('<w:r><w:tab/></w:r>'); i++; continue; }
    const rId = inline.link !== undefined ? resolve.relId?.(inline.link) : undefined;
    if (rId === undefined) { out.push(runXml(inline, false)); i++; continue; }
    // The runs of one address under one w:hyperlink, in Word's look.
    const runs: string[] = [];
    while (i < inlines.length) {
      const next = inlines[i]!;
      if (!('text' in next) || next.link !== inline.link) break;
      runs.push(runXml(next, true));
      i++;
    }
    out.push(`<w:hyperlink r:id="${attr(rId)}">${runs.join('')}</w:hyperlink>`);
  }
  return out.join('');
}

function pPrXml(p: PasteParagraph, resolve: BuildResolvers): string {
  const out: string[] = [];
  // In the schema's order: pStyle, numPr, spacing, ind, jc.
  if (p.style) out.push(`<w:pStyle w:val="${attr(p.style)}"/>`);
  const numId = p.list ? resolve.numId?.(p.list.key) : undefined;
  if (p.list && numId !== undefined) out.push(`<w:numPr><w:ilvl w:val="${p.list.level}"/><w:numId w:val="${attr(numId)}"/></w:numPr>`);
  if (p.spacing) out.push(`<w:spacing w:before="${p.spacing.before}"/>`);
  if (p.ind) {
    const a: string[] = [];
    if (p.ind.start !== undefined) a.push(`w:left="${p.ind.start}"`);
    if (p.ind.end !== undefined) a.push(`w:right="${p.ind.end}"`);
    if (p.ind.hanging !== undefined) a.push(`w:hanging="${p.ind.hanging}"`);
    else if (p.ind.firstLine !== undefined) a.push(`w:firstLine="${p.ind.firstLine}"`);
    if (a.length > 0) out.push(`<w:ind ${a.join(' ')}/>`);
  }
  if (p.jc) out.push(`<w:jc w:val="${attr(p.jc)}"/>`);
  return out.length > 0 ? `<w:pPr>${out.join('')}</w:pPr>` : '';
}

function paragraphXml(p: PasteParagraph, resolve: BuildResolvers): string {
  return `<w:p>${pPrXml(p, resolve)}${inlinesXml(p.inlines, resolve)}</w:p>`;
}

/** Word's default text width shared equally where the table does not say (A4 or Letter with Word's margins: 9,026 twips). */
const TEXT_WIDTH = 9026;

function tableXml(t: PasteTable, resolve: BuildResolvers): string {
  const columns = Math.max(1, ...t.rows.map((r) => r.reduce((n, c) => n + c.colspan, 0)));
  const widths: number[] = Array.from({ length: columns }, () => Math.floor(TEXT_WIDTH / columns));
  const first = t.rows[0];
  if (first && first.every((c) => c.width !== undefined && c.colspan === 1) && first.length === columns) first.forEach((c, i) => { widths[i] = c.width!; });
  const widthOf = (col: number, span: number): number => { let w = 0; for (let k = 0; k < span; k++) w += widths[col + k] ?? widths[0]!; return w; };
  const tcPr = (col: number, span: number, merge: 'restart' | 'continue' | undefined): string =>
    `<w:tcPr><w:tcW w:w="${widthOf(col, span)}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}${merge === 'restart' ? '<w:vMerge w:val="restart"/>' : merge === 'continue' ? '<w:vMerge/>' : ''}</w:tcPr>`;
  const cellXml = (c: PasteCell, col: number): string => {
    const blocks = c.blocks.length > 0 ? blocksToXml(c.blocks, resolve) : '<w:p/>';
    return `<w:tc>${tcPr(col, c.colspan, c.rowspan > 1 ? 'restart' : undefined)}${blocks}</w:tc>`;
  };
  // A cell spanning rows holds its place in the rows below with a continuation cell, as WordprocessingML writes a vertical merge.
  const pending = new Map<number, { rows: number; span: number }>();
  const rows = t.rows.map((cells, r) => {
    const out: string[] = [];
    let col = 0;
    const continuation = (): void => {
      const p = pending.get(col);
      if (!p) return;
      out.push(`<w:tc>${tcPr(col, p.span, 'continue')}<w:p/></w:tc>`);
      if (--p.rows <= 0) pending.delete(col);
      col += p.span;
    };
    for (const c of cells) {
      while (pending.has(col)) continuation();
      out.push(cellXml(c, col));
      if (c.rowspan > 1) pending.set(col, { rows: c.rowspan - 1, span: c.colspan });
      col += c.colspan;
    }
    while (col < columns && pending.has(col)) continuation();
    const trPr = r === 0 && t.headerRow ? '<w:trPr><w:tblHeader/></w:trPr>' : '';
    return `<w:tr>${trPr}${out.join('')}</w:tr>`;
  });
  const tblPr = `<w:tblPr>${t.style ? `<w:tblStyle w:val="${attr(t.style)}"/>` : ''}<w:tblW w:w="0" w:type="auto"/></w:tblPr>`;
  const grid = `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
  return `<w:tbl>${tblPr}${grid}${rows.join('')}</w:tbl>`;
}

/** The blocks as a WordprocessingML fragment: `w:p` and `w:tbl` elements, the prefixes undeclared (`insertXml` declares them). */
export function blocksToXml(blocks: readonly PasteBlock[], resolve: BuildResolvers = {}): string {
  return blocks.map((b) => (b.kind === 'table' ? tableXml(b, resolve) : paragraphXml(b, resolve))).join('');
}
