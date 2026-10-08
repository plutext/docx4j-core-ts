// SPDX-License-Identifier: Apache-2.0
//
// HTML to an intermediate form (CR-005 section 8; moved from docx4j-ts-editor's `view/paste.ts` on
// 2026-10-08 by its copyright owner's decision, docx4j-ts-editor ED-001 decision 26, relicensed from
// GPL-3.0-or-later; the editor imports it back). Written for Word's clipboard HTML (the editor's
// ED-003 section 8.3): `<h1>` to `<h6>` for the heading styles, `<p class=MsoNormal>` and `MsoTitle`,
// `<b>`, `<i>`, `<u>`, `<s>`, `<sup>`, `<sub>` and their `span` style forms, `<br>`, a tab as a span
// with `mso-tab-count`, lists as `<ul>` and `<ol>` or as `MsoListParagraph` paragraphs with
// `mso-list`, tables as `<table class=MsoTableGrid>`, and a great deal else; it serves any XHTML a
// data node holds for an OpenDoPE binding (the specification's section 9) the same way. The
// converter keeps what a document's styles name and the direct formatting an element states
// (alignment, indents, space before, the font, size, colour and highlight of a run, a web or mail
// link), drops the rest and counts it, so a caller can say what was dropped. `margin-bottom` and
// `line-height` stay dropped: Word writes them on every paragraph from the source's style. Pure over
// its inputs: a DOM walk to the intermediate form, which `elements.mts` turns into
// WordprocessingML. The HTML parser is the caller's: a browser's `DOMParser` by default, jsdom's or
// linkedom's in Node.

/** Word's sixteen highlight names and their colours. */
const HIGHLIGHT: Record<string, string> = {
  black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#008000', magenta: '#FF00FF', red: '#FF0000', yellow: '#FFFF00',
  white: '#FFFFFF', darkBlue: '#00008B', darkCyan: '#008B8B', darkGreen: '#006400', darkMagenta: '#8B008B', darkRed: '#8B0000',
  darkYellow: '#FFD700', darkGray: '#A9A9A9', lightGray: '#D3D3D3',
};

/** The '#RRGGBB' of one of Word's sixteen highlight names, or undefined for anything else (`mso-highlight` is read through it). */
export function highlightHexOf(name: string): string | undefined {
  return HIGHLIGHT[name];
}

/** The highlight name for one of the sixteen colours, by its '#RRGGBB'. */
export function highlightNameOf(hex: string): string | undefined {
  const upper = hex.toUpperCase();
  for (const [name, value] of Object.entries(HIGHLIGHT)) if (value === upper) return name;
  return undefined;
}

/** A web or mail address a link may keep (`http`, `https`, `mailto`), trimmed; anything else undefined. */
export function webAddress(url: string): string | undefined {
  const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(url.trim())?.[1]?.toLowerCase();
  return scheme === 'http' || scheme === 'https' || scheme === 'mailto' ? url.trim() : undefined;
}

// ---------------------------------------------------------------- the intermediate form

export interface PasteRun {
  text: string;
  bold?: true;
  italic?: true;
  underline?: true;
  strike?: true;
  vertAlign?: 'superscript' | 'subscript';
  /** The first family of `font-family`, when it is a name and not a generic. */
  fontName?: string;
  /** Points. */
  fontSize?: number;
  /** '#RRGGBB', or 'auto' for `windowtext`. */
  color?: string;
  /** '#RRGGBB' of one of Word's highlight names (`mso-highlight`, else a `background` naming one). */
  highlight?: string;
  /** A character style id of the document (Markdown's inline code, `HTMLCode`). */
  style?: string;
  /** A link's address (Markdown's links): a relationship made before the transaction, the text in Word's look. */
  link?: string;
}

export type PasteInline = PasteRun | { br: true } | { tab: true };

export interface PasteList {
  /** Word's `lfoN`, or the list element's own key: one definition per key. */
  key: string;
  level: number;
  bullet: boolean;
}

export interface PasteParagraph {
  kind: 'paragraph';
  /** A style id of the document, or null for the default. */
  style: string | null;
  list?: PasteList;
  /** `w:jc` from `text-align` or the `align` attribute. */
  jc?: string;
  /** Twips from `margin-left`, `margin-right` and `text-indent` (a negative one a hanging indent). */
  ind?: { start?: number; end?: number; firstLine?: number; hanging?: number };
  /** Twips from `margin-top`; `margin-bottom` and `line-height` are dropped and counted. */
  spacing?: { before: number };
  inlines: PasteInline[];
}

export interface PasteCell {
  blocks: PasteBlock[];
  colspan: number;
  rowspan: number;
  /** Twips, when the cell said. */
  width?: number;
}

export interface PasteTable {
  kind: 'table';
  rows: PasteCell[][];
  /** The document's table grid style, when it has one. */
  style: string | null;
  /** The first row repeated at the top of each page (`w:tblHeader`): a Markdown table's header row. */
  headerRow?: boolean;
}

export type PasteBlock = PasteParagraph | PasteTable;

export interface PasteReport {
  paragraphs: number;
  lists: number;
  tables: number;
  /** What was dropped, by kind, with how many times. */
  dropped: Record<string, number>;
  /** Built-in styles the paste names that the document lacked: a caller splices them in from the defaults (`styles.ensure`) before building, else the paragraphs stay plain. */
  stylesWanted: string[];
}

export interface PasteResult {
  blocks: PasteBlock[];
  report: PasteReport;
}

/** What the converter asks of the document: its styles, by id and by name. */
export interface StyleLookup {
  /** The style id for a name (`heading 1`, `Title`), case-insensitively, or undefined. */
  idFor(name: string): string | undefined;
  hasId(id: string): boolean;
}

/** The lookup over the document's paragraph and character styles, plus the table style ids the caller found (the styles list carries no table styles). */
/** A style's id and name, as a document's styles part lists them. */
export interface StyleNaming {
  id: string;
  name: string;
}

export function styleLookupOf(entries: readonly StyleNaming[], tableStyleIds: readonly string[] = []): StyleLookup {
  const byName = new Map(entries.map((e) => [e.name.toLowerCase(), e.id]));
  const ids = new Set([...entries.map((e) => e.id), ...tableStyleIds]);
  return { idFor: (name) => byName.get(name.toLowerCase()), hasId: (id) => ids.has(id) };
}


// ---------------------------------------------------------------- the clipboard text

/** Word's `CF_HTML` header (`Version:`, `StartHTML:` and the rest) precedes `<html` in a saved capture; a browser hands the markup alone. */
export function clipboardMarkup(text: string): string {
  // Word (and a browser handing on Word's clipboard) marks the selection: what is inside the
  // markers is the paste; the head's styles, the body tag and a capture's tail are not.
  const s = text.indexOf('<!--StartFragment-->');
  const e = text.indexOf('<!--EndFragment-->');
  if (s >= 0 && e > s) return text.slice(s + '<!--StartFragment-->'.length, e);
  const at = text.search(/<html[\s>]/i);
  if (at > 0 && text.slice(0, at).includes('StartHTML:')) return text.slice(at);
  return text.replace(/^\uFEFF/, '');
}

/** Elements that make a block of their own: HTML with any of them is a document's, not a screen's. */
const BLOCK_ELEMENTS = 'p,div,h1,h2,h3,h4,h5,h6,ul,ol,li,dl,table,pre,blockquote,hr,section,article,header,footer,main,nav,aside,figure,img';
/** A font-family a terminal or a code view would name. */
const MONOSPACE = /mono|courier|consolas|menlo|monaco/i;

/**
 * True when the HTML is a terminal's screen (ED-005 section 12.67): inline content only, every
 * piece of text in a monospace font, as Konsole's "Copy text as HTML" writes it (a span in
 * `font-family:monospace` holding a coloured span per run and a `<br>` per line,
 * `fixtures/paste/terminal/`). Its colours and font are the terminal's theme, not content, and the
 * converter would make a paragraph of each span, so the paste takes the plain text instead.
 */
export function isTerminalHtml(html: string, options: { parser?: (html: string) => Document } = {}): boolean {
  const parse = options.parser ?? ((h: string) => new DOMParser().parseFromString(h, 'text/html'));
  const doc = parse(clipboardMarkup(html));
  const body = rootOf(doc);
  if (!body || body.querySelector(BLOCK_ELEMENTS)) return false;
  const monospaced = (node: Node): boolean => {
    for (let el = node.parentElement; el && el !== body; el = el.parentElement) {
      if (MONOSPACE.test(styleOf(el).get('font-family') ?? '') || /^(code|kbd|samp|tt)$/i.test(el.tagName)) return true;
    }
    return false;
  };
  let text = false;
  const walker = doc.createTreeWalker(body, 4 /* NodeFilter.SHOW_TEXT */);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (!/\S/.test(node.textContent ?? '')) continue;
    if (!monospaced(node)) return false;
    text = true;
  }
  return text;
}


// ---------------------------------------------------------------- the walk

const HEADING = /^h([1-6])$/i;
const BULLET_GLYPHS = /^[•·●◦▪■–—o\-∙‣⁃]+$/;

type InlineState = Omit<PasteRun, 'text'>;

/** A CSS length as twips: points, centimetres, millimetres, inches and pixels (15 twips each, as the schema counts), else undefined. */
function twipsOf(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const m = /^(-?\d*\.?\d+)\s*(pt|cm|mm|in|px)?$/.exec(value.trim());
  if (!m) return undefined;
  const n = parseFloat(m[1]!);
  const per: Record<string, number> = { pt: 20, cm: 566.93, mm: 56.693, in: 1440, px: 15 };
  if (n === 0) return 0;
  const unit = m[2];
  if (!unit) return undefined;
  return Math.round(n * per[unit]!);
}

/** The CSS colour names Word writes instead of a hex value (the basic sixteen and orange), and `windowtext` as the automatic colour. */
const COLOUR_NAMES: Record<string, string> = {
  black: '#000000', silver: '#C0C0C0', gray: '#808080', grey: '#808080', white: '#FFFFFF', maroon: '#800000', red: '#FF0000', purple: '#800080', fuchsia: '#FF00FF',
  green: '#008000', lime: '#00FF00', olive: '#808000', yellow: '#FFFF00', navy: '#000080', blue: '#0000FF', teal: '#008080', aqua: '#00FFFF', orange: '#FFA500',
  windowtext: 'auto',
};

/** A CSS colour as '#RRGGBB' (or 'auto'), else undefined. */
function colourOf(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  const v = value.trim();
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(v);
  if (hex) {
    const h = hex[1]!;
    return `#${(h.length === 3 ? h.split('').map((c) => c + c).join('') : h).toUpperCase()}`;
  }
  return COLOUR_NAMES[v.toLowerCase()];
}

/** The first family of a `font-family` list, unquoted, unless it is a generic family. */
function fontNameOf(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  const first = value.split(',')[0]!.trim().replace(/^["']|["']$/g, '').trim();
  if (first === '' || /^(serif|sans-serif|monospace|cursive|fantasy|system-ui|inherit|initial)$/i.test(first)) return undefined;
  return first;
}

/** A `font-size` in points (pixels at three quarters of a point), else undefined. */
function fontSizeOf(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const m = /^(\d*\.?\d+)\s*(pt|px)$/.exec(value.trim());
  if (!m) return undefined;
  const n = parseFloat(m[1]!);
  const pt = m[2] === 'px' ? n * 0.75 : n;
  return pt > 0 ? Math.round(pt * 2) / 2 : undefined;
}

/** `w:jc` for a `text-align` value or an `align` attribute. */
function jcOf(value: string | undefined): string | undefined {
  switch ((value ?? '').trim().toLowerCase()) {
    case 'left': case 'start': return 'left';
    case 'center': return 'center';
    case 'right': case 'end': return 'right';
    case 'justify': return 'both';
    default: return undefined;
  }
}

interface Walk {
  lookup: StyleLookup;
  report: PasteReport;
  /** Distinct list keys met. */
  listKeys: Set<string>;
  /** Built-in style ids named but absent from the document. */
  wanted: Set<string>;
  /** The list element chain the walk is in: bullet or number, by `<ul>` / `<ol>`. */
  listStack: boolean[];
  tableStyle: string | null;
  /** Counts a dropped kind. */
  drop(kind: string, times?: number): void;
}

function styleOf(el: Element): Map<string, string> {
  const out = new Map<string, string>();
  const raw = el.getAttribute('style');
  if (!raw) return out;
  for (const part of raw.split(';')) {
    const i = part.indexOf(':');
    if (i < 0) continue;
    out.set(part.slice(0, i).trim().toLowerCase(), part.slice(i + 1).trim().replace(/\s+/g, ' '));
  }
  return out;
}

/** Word's built-in styles a paste may name that a document can lack: the id docx4j's defaults define them under. */
const BUILT_IN: Record<string, string> = { Title: 'Title', Subtitle: 'Subtitle', Quote: 'Quote', IntenseQuote: 'IntenseQuote', ListParagraph: 'ListParagraph', Caption: 'Caption' };

/** The style a Word paragraph names: `<hN>`, `MsoTitle`, `MsoSubtitle`, a class the document has a style for, else the default; a built-in the document lacks is wanted. */
function paragraphStyle(el: Element, w: Walk): string | null {
  const m = HEADING.exec(el.tagName);
  if (m) {
    const have = w.lookup.idFor(`heading ${m[1]}`) ?? w.lookup.idFor(`Heading ${m[1]}`) ?? (w.lookup.hasId(`Heading${m[1]}`) ? `Heading${m[1]}` : undefined);
    if (have) return have;
    w.wanted.add(`Heading${m[1]}`);
    return `Heading${m[1]}`;
  }
  const cls = el.getAttribute('class') ?? '';
  for (const c of cls.split(/\s+/)) {
    if (c === '' || c === 'MsoNormal' || c === 'MsoListParagraph' || c === 'MsoListParagraphCxSpFirst' || c === 'MsoListParagraphCxSpMiddle' || c === 'MsoListParagraphCxSpLast') continue;
    const bare = c.startsWith('Mso') ? c.slice(3) : c;
    if (w.lookup.hasId(bare)) return bare;
    const spaced = bare.replace(/([a-z])([A-Z0-9])/g, '$1 $2');
    const byName = w.lookup.idFor(spaced) ?? w.lookup.idFor(bare);
    if (byName) return byName;
    const builtIn = BUILT_IN[bare];
    if (builtIn) { w.wanted.add(builtIn); return builtIn; }
  }
  return null;
}

/** Word's `mso-list:l0 level2 lfo1` on a list paragraph. */
function listOf(el: Element, w: Walk, ignoreLabel: string | undefined): PasteList | undefined {
  const mso = styleOf(el).get('mso-list');
  const inList = w.listStack.length > 0;
  if (!mso && !inList) return undefined;
  const level = mso ? Number(/level(\d+)/.exec(mso)?.[1] ?? '1') - 1 : Math.max(0, w.listStack.length - 1);
  const lfo = mso ? /lfo(\d+)/.exec(mso)?.[1] : undefined;
  const key = lfo ? `lfo${lfo}` : `list${w.listStack.length}`;
  // Bullets: the enclosing list element says; without one, the label's glyph (a digit or a letter numbers, a symbol bullets).
  let bullet: boolean;
  if (inList) bullet = w.listStack[w.listStack.length - 1]!;
  else bullet = ignoreLabel !== undefined ? BULLET_GLYPHS.test(ignoreLabel.trim()) || ignoreLabel.trim() === '' : true;
  return { key, level: Math.max(0, Math.min(8, level)), bullet };
}

function withMark(state: InlineState, el: Element, w: Walk): InlineState {
  const tag = el.tagName.toLowerCase();
  const st = styleOf(el);
  const next: InlineState = { ...state };
  if (tag === 'b' || tag === 'strong' || /^(bold|[6-9]00)$/.test(st.get('font-weight') ?? '')) next.bold = true;
  if (tag === 'i' || tag === 'em' || st.get('font-style') === 'italic') next.italic = true;
  if (tag === 'u' || (st.get('text-decoration') ?? '').includes('underline')) next.underline = true;
  if (tag === 's' || tag === 'strike' || tag === 'del' || (st.get('text-decoration') ?? '').includes('line-through')) next.strike = true;
  if (tag === 'sup' || st.get('vertical-align') === 'super') next.vertAlign = 'superscript';
  if (tag === 'sub' || st.get('vertical-align') === 'sub') next.vertAlign = 'subscript';
  // A link to the web or to mail is kept (ED-005 section 12.67), its relationship made when the paste
  // lands; one to a place in its source (`#name`) or to a file is dropped and counted, its text kept.
  // An `<a name>` with no `href` is a bookmark, not a link.
  if (tag === 'a' && el.hasAttribute('href')) {
    const address = webAddress(el.getAttribute('href') ?? '');
    if (address !== undefined) next.link = address; else w.drop('links');
  }
  // The direct formatting the element states (section 8.6): taken where it reads, counted as dropped where it does not.
  if (st.has('font-family')) { const name = fontNameOf(st.get('font-family')); if (name) next.fontName = name; else w.drop('fonts'); }
  if (st.has('font-size')) { const size = fontSizeOf(st.get('font-size')); if (size) next.fontSize = size; else w.drop('sizes'); }
  if (st.has('color')) { const color = colourOf(st.get('color')); if (color) next.color = color; else w.drop('colours'); }
  if (st.has('mso-highlight') || st.has('background') || st.has('background-color')) {
    const highlight = highlightHexOf(st.get('mso-highlight') ?? '') ?? highlightHexOf(st.get('background') ?? st.get('background-color') ?? '');
    if (highlight) next.highlight = highlight; else w.drop('highlights');
  }
  return next;
}

const RUN_FIELDS = ['bold', 'italic', 'underline', 'strike', 'vertAlign', 'fontName', 'fontSize', 'color', 'highlight', 'style', 'link'] as const;

function pushText(inlines: PasteInline[], text: string, state: InlineState): void {
  const last = inlines[inlines.length - 1];
  // After a break a line starts afresh: its leading space is not rendered.
  if (last && 'br' in last) text = text.replace(/^ +/, '');
  if (text === '') return;
  const same = last && 'text' in last && RUN_FIELDS.every((f) => last[f] === state[f]);
  if (same) { (last as PasteRun).text += text; return; }
  inlines.push({ text, ...state });
}

/** The inline content of a paragraph-level element, whitespace collapsed as HTML renders it. */
function walkInline(node: Node, inlines: PasteInline[], state: InlineState, w: Walk, pre: boolean): void {
  if (node.nodeType === 3) {
    const raw = node.textContent ?? '';
    const text = pre ? raw : raw.replace(/[ \t\r\n]+/g, ' ');
    pushText(inlines, text, state);
    return;
  }
  if (node.nodeType !== 1) return;
  const el = node as Element;
  const tag = el.tagName.toLowerCase();
  if (tag === 'o:p' || tag === 'script' || tag === 'style') return;
  if (tag === 'br') { inlines.push({ br: true }); return; }
  if (tag === 'img') { w.drop('images'); return; }
  const st = styleOf(el);
  if (st.has('mso-tab-count')) { inlines.push({ tab: true }); return; }
  if (st.get('mso-list') === 'Ignore') return;
  const next = withMark(state, el, w);
  el.childNodes.forEach((child) => walkInline(child, inlines, next, w, pre || tag === 'pre'));
}

/** The label span Word puts before a list paragraph's text (`mso-list:Ignore`), for the bullet-or-number decision. */
function ignoreLabelOf(el: Element): string | undefined {
  for (const span of Array.from(el.querySelectorAll('span'))) {
    if (styleOf(span).get('mso-list') === 'Ignore') return span.textContent ?? '';
  }
  return undefined;
}

/** The paragraph formatting an element states (section 8.6): alignment, indents and the space before kept; `margin-bottom` and `line-height` dropped and counted. */
function paragraphFormatting(el: Element, st: Map<string, string>, w: Walk): Pick<PasteParagraph, 'jc' | 'ind' | 'spacing'> {
  const out: Pick<PasteParagraph, 'jc' | 'ind' | 'spacing'> = {};
  if (el.hasAttribute('align') || st.has('text-align')) {
    const jc = jcOf(st.get('text-align') ?? el.getAttribute('align') ?? undefined);
    if (jc) out.jc = jc; else w.drop('alignment');
  }
  if (st.has('margin-left') || st.has('text-indent') || st.has('margin-right')) {
    const ind: NonNullable<PasteParagraph['ind']> = {};
    const start = twipsOf(st.get('margin-left'));
    const end = twipsOf(st.get('margin-right'));
    const first = twipsOf(st.get('text-indent'));
    if (start !== undefined) ind.start = start;
    if (end !== undefined) ind.end = end;
    if (first !== undefined) { if (first < 0) ind.hanging = -first; else ind.firstLine = first; }
    if (Object.keys(ind).length > 0) out.ind = ind; else w.drop('indents');
  }
  if (st.has('margin-top')) {
    const before = twipsOf(st.get('margin-top'));
    if (before !== undefined) out.spacing = { before }; else w.drop('spacing');
  }
  if (st.has('line-height') || st.has('margin-bottom')) w.drop('spacing');
  return out;
}

function paragraphFrom(el: Element, w: Walk): PasteParagraph {
  const st = styleOf(el);
  const formatting = paragraphFormatting(el, st, w);
  const list = listOf(el, w, ignoreLabelOf(el));
  if (list) w.listKeys.add(list.key);
  const inlines: PasteInline[] = [];
  el.childNodes.forEach((child) => walkInline(child, inlines, {}, w, el.tagName.toLowerCase() === 'pre'));
  trimEdges(inlines);
  w.report.paragraphs++;
  return { kind: 'paragraph', style: paragraphStyle(el, w), ...(list ? { list } : {}), ...formatting, inlines };
}

/** A paragraph's leading and trailing whitespace goes, as a browser renders it. */
function trimEdges(inlines: PasteInline[]): void {
  const first = inlines[0];
  if (first && 'text' in first) { first.text = first.text.replace(/^ +/, ''); if (first.text === '') inlines.shift(); }
  const last = inlines[inlines.length - 1];
  if (last && 'text' in last) { last.text = last.text.replace(/ +$/, ''); if (last.text === '') inlines.pop(); }
}

function cellFrom(el: Element, w: Walk): PasteCell {
  const blocks: PasteBlock[] = [];
  walkBlocks(el, blocks, w);
  if (blocks.length === 0) { blocks.push({ kind: 'paragraph', style: null, inlines: [] }); w.report.paragraphs++; }
  const cell: PasteCell = { blocks, colspan: Math.max(1, Number(el.getAttribute('colspan') ?? '1') || 1), rowspan: Math.max(1, Number(el.getAttribute('rowspan') ?? '1') || 1) };
  const width = el.getAttribute('width');
  const styleWidth = styleOf(el).get('width');
  if (styleWidth && /pt$/.test(styleWidth)) cell.width = Math.round(parseFloat(styleWidth) * 20);
  else if (width && /^\d+$/.test(width)) cell.width = Number(width) * 15;
  return cell;
}

function tableFrom(el: Element, w: Walk): PasteTable {
  const rows: PasteCell[][] = [];
  const rowElements: Element[] = [];
  const collect = (parent: Element): void => {
    for (const child of Array.from(parent.children)) {
      const tag = child.tagName.toLowerCase();
      if (tag === 'tr') rowElements.push(child);
      else if (tag === 'thead' || tag === 'tbody' || tag === 'tfoot') collect(child);
    }
  };
  collect(el);
  for (const tr of rowElements) {
    const cells: PasteCell[] = [];
    for (const child of Array.from(tr.children)) {
      const tag = child.tagName.toLowerCase();
      if (tag === 'td' || tag === 'th') cells.push(cellFrom(child, w));
    }
    if (cells.length > 0) rows.push(cells);
  }
  w.report.tables++;
  return { kind: 'table', rows, style: w.tableStyle };
}

/** Block-level content in document order: paragraphs, headings, list items, tables; containers descended into. */
function walkBlocks(parent: Node, out: PasteBlock[], w: Walk): void {
  const pending: { inlines: PasteInline[] } | null = null;
  void pending;
  for (const node of Array.from(parent.childNodes)) {
    if (node.nodeType === 3) {
      // Bare text under a container (a cell without a paragraph): its own paragraph.
      const text = (node.textContent ?? '').replace(/[ \t\r\n]+/g, ' ').trim();
      if (text !== '') { out.push({ kind: 'paragraph', style: null, inlines: [{ text }] }); w.report.paragraphs++; }
      continue;
    }
    if (node.nodeType !== 1) continue;
    const el = node as Element;
    const tag = el.tagName.toLowerCase();
    if (tag === 'script' || tag === 'style' || tag === 'head' || tag === 'meta' || tag === 'link' || tag === 'title' || tag === 'o:p') continue;
    if (tag === 'p' || HEADING.test(tag) || tag === 'li' || tag === 'pre' || tag === 'blockquote' && el.querySelector('p') === null) { out.push(paragraphFrom(el, w)); continue; }
    if (tag === 'ul' || tag === 'ol') { w.listStack.push(tag === 'ul'); walkBlocks(el, out, w); w.listStack.pop(); continue; }
    if (tag === 'table') { if (w.listStack.length === 0 || true) out.push(tableFrom(el, w)); continue; }
    if (tag === 'img') { w.drop('images'); continue; }
    if (tag === 'br') continue;
    // A container (div, body, html, span at block level, blockquote with paragraphs, td outside a table): its blocks.
    walkBlocks(el, out, w);
  }
}

/**
 * Where a parsed fragment's elements are: a browser's `DOMParser` wraps a bare fragment (the text
 * between Word's fragment markers, a data node's XHTML) in `html` and `body`; linkedom and other
 * light DOMs leave the body empty and put them under the document element or the document. The
 * first of those holding an element is the root.
 */
function rootOf(doc: Document): ParentNode {
  const html = doc.documentElement && /^html$/i.test(doc.documentElement.tagName) ? doc.documentElement : null;
  for (const candidate of [doc.body, html, doc] as (ParentNode | null)[]) {
    if (candidate && candidate.children.length > 0) return candidate;
  }
  return doc.body ?? doc.documentElement ?? doc;
}

/** Convert clipboard HTML to the intermediate form, keeping what the outline names and counting the rest. */
export function convertHtml(html: string, lookup: StyleLookup, options: { parser?: (html: string) => Document } = {}): PasteResult {
  const markup = clipboardMarkup(html);
  const parse = options.parser ?? ((h: string) => new DOMParser().parseFromString(h, 'text/html'));
  const doc = parse(markup);
  const report: PasteReport = { paragraphs: 0, lists: 0, tables: 0, dropped: {}, stylesWanted: [] };
  const w: Walk = {
    lookup,
    report,
    listKeys: new Set(),
    wanted: new Set(),
    listStack: [],
    tableStyle: lookup.hasId('TableGrid') ? 'TableGrid' : null,
    drop: (kind, times = 1) => { report.dropped[kind] = (report.dropped[kind] ?? 0) + times; },
  };
  const blocks: PasteBlock[] = [];
  walkBlocks(rootOf(doc), blocks, w);
  report.lists = w.listKeys.size;
  report.stylesWanted = [...w.wanted];
  return { blocks, report };
}

// ---------------------------------------------------------------- nodes of the schema

/** The list definitions a paste needs: a `w:numId` per list key. */
