// The text model of a paragraph and the block-level structure, over the objects package's
// builders (walk, find, linkParents, textOf and the element helpers come from there).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { isElement, typeNameOf, type Element } from '@docx4j/generated-objects-ts/builders/wml';

export { isElement, typeNameOf, walk, find, linkParents, textOf, W_NS, type Element } from '@docx4j/generated-objects-ts/builders/wml';

/** One text-bearing item of a run, with its position in the paragraph's text. */
export interface TextSegment {
  /** The item: w:t, w:tab, w:br, ... */
  item: Element;
  /** The run's content array holding the item, and the item's index in it. */
  owner: Element[];
  index: number;
  /** The run. */
  run: wml.R;
  /** The array holding the run and its index there. */
  runOwner: Element[];
  runIndex: number;
  text: string;
  start: number;
  end: number;
  /** True for w:t: text that can be edited in place. */
  editable: boolean;
}

/** Types whose value holds runs (directly or through sdtContent); w:del is not among them: deleted text is not the text. */
const RUN_HOLDERS = new Set([
  'org_docx4j_wml.P.Hyperlink', 'org_docx4j_wml.SdtRun', 'org_docx4j_wml.RunIns', 'org_docx4j_wml.CTSmartTagRun',
  'org_docx4j_wml.CTCustomXmlRun', 'org_docx4j_wml.P.Dir', 'org_docx4j_wml.P.Bdo', 'org_docx4j_wml.CTSimpleField',
  'org_docx4j_wml.RunTrackChange', 'org_docx4j_wml.CTSdtContentRun',
]);

/**
 * The run-level items a holder keeps, by the model's own property names: `content` for most,
 * `customXmlOrSmartTagOrSdt` for w:ins, w:del, w:moveFrom and w:moveTo (docx4j's names for
 * the choice groups), `sdtContent.content` for a run-level content control.
 */
export function runItemsOf(value: object): Element[] | undefined {
  const v = value as { content?: Element[]; customXmlOrSmartTagOrSdt?: Element[]; sdtContent?: { content?: Element[] } };
  if (Array.isArray(v.content)) return v.content;
  if (Array.isArray(v.customXmlOrSmartTagOrSdt)) return v.customXmlOrSmartTagOrSdt;
  if (v.sdtContent) return runItemsOf(v.sdtContent);
  return undefined;
}

function itemText(item: Element): string | undefined {
  const t = item.name.localPart;
  const v = item.value as Record<string, unknown>;
  switch (t) {
    case 't': return String(v.value ?? '');
    case 'tab': return '\t';
    case 'br': return '\n';
    case 'cr': return '\n';
    case 'noBreakHyphen': return '‑';
    case 'softHyphen': return '­';
    case 'sym': {
      const code = v._char ?? v.char;
      return typeof code === 'string' && /^[0-9A-Fa-f]{4}$/.test(code) ? String.fromCharCode(parseInt(code, 16)) : '';
    }
    default: return undefined;
  }
}

/** The text segments of a paragraph (or any run holder) in order. Deleted text (w:del) is skipped. */
export function segmentsOf(container: object): TextSegment[] {
  const out: TextSegment[] = [];
  let pos = 0;
  const visitRuns = (items: Element[] | undefined): void => {
    if (!items) return;
    for (let i = 0; i < items.length; i++) {
      const el = items[i]!;
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.R') {
        const run = el.value as wml.R;
        const content = run.content ?? [];
        for (let j = 0; j < content.length; j++) {
          const item = content[j]!;
          const text = itemText(item);
          if (text === undefined) continue;
          out.push({ item, owner: content, index: j, run, runOwner: items, runIndex: i, text, start: pos, end: pos + text.length, editable: item.name.localPart === 't' });
          pos += text.length;
        }
      } else if (tn !== undefined && RUN_HOLDERS.has(tn)) {
        visitRuns(runItemsOf(el.value as object));
      }
    }
  };
  visitRuns(runItemsOf(container));
  return out;
}

/** The runs of a paragraph, direct and nested in hyperlinks, content controls and insertions, in order. */
export function runsOf(container: object): Element<wml.R>[] {
  const out: Element<wml.R>[] = [];
  const visit = (items: Element[] | undefined): void => {
    if (!items) return;
    for (const el of items) {
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.R') out.push(el as Element<wml.R>);
      else if (tn !== undefined && RUN_HOLDERS.has(tn)) visit(runItemsOf(el.value as object));
    }
  };
  visit(runItemsOf(container));
  return out;
}

/**
 * The block-level children of a container: body, header, footer, cell, sdt content, table (rows),
 * row (cells). A content control of any kind (block, row, cell, run) answers with its
 * `sdtContent`'s content, so that a traversal does not have to know the four types.
 */
export function childrenOf(value: object): Element[] | undefined {
  const v = value as { content?: Element[]; sdtContent?: { content?: Element[] } };
  if (v.sdtContent) return v.sdtContent.content;
  return Array.isArray(v.content) ? v.content : undefined;
}

/** An element with the array that holds it: what a view needs to edit it in place. */
export interface Located<T = unknown> {
  element: Element<T>;
  container: Element[];
}

/** w:sdt and w:customXml at row level: they hold the rows of a table (docx4j CTSdtRow, CTCustomXmlRow). */
const ROW_HOLDERS = new Set(['org_docx4j_wml.CTSdtRow', 'org_docx4j_wml.CTCustomXmlRow']);
/** The same at cell level. */
const CELL_HOLDERS = new Set(['org_docx4j_wml.CTSdtCell', 'org_docx4j_wml.CTCustomXmlCell']);

/** The rows of a table, in order, descending into row-level content controls (an OpenDoPE repeat). */
export function rowsOf(tbl: object): Located<wml.Tr>[] {
  return located(childrenOf(tbl), 'org_docx4j_wml.Tr', ROW_HOLDERS) as Located<wml.Tr>[];
}

/** The cells of a row, in order, descending into cell-level content controls. */
export function cellsOf(tr: object): Located<wml.Tc>[] {
  return located(childrenOf(tr), 'org_docx4j_wml.Tc', CELL_HOLDERS) as Located<wml.Tc>[];
}

function located(items: Element[] | undefined, typeName: string, holders: Set<string>): Located[] {
  const out: Located[] = [];
  const visit = (list: Element[] | undefined): void => {
    if (!list) return;
    for (const element of list) {
      const tn = typeNameOf(element);
      if (tn === typeName) out.push({ element, container: list });
      else if (tn !== undefined && holders.has(tn) && typeof element.value === 'object' && element.value !== null) visit(childrenOf(element.value));
    }
  };
  visit(items);
  return out;
}

/** Content controls: `w:sdt` at block, row, cell and run level, by TYPE_NAME. */
export const SDT_TYPES = new Set([
  'org_docx4j_wml.SdtBlock', 'org_docx4j_wml.SdtRun', 'org_docx4j_wml.CTSdtRow', 'org_docx4j_wml.CTSdtCell',
]);

/** True for what a body or cell may hold directly (the Body content union), by TYPE_NAME. */
export const BLOCK_LEVEL_TYPES = new Set([
  'org_docx4j_wml.P', 'org_docx4j_wml.Tbl', 'org_docx4j_wml.SdtBlock', 'org_docx4j_wml.CTCustomXmlBlock', 'org_docx4j_wml.CTAltChunk',
  'org_docx4j_wml.ProofErr', 'org_docx4j_wml.RangePermissionStart', 'org_docx4j_wml.CTPerm', 'org_docx4j_wml.CTBookmark',
  'org_docx4j_wml.CTMarkupRange', 'org_docx4j_wml.CTMoveBookmark', 'org_docx4j_wml.CTMoveFromRangeEnd', 'org_docx4j_wml.CTMoveToRangeEnd',
  'org_docx4j_wml.CommentRangeStart', 'org_docx4j_wml.CommentRangeEnd', 'org_docx4j_wml.CTTrackChange', 'org_docx4j_wml.CTMarkup',
  'org_docx4j_wml.RunIns', 'org_docx4j_wml.RunDel', 'org_docx4j_wml.RunTrackChange', 'org_docx4j_math.CTOMathPara', 'org_docx4j_math.CTOMath',
]);

/** A run element with the given items and, optionally, run properties (the builders' r() takes text; this takes items). */
export function runOf(items: Element[], rPr?: wml.RPr): Element<wml.R> {
  const value: wml.R = { TYPE_NAME: 'org_docx4j_wml.R', content: items as wml.R['content'] };
  if (rPr) value.rPr = rPr;
  return { name: { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localPart: 'r' }, value };
}

/** A paragraph element with the given runs and, optionally, paragraph properties. */
export function paragraphOf(runs: Element[], pPr?: wml.PPr): Element<wml.P> {
  const value: wml.P = { TYPE_NAME: 'org_docx4j_wml.P', content: runs as wml.P['content'] };
  if (pPr) value.pPr = pPr;
  return { name: { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localPart: 'p' }, value };
}

export { isElement as isTypedElement };

// CR-002 phase G (comments) needs the two internals above: the run holders to walk to the
// markers inside hyperlinks and insertions, and the per-item text to give a marker its offset.
export { RUN_HOLDERS, itemText as itemTextOf };
