// The text model of a paragraph and the block-level structure, over the objects package's
// builders (walk, find, linkParents, textOf and the element helpers come from there).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import { isElement, typeNameOf, runItemsOf, mcBranchOf, type Element } from '@docx4j/generated-objects-ts/builders/wml';
import { READ_UNDERSTOOD_PREFIXES } from '../../opc/mce/understood.mjs';

export { isElement, typeNameOf, walk, find, linkParents, textOf, W_NS, type Element } from '@docx4j/generated-objects-ts/builders/wml';

/** The two views of a tracked document: as it will read once accepted, or as it read before the changes. */
export type TextView = 'accepted' | 'original';

export interface TextViewOptions {
  /** 'accepted' (the default: w:ins included, w:del excluded) or 'original' (the other way). */
  view?: TextView;
}

/** `segmentsOf`'s options: a view, and whether drawn objects get a segment of their own. */
export interface SegmentOptions extends TextViewOptions {
  /**
   * True to list each picture or other drawn object (`OBJECT_ITEMS`) as a segment of no text at its
   * place in the text (`object` set), so that a deletion of a span can find it (CR-002 section 48).
   */
  objects?: boolean;
}

/**
 * Run content that stands in the text and has no text of its own: a `w:drawing` (a picture, a chart, a
 * text box), a VML `w:pict` and an embedded `w:object`. Office JS counts an inline picture as in a
 * range, so a span's deletion takes the objects in it (CR-002 section 48). A kept `mc:AlternateContent`
 * in a run is one such object too, whatever its branches hold (`ALTERNATE_CONTENT`, CR-010 phase 3).
 */
export const OBJECT_ITEMS: ReadonlySet<string> = new Set(['drawing', 'pict', 'object']);

/** The TYPE_NAME of a kept `mc:AlternateContent` (the default load since CR-010 phase 4; under `KEPT_HOSTS` with `mcePreprocess: true`). */
export const ALTERNATE_CONTENT = 'org_docx4j_mce.AlternateContent';

/**
 * The branch of a kept `mc:AlternateContent` the content API reads (docx4j's `McMode.READ`, CR-010
 * phase 3): the objects package's `mcBranchOf` over `READ_UNDERSTOOD_PREFIXES`, the prefixes load
 * understands and the four Word 365 understands that load leaves alone (CR-010 section 15). So a reader
 * takes the branch the preprocessor would have resolved the element to, unless a Choice requires one of
 * those four: then it takes that Choice, as Word does, where a resolved load takes the next one or the
 * Fallback. A mutator walks every branch instead (docx4j's `McMode.ALL`). Undefined for an element with
 * no branches.
 */
export function selectedBranchOf(value: mce.AlternateContent | undefined): Element[] | undefined {
  return mcBranchOf(value, { understood: READ_UNDERSTOOD_PREFIXES });
}

/** True for a kept `mc:AlternateContent`: its element pair, or its value. */
export function isAlternateContent(value: unknown): boolean {
  const v = isElement(value) ? value.value : value;
  return typeof v === 'object' && v !== null && (v as { TYPE_NAME?: string }).TYPE_NAME === ALTERNATE_CONTENT;
}

/**
 * Which objects at a span's edges are in the span. An object stands between two characters, so one at
 * a span's start or end offset may be inside it or outside it, and offsets alone cannot say which. A
 * paragraph's whole range holds the objects at both of its edges (`WHOLE`); a span of text, such as a
 * search found, holds only the objects strictly inside it (CR-002 section 48).
 */
export interface SpanEdges { start: boolean; end: boolean }

/** Both edges: the objects at a span's start and at its end are in it. */
export const WHOLE: SpanEdges = Object.freeze({ start: true, end: true });

/** True when an object at offset `at` is in the span [start, end) with these edges. */
export function objectInSpan(at: number, start: number, end: number, edges?: SpanEdges): boolean {
  return (start < at && at < end) || (at === start && edges?.start === true) || (at === end && edges?.end === true);
}

/** The four run-level revision holders (CR-002 phase F); their runs live under `customXmlOrSmartTagOrSdt`. */
export type RevisionKind = 'ins' | 'del' | 'moveFrom' | 'moveTo';

/** The run-level revision a run sits in. */
export interface RevisionHolder {
  kind: RevisionKind;
  /** The `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` element pair. */
  element: Element;
  value: wml.CTTrackChange;
  /** Its run items (the model's `customXmlOrSmartTagOrSdt`). */
  items: Element[];
  /** The array holding `element`. */
  owner: Element[];
}

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
  /** True for a drawn object (`OBJECT_ITEMS`, or a kept `mc:AlternateContent`), whose segment has no text: listed only when asked for. */
  object?: boolean;
  /**
   * The nearest enclosing w:ins / w:del / w:moveFrom / w:moveTo, when the run is in one.
   * A deletion of this run goes inside it (`runOwner` is already its item list); an insertion
   * next to it goes beside it (`revision.owner`) unless the same author may extend it.
   */
  revision?: RevisionHolder;
}

/** Types whose value holds runs (directly or through sdtContent); the revision holders are separate. */
const RUN_HOLDERS = new Set([
  'org_docx4j_wml.P.Hyperlink', 'org_docx4j_wml.SdtRun', 'org_docx4j_wml.CTSmartTagRun',
  'org_docx4j_wml.CTCustomXmlRun', 'org_docx4j_wml.P.Dir', 'org_docx4j_wml.P.Bdo', 'org_docx4j_wml.CTSimpleField',
  'org_docx4j_wml.CTSdtContentRun',
]);

/** w:ins, w:del and the two w:move forms (w:moveFrom and w:moveTo share docx4j's RunTrackChange). */
const REVISION_HOLDERS = new Set(['org_docx4j_wml.RunIns', 'org_docx4j_wml.RunDel', 'org_docx4j_wml.RunTrackChange']);

/** The kind of a run-level revision element, by its name; undefined when it is not one. */
export function revisionKindOf(el: Element): RevisionKind | undefined {
  if (!REVISION_HOLDERS.has(typeNameOf(el) ?? '')) return undefined;
  const name = el.name.localPart;
  return name === 'ins' || name === 'del' || name === 'moveFrom' || name === 'moveTo' ? name : undefined;
}

/** True when this view shows the runs of a revision of that kind. */
function shows(view: TextView, kind: RevisionKind): boolean {
  return view === 'accepted' ? kind === 'ins' || kind === 'moveTo' : kind === 'del' || kind === 'moveFrom';
}

export { runItemsOf };

function itemText(item: Element): string | undefined {
  const t = item.name.localPart;
  const v = item.value as Record<string, unknown>;
  switch (t) {
    case 't': case 'delText': return String(v.value ?? '');
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

/**
 * The text segments of a paragraph (or any run holder) in order. The default view is the
 * accepted one: the runs of a `w:ins` or `w:moveTo` count, those of a `w:del` or `w:moveFrom`
 * do not. `{ view: 'original' }` is the other way round (CR-002 phase F). `{ objects: true }`
 * lists the drawn objects too, each a segment of no text where it stands (CR-002 section 48).
 *
 * A kept `mc:AlternateContent` in a run is one object segment, whatever its branches hold, so that
 * removing a span removes every branch with it (CR-010 phase 3). A drawing adds no text, so the
 * offsets are a resolved load's. One in a paragraph's content, outside a run, is not listed: Word
 * writes none, and the corpus has none (CR-010 section 11.3).
 */
export function segmentsOf(container: object, options?: SegmentOptions): TextSegment[] {
  const view = options?.view ?? 'accepted';
  const objects = options?.objects === true;
  const out: TextSegment[] = [];
  let pos = 0;
  const visitRuns = (items: Element[] | undefined, revision: RevisionHolder | undefined): void => {
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
          if (text === undefined) {
            if (objects && (OBJECT_ITEMS.has(item.name.localPart) || isAlternateContent(item))) {
              const seg: TextSegment = { item, owner: content, index: j, run, runOwner: items, runIndex: i, text: '', start: pos, end: pos, editable: false, object: true };
              if (revision) seg.revision = revision;
              out.push(seg);
            }
            continue;
          }
          const seg: TextSegment = {
            item, owner: content, index: j, run, runOwner: items, runIndex: i, text,
            start: pos, end: pos + text.length, editable: item.name.localPart === 't',
          };
          if (revision) seg.revision = revision;
          out.push(seg);
          pos += text.length;
        }
        continue;
      }
      const kind = revisionKindOf(el);
      if (kind !== undefined) {
        if (!shows(view, kind)) continue;
        const nested = runItemsOf(el.value as object);
        if (nested) visitRuns(nested, { kind, element: el, value: el.value as wml.CTTrackChange, items: nested, owner: items });
      } else if (tn !== undefined && RUN_HOLDERS.has(tn)) {
        visitRuns(runItemsOf(el.value as object), revision);
      }
    }
  };
  visitRuns(runItemsOf(container), undefined);
  return out;
}

/** The text of a paragraph or container in one of the two views (`textOf` is the accepted one). */
export function textOfView(container: object, options?: TextViewOptions): string {
  return segmentsOf(container, options).map((s) => s.text).join('');
}

/** The runs of a paragraph, direct and nested in hyperlinks, content controls and insertions, in order. */
export function runsOf(container: object, options?: TextViewOptions): Element<wml.R>[] {
  const view = options?.view ?? 'accepted';
  const out: Element<wml.R>[] = [];
  const visit = (items: Element[] | undefined): void => {
    if (!items) return;
    for (const el of items) {
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.R') { out.push(el as Element<wml.R>); continue; }
      const kind = revisionKindOf(el);
      if (kind !== undefined) { if (shows(view, kind)) visit(runItemsOf(el.value as object)); }
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
