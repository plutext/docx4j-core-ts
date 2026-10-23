// The text model of a paragraph and the block-level structure, over the objects package's
// builders (walk, find, linkParents, textOf and the element helpers come from there).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import { isElement, typeNameOf, runItemsOf, mcBranchOf, sdtProperty, type Element } from '@docx4j/generated-objects-ts/builders/wml';
import { READ_UNDERSTOOD_PREFIXES } from '../../opc/mce/understood.mjs';
import { isMath, mathTextOf } from './mathText.mjs';
import { NotSupportedError } from '../../office-js/errors.mjs';

export { isElement, typeNameOf, walk, find, linkParents, textOf, W_NS, type Element } from '@docx4j/generated-objects-ts/builders/wml';

/** The two views of a tracked document: as it will read once accepted, or as it read before the changes. */
export type TextView = 'accepted' | 'original';

export interface TextViewOptions {
  /** 'accepted' (the default: w:ins included, w:del excluded) or 'original' (the other way). */
  view?: TextView;
}

/** Whether a run is hidden: its effective `w:vanish` (`hiddenRunsOf` makes one for a paragraph). */
export type HiddenRuns = (run: wml.R) => boolean;

/** `segmentsOf`'s options: a view, whether drawn objects get a segment of their own, which runs are hidden. */
export interface SegmentOptions extends TextViewOptions {
  /**
   * True to list each picture or other drawn object (`OBJECT_ITEMS`) as a segment of no text at its
   * place in the text (`object` set), so that a deletion of a span can find it (CR-002 section 48).
   * Hidden runs' items are listed with them.
   */
  objects?: boolean;
  /**
   * Which runs are hidden. Office JS's text leaves hidden text out, a run's own `w:vanish` or its
   * styles' (check 48b), so `Paragraph.text` does too (CR-012 phase 4): a hidden run's items take no
   * offset, and are listed only with the objects, each `object` and `hidden`, so that a span's deletion
   * takes the hidden text strictly inside it as it takes a picture. Undefined: no run is hidden.
   */
  hidden?: HiddenRuns;
  /**
   * Called for every element the walk passes that adds no text: each element of a run holder's list
   * (runs, revision holders, markers, before it is entered) and each item of a run that adds none (a
   * comment reference, a field character, a drawing, a hidden run's text), with the offset it stands at,
   * the list holding it and its run. What gives a comment marker or a revision its offset in the same text.
   */
  visit?: (item: Element, at: number, owner: Element[], run: { element: Element<wml.R>; owner: Element[] } | undefined) => void;
  /**
   * `office` (the default): Office JS's characters (`officeItemText`), a kept `mc:AlternateContent`'s
   * selected branch read, an equation's drawn text. `textOf`: the characters of the objects package's
   * `textOf` and docx4j's `TextUtils`, as this function read before CR-012 phase 4 (a break `\n`, the
   * hyphens U+2011 and U+00AD, a symbol its code point, a note reference nothing), a run-level element
   * one object, a paragraph-level one and an equation nothing: what a tracked change's `text` is, which
   * the revisions goldens hold to docx4j's.
   */
  characters?: 'office' | 'textOf';
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
  /**
   * An item of a hidden run (`SegmentOptions.hidden`): no text, listed with the objects, `object` set too;
   * a drawing in a hidden run is one (Word's default view doesn't show it: `search`'s `^g` passes it by).
   */
  hidden?: boolean;
  /**
   * An equation (`m:oMath`, `m:oMathPara`): one segment holding its text as Word draws it
   * (`mathTextOf`). `item` is the equation and `owner` the list holding it, as `runOwner` is; `run`
   * is a frozen stand-in, since an equation is in no run. An edit may take it whole, not in part.
   */
  math?: boolean;
  /**
   * The kept `mc:AlternateContent` whose selected branch holds this item (a run-level element) or its
   * run (a paragraph-level one), the innermost: an edit lands in that branch, the others kept as they
   * were read (CR-012 section 12.2).
   */
  alternate?: AlternateLocation;
}

/** Where a kept `mc:AlternateContent` whose selected branch holds text stands (CR-012 section 12.2). */
export interface AlternateLocation {
  /** The element pair. */
  element: Element;
  /** The array holding it: a run's content (`run` level) or a run holder's list (`paragraph` level). */
  owner: Element[];
  /** Its selected branch's items (`selectedBranchOf`). */
  branch: Element[];
  /** `run`: the element is run content and its branch holds run items; `paragraph`: it stands among runs and its branch holds runs. */
  level: 'run' | 'paragraph';
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

/**
 * An item's text in Office JS's text, Word's own character stream (CR-012 phase 4; check 48, Word 365):
 * a `w:br` U+000B, of type page U+000C and column U+000E; a `w:cr` U+000D; a `w:ptab` `0`; the
 * non-breaking and optional hyphens U+001E and U+001F; any `w:sym` `(`, whatever its font or code
 * (check 48b); a footnote or endnote reference U+0002, and so the note's own mark (`w:footnoteRef`,
 * `w:endnoteRef`, not measured). Undefined for an item that adds no text: a comment reference, a
 * field's characters and code, a drawing.
 */
export function officeItemText(item: Element): string | undefined {
  const v = item.value as Record<string, unknown>;
  switch (item.name.localPart) {
    case 't': case 'delText': return String(v.value ?? '');
    case 'tab': return '\t';
    case 'br': return v.type === 'page' ? '\f' : v.type === 'column' ? '\u000E' : '\u000B';
    case 'cr': return '\r';
    case 'ptab': return '0';
    case 'noBreakHyphen': return '\u001E';
    case 'softHyphen': return '\u001F';
    case 'sym': return '(';
    case 'footnoteReference': case 'endnoteReference': case 'footnoteRef': case 'endnoteRef': return '\u0002';
    default: return undefined;
  }
}

/** An item a run-level element's selected branch may hold that makes the element one object (CR-012 section 12.2 item 5). */
function isDrawn(item: Element): boolean {
  return OBJECT_ITEMS.has(item.name.localPart);
}

/** The `run` of an equation's segment: it is in no run, and nothing may be written to this one. */
function standIn(math: Element): wml.R {
  return Object.freeze({ TYPE_NAME: 'org_docx4j_wml.R', PARENT: (math.value as { PARENT?: object }).PARENT }) as unknown as wml.R;
}

/**
 * The item text the revision listings and the text model's plan count (`itemTextOf`): whether a run
 * holds text, in the characters `textOf` gives. Not Office JS's (`officeItemText`).
 */
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
 * The text segments of a paragraph (or any run holder) in order, in Office JS's characters
 * (`officeItemText`). The default view is the accepted one: the runs of a `w:ins` or `w:moveTo`
 * count, those of a `w:del` or `w:moveFrom` do not. `{ view: 'original' }` is the other way round
 * (CR-002 phase F). `{ objects: true }` lists the drawn objects too, each a segment of no text where
 * it stands (CR-002 section 48), and the items of hidden runs (`hidden`).
 *
 * A kept `mc:AlternateContent` reads as its selected branch (CR-012 section 12.2, ED-009): in a run,
 * a branch holding run items reads as those items, at the run's offsets, and one holding a drawing,
 * a picture or an object is one object segment, as a drawing is (CR-010 phase 3); among a paragraph's
 * runs, it reads as the runs its branch holds. Either way the segments carry `alternate`, so that an
 * edit lands in that branch. An equation is one segment of its drawn text (`math`).
 */
export function segmentsOf(container: object, options?: SegmentOptions): TextSegment[] {
  const view = options?.view ?? 'accepted';
  const objects = options?.objects === true;
  const hiddenRun = options?.hidden;
  const visit = options?.visit;
  const office = options?.characters !== 'textOf';
  const charactersOf = office ? officeItemText : itemText;
  const out: TextSegment[] = [];
  let pos = 0;
  const runContent = (run: wml.R, runEl: Element<wml.R>, content: Element[], runOwner: Element[], runIndex: number,
    revision: RevisionHolder | undefined, alternate: AlternateLocation | undefined, hidden: boolean): void => {
    for (let j = 0; j < content.length; j++) {
      const item = content[j]!;
      const where = { owner: content, index: j, run, runOwner, runIndex };
      if (office && isAlternateContent(item)) {
        const branch = selectedBranchOf(item.value as mce.AlternateContent);
        // a branch holding content the model doesn't bind (DOM) is one object, as a drawing is
        if (branch && branch.every(isElement) && !branch.some(isDrawn)) {
          runContent(run, runEl, branch, runOwner, runIndex, revision, { element: item, owner: content, branch, level: 'run' }, hidden);
          continue;
        }
      }
      const text = charactersOf(item);
      if (text === undefined || (hidden && text !== '')) {
        const drawn = OBJECT_ITEMS.has(item.name.localPart) || isAlternateContent(item);
        if (objects && (drawn || (hidden && text !== undefined))) {
          const seg: TextSegment = { item, ...where, text: '', start: pos, end: pos, editable: false, object: true };
          if (hidden) seg.hidden = true;
          if (revision) seg.revision = revision;
          if (alternate) seg.alternate = alternate;
          out.push(seg);
        }
        visit?.(item, pos, content, { element: runEl, owner: runOwner });
        continue;
      }
      const seg: TextSegment = { item, ...where, text, start: pos, end: pos + text.length, editable: item.name.localPart === 't' };
      if (revision) seg.revision = revision;
      if (alternate) seg.alternate = alternate;
      out.push(seg);
      pos += text.length;
    }
  };
  const visitRuns = (items: Element[] | undefined, revision: RevisionHolder | undefined, alternate: AlternateLocation | undefined): void => {
    if (!items) return;
    for (let i = 0; i < items.length; i++) {
      const el = items[i]!;
      if (!isElement(el)) continue;                                   // a kept branch's unbound content (DOM)
      visit?.(el, pos, items, undefined);
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.R') {
        const run = el.value as wml.R;
        runContent(run, el as Element<wml.R>, run.content ?? [], items, i, revision, alternate, hiddenRun?.(run) === true);
        continue;
      }
      if (!office) {
        // textOf's reader: no branch among the runs, no equation
      } else if (isAlternateContent(el)) {
        const branch = selectedBranchOf(el.value as mce.AlternateContent);
        if (branch) visitRuns(branch, revision, { element: el, owner: items, branch, level: 'paragraph' });
        continue;
      } else if (isMath(el)) {
        const text = mathTextOf(el.value, view);
        if (text === '') continue;
        const seg: TextSegment = {
          item: el, owner: items, index: i, run: standIn(el), runOwner: items, runIndex: i, text,
          start: pos, end: pos + text.length, editable: false, math: true,
        };
        if (revision) seg.revision = revision;
        if (alternate) seg.alternate = alternate;
        out.push(seg);
        pos += text.length;
        continue;
      }
      const kind = revisionKindOf(el);
      if (kind !== undefined) {
        if (!shows(view, kind)) continue;
        const nested = runItemsOf(el.value as object);
        if (nested) visitRuns(nested, { kind, element: el, value: el.value as wml.CTTrackChange, items: nested, owner: items }, alternate);
      } else if (tn !== undefined && RUN_HOLDERS.has(tn)) {
        visitRuns(runItemsOf(el.value as object), revision, alternate);
      }
    }
  };
  visitRuns(runItemsOf(container), undefined, undefined);
  return out;
}

/** What `hiddenRunsOf` asks of a `PropertyResolver`. */
export interface VanishResolver {
  isVanishedByStyles(paragraphStyleId: string | undefined, runStyleId: string | undefined): boolean;
}

/**
 * Which of a paragraph's runs are hidden, as the text model's rule has it (REQ-507) and Office JS's text
 * leaves them out (check 48b): a run's own `w:vanish`, else its paragraph style's and character style's.
 * Without a resolver only a run's own counts.
 */
export function hiddenRunsOf(resolver: VanishResolver | undefined, p: wml.P): HiddenRuns {
  const pStyle = p.pPr?.pStyle?.val;
  return (run) => {
    if (inShowingPlaceholder(run)) return true;
    const own = run.rPr?.vanish;
    if (own !== undefined && own !== null) return own.val !== false;
    return resolver?.isVanishedByStyles(pStyle, run.rPr?.rStyle?.val) ?? false;
  };
}

/**
 * Whether a run is in a content control showing its placeholder (`w:showingPlcHdr`), run-level or block-level, up its
 * `PARENT`s. Office JS reads a placeholder as no text, and finds no text there to keep (check 53 cases 34 and 35, check 58
 * cases 11 and 12; CR-012 section 18), so the readers take its runs as they take hidden ones.
 */
export function inShowingPlaceholder(run: object): boolean {
  let o = (run as { PARENT?: object }).PARENT;
  for (let guard = 0; o && guard < 32; guard++, o = (o as { PARENT?: object }).PARENT) {
    const v = o as { sdtPr?: wml.SdtPr; sdtContent?: unknown };
    if (v.sdtPr === undefined || v.sdtContent === undefined) continue;
    const item = sdtProperty(v.sdtPr, 'showingPlcHdr');
    if (item !== undefined && (item.value as wml.BooleanDefaultTrue | undefined)?.val !== false) return true;
  }
  return false;
}

/** The text of a paragraph or container in one of the two views, in Office JS's characters (`segmentsOf`). */
export function textOfView(container: object, options?: SegmentOptions): string {
  return segmentsOf(container, options).map((s) => s.text).join('');
}

/**
 * The runs of a paragraph, direct and nested in hyperlinks, content controls and insertions, in order,
 * and in the selected branch of a kept `mc:AlternateContent` among them (CR-012 section 12.2).
 */
export function runsOf(container: object, options?: TextViewOptions): Element<wml.R>[] {
  const view = options?.view ?? 'accepted';
  const out: Element<wml.R>[] = [];
  const visit = (items: Element[] | undefined): void => {
    if (!items) return;
    for (const el of items) {
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.R') { out.push(el as Element<wml.R>); continue; }
      if (isAlternateContent(el)) { visit(selectedBranchOf(el.value as mce.AlternateContent)); continue; }
      const kind = revisionKindOf(el);
      if (kind !== undefined) { if (shows(view, kind)) visit(runItemsOf(el.value as object)); }
      else if (tn !== undefined && RUN_HOLDERS.has(tn)) visit(runItemsOf(el.value as object));
    }
  };
  visit(runItemsOf(container));
  return out;
}

/** The TYPE_NAMEs of a kept `mc:AlternateContent`'s branches, whose items are their `any`. */
export const BRANCH_TYPES: ReadonlySet<string> = new Set(['org_docx4j_mce.AlternateContent.Choice', 'org_docx4j_mce.AlternateContent.Fallback']);

/**
 * How a walk takes a kept `mc:AlternateContent` among blocks, rows or cells (CR-013): `'read'`, a reader's,
 * its selected branch (`selectedBranchOf`), as Word reads it; `'every'`, a mutator's, every branch, Choices
 * first (phase 3). With no mode the element is stepped over, as every walk did before CR-013 phase 2.
 */
export type BranchMode = 'read' | 'every';

/**
 * The block-level children of a container: body, header, footer, cell, sdt content, table (rows),
 * row (cells). A content control of any kind (block, row, cell, run) answers with its
 * `sdtContent`'s content, so that a traversal does not have to know the four types. A branch of a kept
 * `mc:AlternateContent` (a Choice or the Fallback) answers with its items, so that a locator climbing
 * `PARENT` from a branch's paragraph, row or cell finds the list it is in; the element itself answers
 * in a reader's walk (`'read'`) with its selected branch (CR-013 phase 2), and otherwise not at all.
 */
export function childrenOf(value: object, mode?: 'read'): Element[] | undefined {
  const v = value as { content?: Element[]; sdtContent?: { content?: Element[] }; any?: Element[]; TYPE_NAME?: string };
  if (v.sdtContent) return v.sdtContent.content;
  if (Array.isArray(v.content)) return v.content;
  const tn = v.TYPE_NAME;
  if (tn === undefined) return undefined;
  if (BRANCH_TYPES.has(tn)) return v.any;
  if (mode === 'read' && tn === ALTERNATE_CONTENT) return selectedBranchOf(v as mce.AlternateContent);
  return undefined;
}

/** Every branch of a kept `mc:AlternateContent`, the Choices in order, then the Fallback. */
export function branchesOf(value: mce.AlternateContent): Element[][] {
  const out: Element[][] = [];
  for (const choice of value.choice ?? []) if (choice.any) out.push(choice.any as Element[]);
  if (value.fallback?.any) out.push(value.fallback.any as Element[]);
  return out;
}

/** How many bulk edits are writing in every branch now (`writingEveryBranch`). */
let everyBranch = 0;

/**
 * Runs a bulk edit that reaches every branch of a kept `mc:AlternateContent` (CR-013 phase 3): a
 * `replaceText`, a tracked clear or row deletion, the binding passes. Inside it the views it makes of
 * what is in a branch write, where a caller's view refuses (D4): the edit itself reaches every branch,
 * the selected one and the others alike, as CR-010's rule has mutators do.
 */
export function writingEveryBranch<T>(edit: () => T): T {
  everyBranch++;
  try {
    return edit();
  } finally {
    everyBranch--;
  }
}

/**
 * Refuses a write through a view in a branch of a kept `mc:AlternateContent` at block, row or cell level
 * (CR-013 D4): the view reads the selected branch, and a write would have to reach the element's other
 * branches too, which nothing maps. A bulk edit reaching every branch writes there (`writingEveryBranch`).
 * `className` is Office JS's ('Paragraph', 'Body').
 */
export function assertOutsideBranch(value: object | undefined, className: string, what = `this ${className.toLowerCase()}`): void {
  if (everyBranch === 0 && inBranch(value)) {
    throw new NotSupportedError(className, 'edit', `${what} is in a branch of an mc:AlternateContent, which is read-only (CR-013 D4): `
      + "a write would have to reach the element's other branches too, which nothing maps yet");
  }
}

/**
 * True for a value inside a branch of a kept `mc:AlternateContent` at block, row or cell level: a
 * paragraph, table, row, cell or content control a reader sees in the selected branch (CR-013 phase 2).
 * A view of one is read-only (D4). From `PARENT` links; a run-level element's content is not reached
 * from a paragraph, which is outside it.
 */
export function inBranch(value: object | undefined): boolean {
  let o = (value as { PARENT?: object } | undefined)?.PARENT;
  for (let guard = 0; o && guard < 256; guard++, o = (o as { PARENT?: object }).PARENT) {
    const tn = (o as { TYPE_NAME?: string }).TYPE_NAME;
    if (tn !== undefined && BRANCH_TYPES.has(tn)) return true;
  }
  return false;
}

/**
 * True for a value no reader sees: inside a branch of a kept `mc:AlternateContent` other than the one
 * `selectedBranchOf` takes, at any depth (CR-013 phase 3). What a bulk edit writes there it does not
 * count, as resolve-all counts the listed changes: the count is of what a reader sees.
 */
export function inUnreadBranch(value: object | undefined): boolean {
  let o = (value as { PARENT?: object } | undefined)?.PARENT;
  for (let guard = 0; o && guard < 256; guard++, o = (o as { PARENT?: object }).PARENT) {
    const tn = (o as { TYPE_NAME?: string }).TYPE_NAME;
    if (tn === undefined || !BRANCH_TYPES.has(tn)) continue;
    const ac = (o as { PARENT?: object }).PARENT as mce.AlternateContent | undefined;
    if (ac && selectedBranchOf(ac) !== (o as { any?: Element[] }).any) return true;
  }
  return false;
}

/**
 * The lists a walk takes for a kept `mc:AlternateContent` among blocks, rows or cells: a reader's
 * (`'read'`) its selected branch, a mutator's (`'every'`) every branch, the Choices first.
 */
export function branchListsOf(value: mce.AlternateContent, mode: BranchMode): Element[][] {
  if (mode === 'every') return branchesOf(value);
  const read = selectedBranchOf(value);
  return read ? [read] : [];
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

/**
 * The rows of a table, in order, descending into row-level content controls (an OpenDoPE repeat), and,
 * with a mode, into a row-level `mc:AlternateContent` (`BranchMode`; CR-013 phase 2).
 */
export function rowsOf(tbl: object, mode?: BranchMode): Located<wml.Tr>[] {
  return located(childrenOf(tbl), 'org_docx4j_wml.Tr', ROW_HOLDERS, mode) as Located<wml.Tr>[];
}

/** The cells of a row, in order, descending into cell-level content controls, and with a mode into a cell-level `mc:AlternateContent`. */
export function cellsOf(tr: object, mode?: BranchMode): Located<wml.Tc>[] {
  return located(childrenOf(tr), 'org_docx4j_wml.Tc', CELL_HOLDERS, mode) as Located<wml.Tc>[];
}

function located(items: Element[] | undefined, typeName: string, holders: Set<string>, mode: BranchMode | undefined, reading?: { ac: object; branch: Element[] }): Located[] {
  const out: Located[] = [];
  const visit = (list: Element[] | undefined): void => {
    if (!list) return;
    for (const element of list) {
      const tn = typeNameOf(element);
      if (tn === typeName) out.push({ element, container: list });
      else if (tn !== undefined && holders.has(tn) && typeof element.value === 'object' && element.value !== null) visit(childrenOf(element.value));
      else if (tn === ALTERNATE_CONTENT && mode !== undefined) {
        if (reading && element.value === reading.ac) visit(reading.branch);
        else for (const branch of branchListsOf(element.value as mce.AlternateContent, mode)) visit(branch);
      }
    }
  };
  visit(items);
  return out;
}

/**
 * A row's cells as laid along the grid, one list per reading of the row (docx4j's `Lister.cellLayoutsOf`,
 * its CR-051 phase 2): the selected branches' first (`cellsOf(tr, 'read')`), then, for each cell-level
 * `mc:AlternateContent` and each of its other branches, the row read with that branch in the element's
 * place. Every branch spans the same grid columns, so every cell is in some list at its grid position:
 * what a pass over the grid's columns walks (`collapseGrid`, the column passes; CR-013 phase 3).
 */
export function cellLayoutsOf(tr: object): Located<wml.Tc>[][] {
  const out: Located<wml.Tc>[][] = [cellsOf(tr, 'read')];
  const elements: mce.AlternateContent[] = [];
  const collect = (list: Element[] | undefined): void => {
    for (const element of list ?? []) {
      const tn = typeNameOf(element);
      if (tn === ALTERNATE_CONTENT) {
        elements.push(element.value as mce.AlternateContent);
        for (const branch of branchesOf(element.value as mce.AlternateContent)) collect(branch);
      } else if (tn !== undefined && CELL_HOLDERS.has(tn) && typeof element.value === 'object' && element.value !== null) {
        collect(childrenOf(element.value));
      }
    }
  };
  collect(childrenOf(tr));
  for (const ac of elements) {
    const read = selectedBranchOf(ac);
    for (const branch of branchesOf(ac)) {
      if (branch !== read) out.push(located(childrenOf(tr), 'org_docx4j_wml.Tc', CELL_HOLDERS, 'read', { ac, branch }) as Located<wml.Tc>[]);
    }
  }
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
