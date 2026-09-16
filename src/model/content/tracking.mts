// Change tracking (CR-002 phase F, section 3.7): the revision markup every mutation of the
// content API writes when `pkg.changeTrackingMode` is on. Word's shapes are ECMA-376 17.13.5:
// `w:ins` and `w:del` around runs, `w:delText` for deleted text, `w:pPr/w:rPr/w:ins` and
// `w:pPr/w:rPr/w:del` for a paragraph mark, `w:rPrChange` and `w:pPrChange` for formatting,
// `w:trPr/w:ins` and `w:trPr/w:del` for a table row.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy, deepCopyAsSync } from '@docx4j/generated-objects-ts';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import {
  createCTTrackChange, createCTRPrChange, createCTRPrChangeRPr, createCTPPrChange,
  createDelText, createPPr, createParaRPr, createTrPr,
} from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { rPrToElements } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { type Element, type RevisionHolder, W_NS, linkParents, walk, typeNameOf, revisionKindOf } from './tree.mjs';
import type { Author } from './comments.mjs';

/** Office JS `Word.ChangeTrackingMode`. */
export type ChangeTrackingMode = 'Off' | 'TrackAll' | 'TrackMineOnly';

/**
 * `w:id` on these is the one annotation id space (ECMA-376 17.13.5.4): revisions, bookmarks,
 * comment range marks and permissions share it. `w:comment/@w:id` (docx4j `Comments.Comment`,
 * which also extends `CTMarkup`) is a different space and is excluded, so that phase G's
 * comment ids and phase F's revision ids never have to agree.
 */
const MARKUP_TYPES = new Set([
  'org_docx4j_wml.CTMarkup', 'org_docx4j_wml.CTBookmark', 'org_docx4j_wml.CTBookmarkRange',
  'org_docx4j_wml.CTCellMergeTrackChange', 'org_docx4j_wml.CTMarkupRange', 'org_docx4j_wml.CTMathRunTrackChange',
  'org_docx4j_wml.CTMoveBookmark', 'org_docx4j_wml.CTMoveFromRangeEnd', 'org_docx4j_wml.CTMoveToRangeEnd',
  'org_docx4j_wml.CTPPrChange', 'org_docx4j_wml.CTRPrChange', 'org_docx4j_wml.CTSectPrChange',
  'org_docx4j_wml.CTTblGridChange', 'org_docx4j_wml.CTTblPrChange', 'org_docx4j_wml.CTTblPrExChange',
  'org_docx4j_wml.CTTcPrChange', 'org_docx4j_wml.CTTrPrChange', 'org_docx4j_wml.CTTrackChange',
  'org_docx4j_wml.CTTrackChangeNumbering', 'org_docx4j_wml.CTTrackChangeRange',
  'org_docx4j_wml.ParaRPrChange', 'org_docx4j_wml.RunDel', 'org_docx4j_wml.RunIns', 'org_docx4j_wml.RunTrackChange',
]);

/** What a package offers so that its edits can be tracked; `WordprocessingMLPackage` implements it. */
export interface TrackingHost {
  /**
   * The revision author (Office JS takes it from the signed-in user; here it is the package's).
   * The same `author` phase G's comments use, so a revision and a comment made in one session
   * carry one identity; only its `name` reaches `w:author`.
   */
  readonly author: Author;
  /** A fixed date for new revisions; `undefined` means now, read once per revision. */
  readonly trackedChangeDate?: Date | undefined;
  /** The trees to scan for the highest annotation id in use: the parts already unmarshalled. */
  markupRoots(): object[];
}

/** The tracker of a package whose mode is not `Off`; `undefined` when it is (duck typed, no import cycle). */
export function trackerOf(pkg: unknown): ChangeTracker | undefined {
  return (pkg as { changeTracker?: ChangeTracker } | undefined)?.changeTracker;
}

/** An `XmlCalendar` in UTC, as Word writes `w:date` (an xsd:dateTime ending in `Z`). */
export function calendarOf(date: Date): wml.XmlCalendar {
  return {
    year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate(),
    hour: date.getUTCHours(), minute: date.getUTCMinutes(), second: date.getUTCSeconds(), timezone: 0,
  };
}

/** The `Date` of a `w:date`, or undefined when there is none. */
export function dateOf(cal: wml.XmlCalendar | undefined): Date | undefined {
  if (!cal || cal.year === undefined) return undefined;
  const ms = Date.UTC(cal.year, (cal.month ?? 1) - 1, cal.day ?? 1, cal.hour ?? 0, cal.minute ?? 0, cal.second ?? 0);
  return new Date(ms - (cal.timezone ?? 0) * 60_000);
}

function element<T>(localPart: string, value: T): Element<T> {
  return { name: { namespaceURI: W_NS, localPart }, value };
}

/**
 * The per-package writer of revision markup. Created by the package while
 * `changeTrackingMode` is not `Off`; every content-API mutation asks its `Body` for one and,
 * when it gets one, writes revisions instead of editing in place.
 *
 * Departure from docx4j, which has no such class: docx4j leaves revision markup to the caller.
 * The name follows Office JS's vocabulary (`Word.ChangeTrackingMode`).
 */
export class ChangeTracker {
  private counter: number | undefined;

  constructor(readonly host: TrackingHost, readonly mode: ChangeTrackingMode) {}

  /** The `w:author` a new revision carries: the package author's name. */
  get author(): string {
    return this.host.author.name;
  }

  /** The next annotation id: one above the highest in use in the parts already unmarshalled. */
  nextId(): number {
    if (this.counter === undefined) {
      let highest = 0;
      for (const root of this.host.markupRoots()) {
        walk(root, (value) => {
          const tn = (value as { TYPE_NAME?: string }).TYPE_NAME;
          const id = (value as { id?: unknown }).id;
          if (tn !== undefined && MARKUP_TYPES.has(tn) && typeof id === 'number' && id > highest) highest = id;
        });
      }
      this.counter = highest;
    }
    return ++this.counter;
  }

  /** The `w:id`, `w:author` and `w:date` a new revision carries. */
  markup(): { id: number; author: string; date: wml.XmlCalendar } {
    return { id: this.nextId(), author: this.author, date: calendarOf(this.host.trackedChangeDate ?? new Date()) };
  }

  /** A `w:ins` / `w:del` attribute set: the id, this package's author and the date. */
  trackChange(): wml.CTTrackChange {
    return createCTTrackChange(this.markup());
  }

  /** `w:ins` around run-level items (they go under `customXmlOrSmartTagOrSdt`, not `content`). */
  ins(items: Element[]): Element<wml.RunIns> {
    const value: wml.RunIns = { TYPE_NAME: 'org_docx4j_wml.RunIns', ...this.markup(), customXmlOrSmartTagOrSdt: items as wml.RunIns['customXmlOrSmartTagOrSdt'] };
    const wrapper = element('ins', value);
    linkParents(items, value);
    return wrapper;
  }

  /** `w:del` around run-level items; the caller has already turned their `w:t` into `w:delText`. */
  del(items: Element[]): Element<wml.RunDel> {
    const value: wml.RunDel = { TYPE_NAME: 'org_docx4j_wml.RunDel', ...this.markup(), customXmlOrSmartTagOrSdt: items as wml.RunDel['customXmlOrSmartTagOrSdt'] };
    const wrapper = element('del', value);
    linkParents(items, value);
    return wrapper;
  }

  /** True when this revision is an insertion this package's author may extend rather than nest in. */
  ownInsertion(revision: RevisionHolder | undefined): boolean {
    return revision?.kind === 'ins' && revision.value.author === this.author;
  }

  /** Refuses an edit that would change deleted text; rejecting the deletion first is the way. */
  assertEditable(revision: RevisionHolder | undefined): void {
    if (revision?.kind === 'del' || revision?.kind === 'moveFrom') {
      throw new Docx4JException(`Cannot edit text inside a w:${revision.kind} (deleted by ${revision.value.author}); reject the deletion first`);
    }
  }

  /** Records the run's properties as they are now in `w:rPrChange`, once (a later write keeps the first original). */
  recordRPrChange(rPr: wml.RPr): void {
    if (rPr.rPrChange) return;
    const original = createCTRPrChangeRPr({ egrPrBase: rPrToElements(rPr) });
    rPr.rPrChange = createCTRPrChange({ ...this.markup(), rPr: original });
    linkParents(rPr.rPrChange, rPr);
  }

  /** Records the paragraph's properties as they are now in `w:pPrChange`, once. */
  recordPPrChange(pPr: wml.PPr): void {
    if (pPr.pPrChange) return;
    // w:pPrChange's child is a CT_PPrBase: `deepCopyAsSync` types the copy as the base and drops
    // what CT_PPrBase does not declare (w:rPr, w:sectPr, w:pPrChange), so no xsi:type is marshalled.
    // Synchronous: the part this runs on is unmarshalled, so the Jsonix context is built.
    const original = deepCopyAsSync<wml.PPrBase>(pPr, 'org_docx4j_wml.PPrBase');
    pPr.pPrChange = createCTPPrChange({ ...this.markup(), pPr: original });
    linkParents(pPr.pPrChange, pPr);
  }

  /** `w:pPr/w:rPr/w:ins`: the paragraph mark is an insertion. */
  markParagraphInserted(p: wml.P): void {
    const rPr = paraRPrOf(p);
    rPr.ins = this.trackChange();
    linkParents(rPr.ins, rPr);
  }

  /** `w:pPr/w:rPr/w:del`: the paragraph mark is deleted (accepting joins the paragraph with the next). */
  markParagraphDeleted(p: wml.P): void {
    const rPr = paraRPrOf(p);
    if (rPr.del) throw new Docx4JException('This paragraph mark is already marked deleted');
    rPr.del = this.trackChange();
    linkParents(rPr.del, rPr);
  }

  /** `w:trPr/w:ins`: the table row is an insertion (CR-002 phase C's `Table` calls this). */
  markRowInserted(tr: wml.Tr): void {
    const trPr = rowPrOf(tr);
    trPr.ins = this.trackChange();
    linkParents(trPr.ins, trPr);
  }

  /** `w:trPr/w:del`: the table row is deleted. */
  markRowDeleted(tr: wml.Tr): void {
    const trPr = rowPrOf(tr);
    if (trPr.del) throw new Docx4JException('This row is already marked deleted');
    trPr.del = this.trackChange();
    linkParents(trPr.del, trPr);
  }
}

/**
 * A whole paragraph inserted: the mark is marked inserted and every run of it goes into a
 * `w:ins`. Word, splitting a paragraph, marks the *first* paragraph's mark instead and leaves
 * the new one the original mark; marking the new paragraph's own mark is the same document once
 * accepted or rejected and keeps the edit to the element that was added (CR-002 section 13).
 */
export function trackInsertedParagraph(tracker: ChangeTracker, p: wml.P): void {
  tracker.markParagraphInserted(p);
  wrapNewRuns(tracker, p);
}

/** Every top-level run of the paragraph that is not already in a revision, wrapped in one `w:ins` per group. */
export function wrapNewRuns(tracker: ChangeTracker, p: wml.P): void {
  const content = p.content as Element[] | undefined;
  if (!content) return;
  for (let i = 0; i < content.length; i++) {
    if (typeNameOf(content[i]!) !== 'org_docx4j_wml.R' || revisionKindOf(content[i]!) !== undefined) continue;
    let j = i;
    while (j < content.length && typeNameOf(content[j]!) === 'org_docx4j_wml.R' && revisionKindOf(content[j]!) === undefined) j++;
    const runs = content.splice(i, j - i);
    const ins = tracker.ins(runs);
    content.splice(i, 0, ins);
    linkParents(ins, p);
  }
}

/** A whole table inserted: every row is marked inserted and every paragraph in it is an insertion. */
export function trackInsertedTable(tracker: ChangeTracker, tbl: wml.Tbl): void {
  for (const row of (tbl.content ?? []) as Element[]) {
    if (typeNameOf(row) !== 'org_docx4j_wml.Tr') continue;
    const tr = row.value as wml.Tr;
    tracker.markRowInserted(tr);
    for (const cell of (tr.content ?? []) as Element[]) {
      if (typeNameOf(cell) !== 'org_docx4j_wml.Tc') continue;
      for (const block of ((cell.value as wml.Tc).content ?? []) as Element[]) {
        if (typeNameOf(block) === 'org_docx4j_wml.P') trackInsertedParagraph(tracker, block.value as wml.P);
      }
    }
  }
}

/** The paragraph mark's run properties (`w:pPr/w:rPr`), created when absent. */
export function paraRPrOf(p: wml.P): wml.ParaRPr {
  const pPr = (p.pPr ??= createPPr());
  linkParents(pPr, p);
  const rPr = (pPr.rPr ??= createParaRPr());
  linkParents(rPr, pPr);
  return rPr;
}

/** The row's properties (`w:trPr`), created when absent. */
export function rowPrOf(tr: wml.Tr): wml.TrPr {
  const trPr = (tr.trPr ??= createTrPr());
  linkParents(trPr, tr);
  return trPr;
}

/** Drops `w:pPr/w:rPr` and `w:pPr` once nothing is left in them, so that accepting or rejecting leaves no husk. */
export function pruneParagraphProperties(p: wml.P): void {
  const pPr = p.pPr;
  if (!pPr) return;
  const rPr = pPr.rPr;
  if (rPr && isEmpty(rPr)) delete pPr.rPr;
  if (isEmpty(pPr)) delete p.pPr;
}

function isEmpty(value: object): boolean {
  return Object.keys(value).every((key) => key === 'TYPE_NAME' || (value as Record<string, unknown>)[key] === undefined);
}

/** True when a paragraph's mark is marked deleted. */
export function markDeleted(p: wml.P): boolean {
  return p.pPr?.rPr?.del !== undefined;
}

/** True when a paragraph's mark is marked inserted. */
export function markInserted(p: wml.P): boolean {
  return p.pPr?.rPr?.ins !== undefined;
}

// --- w:t <-> w:delText -------------------------------------------------------------------

const DELETED_NAME: Record<string, string> = { t: 'delText', instrText: 'delInstrText' };
const RESTORED_NAME: Record<string, string> = { delText: 't', delInstrText: 'instrText' };

/** Turns a run's `w:t` into `w:delText` and `w:instrText` into `w:delInstrText`, in place. */
export function toDeletedText(run: wml.R): void {
  rename(run, DELETED_NAME, true);
}

/** The inverse, for `reject()` of a deletion. */
export function toRestoredText(run: wml.R): void {
  rename(run, RESTORED_NAME, false);
}

function rename(run: wml.R, table: Record<string, string>, deleting: boolean): void {
  const content = run.content;
  if (!content) return;
  for (let i = 0; i < content.length; i++) {
    const item = content[i]!;
    const to = table[item.name.localPart];
    if (to === undefined) continue;
    const v = item.value as wml.Text;
    const value: wml.Text | wml.DelText = deleting && to === 'delText'
      ? createDelText({ value: v.value, ...(v.space === undefined ? {} : { space: v.space }) })
      : { TYPE_NAME: 'org_docx4j_wml.Text', value: v.value, ...(v.space === undefined ? {} : { space: v.space }) };
    content[i] = element(to, value) as (typeof content)[number];
    linkParents(content[i], run);
  }
}

// --- w:rPr <-> w:rPrChange/w:rPr ---------------------------------------------------------
// The conversion itself is the builders' `rPrToElements` / `rPrFromElements` (objects CR-003
// section 3.5, in 0.1.4): every EG_RPrBase member in schema order, the w14 effects included.

/** Replaces a run's direct formatting with `original`, keeping nothing of what is there. */
export function restoreRPr(holder: { rPr?: wml.RPr }, original: wml.RPr): void {
  if (Object.keys(original).filter((k) => k !== 'TYPE_NAME').length === 0) delete holder.rPr;
  else { holder.rPr = original; linkParents(original, holder); }
}

/** Replaces a paragraph's properties with `original` (a `w:pPrChange/w:pPr`), keeping the mark and the section. */
export function restorePPr(p: wml.P, original: wml.PPrBase | undefined): void {
  const pPr = p.pPr;
  if (!pPr) return;
  const keep = { rPr: pPr.rPr, sectPr: pPr.sectPr };
  for (const key of Object.keys(pPr)) if (key !== 'TYPE_NAME' && key !== 'PARENT') delete (pPr as unknown as Record<string, unknown>)[key];
  if (original) {
    for (const [key, value] of Object.entries(original)) {
      if (key === 'TYPE_NAME' || key === 'PARENT') continue;
      (pPr as unknown as Record<string, unknown>)[key] = deepCopy(value);
    }
  }
  if (keep.rPr) pPr.rPr = keep.rPr;
  if (keep.sectPr) pPr.sectPr = keep.sectPr;
  linkParents(pPr, p);
  pruneParagraphProperties(p);
}

/** A copy of a run's direct formatting for a new run: the recorded original (`w:rPrChange`) is not carried over. */
export function copyRPr(rPr: wml.RPr | undefined): wml.RPr | undefined {
  if (!rPr) return undefined;
  const copy = deepCopy(rPr);
  delete copy.rPrChange;
  return copy;
}

export type { RevisionHolder, RevisionKind } from './tree.mjs';
