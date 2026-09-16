// A subset of Office JS `Word.TrackedChange` over Word's revision markup (ECMA-376 17.13.5).
// Accepting and rejecting follow docx4j's AcceptTrackedChanges (docx4j-core
// org.docx4j.convert.out.common.preprocess): a w:ins is unwrapped, a w:del removed, a deleted
// paragraph mark joins its paragraph with the next. CR-002 phase F.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { textOf } from '@docx4j/generated-objects-ts/builders/wml';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { type Element, type RevisionKind, typeNameOf, runItemsOf, revisionKindOf, linkParents, segmentsOf } from './tree.mjs';
import { Range } from './Range.mjs';
import type { Paragraph } from './Paragraph.mjs';
import { rPrFromElements } from '@docx4j/generated-objects-ts/builders/wml';
import { dateOf, toRestoredText, restoreRPr, restorePPr, pruneParagraphProperties } from './tracking.mjs';

/** Office JS `Word.ChangeTrackingState` (`Unknown` is not produced here). */
export type TrackedChangeType = 'Added' | 'Deleted' | 'Formatted' | 'None' | 'Unknown';

/** Which piece of revision markup a `TrackedChange` is over (an extension: Office JS has no such member). */
export type TrackedChangeTarget =
  /** `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` around runs. */
  | { kind: 'run'; revision: RevisionKind; element: Element; owner: Element[]; value: wml.CTTrackChange }
  /** `w:pPr/w:rPr/w:ins` or `w:pPr/w:rPr/w:del`: the paragraph mark. */
  | { kind: 'mark'; mark: 'ins' | 'del'; value: wml.CTTrackChange }
  /** `w:rPr/w:rPrChange`: a run's formatting. */
  | { kind: 'runProperties'; run: wml.R; value: wml.CTRPrChange }
  /** `w:pPr/w:pPrChange`: the paragraph's properties. */
  | { kind: 'paragraphProperties'; value: wml.CTPPrChange }
  /** `w:trPr/w:ins` or `w:trPr/w:del`: a table row. */
  | { kind: 'row'; row: 'ins' | 'del'; tr: Element<wml.Tr>; owner: Element[]; value: wml.CTTrackChange };

/**
 * A subset of Office JS `Word.TrackedChange`: `type`, `author`, `date`, `text`, `accept()`,
 * `reject()` and `getRange()`, plus `element`, `id` and `target` as extensions. A view, like
 * `Paragraph` and `Range`: it holds the markup it is over, nothing is cached.
 */
export class TrackedChange {
  constructor(
    readonly target: TrackedChangeTarget,
    /** The paragraph the change is in; undefined for a row revision. */
    readonly paragraph: Paragraph | undefined,
  ) {}

  /** Office JS `Word.ChangeTrackingState`. */
  get type(): TrackedChangeType {
    const t = this.target;
    switch (t.kind) {
      case 'run': return t.revision === 'ins' || t.revision === 'moveTo' ? 'Added' : 'Deleted';
      case 'mark': return t.mark === 'ins' ? 'Added' : 'Deleted';
      case 'row': return t.row === 'ins' ? 'Added' : 'Deleted';
      default: return 'Formatted';
    }
  }

  get author(): string {
    return this.target.value.author;
  }

  /** `w:date`, when the markup carries one. */
  get date(): Date | undefined {
    return dateOf(this.target.value.date);
  }

  /** The annotation id (`w:id`; an extension). */
  get id(): number {
    return this.target.value.id;
  }

  /** The markup this change is over (an extension): the `w:ins` element, the `w:rPrChange`, ... */
  get element(): object {
    return this.target.kind === 'run' ? this.target.element : this.target.kind === 'row' ? this.target.tr : this.target.value;
  }

  /** The text the change covers; '' for a paragraph mark, whose range is the paragraph. */
  get text(): string {
    const t = this.target;
    switch (t.kind) {
      case 'run': return segmentsOf(t.element.value as object, { view: t.revision === 'ins' || t.revision === 'moveTo' ? 'accepted' : 'original' }).map((s) => s.text).join('');
      case 'runProperties': return textOf(t.run);
      case 'paragraphProperties': return this.paragraph?.text ?? '';
      case 'row': return textOf(t.tr.value as object);
      default: return '';
    }
  }

  /** The range the change covers; a deletion is a collapsed range where it sits in the accepted text. */
  getRange(): Range | undefined {
    const p = this.paragraph;
    if (!p) return undefined;
    const t = this.target;
    if (t.kind === 'run') {
      const [start, end] = spanOf(p, t.element);
      return new Range(p, start, end);
    }
    if (t.kind === 'runProperties') {
      const segs = p.segments().filter((s) => s.run === t.run);
      if (segs.length === 0) return p.getRange('Start');
      return new Range(p, segs[0]!.start, segs[segs.length - 1]!.end);
    }
    return p.getRange('Whole');
  }

  /** Keeps the change: a `w:ins` is unwrapped, a `w:del` removed, a deleted mark joins the paragraphs. */
  accept(): void {
    const t = this.target;
    switch (t.kind) {
      case 'run':
        if (t.revision === 'ins' || t.revision === 'moveTo') unwrap(t); else remove(t.owner, t.element);
        return;
      case 'mark':
        if (t.mark === 'ins') dropMark(this.paragraph, 'ins'); else joinWithNext(this.requireParagraph());
        return;
      case 'runProperties':
        delete t.run.rPr?.rPrChange;
        return;
      case 'paragraphProperties': {
        const p = this.requireParagraph().p;
        delete p.pPr?.pPrChange;
        pruneParagraphProperties(p);
        return;
      }
      case 'row':
        if (t.row === 'ins') delete t.tr.value.trPr?.ins; else remove(t.owner, t.tr as Element);
        return;
    }
  }

  /** Puts back what was there: a `w:ins` is removed, a `w:del` restored, a `w:rPrChange` re-applied. */
  reject(): void {
    const t = this.target;
    switch (t.kind) {
      case 'run':
        if (t.revision === 'ins' || t.revision === 'moveTo') remove(t.owner, t.element); else restoreDeleted(t);
        return;
      case 'mark':
        if (t.mark === 'del') dropMark(this.paragraph, 'del'); else joinWithNext(this.requireParagraph(), true);
        return;
      case 'runProperties':
        restoreRPr(t.run, rPrFromElements(t.value.rPr));
        return;
      case 'paragraphProperties': {
        const p = this.requireParagraph().p;
        restorePPr(p, t.value.pPr);
        return;
      }
      case 'row':
        if (t.row === 'del') delete t.tr.value.trPr?.del; else remove(t.owner, t.tr as Element);
        return;
    }
  }

  private requireParagraph(): Paragraph {
    if (!this.paragraph) throw new Docx4JException('This tracked change has no paragraph');
    return this.paragraph;
  }
}

// --- the operations ------------------------------------------------------------------------

function remove(owner: Element[], element: Element): void {
  const i = owner.indexOf(element);
  if (i >= 0) owner.splice(i, 1);
}

/** A `w:ins` or `w:moveTo` accepted: its runs take its place. */
function unwrap(t: Extract<TrackedChangeTarget, { kind: 'run' }>): void {
  const i = t.owner.indexOf(t.element);
  if (i < 0) return;
  const items = runItemsOf(t.element.value as object) ?? [];
  t.owner.splice(i, 1, ...items);
  const parent = (t.value as { PARENT?: object }).PARENT;
  for (const item of items) linkParents(item, parent);
}

/** A `w:del` or `w:moveFrom` rejected: its runs come back, with `w:delText` turned into `w:t`. */
function restoreDeleted(t: Extract<TrackedChangeTarget, { kind: 'run' }>): void {
  for (const item of runItemsOf(t.element.value as object) ?? []) {
    if (typeNameOf(item) === 'org_docx4j_wml.R') toRestoredText(item.value as wml.R);
  }
  unwrap(t);
}

function dropMark(paragraph: Paragraph | undefined, which: 'ins' | 'del'): void {
  if (!paragraph) return;
  const rPr = paragraph.p.pPr?.rPr;
  if (rPr) delete rPr[which];
  pruneParagraphProperties(paragraph.p);
}

/**
 * docx4j AcceptTrackedChanges: the paragraph takes the next one's content and properties, so the
 * mark that survives is the next one's. Called when a deleted mark is accepted and when an
 * inserted one is rejected.
 *
 * `fallbackToPrevious` is for the second case at the end of a container, where this package's
 * inserted paragraph carries its own mark (section 13) and there is nothing after it to join
 * with: the paragraph then goes, its content joining the one before. With neither neighbour the
 * mark is simply dropped.
 */
export function joinWithNext(paragraph: Paragraph, fallbackToPrevious = false): void {
  const container = paragraph.container;
  const i = container.indexOf(paragraph.element);
  const next = i >= 0 ? container[i + 1] : undefined;
  const p = paragraph.p;
  if (!next || typeNameOf(next) !== 'org_docx4j_wml.P') {
    const previous = fallbackToPrevious && i > 0 ? container[i - 1] : undefined;
    if (previous && typeNameOf(previous) === 'org_docx4j_wml.P') {
      const into = previous.value as wml.P;
      const items = (into.content ??= []);
      for (const item of p.content ?? []) items.push(item as never);
      linkParents(items, into);
      container.splice(i, 1);
      return;
    }
    const rPr = p.pPr?.rPr;
    if (rPr) { delete rPr.ins; delete rPr.del; }
    pruneParagraphProperties(p);
    return;
  }
  const nextP = next.value as wml.P;
  const content = (p.content ??= []);
  for (const item of nextP.content ?? []) content.push(item as never);
  linkParents(content, p);
  p.pPr = nextP.pPr;
  if (p.pPr) linkParents(p.pPr, p);
  container.splice(i + 1, 1);
}

/** The span of a run-level revision in the paragraph's accepted text; collapsed for a deletion. */
function spanOf(paragraph: Paragraph, element: Element): [number, number] {
  const segs = paragraph.segments();
  const inside = segs.filter((s) => s.revision?.element === element);
  if (inside.length > 0) return [inside[0]!.start, inside[inside.length - 1]!.end];
  const at = acceptedOffsetOf(paragraph.p, element);
  return [at, at];
}

/** How much accepted text comes before `target` in the paragraph. */
function acceptedOffsetOf(p: wml.P, target: Element): number {
  let pos = 0;
  let found = -1;
  const visit = (items: Element[] | undefined): void => {
    if (!items || found >= 0) return;
    for (const el of items) {
      if (found >= 0) return;
      if (el === target) { found = pos; return; }
      if (typeNameOf(el) === 'org_docx4j_wml.R') { pos += segmentsOf({ content: [el] }).map((s) => s.text).join('').length; continue; }
      const kind = revisionKindOf(el);
      if (kind === 'del' || kind === 'moveFrom') continue;
      visit(runItemsOf(el.value as object));
    }
  };
  visit(runItemsOf(p));
  return found >= 0 ? found : pos;
}

// --- collecting ----------------------------------------------------------------------------

/** The tracked changes of one paragraph, in document order. */
export function trackedChangesOfParagraph(paragraph: Paragraph): TrackedChange[] {
  const out: TrackedChange[] = [];
  const p = paragraph.p;
  const pPrChange = p.pPr?.pPrChange;
  if (pPrChange) out.push(new TrackedChange({ kind: 'paragraphProperties', value: pPrChange }, paragraph));
  const rPr = p.pPr?.rPr;
  if (rPr?.ins) out.push(new TrackedChange({ kind: 'mark', mark: 'ins', value: rPr.ins }, paragraph));
  if (rPr?.del) out.push(new TrackedChange({ kind: 'mark', mark: 'del', value: rPr.del }, paragraph));
  collectRunLevel(paragraph, (p.content ?? []) as Element[], out);
  return out;
}

function collectRunLevel(paragraph: Paragraph, items: Element[], out: TrackedChange[]): void {
  for (const el of items) {
    const kind = revisionKindOf(el);
    if (kind !== undefined) {
      out.push(new TrackedChange({ kind: 'run', revision: kind, element: el, owner: items, value: el.value as wml.CTTrackChange }, paragraph));
      const nested = runItemsOf(el.value as object);
      if (nested) collectRunLevel(paragraph, nested, out);
      continue;
    }
    if (typeNameOf(el) === 'org_docx4j_wml.R') {
      const run = el.value as wml.R;
      if (run.rPr?.rPrChange) out.push(new TrackedChange({ kind: 'runProperties', run, value: run.rPr.rPrChange }, paragraph));
      continue;
    }
    const nested = runItemsOf(el.value as object);
    if (nested) collectRunLevel(paragraph, nested, out);
  }
}

/** The `w:trPr/w:ins` and `w:trPr/w:del` revisions of a table row. */
export function trackedChangesOfRow(tr: Element<wml.Tr>, owner: Element[]): TrackedChange[] {
  const trPr = tr.value.trPr;
  const out: TrackedChange[] = [];
  if (trPr?.ins) out.push(new TrackedChange({ kind: 'row', row: 'ins', tr, owner, value: trPr.ins }, undefined));
  if (trPr?.del) out.push(new TrackedChange({ kind: 'row', row: 'del', tr, owner, value: trPr.del }, undefined));
  return out;
}
