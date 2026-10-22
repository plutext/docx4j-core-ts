// The writes of a Range across paragraphs, and of a range holding its last paragraph's mark (CR-014 phase 2), as
// Word 365 makes them through Office JS (check 52, test/README.md; CR-014 sections 6.2 and 10). `Range` calls in
// here for every write whose answer isn't one paragraph's splice; what it gets back is where the text closed up.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy } from '@docx4j/generated-objects-ts';
import { applyRunOptions, readRunOptions, t as textItem, type RunOptions } from '@docx4j/generated-objects-ts/builders/wml';
import { createPPr, createParaRPr } from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { NotSupportedError } from '../../office-js/errors.mjs';
import { Docx4JException } from '../../opc/exceptions.mjs';
import type { PropertyResolver } from '../properties/PropertyResolver.mjs';
import { type Element, typeNameOf, childrenOf, linkParents, runItemsOf, runOf, paragraphOf, revisionKindOf, objectInSpan, BLOCK_LEVEL_TYPES, ALTERNATE_CONTENT, type SpanEdges } from './tree.mjs';
import { joinWithNext, queueNoteRemovals } from './TrackedChange.mjs';
import { markDeleted, copyRPr, type ChangeTracker } from './tracking.mjs';
import { storyLayout, storyOf } from './story.mjs';
import { Font } from './Font.mjs';
import { GeneralException } from './generalException.mjs';
import type { Paragraph } from './Paragraph.mjs';
import type { Body } from './Body.mjs';

/** What a write needs of the `Range` it writes: its two ends, and whether it holds its last paragraph's mark. */
export interface Span {
  readonly paragraph: Paragraph;
  readonly endParagraph: Paragraph;
  readonly start: number;
  readonly end: number;
  readonly whole: boolean;
  /** The range holds its last paragraph's mark (a 'Whole' range), or a table's end in its place. */
  readonly holdsEndMark: boolean;
}

/**
 * Where a range's two ends stand in the innermost list of blocks holding both: the block holding the start and
 * the block holding the end, each the end's paragraph itself, or a table or content control the range holds
 * whole (a table's 'Whole' range, a body's that starts with a table).
 */
export interface Extent {
  list: Element[];
  first: number;
  last: number;
  /** The start's paragraph, where it is the first block itself; else that block goes whole. */
  head?: Paragraph;
  /** The end's paragraph, where it is the last block itself; else that block goes whole. */
  tail?: Paragraph;
}

/** Where a deletion closed up: the paragraph, and the offset in it a replacement's text goes at. */
export interface Landing {
  paragraph: Paragraph;
  offset: number;
}

/**
 * The extent of a range, refusing one that starts or ends inside a block it doesn't hold whole. Inside a table
 * that is Word's own refusal (check 52: `delete()` and `Replace` on a range from a cell to a paragraph after the
 * table, or from a paragraph into a cell, threw `GeneralException` and changed nothing, tracked or not; D3).
 * Inside a content control, unmeasured, it is `NotSupportedError`.
 */
export function extentOf(span: Span, member: string): Extent {
  const story = storyOf(span.paragraph.p).items;
  const up = blocksHolding(span.paragraph.p, story);
  const down = blocksHolding(span.endParagraph.p, story);
  if (up[up.length - 1]?.list !== story || down[down.length - 1]?.list !== story) {
    throw new Docx4JException("This range's paragraphs are not in the document any more: an edit through another view removed or joined them");
  }
  for (const [i, link] of up.entries()) {
    const j = down.findIndex((other) => other.list === link.list);
    if (j < 0) continue;
    const a = link.element;
    const b = down[j]!.element;
    if (i > 0 && !startsWhole(span, a)) refusePartial(a, member, 'starts');
    if (j > 0 && !endsWhole(span, b)) refusePartial(b, member, 'ends');
    const extent: Extent = { list: link.list, first: link.list.indexOf(a), last: link.list.indexOf(b) };
    if (i === 0) extent.head = span.paragraph;
    if (j === 0) extent.tail = span.endParagraph;
    return extent;
  }
  throw new Docx4JException("This range's two ends are not in one story");
}

/** The blocks holding a paragraph, innermost first: the paragraph in its list, each block holding that in the list holding it, up to the story's. */
function blocksHolding(p: wml.P, story: Element[]): { list: Element[]; element: Element }[] {
  const out: { list: Element[]; element: Element }[] = [];
  let o: object = p;
  for (let guard = 0; guard < 256; guard++) {
    const parent = (o as { PARENT?: object }).PARENT;
    if (!parent) break;
    const list = childrenOf(parent);
    const element = list?.find((e) => e.value === o);
    const tn = element ? typeNameOf(element) : undefined;
    if (list && element && tn !== undefined && (BLOCK_LEVEL_TYPES.has(tn) || tn === ALTERNATE_CONTENT)) {
      out.push({ list, element });
      if (list === story) break;
    }
    o = parent;
  }
  return out;
}

/** The range starts at the start of the block: its first paragraph, as a reader reads it, at offset 0. */
function startsWhole(span: Span, block: Element): boolean {
  return span.start === 0 && storyLayout([block], () => '').paragraphs[0]?.element.value === span.paragraph.p;
}

/** The range ends at the end of the block: its last paragraph's, as a reader reads it, the mark (or a table's row end) included. */
function endsWhole(span: Span, block: Element): boolean {
  const paragraphs = storyLayout([block], () => '').paragraphs;
  return span.holdsEndMark && paragraphs[paragraphs.length - 1]?.element.value === span.endParagraph.p;
}

function refusePartial(block: Element, member: string, side: 'starts' | 'ends'): never {
  if (typeNameOf(block) === 'org_docx4j_wml.Tbl') {
    throw new GeneralException(`Range.${member}: this range ${side} inside a table, which Office JS refuses too (check 52); `
      + 'take the whole table, or edit inside one cell');
  }
  throw new NotSupportedError('Range', member, `this range ${side} inside a content control (or another block) it doesn't hold whole, `
    + "and what Word does with that isn't measured (CR-014 phase 2)");
}

// --- delete -------------------------------------------------------------------------------------------

/**
 * Deletes a range across paragraphs, or one holding its last paragraph's mark, as Word does (check 52, cases 3 to
 * 28): the start's paragraph loses its text from the start, every block between goes whole (a table too), the end's
 * paragraph loses its text to the end, and the two join in the end paragraph's properties, each run keeping its own
 * formatting. Where the range holds the end paragraph's mark, that goes too and the paragraph after it is joined,
 * in its properties (W1, W2); with none after it in the same list, a cell's last paragraph or the body's, the mark
 * stays (CR-002 section 55: Word keeps a cell's last paragraph). `keepLastMark` is `Replace`'s, which keeps it.
 *
 * Tracked, nothing is joined: the text is a `w:del`, every mark in the range is marked deleted but a last one the
 * range doesn't hold, and a table's rows and its cells' marks are deleted, with no `w:pPrChange` (Word's markup in
 * its save after the run). Accepting gives the untracked answer.
 */
export function deleteSpan(span: Span, extent: Extent, keepLastMark: boolean): Landing | undefined {
  const tracker = span.paragraph.changeTracker;
  return tracker ? deleteTracked(span, extent, tracker, keepLastMark) : deleteUntracked(span, extent, keepLastMark);
}

function deleteUntracked(span: Span, { list, first, last, head, tail }: Extent, keepLastMark: boolean): Landing | undefined {
  const takesMark = span.holdsEndMark && !keepLastMark;
  if (head && tail && head.element === tail.element) {
    cut(head, span.start, span.end, { start: span.whole, end: true });
    if (takesMark) joinForward(head);
    return { paragraph: head, offset: span.start };
  }
  // the end first, so that the indexes before it hold
  if (tail) {
    cut(tail, 0, span.end, { start: true, end: span.whole || span.holdsEndMark });
    if (takesMark) joinForward(tail);
  } else {
    removeBlock(span.paragraph, list, last);
  }
  // the blocks between; a range marker between blocks stays, for the join to bring into the joined paragraph
  for (let i = last - 1; i > first; i--) if (!isRangeMarker(list[i]!)) removeBlock(span.paragraph, list, i);
  if (head) {
    cut(head, span.start, head.text.length, { start: span.whole, end: true });
    // a whole table or control at the end goes, and the start's paragraph keeps its mark (unmeasured)
    if (tail) join(head);
    return { paragraph: head, offset: span.start };
  }
  removeBlock(span.paragraph, list, first);
  return tail ? { paragraph: tail, offset: 0 } : undefined;
}

/**
 * Refuses, before anything changes, a tracked deletion of a range holding a whole block that isn't a paragraph or a
 * table: a content control or custom XML element, an `w:altChunk`, a kept `mc:AlternateContent`. A control's content
 * can be marked deleted, but what Word writes for the element itself isn't measured (check 30 measured an inserted
 * control's markers only), and an `w:altChunk` has no deleted form; Accept All would then leave the element where
 * the untracked deletion takes it, the paragraphs either side unjoined. Untracked, such a block goes whole, as a
 * table does.
 */
export function assertTrackable(extent: Extent, member: string): void {
  for (let i = extent.first; i <= extent.last; i++) {
    if ((i === extent.first && extent.head) || (i === extent.last && extent.tail)) continue;
    const block = extent.list[i]!;
    const tn = typeNameOf(block) ?? '';
    if (tn === 'org_docx4j_wml.P' || tn === 'org_docx4j_wml.Tbl' || isRangeMarker(block)) continue;
    throw new NotSupportedError('Range', member, `this range holds a whole ${block.name?.localPart ?? tn}, and there is no measured way to delete one `
      + 'under change tracking (CR-014 phase 2); delete it untracked, or the paragraphs either side of it');
  }
}

function deleteTracked(span: Span, { list, first, last, head, tail }: Extent, tracker: ChangeTracker, keepLastMark: boolean): Landing | undefined {
  const takesMark = span.holdsEndMark && !keepLastMark;
  const body = span.paragraph.parentBody;
  if (head && tail && head.element === tail.element) {
    head.splice(span.start, span.end, '', 'before', { start: span.whole, end: true });
    if (takesMark) deleteMark(head, tracker);
    return { paragraph: head, offset: span.start };
  }
  if (head) {
    head.splice(span.start, head.text.length, '', 'before', { start: span.whole, end: true });
    if (tail) deleteMark(head, tracker);
  } else {
    markBlockDeleted(body, list[first]!, tracker);
  }
  for (let i = first + 1; i < last; i++) if (!isRangeMarker(list[i]!)) markBlockDeleted(body, list[i]!, tracker);
  if (tail) {
    tail.splice(0, span.end, '', 'before', { start: true, end: span.whole || span.holdsEndMark });
    if (takesMark) deleteMark(tail, tracker);
  } else {
    markBlockDeleted(body, list[last]!, tracker);
  }
  return head ? { paragraph: head, offset: span.start } : tail ? { paragraph: tail, offset: 0 } : undefined;
}

/** Removes [from, to) of a paragraph's text, untracked, the notes its references go with queued to go too (check 52 case 12). */
function cut(p: Paragraph, from: number, to: number, edges: SpanEdges): void {
  const removed = p.segments({ objects: true })
    .filter((s) => (s.object ? objectInSpan(s.start, from, to, edges) : s.start < to && s.end > from))
    .map((s) => s.item);
  if (to <= from && removed.length === 0) return;
  queueNoteRemovals(p.parentBody.package_, removed);
  p.splice(from, to, '', 'before', edges);
}

/** Removes a block, untracked: its notes and the relationships it held go on save where nothing else holds them. */
function removeBlock(p: Paragraph, list: Element[], index: number): void {
  const [removed] = list.splice(index, 1);
  if (!removed) return;
  queueNoteRemovals(p.parentBody.package_, removed);
  p.releaseRelationships(removed);
}

/** The paragraph after `p` in its list, past range markers, joined into it where there is one: its content and properties. */
function joinForward(p: Paragraph): void {
  const list = listOf(p);
  let j = list.indexOf(p.element) + 1;
  while (j < list.length && isRangeMarker(list[j]!)) j++;
  if (j < list.length && typeNameOf(list[j]!) === 'org_docx4j_wml.P') join(p);
}

/**
 * `p` takes the next paragraph's content and properties, which is Word's join (`joinWithNext`, the accept side's).
 * A section break in `p`'s mark goes with the mark, and the relationships its headers and footers held with it.
 */
function join(p: Paragraph): void {
  const sectPr = (p.p.pPr as { sectPr?: unknown } | undefined)?.sectPr;
  if (sectPr) p.releaseRelationships(sectPr);
  joinWithNext(p);
}

function listOf(p: Paragraph): Element[] {
  const list = childrenOf(((p.p as { PARENT?: object }).PARENT ?? {}) as object);
  return list?.includes(p.element) ? list : p.container;
}

/** A paragraph's mark marked deleted, unless it is already. */
function deleteMark(p: Paragraph, tracker: ChangeTracker): void {
  if (!markDeleted(p.p)) tracker.markParagraphDeleted(p.p);
}

/** Everything in a block marked deleted in place: each paragraph's text and mark, each row of its tables (check 52 case 21). */
function markBlockDeleted(body: Body, block: Element, tracker: ChangeTracker): void {
  body.sub({ content: [block] }).markContentDeleted(tracker);
}

/** A marker that stands between two paragraphs without keeping them apart (as `joinWithNext` reads one). */
function isRangeMarker(el: Element): boolean {
  const name = el.name?.localPart ?? '';
  return name.endsWith('RangeStart') || name.endsWith('RangeEnd') || name === 'bookmarkStart' || name === 'bookmarkEnd'
    || name === 'permStart' || name === 'permEnd' || name === 'proofErr';
}

// --- Replace ------------------------------------------------------------------------------------------

/** The first replaced character's run properties and where it stood, read before anything is deleted. */
export interface FirstCharacter {
  rPr: wml.RPr | undefined;
  pPr: wml.PPr | undefined;
  p: wml.P;
}

/** The character at a range's start: its run's properties, or its paragraph's mark's where the range starts at the mark. */
export function firstCharacterOf(span: Span): FirstCharacter {
  const p = span.paragraph;
  const seg = p.segments().find((s) => s.start <= span.start && span.start < s.end);
  return { rPr: seg ? seg.run.rPr : markRPrOf(p.p), pPr: p.p.pPr, p: p.p };
}

/**
 * The run properties that give text in `into` the look the first replaced character had (check 52, cases 29 to
 * 38): its run's own properties, and each property its paragraph's style gave it that `into`'s doesn't, written
 * direct - a heading's font, size, colour and bold in a body paragraph (case 29), and a body paragraph's
 * `w:b w:val="0"` and automatic colour in a heading (case 30). A property no style sets anywhere and `into`'s does
 * is written as its neutral value where it has one (`NEUTRAL`).
 */
export function lookOf(first: FirstCharacter, into: Paragraph): wml.RPr | undefined {
  const direct = copyRPr(first.rPr);
  const resolver = (into.parentBody.package_ as { propertyResolverOrUndefined?: PropertyResolver } | undefined)?.propertyResolverOrUndefined;
  // the same paragraph properties (a range within one paragraph): its own properties give it its look
  if (!resolver || first.pPr === into.p.pPr) return direct;
  const was = resolver.getEffectiveRPr(first.rPr, first.pPr, resolver.cellContextOf(first.p)) as Record<string, unknown>;
  const now = resolver.getEffectiveRPr(first.rPr, into.p.pPr, resolver.cellContextOf(into.p)) as Record<string, unknown>;
  const out = (direct ?? { TYPE_NAME: 'org_docx4j_wml.RPr' }) as Record<string, unknown>;
  for (const key of new Set([...Object.keys(was), ...Object.keys(now)])) {
    if (SKIPPED.has(key) || same(was[key], now[key])) continue;
    const value = was[key] !== undefined ? deepCopy(was[key] as object) : NEUTRAL[key]?.();
    if (value !== undefined) out[key] = value;
  }
  const rPr = out as wml.RPr;
  return Object.keys(rPr).some((key) => key !== 'TYPE_NAME') ? rPr : undefined;
}

const SKIPPED = new Set(['TYPE_NAME', 'PARENT', 'rPrChange', 'rStyle']);

/** What a property reads as where nothing sets it, for the properties that have one. */
const NEUTRAL: Record<string, () => object> = {
  ...Object.fromEntries(['b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'vanish', 'webHidden']
    .map((key) => [key, () => ({ val: false })])),
  color: () => ({ val: 'auto' }),
  u: () => ({ val: 'none' }),
  highlight: () => ({ val: 'none' }),
  vertAlign: () => ({ val: 'baseline' }),
  sz: () => ({ val: 20 }),
  szCs: () => ({ val: 20 }),
};

function same(a: unknown, b: unknown): boolean {
  return JSON.stringify(a, (key, v) => (key === 'PARENT' || key === 'TYPE_NAME' ? undefined : v)) === JSON.stringify(b, (key, v) => (key === 'PARENT' || key === 'TYPE_NAME' ? undefined : v));
}

/** The text of an untracked `Replace`, put where the deletion closed up, in the first replaced character's look. */
export function placeReplacement(landing: Landing, first: FirstCharacter, text: string): void {
  if (text === '') return;
  landing.paragraph.insertItemsAt(landing.offset, [runOf([textItem(text)], lookOf(first, landing.paragraph))]);
}

/** A paragraph mark's formatting as run properties: the mark's own revisions left out. */
function markRPrOf(p: wml.P): wml.RPr | undefined {
  const rPr = p.pPr?.rPr;
  if (!rPr) return undefined;
  const { ins: _ins, del: _del, moveFrom: _from, moveTo: _to, rPrChange: _change, ...rest } = rPr;
  return Object.keys(rest).some((key) => key !== 'TYPE_NAME') ? { ...rest, TYPE_NAME: 'org_docx4j_wml.RPr' } as wml.RPr : undefined;
}

// --- insertParagraph ------------------------------------------------------------------------------------

/**
 * A paragraph of `text` put at `offset` in `p`, as Word puts one where a range across paragraphs starts
 * ('Before') or ends ('After') (check 52, cases 57 to 70): at the paragraph's start it goes before it and at its
 * end after it, in its properties (`Paragraph.insertParagraph`); inside it `p` is split there (`splitParagraph`)
 * and the new paragraph goes between the halves, its mark and its text in the formatting of the run before the
 * split (case 63's italic). Tracked, the new paragraph and the first half's mark are inserted (cases 69 and 70).
 * Returns the new paragraph, and the first half where `p` was split.
 */
export function insertParagraphAt(p: Paragraph, offset: number, text: string): { inserted: Paragraph; firstHalf?: Paragraph } {
  if (offset <= 0) return { inserted: p.insertParagraph(text, 'Before') };
  if (offset >= p.text.length) return { inserted: p.insertParagraph(text, 'After') };
  const rPr = rPrBefore(p, offset);
  const firstHalf = splitParagraph(p, offset, rPr);
  const pPr = propertiesOf(p, rPr);
  const el = paragraphOf(text === '' ? [] : [runOf([textItem(text)], copyRPr(rPr))], pPr);
  return { inserted: p.parentBody.insertElement(el, 'Before', p) as Paragraph, firstHalf };
}

/** The run properties of the text just before `offset`, which a split there gives its new marks (check 52 case 63). */
function rPrBefore(p: Paragraph, offset: number): wml.RPr | undefined {
  const before = p.segments().filter((s) => s.start < offset && s.text.length > 0);
  return copyRPr(before[before.length - 1]?.run.rPr);
}

/** A copy of `p`'s properties for a new paragraph: no section break, no tracked change, the mark in `rPr`'s formatting. */
function propertiesOf(p: Paragraph, rPr: wml.RPr | undefined): wml.PPr | undefined {
  const pPr = p.p.pPr ? deepCopy(p.p.pPr) : undefined;
  if (pPr) { delete pPr.rPr; delete (pPr as { sectPr?: unknown }).sectPr; delete pPr.pPrChange; }
  if (!rPr) return pPr;
  const out = pPr ?? createPPr();
  out.rPr = { ...deepCopy(rPr), TYPE_NAME: 'org_docx4j_wml.ParaRPr' } as wml.ParaRPr;
  linkParents(out.rPr, out);
  return out;
}

/**
 * Splits `p` at `offset`, inside its text: what is before it becomes a new paragraph in front of `p`, in a copy of
 * its properties with a mark of its own in `rPr`'s formatting, and `p` keeps the rest with its own mark, its
 * section break and its identity, as Word keeps them (check 52 cases 57 and 63). A link, a revision or a smart tag
 * the offset is inside is split into two of the same; a content control, a field or a kept `mc:AlternateContent`
 * there is refused, since its two halves would be two of it. Tracked, the new mark is inserted, as Word marks the
 * first half's (cases 69 and 70). Returns the first half.
 */
export function splitParagraph(p: Paragraph, offset: number, rPr: wml.RPr | undefined): Paragraph {
  p.assertWritable();
  const length = p.text.length;
  // the character at the split: the holders above it must be ones that can be cut in two
  const seg = p.segments().find((s) => s.start <= offset && offset < s.end);
  if (!seg || offset <= 0 || offset >= length) throw new Docx4JException(`A paragraph splits inside its text, not at ${offset} of ${length}`);
  if (seg.alternate || seg.math) throw new NotSupportedError('Range', 'insertParagraph', 'the split falls inside an equation or an mc:AlternateContent');
  // every check before anything changes
  for (let o = (seg.run as { PARENT?: object }).PARENT; o && o !== p.p; o = (o as { PARENT?: object }).PARENT) {
    const tn = (o as { TYPE_NAME?: string }).TYPE_NAME ?? '';
    if (!SPLITTABLE.has(tn)) throw new NotSupportedError('Range', 'insertParagraph', `the split falls inside a ${tn.replace('org_docx4j_wml.', '')}, which can't be cut in two`);
  }
  p.splitRunsAround(offset, length);
  const at = p.segments().find((s) => s.start >= offset && s.text.length > 0)!;
  let owner = at.runOwner;
  let index = at.runIndex;
  const content = (p.p.content ??= []) as Element[];
  while (owner !== content) {
    const holder = (owner[index]!.value as { PARENT?: object }).PARENT as Record<string, unknown> | undefined;
    const parent = holder?.PARENT as object | undefined;
    const outer = parent ? runItemsOf(parent) as Element[] | undefined : undefined;
    const element = outer?.find((e) => e.value === holder);
    const key = holder ? Object.keys(holder).find((k) => holder[k] === owner) : undefined;
    if (!holder || !outer || !element || !key) throw new Docx4JException('The split could not find the run holder it is inside');
    let next = outer.indexOf(element);
    if (index > 0) {
      const copy: Record<string, unknown> = {};
      for (const [k, v] of Object.entries(holder)) if (k !== key && k !== 'PARENT') copy[k] = v;
      if (revisionKindOf(element) !== undefined) copy.id = p.nextAnnotationId();
      copy[key] = owner.splice(index);
      linkParents(copy[key], copy);
      outer.splice(next + 1, 0, { name: element.name, value: copy } as Element);
      linkParents(outer[next + 1], parent);
    }
    next += index > 0 ? 1 : 0;
    owner = outer;
    index = next;
  }
  const head = content.splice(0, index);
  const el = paragraphOf(head, propertiesOf(p, rPr));
  linkParents(head, el.value);
  const list = listOf(p);
  list.splice(list.indexOf(p.element), 0, el as Element);
  linkParents(el, (p.p as { PARENT?: object }).PARENT);
  p.parentBody.assignParaId(el.value);
  p.changeTracker?.markParagraphInserted(el.value);
  return p.viewOf(el, list);
}

/** The run holders a paragraph split may cut in two: each half is a whole one of the same. */
const SPLITTABLE = new Set(['org_docx4j_wml.P.Hyperlink', 'org_docx4j_wml.RunIns', 'org_docx4j_wml.RunDel', 'org_docx4j_wml.RunTrackChange',
  'org_docx4j_wml.CTSmartTagRun', 'org_docx4j_wml.CTCustomXmlRun', 'org_docx4j_wml.P.Dir', 'org_docx4j_wml.P.Bdo']);

// --- font ---------------------------------------------------------------------------------------------

/** The run options Word writes only where a run doesn't read the value already (check 52: `w:b` on no heading run). */
const TOGGLES = ['bold', 'italic', 'strikeThrough', 'doubleStrikeThrough', 'subscript', 'superscript'] as const;

/** True where every option is a toggle the effective properties already read as given: nothing to write. */
function reads(rPr: wml.RPr | undefined, opts: RunOptions): boolean {
  const read = readRunOptions(rPr) as unknown as Record<string, unknown>;
  return Object.entries(opts).every(([key, value]) => (TOGGLES as readonly string[]).includes(key) && read[key] === value);
}

/** A piece of a range in one paragraph: its runs, and whether the range holds the paragraph's mark. */
export interface FontPiece {
  paragraph: Paragraph;
  start: number;
  end: number;
  runs: () => Element<wml.R>[];
  holdsMark: boolean;
}

/**
 * A `font` write over a range across paragraphs (check 52, cases 49 to 56): every run in the range, a table's
 * cells' too, split at the range's ends, and each mark the range holds, as Word formats them. A toggle (`bold`,
 * `italic`, ...) goes only where a run or mark doesn't read the value already: Word wrote `w:b` on none of a
 * heading's runs, bold through its style, and tracked no change there. Tracked, each run and mark changed has its
 * `w:rPrChange` (case 56: the body paragraph's mark's).
 */
export function formatSpan(pieces: FontPiece[], opts: RunOptions): void {
  for (const piece of pieces) {
    const p = piece.paragraph;
    const resolver = (p.parentBody.package_ as { propertyResolverOrUndefined?: PropertyResolver } | undefined)?.propertyResolverOrUndefined;
    const cell = resolver?.cellContextOf(p.p);
    if (piece.start < piece.end) p.splitRunsAround(piece.start, piece.end);
    const runs = piece.runs().filter((run) => !reads(resolver ? resolver.getEffectiveRPr(run.value.rPr, p.p.pPr, cell) : run.value.rPr, opts));
    if (runs.length > 0) Object.assign(new Font(() => runs.map((run) => run.value), () => p.fontTracking()), opts);
    if (piece.holdsMark && !reads(resolver ? resolver.getEffectiveParagraphMarkRPr(p.p.pPr, cell) : markRPrOf(p.p), opts)) formatMark(p, opts);
  }
}

/** A paragraph mark's direct formatting written; tracked, its `w:rPrChange` first, unless the mark is this author's insertion. */
function formatMark(p: Paragraph, opts: RunOptions): void {
  p.assertWritable();
  const pPr = (p.p.pPr ??= createPPr());
  linkParents(pPr, p.p);
  const rPr = (pPr.rPr ??= createParaRPr());
  linkParents(rPr, pPr);
  const tracker = p.changeTracker;
  if (tracker && rPr.ins?.author !== tracker.author) tracker.recordMarkRPrChange(rPr);
  applyRunOptions(rPr as wml.RPr, opts);
  if (tracker) tracker.settleMarkRPrChange(p.p);
}
