import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { Font } from './Font.mjs';
import { runFontSelectorOf } from '../fonts/lookup.mjs';
import type { Paragraph, FormattingOptions } from './Paragraph.mjs';
import type { BlockElement } from './Body.mjs';
import { type Element, typeNameOf, linkParents, runItemsOf, runsOf, type TextViewOptions } from './tree.mjs';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { ContentControl, type ContentControlType } from './ContentControl.mjs';
import { sdt as sdtOf, nextSdtId, applyRunOptions, t as textElement } from '@docx4j/generated-objects-ts/builders/wml';
import * as f from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { hyperlink as hyperlinkOf } from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { Namespaces } from '../../parts/Namespaces.mjs';
import { controlIdScope, sdtKindFor } from '../customxml/insert.mjs';
import { contentOf } from './ooxml.mjs';
import { searchPattern, findAll, matchesOf, expandReplacement, type SearchOptions } from './search.mjs';
import { commentApi, type CommentContent, type CommentOptions } from './comments.mjs';
import type { Comment } from './Comment.mjs';
import type { TrackedChange } from './TrackedChange.mjs';

/**
 * A subset of Office JS `Word.Range`: a span of text within one paragraph, [start, end) in the
 * paragraph's text. Positions are recomputed from the tree on every use, so a range stays
 * valid across edits made through it; edits made elsewhere may shift it.
 */
export class Range {
  constructor(readonly paragraph: Paragraph, public start: number, public end: number) {}

  get paragraphs(): Paragraph[] {
    return [this.paragraph];
  }

  get text(): string {
    return this.paragraph.text.substring(this.start, this.end);
  }
  set text(value: string) {
    this.insertText(value, 'Replace');
  }

  /** The paragraph's style id (extension, as on Paragraph). */
  get styleId(): string {
    return this.paragraph.styleId;
  }
  set styleId(id: string) {
    this.paragraph.styleId = id;
  }

  /** The paragraph's style name (Office JS reports the range's paragraph style). */
  get style(): string {
    return this.paragraph.style;
  }
  set style(id: string) {
    this.paragraph.style = id;
  }

  /**
   * The text in one of the two views (extension; `text` is the accepted one). `{ view: 'original' }`
   * is offered on `Paragraph` and `Body`; on a `Range` it is the accepted-view span only, since a
   * range's offsets are accepted-view offsets.
   */
  getText(options?: TextViewOptions): string {
    if (options?.view === 'original') throw new Docx4JException("A Range's offsets are accepted-view offsets; read the original view on its Paragraph");
    return this.text;
  }

  /**
   * The formatting of exactly this span: reads the first run's *effective* properties, writes
   * direct formatting, splitting the runs at the boundaries so that a write touches only the
   * span.
   */
  get font(): Font {
    return this.getFont();
  }

  /** `font`, with `{ direct: true }` for the direct formatting of the first run (extension). */
  getFont(options?: FormattingOptions): Font {
    const holders = (): wml.R[] => {
      if (this.start === this.end) return [];
      this.paragraph.splitAt(this.start);
      this.paragraph.splitAt(this.end);
      return this.runs.map((r) => r.value);
    };
    if (options?.direct === true) return new Font(holders, () => this.paragraph.fontTracking());
    return new Font(holders, () => this.paragraph.fontTracking(),
      (rPr) => this.paragraph.parentBody.propertyResolver.getEffectiveRPr(rPr, this.paragraph.p.pPr),
      (rPr) => runFontSelectorOf(this.paragraph.parentBody.package_)?.asciiFontName(rPr));
  }

  /** The runs the span covers (a run partly inside counts). */
  get runs(): Element<wml.R>[] {
    const out: Element<wml.R>[] = [];
    const seen = new Set<wml.R>();
    for (const seg of this.paragraph.segments()) {
      if (seg.end <= this.start || seg.start >= this.end) continue;
      if (seen.has(seg.run)) continue;
      seen.add(seg.run);
      out.push(seg.runOwner[seg.runIndex] as Element<wml.R>);
    }
    return out;
  }

  insertText(text: string, location: 'Before' | 'After' | 'Start' | 'End' | 'Replace'): Range {
    let r: Range;
    if (location === 'Replace') r = this.paragraph.splice(this.start, this.end, text);
    else if (location === 'Before' || location === 'Start') r = this.paragraph.splice(this.start, this.start, text);
    else r = this.paragraph.splice(this.end, this.end, text);
    // this range keeps covering the same text: a replacement resizes it, text before it shifts it, text after it leaves it
    if (location === 'Replace') this.end = this.start + text.length;
    else if (location === 'Before' || location === 'Start') { this.start += text.length; this.end += text.length; }
    return r;
  }

  insertParagraph(text: string, location: 'Before' | 'After'): Paragraph {
    return this.paragraph.insertParagraph(text, location);
  }

  /**
   * Word's `insertOoxml` at range level: a flat OPC `pkg:package` (or a bare fragment). One
   * incoming paragraph's runs go at the start or the end of the span, or in its place; anything
   * else is inserted as blocks before or after the paragraph. Returns what was inserted, as
   * `insertXml` does.
   */
  async insertOoxml(ooxml: string, location: 'Before' | 'After' | 'Replace'): Promise<(Paragraph | BlockElement)[]> {
    const body = this.paragraph.parentBody;
    const preprocess = body.package_?.loadOptions.preprocessor;
    const elements = await contentOf(ooxml, { preprocess: preprocess ? (doc) => preprocess(doc) : undefined, target: body.part });
    if (elements.length === 0) return [];
    const only = elements.length === 1 && typeNameOf(elements[0]!) === 'org_docx4j_wml.P' ? elements[0]!.value as { content?: Element[] } : undefined;
    if (only) {
      if (location === 'Replace') { this.delete(); this.paragraph.insertItemsAt(this.start, (only.content ?? [])); }
      else this.paragraph.insertItemsAt(location === 'Before' ? this.start : this.end, (only.content ?? []));
      return [this.paragraph];
    }
    body.insertElement(elements, location === 'Before' ? 'Before' : 'After', this.paragraph);
    if (location === 'Replace') this.delete();
    return elements.map((el) => typeNameOf(el) === 'org_docx4j_wml.P'
      ? body.paragraphFor(el.value as never) ?? { element: el, container: this.paragraph.container }
      : { element: el, container: this.paragraph.container });
  }

  /** Removes the span's text. */
  delete(): void {
    this.paragraph.splice(this.start, this.end, '');
    this.end = this.start;
  }

  search(text: string, options?: SearchOptions): Range[] {
    const base = this.start;
    return findAll(this.text, searchPattern(text, options)).map(([s, e]) => new Range(this.paragraph, base + s, base + e));
  }

  /**
   * Replaces every match in this span, last first so the offsets stay valid; returns the count
   * (CR-002 section 3.7). The replacement's `$1` or Word's `\1` expands as it does on a paragraph
   * (section 26; this one was missed there, section 31).
   */
  replaceText(find: string, replace: string, options?: SearchOptions): number {
    const base = this.start;
    const matches = matchesOf(this.text, searchPattern(find, options));
    for (let i = matches.length - 1; i >= 0; i--) {
      const match = matches[i]!;
      new Range(this.paragraph, base + match.index, base + match.index + match[0].length)
        .insertText(expandReplacement(replace, match, options), 'Replace');
    }
    return matches.length;
  }

  /** The tracked changes this span covers, in document order (CR-002 phase F). */
  getTrackedChanges(): TrackedChange[] {
    return this.paragraph.getTrackedChanges().filter((change) => {
      const range = change.getRange();
      if (!range) return false;
      return range.start <= this.end && range.end >= this.start;
    });
  }

  getRange(location: 'Whole' | 'Start' | 'End' | 'Content' = 'Whole'): Range {
    if (location === 'Start') return new Range(this.paragraph, this.start, this.start);
    if (location === 'End') return new Range(this.paragraph, this.end, this.end);
    return new Range(this.paragraph, this.start, this.end);
  }

  /** The span's text as it stands (extension; Office JS getOoxml wraps it in a package). */
  toString(): string {
    return this.text;
  }

  // --- comments (CR-002 phase G) ---

  /** The comments this span touches, replies nested under their parent. */
  async getComments(): Promise<Comment[]> {
    return commentApi().commentsOf(this);
  }

  /**
   * Comments the span: the markers and the reference run around it (runs are split at the
   * boundaries, as `font` does), the comment itself and the side-part entries, creating any of
   * the comment parts the document lacks. Asynchronous because it unmarshals those parts.
   */
  async insertComment(content: CommentContent, options?: CommentOptions): Promise<Comment> {
    return commentApi().insertComment(this, content, options);
  }

  // --- content controls (CR-002 phase E) ---

  /**
   * Wraps the runs this span covers in a run-level content control (Office JS
   * `Range.insertContentControl`), splitting the runs at the boundaries as `font` does. An empty
   * span gets an empty control at its position. Returns the control.
   */
  insertContentControl(kind?: ContentControlType): ContentControl {
    const paragraph = this.paragraph;
    const body = paragraph.parentBody;
    // The control is built empty first: `sdt` refuses a kind that cannot be run-level (a repeating
    // section), and that refusal must come before any run is split.
    const sdt = sdtOf([], { kind: sdtKindFor(kind), id: nextSdtId(controlIdScope(body)), form: 'run' }) as Element<wml.SdtRun>;
    if (this.start === this.end) {
      paragraph.insertItemsAt(this.start, [sdt as Element]);
      return new ContentControl(sdt, containerOf(sdt, paragraph), body);
    }
    paragraph.splitAt(this.start);
    paragraph.splitAt(this.end);
    const segments = paragraph.segments().filter((s) => s.start >= this.start && s.end <= this.end);
    if (segments.length === 0) throw new Docx4JException('This range covers no run');
    const owner = segments[0]!.runOwner;
    if (segments.some((s) => s.runOwner !== owner)) {
      throw new Docx4JException('This range spans more than one run holder (a hyperlink or a tracked change); wrap a narrower span');
    }
    const items: Element[] = [];
    for (const segment of segments) {
      const element = owner[segment.runIndex] as Element;
      if (!items.includes(element)) items.push(element);
    }
    const at = owner.indexOf(items[0]!);
    (sdt.value.sdtContent as { content?: Element[] }).content = items;
    owner.splice(at, items.length, sdt as Element);
    linkParents(sdt, (segments[0]!.run as { PARENT?: object }).PARENT ?? paragraph.p);
    return new ContentControl(sdt, owner, body);
  }

  /**
   * Office JS `Range.hyperlink`: the address of the first hyperlink in this range, or a new
   * hyperlink over it. `address#location` separates the address from an optional location within
   * it, so `"#heading"` is a link inside this document (a `w:anchor`) and
   * `"https://example.com"` an external one (a relationship). Reading gives `""` where the range
   * is in no hyperlink.
   *
   * Setting removes the hyperlinks the range already carries and wraps its runs in one
   * `w:hyperlink`, splitting runs at the boundaries as `font` and `insertContentControl` do;
   * setting `""` removes them and wraps nothing. An external address becomes a relationship of
   * this range's part, which is why the part must exist (CR-002 section 20). Only the hyperlinks
   * the range touches are removed; before 0.2.1 every one in the paragraph was (section 25). An
   * empty range is refused with a `GeneralException`, as Office JS refuses it (section 25, check 19):
   * to change a link under a caret, set it on the link's whole range.
   *
   * The runs are styled as Word styles them, measured through Office JS (`test/README.md` check
   * 15, CR-002 section 25): each run a link wraps takes the `Hyperlink` character style, over any
   * character style it had, and loses its direct colour, keeping the rest of its direct formatting;
   * a link removed takes the style off its runs again. The `Hyperlink` definition is added to the
   * styles part when the document lacks it, at the latest when the package is saved - a setter
   * cannot await it - so `await pkg.styles.ensure('Hyperlink')` first where an effective read
   * before the save must see it. Under change tracking the restyle is a formatting revision, as a
   * `font` write is.
   */
  get hyperlink(): string {
    const holder = this.hyperlinkHolder();
    if (holder === undefined) return '';
    const anchor = holder.value.anchor ?? '';
    const address = holder.value.id === undefined ? '' : this.relationshipTarget(holder.value.id);
    return anchor === '' ? address : `${address}#${anchor}`;
  }

  set hyperlink(value: string) {
    // Office JS refuses a hyperlink set on an empty range - inside a link, in plain text, just after a
    // link, and "" alike - with GeneralException, and changes nothing (test/README.md check 19, Word
    // 16.0.20326.20158; CR-002 section 25). Before, the setter removed the link under a caret and
    // left an empty w:hyperlink there.
    if (this.start === this.end) {
      throw new GeneralException('A hyperlink cannot be set on an empty range, as in Office JS: select the text to link, or the whole of a link to change it');
    }
    const paragraph = this.paragraph;
    // A range that crosses the edge of a link takes the whole of that link into the new one; a range
    // inside a link replaces it with a link over the range alone. Measured in Word through Office JS
    // (test/README.md check 16, cases F and G; CR-002 section 25): "alpha beta" linked and then
    // "beta gamma" is one link over "alpha beta gamma", where "alpha beta gamma" linked and then
    // "beta" leaves "alpha " and " gamma" unlinked.
    let start = this.start;
    let end = this.end;
    for (const holder of hyperlinkHolders(paragraph)) {
      const span = this.spanOf(holder);
      if (span === undefined || !this.touchesSpan(span)) continue;
      if (span[0] <= this.start && this.end <= span[1]) continue;
      start = Math.min(start, span[0]);
      end = Math.max(end, span[1]);
    }
    this.removeHyperlinks();
    if (value === '') return;

    const hash = value.indexOf('#');
    const address = hash === -1 ? value : value.slice(0, hash);
    const location = hash === -1 ? '' : value.slice(hash + 1);
    // w:history="1" is what Word writes on every hyperlink it makes, and what docx4j writes when
    // it builds one (MainDocumentPart.addHyperlink): "add the target to the viewed hyperlinks"
    // (ECMA-376 17.16.22). Readers do not depend on it, but a document holding one holder from
    // each source should not differ over it. Found by the docx4j-ts-editor session, ED-003
    // section 11.8.
    const holder = hyperlinkOf({ content: [], history: true }) as Element<wml.P.Hyperlink>;
    if (location !== '') holder.value.anchor = location;
    if (address !== '') holder.value.id = this.addHyperlinkRelationship(address);

    paragraph.splitAt(start);
    paragraph.splitAt(end);
    const segments = paragraph.segments().filter((seg) => seg.start >= start && seg.end <= end);
    if (segments.length === 0) throw new Docx4JException('This range covers no run');
    const owner = segments[0]!.runOwner;
    if (segments.some((seg) => seg.runOwner !== owner)) {
      throw new Docx4JException('This range spans more than one run holder (a hyperlink or a tracked change); wrap a narrower span');
    }
    const items: Element[] = [];
    for (const segment of segments) {
      const element = owner[segment.runIndex] as Element;
      if (!items.includes(element)) items.push(element);
    }
    const at = owner.indexOf(items[0]!);
    holder.value.content = items as never;
    owner.splice(at, items.length, holder as Element);
    linkParents(holder, (segments[0]!.run as { PARENT?: object }).PARENT ?? paragraph.p);
    this.restyle(holder.value, true);
    mergeTextRuns(holder.value.content as Element[], holder.value);
    (paragraph.parentBody.package_ as { requireStyles?: (ids: string[]) => void } | undefined)
      ?.requireStyles?.([HYPERLINK_STYLE]);
  }

  /** The first `w:hyperlink` this range touches, if any. */
  private hyperlinkHolder(): Element<wml.P.Hyperlink> | undefined {
    return hyperlinkHolders(this.paragraph).find((holder) => this.touches(holder));
  }

  /**
   * Unwraps every `w:hyperlink` the range touches, keeping its runs where they are and taking the
   * `Hyperlink` style off them. Until 0.2.1 this unwrapped every hyperlink of the paragraph, so
   * linking one word unlinked the others (CR-002 section 25).
   */
  private removeHyperlinks(): void {
    for (const holder of hyperlinkHolders(this.paragraph).filter((h) => this.touches(h))) {
      const owner = containerOf(holder as Element, this.paragraph);
      const at = owner.indexOf(holder as Element);
      if (at === -1) continue;
      this.restyle(holder.value, false);
      owner.splice(at, 1, ...((holder.value.content ?? []) as Element[]));
      for (const item of (holder.value.content ?? []) as Element[]) {
        linkParents(item, (holder as { PARENT?: object }).PARENT ?? this.paragraph.p);
      }
    }
  }

  /**
   * Whether this range touches a hyperlink's text: overlaps it, or for an empty range lies in it or
   * at either end of it - Office JS reads a caret just after a link as in that link (check 19, case
   * K; the start edge was not measured). A hyperlink holding no text touches nothing.
   */
  private touches(holder: Element<wml.P.Hyperlink>): boolean {
    const span = this.spanOf(holder);
    return span !== undefined && this.touchesSpan(span);
  }

  private touchesSpan([start, end]: [number, number]): boolean {
    return this.start === this.end ? start <= this.start && this.start <= end : start < this.end && end > this.start;
  }

  /** The text a hyperlink holds, as [start, end) in the paragraph, or undefined when it holds none. */
  private spanOf(holder: Element<wml.P.Hyperlink>): [number, number] | undefined {
    let start: number | undefined;
    let end = 0;
    for (const segment of this.paragraph.segments()) {
      if (!within(segment.run, holder.value, this.paragraph.p)) continue;
      start ??= segment.start;
      end = segment.end;
    }
    return start === undefined ? undefined : [start, end];
  }

  /**
   * The `Hyperlink` character style on every run of a hyperlink, in place of any character style
   * and with the direct colour removed, or off again when the link goes: what Word does (CR-002
   * section 25). A run whose properties are then empty loses its `w:rPr`, as in Word's output.
   */
  private restyle(holder: wml.P.Hyperlink, linked: boolean): void {
    const tracking = this.paragraph.fontTracking();
    const runs = new Set([...runsOf(holder), ...runsOf(holder, { view: 'original' })].map((r) => r.value));
    for (const run of runs) {
      if (!linked && run.rPr?.rStyle?.val !== HYPERLINK_STYLE) continue;
      run.rPr ??= { TYPE_NAME: 'org_docx4j_wml.RPr' };
      // the properties as they are now become w:rPrChange, unless this run is our own insertion
      if (tracking && !tracking.ownInsertions.has(run)) tracking.tracker.recordRPrChange(run.rPr);
      applyRunOptions(run.rPr, linked ? { style: HYPERLINK_STYLE, color: '' } : { style: '' });
      if (Object.keys(run.rPr).every((key) => key === 'TYPE_NAME')) delete run.rPr;
      else linkParents(run.rPr, run);
    }
  }

  /** The target of a relationship of this range's part, for the getter. */
  private relationshipTarget(id: string): string {
    const part = this.paragraph.parentBody.part;
    return part?.relationshipsPart?.getRelationshipById(id)?.target ?? '';
  }

  /** An external relationship of this range's part for the address, and its id. */
  private addHyperlinkRelationship(address: string): string {
    const part = this.paragraph.parentBody.part;
    if (!part) throw new Docx4JException('This range has no part, so a hyperlink cannot be related to it');
    return part.getRelationshipsPart(true)!.addExternalRelationship(Namespaces.HYPERLINK, address).id;
  }
}

/**
 * What Office JS throws where the host refuses an operation it has no answer for: its `code`, so that
 * add-in code branching on `error.code` behaves the same here (check 19).
 */
class GeneralException extends Docx4JException {
  readonly code = 'GeneralException';
  constructor(message: string) {
    super(message);
    this.name = 'GeneralException';
  }
}

/** The character style Word gives the runs of a hyperlink it makes. */
const HYPERLINK_STYLE = 'Hyperlink';

/**
 * Adjacent runs holding nothing but text and the same properties become one, as Word writes a link
 * it makes: linking "beta gamma" over the end of a link on "alpha beta" gives Word one run
 * "alpha beta gamma" (`test/README.md` check 16, case F), where splitting and restyling left two.
 * Runs whose properties differ at all - a `w:rPrChange` under tracking among them - stay apart.
 */
function mergeTextRuns(items: Element[], parent: object): void {
  for (let i = items.length - 1; i > 0; i--) {
    const [a, b] = [items[i - 1]!, items[i]!];
    if (!isTextRun(a) || !isTextRun(b)) continue;
    const [ra, rb] = [a.value as wml.R, b.value as wml.R];
    if (canonical(ra.rPr) !== canonical(rb.rPr)) continue;
    const text = [...(ra.content ?? []), ...(rb.content ?? [])].map((item) => (item.value as wml.Text).value ?? '').join('');
    ra.content = [textElement(text)] as never;
    linkParents(ra.content, ra);
    items.splice(i, 1);
  }
  linkParents(items, parent);
}

function isTextRun(item: Element): boolean {
  return typeNameOf(item) === 'org_docx4j_wml.R'
    && ((item.value as wml.R).content ?? []).every((c) => (c as Element).name?.localPart === 't');
}

/** A value as text with its keys in order and `PARENT` left out, to compare run properties by content. */
function canonical(value: unknown): string {
  return JSON.stringify(value, (key, v) => {
    if (key === 'PARENT') return undefined;
    if (v === null || typeof v !== 'object' || Array.isArray(v)) return v;
    return Object.fromEntries(Object.keys(v).sort().map((k) => [k, (v as Record<string, unknown>)[k]]));
  });
}

/** Whether `object` is `ancestor` or below it, following `PARENT` no further than `stop`. */
function within(object: object, ancestor: object, stop: object): boolean {
  for (let o: object | undefined = object; o !== undefined && o !== stop; o = (o as { PARENT?: object }).PARENT) {
    if (o === ancestor) return true;
  }
  return false;
}

/** Every `w:hyperlink` of a paragraph, outermost first (they do not nest in practice). */
function hyperlinkHolders(paragraph: Paragraph): Element<wml.P.Hyperlink>[] {
  const out: Element<wml.P.Hyperlink>[] = [];
  const visit = (items: Element[] | undefined): void => {
    for (const item of items ?? []) {
      if (typeNameOf(item) === 'org_docx4j_wml.P.Hyperlink') out.push(item as Element<wml.P.Hyperlink>);
      const value = item.value;
      if (typeof value === 'object' && value !== null) visit(runItemsOf(value) as Element[] | undefined);
    }
  };
  visit(runItemsOf(paragraph.p) as Element[] | undefined);
  return out;
}

/** The run-level array holding an element, after `insertItemsAt` put it somewhere in the paragraph. */
function containerOf(element: Element, paragraph: Paragraph): Element[] {
  const seek = (items: Element[] | undefined): Element[] | undefined => {
    if (!items) return undefined;
    if (items.includes(element)) return items;
    for (const item of items) {
      const value = item.value;
      if (typeof value !== 'object' || value === null) continue;
      const found = seek(runItemsOf(value));
      if (found) return found;
    }
    return undefined;
  };
  return seek(runItemsOf(paragraph.p)) ?? ((paragraph.p.content ??= []) as Element[]);
}
