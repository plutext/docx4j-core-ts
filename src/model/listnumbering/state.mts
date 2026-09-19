// The counters (docx4j `NumberingState` and `NumberingStates`, CR-014 phase 4).
//
// Definitions hold no counters: a traversal owns its states, one per **story**, and passes the
// right one to the `Emulator`.  So two walks of one package do not interleave, a footer's list
// starts at its own start whenever the footer is read, and `Emulator.peek` can ask what a
// paragraph's number would be without taking it.
//
// Counters are keyed by the **referencing** `w:abstractNum` and the level, which is the sharing
// rule Word applies: every `w:num` over one abstract definition continues a single sequence
// (CR-014 probe P1, measured), and a `w:numStyleLink` abstract definition is a list of its own
// (P2).  A `w:num` with a `w:startOverride` for a level resets that shared counter the first
// time that `w:num` is met at that level, in this story.
import { ContentTypes } from '../../opc/ContentTypes.mjs';

/** One level's count in one {@link NumberingState}. docx4j `ListLevel.Counter`. */
export class Counter {

  /** The current count. */
  value: number;

  /** The level has been counted in this state; its start value is spent. */
  encounteredAlready = false;

  /**
   * A shallower level reset it: it holds its start value, and the next use of the level takes
   * that value rather than incrementing (CR-014 probe P8, measured - a level-2 item straight
   * after a level-0 one reads "2.1.1", where a counter left at start-1 read "2.0.1").
   */
  resetPending = false;

  constructor(value = 0) {
    this.value = value;
  }

  /** One more item at this level. */
  increment(): void {
    this.value += 1;
  }

  /** An independent copy; what {@link NumberingState.copy} is made of. */
  copy(): Counter {
    const other = new Counter(this.value);
    other.encounteredAlready = this.encounteredAlready;
    other.resetPending = this.resetPending;
    return other;
  }

  toString(): string {
    return `${this.value}${this.encounteredAlready ? '' : ' (unused)'}${this.resetPending ? ' (reset)' : ''}`;
  }
}

/**
 * The counters of one story of one traversal. docx4j `NumberingState`.
 *
 * Not shared: one state belongs to one walk of one story.
 */
export class NumberingState {

  private readonly byLevel = new Map<string, Counter>();
  private readonly overridesApplied = new Set<string>();

  /**
   * The counter of a level of an abstract list, created at `initial` (the level's start value
   * less one) on first use.
   */
  counter(abstractNumId: string, ilvl: string, initial: number | undefined): Counter {
    const key = `${abstractNumId}/${ilvl}`;
    let found = this.byLevel.get(key);
    if (found === undefined) {
      found = new Counter(initial ?? 0);
      this.byLevel.set(key, found);
    }
    return found;
  }

  /**
   * Every counter this story stands at, keyed `"<abstractNumId>/<ilvl>"`. docx4j
   * `NumberingState.counters()`, one of the parity accessors it made public in 17.1.1 so that a
   * port can be held to the *counters* and not only to the labels. The live map, as docx4j's
   * unmodifiable view is of the live one: do not hold it across a later `getNumber`.
   */
  get counters(): ReadonlyMap<string, Counter> {
    return this.byLevel;
  }

  /**
   * The `w:num` start overrides spent, keyed `"<numId>/<ilvl>"`. docx4j
   * `NumberingState.startOverridesApplied()`. A `w:num`'s first use at a level is recorded here
   * whether or not it carries a `w:startOverride`, which is docx4j's own behaviour (the flag is
   * set wherever the counter takes a start value), and what its goldens show.
   */
  get startOverridesApplied(): ReadonlySet<string> {
    return this.overridesApplied;
  }

  /** Whether this `w:num`'s `w:startOverride` has been spent at a level. */
  startOverrideApplied(numId: string, ilvl: string): boolean {
    return this.overridesApplied.has(`${numId}/${ilvl}`);
  }

  /** Remember that it has. */
  markStartOverrideApplied(numId: string, ilvl: string): void {
    this.overridesApplied.add(`${numId}/${ilvl}`);
  }

  /** Every list starts again, as at the head of a story. */
  reset(): void {
    this.byLevel.clear();
    this.overridesApplied.clear();
  }

  /** Nothing has been numbered in this state yet. */
  get isEmpty(): boolean {
    return this.byLevel.size === 0 && this.overridesApplied.size === 0;
  }

  /** An independent copy: what `Emulator.peek` numbers against. */
  copy(): NumberingState {
    const other = new NumberingState();
    for (const [key, counter] of this.byLevel) other.byLevel.set(key, counter.copy());
    for (const key of this.overridesApplied) other.overridesApplied.add(key);
    return other;
  }

  toString(): string {
    return `NumberingState(${[...this.byLevel].map(([k, c]) => `${k}=${c}`).join(', ')})`;
  }
}

/** As much of a `Part` as {@link NumberingStates.forPart} reads. */
export interface PartLike {
  readonly contentType: string;
}

/** Which story a part's paragraphs number in. */
export type StoryKind = 'main' | 'headersFooters' | 'own';

/**
 * The story a part belongs to, by content type (never by class, so that this module imports no
 * part): every header and footer is one story, the footnotes, endnotes and comments parts each
 * their own, everything else the main story.
 */
export function storyKindOf(part: PartLike | undefined): StoryKind {
  switch (part?.contentType) {
    case ContentTypes.WORDPROCESSINGML_HEADER:
    case ContentTypes.WORDPROCESSINGML_FOOTER:
      return 'headersFooters';
    case ContentTypes.WORDPROCESSINGML_FOOTNOTES:
    case ContentTypes.WORDPROCESSINGML_ENDNOTES:
    case ContentTypes.WORDPROCESSINGML_COMMENTS:
      return 'own';
    default:
      return 'main';
  }
}

/**
 * The numbering states of one traversal, one per story. docx4j `NumberingStates`.
 *
 * Word numbers each story from its own counters (CR-014 probe P7, measured 2026-09-12): the body
 * is one story; a section's header and footer share one (header 1 2 3, footer 4 5 6); the
 * footnotes part is one (the second footnote continues the first's count); the endnotes part
 * another; each text box and each comment counts on its own; and the body's count runs past all
 * of them untouched.
 *
 * Only one section was probed, so headers and footers of every section are folded into one story
 * here until a document says otherwise.
 */
export class NumberingStates {

  private readonly mainState = new NumberingState();
  private headersFooters: NumberingState | undefined;
  private readonly byPart = new Map<PartLike, NumberingState>();

  /** The main document's story. */
  main(): NumberingState {
    return this.mainState;
  }

  /**
   * The story the given part's paragraphs number in: the main document (also for undefined, and
   * for any part not listed in {@link storyKindOf}), the one shared by headers and footers, or
   * the footnotes, endnotes or comments part's own.
   */
  forPart(part: PartLike | undefined): NumberingState {
    switch (storyKindOf(part)) {
      case 'headersFooters':
        this.headersFooters ??= new NumberingState();
        return this.headersFooters;
      case 'own': {
        let found = this.byPart.get(part!);
        if (found === undefined) {
          found = new NumberingState();
          this.byPart.set(part!, found);
        }
        return found;
      }
      default:
        return this.mainState;
    }
  }

  /** A fresh story: a text box's, or a comment's when comments are numbered one by one. */
  newStory(): NumberingState {
    return new NumberingState();
  }
}
