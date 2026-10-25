// A story's text as Office JS's `body.text` gives it, with where each paragraph's text stands in it (CR-014
// phase 1). `Body.text` is this text; a `Range` across paragraphs reads a slice of it, and lists the
// paragraphs between its two ends from it.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { type Element, typeNameOf, childrenOf, rowsOf, cellsOf, selectedBranchOf, BRANCH_TYPES } from './tree.mjs';

/** One paragraph of a story: its element, the array holding it, and its text's place in the story's text. */
export interface StoryParagraph {
  element: Element<wml.P>;
  container: Element[];
  /** Where its text starts and ends in `StoryLayout.text`. */
  start: number;
  end: number;
}

/**
 * A story's text in Word's default view (CR-012 phase 4; checks 30, 31, 48 and 51): each paragraph's text
 * followed by its mark, `\r`, but a mark tracked as deleted (CR-012 section 19, check 58); in a table a row's cells a `\t` apart, each cell's paragraphs a mark apart
 * (its last mark not written), and the row ended by `\r\n`. Unlike `Body.text` the story's last mark is
 * kept, so that every paragraph's text is followed by what Word puts after it.
 */
export interface StoryLayout {
  text: string;
  /** The paragraphs a reader takes, in document order (tables and controls descended into). */
  paragraphs: StoryParagraph[];
  /** The index in `paragraphs` of a `w:p`, or -1 where the walk didn't meet it. */
  indexOf(p: wml.P): number;
}

/**
 * Whether a paragraph's mark is written in Word's default view: not where it is tracked as deleted, which reads as
 * nothing there, the paragraphs staying apart (check 58 cases 6 to 8, CR-012 section 19), nor where it is tracked as
 * moved from, as the paragraph's runs are (CR-002 section 64). A cell's last paragraph is ended by the cell's `\t`
 * whatever its mark (cases 9 and 10).
 */
export const shownInDefaultView = (p: wml.P): boolean => p.pPr?.rPr?.del === undefined && p.pPr?.rPr?.moveFrom === undefined;

/** Whether a paragraph's mark is written in the original view, the marks Reject All keeps: not one tracked as inserted, nor as moved to. */
export const shownInOriginalView = (p: wml.P): boolean => p.pPr?.rPr?.ins === undefined && p.pPr?.rPr?.moveTo === undefined;

/**
 * Lays out the blocks `items`, each paragraph's text read by `textOf` (its `Paragraph.text`, or another view's), its
 * mark written where `markShown` says (Word's default view's rule unless another view's is given).
 */
export function storyLayout(items: Element[], textOf: (element: Element<wml.P>, container: Element[]) => string, markShown: (p: wml.P) => boolean = shownInDefaultView): StoryLayout {
  const pieces: string[] = [];
  let length = 0;
  const push = (piece: string): void => { pieces.push(piece); length += piece.length; };
  const pop = (): void => { length -= pieces.pop()!.length; };
  const paragraphs: StoryParagraph[] = [];
  // Word's default view: the cells a tab apart, the row ended by \r\n (checks 30, 31 and 51; check 48's \t after
  // every cell and row was All Markup's, with hidden text shown)
  const visitRow = (row: object): void => {
    cellsOf(row, 'read').forEach((cell, i) => {
      if (i > 0) push('\t');
      visit(childrenOf(cell.element.value) ?? []);
      if (pieces[pieces.length - 1] === '\r') pop();
    });
    push('\r\n');
  };
  const visit = (blocks: Element[]): void => {
    for (const el of blocks) {
      const tn = typeNameOf(el);
      if (tn === 'org_docx4j_wml.P') {
        const text = textOf(el as Element<wml.P>, blocks);
        paragraphs.push({ element: el as Element<wml.P>, container: blocks, start: length, end: length + text.length });
        push(text);
        if (markShown(el.value as wml.P)) push('\r');
        continue;
      }
      const v = el.value;
      if (typeof v !== 'object' || v === null) continue;
      if (tn === 'org_docx4j_wml.Tbl') {
        for (const row of rowsOf(v, 'read')) visitRow(row.element.value);
        continue;
      }
      // a row-level content control's content: its rows (check 53: its text is its cells a \t apart)
      if (tn === 'org_docx4j_wml.Tr') {
        visitRow(v);
        continue;
      }
      const children = childrenOf(v, 'read');
      if (children) visit(children);
    }
  };
  visit(items);
  const index = new Map(paragraphs.map((p, i) => [p.element.value, i]));
  return { text: pieces.join(''), paragraphs, indexOf: (p) => index.get(p) ?? -1 };
}

/**
 * The story a paragraph stands in: the outermost array of block content above it, and the object holding
 * that array. A text box's paragraphs are a story of their own (the array in its `w:txbxContent`), since a
 * range never crosses a text box's edge; so is a branch of a kept `mc:AlternateContent` that Word doesn't
 * read (CR-013's twin stories), where the selected branch is part of the story around it.
 */
export function storyOf(p: wml.P): { items: Element[]; owner: object } {
  let story: { items: Element[]; owner: object } | undefined;
  for (let o: object | undefined = (p as { PARENT?: object }).PARENT; o !== undefined; o = (o as { PARENT?: object }).PARENT) {
    const tn = (o as { TYPE_NAME?: string }).TYPE_NAME;
    // a paragraph above this one: p is in a text box, whose content is its story
    if (tn === 'org_docx4j_wml.P') break;
    if (tn !== undefined && BRANCH_TYPES.has(tn)) {
      const ac = (o as { PARENT?: object }).PARENT as mce.AlternateContent | undefined;
      const items = (o as { any?: Element[] }).any;
      if (ac && items && selectedBranchOf(ac) !== items) return { items, owner: o };
      continue;
    }
    const content = (o as { content?: unknown }).content;
    if (Array.isArray(content)) story = { items: content as Element[], owner: o };
  }
  if (!story) throw new Docx4JException('This paragraph is in no story: it is not in a body, a cell, a note or a comment');
  return story;
}
