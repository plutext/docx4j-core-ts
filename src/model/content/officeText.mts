// Office JS's characters written back (CR-002 section 60). The readers give Word's character stream (CR-012 phase 4,
// `officeItemText`): a `w:br` U+000B, a page break U+000C, a column break U+000E, a `w:cr` U+000D, the non-breaking and
// optional hyphens U+001E and U+001F, a note reference U+0002. A writer turns each back into the element it reads as
// (D1, as Word 365's writers do in check 55), drops the characters XML 1.0 excludes (E3: Word drops them too), and
// takes a line feed or carriage return as its kind of writer does: a paragraph break for `insertText` (E1), a break
// for `insertParagraph` (E2), a break or nothing in a bound or plain-text control (section 61 item 3, F3).
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { t as textItem, br as breakItem, tab as tabItem } from '@docx4j/generated-objects-ts/builders/wml';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { runOf, type Element } from './tree.mjs';

/**
 * The characters XML 1.0 excludes that are no element: U+0002 (a note reference, which no text makes) and the other C0
 * controls but tab, line feed, carriage return and the five that are elements, U+FFFE, U+FFFF and an unpaired surrogate.
 */
const EXCLUDED = /[\u0000-\u0008\u000F-\u001D￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

/** The characters a writer writes as elements: tab, the three breaks, a carriage return, a line feed, the two hyphens. */
const ELEMENT_CHARACTERS = /[\t\n\r\u000B\u000C\u000E\u001E\u001F]/;

/** The text without the characters XML 1.0 excludes (E3). */
export function dropExcluded(text: string): string {
  return text.replace(EXCLUDED, '');
}

/** True when the text holds a character a writer writes as an element. */
export function holdsElementCharacters(text: string): boolean {
  return ELEMENT_CHARACTERS.test(text);
}

/**
 * How a writer takes a line feed and a carriage return:
 * - `paragraphs`: a paragraph break each, `\r\n` one (E1: `insertText` and every writer through `Paragraph.splice`);
 *   the caller splits the text into lines (`linesOf`);
 * - `one-paragraph`: `\n` and `\r\n` a `w:br`, `\r` a `w:cr` (E2: `insertParagraph`, `Paragraph.text =`), a recorded
 *   departure from Word, which writes their escapes as text;
 * - `multi-line`: each a `w:br` (a multi-line plain-text control, or a bound block control of one: section 61 item 3);
 * - `single-line`: none, and no break either (a single-line plain-text control: F3).
 */
export type LineBreaks = 'paragraphs' | 'one-paragraph' | 'multi-line' | 'single-line';

/** The text a writer of that kind writes, its line feeds and carriage returns turned into its characters. */
export function forWriter(text: string, lineBreaks: LineBreaks): string {
  switch (lineBreaks) {
    case 'one-paragraph': return text.replace(/\r\n|\n/g, '\u000B');
    case 'multi-line': return text.replace(/\r\n|\r|\n/g, '\u000B');
    case 'single-line': return text.replace(/\r\n|[\r\n\u000B\u000C\u000E]/g, '');
    default: return text;
  }
}

/** The text's lines, at each line feed, carriage return or `\r\n` (E1). */
export function linesOf(text: string): string[] {
  return text.split(/\r\n|\r|\n/);
}

/** The element a character reads as (`officeItemText` in reverse); a line feed a `w:br`. */
function elementFor(c: string): Element {
  switch (c) {
    case '\t': return tabItem() as Element;
    case '\u000C': return breakItem('page') as Element;
    case '\u000E': return breakItem('column') as Element;
    case '\r': return el.cr({}) as Element;
    case '\u001E': return el.noBreakHyphen({}) as Element;
    case '\u001F': return el.softHyphen({}) as Element;
    default: return breakItem() as Element;   // U+000B, and a line feed a writer left
  }
}

/** A run's items for a text: a `w:t` for each plain stretch, the element each mapped character reads as. */
export function officeItems(text: string): Element[] {
  const out: Element[] = [];
  let plain = '';
  for (const c of text) {
    if (ELEMENT_CHARACTERS.test(c)) {
      if (plain !== '') out.push(textItem(plain) as Element);
      plain = '';
      out.push(elementFor(c));
    } else plain += c;
  }
  if (plain !== '') out.push(textItem(plain) as Element);
  return out;
}

/**
 * A run of text a writer keeps in one paragraph (E2: `insertParagraph`, `Paragraph.text =`), or in a plain-text control
 * as its kind takes it: the excluded characters dropped, every other character that is an element written as it.
 */
export function oneParagraphRun(text: string, rPr?: wml.RPr, lineBreaks: LineBreaks = 'one-paragraph'): Element<wml.R> {
  return runOf(officeItems(forWriter(dropExcluded(text), lineBreaks)), rPr);
}

