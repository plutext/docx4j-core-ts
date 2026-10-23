// Marshalling a view back to a fragment: what `getOoxml()` and `toApiScript`'s fallback need.
// Copies of the blocks a range holds, its paragraphs trimmed to it, are the one piece the views do
// not offer (Range has no element of its own: it is a span of its paragraphs' text).

import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy, marshalString } from '@docx4j/generated-objects-ts';
import { Body } from '../model/content/Body.mjs';
import { Paragraph } from '../model/content/Paragraph.mjs';
import type { Range } from '../model/content/Range.mjs';
import { extentOf } from '../model/content/spanEdits.mjs';
import { GeneralException } from '../model/content/generalException.mjs';
import type { Element } from '../model/content/tree.mjs';
import { NotSupportedError } from './errors.mjs';

/**
 * The body a detached copy is trimmed through. It reads as the paragraph's own body does (its package's styles
 * decide which runs are hidden, and so the offsets), and writes untracked: trimming a copy is no edit of the
 * document, and under tracking the trim would otherwise mark the rest of the copy deleted. What the trim queues on
 * the package (a picture's relationship, a comment whose reference the copy lost) is settled at save against the
 * document as it stands, which still holds them.
 */
class CopyBody extends Body {
  override get changeTracker(): undefined {
    return undefined;
  }
}

/** A detached copy of `paragraph` holding only its text from `start` to `end`, with the runs and their formatting. */
function trimmed(paragraph: Paragraph, start: number, end: number): Element<wml.P> {
  const copy = deepCopy(paragraph.element) as Element<wml.P>;
  const own = paragraph.parentBody;
  const view = new Paragraph(copy, [copy], new CopyBody(own.part, { content: [copy] }, own.prefix, own.package_));
  const length = view.text.length;
  if (end < length) view.splice(end, length, '');
  if (start > 0) view.splice(0, start, '');
  return copy;
}

/**
 * The blocks a range holds, as detached copies: within one paragraph, that paragraph trimmed to the range; across
 * paragraphs, the first paragraph from the range's start, every block between whole (a table, a content control),
 * and the last paragraph to the range's end, left out where the range ends at its start. The blocks are those
 * `delete()` takes (`extentOf`), so a range starting or ending inside a table or a content control it doesn't hold
 * whole is refused: a fragment of part of one isn't written.
 */
export function blocksOfRange(range: Range): Element[] {
  if (!range.spansParagraphs) return [trimmed(range.paragraph, range.start, range.end)];
  let extent;
  try {
    extent = extentOf(range, 'getOoxml');
  } catch (error) {
    if (!(error instanceof GeneralException) && !(error instanceof NotSupportedError)) throw error;
    throw new NotSupportedError('Range', 'getOoxml', "this range starts or ends inside a table or a content control it doesn't hold whole, "
      + "and a fragment of part of one isn't written; take the whole table or control, or a range inside one cell");
  }
  const out: Element[] = [];
  for (let k = extent.first; k <= extent.last; k++) {
    const block = extent.list[k]!;
    if (k === extent.first && extent.head) out.push(trimmed(extent.head, range.start, extent.head.text.length));
    else if (k === extent.last && extent.tail) {
      if (range.end > 0) out.push(trimmed(extent.tail, 0, range.end));
    } else out.push(deepCopy(block) as Element);
  }
  return out;
}

/** The range as a fragment: its blocks' XML, one after another (extension; Office JS wraps a range in a pkg:package). */
export async function xmlOfRange(range: Range): Promise<string> {
  const xml = await Promise.all(blocksOfRange(range).map((block) => marshalString(block)));
  return xml.join('');
}
