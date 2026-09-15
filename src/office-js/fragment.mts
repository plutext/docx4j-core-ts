// Marshalling a view back to a fragment: what `getOoxml()` and `toApiScript`'s fallback need.
// A `w:p` copy holding only a range's text is the one piece the views do not offer (Range has no
// element of its own: it is a span of the paragraph's text).

import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy, marshalString } from '@docx4j/generated-objects-ts';
import { Paragraph } from '../model/content/Paragraph.mjs';
import type { Range } from '../model/content/Range.mjs';
import type { Element } from '../model/content/tree.mjs';

/**
 * A detached `w:p` holding exactly the range's text, with the runs and their formatting. The
 * paragraph is copied first, so the document is untouched.
 */
export function paragraphElementOfRange(range: Range): Element<wml.P> {
  const copy = deepCopy(range.paragraph.element) as Element<wml.P>;
  const view = new Paragraph(copy, [copy], range.paragraph.parentBody);
  const length = view.text.length;
  if (range.end < length) view.splice(range.end, length, '');
  if (range.start > 0) view.splice(0, range.start, '');
  return copy;
}

/** The range as a `w:p` fragment (extension; Office JS wraps a range in a pkg:package). */
export function xmlOfRange(range: Range): Promise<string> {
  return marshalString(paragraphElementOfRange(range) as Element);
}
