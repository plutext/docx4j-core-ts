// Members the shim adds over a view, per Word class (CR-002 section 3.10). The views keep to what
// the content API promises (section 3.12), so `getOoxml()`, which needs the package, lives here
// rather than on `Body`; the proxy offers these only when the view itself has no such member, so
// a later phase that implements one on the view wins without a change here.

import type { Body } from '../model/content/Body.mjs';
import type { Paragraph } from '../model/content/Paragraph.mjs';
import type { Range } from '../model/content/Range.mjs';
import { xmlOfRange } from './fragment.mjs';

/** A member the shim adds: called with the view and the call's arguments. */
export type ExtraMember = (target: never, ...args: never[]) => unknown;

/** By Word class name, then member name. */
export type ExtraMembers = Readonly<Record<string, Readonly<Record<string, ExtraMember>>>>;

/**
 * `body.getOoxml()` is Office JS's whole-document flat OPC package, which is exactly
 * `OpcPackage.saveFlatOpc()`; a body with no package (a fragment) gives its part's XML.
 */
function bodyOoxml(body: Body): Promise<string> {
  const pkg = body.package_;
  if (pkg) return pkg.saveFlatOpc();
  return body.getXml();
}

export const EXTRA_MEMBERS: ExtraMembers = {
  Body: {
    getOoxml: (body: Body) => bodyOoxml(body),
  },
  Paragraph: {
    // Word wraps a paragraph in a pkg:package with the styles it needs; that is CR-002 phase C's
    // insertOoxml work, so this is the bare `w:p` fragment (which `insertXml` takes back).
    getOoxml: (paragraph: Paragraph) => paragraph.getXml(),
  },
  Range: {
    getOoxml: (range: Range) => xmlOfRange(range),
    getXml: (range: Range) => xmlOfRange(range),
  },
};
