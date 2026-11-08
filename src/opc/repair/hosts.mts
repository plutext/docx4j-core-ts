// The content hosts an unknown w: element is read through in (CR-026 section 3.7), apart from the pass, so that CR-025 section 10's
// classes can name them without importing the pass, which imports the report.
import { WML_ELEMENT_NAMES } from './wmlNames.generated.mjs';

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
export const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** The twenty hosts check 76 measured. */
export const HOSTS: ReadonlySet<string> = new Set(['p', 'body', 'tc', 'r', 'tbl', 'tr', 'sdtContent', 'hyperlink', 'ins', 'del', 'moveFrom', 'moveTo',
  'hdr', 'ftr', 'comment', 'fldSimple', 'smartTag', 'customXml', 'footnote', 'endnote']);

/** The same names under parents that make them something else, none measured: no host there (WD REQ-308's NOTE). */
export const NOT_A_HOST_UNDER: Readonly<Record<string, readonly (readonly [string, string])[]>> = {
  ins: [[W, 'numPr'], [W, 'rPr'], [W, 'trPr'], [M, 'r'], [M, 'ctrlPr']],
  del: [[W, 'rPr'], [W, 'trPr'], [M, 'r'], [M, 'ctrlPr']],
  moveFrom: [[W, 'rPr']],
  moveTo: [[W, 'rPr']],
  footnote: [[W, 'footnotePr']],
  endnote: [[W, 'endnotePr']],
};

export const isW = (e: Element, local: string): boolean => e.namespaceURI === W && e.localName === local;
export const isUnknown = (e: Element): boolean => e.namespaceURI === W && !WML_ELEMENT_NAMES.has(e.localName);
export const isBranch = (e: Element): boolean => e.namespaceURI === MC && (e.localName === 'Choice' || e.localName === 'Fallback' || e.localName === 'AlternateContent');

/** The content host `parent` is, through a kept branch (`76t44`, `76t45`); undefined for anything else, another namespace's element above all (`76k`). */
export function hostOf(parent: Element): Element | undefined {
  let host: Element = parent;
  while (isBranch(host)) host = host.parentNode as Element;
  if (host.namespaceURI !== W || !HOSTS.has(host.localName)) return undefined;
  const above = host.parentNode as Element | null;
  const excluded = NOT_A_HOST_UNDER[host.localName]?.some(([ns, local]) => above?.namespaceURI === ns && above.localName === local);
  return excluded ? undefined : host;
}
