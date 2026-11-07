// An element Word doesn't know, read as Word 365 reads it (CR-026 section 3.7, D7 to D12; WD REQ-308 to REQ-316; check 76). The pass
// runs on the DOM ahead of CR-021's and of the runtime, which drops what it isn't given a place for. An unknown w: element in a
// content host loses its tags and keeps what it holds. A run in one in a paragraph's w:pPr goes to the head of the paragraph. A w:p
// in one, last in a w:p with a run before it, ends that paragraph. Everything else is left for the runtime to skip and report.
import { childElements } from '../../xml/dom.mjs';
import { skippedElement, skippedText, type ContentSkippedWarning } from '../skipped.mjs';
import { UNDERSTOOD_NAMESPACES } from '../mce/understood.mjs';
import { WML_ELEMENT_NAMES } from './wmlNames.generated.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** The twenty hosts check 76 measured. */
const HOSTS: ReadonlySet<string> = new Set(['p', 'body', 'tc', 'r', 'tbl', 'tr', 'sdtContent', 'hyperlink', 'ins', 'del', 'moveFrom', 'moveTo',
  'hdr', 'ftr', 'comment', 'fldSimple', 'smartTag', 'customXml', 'footnote', 'endnote']);

/** The same names under parents that make them something else, none measured: no host there (WD REQ-308's NOTE). */
const NOT_A_HOST_UNDER: Readonly<Record<string, readonly (readonly [string, string])[]>> = {
  ins: [[W, 'numPr'], [W, 'rPr'], [W, 'trPr'], [M, 'r'], [M, 'ctrlPr']],
  del: [[W, 'rPr'], [W, 'trPr'], [M, 'r'], [M, 'ctrlPr']],
  moveFrom: [[W, 'rPr']],
  moveTo: [[W, 'rPr']],
  footnote: [[W, 'footnotePr']],
  endnote: [[W, 'endnotePr']],
};

const isW = (e: Element, local: string): boolean => e.namespaceURI === W && e.localName === local;
const isUnknown = (e: Element): boolean => e.namespaceURI === W && !WML_ELEMENT_NAMES.has(e.localName);
const isBranch = (e: Element): boolean => e.namespaceURI === MC && (e.localName === 'Choice' || e.localName === 'Fallback' || e.localName === 'AlternateContent');

/** The content host `parent` is, through a kept branch (`76t44`, `76t45`); undefined for anything else, another namespace's element above all (`76k`). */
function hostOf(parent: Element): Element | undefined {
  let host: Element = parent;
  while (isBranch(host)) host = host.parentNode as Element;
  if (host.namespaceURI !== W || !HOSTS.has(host.localName)) return undefined;
  const above = host.parentNode as Element | null;
  const excluded = NOT_A_HOST_UNDER[host.localName]?.some(([ns, local]) => above?.namespaceURI === ns && above.localName === local);
  return excluded ? undefined : host;
}

function outermostUnknown(root: Element): Element[] {
  const found: Element[] = [];
  const visit = (e: Element): void => {
    for (const c of childElements(e)) {
      if (isUnknown(c)) found.push(c);
      else visit(c);
    }
  };
  visit(root);
  return found;
}

/** What `e` holds with any unknown element in it taken away, in order. */
function heldElements(e: Element): Element[] {
  return childElements(e).flatMap((c) => (isUnknown(c) ? heldElements(c) : [c]));
}

function reportAll(e: Element, skipped: ContentSkippedWarning[], partName: string | undefined): void {
  skipped.push(skippedElement(e, partName));
  for (const c of childElements(e)) if (isUnknown(c)) reportAll(c, skipped, partName);
}

function reportText(e: Element, skipped: ContentSkippedWarning[], partName: string | undefined): void {
  for (let c = e.firstChild; c !== null; c = c.nextSibling) {
    if ((c.nodeType === 3 || c.nodeType === 4) && /[^ \t\r\n]/.test((c as CharacterData).data)) skipped.push(skippedText((c as CharacterData).data, c, partName));
    if (c.nodeType === 1 && isUnknown(c as Element)) reportText(c as Element, skipped, partName);
  }
}

/** `e`'s tags gone and what it holds in its place, any text with them: text in no run goes (REQ-314). */
function unwrap(e: Element): void {
  const parent = e.parentNode as Node;
  for (const held of heldElements(e)) parent.insertBefore(held, e);
  parent.removeChild(e);
}

const insideAParagraph = (host: Element): boolean => {
  for (let n: Node | null = host; n !== null && n.nodeType === 1; n = n.parentNode) if (isW(n as Element, 'p')) return true;
  return false;
};

/**
 * REQ-316, D11: a w:p in the element, last in its w:p, with a run before it, ends that paragraph, the held w:p's mark taking the
 * host's place. The host takes the held w:p's properties if it has any, a further held w:p stands alone, and what's left of the host
 * is an empty paragraph with the host's properties (`76t43`, `76t56`, `76t58`, `76t68`, `76t69`). Word refuses every other place.
 */
function endHost(e: Element, host: Element, held: Element[]): void {
  const doc = host.ownerDocument as Document;
  const pPrOf = (p: Element): Element | undefined => childElements(p).find((c) => isW(c, 'pPr'));
  const [first, ...further] = held as [Element, ...Element[]];
  const hostPPr = pPrOf(host);
  const firstPPr = pPrOf(first);
  host.removeChild(e);
  for (let c = first.firstChild; c !== null;) {
    const next = c.nextSibling;
    if (c !== firstPPr) host.appendChild(c);
    c = next;
  }
  if (firstPPr) {
    if (hostPPr) host.removeChild(hostPPr);
    host.insertBefore(firstPPr, host.firstChild);
  }
  const parent = host.parentNode as Node;
  const after = host.nextSibling;
  for (const p of further) parent.insertBefore(p, after);
  const rest = doc.createElementNS(W, host.prefix ? `${host.prefix}:p` : 'p');
  if (hostPPr) rest.appendChild(firstPPr ? hostPPr : hostPPr.cloneNode(true));
  parent.insertBefore(rest, after);
}

/** The paragraph a w:pPr, or a w:sectPr in one, belongs to, through a w:pPrChange (`76t40`, `76t50`). */
function paragraphOfProperties(parent: Element): { paragraph: Element; sectPr: Element | undefined } | undefined {
  let n: Element | null = parent;
  const sectPr = isW(n, 'sectPr') ? n : undefined;
  if (sectPr) n = sectPr.parentNode as Element | null;
  if (!n || !isW(n, 'pPr')) return undefined;
  n = n.parentNode as Element | null;
  if (n && isW(n, 'pPrChange')) {
    n = (n.parentNode as Element | null)?.parentNode as Element | null;
  }
  return n && isW(n, 'p') ? { paragraph: n, sectPr } : undefined;
}

/** REQ-311, D8 to D10: runs in the element go to the head of the paragraph, in order; a property goes with the element. */
function runsToHead(e: Element, at: { paragraph: Element; sectPr: Element | undefined }, heads: Map<Element, Node | null>, skipped: ContentSkippedWarning[], partName: string | undefined): void {
  const runs = childElements(e).filter((c) => isW(c, 'r'));
  if (runs.length === 0) return;
  const { paragraph, sectPr } = at;
  const stays = childElements(e).some((c) => !isW(c, 'r'));
  if (!stays) skipped.push(skippedElement(e, partName));
  if (!heads.has(paragraph)) {
    const pPr = childElements(paragraph).find((c) => isW(c, 'pPr'));
    heads.set(paragraph, pPr ? pPr.nextSibling : paragraph.firstChild);
  }
  const head = heads.get(paragraph) ?? null;
  for (const r of runs) paragraph.insertBefore(r, head);
  if (!stays) (e.parentNode as Node).removeChild(e);
  if (sectPr) {
    skipped.push(skippedElement(sectPr, partName));
    (sectPr.parentNode as Node).removeChild(sectPr);
  }
}

/** Reads a part's unknown w: elements as Word does, reporting each it takes. What it leaves, the runtime skips and reports. */
export function readThroughUnknownElements(doc: Document, partName?: string): ContentSkippedWarning[] {
  const root = doc.documentElement;
  if (!root || root.namespaceURI !== W) return [];
  const found = outermostUnknown(root);
  if (found.length === 0) return [];
  const skipped: ContentSkippedWarning[] = [];
  const heads = new Map<Element, Node | null>();
  for (const e of found) {
    const parent = e.parentNode as Element;
    const host = hostOf(parent);
    if (!host) {
      const at = paragraphOfProperties(parent);
      if (at) runsToHead(e, at, heads, skipped, partName);
      continue;
    }
    const held = heldElements(e);
    if (held.some((c) => isW(c, 'p')) && insideAParagraph(host)) {
      const inHost = childElements(host);
      const last = inHost[inHost.length - 1] === e;
      const runBefore = inHost.slice(0, -1).some((c) => isW(c, 'r'));
      if (!isW(host, 'p') || parent !== host || !last || !runBefore || !held.every((c) => isW(c, 'p'))) continue; // Word refuses, or nothing measures it
      reportAll(e, skipped, partName);
      endHost(e, host, held);
      continue;
    }
    reportAll(e, skipped, partName);
    reportText(e, skipped, partName);
    unwrap(e);
  }
  return skipped;
}

/**
 * The elements a caller's fragment holds that Word doesn't know, for `insertOoxml` to refuse (D12): an unknown w: element, and
 * under a w: element one of a namespace the reader doesn't understand. What a w:t holds counts: CR-021's pass would remove it unreported.
 */
export function unknownElementsIn(doc: Document): Element[] {
  const out: Element[] = [];
  const visit = (e: Element): void => {
    for (const c of childElements(e)) {
      const foreign = e.namespaceURI === W && c.namespaceURI !== null && c.namespaceURI !== W && c.namespaceURI !== MC && !UNDERSTOOD_NAMESPACES.has(c.namespaceURI);
      if (isUnknown(c) || foreign) out.push(c);
      else visit(c);
    }
  };
  if (doc.documentElement) visit(doc.documentElement);
  return out;
}
