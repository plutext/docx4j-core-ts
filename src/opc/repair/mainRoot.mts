// A main document part whose root isn't w:document, read as WD REQ-318 to REQ-320 have it (CR-028 D10 and D11; docx4j CR-079,
// python CR-039). Word reads some such roots as a document, refuses some, and opens the rest showing nothing. The pass runs on
// the DOM ahead of CR-026's and of the runtime, so a root it reads reaches them as a w:document holding a w:body.
// mc:AlternateContent as the root is REQ-317's, and left as it is.
import { PartUnmarshalException } from '../exceptions.mjs';
import { skippedElement, type ContentSkippedWarning } from '../skipped.mjs';
import { UNDERSTOOD_NAMESPACES } from '../mce/understood.mjs';
import { WML_ELEMENT_NAMES } from './wmlNames.generated.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const WORD_2003 = 'http://schemas.microsoft.com/office/word/2003/wordml';
const XML_PACKAGE = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const XMLNS = 'http://www.w3.org/2000/xmlns/';

/**
 * `part.unexpected_root` (WD REQ-320): a main document part whose root Word opens and shows nothing of, such as `w:hdr`,
 * `w:styles` or an element in a namespace its `mc:Ignorable` names. It reads as an empty document, and a write through `Body` is
 * refused (CR-028 D10). `root` is the root's qualified name, as the file spells it. The root is also a `part.skipped`, so a save
 * copies the part as written while it is unchanged, and CR-025's gate refuses one that changed.
 */
export interface UnexpectedRootWarning {
  code: 'part.unexpected_root';
  partName: string;
  root: string;
  message: string;
}

/**
 * Reads the root of a main document part's DOM: a root Word reads becomes a `w:document` holding a `w:body`, a root Word refuses
 * throws `PartUnmarshalException`, and any other gives an empty document with `part.unexpected_root`. A `w:document` and an
 * `mc:AlternateContent` are left alone. Returns the warnings: what a read root dropped, as `part.skipped`, or the unexpected root.
 */
export function readMainPartRoot(doc: Document, partName: string): (UnexpectedRootWarning | ContentSkippedWarning)[] {
  const root = doc.documentElement;
  if (!root) return [];
  if (isW(root, 'document') || (root.namespaceURI === MC && root.localName === 'AlternateContent')) return [];
  if (refused(root)) {
    throw new PartUnmarshalException(`${partName}: its root, ${root.nodeName}, is one Word refuses a document for (WD REQ-318)`, partName);
  }
  const body = bodyFrom(root, partName);
  if (body) {
    replaceRoot(doc, root, body.body);
    return body.skipped;
  }
  const skipped = skippedElement(root, partName);
  replaceRoot(doc, root, doc.createElementNS(W, 'w:body'));
  return [{
    code: 'part.unexpected_root',
    partName,
    root: root.nodeName,
    message: `${partName}: its root, ${root.nodeName}, is one Word opens showing nothing of, so it reads as an empty document (WD REQ-320)`,
  }, skipped];
}

/** WD REQ-318: a root Word refuses, by its name or by what it holds. */
function refused(root: Element): boolean {
  const ns = root.namespaceURI;
  const local = root.localName;
  if (ns === W && (local === 'footnotes' || local === 'endnotes' || local === 'tc' || local === 'txbxContent')) return true;
  if ((ns === WORD_2003 && local === 'wordDocument') || (ns === XML_PACKAGE && local === 'package')) return true;
  if (isW(root, 'sdtContent')) return childrenOf(root).some((c) => isW(c, 'tr') || (isW(c, 'sdt') && sdtLevel(c) === 'row'));
  if (ns === W && !WML_ELEMENT_NAMES.has(local)) {
    return childrenOf(root).some((c) => isW(c, 'document') || isW(c, 'r') || (isW(c, 'sdt') && sdtLevel(c) === 'run'));
  }
  if (ns !== null && (ns === W || UNDERSTOOD_NAMESPACES.has(ns))) return false;
  return !ignorable(root) && childrenOf(root).some((c) => isW(c, 'document') || isW(c, 'p'));
}

/**
 * WD REQ-319: the `w:body` a root Word reads stands for, made from what it holds, and what that dropped; undefined for a root
 * Word doesn't read (REQ-320).
 */
function bodyFrom(root: Element, partName: string): { body: Element; skipped: ContentSkippedWarning[] } | undefined {
  const doc = root.ownerDocument;
  const wrap = (nodes: Node[]): Element => {
    const body = doc.createElementNS(W, 'w:body');
    for (const n of nodes) body.appendChild(n);
    return body;
  };
  if (isW(root, 'body')) return { body: root, skipped: [] };
  if (isW(root, 'p') || (isW(root, 'sdt') && sdtLevel(root) === 'block')) return { body: wrap([root]), skipped: [] };
  if (isW(root, 'sdtContent')) return { body: wrap(nodesOf(root)), skipped: [] };
  if (root.namespaceURI !== W || WML_ELEMENT_NAMES.has(root.localName)) return undefined;
  const children = childrenOf(root);
  const held = children.find((c) => isW(c, 'body'));
  if (held) {
    // that body is the body, its w:sectPr read (mr-15); what stands beside it nothing measures, so it goes and is said to
    const skipped = children.filter((c) => c !== held).map((c) => skippedElement(c, partName));
    return { body: held, skipped };
  }
  if (children.length > 0 && !children.some(readsAsBody)) return undefined;
  // a w:sectPr directly in it has no effect (mr-11), so it goes, said as CR-026 D10 says a w:sectPr it drops
  const sections = children.filter((c) => isW(c, 'sectPr'));
  const skipped = sections.map((c) => skippedElement(c, partName));
  for (const s of sections) root.removeChild(s);
  return { body: wrap(nodesOf(root)), skipped };
}

/** What an unknown `w:` root holds that makes Word read it as a body (WD REQ-319). */
function readsAsBody(c: Element): boolean {
  if (isW(c, 'p') || isW(c, 'tbl') || (isW(c, 'sdt') && sdtLevel(c) === 'block')) return true;
  if (c.namespaceURI === MC && c.localName === 'AlternateContent') return true;
  return c.namespaceURI === W && !WML_ELEMENT_NAMES.has(c.localName);
}

/**
 * Puts a `w:document` holding `body` where `root` stood. The root's namespace declarations are copied onto it, so a prefix only
 * the root bound still resolves, and its `mc:` attributes moved there, so its `mc:Ignorable` still stands and a root kept as the
 * body or its paragraph holds none the model has no place for. `body` may be the root, or hold it.
 */
function replaceRoot(doc: Document, root: Element, body: Element): void {
  const document = doc.createElementNS(W, 'w:document');
  const moved: Attr[] = [];
  for (let i = 0; i < root.attributes.length; i++) {
    const a = root.attributes.item(i)!;
    if (a.namespaceURI === XMLNS || a.namespaceURI === MC) document.setAttributeNS(a.namespaceURI, a.name, a.value);
    if (a.namespaceURI === MC) moved.push(a);
  }
  for (const a of moved) root.removeAttributeNode(a);
  if (root.parentNode === doc) doc.removeChild(root);
  if (body.parentNode && body.parentNode !== document) body.parentNode.removeChild(body);
  document.appendChild(body);
  doc.appendChild(document);
}

/** The level a `w:sdt` stands at, by what its `w:sdtContent` holds first that says: an empty one is taken as a block. */
function sdtLevel(sdt: Element): 'block' | 'row' | 'cell' | 'run' {
  const content = childrenOf(sdt).find((c) => isW(c, 'sdtContent'));
  for (const c of content ? childrenOf(content) : []) {
    if (isW(c, 'p') || isW(c, 'tbl') || isW(c, 'altChunk')) return 'block';
    if (isW(c, 'tr')) return 'row';
    if (isW(c, 'tc')) return 'cell';
    if (isW(c, 'r') || isW(c, 'hyperlink') || isW(c, 'fldSimple') || isW(c, 'smartTag') || isW(c, 'ins') || isW(c, 'del')) return 'run';
    if (isW(c, 'sdt')) return sdtLevel(c);
  }
  return 'block';
}

/** Whether the root's own namespace is one its `mc:Ignorable` names. */
function ignorable(root: Element): boolean {
  const ns = root.namespaceURI;
  if (!ns) return false;
  const listed = root.getAttributeNS(MC, 'Ignorable') ?? '';
  return listed.trim().split(/\s+/).some((prefix) => prefix !== '' && root.lookupNamespaceURI(prefix) === ns);
}

function isW(el: Element, local: string): boolean {
  return el.namespaceURI === W && el.localName === local;
}

function childrenOf(el: Element): Element[] {
  const out: Element[] = [];
  for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) out.push(c as Element);
  return out;
}

function nodesOf(el: Element): Node[] {
  const out: Node[] = [];
  for (let c = el.firstChild; c; c = c.nextSibling) out.push(c);
  return out;
}
