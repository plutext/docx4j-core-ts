// A w:br outside a w:r, read as Word 365 reads it (CR-021, the port of docx4j CR-073; the rules are
// docx4j's mc-preprocessor.xslt at 95c6b42d3, its section 8). Word opens a w:br directly in a w:p, a
// w:hyperlink or another element that holds runs, one inside a w:t or a w:instrText, and one directly
// in w:body, with no repair, and reads each as a break at that point, all the text kept. The object
// model has no place for one there: inside a text element the runtime refuses the whole part ("Element
// text content may not contain START_ELEMENT"), and elsewhere it is a misplaced child (CR-016), unseen
// by the readers and refused at save. docx4j's load survey found one in 206,351 of 3.7 million
// documents, 89% of them a web-to-docx generator's blank line, <w:p><w:br/></w:p>.
//
// The pass runs over a WordprocessingML part's DOM before anything else does (`XmlPart`, ahead of the
// package's hook, and `insertOoxml`'s fragment route; CR-021 D1):
// 1. a w:br in a run container gets a w:r of its own, with the paragraph mark's run properties while no
//    run precedes it in its paragraph, none after one;
// 2. a w:t, w:instrText, w:delText or w:delInstrText holding elements becomes its stretches of text and
//    its breaks, in order, in the same run, a w:t nested in it giving its text, any other w: element and
//    one in a namespace Word doesn't know going with all it holds (Word check 76, CR-021 section 11; docx4j gives its text);
// 3. a w:br directly in w:body goes into the next paragraph as its first run, with that mark's
//    properties, else becomes a paragraph of its own.
// Each rule is decided on the tree as read, as the stylesheet's templates are, before anything moves.
// A clean part pays one walk over its DOM, which allocates nothing (CR-021 section 2: 0.4% of a load).
import { childElements } from '../../xml/dom.mjs';
import { UNDERSTOOD_NAMESPACES } from '../mce/understood.mjs';
import { skippedElement, type ContentSkippedWarning } from '../skipped.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

/** The run containers rule 1 takes a w:br from; named, since a math run (`m:r`) holds a w:br legally. */
const RUN_CONTAINERS: ReadonlySet<string> = new Set(['p', 'hyperlink', 'smartTag', 'customXml', 'fldSimple', 'ins', 'del', 'moveFrom', 'moveTo', 'sdtContent', 'dir', 'bdo']);

/**
 * The paragraph-or-block ancestors: rule 1 takes a w:br only where the nearest is a w:p, so a block-level
 * w:sdtContent or w:customXml is not taken for an inline one.
 */
const BLOCKS: ReadonlySet<string> = new Set(['p', 'body', 'tc', 'txbxContent', 'hdr', 'ftr', 'footnote', 'endnote', 'comment', 'docPartBody']);

/** The text elements rule 2 splits. */
const TEXTS: ReadonlySet<string> = new Set(['t', 'instrText', 'delText', 'delInstrText']);

/** What only a paragraph mark's run properties hold, left out of a run's. */
const MARK_ONLY: ReadonlySet<string> = new Set(['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange']);

/**
 * A part holding a w:br outside a w:r, or a text element holding elements, read as Word reads them (CR-021 D2): one per
 * part, as docx4j's one WARN a part. The code and the fields are docx4j-python's (`part_name`, `breaks`,
 * `text_elements` there), settled through the portfolio session on 2026-11-03. The counts are the stylesheet's: every
 * w:br whose parent is a run container or w:body, and every text element holding an element, at any depth.
 */
export interface BrOutsideRunWarning {
  code: 'br.outside_run';
  partName: string | undefined;
  /** The w:brs found directly in a run container or in w:body. */
  breaks: number;
  /** The w:t, w:instrText, w:delText and w:delInstrText found holding elements. */
  textElements: number;
  /** docx4j's WARN, after the part: "w:br outside a w:r: 2 outside a run, 1 text elements holding elements; read as Word reads them". */
  message: string;
}

/**
 * Puts each w:br found outside a w:r in a run, as Word reads it, over a part's DOM (CR-021; docx4j CR-073's three
 * rules, above). A part whose root is not WordprocessingML's is left alone. Returns the part's warning, or nothing
 * where the part holds none.
 */
export function repairBreaksOutsideRuns(doc: Document, partName?: string): (BrOutsideRunWarning | ContentSkippedWarning)[] {
  const root = doc.documentElement;
  if (!root || root.namespaceURI !== W) return [];
  let breaks: Element[] | undefined;
  let texts: Element[] | undefined;
  let textElements = 0;
  for (let n: Node | null = root; n !== null;) {
    if (n.nodeType === 1) {
      const e = n as Element;
      if (e.namespaceURI === W) {
        const local = e.localName;
        if (local === 'br') {
          const parent = e.parentNode as Element;
          if (parent.namespaceURI === W && (RUN_CONTAINERS.has(parent.localName) || parent.localName === 'body')) (breaks ??= []).push(e);
        } else if (TEXTS.has(local) && holdsElement(e)) {
          textElements++;
          if (!insideText(e)) (texts ??= []).push(e);
        }
      }
      if (e.firstChild !== null) {
        n = e.firstChild;
        continue;
      }
    }
    while (n !== root && n.nextSibling === null) n = n.parentNode as Node;
    n = n === root ? null : n.nextSibling;
  }
  if (breaks === undefined && texts === undefined) return [];

  // decided on the tree as read; a break inside a text element is rule 2's
  const inRuns: { br: Element; rPr: Element | undefined }[] = [];
  const inBody: { br: Element; p: Element | undefined; rPr: Element | undefined }[] = [];
  for (const br of breaks ?? []) {
    if (insideText(br)) continue;
    if ((br.parentNode as Element).localName === 'body') {
      const next = nextBlock(br);
      const p = next !== undefined && takesBodyBreaks(next) ? next : undefined;
      inBody.push({ br, p, rPr: p && markRunProperties(p) });
      continue;
    }
    const p = paragraphOf(br);
    if (p !== undefined) inRuns.push({ br, rPr: runPrecedes(br, p) ? undefined : markRunProperties(p) });
  }

  // (3) first, each paragraph's place for its breaks fixed before any goes in: after its w:pPr, in order
  const places = new Map<Element, Node | null>();
  for (const { br, p, rPr } of inBody) {
    const r = create(br, 'r');
    if (rPr) r.appendChild(rPr);
    if (p === undefined) {
      const own = create(br, 'p');
      (br.parentNode as Node).insertBefore(own, br);
      own.appendChild(r);
    } else {
      if (!places.has(p)) {
        const pPr = childElements(p).find((c) => c.namespaceURI === W && c.localName === 'pPr');
        places.set(p, pPr ? pPr.nextSibling : p.firstChild);
      }
      p.insertBefore(r, places.get(p) ?? null);
    }
    r.appendChild(bare(br));
  }
  // (1)
  for (const { br, rPr } of inRuns) {
    const r = create(br, 'r');
    if (rPr) r.appendChild(rPr);
    (br.parentNode as Node).insertBefore(r, br);
    r.appendChild(bare(br));
  }
  // (2)
  const skipped: ContentSkippedWarning[] = [];
  for (const t of texts ?? []) splitText(t, skipped, partName);

  const outside = breaks?.length ?? 0;
  return [{
    code: 'br.outside_run',
    partName,
    breaks: outside,
    textElements,
    message: `${partName ?? 'a part'}: w:br outside a w:r: ${outside} outside a run, ${textElements} text elements holding elements; read as Word reads them`,
  }, ...skipped];
}

function holdsElement(e: Element): boolean {
  for (let c = e.firstChild; c !== null; c = c.nextSibling) if (c.nodeType === 1) return true;
  return false;
}

function insideText(node: Node): boolean {
  for (let a = node.parentNode; a !== null && a.nodeType === 1; a = a.parentNode) {
    if ((a as Element).namespaceURI === W && TEXTS.has((a as Element).localName)) return true;
  }
  return false;
}

/** The break's paragraph, where its nearest paragraph-or-block ancestor is one. */
function paragraphOf(br: Element): Element | undefined {
  for (let a = br.parentNode; a !== null && a.nodeType === 1; a = a.parentNode) {
    const e = a as Element;
    if (e.namespaceURI === W && BLOCKS.has(e.localName)) return e.localName === 'p' ? e : undefined;
  }
  return undefined;
}

/** True where a run precedes the break in its paragraph: on its path from the paragraph, an earlier sibling is or holds one. */
function runPrecedes(br: Element, p: Element): boolean {
  for (let n: Node = br; n !== p; n = n.parentNode as Node) {
    for (let s = n.previousSibling; s !== null; s = s.previousSibling) {
      if (s.nodeType !== 1) continue;
      const e = s as Element;
      if ((e.namespaceURI === W && e.localName === 'r') || e.getElementsByTagNameNS(W, 'r').length > 0) return true;
    }
  }
  return false;
}

/** The paragraph mark's run properties as a run's, a copy: without what only a mark has; undefined where none is left. */
function markRunProperties(p: Element): Element | undefined {
  let rPr: Element | undefined;
  for (const pPr of childElements(p)) {
    if (pPr.namespaceURI !== W || pPr.localName !== 'pPr') continue;
    for (const markPr of childElements(pPr)) {
      if (markPr.namespaceURI !== W || markPr.localName !== 'rPr') continue;
      for (const property of childElements(markPr)) {
        if (property.namespaceURI === W && MARK_ONLY.has(property.localName)) continue;
        (rPr ??= create(p, 'rPr')).appendChild(property.cloneNode(true));
      }
    }
  }
  return rPr;
}

/** The first element after a body-level break that is not one. */
function nextBlock(br: Element): Element | undefined {
  for (let s = br.nextSibling; s !== null; s = s.nextSibling) {
    if (s.nodeType !== 1) continue;
    const e = s as Element;
    if (!(e.namespaceURI === W && e.localName === 'br')) return e;
  }
  return undefined;
}

/**
 * A w:p the body's breaks before it go into. Not one docx4j splits by its nested-paragraph rule (a w:p in a w:hyperlink
 * or a w:r), which this package doesn't port: there, as in docx4j, each break is a paragraph of its own.
 */
function takesBodyBreaks(e: Element): boolean {
  if (e.namespaceURI !== W || e.localName !== 'p') return false;
  for (const c of childElements(e)) {
    if (c.namespaceURI !== W || (c.localName !== 'hyperlink' && c.localName !== 'r')) continue;
    for (const d of childElements(c)) {
      if (d.namespaceURI !== W) continue;
      if (d.localName === 'p') return false;
      if (d.localName === 'r' && childElements(d).some((x) => x.namespaceURI === W && x.localName === 'p')) return false;
    }
  }
  return true;
}

/** The break as the stylesheet writes it: its attributes, no content. */
function bare(br: Element): Element {
  while (br.firstChild !== null) br.removeChild(br.firstChild);
  return br;
}

/** A WordprocessingML element under the prefix the part gives `like`, which is in that namespace. */
function create(like: Element, local: string): Element {
  return (like.ownerDocument as Document).createElementNS(W, like.prefix ? `${like.prefix}:${local}` : local);
}

/**
 * Rule 2: the text element's stretches of text, each under its name, and its w:brs, in order, in its place. A
 * stretch is XPath's text node (adjacent text and CDATA together), its `xml:space` the nearest ancestor's, and kept
 * only where that preserves it or it holds more than whitespace. Any other element gives its text only as `givesItsText` says.
 */
function splitText(t: Element, skipped: ContentSkippedWarning[], partName: string | undefined): void {
  const out: Node[] = [];
  const visit = (holder: Element): void => {
    for (let c = holder.firstChild; c !== null;) {
      if (c.nodeType === 3 || c.nodeType === 4) {
        let text = '';
        for (; c !== null && (c.nodeType === 3 || c.nodeType === 4); c = c.nextSibling) text += (c as CharacterData).data;
        const space = spaceOf(holder);
        if (space === 'preserve' || /[^ \t\r\n]/.test(text)) {
          const piece = create(t, t.localName);
          if (space !== '') piece.setAttributeNS(XML_NS, 'xml:space', space);
          piece.appendChild((t.ownerDocument as Document).createTextNode(text));
          out.push(piece);
        }
        continue;
      }
      if (c.nodeType === 1) {
        const e = c as Element;
        if (e.namespaceURI === W && e.localName === 'br') out.push(bare(e));
        else if (givesItsText(e)) visit(e);
        else skipped.push(skippedElement(e, partName)); // the runtime never sees it (CR-025 D9)
      }
      c = c.nextSibling;
    }
  };
  visit(t);
  const parent = t.parentNode as Node;
  for (const piece of out) parent.insertBefore(piece, t);
  parent.removeChild(t);
}

/**
 * Whether an element in a text element gives its text (CR-021 section 11). As Word 365 reads it, a nested text element does, and
 * any other w: element and one in a namespace Word doesn't know go with all they hold (check 76's `76n`, `76u`, `76w`, `76z`). Word
 * refuses an element in no namespace (`76x`, `76y`), so docx4j's reading stands there; a namespace Word knows is unmeasured.
 */
function givesItsText(e: Element): boolean {
  const namespace = e.namespaceURI;
  if (namespace === W) return TEXTS.has(e.localName);
  return !namespace || UNDERSTOOD_NAMESPACES.has(namespace);
}

function spaceOf(holder: Element): string {
  for (let a: Node | null = holder; a !== null && a.nodeType === 1; a = a.parentNode) {
    if ((a as Element).hasAttributeNS(XML_NS, 'space')) return (a as Element).getAttributeNS(XML_NS, 'space') ?? '';
  }
  return '';
}
