// CR-006: XPath over the WordprocessingML tree, through a read-only binder recorded while marshalling.
//
// docx4j answers `getJAXBNodesViaXPath` with a JAXB `Binder<Node>`, which keeps the link between a
// DOM node and the object it came from in both directions. This runtime has no binder, so CR-002
// section 3.6 left XPath over the typed tree out; the demand appeared (the editor's Developer
// perspective selects by XPath) and this is the cheap half - DOM node to object, which is what a
// selection needs. Nothing is written back through the DOM (CR-006 section 7).
//
// The link is recorded by `@docx4j/jsonix` 3.4.0's `onElement` callback, reached through the objects
// facade's `marshalNode(value, { onElement })` (objects CR-007): the runtime calls it once per
// class-typed value with the element it was written into, and once per wildcard DOM child with the
// **imported copy** of that child - `writeNode` does `importNode`, so the element in the snapshot is
// not the node in the tree and a map keyed on it would miss exactly the content a caller most wants
// (jsonix-CR-006 section 2, finding 3).
import { marshalNode, type Jsonix } from '@docx4j/generated-objects-ts';

/** What an XPath selected, and the tree object behind it (CR-006 section 2). */
export interface Bound {
  /**
   * The DOM node the expression selected. It is in a snapshot marshalled for this call, so changing
   * it changes nothing: edits go through the content API or `setContents`, never here.
   */
  node: Node;
  /**
   * The tree object: for an element, what a walk of the typed tree would hand you at that path,
   * never a `{ name, value }` wrapper; for an attribute or a text node, the object of the element
   * that owns it. Undefined for an element the model does not build - a wildcard child is its own
   * object, so this is undefined only above the root or for a node the marshal did not report.
   */
  object: unknown;
  /** Set when an attribute was selected: its qualified name on `object`'s element. */
  attribute?: string;
  /** The node's path from the part's root, `/w:document/w:body/w:p[3]/w:r[2]`. */
  path: string;
}

/** A marshalled snapshot and the element-to-object map recorded while it was written. */
export interface BoundSnapshot {
  root: Element;
  /** The object each element was marshalled from; a wildcard child maps to the original tree node. */
  objects: WeakMap<Element, unknown>;
}

/**
 * Marshals a root element and records what each DOM element came from. One marshal per call: the
 * tree has no change log a cached snapshot could be checked against, and a stale answer is worse
 * than a slow one (docx4j's `refreshXmlFirst`, always true here). The editor measured 349 to 411 ms
 * for a 255-page main document part, against 975 ms for its first unmarshal, and runs a selection
 * on Enter - so no cache is needed, which is what CR-006 section 2 was uncertain about.
 */
export async function bindSnapshot<T>(element: Jsonix.TypedNamedValue<T>): Promise<BoundSnapshot> {
  const objects = new WeakMap<Element, unknown>();
  const root = await marshalNode(element as Jsonix.TypedNamedValue, {
    onElement: (node: unknown, value: unknown) => {
      if (isElement(node)) objects.set(node, value);
    },
  });
  return { root, objects };
}

/** Turns selected nodes into `Bound`s against a snapshot, with one path cache for them all. */
export function boundsOf(nodes: readonly Node[], snapshot: BoundSnapshot): Bound[] {
  const paths = new Paths();
  return nodes.map((node) => {
    const owner = owningElement(node);
    const bound: Bound = {
      node,
      object: owner === undefined ? undefined : snapshot.objects.get(owner),
      path: paths.of(node),
    };
    if (node.nodeType === 2) bound.attribute = (node as Attr).name;
    return bound;
  });
}

/**
 * The element an XPath result belongs to: itself for an element, its owner for an attribute, its
 * parent for a text, comment or processing-instruction node. A document node has none.
 */
function owningElement(node: Node): Element | undefined {
  if (node.nodeType === 1) return node as Element;
  if (node.nodeType === 2) {
    const owner = (node as Attr).ownerElement;
    return owner === null ? undefined : owner;
  }
  const parent = node.parentNode;
  return parent !== null && parent.nodeType === 1 ? (parent as Element) : undefined;
}

/**
 * The node's path from the document element, in the prefixed form the part itself uses, with a
 * position on every step: `/w:document/w:body/w:p[3]/w:r[2]`, and `/@w:val` for an attribute. The
 * position counts siblings of the same qualified name, as XPath does, and is written even when
 * there is only one - a path that is stable whether or not a sibling is added later is worth more
 * than a short one, and docx4j's harness records the same form.
 */
export function pathOf(node: Node): string {
  return new Paths().of(node);
}

/**
 * Paths for the nodes of one snapshot. Each parent's children are counted once, on first need,
 * and each element's path is kept for its descendants, so a hit list costs one pass over the
 * parents it touches. Counting a node's preceding siblings afresh for every node and every
 * ancestor was quadratic in a parent's children: the 6,972 hits of `//w:p[w:pPr/w:pStyle]` over a
 * 7,088-paragraph body took 21 s in Chromium, where the marshal took 0.5 s (CR-006 section 10).
 */
class Paths {
  /** For a parent, the position of each element child among its siblings of the same name. */
  private readonly positions = new Map<Node, Map<Node, number>>();
  private readonly elementPaths = new Map<Node, string>();

  of(node: Node): string {
    if (node.nodeType === 2) {
      const owner = (node as Attr).ownerElement;
      return `${owner === null ? '' : this.of(owner)}/@${(node as Attr).name}`;
    }
    if (node.nodeType === 3 || node.nodeType === 4) {
      const parent = node.parentNode;
      return `${parent === null ? '' : this.of(parent)}/text()`;
    }
    if (node.nodeType !== 1) return '';
    const known = this.elementPaths.get(node);
    if (known !== undefined) return known;
    const element = node as Element;
    const parent = element.parentNode;
    const path = parent === null || parent.nodeType !== 1
      ? `/${element.nodeName}`
      : `${this.of(parent)}/${element.nodeName}[${this.position(parent, element)}]`;
    this.elementPaths.set(node, path);
    return path;
  }

  private position(parent: Node, element: Element): number {
    let positions = this.positions.get(parent);
    if (positions === undefined) {
      positions = new Map();
      const counts = new Map<string, number>();
      for (let sibling = parent.firstChild; sibling; sibling = sibling.nextSibling) {
        if (sibling.nodeType !== 1) continue;
        const name = (sibling as Element).nodeName;
        const n = (counts.get(name) ?? 0) + 1;
        counts.set(name, n);
        positions.set(sibling, n);
      }
      this.positions.set(parent, positions);
    }
    return positions.get(element) ?? 0;
  }
}

function isElement(node: unknown): node is Element {
  return typeof node === 'object' && node !== null && (node as { nodeType?: unknown }).nodeType === 1;
}
