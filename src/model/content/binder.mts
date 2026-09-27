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
import { Docx4JException } from '../../opc/exceptions.mjs';

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

/** Turns selected nodes into `Bound`s against a snapshot. */
export function boundsOf(nodes: readonly Node[], snapshot: BoundSnapshot): Bound[] {
  return nodes.map((node) => {
    const owner = owningElement(node);
    const bound: Bound = {
      node,
      object: owner === undefined ? undefined : snapshot.objects.get(owner),
      path: pathOf(node),
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
  if (node.nodeType === 2) {
    const owner = (node as Attr).ownerElement;
    return `${owner === null ? '' : pathOf(owner)}/@${(node as Attr).name}`;
  }
  if (node.nodeType === 3 || node.nodeType === 4) {
    const parent = node.parentNode;
    return `${parent === null ? '' : pathOf(parent)}/text()`;
  }
  if (node.nodeType !== 1) return '';
  const element = node as Element;
  const parent = element.parentNode;
  if (parent === null || parent.nodeType !== 1) return `/${element.nodeName}`;
  let index = 0;
  for (let sibling = parent.firstChild; sibling; sibling = sibling.nextSibling) {
    if (sibling.nodeType === 1 && (sibling as Element).nodeName === element.nodeName) {
      index++;
      if (sibling === element) break;
    }
  }
  return `${pathOf(parent)}/${element.nodeName}[${index}]`;
}

function isElement(node: unknown): node is Element {
  return typeof node === 'object' && node !== null && (node as { nodeType?: unknown }).nodeType === 1;
}

/** The engine an expression goes to, and the error when there is none ready. */
export function requireReadyEngine<E extends { isReady: boolean }>(engine: E | undefined, what: string): E {
  if (engine === undefined) throw new Docx4JException(`${what} needs an XPath engine; set pkg.xpathEngine`);
  if (!engine.isReady) {
    throw new Docx4JException(`${what} needs its XPath engine ready; await pkg.xpathEngine.ready(node) first`);
  }
  return engine;
}
