// Markup compatibility preprocessing (ECMA-376 Part 3; docx4j mc-preprocessor.xslt). Runs over a
// part's DOM before it is unmarshalled: each mc:AlternateContent is replaced by the content of
// the first mc:Choice whose Requires prefixes all name namespaces the object model understands,
// else by its mc:Fallback content, else dropped. This is what Word does on open. It runs only for
// parts that are unmarshalled, so untouched parts keep every branch.
import type { XmlPart } from '../../parts/XmlPart.mjs';
import { UNDERSTOOD_NAMESPACES } from './understood.mjs';

export const MCE_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

export interface McePreprocessOptions {
  /** Namespaces counted as understood; defaults to the object model's modules. */
  understood?: ReadonlySet<string>;
}

/** The default preprocessor, over the object model's namespaces. */
export function mcePreprocess(doc: Document, _part?: XmlPart<unknown>): void {
  resolveAlternateContent(doc, UNDERSTOOD_NAMESPACES);
}

/** A preprocessor over a namespace set of your own. */
export function createMcePreprocessor(options: McePreprocessOptions): (doc: Document, part?: XmlPart<unknown>) => void {
  const understood = options.understood ?? UNDERSTOOD_NAMESPACES;
  return (doc) => resolveAlternateContent(doc, understood);
}

/**
 * Resolves every mc:AlternateContent under a node, innermost first (a chosen branch may contain
 * another). Returns the number resolved.
 */
export function resolveAlternateContent(root: Node, understood: ReadonlySet<string> = UNDERSTOOD_NAMESPACES): number {
  let count = 0;
  const visit = (node: Node): void => {
    // Snapshot: replacing a child mutates the list.
    const children: Node[] = [];
    for (let c = node.firstChild; c; c = c.nextSibling) children.push(c);
    for (const child of children) {
      if (child.nodeType !== 1) continue;
      visit(child);
      const el = child as Element;
      if (el.namespaceURI === MCE_NS && el.localName === 'AlternateContent') {
        replaceAlternateContent(el, understood);
        count++;
      }
    }
  };
  visit(root);
  return count;
}

function replaceAlternateContent(ac: Element, understood: ReadonlySet<string>): void {
  let chosen: Element | undefined;
  let fallback: Element | undefined;
  for (let c = ac.firstChild; c; c = c.nextSibling) {
    if (c.nodeType !== 1) continue;
    const el = c as Element;
    if (el.namespaceURI !== MCE_NS) continue;
    if (el.localName === 'Choice' && !chosen && choiceUnderstood(el, understood)) chosen = el;
    else if (el.localName === 'Fallback' && !fallback) fallback = el;
  }
  const branch = chosen ?? fallback;
  const parent = ac.parentNode!;
  if (branch) {
    const moved: Node[] = [];
    for (let c = branch.firstChild; c; c = c.nextSibling) moved.push(c);
    for (const n of moved) parent.insertBefore(n, ac);
  }
  parent.removeChild(ac);
}

function choiceUnderstood(choice: Element, understood: ReadonlySet<string>): boolean {
  const requires = (choice.getAttribute('Requires') ?? '').trim();
  if (requires === '') return false;
  for (const prefix of requires.split(/\s+/)) {
    const ns = choice.lookupNamespaceURI(prefix);
    if (!ns || !understood.has(ns)) return false;
  }
  return true;
}
