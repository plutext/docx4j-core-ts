// CR-002 phase E, section 3.6: XPath 1.0 over the custom XML DOM.
//
// Two runtimes, one interface. In a browser or an Office add-in `document.evaluate` is native and
// complete; in Node the `xpath` package (goto100, MIT) is the companion of the @xmldom/xmldom DOM
// the runtime brings, and is an **optional peer dependency** so that an add-in bundle never carries
// it (CR-002 open question 7). `pkg.xpathEngine` is settable, so a consumer with another DOM, or
// one who wants no dynamic import at all, plugs in its own.
//
// Readiness: loading the `xpath` package is asynchronous, and section 3's rule is that only what
// marshals or unmarshals is asynchronous. So the engine is **warmed once** — `await
// pkg.xpathEngine.ready()`, which `pkg.customXmlParts.load()` does for you — and every node call
// afterwards is synchronous. A call before the engine is ready throws, naming the fix.
import { Docx4JException } from '../../opc/exceptions.mjs';
import { serializeXml } from '../../xml/dom.mjs';

/** The XPath result kinds `selectValue` can report. */
export type XPathValue = string | number | boolean | undefined;

/**
 * XPath 1.0 over a DOM (CR-002 section 3.6). `select` answers a node set; `selectValue` answers a
 * string, number or boolean result. `namespaces` maps prefixes used in the expression to URIs.
 */
export interface XPathEngine {
  /** Prepares the engine (loads the `xpath` package in Node); idempotent, and cheap once done. */
  ready(): Promise<void>;
  /** True when `select` and `selectValue` can be called synchronously. */
  readonly isReady: boolean;
  select(expression: string, context: Node, namespaces?: Record<string, string>): Node[];
  selectValue(expression: string, context: Node, namespaces?: Record<string, string>): XPathValue;
}

/** What this engine uses of the `xpath` package (typed here: the package is an optional peer). */
interface XPathModule {
  useNamespaces(map: Record<string, string>): (expression: string, node: Node, single?: boolean) => unknown;
  select(expression: string, node: Node, single?: boolean): unknown;
}

// XPathResult constants, spelled out because `XPathResult` is not a global in Node.
const ANY_TYPE = 0;
const NUMBER_TYPE = 1;
const STRING_TYPE = 2;
const BOOLEAN_TYPE = 3;
const ORDERED_NODE_SNAPSHOT_TYPE = 7;

function documentOf(context: Node): Document | undefined {
  return (context.nodeType === 9 ? (context as Document) : context.ownerDocument) ?? undefined;
}

function nativeEvaluatorFor(context: Node): Document | undefined {
  const doc = documentOf(context);
  return doc && typeof (doc as { evaluate?: unknown }).evaluate === 'function' ? doc : undefined;
}

/**
 * The default engine: `document.evaluate` where the runtime has it, else the `xpath` package,
 * imported on `ready()`. The import specifier is a variable so that a bundler does not follow it
 * and a consumer without the optional peer still type-checks.
 */
export class DefaultXPathEngine implements XPathEngine {
  private module: XPathModule | undefined;
  private loading: Promise<void> | undefined;
  /** True when the runtime has a native `document.evaluate` (a browser, a Word add-in). */
  private native = typeof document !== 'undefined' && typeof (document as { evaluate?: unknown }).evaluate === 'function';

  get isReady(): boolean {
    return this.native || this.module !== undefined;
  }

  async ready(): Promise<void> {
    if (this.isReady) return;
    this.loading ??= (async () => {
      const specifier = 'xpath';
      try {
        this.module = (await import(/* @vite-ignore */ specifier)) as XPathModule;
      } catch (e) {
        throw new Docx4JException(
          "XPath needs either a DOM with document.evaluate or the optional peer dependency 'xpath' (npm install xpath); "
          + 'set pkg.xpathEngine to your own engine to avoid both',
          { cause: e },
        );
      }
    })();
    await this.loading;
  }

  select(expression: string, context: Node, namespaces: Record<string, string> = {}): Node[] {
    const doc = nativeEvaluatorFor(context);
    if (doc) {
      const result = doc.evaluate(expression, context, resolverFor(namespaces), ORDERED_NODE_SNAPSHOT_TYPE, null);
      const out: Node[] = [];
      for (let i = 0; i < result.snapshotLength; i++) {
        const node = result.snapshotItem(i);
        if (node) out.push(node);
      }
      return out;
    }
    const value = this.call(expression, context, namespaces);
    return Array.isArray(value) ? (value as Node[]) : isNode(value) ? [value] : [];
  }

  selectValue(expression: string, context: Node, namespaces: Record<string, string> = {}): XPathValue {
    const doc = nativeEvaluatorFor(context);
    if (doc) {
      const result = doc.evaluate(expression, context, resolverFor(namespaces), ANY_TYPE, null);
      switch (result.resultType) {
        case NUMBER_TYPE: return result.numberValue;
        case STRING_TYPE: return result.stringValue;
        case BOOLEAN_TYPE: return result.booleanValue;
        default: {
          const node = result.iterateNext();
          return node ? (node.textContent ?? '') : undefined;
        }
      }
    }
    const value = this.call(expression, context, namespaces);
    if (Array.isArray(value)) {
      const node = value[0] as Node | undefined;
      return node ? (node.textContent ?? '') : undefined;
    }
    if (isNode(value)) return value.textContent ?? '';
    return value === null || value === undefined ? undefined : (value as XPathValue);
  }

  private call(expression: string, context: Node, namespaces: Record<string, string>): unknown {
    if (!this.module) {
      throw new Docx4JException(
        'The XPath engine is not ready yet: await pkg.customXmlParts.load() (or pkg.xpathEngine.ready()) once before selecting nodes',
      );
    }
    const select = Object.keys(namespaces).length > 0 ? this.module.useNamespaces(namespaces) : this.module.select;
    return select(expression, context);
  }
}

function isNode(value: unknown): value is Node {
  return typeof value === 'object' && value !== null && typeof (value as { nodeType?: unknown }).nodeType === 'number';
}

function resolverFor(namespaces: Record<string, string>): XPathNSResolver {
  return { lookupNamespaceURI: (prefix: string | null) => (prefix ? namespaces[prefix] ?? null : null) };
}

/**
 * Word's prefix mapping string, `xmlns:ns0='http://x' xmlns:ns1="http://y"`, as a map. The default
 * declaration (`xmlns='...'`) is read under the empty prefix, though XPath 1.0 never uses it.
 */
export function parsePrefixMappings(mappings: string | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  if (!mappings) return out;
  const re = /xmlns(?::([\w.\-]+))?\s*=\s*(['"])(.*?)\2/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(mappings)) !== null) out[m[1] ?? ''] = m[3]!;
  return out;
}

/** The inverse: the `xmlns:ns0='...'` form Word writes into `w:prefixMappings`. */
export function formatPrefixMappings(namespaces: Record<string, string>): string {
  return Object.entries(namespaces)
    .map(([prefix, uri]) => (prefix === '' ? `xmlns='${uri}'` : `xmlns:${prefix}='${uri}'`))
    .join(' ');
}

/** A node's canonical XPath and the prefixes it uses. */
export interface CanonicalXPath {
  xpath: string;
  prefixMappings: string;
}

/**
 * The canonical `/ns0:a[1]/ns0:b[2]` form Word writes for a node (`setMappingByNode`,
 * `CustomXmlNode.xpath`): a location path from the document element down, every step with its
 * position among the siblings of the same name, prefixes allocated ns0, ns1, ... in the order the
 * namespaces are met. A node in no namespace gets no prefix, which is what Word writes too
 * (`/invoice[1]/customer[1]/name[1]` in the OpenDoPE invoice).
 */
export function canonicalXPathOf(node: Node): CanonicalXPath {
  const chain: Node[] = [];
  let current: Node | null = node;
  while (current && current.nodeType !== 9) {
    chain.unshift(current);
    current = current.nodeType === 2 ? ((current as Attr).ownerElement as Node | null) : current.parentNode;
  }
  const prefixes = new Map<string, string>();
  const qname = (namespaceURI: string | null, localName: string): string => {
    if (!namespaceURI) return localName;
    let prefix = prefixes.get(namespaceURI);
    if (prefix === undefined) {
      prefix = `ns${prefixes.size}`;
      prefixes.set(namespaceURI, prefix);
    }
    return `${prefix}:${localName}`;
  };
  const steps: string[] = [];
  for (const item of chain) {
    if (item.nodeType === 2) {
      const attr = item as Attr;
      steps.push(`@${qname(attr.namespaceURI, attr.localName ?? attr.name)}`);
      continue;
    }
    if (item.nodeType === 3 || item.nodeType === 4) {
      steps.push(`text()[${positionOf(item, (n) => n.nodeType === 3 || n.nodeType === 4)}]`);
      continue;
    }
    if (item.nodeType === 8) {
      steps.push(`comment()[${positionOf(item, (n) => n.nodeType === 8)}]`);
      continue;
    }
    const element = item as Element;
    const local = element.localName ?? element.nodeName;
    const ns = element.namespaceURI;
    const position = positionOf(element, (n) => n.nodeType === 1 && ((n as Element).localName ?? n.nodeName) === local && ((n as Element).namespaceURI ?? null) === (ns ?? null));
    steps.push(`${qname(ns, local)}[${position}]`);
  }
  const namespaces: Record<string, string> = {};
  for (const [uri, prefix] of prefixes) namespaces[prefix] = uri;
  return { xpath: `/${steps.join('/')}`, prefixMappings: formatPrefixMappings(namespaces) };
}

function positionOf(node: Node, matches: (candidate: Node) => boolean): number {
  const parent = node.parentNode;
  if (!parent) return 1;
  let position = 0;
  for (let child = parent.firstChild; child; child = child.nextSibling) {
    if (!matches(child)) continue;
    position++;
    if (child === node) return position;
  }
  return position || 1;
}

/** The XML of a node, without an XML declaration (`CustomXmlNode.xml`). */
export function xmlOf(node: Node): string {
  if (node.nodeType === 2) {
    const attr = node as Attr;
    return `${attr.name}="${attr.value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;')}"`;
  }
  return serializeXml(node);
}
