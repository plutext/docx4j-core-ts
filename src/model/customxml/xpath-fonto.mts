// CR-005 phase A: an XPath 2.0 (in fact 3.1) `XPathEngine`, for the OpenDoPE specification's
// `xpath2` boolean conversion mode (its section 7.2, table 7).
//
// `DefaultXPathEngine` is XPath 1.0, which is what `document.evaluate` and the `xpath` package
// give. The specification's `xpath2` mode needs a 2.0 evaluator, and expressions written for it
// may use 2.0 syntax; docx4j uses Saxon. [FontoXPath](https://github.com/FontoXML/fontoxpath) is
// XPath 3.1 in JavaScript (MIT), over DOM nodes in browsers and in Node, and XPath 3.1 is a
// superset of 2.0.
//
// It is an **optional peer dependency** and this module is its own entry point
// (`@docx4j/core-ts/xpath-fonto`), so nothing that imports the package pays its ~665 KB: only a
// consumer who needs the `xpath2` mode installs it and imports this.
//
//     import { FontoXPathEngine } from '@docx4j/core-ts/xpath-fonto';
//     pkg.xpathEngine = new FontoXPathEngine();
//     await pkg.customXmlParts.load();
import { Docx4JException } from '../../opc/exceptions.mjs';
import type { BooleanConversionMode, XPathEngine, XPathValue } from './xpath.mjs';

/** What this engine uses of `fontoxpath` (typed here: the package is an optional peer). */
interface FontoXPathModule {
  evaluateXPath(expression: string, node: Node, domFacade?: unknown, variables?: unknown, returnType?: number, options?: FontoOptions): unknown;
  evaluateXPathToNodes(expression: string, node: Node, domFacade?: unknown, variables?: unknown, options?: FontoOptions): Node[];
  evaluateXPathToBoolean(expression: string, node: Node, domFacade?: unknown, variables?: unknown, options?: FontoOptions): boolean;
  evaluateXPathToString(expression: string, node: Node, domFacade?: unknown, variables?: unknown, options?: FontoOptions): string;
}

interface FontoOptions {
  namespaceResolver?: (prefix: string) => string | null;
  language?: string;
}

/** `evaluateXPath`'s ANY_TYPE: the value in its natural type, which is what `selectValue` reports. */
const ANY_TYPE = 0;

/**
 * XPath 3.1 through FontoXPath, for the `xpath2` boolean conversion mode and for expressions
 * written in XPath 2.0 syntax. `select`, `selectValue` and the `java` and `xpath1` modes answer as
 * `DefaultXPathEngine` does; `xpath2` is the mode this engine exists for.
 */
export class FontoXPathEngine implements XPathEngine {
  private module: FontoXPathModule | undefined;
  private loading: Promise<void> | undefined;

  get isReady(): boolean {
    return this.module !== undefined;
  }

  async ready(): Promise<void> {
    if (this.isReady) return;
    this.loading ??= (async () => {
      const specifier = 'fontoxpath';
      try {
        // The published ESM build has no named exports under Node's interop, only a default.
        const loaded = await import(/* @vite-ignore */ specifier) as FontoXPathModule & { default?: FontoXPathModule };
        this.module = loaded.default ?? loaded;
      } catch (e) {
        throw new Docx4JException(
          "FontoXPathEngine needs the optional peer dependency 'fontoxpath' (npm install fontoxpath)",
          { cause: e },
        );
      }
    })();
    await this.loading;
  }

  select(expression: string, context: Node, namespaces: Record<string, string> = {}): Node[] {
    return this.fonto().evaluateXPathToNodes(expression, context, null, null, this.options(namespaces));
  }

  selectValue(expression: string, context: Node, namespaces: Record<string, string> = {}): XPathValue {
    const value = this.fonto().evaluateXPath(expression, context, null, null, ANY_TYPE, this.options(namespaces));
    return valueOf(value);
  }

  /**
   * Table 7's three modes. `java` and `xpath1` are the rules the specification states, evaluated
   * here rather than through the `booleanValue` function so that one engine answers all three.
   * `xpath2` is the effective boolean value **with the specification's string rule**: see
   * {@link xpath2BooleanValue}.
   */
  booleanValue(expression: string, context: Node, namespaces: Record<string, string>, mode: BooleanConversionMode): boolean {
    switch (mode) {
      case 'java': {
        const value = this.selectValue(expression, context, namespaces);
        return typeof value === 'boolean' ? value : String(value ?? '').toLowerCase() === 'true';
      }
      case 'xpath1':
        // XPath 1.0's boolean() rules, asked for in that language so that 3.1's do not apply.
        return this.fonto().evaluateXPathToBoolean(`boolean(${expression})`, context, null, null,
          { ...this.options(namespaces), language: 'XPath3.1' });
      default:
        return this.xpath2BooleanValue(expression, context, namespaces);
    }
  }

  /**
   * The `xpath2` mode of table 7, **as docx4j does it**: the expression is wrapped in
   * `xs:boolean(...)` and the engine performs the cast, which is `XmlPart.cachedXPathGetBoolean`
   * with `opendope.conditions.Xpathref.XPathBoolean=cast2` (measured against Saxon-HE 9.9.0-2 by
   * the docx4j session, 2026-09-25, and matched here case for case). Doing what docx4j does,
   * rather than casting by hand, is what keeps the two from drifting: the lexical space, the
   * whitespace collapse, the multi-item rule and the empty sequence all come from the engine.
   *
   * So: `true`, `1` are true; `false`, `0` are false; leading and trailing space is collapsed
   * first, so `" false "` is false; an expression selecting **nothing** is false; and anything
   * else raises - `"yes"`, `"True"` (`xs:boolean`'s lexical space is case-sensitive where the
   * `java` mode ignores case), `""` from an empty element, and a sequence of more than one node.
   *
   * A raise is not a false. docx4j turns it into an `InputIntegrityException` and the whole bind
   * fails; here it is a `Docx4JException` from this call, and what a processor does with it is
   * the processor's business (CR-005 phase B). That is the one place the two behaviours are not
   * yet the same thing, and the specification does not say which is right.
   */
  private xpath2BooleanValue(expression: string, context: Node, namespaces: Record<string, string>): boolean {
    try {
      return this.fonto().evaluateXPathToBoolean(`xs:boolean(${expression})`, context, null, null, this.options(namespaces));
    } catch (e) {
      throw new Docx4JException(
        `The xpath2 boolean conversion mode casts to xs:boolean, which accepts only "true", "false", "1" and "0", `
        + `and one value at a time: ${expression} gives none of them`,
        { cause: e },
      );
    }
  }

  private options(namespaces: Record<string, string>): FontoOptions {
    return { namespaceResolver: (prefix: string) => namespaces[prefix] ?? null };
  }

  private fonto(): FontoXPathModule {
    if (!this.module) {
      throw new Docx4JException(
        'The XPath engine is not ready yet: await pkg.customXmlParts.load() (or pkg.xpathEngine.ready()) once before selecting nodes',
      );
    }
    return this.module;
  }
}

/** A FontoXPath result in the four-type union `selectValue` reports. */
function valueOf(value: unknown): XPathValue {
  if (value === null || value === undefined) return undefined;
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  if (Array.isArray(value)) return value.length === 0 ? undefined : valueOf(value[0]);
  if (typeof value === 'object' && typeof (value as { nodeType?: unknown }).nodeType === 'number') {
    return (value as Node).textContent ?? '';
  }
  return String(value);
}
