// What a read drops, as a part's warning (CR-025). `part.skipped` is python's code, agreed through its CR-033 Q9, and
// `kind`, `path` and `line` its `SkippedNode`'s. The runtime reports what it drops through three callbacks, which
// `XmlPart.unmarshal` alone passes: objects CR-007 keeps them per call, since a drop is sometimes the mechanism
// (`recordPPrChange`'s deep copy). CR-021's pass removes what it removes before the runtime sees it, so it reports its own.

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** One element, attribute or stretch of text a read dropped. */
export interface ContentSkippedWarning {
  code: 'part.skipped';
  partName: string | undefined;
  kind: 'element' | 'attribute' | 'text';
  /** As the part writes it, `ext:opaque`; for text, its first 80 characters. */
  name: string;
  /** The element holding it, `w:sdtPr`. */
  parentName: string;
  /** That element's class, `org_docx4j_wml.SdtPr`. Absent for what CR-021's pass removes, which no class has met. */
  parentClass?: string;
  /** From the part's root, `/w:document/w:body/w:sdt/w:sdtPr/ext:opaque`; an attribute's last step is `@ext:a`, text's `text()`. */
  path: string;
  /** Its namespace is listed in `mc:Ignorable` where it stands. Always false for text. */
  ignorable: boolean;
  /** Where the parser gives one, as `target.misplaced`: an attribute's is its element's. */
  line?: number;
  column?: number;
  message: string;
}

type Located = Node & { lineNumber?: number; columnNumber?: number };

const position = (node: Node): { line?: number; column?: number } => {
  const { lineNumber, columnNumber } = node as Located;
  return { ...(typeof lineNumber === 'number' ? { line: lineNumber } : {}), ...(typeof columnNumber === 'number' ? { column: columnNumber } : {}) };
};

function pathOf(element: Node): string {
  const steps: string[] = [];
  for (let n: Node | null = element; n !== null && n.nodeType === 1; n = n.parentNode) steps.unshift(n.nodeName);
  return `/${steps.join('/')}`;
}

/**
 * Whether `mc:Ignorable` lists the namespace, there or on an ancestor, each prefix read where the attribute stands. A prefix
 * that no longer resolves counts by its name: `mcePreprocess` takes a branch out of the `mc:Choice` declaring it.
 */
function isIgnorable(namespace: string | null | undefined, prefix: string | null | undefined, from: Element): boolean {
  if (!namespace) return false;
  for (let n: Node | null = from; n !== null && n.nodeType === 1; n = n.parentNode) {
    const element = n as Element;
    const listed = element.getAttributeNS(MC, 'Ignorable');
    if (!listed) continue;
    for (const candidate of listed.trim().split(/\s+/)) {
      const uri = element.lookupNamespaceURI(candidate);
      if (uri === namespace || (!uri && candidate === prefix)) return true;
    }
  }
  return false;
}

function warning(kind: ContentSkippedWarning['kind'], name: string, parent: Element, path: string, ignorable: boolean, located: Node,
  partName: string | undefined, parentClass: string | undefined): ContentSkippedWarning {
  return {
    code: 'part.skipped', partName, kind, name, parentName: parent.nodeName,
    ...(parentClass !== undefined ? { parentClass } : {}),
    path, ignorable, ...position(located),
    message: `${partName ?? 'a part'}: ${path}: the read dropped this ${kind}, which the model has no place for`,
  };
}

/** An element and all it holds. */
export function skippedElement(element: Element, partName: string | undefined, parentClass?: string): ContentSkippedWarning {
  return warning('element', element.nodeName, element.parentNode as Element, pathOf(element), isIgnorable(element.namespaceURI, element.prefix, element), element, partName, parentClass);
}

/** An attribute of `owner`, named as the runtime reports it. */
export function skippedAttribute(name: { namespaceURI?: string; localPart: string; prefix?: string }, owner: Element, partName: string | undefined, parentClass?: string): ContentSkippedWarning {
  const shown = name.prefix ? `${name.prefix}:${name.localPart}` : name.localPart;
  return warning('attribute', shown, owner, `${pathOf(owner)}/@${shown}`, isIgnorable(name.namespaceURI, name.prefix, owner), owner, partName, parentClass);
}

/** A stretch of text, as a text or CDATA node. */
export function skippedText(text: string, node: Node, partName: string | undefined, parentClass?: string): ContentSkippedWarning {
  const parent = node.parentNode as Element;
  return warning('text', text.length > 80 ? text.slice(0, 80) : text, parent, `${pathOf(parent)}/text()`, false, node, partName, parentClass);
}
