// What a read drops, as a part's warning (CR-025). `part.skipped` is python's code, agreed through its CR-033 Q9, and
// `kind`, `path` and `line` its `SkippedNode`'s. The runtime reports what it drops through three callbacks, which
// `XmlPart.unmarshal` alone passes: objects CR-007 keeps them per call, since a drop is sometimes the mechanism
// (`recordPPrChange`'s deep copy). CR-021's pass removes what it removes before the runtime sees it, so it reports its own.

import { UNDERSTOOD_NAMESPACES } from './mce/understood.mjs';
import { W, hostOf, isUnknown, isW } from './repair/hosts.mjs';

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/**
 * What a save does with an item when the part holding it has changed (CR-025 section 10, D14): `dropped`, written without it
 * as Word 365's own save writes it, listed on `lastSave.dropped` and never refused; or `lost`, which Word's save keeps, or
 * which Word won't open, or which nobody has measured, refused unless `checkLossOnSave` is off (D15).
 */
export type OnSave = 'dropped' | 'lost';

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
  /** Its namespace URI, `null` for none and for text (CR-025 section 10). */
  namespace: string | null;
  /** What a save of the changed part does with it (CR-025 section 10.3): written without it, or refused. */
  onSave: OnSave;
  /** The Word file or statement that decides `onSave` (`'76i'`, `'REQ-314'`); absent for a lost item nobody measured. */
  measured?: string;
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
 * that no longer resolves counts by its name, since `mcePreprocess` takes a branch out of the `mc:Choice` declaring it, but not
 * where the item's own prefix still resolves to it: then it was declared only below the list, which doesn't put it in scope
 * there, and the list names nothing (WD REQ-323, `ui-05`).
 */
function isIgnorable(namespace: string | null | undefined, prefix: string | null | undefined, from: Element): boolean {
  if (!namespace) return false;
  for (let n: Node | null = from; n !== null && n.nodeType === 1; n = n.parentNode) {
    const element = n as Element;
    const listed = element.getAttributeNS(MC, 'Ignorable');
    if (!listed) continue;
    for (const candidate of listed.trim().split(/\s+/)) {
      const uri = element.lookupNamespaceURI(candidate);
      if (uri === namespace || (!uri && candidate === prefix && from.lookupNamespaceURI(prefix) !== namespace)) return true;
    }
  }
  return false;
}

export interface SaveClass { readonly onSave: OnSave; readonly measured?: string }
/** Written without, as Word's own save writes it; `measured` the Word file or statement. */
export const dropped = (measured: string): SaveClass => ({ onSave: 'dropped', measured });
const lost = (measured?: string): SaveClass => (measured === undefined ? { onSave: 'lost' } : { onSave: 'lost', measured });
const isElement = (n: Node | null): n is Element => n !== null && n.nodeType === 1;

/** Check 76's file for an element in a namespace Word doesn't understand, by its host: ignorable (`76a` ...) and not (`76e` ...). */
const FOREIGN_IGNORABLE: Readonly<Record<string, string>> = { sdtPr: '76a', p: '76b', r: '76c', tc: '76d', b: '76l', t: '76m' };
const FOREIGN_REFUSED: Readonly<Record<string, string>> = { p: '76e', sdtPr: '76o', pPr: '76p', r: '76q', tc: '76r', b: '76s', t: '76v' };
/** The property hosts where a run held by an unknown `w:` element makes Word refuse the file (CR-026 section 3.7). */
const RUN_REFUSED: Readonly<Record<string, string>> = { rPr: '76t28', sdtPr: '76t46', trPr: '76t47', tcPr: '76t48', tblPr: '76t49' };

/**
 * An element's class (CR-025 section 10.3's table, D15): by the Word file that measured its shape, a shape nobody measured lost.
 * `el` is the element as the read met it, in its parent.
 */
function elementClass(el: Element, ignorable: boolean): SaveClass {
  const parent = el.parentNode;
  if (!isElement(parent)) return lost('REQ-320'); // a part's root: what an unexpected root held, which has no reading to copy
  const ns = el.namespaceURI;
  if (!ns) return lost('76x'); // in no namespace: Word won't open the file (76x, 76y)
  const host = parent.namespaceURI === W ? parent.localName : undefined;
  if (ns === W) {
    if (isUnknown(el)) {
      // an unknown w: element. CR-026's pass reads one through where Word does and gives its items that class itself; one it
      // leaves in a content host is a shape Word refuses or nobody measured (76t53, 76t54, 76t57, 76t59), left to the runtime
      if (isUnknown(parent) || hostOf(parent)) return lost();
      if (host === 't') return dropped('76u'); // CR-021's pass takes it with what it holds, as Word's save does
      if (host === 'pPr') return dropped('76g');
      if (host !== undefined && RUN_REFUSED[host] !== undefined) {
        if (el.getElementsByTagNameNS(W, 'r').length > 0) return lost(RUN_REFUSED[host]);
        return host === 'rPr' ? dropped('76t6') : lost();
      }
      if (host === 'sectPr' && isElement(parent.parentNode) && isW(parent.parentNode, 'pPr')) return dropped('76t50');
      return lost();
    }
    // a known w: element out of place
    if (el.localName === 'document') return lost('nd-10'); // a w:document in a w:document: Word won't open the file
    if (el.localName === 'sectPr' && isUnknown(parent)) return dropped('mr-11');
    if (el.localName === 'sectPr' && host === 'pPr') return dropped('76t50');
    if (el.localName === 'tab' && host === 't') return dropped('76z');
    return lost();
  }
  if (UNDERSTOOD_NAMESPACES.has(ns)) return lost(); // Word understands it and its save keeps it (python's 1b)
  // a namespace Word doesn't understand: ignorable, Word ignores it wherever it stands (ISO/IEC 29500-3) and its save drops it
  if (ignorable) return dropped((host !== undefined ? FOREIGN_IGNORABLE[host] : undefined) ?? 'mc:Ignorable');
  return lost(host !== undefined ? FOREIGN_REFUSED[host] : undefined);
}

/** An attribute's class: an unknown `w:` one, or one in a namespace Word doesn't understand, ignorable or not, Word's save drops (76h to 76j). */
function attributeClass(ns: string | null | undefined, ignorable: boolean): SaveClass {
  if (ns === W) return dropped('76j');
  if (!ns || UNDERSTOOD_NAMESPACES.has(ns)) return lost();
  return dropped(ignorable ? '76h' : '76i');
}

/** Text's class: held by an unknown `w:` element it's absent in Word's reading (REQ-314); in a run, Word's save drops it (`72m`). */
function textClass(parent: Element): SaveClass {
  if (isUnknown(parent)) return dropped('REQ-314');
  if (isW(parent, 'r')) return dropped('72m');
  return lost();
}

function warning(kind: ContentSkippedWarning['kind'], name: string, parent: Element, path: string, ignorable: boolean, located: Node,
  partName: string | undefined, parentClass: string | undefined, namespace: string | null, saveClass: SaveClass): ContentSkippedWarning {
  return {
    code: 'part.skipped', partName, kind, name, parentName: parent.nodeName,
    ...(parentClass !== undefined ? { parentClass } : {}),
    path, ignorable, namespace, ...saveClass, ...position(located),
    message: `${partName ?? 'a part'}: ${path}: the read dropped this ${kind}, which the model has no place for`,
  };
}

/** An element and all it holds. */
export function skippedElement(element: Element, partName: string | undefined, parentClass?: string, saveClass?: SaveClass): ContentSkippedWarning {
  const ignorable = isIgnorable(element.namespaceURI, element.prefix, element);
  return warning('element', element.nodeName, element.parentNode as Element, pathOf(element), ignorable, element, partName, parentClass,
    element.namespaceURI || null, saveClass ?? elementClass(element, ignorable));
}

/** For a pass that goes beyond what Word was measured on: lost, whatever its shape (CR-025 D15). */
export const UNMEASURED: SaveClass = Object.freeze({ onSave: 'lost' });

/** An attribute of `owner`, named as the runtime reports it. */
export function skippedAttribute(name: { namespaceURI?: string; localPart: string; prefix?: string }, owner: Element, partName: string | undefined, parentClass?: string): ContentSkippedWarning {
  const shown = name.prefix ? `${name.prefix}:${name.localPart}` : name.localPart;
  const ignorable = isIgnorable(name.namespaceURI, name.prefix, owner);
  return warning('attribute', shown, owner, `${pathOf(owner)}/@${shown}`, ignorable, owner, partName, parentClass,
    name.namespaceURI || null, attributeClass(name.namespaceURI, ignorable));
}

/** A stretch of text, as a text or CDATA node. */
export function skippedText(text: string, node: Node, partName: string | undefined, parentClass?: string): ContentSkippedWarning {
  const parent = node.parentNode as Element;
  return warning('text', text.length > 80 ? text.slice(0, 80) : text, parent, `${pathOf(parent)}/text()`, false, node, partName, parentClass,
    null, textClass(parent));
}
