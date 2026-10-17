// Markup compatibility preprocessing (ECMA-376 Part 3; docx4j mc-preprocessor.xslt). Runs over a
// part's DOM before it is unmarshalled: each mc:AlternateContent is replaced by the content of
// the first mc:Choice whose Requires prefixes all name namespaces the object model understands,
// else by its mc:Fallback content. This is what Word does on open. It runs only for parts that are
// unmarshalled, so untouched parts keep every branch.
//
// Two kinds are left whole instead (CR-010 phase 1), so a save writes them back with every branch:
// - those under a host in `KEPT_HOSTS`, where the model binds the element and nothing in this
//   package reads the branches' content (a chart part's c14:style, a workbook's x15ac:absPath,
//   which promoting lost outright);
// - those nothing can be selected from (Choices only, none understood). ISO removes them. Promoting
//   the first Choice, the reading rule (the text extraction paper's D4 item 4), would put content
//   in no understood namespace where its host may refuse it, and does: a w:r holding one then
//   fails to unmarshal. Kept, the model types the element where its host accepts one and drops it
//   where it does not, as it drops any unexpected element, so keeping never loses more than
//   removing did.
import type { XmlPart } from '../../parts/XmlPart.mjs';
import { UNDERSTOOD_NAMESPACES } from './understood.mjs';

export const MCE_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/**
 * The hosts whose mc:AlternateContent is kept whole on load, as `{namespace}localName`
 * (CR-010 phase 1): a chart's `c:chartSpace` and SpreadsheetML's `workbook`.
 */
export const KEPT_HOSTS: ReadonlySet<string> = new Set([
  '{http://schemas.openxmlformats.org/drawingml/2006/chart}chartSpace',
  '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}workbook',
]);

export interface McePreprocessOptions {
  /** Namespaces counted as understood; defaults to the object model's modules. */
  understood?: ReadonlySet<string>;
  /** Hosts, as `{namespace}localName`, whose mc:AlternateContent is left whole; defaults to `KEPT_HOSTS`. */
  keptHosts?: ReadonlySet<string>;
}

/**
 * What load reports about a part (CR-010 phase 1; the text extraction paper's warning channel).
 * `mce.branches_differ`: an mc:AlternateContent whose branches' text differs, so a reader that
 * takes another branch than this package's reads something else. `selected` is the branch this
 * package took (`Choice(wps)`, `Fallback`); `branches` gives every branch's text, in document
 * order, as the w:t, a:t and m:t text it holds, with whitespace collapsed.
 */
export interface McBranchesDifferWarning {
  code: 'mce.branches_differ';
  partName: string | undefined;
  /** Where the element starts in the part, where the DOM records it (xmldom does; a browser's DOMParser does not). */
  line?: number;
  column?: number;
  selected: string;
  branches: { branch: string; text: string }[];
  message: string;
}

/** A warning load records on a part. */
export type PartWarning = McBranchesDifferWarning;

/** What one pass over a DOM did. */
export interface MceResult {
  resolved: number;
  kept: number;
  warnings: PartWarning[];
}

/** The default preprocessor, over the object model's namespaces. Returns the part's warnings. */
export function mcePreprocess(doc: Document, part?: XmlPart<unknown>): PartWarning[] {
  return preprocessAlternateContent(doc, { partName: part?.partName.name }).warnings;
}

/** A preprocessor over a namespace set (or kept hosts) of your own. */
export function createMcePreprocessor(options: McePreprocessOptions): (doc: Document, part?: XmlPart<unknown>) => PartWarning[] {
  return (doc, part) => preprocessAlternateContent(doc, { ...options, partName: part?.partName.name }).warnings;
}

/**
 * Resolves every mc:AlternateContent under a node, outermost first, but those `KEPT_HOSTS` and the
 * nothing-selectable rule keep (see the file comment). Returns the number resolved.
 */
export function resolveAlternateContent(root: Node, understood: ReadonlySet<string> = UNDERSTOOD_NAMESPACES): number {
  return preprocessAlternateContent(root, { understood }).resolved;
}

/** The pass itself: what it resolved, kept and reported. */
export function preprocessAlternateContent(root: Node, options: McePreprocessOptions & { partName?: string } = {}): MceResult {
  const understood = options.understood ?? UNDERSTOOD_NAMESPACES;
  const keptHosts = options.keptHosts ?? KEPT_HOSTS;
  const result: MceResult = { resolved: 0, kept: 0, warnings: [] };
  const visit = (node: Node): void => {
    const hostKept = node.nodeType === 1 && keptHosts.has(`{${(node as Element).namespaceURI ?? ''}}${(node as Element).localName}`);
    // A work list, not the live child list: resolving replaces a child with the branch's content,
    // which is visited in its turn (it may hold, or be, another mc:AlternateContent).
    const queue: Node[] = [];
    for (let c = node.firstChild; c; c = c.nextSibling) queue.push(c);
    for (let i = 0; i < queue.length; i++) {
      if (queue[i].nodeType !== 1) continue;
      const el = queue[i] as Element;
      if (!isAlternateContent(el)) { visit(el); continue; }
      const { choices, fallback } = branchesOf(el);
      const chosen = choices.find((c) => choiceUnderstood(c, understood));
      const branch = chosen ?? fallback;
      if (hostKept || (branch === undefined && choices.length > 0)) { result.kept++; continue; }
      const warning = branchesDiffer(el, choices, fallback, branch, understood, options.partName);
      if (warning) result.warnings.push(warning);
      const moved: Node[] = [];
      if (branch) for (let c = branch.firstChild; c; c = c.nextSibling) moved.push(c);
      for (const n of moved) node.insertBefore(n, el);
      node.removeChild(el);
      queue.splice(i + 1, 0, ...moved);
      result.resolved++;
    }
  };
  visit(root);
  return result;
}

function isAlternateContent(el: Element): boolean {
  return el.namespaceURI === MCE_NS && el.localName === 'AlternateContent';
}

function branchesOf(ac: Element): { choices: Element[]; fallback: Element | undefined } {
  const choices: Element[] = [];
  let fallback: Element | undefined;
  for (let c = ac.firstChild; c; c = c.nextSibling) {
    if (c.nodeType !== 1) continue;
    const el = c as Element;
    if (el.namespaceURI !== MCE_NS) continue;
    if (el.localName === 'Choice') choices.push(el);
    else if (el.localName === 'Fallback' && !fallback) fallback = el;
  }
  return { choices, fallback };
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

/** The text elements a branch's text is read from: WordprocessingML's, DrawingML's and OMML's `t`. */
const TEXT_NAMESPACES: ReadonlySet<string> = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://schemas.openxmlformats.org/officeDocument/2006/math',
]);

/**
 * The text a branch holds: its w:t, a:t and m:t, whitespace collapsed. A nested
 * mc:AlternateContent contributes the branch the same rule selects, or its first Choice when
 * nothing is selectable (the reading rule).
 */
function textOfBranch(node: Node, understood: ReadonlySet<string>): string {
  const out: string[] = [];
  const walk = (n: Node): void => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const el = c as Element;
      if (isAlternateContent(el)) {
        const { choices, fallback } = branchesOf(el);
        const branch = choices.find((x) => choiceUnderstood(x, understood)) ?? fallback ?? choices[0];
        if (branch) walk(branch);
      } else if (el.localName === 't' && TEXT_NAMESPACES.has(el.namespaceURI ?? '')) out.push(el.textContent ?? '');
      else walk(el);
    }
  };
  walk(node);
  return out.join('').replace(/\s+/g, ' ').trim();
}

function branchName(branch: Element): string {
  return branch.localName === 'Choice' ? `Choice(${(branch.getAttribute('Requires') ?? '').trim()})` : 'Fallback';
}

function branchesDiffer(ac: Element, choices: Element[], fallback: Element | undefined, selected: Element | undefined,
  understood: ReadonlySet<string>, partName: string | undefined): McBranchesDifferWarning | undefined {
  const all = fallback ? [...choices, fallback] : choices;
  if (all.length < 2 || selected === undefined) return undefined;
  const branches = all.map((b) => ({ branch: branchName(b), text: textOfBranch(b, understood) }));
  const taken = branches[all.indexOf(selected)];
  if (branches.every((b) => b.text === taken.text)) return undefined;
  const located = ac as Element & { lineNumber?: number; columnNumber?: number };
  const where = typeof located.lineNumber === 'number' ? ` (line ${located.lineNumber}, column ${located.columnNumber})` : '';
  const others = branches.filter((b) => b !== taken).map((b) => `${b.branch} ${JSON.stringify(abbreviate(b.text))}`).join(', ');
  return {
    code: 'mce.branches_differ',
    partName,
    ...(typeof located.lineNumber === 'number' ? { line: located.lineNumber, column: located.columnNumber } : {}),
    selected: taken.branch,
    branches,
    message: `${partName ?? 'a part'}${where}: mc:AlternateContent's branches say different things; ${taken.branch} was read (${JSON.stringify(abbreviate(taken.text))}), not ${others}`,
  };
}

function abbreviate(text: string): string {
  return text.length > 60 ? `${text.slice(0, 57)}...` : text;
}
