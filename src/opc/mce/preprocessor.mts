// Markup compatibility preprocessing (ECMA-376 Part 3; docx4j mc-preprocessor.xslt). Runs over a
// part's DOM before it is unmarshalled. Since CR-010 phase 4 load keeps every mc:AlternateContent,
// as docx4j does since its CR-021, and the default pass (`mceKeep`) changes nothing: it reports what
// resolving would, the `mce.branches_differ` and `mce.nothing_selected` warnings below, so a reader learns
// where the branch the content API reads says something other than one it doesn't, and where there is none.
//
// Resolving (`mcePreprocess: true` on load, the pass `mcePreprocess`) replaces each element by the
// content of the first mc:Choice whose Requires prefixes all name namespaces in `UNDERSTOOD_NAMESPACES`,
// what Word 365 was measured to understand, else by its mc:Fallback content, else, in a word-processing
// part, by nothing (WD REQ-305, CR-029). This is what Word does on open. It runs only for parts that are
// unmarshalled, so untouched parts keep every branch. A `w:tc` a removal leaves with no block gets an empty
// `w:p`, since Word repairs a cell with none (the `tc-empty` round).
//
// Three kinds are left whole even so, so a save writes them back with every branch:
// - those under a host in `KEPT_HOSTS`, where the model binds the element and nothing in this
//   package reads the branches' content (a chart part's c14:style, a workbook's x15ac:absPath,
//   which promoting lost outright);
// - those nothing can be selected from in any other part (Choices only, none understood): the understood
//   set was measured in a word-processing package, so a Choice in a slide or a sheet could otherwise go
//   for a namespace nobody asked Word about (CR-029 D3);
// - one nothing can be selected from at a part's root, whose removal would leave a part with no root. CR-028
//   reads the root.
import type { XmlPart } from '../../parts/XmlPart.mjs';
import { UNDERSTOOD_NAMESPACES } from './understood.mjs';
import type { BrOutsideRunWarning } from '../repair/breaks.mjs';
import type { UnexpectedRootWarning } from '../repair/mainRoot.mjs';
import type { ContentSkippedWarning } from '../skipped.mjs';
import type { RelationshipsNamespaceWarning, RelationshipsNotReadWarning } from '../exceptions.mjs';

export const MCE_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * The hosts whose mc:AlternateContent is kept whole on load, as `{namespace}localName`
 * (CR-010 phase 1): a chart's `c:chartSpace` and SpreadsheetML's `workbook`.
 */
export const KEPT_HOSTS: ReadonlySet<string> = new Set([
  '{http://schemas.openxmlformats.org/drawingml/2006/chart}chartSpace',
  '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}workbook',
]);

export interface McePreprocessOptions {
  /** Namespaces counted as understood; defaults to `UNDERSTOOD_NAMESPACES`, the 114 Word 365 was measured to understand (WD REQ-306). */
  understood?: ReadonlySet<string>;
  /** Hosts, as `{namespace}localName`, whose mc:AlternateContent is left whole; defaults to `KEPT_HOSTS`. */
  keptHosts?: ReadonlySet<string>;
  /**
   * False leaves every element whole and reports what resolving would (CR-010 phase 4): the same
   * warnings, from the same elements, and no change to the DOM. Defaults to true.
   */
  resolve?: boolean;
  /**
   * Whether this is a word-processing part, where an element nothing can be selected from is reported and removed (WD
   * REQ-305); in any other it is left whole and said nothing of (CR-029 D3). Defaults to true; the preprocessors read it
   * from the part's content type.
   */
  wordProcessing?: boolean;
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

/**
 * `mce.nothing_selected` (WD REQ-305, CR-029): an mc:AlternateContent with Choices, none understood, and no Fallback, so
 * that the element and all it holds are absent, as ISO and Word read it. `namespaces` gives what its Choices require, each
 * once, in order, a prefix nothing binds having none. `removed` says a resolving load took it out; `root` that it is a part's
 * root, which is left whole. A reader's walk meets neither, so the text model carries such a warning with no address.
 */
export interface McNothingSelectedWarning {
  code: 'mce.nothing_selected';
  partName: string | undefined;
  /** Where the element starts in the part, where the DOM records it (xmldom does; a browser's DOMParser does not). */
  line?: number;
  column?: number;
  namespaces: string[];
  removed: boolean;
  root: boolean;
  message: string;
}

/**
 * `mce.undeclared_prefix` (WD REQ-323, CR-030): an `mc:Ignorable` listing a prefix with no namespace declaration in scope on the
 * element carrying it. Such a prefix names no namespace, so it is read as though the list didn't hold it: nothing it seems to
 * name is ignorable. Word repairs such a file and loses nothing (the `undeclared-ignorable` round). One warning a part, with
 * every such prefix once, in document order. A declaration on a descendant doesn't put a prefix in scope. A save that writes
 * the part anew holds none: the objects facade drops a prefix it doesn't know and declares one it does, and a part copied from
 * its source keeps it as written.
 */
export interface McUndeclaredPrefixWarning {
  code: 'mce.undeclared_prefix';
  partName: string | undefined;
  prefixes: string[];
  message: string;
}

/**
 * An attribute value of Word's own WordprocessingML that its type refuses, read as Word 365 reads it
 * rather than failing the part (CR-015 D8; the objects facade's `wordInvalidValue`, its CR-013):
 * `value.kept.exact` where the text, trimmed of XML whitespace, is a form its type reads (`" false"`),
 * else `value.kept.invalid` (`"yes"`, `"Single"`). The codes are this package's (CR-015 D8); the reasons, exact and invalid,
 * are docx4j-python CR-018's (its `KeptValue.reason`). An integer attribute's value Word reads by its own rule (the objects
 * package's CR-015, CR-017 P1): `value.truncated` for a decimal, cut toward zero as Word cuts it (`"720.5"` reads 720);
 * `value.refused` for a form Word refuses a document for (`" 240 "`, `"0x10"`), dropped; `""` and the forms with no
 * digit (`"."`, `"+"`) read as 0, `value.kept.invalid`.
 */
export interface ValueKeptWarning {
  code: 'value.kept.exact' | 'value.kept.invalid' | 'value.truncated' | 'value.refused';
  partName: string | undefined;
  /** Where the element starts in the part, where the DOM records it (xmldom does; a browser's DOMParser does not). */
  line?: number;
  column?: number;
  /** The element and the attribute, as the part writes them: `w:b`, `w:val`. */
  element: string;
  attribute: string;
  /** The attribute's text, as written. */
  text: string;
  /**
   * What the tree holds instead: `true` for an on/off value (Word reads only exactly `false`, `off` and
   * `0` as off), `'none'` for `w:u`'s on/off words, an integer Word's reading of a decimal or of `""`, and
   * undefined where the value was dropped, so the attribute reads as absent (a style's value beneath shows).
   */
  value: boolean | string | number | undefined;
  message: string;
}

/**
 * An integer attribute value Word 365 reads differently from the tree (CR-017 P2 and P3; the objects facade's
 * `onOutOfRange`, its CR-016, docx4j CR-066's rounds 5 and 6): outside the width Word stores the attribute in, 16 bits
 * for most, where Word reads it wrapped, sometimes then clamped, with no error or repair. The tree keeps the value as
 * written, and a save writes it as held: the warning says what Word would show. The code and the fields are shared
 * with docx4j-python (its CR-020 section 4.7, `word_reads` there), one name across the ports for the one condition
 * (CR-017 E1). On load the warning is on the part, with the line and column where the DOM records them; on save it is
 * on `pkg.lastSave.outOfRange`, with the `path` from the part's root.
 */
export interface ValueOutOfRangeWarning {
  code: 'value.out_of_range';
  partName: string | undefined;
  /** Where the element starts in the part, where the DOM records it (xmldom does; a browser's DOMParser does not). */
  line?: number;
  column?: number;
  /** On save, the attribute's path from the part's root: `/w:document/w:body[1]/w:p[3]/w:pPr[1]/w:ind[1]/@w:left`. */
  path?: string;
  /** The element and the attribute as the part writes them (`w:ind`, `w:left`). */
  element: string;
  attribute: string;
  /** The value the tree holds, and a save writes. */
  value: number;
  /** What Word reads for it, after the wrap and any clamp; null for an attribute no round measured. */
  wordReads: number | null;
  /** The attribute Word reads the value in, where the shared field's sign moved it: `w:hanging` for a `w:firstLine` of 32768. */
  readAs?: string;
  /** For a line, a row height or a frame's height, the rule Word reads with the value (`exact`, `atLeast`, `auto`). */
  rule?: 'auto' | 'exact' | 'atLeast';
  /** How Word stores the attribute, the objects package's wording: `int16`, `uint16, then 144..31681`, `int16, the sign the rule`, ... */
  storage: string;
  /** The values Word keeps as written there, `[min, max]` in the attribute's own unit (`keptRange`); null where none is known. */
  range: readonly [number, number] | null;
  message: string;
}

/**
 * A child its container never admits (CR-016 D7, the objects facade's `onMisplaced`, its CR-017; docx4j CR-064's rules):
 * a run straight in the body, a table in a paragraph, a cell beside a table's rows, a string in a run, or a name and
 * value that disagree. On load, one per child the facade reports as it reads the part, the tree keeping what it read
 * (a string in a run excepted, which the runtime drops: `text` carries it); on a save with `checkContentOnSave` off,
 * one per child written as held, on `pkg.lastSave.misplaced`; with it on, the save is refused with them
 * (`MisplacedContentException`). The code and the fields are docx4j-python's (its CR-020 section 11.12), one code across
 * the ports; `rule`, `text` and `line` are this package's own.
 */
export interface MisplacedContentWarning {
  code: 'target.misplaced';
  partName: string | undefined;
  /** Where the child (else its container) starts in the part, where the DOM records it (xmldom does). */
  line?: number;
  column?: number;
  /** On save, the child's path from the part's root (`/w:document/w:body[1]/w:p[1]/w:r[1]/w:tbl[1]`); the container's for a string. */
  path?: string;
  /** The container's class, docx4j's name (`P`, `R`, `Tbl`, `CTSdtContentRun`): the objects `TYPE_NAME`'s last segment. */
  holder: string;
  /** The child's position in the list holding it. */
  index: number;
  /** The child's class the same way (`Tbl`, `R`), where it has one; absent for a string. */
  item?: string;
  /** The child's element as the part writes it (`w:tbl`); absent for a string. */
  element?: string;
  /** `misplaced`, a class the container refuses; `name-value`, a value whose class isn't the one its name declares there. */
  rule: 'misplaced' | 'name-value';
  /** For a string in a container that admits none, the string (on load; a save never sees it). */
  text?: string;
  /** docx4j's wording, with the fix: "You can't add a run (w:r) directly to the body (w:body); ...", after the part and the place. */
  message: string;
}

/** A warning load records on a part. */
export type PartWarning = McBranchesDifferWarning | McNothingSelectedWarning | McUndeclaredPrefixWarning | ValueKeptWarning | ValueOutOfRangeWarning | MisplacedContentWarning | BrOutsideRunWarning | UnexpectedRootWarning | ContentSkippedWarning | RelationshipsNamespaceWarning | RelationshipsNotReadWarning;

/**
 * What one pass over a DOM did. `kept` counts the elements it met and left whole; it doesn't enter a
 * branch the rule doesn't select, so an element nested there is not counted.
 */
export interface MceResult {
  resolved: number;
  kept: number;
  warnings: PartWarning[];
}

/**
 * Load's preprocessor since CR-010 phase 4: every mc:AlternateContent is kept, and the part's
 * warnings are what resolving would report. Returns them.
 */
export function mceKeep(doc: Document, part?: XmlPart<unknown>): PartWarning[] {
  return preprocessAlternateContent(doc, { resolve: false, partName: part?.partName.name, wordProcessing: isWordProcessing(part) }).warnings;
}

/** The resolving preprocessor (`mcePreprocess: true`), over `UNDERSTOOD_NAMESPACES`. Returns the part's warnings. */
export function mcePreprocess(doc: Document, part?: XmlPart<unknown>): PartWarning[] {
  return preprocessAlternateContent(doc, { partName: part?.partName.name, wordProcessing: isWordProcessing(part) }).warnings;
}

/** A preprocessor over a namespace set (or kept hosts) of your own. */
export function createMcePreprocessor(options: McePreprocessOptions): (doc: Document, part?: XmlPart<unknown>) => PartWarning[] {
  return (doc, part) => preprocessAlternateContent(doc, { wordProcessing: isWordProcessing(part), ...options, partName: part?.partName.name }).warnings;
}

/** Whether a part is WordprocessingML's, by its content type (the transitional and macro-enabled forms); a bare node is taken as one. */
function isWordProcessing(part: XmlPart<unknown> | undefined): boolean {
  return part === undefined || /wordprocessingml|vnd\.ms-word/.test(part.contentType);
}

/**
 * Resolves every mc:AlternateContent under a node, outermost first, but those the file comment keeps. Returns the number
 * resolved, one nothing was selected from among them.
 */
export function resolveAlternateContent(root: Node, understood: ReadonlySet<string> = UNDERSTOOD_NAMESPACES): number {
  return preprocessAlternateContent(root, { understood }).resolved;
}

/**
 * The pass itself: what it resolved, kept and reported. With `resolve: false` it visits the elements
 * resolving would, outermost first and into the branch the rule selects, reports the same warnings,
 * and leaves the DOM as it was.
 */
export function preprocessAlternateContent(root: Node, options: McePreprocessOptions & { partName?: string } = {}): MceResult {
  const understood = options.understood ?? UNDERSTOOD_NAMESPACES;
  const keptHosts = options.keptHosts ?? KEPT_HOSTS;
  const resolve = options.resolve ?? true;
  const wordProcessing = options.wordProcessing ?? true;
  const result: MceResult = { resolved: 0, kept: 0, warnings: [] };
  const undeclared = new Set<string>();
  // a list read where its element stood in the source: resolving moves a branch's content out of the mc:Choice that may declare
  // its prefixes, as Excel's drawings declare a14 there and list it on an a:srgbClr inside
  const asRead = new Map<Element, string[]>();
  const visit = (node: Node): void => {
    if (node.nodeType === 1) for (const p of asRead.get(node as Element) ?? undeclaredIgnorable(node as Element)) undeclared.add(p);
    const hostKept = node.nodeType === 1 && keptHosts.has(`{${(node as Element).namespaceURI ?? ''}}${(node as Element).localName}`);
    // A work list, not the live child list: resolving replaces a child with the branch's content,
    // which is visited in its turn (it may hold, or be, another mc:AlternateContent).
    const queue: Node[] = [];
    for (let c = node.firstChild; c; c = c.nextSibling) queue.push(c);
    let removed = false;
    for (let i = 0; i < queue.length; i++) {
      if (queue[i].nodeType !== 1) continue;
      const el = queue[i] as Element;
      if (!isAlternateContent(el)) { visit(el); continue; }
      const { choices, fallback } = branchesOf(el);
      const chosen = choices.find((c) => choiceUnderstood(c, understood));
      const branch = chosen ?? fallback;
      const nothing = branch === undefined && choices.length > 0;
      if (hostKept || (nothing && !wordProcessing)) { result.kept++; continue; }
      const atRoot = el.parentNode?.nodeType === 9;
      const warning = nothing ? nothingSelected(el, choices, resolve && !atRoot, atRoot, options.partName) : branchesDiffer(el, choices, fallback, branch, understood, options.partName);
      if (warning) result.warnings.push(warning);
      // a part's root is left whole, as a resolving load can't put a branch's content where the root stood beside it, and
      // reading the branch as the part is WD REQ-317's, not built (CR-028 D10)
      if (!resolve || atRoot) {
        result.kept++;
        if (branch) visit(branch);
        continue;
      }
      const moved: Node[] = [];
      if (branch) readInPlace(branch, asRead);
      if (branch) for (let c = branch.firstChild; c; c = c.nextSibling) moved.push(c);
      for (const n of moved) node.insertBefore(n, el);
      node.removeChild(el);
      removed = true;
      queue.splice(i + 1, 0, ...moved);
      result.resolved++;
    }
    if (removed) keepACellsParagraph(node);
  };
  visit(root);
  if (undeclared.size > 0) result.warnings.push(undeclaredPrefix([...undeclared], options.partName));
  return result;
}

/** The prefixes an element's `mc:Ignorable` lists that have no declaration in scope on it, where it stands now (WD REQ-323). */
function undeclaredIgnorable(el: Element): string[] {
  const listed = el.getAttributeNS(MCE_NS, 'Ignorable');
  if (!listed) return [];
  return listed.trim().split(/\s+/).filter((prefix) => prefix !== '' && !el.lookupNamespaceURI(prefix));
}

/**
 * Before a branch's content is moved out, each list in it read where it stands, unless an outer branch's move already read it
 * there. A list in a nested branch not taken is read too, and never visited, so never said.
 */
function readInPlace(branch: Element, into: Map<Element, string[]>): void {
  const all = branch.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    const el = all.item(i)!;
    if (!into.has(el) && el.hasAttributeNS(MCE_NS, 'Ignorable')) into.set(el, undeclaredIgnorable(el));
  }
}

function undeclaredPrefix(prefixes: string[], partName: string | undefined): McUndeclaredPrefixWarning {
  return {
    code: 'mce.undeclared_prefix',
    partName,
    prefixes,
    message: `${partName ?? 'a part'}: mc:Ignorable lists ${prefixes.join(' ')}, which no declaration in scope binds, so it names nothing and is read as though the list didn't hold it (WD REQ-323)`,
  };
}

/**
 * A `w:tc` left with no block gets the paragraph Word's repair gives a cell with none (CR-029 D4 and D9): single-spaced with no
 * space after, `w:spacing w:after="0" w:line="240" w:lineRule="auto"`, whatever the document's defaults, its Normal style or a
 * table style say (the `cell-paragraph` round). Word writes no spacing where the paragraph already reads so (`cp-02`); there the
 * element changes nothing drawn. Word repairs such a cell on open (`te-01` to `te-03`), and opens this silently.
 */
function keepACellsParagraph(node: Node): void {
  if (node.nodeType !== 1) return;
  const cell = node as Element;
  if (cell.namespaceURI !== W || cell.localName !== 'tc') return;
  for (let c = cell.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1 && (c as Element).localName !== 'tcPr') return;
  }
  const doc = cell.ownerDocument;
  const name = (local: string): string => (cell.prefix ? `${cell.prefix}:${local}` : local);
  const spacing = doc.createElementNS(W, name('spacing'));
  for (const [attribute, value] of [['after', '0'], ['line', '240'], ['lineRule', 'auto']] as const) spacing.setAttributeNS(W, name(attribute), value);
  const pPr = doc.createElementNS(W, name('pPr'));
  pPr.appendChild(spacing);
  const p = doc.createElementNS(W, name('p'));
  p.appendChild(pPr);
  cell.appendChild(p);
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
 * mc:AlternateContent contributes the branch the same rule selects, and nothing where it selects nothing.
 */
function textOfBranch(node: Node, understood: ReadonlySet<string>): string {
  const out: string[] = [];
  const walk = (n: Node): void => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const el = c as Element;
      if (isAlternateContent(el)) {
        const { choices, fallback } = branchesOf(el);
        const branch = choices.find((x) => choiceUnderstood(x, understood)) ?? fallback;
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

/** `mce.nothing_selected`'s record: the namespaces the Choices require, each once, in order (the objects package's `NothingSelected`). */
function nothingSelected(ac: Element, choices: Element[], removed: boolean, root: boolean, partName: string | undefined): McNothingSelectedWarning {
  const namespaces = new Set<string>();
  for (const choice of choices) {
    for (const prefix of (choice.getAttribute('Requires') ?? '').trim().split(/\s+/)) {
      const ns = prefix === '' ? null : choice.lookupNamespaceURI(prefix);
      if (ns) namespaces.add(ns);
    }
  }
  const located = ac as Element & { lineNumber?: number; columnNumber?: number };
  const where = typeof located.lineNumber === 'number' ? ` (line ${located.lineNumber}, column ${located.columnNumber})` : '';
  return {
    code: 'mce.nothing_selected',
    partName,
    ...(typeof located.lineNumber === 'number' ? { line: located.lineNumber, column: located.columnNumber } : {}),
    namespaces: [...namespaces],
    removed,
    root,
    message: `${partName ?? 'a part'}${where}: mc:AlternateContent selects nothing: no Choice is understood (${[...namespaces].join(', ') || 'no namespace bound'}) and it has no Fallback, so it reads as absent (REQ-305)`,
  };
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
