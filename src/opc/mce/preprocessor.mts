// Markup compatibility preprocessing (ECMA-376 Part 3; docx4j mc-preprocessor.xslt). Runs over a
// part's DOM before it is unmarshalled. Since CR-010 phase 4 load keeps every mc:AlternateContent,
// as docx4j does since its CR-021, and the default pass (`mceKeep`) changes nothing: it reports what
// resolving would, the `mce.branches_differ` warning below, so a reader learns where the branch the
// content API reads says something other than one it doesn't.
//
// Resolving (`mcePreprocess: true` on load, the pass `mcePreprocess`) replaces each element by the
// content of the first mc:Choice whose Requires prefixes all name namespaces in
// `READ_UNDERSTOOD_NAMESPACES` (the object model's, and the four Word 365 understands that no module
// binds: CR-013 D3), else by its mc:Fallback content. This is what Word does on open. It runs only for
// parts that are unmarshalled, so untouched parts keep every branch.
//
// Two kinds are left whole even so (CR-010 phase 1), so a save writes them back with every branch:
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
import { READ_UNDERSTOOD_NAMESPACES } from './understood.mjs';

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
  /** Namespaces counted as understood; defaults to `READ_UNDERSTOOD_NAMESPACES` (CR-013 D3). */
  understood?: ReadonlySet<string>;
  /** Hosts, as `{namespace}localName`, whose mc:AlternateContent is left whole; defaults to `KEPT_HOSTS`. */
  keptHosts?: ReadonlySet<string>;
  /**
   * False leaves every element whole and reports what resolving would (CR-010 phase 4): the same
   * warnings, from the same elements, and no change to the DOM. Defaults to true.
   */
  resolve?: boolean;
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
export type PartWarning = McBranchesDifferWarning | ValueKeptWarning | ValueOutOfRangeWarning | MisplacedContentWarning;

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
  return preprocessAlternateContent(doc, { resolve: false, partName: part?.partName.name }).warnings;
}

/** The resolving preprocessor (`mcePreprocess: true`), over `READ_UNDERSTOOD_NAMESPACES`. Returns the part's warnings. */
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
export function resolveAlternateContent(root: Node, understood: ReadonlySet<string> = READ_UNDERSTOOD_NAMESPACES): number {
  return preprocessAlternateContent(root, { understood }).resolved;
}

/**
 * The pass itself: what it resolved, kept and reported. With `resolve: false` it visits the elements
 * resolving would, outermost first and into the branch the rule selects, reports the same warnings,
 * and leaves the DOM as it was.
 */
export function preprocessAlternateContent(root: Node, options: McePreprocessOptions & { partName?: string } = {}): MceResult {
  const understood = options.understood ?? READ_UNDERSTOOD_NAMESPACES;
  const keptHosts = options.keptHosts ?? KEPT_HOSTS;
  const resolve = options.resolve ?? true;
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
      if (!resolve) {
        result.kept++;
        if (branch) visit(branch);
        continue;
      }
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
