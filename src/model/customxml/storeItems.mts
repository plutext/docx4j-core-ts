// CR-020: docx4j's StoreItemResolver (its CR-058 D4, `35cdeaee6`, released in 17.3.5): the custom XML part a binding is
// evaluated against. The part its `w:storeItemID` names, or, where that names no part, the one its XPath fits, as Word
// repairs such a binding on open (Word 365, 22 and 23 October 2026, and checks 73 and 73b: the control shows the part's
// value, Office JS's `customXmlPart` is that part, and the save writes its ID into `w:storeItemID`). Shared by every
// reader and the bind (CR-020 D1): the text model's bound values, `XmlMapping.customXmlPart` and `customXmlNode`,
// `applyBindings`, `updateFromContentControls` and the `od:xpath` passes, so they all name the part Word binds.
import { NAMESPACE_PREFIXES } from '@docx4j/generated-objects-ts';
import { parsePrefixMappings } from './xpath.mjs';

/** The `w:storeItemID`s of the core and application properties (docx4j `BindingHandler`): not custom XML parts, never repaired. */
export const CORE_PROPERTIES_STOREITEMID = '{6C3C8BC8-F283-45AE-878A-BAB7291924A1}';
export const EXTENDED_PROPERTIES_STOREITEMID = '{6668398D-A668-4E3E-A5EB-62B293D839F1}';

/** An itemID without braces, lower-cased: how `getItem` compares ids, and so how a binding's names its part. */
export function normaliseStoreItemId(id: string): string {
  return id.trim().replace(/[{}]/g, '').toLowerCase();
}

/** Whether an ID is the core or the application properties'. */
export function isPropertiesStoreItem(id: string): boolean {
  const key = normaliseStoreItemId(id);
  return key === normaliseStoreItemId(CORE_PROPERTIES_STOREITEMID) || key === normaliseStoreItemId(EXTENDED_PROPERTIES_STOREITEMID);
}

/** One custom XML part as the resolver sees it. */
export interface StoreItemCandidate<P> {
  readonly part: P;
  /** Its itemID, as its properties part spells it. */
  readonly id: string;
  /** Its part name: the order among several that fit. */
  readonly partName: string;
  /** Its document element; none where the part can't be read, which then fits nothing. */
  readonly root: Element | null | undefined;
  /** Whether the XPath, with these prefix mappings, selects a node in it. */
  selects(xpath: string, prefixMappings: string | undefined): boolean;
}

/** The part a binding is evaluated against. */
export interface StoreItemResolution<P> {
  readonly part: P;
  /** True where the binding's ID names no part and this one fits its XPath: Word's repair. */
  readonly repaired: boolean;
  /** Where repaired, how many parts fit the XPath's first step; this one is the first, in part-name order, that selects a node. */
  readonly fitting: number;
}

/**
 * docx4j's `StoreItemResolver.resolve` rule over a package's custom XML parts:
 * - the part the ID names, the comparison `getItem`'s (case and braces ignored); the core and application properties'
 *   IDs are left to the caller, and never repaired;
 * - else the parts whose document element's namespace and local name match the XPath's first step, its prefix read with
 *   the binding's `w:prefixMappings`, else the objects package's `NAMESPACE_PREFIXES` (docx4j reads its own table); an
 *   XPath with no named first step (`//`, `*`, an axis, an attribute) fits every part. In part-name order: one fitting part
 *   is taken; of several, the first whose XPath selects a node, else the first (Word's choice in every fresh session,
 *   check 73b; CR-020 D7); none, and the binding is unresolved.
 *
 * What a stale ID resolves to is remembered per (ID, XPath, prefix mappings), since one stale ID fits different parts for
 * different XPaths (the two-parts probe: B7 part 1, B11 part 2), and forgotten when the parts change (an add, a delete).
 */
export class StoreItemResolver<P> {
  private readonly memo = new Map<string, StoreItemResolution<P> | null>();
  private signature = '';

  constructor(private readonly candidates: () => readonly StoreItemCandidate<P>[]) {}

  /** The part for a binding: named by its ID, else the one its XPath fits; undefined where none is named and none fits. */
  resolve(storeItemId: string | undefined, xpath: string | undefined, prefixMappings: string | undefined): StoreItemResolution<P> | undefined {
    if (!storeItemId) return undefined;
    const all = this.candidates();
    const key = normaliseStoreItemId(storeItemId);
    const named = all.find((candidate) => normaliseStoreItemId(candidate.id) === key);
    if (named) return { part: named.part, repaired: false, fitting: 0 };
    if (isPropertiesStoreItem(storeItemId)) return undefined;
    const signature = all.map((candidate) => `${normaliseStoreItemId(candidate.id)} ${candidate.partName}`).join('|');
    if (signature !== this.signature) {
      this.memo.clear();
      this.signature = signature;
    }
    const memoKey = `${key}|${xpath ?? ''}|${prefixMappings ?? ''}`;
    let resolved = this.memo.get(memoKey);
    if (resolved === undefined) {
      resolved = repair(all, xpath ?? '', prefixMappings) ?? null;
      this.memo.set(memoKey, resolved);
    }
    return resolved ?? undefined;
  }
}

function repair<P>(all: readonly StoreItemCandidate<P>[], xpath: string, prefixMappings: string | undefined): StoreItemResolution<P> | undefined {
  const step = firstStep(xpath, prefixMappings);
  const fitting = all.filter((candidate) => {
    const root = candidate.root;
    if (!root) return false;
    return step === undefined || ((root.namespaceURI ?? '') === step.namespaceURI && (root.localName ?? root.nodeName) === step.localName);
  }).sort((a, b) => (a.partName < b.partName ? -1 : a.partName > b.partName ? 1 : 0));
  if (fitting.length === 0) return undefined;
  if (fitting.length === 1) return { part: fitting[0]!.part, repaired: true, fitting: 1 };
  const selecting = fitting.find((candidate) => {
    try {
      return candidate.selects(xpath, prefixMappings);
    } catch {
      return false;
    }
  });
  return { part: (selecting ?? fitting[0]!).part, repaired: true, fitting: fitting.length };
}

/** The objects package's prefixes, prefix to namespace: what a prefix the binding doesn't map reads as. */
const NAMESPACES_BY_PREFIX: Record<string, string> = (() => {
  const out: Record<string, string> = {};
  for (const [uri, prefix] of Object.entries(NAMESPACE_PREFIXES)) if (prefix !== '' && !(prefix in out)) out[prefix] = uri;
  return out;
})();

/**
 * The XPath's first step as its namespace and local name, with the binding's prefix mappings (`/p:root[1]/p:b7[1]` with
 * `xmlns:p='urn:x'` is `urn:x` and `root`); undefined where the XPath does not start with a named step, when every part
 * fits. docx4j's `StoreItemResolver.firstStep`.
 */
export function firstStep(xpath: string | undefined, prefixMappings: string | undefined): { namespaceURI: string; localName: string } | undefined {
  if (xpath === undefined) return undefined;
  let s = xpath.trim();
  if (s.startsWith('//')) return undefined;
  if (s.startsWith('/')) s = s.slice(1);
  const end = s.search(/[/[(\s]/);
  const step = end < 0 ? s : s.slice(0, end);
  if (step === '' || step === '*' || step === '.' || step === '..' || step.includes('::') || step.startsWith('@')) return undefined;
  const colon = step.indexOf(':');
  const prefix = colon < 0 ? '' : step.slice(0, colon);
  const localName = colon < 0 ? step : step.slice(colon + 1);
  if (localName === '' || localName === '*') return undefined;
  const namespaceURI = prefix === '' ? '' : (parsePrefixMappings(prefixMappings)[prefix] ?? NAMESPACES_BY_PREFIX[prefix] ?? '');
  return { namespaceURI, localName };
}
