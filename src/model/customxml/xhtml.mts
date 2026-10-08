// SPDX-License-Identifier: Apache-2.0
//
// The bind step for escaped XHTML (CR-005 section 8.2; the specification's section 9). A control
// whose tag says `od:ContentType=application/xhtml+xml` and names an `od:xpath` entry carries no
// `w:dataBinding` (Word allows none on a block-level control, and block content needs one; docx4j's
// `bind.xslt` reads the XPath from the entry and leaves `w:sdtPr` as it is). The entry's XPath over
// its data part gives the node, whose string value is XHTML markup; the `html` module converts it
// and writes it as WordprocessingML, which replaces the control's content: the blocks for a
// block-level control, the first paragraph's inline content for a run-level one (docx4j's
// `BindingTraverserXSLT.convertXHTML` refuses block content there), the rest reported (REQ-076,
// REC-009). The control's own run properties apply where the markup states none, as a text
// binding's do. The reverse direction leaves such a control alone, as docx4j's does.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { deepCopy } from '@docx4j/generated-objects-ts';
import { Namespaces } from '../../parts/Namespaces.mjs';
import type { PropertyResolver } from '../properties/index.mjs';
import type { Element } from '../content/tree.mjs';
import { typeNameOf, linkParents, childrenOf } from '../content/tree.mjs';
import type { ContentControl } from '../content/ContentControl.mjs';
import { contentOf } from '../content/ooxml.mjs';
import { convertHtml, styleLookupOf, type PasteResult, type StyleNaming } from '../../html/convert.mjs';
import { blocksToXml, needsOf, HYPERLINK_STYLE } from '../../html/elements.mjs';
import type { CustomXmlPartLookup } from './XmlMapping.mjs';
import { tagParamsOf, xpathsEntriesOf, XHTML_CONTENT_TYPE } from './opendope.mjs';
import type { BindingResult } from './bindings.mjs';

/** What the XHTML pass needs of the package. */
export interface XhtmlHost {
  numbering: { newList(options?: { bullet?: boolean }): Promise<string>; restart(numId: string): string };
  styles: { ensure(ids: readonly string[]): Promise<string[]> };
  requireStyles(ids: Iterable<string>): void;
  getPropertyResolver(): Promise<PropertyResolver>;
  getMainDocumentPart(): { styleDefinitionsPart?: { readContents(): Promise<wml.Styles> } | undefined };
}

export interface XhtmlBindingOptions {
  /** The HTML parser: a browser's `DOMParser` by default; jsdom's or linkedom's in Node. */
  parser?: (html: string) => Document;
}

/** Whether a control is bound to XHTML through its tag. */
export function isXhtmlBound(control: ContentControl): boolean {
  const params = tagParamsOf(control.tag);
  return params.get('od:ContentType') === XHTML_CONTENT_TYPE && params.has('od:xpath');
}

async function lookupFor(host: XhtmlHost): Promise<ReturnType<typeof styleLookupOf>> {
  const entries: StyleNaming[] = [];
  const part = host.getMainDocumentPart().styleDefinitionsPart;
  if (part) {
    const styles = await part.readContents();
    for (const s of styles.style ?? []) {
      if (!s.styleId) continue;
      if (s.type !== undefined && s.type !== 'paragraph' && s.type !== 'character') continue;
      if (s.hidden !== undefined && s.hidden.val !== false) continue;
      entries.push({ id: s.styleId, name: s.name?.val ?? s.styleId });
    }
  }
  const resolver = await host.getPropertyResolver();
  return styleLookupOf(entries, resolver.getStyle('TableGrid') ? ['TableGrid'] : []);
}

/** The object behind an element: a parsed fragment's elements are wrapped (`{ name, value }`), the tree's own may not be. */
function innerOf(element: Element): object {
  const e = element as unknown as { name?: unknown; value?: object };
  return e.name !== undefined && e.value !== undefined && typeof e.value === 'object' ? e.value : (element as unknown as object);
}

/** The control's run properties onto every run the markup left without any (a hyperlink's runs included). */
function applyRunProperties(elements: Element[], rPr: wml.RPr | undefined): void {
  if (!rPr) return;
  const visit = (items: Element[] | undefined): void => {
    if (!items) return;
    for (const item of items) {
      const name = typeNameOf(item);
      const inner = innerOf(item);
      if (name === 'org_docx4j_wml.R') {
        const run = inner as { rPr?: wml.RPr };
        if (!run.rPr) run.rPr = deepCopy(rPr) as wml.RPr;
      } else if (name !== undefined) {
        visit(childrenOf(inner));
      }
    }
  };
  visit(elements);
}

function inlineContentOf(paragraph: Element): Element[] {
  const content = childrenOf(innerOf(paragraph)) ?? [];
  return content.filter((item) => { const n = typeNameOf(item); return n === 'org_docx4j_wml.R' || n === 'org_docx4j_wml.P.Hyperlink'; });
}

/**
 * Binds one XHTML control. Returns what was done, or a note saying why not: no entry, no part, no
 * node, no parser in this runtime. The content the markup gave that the control cannot hold is in
 * `dropped` (a run-level control's blocks after the first, and what the converter dropped).
 */
export async function applyXhtmlBindingTo(control: ContentControl, parts: CustomXmlPartLookup, host: XhtmlHost, options: XhtmlBindingOptions = {}): Promise<{ done: boolean; note?: string; dropped?: Record<string, number> }> {
  const params = tagParamsOf(control.tag);
  const entryId = params.get('od:xpath')!;
  const entry = xpathsEntriesOf(parts).get(entryId);
  if (!entry) return { done: false, note: `od:xpath=${entryId}: no such entry in the XPaths part` };
  const part = parts.getItem(entry.storeItemID) ?? parts.items.find((p) => p.id.replace(/[{}]/g, '').toLowerCase() === entry.storeItemID.replace(/[{}]/g, '').toLowerCase());
  if (!part) return { done: false, note: `od:xpath=${entryId}: no custom XML part ${entry.storeItemID}` };
  const node = part.selectSingleNode(entry.xpath, entry.prefixMappings);
  if (!node) return { done: false, note: `od:xpath=${entryId}: ${entry.xpath} selects nothing` };
  const parser = options.parser ?? (typeof DOMParser === 'function' ? (h: string) => new DOMParser().parseFromString(h, 'text/html') : undefined);
  if (!parser) return { done: false, note: `od:xpath=${entryId}: XHTML needs an HTML parser (applyBindings({ html: { parser } }))` };
  const converted: PasteResult = convertHtml(node.text, await lookupFor(host), { parser });
  const dropped: Record<string, number> = { ...converted.report.dropped };
  const { blocks } = converted;
  // What the document must give first: the styles the markup names, a definition per list, a relationship per link.
  if (converted.report.stylesWanted.length > 0) {
    try { await host.styles.ensure(converted.report.stylesWanted); } catch { dropped['styles the defaults lack'] = (dropped['styles the defaults lack'] ?? 0) + converted.report.stylesWanted.length; }
  }
  const needs = needsOf(blocks);
  const lists = new Map<string, string>();
  let firstNumbered: string | undefined;
  for (const { key, bullet } of needs.lists) {
    if (bullet) lists.set(key, await host.numbering.newList({ bullet: true }));
    else if (firstNumbered === undefined) { firstNumbered = await host.numbering.newList({ bullet: false }); lists.set(key, firstNumbered); }
    else lists.set(key, host.numbering.restart(firstNumbered));
  }
  const links = new Map<string, string>();
  const rels = needs.links.length > 0 ? control.parentBody.part?.getRelationshipsPart(true) : undefined;
  if (rels) {
    for (const address of needs.links) links.set(address, rels.addExternalRelationship(Namespaces.HYPERLINK, address).id);
    host.requireStyles([HYPERLINK_STYLE]);
  } else if (needs.links.length > 0) dropped['links'] = (dropped['links'] ?? 0) + needs.links.length;
  const elements = await contentOf(blocksToXml(blocks, { numId: (key) => lists.get(key), relId: (address) => links.get(address) }));
  applyRunProperties(elements, control.runProperties);
  const sdt = control.sdt;
  sdt.sdtContent ??= {};
  const content = ((sdt.sdtContent as { content?: Element[] }).content ??= []);
  if (control.form === 'Run') {
    // A run-level control holds inline content: the first paragraph's, the rest reported.
    const first = elements.find((e) => typeNameOf(e) === 'org_docx4j_wml.P');
    const inlines = first ? inlineContentOf(first) : [];
    const rest = elements.length - (first ? 1 : 0);
    if (rest > 0) dropped['blocks after the first, in a run-level control'] = (dropped['blocks after the first, in a run-level control'] ?? 0) + rest;
    content.length = 0;
    content.push(...inlines);
    linkParents(content, sdt.sdtContent as object);
  } else {
    content.length = 0;
    content.push(...elements);
    linkParents(content, sdt.sdtContent as object);
  }
  control.isShowingPlaceholder = false;
  return Object.keys(dropped).length > 0 ? { done: true, dropped } : { done: true };
}

/** The XHTML pass over the controls, after the text bindings: counted into the result, what was left said in its notes. */
export async function applyXhtmlBindingsTo(controls: ContentControl[], parts: CustomXmlPartLookup, host: XhtmlHost, result: BindingResult, options: XhtmlBindingOptions = {}): Promise<void> {
  for (const control of controls) {
    if (!isXhtmlBound(control)) continue;
    result.bound++;
    const outcome = await applyXhtmlBindingTo(control, parts, host, options);
    if (outcome.done) result.updated++; else result.skipped++;
    if (outcome.note) (result.notes ??= []).push(outcome.note);
    if (outcome.dropped) for (const [kind, n] of Object.entries(outcome.dropped)) (result.notes ??= []).push(`${control.tag}: dropped ${kind} (${n})`);
  }
}
