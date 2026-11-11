// SPDX-License-Identifier: Apache-2.0
//
// The bind step for Word fragments (CR-005 section 11; the specification's "Word Fragments", REQ-059 and REQ-060; the editor's
// ED-005 proposal 46, part 1). A control whose tag says `od:progid=Word.Document` and names an `od:xpath` entry, carrying no
// `w:dataBinding`, is bound to a node holding a Word document as escaped Flat OPC: the control's content is replaced with the
// document's, resolved here as Word 365 resolves docx4j's altChunk (`route: 'resolve'`, the default: `importFlatOpc`'s styles,
// lists, notes and ids), or with that altChunk itself (`route: 'altChunk'`, docx4j's `BindingTraverserXSLT.convertFlatOPC`,
// which Word resolves when it opens the file: 46a). A block-level control's content is replaced; a cell-level one's cell; a
// run-level one is left, with a note (REQ-060, REC-006). Controls inside the fragment keep their tags and aren't bound.
//
// Word's own form is taken too (`w15FlatOpc: 'resolve'`, the default): a rich text control mapped by `w15:dataBinding` whose
// node holds Flat OPC, which Word 365 fills from the node when it opens the file and reconciles the same way (46c). The text
// pass leaves such a control (`holdsFlatOpc`); here it is resolved, and its mapping kept (`w15Mapping: 'keep'`, so Word refills
// it and writes an edit back) or dropped (`'drop'`: the `w15:dataBinding` alone goes, the control stays).
//
// The reverse direction, a control's content written back to its node as Flat OPC, is out of scope (decision (c)).
import { AlternativeFormatInputPart } from '../../parts/BinaryPart.mjs';
import { sdtProperty, W15_NAMESPACE } from '@docx4j/generated-objects-ts/builders/wml';
import type { Element } from '../content/tree.mjs';
import { typeNameOf, linkParents, writingEveryBranch } from '../content/tree.mjs';
import type { ContentControl } from '../content/ContentControl.mjs';
import { contentOf } from '../content/ooxml.mjs';
import { importFlatOpc, freePartName, partNamesOf, type FragmentTarget, type FragmentCreated } from '../content/fragments.mjs';
import { wrappingCellOf, endWithParagraph } from '../content/tableRules.mjs';
import { assignParaIds } from '../content/paraIds.mjs';
import type { CustomXmlPartLookup } from './XmlMapping.mjs';
import { tagParamsOf, selectEntryNode } from './opendope.mjs';
import { holdsFlatOpc, type BindingResult } from './bindings.mjs';

/** The tag's `od:progid` value for a Word fragment (the specification's "Word Fragments"). */
export const WORD_FRAGMENT_PROGID = 'Word.Document';

/** What the fragment pass needs of the package: `importFlatOpc`'s target. */
export type FragmentHost = FragmentTarget;

export interface FragmentBindingOptions {
  /** `'resolve'` (the default): the fragment's content, reconciled as Word reconciles it; `'altChunk'`: docx4j's altChunk, for Word to resolve. */
  route?: 'resolve' | 'altChunk';
  /** Word's `w:matchSrc`: the fragment's styles as direct formatting (on the altChunk route, the flag on the chunk). */
  keepSourceFormatting?: boolean;
  /** A rich text control mapped by `w15:dataBinding` to a node holding Flat OPC: `'resolve'` (the default) or `'leave'`. */
  w15FlatOpc?: 'resolve' | 'leave';
  /** Such a control's mapping once resolved: `'keep'` (the default: live in Word 2013 and later) or `'drop'` (one-way). */
  w15Mapping?: 'keep' | 'drop';
}

/**
 * What one control's binding did: done, or left with a note; what the fragment held that the document doesn't take, counted;
 * and what was added to the document (`importFlatOpc`'s `created`; on the altChunk route, the chunk's part alone).
 */
export interface FragmentOutcome {
  done: boolean;
  note?: string;
  dropped?: Record<string, number>;
  created?: FragmentCreated;
}

/** Whether a control is bound to a Word fragment through its tag: `od:progid=Word.Document` and an `od:xpath` entry, no `w:dataBinding`. */
export function isFragmentBound(control: ContentControl): boolean {
  const params = tagParamsOf(control.tag);
  return params.get('od:progid') === WORD_FRAGMENT_PROGID && params.has('od:xpath') && control.xmlMapping.dataBinding === undefined;
}

/** Whether a control is Word's own mapped rich text over Flat OPC: rich text, a `w15:dataBinding`, a node holding a `pkg:package`. */
export function isW15FlatOpcMapped(control: ContentControl): boolean {
  if (control.type !== 'RichText') return false;
  if (sdtProperty(control.sdt.sdtPr, 'dataBinding', W15_NAMESPACE) === undefined) return false;
  let text: string | undefined;
  try {
    text = control.xmlMapping.customXmlNode?.text;
  } catch {
    text = undefined;
  }
  return text !== undefined && holdsFlatOpc(text);
}

function nameOf(control: ContentControl): string {
  return control.tag || control.title || `control ${control.id}`;
}

/** The content into the control: a block-level control's content, a cell-level one's cell; undefined, or a note why not. */
function place(control: ContentControl, elements: Element[]): string | undefined {
  const sdt = control.sdt;
  if (control.form === 'Run') return `${nameOf(control)}: a Word fragment in a run-level control is left (REQ-060: block-level or in a cell)`;
  if (control.form === 'Row') return `${nameOf(control)}: a Word fragment in a row-level control is left (REQ-060: block-level or in a cell)`;
  sdt.sdtContent ??= {};
  const content = ((sdt.sdtContent as { content?: Element[] }).content ??= []);
  if (control.form === 'Cell') {
    // docx4j's applyImportedContent: the cell's content, the cell kept
    const cell = content.find((e) => typeNameOf(e) === 'org_docx4j_wml.Tc');
    const tc = cell ? ((cell as { value?: object }).value ?? cell) as { content?: Element[] } : undefined;
    if (!tc) return `${nameOf(control)}: a cell-level control holding no cell is left`;
    tc.content = [...elements];
    linkParents(tc.content, tc);
    endWithParagraph(tc as never);
    assignParaIds(tc, elements);
  } else {
    content.length = 0;
    content.push(...elements);
    linkParents(content, sdt.sdtContent as object);
    assignParaIds(sdt.sdtContent as object, elements);
    // in a block control in a cell, the cell ends with a paragraph after the fragment's table (CR-002 section 55)
    const cell = wrappingCellOf((sdt as { PARENT?: object }).PARENT);
    if (cell) endWithParagraph(cell);
  }
  writingEveryBranch(() => { control.isShowingPlaceholder = false; });
  return undefined;
}

/** docx4j's form: the Flat OPC as an `application/xml` part beside the control's, an `aFChunk` relationship, a `w:altChunk`. */
async function altChunkOf(control: ContentControl, flatOpc: string, keepSourceFormatting: boolean): Promise<Element[]> {
  const owner = control.parentBody.part;
  const pkg = owner?.package;
  if (!owner || !pkg) throw new Error('the control has no part');
  // Word can't open a chunk without its XML declaration (docx4j's comment)
  const text = flatOpc.trimStart().startsWith('<?xml') ? flatOpc : `<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n${flatOpc}`;
  const part = new AlternativeFormatInputPart(freePartName(pkg, owner, 'chunk', '.xml'), 'application/xml');
  part.setBytes(new TextEncoder().encode(text));
  const rel = owner.addTargetPart(part);
  return contentOf(`<w:altChunk r:id="${rel.id}">${keepSourceFormatting ? '<w:altChunkPr><w:matchSrc/></w:altChunkPr>' : ''}</w:altChunk>`);
}

/** One fragment into one control, by the route: done, or left with a note; what was dropped, counted. */
async function bindFragment(control: ContentControl, flatOpc: string, host: FragmentHost, options: FragmentBindingOptions): Promise<FragmentOutcome> {
  if (control.form === 'Run' || control.form === 'Row') return { done: false, note: place(control, [])! };
  const target = control.parentBody.part;
  if (!target) return { done: false, note: `${nameOf(control)}: the control has no part, so the fragment's parts have nowhere to go` };
  try {
    if (options.route === 'altChunk') {
      const before = partNamesOf(target.package);
      const elements = await altChunkOf(control, flatOpc, options.keepSourceFormatting === true);
      const created: FragmentCreated = { numIds: [], abstractNumIds: [], styleIds: [], parts: [...partNamesOf(target.package)].filter((n) => !before.has(n)) };
      const note = place(control, elements);
      return note ? { done: false, note, created } : { done: true, created };
    }
    const { content, dropped, created } = await importFlatOpc(flatOpc, host, { target, keepSourceFormatting: options.keepSourceFormatting === true });
    const note = place(control, content);
    if (note) return { done: false, note, created };
    return { done: true, ...(Object.keys(dropped).length > 0 ? { dropped } : {}), created };
  } catch (e) {
    return { done: false, note: `${nameOf(control)}: the Word fragment can't be read (${(e as Error).message})` };
  }
}

/**
 * Binds one control to its Word fragment: an `od:progid=Word.Document` control through its `od:xpath` entry, or Word's own
 * `w15`-mapped rich text through its mapping. Returns what was done, or a note saying why not (no entry, no part, no node, a
 * node holding no Flat OPC, a run-level control, a package that can't be read); what the fragment held that the document
 * doesn't take is in `dropped`.
 */
export async function applyFlatOpcBindingTo(control: ContentControl, parts: CustomXmlPartLookup, host: FragmentHost, options: FragmentBindingOptions = {}): Promise<FragmentOutcome> {
  let flatOpc: string;
  let repair: string | undefined;
  if (isFragmentBound(control)) {
    const entryId = tagParamsOf(control.tag).get('od:xpath')!;
    const selected = selectEntryNode(entryId, parts);
    if ('note' in selected) return { done: false, note: selected.note };
    flatOpc = selected.node.text;
    repair = selected.repair;
    if (!holdsFlatOpc(flatOpc)) return { done: false, note: `od:xpath=${entryId}: the node holds no Flat OPC package (pkg:package)` };
  } else if (isW15FlatOpcMapped(control)) {
    flatOpc = control.xmlMapping.customXmlNode!.text;
  } else {
    return { done: false, note: `${nameOf(control)}: neither an od:progid=Word.Document control nor a w15-mapped rich text control over Flat OPC` };
  }
  const outcome = await bindFragment(control, flatOpc, host, options);
  if (outcome.done && !isFragmentBound(control) && options.w15Mapping === 'drop') control.removeProperty('dataBinding', W15_NAMESPACE);
  return repair ? { ...outcome, note: outcome.note ? `${repair}; ${outcome.note}` : repair } : outcome;
}

/**
 * The fragment pass over the controls, after the text, picture and XHTML passes: counted into the result, what was left or
 * dropped said in its notes. A `w15`-mapped control the text pass counted and skipped is counted updated once resolved.
 */
export async function applyFlatOpcBindingsTo(controls: ContentControl[], parts: CustomXmlPartLookup, host: FragmentHost, result: BindingResult, options: FragmentBindingOptions = {}): Promise<void> {
  for (const control of controls) {
    const tagged = isFragmentBound(control);
    const mapped = !tagged && options.w15FlatOpc !== 'leave' && isW15FlatOpcMapped(control);
    if (!tagged && !mapped) continue;
    if (tagged) result.bound++;
    const outcome = await applyFlatOpcBindingTo(control, parts, host, options);
    if (outcome.done) {
      result.updated++;
      // the text pass counted the mapped control and left it (holdsFlatOpc)
      if (mapped && result.skipped > 0) result.skipped--;
    } else if (tagged) result.skipped++;
    if (outcome.note) (result.notes ??= []).push(outcome.note);
    if (outcome.dropped) for (const [kind, n] of Object.entries(outcome.dropped)) (result.notes ??= []).push(`${nameOf(control)}: dropped ${kind} (${n})`);
  }
}
