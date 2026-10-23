// Word's rule for what a table cell holds (CR-002 section 55): docx4j's org.docx4j.model.table.TableRules
// (18a3a31f4, VERSION_17_3_4), its names and rule, with what an edit here does to keep it.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import { type Element, typeNameOf, childrenOf, linkParents, selectedBranchOf, paragraphOf, rowsOf, cellsOf, ALTERNATE_CONTENT, BRANCH_TYPES } from './tree.mjs';

const P = 'org_docx4j_wml.P';
const TBL = 'org_docx4j_wml.Tbl';
const TC = 'org_docx4j_wml.Tc';

/** What counts as a block among a cell's contents; anything else (bookmarks, comment and move ranges, permissions) is range markup. */
const BLOCKS: ReadonlySet<string> = new Set([P, TBL, 'org_docx4j_wml.SdtBlock', 'org_docx4j_wml.CTCustomXmlBlock', ALTERNATE_CONTENT]);

/** What stands between a cell and the blocks Word reads through: block-level controls, their content, custom XML, a kept branch. */
const WRAPPERS: ReadonlySet<string> = new Set([
  'org_docx4j_wml.SdtBlock', 'org_docx4j_wml.SdtContentBlock', 'org_docx4j_wml.CTSdtContentBlock', 'org_docx4j_wml.CTCustomXmlBlock',
  ALTERNATE_CONTENT, ...BRANCH_TYPES,
]);

/** The paragraph a cell ends with, as Word reads it, and the list it is in. */
export interface CellEnd { element: Element<wml.P>; container: Element[] }

/** The last block-level item, unwrapped from its element pair; undefined where there is none. */
function lastBlock(items: readonly Element[] | undefined): Element | undefined {
  for (let i = (items?.length ?? 0) - 1; i >= 0; i--) if (BLOCKS.has(typeNameOf(items![i]!) ?? '')) return items![i];
  return undefined;
}

/**
 * The paragraph these cell contents end with, as Word reads them: the last block, past range markup, through
 * block-level controls and custom XML, and through the branch of a kept `mc:AlternateContent` a reader takes
 * (`selectedBranchOf`; the others aren't checked). Undefined where the last block, so read, is a table or there is
 * no block at all.
 */
export function paragraphEndOf(items: Element[] | undefined): CellEnd | undefined {
  const last = lastBlock(items);
  if (!last) return undefined;
  const tn = typeNameOf(last);
  if (tn === P) return { element: last as Element<wml.P>, container: items! };
  if (tn === TBL) return undefined;
  const inner = tn === ALTERNATE_CONTENT ? selectedBranchOf(last.value as mce.AlternateContent) : childrenOf(last.value as object);
  return paragraphEndOf(inner);
}

/** docx4j's `endsWithParagraph`: whether the last block among these cell contents, as Word reads it, is a paragraph. */
export function endsWithParagraph(items: Element[] | undefined): boolean {
  return paragraphEndOf(items) !== undefined;
}

/**
 * docx4j's `endWithParagraph`, Word's repair: an empty paragraph appended to the cell, after any wrapper, where its
 * last block as Word reads it isn't one. Returns the paragraph appended, undefined where the rule held.
 */
export function endWithParagraph(tc: wml.Tc): wml.P | undefined {
  const content = (tc.content ??= []) as Element[];
  if (endsWithParagraph(content)) return undefined;
  const p = paragraphOf([]);
  content.push(p as Element);
  linkParents(p, tc);
  return p.value;
}

/** The cell whose content holds this container (a `w:tc`, a control's content, a branch, ...): up `PARENT` past the wrappers. */
export function wrappingCellOf(owner: object | undefined): wml.Tc | undefined {
  let o = owner as { TYPE_NAME?: string; PARENT?: object } | undefined;
  for (let guard = 0; o && guard < 32 && WRAPPERS.has(o.TYPE_NAME ?? ''); guard++) o = o.PARENT as typeof o;
  return o?.TYPE_NAME === TC ? o as wml.Tc : undefined;
}

/** The cell a paragraph is in, through the wrappers, and whether it is the paragraph that cell ends with as Word reads it. */
export function endsCell(p: wml.P): boolean {
  const cell = wrappingCellOf((p as { PARENT?: object }).PARENT);
  return cell !== undefined && paragraphEndOf(cell.content as Element[] | undefined)?.element.value === p;
}

/**
 * D3: every cell of the tables in these blocks, at any depth, ends with a paragraph (docx4j's repair). For content
 * that comes from a caller (`insertXml`, `insertOoxml`) before it goes in.
 */
export function endCellsWithParagraphs(blocks: readonly Element[]): void {
  const visit = (items: readonly Element[] | undefined): void => {
    for (const el of items ?? []) {
      const v = el.value;
      if (typeof v !== 'object' || v === null) continue;
      if (typeNameOf(el) === TBL) {
        for (const row of rowsOf(v as wml.Tbl, 'every')) {
          for (const cell of cellsOf(row.element.value, 'every')) {
            const tc = cell.element.value as wml.Tc;
            visit(tc.content as Element[] | undefined);
            endWithParagraph(tc);
          }
        }
      } else if (typeNameOf(el) === ALTERNATE_CONTENT) {
        for (const choice of (v as mce.AlternateContent).choice ?? []) visit(choice.any as Element[] | undefined);
        visit((v as mce.AlternateContent).fallback?.any as Element[] | undefined);
      } else if (typeNameOf(el) !== P) visit(childrenOf(v));
    }
  };
  visit(blocks);
}
