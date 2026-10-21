// The text boxes a paragraph anchors, as the text model counts them (CR-012 D4, REQ-518 and REQ-519):
// what a `textbox:N` address segment after a paragraph's address names, in the content API as in the block
// model (`body/3/textbox:0/1` is the second block of the first text box body/3 anchors).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import { segmentsOf, selectedBranchOf, isElement, ALTERNATE_CONTENT, type HiddenRuns } from './tree.mjs';

/** A text box a paragraph anchors: its `w:txbxContent`, or legacy WordArt's string, which has no paragraphs. */
export interface AnchoredBox {
  content?: wml.CTTxbxContent;
  wordArt?: string;
}

/**
 * The text boxes a paragraph anchors, in the order the text model numbers them (`textbox:0`, `textbox:1`):
 * the drawings, VML pictures and kept `mc:AlternateContent` its runs hold in the accepted view, a hidden
 * run's and a field code's left out, each walked as the text model's `anchored` walks it. In a drawing or a
 * shape every `w:txbxContent` (the selected branch of an element, not entered further), and legacy WordArt
 * (a `v:shape` with a `v:textpath` string); nothing in a chart, SmartArt, a picture or an embedded object.
 *
 * A paragraph the accepted view joins to the next (its mark deleted) is one block in the text model, whose
 * boxes are numbered on from the first paragraph's; here each paragraph numbers its own.
 */
export function textBoxesOf(p: wml.P, hidden?: HiddenRuns): AnchoredBox[] {
  const out: AnchoredBox[] = [];
  // the fields open where the walk is, each true while in its code (from its begin to its separate)
  const fields: boolean[] = [];
  segmentsOf(p, {
    hidden,
    visit: (item, _at, _owner, run) => {
      if (run === undefined) return;
      const name = item.name.localPart;
      if (name === 'fldChar') {
        const type = (item.value as wml.FldChar).fldCharType;
        if (type === 'begin') fields.push(true);
        else if (type === 'separate') { if (fields.length > 0) fields[fields.length - 1] = false; }
        else if (type === 'end') fields.pop();
        return;
      }
      const drawn = name === 'drawing' || name === 'pict' || name === 'object' || (item.value as { TYPE_NAME?: string }).TYPE_NAME === ALTERNATE_CONTENT;
      if (!drawn || fields.includes(true) || hidden?.(run.element.value) === true) return;
      boxesIn(item.value, out, 0);
    },
  });
  return out;
}

const CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CHARTEX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
/** What a graphic holds that the text model doesn't enter for text boxes: a chart, SmartArt, a picture. */
const LEAVES = new Set(['org_docx4j_dml_chart.CTRelId', 'org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex.CTRelId', 'org_docx4j_dml_diagram.CTRelIds', 'org_docx4j_dml_picture.Pic']);

/** The text model's `anchored` walk, text boxes only: typed objects in property and list order, DOM not entered. */
function boxesIn(value: unknown, out: AnchoredBox[], depth: number): void {
  if (typeof value !== 'object' || value === null || depth > 200) return;
  const tn = (value as { TYPE_NAME?: string }).TYPE_NAME;
  if (tn === undefined) return;
  switch (tn) {
    case 'org_docx4j_wml.CTTxbxContent':
      out.push({ content: value as wml.CTTxbxContent });
      return;
    case ALTERNATE_CONTENT:
      for (const o of selectedBranchOf(value as mce.AlternateContent) ?? []) boxesIn(unwrap(o), out, depth + 1);
      return;
    case 'org_docx4j_wml.CTObject':
    case 'org_docx4j_wml.CTControl':
    case 'org_docx4j_dml_picture.Pic':
      return;
    case 'org_docx4j_dml_wordprocessingDrawing.Inline':
    case 'org_docx4j_dml_wordprocessingDrawing.Anchor': {
      const data = (value as { graphic?: { graphicData?: { any?: unknown[]; uri?: string } } }).graphic?.graphicData;
      if (!data) return;
      const uri = data.uri ?? '';
      if (uri === CHART || uri.startsWith(CHARTEX)) return;
      if ((data.any ?? []).some((o) => LEAVES.has((unwrap(o) as { TYPE_NAME?: string } | undefined)?.TYPE_NAME ?? ''))) return;
      break;
    }
    case 'org_docx4j_vml.CTShape': {
      for (const e of ((value as { pathOrFormulasOrHandles?: unknown[] }).pathOrFormulasOrHandles ?? [])) {
        const v = unwrap(e) as { TYPE_NAME?: string; string?: string } | undefined;
        if (v?.TYPE_NAME === 'org_docx4j_vml.CTTextPath' && v.string !== undefined && v.string !== '') {
          out.push({ wordArt: v.string });
          return;
        }
      }
      break;
    }
    default:
  }
  for (const key of Object.keys(value)) {
    if (key === 'PARENT' || key === 'TYPE_NAME') continue;
    const fv = (value as Record<string, unknown>)[key];
    if (Array.isArray(fv)) { for (const o of [...fv]) if (typeof o === 'object' && o !== null) boxesIn(unwrap(o), out, depth + 1); }
    else if (typeof fv === 'object' && fv !== null) boxesIn(unwrap(fv), out, depth + 1);
  }
}

function unwrap(o: unknown): unknown {
  return isElement(o) ? o.value : o;
}
