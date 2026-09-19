// The placeholder shapes a created slide and its layout carry (CR-001 section 17).
//
// PowerPoint shows "Click to add title" only where the *slide* holds a placeholder shape
// (`p:sp` whose `p:nvSpPr/p:nvPr/p:ph` names a type or an index) matching one the layout
// defines; a slide whose `p:spTree` is empty - which is all docx4j's `COMMON_SLIDE_DATA` is -
// opens blank, with nothing to click and nothing to type into. So `createPackage()` gives the
// layout a title and a body placeholder with their geometry, and every slide the matching pair,
// empty (the prompts) or filled from `{ title, body }`.
//
// Built with the generated factories rather than markup, so the shapes are the object model's
// and a caller can go on editing them. Geometry follows PowerPoint's own Title and Content
// layout, as fractions of the slide size, so that it holds at any `slideSize`.
import type * as pml from '@docx4j/generated-objects-ts/modules/org_pptx4j_pml';
import type * as dml from '@docx4j/generated-objects-ts/modules/org_docx4j_dml';
import {
  createShape, createShapeNvSpPr, createNvPr, createCTPlaceholder,
} from '@docx4j/generated-objects-ts/factory/org_pptx4j_pml';
import {
  createCTNonVisualDrawingProps, createCTNonVisualDrawingShapeProps, createCTShapeLocking,
  createCTShapeProperties, createCTTextBody, createCTTextBodyProperties, createCTTextListStyle,
  createCTTextParagraph, createCTRegularTextRun, createCTTextCharacterProperties,
  createCTTransform2D, createCTPoint2D, createCTPositiveSize2D,
} from '@docx4j/generated-objects-ts/factory/org_docx4j_dml';

/** The text of a created slide's placeholders; `body` is one paragraph per line, or per entry. */
export interface SlideTextOptions {
  /** The title placeholder's text; the placeholder is left empty (its prompt) when absent. */
  title?: string;
  /** The body placeholder's text: a string (split on newlines) or one string per paragraph. */
  body?: string | string[];
}

/** A rectangle in EMU, as `a:xfrm` carries it. */
export interface Rect { x: number; y: number; cx: number; cy: number }

/**
 * Where the title and body placeholders sit, as fractions of the slide: PowerPoint's own
 * Title and Content layout at 4:3 (title 685800, 457200, 7772400 x 1143000; body 685800,
 * 1600200, 7772400 x 4525963 of 9144000 x 6858000), so any slide size is laid out the same way.
 */
export function placeholderGeometry(sldSz: { cx?: number; cy?: number } | undefined): { title: Rect; body: Rect } {
  const w = sldSz?.cx ?? 9144000;
  const h = sldSz?.cy ?? 6858000;
  const at = (fx: number, fy: number, fcx: number, fcy: number): Rect =>
    ({ x: Math.round(w * fx), y: Math.round(h * fy), cx: Math.round(w * fcx), cy: Math.round(h * fcy) });
  return { title: at(0.075, 0.0667, 0.85, 0.1667), body: at(0.075, 0.2333, 0.85, 0.66) };
}

/** `a:p/a:r/a:t` paragraphs for a placeholder's `p:txBody`; one empty paragraph when there is no text. */
function paragraphsOf(lines: string[]): dml.CTTextParagraph[] {
  if (lines.length === 0) return [createCTTextParagraph({ endParaRPr: createCTTextCharacterProperties({ lang: 'en-US' }) })];
  return lines.map((text, i) => createCTTextParagraph({
    egTextRun: [createCTRegularTextRun({ rPr: createCTTextCharacterProperties({ lang: 'en-US' }), t: text })],
    ...(i === lines.length - 1 ? { endParaRPr: createCTTextCharacterProperties({ lang: 'en-US' }) } : {}),
  }));
}

/** The body text as paragraphs: a string is split on newlines, as Office JS's `insertText` does. */
export function bodyLines(body: string | string[] | undefined): string[] {
  if (body === undefined) return [];
  const lines = Array.isArray(body) ? body.slice() : body.split('\n');
  while (lines.length > 0 && lines[lines.length - 1] === '') lines.pop();
  return lines;
}

/**
 * One placeholder shape (docx4j's `SAMPLE_SHAPE` in `PresentationMLPackage`, and the layout
 * markup of its `SlidePlaceholder` sample): `p:sp` with `a:spLocks noGrp`, the `p:ph`, an
 * optional `a:xfrm` and a `p:txBody` of the lines given.
 */
export function placeholderShape(options: {
  id: number; name: string; type?: pml.STPlaceholderType; idx?: number; xfrm?: Rect; lines?: string[];
}): pml.Shape {
  const ph = createCTPlaceholder({
    ...(options.type === undefined ? {} : { type: options.type }),
    ...(options.idx === undefined ? {} : { idx: options.idx }),
  });
  const spPr = createCTShapeProperties(options.xfrm === undefined ? {} : {
    xfrm: createCTTransform2D({
      off: createCTPoint2D({ x: options.xfrm.x, y: options.xfrm.y }),
      ext: createCTPositiveSize2D({ cx: options.xfrm.cx, cy: options.xfrm.cy }),
    }),
  });
  return createShape({
    nvSpPr: createShapeNvSpPr({
      cNvPr: createCTNonVisualDrawingProps({ id: options.id, name: options.name }),
      cNvSpPr: createCTNonVisualDrawingShapeProps({ spLocks: createCTShapeLocking({ noGrp: true }) }),
      nvPr: createNvPr({ ph }),
    }),
    spPr,
    txBody: createCTTextBody({
      bodyPr: createCTTextBodyProperties(),
      lstStyle: createCTTextListStyle(),
      p: paragraphsOf(options.lines ?? []),
    }),
  });
}

/**
 * The title (`p:ph type="title"`) and body (`p:ph idx="1"`) shapes a created **slide** carries,
 * with no geometry of their own: each inherits the layout's, which is what a placeholder is for.
 */
export function slidePlaceholders(text: SlideTextOptions = {}): pml.Shape[] {
  return [
    placeholderShape({ id: 2, name: 'Title 1', type: 'title', ...(text.title === undefined ? {} : { lines: [text.title] }) }),
    placeholderShape({ id: 3, name: 'Content Placeholder 2', idx: 1, lines: bodyLines(text.body) }),
  ];
}

/**
 * The same pair for the **layout**, with the geometry the slides inherit and PowerPoint's own
 * prompt text. Without these the slide's placeholders would have nothing to inherit from and
 * PowerPoint would draw them at the origin with no size.
 */
export function layoutPlaceholders(sldSz: { cx?: number; cy?: number } | undefined): pml.Shape[] {
  const geometry = placeholderGeometry(sldSz);
  return [
    placeholderShape({ id: 2, name: 'Title 1', type: 'title', xfrm: geometry.title, lines: ['Click to edit Master title style'] }),
    placeholderShape({ id: 3, name: 'Content Placeholder 2', idx: 1, xfrm: geometry.body, lines: ['Click to edit Master text styles'] }),
  ];
}

/** The placeholder shapes of a slide, layout or master, in shape-tree order. */
export function placeholdersOf(cSld: pml.CommonSlideData | undefined): { shape: pml.Shape; ph: pml.CTPlaceholder }[] {
  const out: { shape: pml.Shape; ph: pml.CTPlaceholder }[] = [];
  for (const item of cSld?.spTree?.spOrGrpSpOrGraphicFrame ?? []) {
    const shape = item as pml.Shape;
    const ph = shape?.nvSpPr?.nvPr?.ph;
    if (shape.TYPE_NAME === 'org_pptx4j_pml.Shape' && ph) out.push({ shape, ph });
  }
  return out;
}

/** Date, footer and slide-number placeholders: a layout's, which a new slide does not repeat
 *  (docx4j's `SlidePlaceholder` sample deletes exactly these three when it derives a slide). */
const NOT_ON_A_NEW_SLIDE = new Set(['dt', 'ftr', 'sldNum']);

/**
 * The placeholders a new slide on this layout carries: one shape per placeholder the layout
 * defines (bar the date, footer and slide-number ones), with the layout's `type` and `idx` so
 * that each inherits that layout shape's position and text style, and with the text given in the
 * first title and the first body one. A layout that defines none - docx4j's own - gets the
 * default title and body pair.
 */
export function slidePlaceholdersFor(layout: pml.CommonSlideData | undefined, text: SlideTextOptions = {}): pml.Shape[] {
  const usable = placeholdersOf(layout).filter(({ ph }) => ph.type === undefined || !NOT_ON_A_NEW_SLIDE.has(ph.type));
  if (usable.length === 0) return slidePlaceholders(text);
  let id = 2;
  let titleTaken = false;
  let bodyTaken = false;
  return usable.map(({ shape, ph }) => {
    const isTitle = ph.type === 'title' || ph.type === 'ctrTitle';
    let lines: string[] = [];
    if (isTitle && !titleTaken) { titleTaken = true; lines = text.title === undefined ? [] : [text.title]; }
    else if (!isTitle && !bodyTaken) { bodyTaken = true; lines = bodyLines(text.body); }
    return placeholderShape({
      id: id++,
      name: shape.nvSpPr.cNvPr.name || (isTitle ? 'Title 1' : 'Content Placeholder 2'),
      ...(ph.type === undefined ? {} : { type: ph.type }),
      ...(ph.idx === undefined ? {} : { idx: ph.idx }),
      lines,
    });
  });
}

/** Puts shapes into a slide's or layout's shape tree (docx4j `cSld.getSpTree().getSpOrGrpSpOrGraphicFrame().add(...)`). */
export function addShapes(spTree: pml.GroupShape, shapes: pml.Shape[]): void {
  const into = (spTree.spOrGrpSpOrGraphicFrame ??= []);
  for (const shape of shapes) into.push(shape);
}
