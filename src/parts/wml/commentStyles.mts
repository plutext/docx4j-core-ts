// Word's two comment styles, `CommentText` (the comment's paragraphs) and `CommentReference` (the
// reference mark), as one definition for the comment parts, which add them when they are created,
// and for `pkg.styles.ensure`, which a caller that places a comment itself uses (CR-002 section 37).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as wmlFactory from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { COMMENT_REFERENCE_STYLE, COMMENT_TEXT_STYLE } from '../../model/content/comments.mjs';

/** Fresh copies of the two definitions, each call. */
export function commentStyles(): wml.Style[] {
  return [
    wmlFactory.createStyle({
      type: 'paragraph', styleId: COMMENT_TEXT_STYLE,
      name: wmlFactory.createStyleName({ val: 'annotation text' }),
      basedOn: wmlFactory.createStyleBasedOn({ val: 'Normal' }),
      uiPriority: wmlFactory.createStyleUiPriority({ val: 99 }),
      semiHidden: wmlFactory.createBooleanDefaultTrue(),
      unhideWhenUsed: wmlFactory.createBooleanDefaultTrue(),
      rPr: wmlFactory.createRPr({ sz: wmlFactory.createHpsMeasure({ val: 20 }), szCs: wmlFactory.createHpsMeasure({ val: 20 }) }),
    }),
    wmlFactory.createStyle({
      type: 'character', styleId: COMMENT_REFERENCE_STYLE,
      name: wmlFactory.createStyleName({ val: 'annotation reference' }),
      basedOn: wmlFactory.createStyleBasedOn({ val: 'DefaultParagraphFont' }),
      uiPriority: wmlFactory.createStyleUiPriority({ val: 99 }),
      semiHidden: wmlFactory.createBooleanDefaultTrue(),
      unhideWhenUsed: wmlFactory.createBooleanDefaultTrue(),
      rPr: wmlFactory.createRPr({ sz: wmlFactory.createHpsMeasure({ val: 16 }), szCs: wmlFactory.createHpsMeasure({ val: 16 }) }),
    }),
  ];
}
