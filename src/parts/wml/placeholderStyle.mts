// Word's `PlaceholderText` character style, which a content control's placeholder run names
// (`PLACEHOLDER_STYLE`) and docx4j's default styles do not carry. `pkg.styles.ensure` splices it,
// and the package ensures it on save once the engine has written a placeholder: Word 2010 drops a
// `w:rStyle` naming a style the styles part lacks (CR-002 section 40, `test/README.md` check 35).
// The definition is current Word's, from `test/fixtures/revisions/check35/35a-word15.docx`.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as wmlFactory from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { PLACEHOLDER_STYLE } from '../../model/customxml/bindings.mjs';

/** A fresh copy of the definition, each call. */
export function placeholderStyle(): wml.Style {
  return wmlFactory.createStyle({
    type: 'character', styleId: PLACEHOLDER_STYLE,
    name: wmlFactory.createStyleName({ val: 'Placeholder Text' }),
    basedOn: wmlFactory.createStyleBasedOn({ val: 'DefaultParagraphFont' }),
    uiPriority: wmlFactory.createStyleUiPriority({ val: 99 }),
    unhideWhenUsed: wmlFactory.createBooleanDefaultTrue(),
    rPr: wmlFactory.createRPr({ color: wmlFactory.createColor({ val: '666666' }) }),
  });
}
