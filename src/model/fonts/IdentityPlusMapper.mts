// docx4j `org.docx4j.fonts.IdentityPlusMapper`: the default mapper.
//
// Maps a document font to the physical font of the same name where the environment has one -
// an identity mapping - plus the shared passes in {@link Mapper} (the metric clones,
// `w:altName`, a face of the same class, Word's own default) for the fonts it lacks.  Measured
// on three real-document corpora in six font environments, including boxes with no Microsoft
// fonts and images with no fonts at all, it is ahead of docx4j's `BestMatchingMapper` in every
// one of them (docx4j CR-016 phase 0c).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { Mapper } from './Mapper.mjs';
import type { PhysicalFont } from './PhysicalFont.mjs';

/** The order a family whose plain face is absent is asked for.  Bold before italic since
 *  docx4j 17.1.1: with italic first, "Franklin Gothic Demi" was set upright in Franklin Gothic
 *  Demi Italic, and a bold upright is nearer a missing regular than an italic is. */
const VARIANTS = [' regular', ' bold', ' italic', ' bold italic'];

export class IdentityPlusMapper extends Mapper {

  /** A variant of the name: "Noto Sans Symbols" is registered as "Noto Sans Symbols Regular". */
  protected override resolveDocumentFont(documentFontName: string,
    _fontTableEntry: wml.Fonts.Font | undefined): PhysicalFont | undefined {
    this.resolvedVia = undefined;
    for (const variant of VARIANTS) {
      const pf = this.registry.get(documentFontName + variant);
      if (pf !== undefined) {
        this.resolvedVia = `a variant of the name: ${documentFontName}${variant}`;
        return pf;
      }
    }
    return undefined;
  }
}
