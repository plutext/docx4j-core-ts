// docx4j `org.docx4j.fonts.PhysicalFont`, without the files.
//
// docx4j's class wraps FOP's `EmbedFontInfo` and carries the font's panose and its file URI, so
// that FOP can be told where to read the outlines.  CR-001 Phase B step 4 ports the *decision*
// half of `org.docx4j.fonts` and not the rendering half (CR-001 section 14.2), so a physical
// font here is a name, the family a CSS `font-family` would use, and the no-bold-face flag -
// everything a `Mapper` answer means to a consumer that draws with a browser, measures with
// `fontkit`, or only reports.  Reading font files (`PhysicalFonts.discover`) is a later CR.

/** The suffix of a no-bold-face alias; a registry lookup strips it (docx4j `PhysicalFonts.get`). */
export const NOBOLD_SUFFIX = '+nobold';

/** The suffixes the docx4j FO layer stacks on a font-family for a twin declared from the same
 *  file: the kerned twin, the twin with no OpenType features, and the no-bold-face alias.
 *  Nothing here produces the first two; they are stripped so that a name written by docx4j's
 *  FO output resolves against a registry built here. */
const SUFFIXES = ['+kerned', '+noliga', NOBOLD_SUFFIX];

/**
 * A font available to the consumer: the full name a font registry keys it by (docx4j registers
 * the first of FOP's triplets, the name table's id 4 - "Carlito Regular", "Arimo Bold"), and
 * the family a `font-family` names ("Carlito").
 */
export class PhysicalFont {
  constructor(
    /** The face's full name, as a registry keys it. */
    readonly name: string,
    /** The family name, as CSS would use it; the name with its twin suffixes stripped when the
     *  registry does not say. */
    familyName?: string,
    /** True for a {@link noBoldFaceAlias}. */
    readonly noBoldFace: boolean = false,
  ) {
    this.familyName = familyName ?? stripSuffixes(name);
  }

  readonly familyName: string;

  /**
   * This font again under `name + '+nobold'`, reporting no bold face of its own, so that a
   * consumer synthesises bold from this file at its own advances rather than taking the
   * family's real bold - which is what Word does for a family that has none (docx4j CR-016
   * probe `fonts-light-bold`: Word draws a `w:b` run in Calibri Light with the Calibri Light
   * font object, 378.55pt for a sentence its regular sets in 377.50, where Carlito Bold set it
   * in 388.70).  Held by the mapper, never put in a registry - it stands for one document's font.
   */
  noBoldFaceAlias(): PhysicalFont {
    return new PhysicalFont(this.name + NOBOLD_SUFFIX, this.familyName, true);
  }

  toString(): string {
    return this.name;
  }
}

/** The physical font name without the suffixes a FO layer adds; they stack, so all are stripped. */
export function stripSuffixes(name: string): string {
  let key = name;
  for (let stripped = true; stripped;) {
    stripped = false;
    for (const suffix of SUFFIXES) {
      if (key.endsWith(suffix)) {
        key = key.slice(0, -suffix.length);
        stripped = true;
      }
    }
  }
  return key;
}
