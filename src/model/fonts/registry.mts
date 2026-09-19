// The fonts a consumer has: docx4j's `org.docx4j.fonts.PhysicalFonts`, as an interface the
// caller supplies rather than a static map this package fills.
//
// docx4j discovers the machine's fonts once, in the `Mapper` static initialiser, by walking the
// font directories and the `fonts/` entries of its own jars.  Nothing of that is portable: a
// browser cannot enumerate fonts, a Word add-in has no file system, and Node needs `fontkit` to
// read a name table.  So a `FontRegistry` is a parameter here (CR-001 section 6.3: "the
// interface is fixed here so consumers can plug their own"), and `PhysicalFonts.discover` over
// installed fonts is a later CR.
//
// The default registry is the environment the parity goldens were made in (CR-001 section 14.6:
// `docx4j.fonts.discoverPhysicalFonts.enabled=false` with the symbol, croscore, crosextra and
// theme2023 jars on the classpath), so a mapping computed here reproduces the one in a golden
// on any machine.
import { PhysicalFont, stripSuffixes } from './PhysicalFont.mjs';

/**
 * The fonts available, by name.  The only thing the mapping passes ask of a font environment.
 *
 * Lookups are case-insensitive and strip a twin suffix, as docx4j's `PhysicalFonts.get` does.
 */
export interface FontRegistry {
  /** The font of this name, or undefined where the environment has none. */
  get(name: string | undefined | null): PhysicalFont | undefined;
  /** Every font, for the passes that search by a short key (`FontFallback`'s candidate lists). */
  all(): readonly PhysicalFont[];
}

/** A `FontRegistry` over a fixed list of faces. */
export class SimpleFontRegistry implements FontRegistry {
  private readonly byKey = new Map<string, PhysicalFont>();
  private readonly fonts: PhysicalFont[] = [];

  constructor(fonts: Iterable<PhysicalFont> = []) {
    for (const font of fonts) this.add(font);
  }

  add(font: PhysicalFont): this {
    const key = stripSuffixes(font.name).toLowerCase();
    if (!this.byKey.has(key)) this.fonts.push(font);
    this.byKey.set(key, font);
    return this;
  }

  get(name: string | undefined | null): PhysicalFont | undefined {
    if (name === undefined || name === null) return undefined; // a slot nothing names
    return this.byKey.get(stripSuffixes(name).toLowerCase());
  }

  all(): readonly PhysicalFont[] {
    return this.fonts;
  }
}

/** A registry with nothing in it: every document font is `UNMAPPED`. */
export const EMPTY_FONT_REGISTRY: FontRegistry = new SimpleFontRegistry();

/**
 * `[name, familyName]` of every face docx4j's four font jars carry, read off docx4j
 * `VERSION_17_1_1` (`PhysicalFonts.getPhysicalFonts()` with system discovery off and
 * `docx4j-export-fo-fonts-symbol`, `-croscore`, `-crosextra` and `-theme2023` on the
 * classpath): Liberation's croscore metric clones, crosextra's Carlito and Caladea, the
 * theme2023 stand-ins for Word 365's Aptos, and the symbol jar's DejaVu Serif and Noto
 * Sans Symbols.  This is the environment CR-016 phase 0c measured and the parity goldens
 * record, so the default mapping here is reproducible against them.
 */
const JAR_FONTS: readonly (readonly [string, string])[] = [
  ['Akasia Black', 'Akasia'],
  ['Akasia Black Italic', 'Akasia Black'],
  ['Akasia Bold', 'Akasia'],
  ['Akasia Bold Italic', 'Akasia'],
  ['Akasia ExtraBold', 'Akasia'],
  ['Akasia ExtraBold Italic', 'Akasia ExtraBold'],
  ['Akasia Italic', 'Akasia'],
  ['Akasia Light', 'Akasia'],
  ['Akasia Light Italic', 'Akasia Light'],
  ['Akasia Regular', 'Akasia'],
  ['Akasia SemiBold', 'Akasia'],
  ['Akasia SemiBold Italic', 'Akasia'],
  ['Arimo Bold', 'Arimo'],
  ['Arimo Bold Italic', 'Arimo'],
  ['Arimo Italic', 'Arimo'],
  ['Arimo Regular', 'Arimo'],
  ['Caladea Bold', 'Caladea'],
  ['Caladea Bold Italic', 'Caladea'],
  ['Caladea Italic', 'Caladea'],
  ['Caladea Regular', 'Caladea'],
  ['Carlito Bold', 'Carlito'],
  ['Carlito Bold Italic', 'Carlito'],
  ['Carlito Italic', 'Carlito'],
  ['Carlito Regular', 'Carlito'],
  ['Cousine Bold', 'Cousine'],
  ['Cousine Bold Italic', 'Cousine'],
  ['Cousine Italic', 'Cousine'],
  ['Cousine Regular', 'Cousine'],
  ['DejaVu Serif', 'DejaVu Serif'],
  ['DejaVu Serif Bold', 'DejaVu Serif'],
  ['DejaVu Serif Bold Italic', 'DejaVu Serif'],
  ['DejaVu Serif Italic', 'DejaVu Serif'],
  ['Intos Display Bold', 'Intos Display'],
  ['Intos Display Bold Italic', 'Intos Display'],
  ['Intos Display Italic', 'Intos Display'],
  ['Intos Display Regular', 'Intos Display'],
  ['Noto Sans Symbols 2 Regular', 'Noto Sans Symbols 2'],
  ['Noto Sans Symbols Bold', 'Noto Sans Symbols'],
  ['Noto Sans Symbols Regular', 'Noto Sans Symbols'],
  ['Tinos Bold', 'Tinos'],
  ['Tinos Bold Italic', 'Tinos'],
  ['Tinos Italic', 'Tinos'],
  ['Tinos Regular', 'Tinos'],
];

/**
 * The default `FontRegistry`: the faces docx4j's four font jars provide, and nothing else.
 *
 * It is what `IdentityPlusMapper` uses where the caller supplies none, and the reason a
 * mapping computed here can be compared with a parity golden.  A consumer with real fonts
 * passes its own (a `SimpleFontRegistry` over what `fontkit` or `@font-face` reports).
 */
export const DEFAULT_FONT_REGISTRY: FontRegistry = new SimpleFontRegistry(
  JAR_FONTS.map(([name, family]) => new PhysicalFont(name, family)),
);
