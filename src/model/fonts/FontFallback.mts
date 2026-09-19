// docx4j `org.docx4j.fonts.FontFallback`, its class and candidate half.
//
// What is ported: the font *class* of a name (from `FontSubstitutions.xml` where it knows the
// font, else from the name), the families deliberately left to the document default, and
// `selectByClass` - a face of the document font's class, which is what
// `Mapper.addClassBasedSubstitutes` takes.  The coverage half (`selectCovering`, `covers`,
// `needsCoverage`, `warnNoCoverage`) needs a font file's cmap and is a later CR with `fontkit`
// (CR-001 section 14.2); `isSymbol`, `isEmoji`, `isEastAsianForm` and `coverageGroupOf` are
// here because `RunFontSelector.spanScript` cuts spans by them.
import type { FontRegistry } from './registry.mjs';
import type { PhysicalFont } from './PhysicalFont.mjs';
import { VCL_ENTRIES, type VclEntry } from './substitutions.generated.mjs';
import { scriptOf } from './scripts.mjs';

export type FontClass = 'SERIF' | 'SANS' | 'MONO' | 'UNKNOWN';

/** The key convention `FontSubstitutions.xml` uses: the name lower-cased, everything but
 *  letters and digits removed. */
export function shortKey(fontName: string): string {
  let key = '';
  for (const c of fontName) {
    // Character.isLetterOrDigit over the whole of Unicode, which is what docx4j's key() uses
    if (/[\p{L}\p{Nd}]/u.test(c)) key += c.toLowerCase();
  }
  return key;
}

/**
 * The `FontSubstitutions.xml` entry for this name, or undefined.  The table knows families,
 * not the foundry and weight suffixes a document adds to them ("Calisto MT", "Segoe UI
 * Light"), so trailing words are dropped one at a time until one matches.
 */
export function vclEntryFor(documentFontName: string | undefined): VclEntry | undefined {
  if (documentFontName === undefined || documentFontName === null) return undefined;
  const direct = VCL_ENTRIES[shortKey(documentFontName)];
  if (direct) return direct;
  const words = documentFontName.trim().split(/\s+/);
  for (let drop = 1; drop < words.length; drop++) {
    const entry = VCL_ENTRIES[shortKey(words.slice(0, words.length - drop).join(''))];
    if (entry) return entry;
  }
  return undefined;
}

/** Whether `FontSubstitutions.xml` calls this font condensed (Arial Narrow and the like): no
 *  ordinary stand-in has its widths, and docx4j's corpus showed the document default is the
 *  lesser evil, so these are deliberately left unmapped. */
export function isCondensed(documentFontName: string | undefined): boolean {
  const entry = vclEntryFor(documentFontName);
  if (entry && entry.fontWidth.toLowerCase().includes('condensed')) return true;
  const n = (documentFontName ?? '').toLowerCase();
  return n.includes('narrow') || n.includes('condensed');
}

/** `.+-(bold|italic|...)(mt|ps)?`: a PostScript name, which no system has a family by, so
 *  Word does not resolve it either and falls back to the document default. */
const POSTSCRIPT_NAME = /^.+-(bold|italic|bolditalic|regular|roman|light|medium|semibold|demibold|black|oblique)(mt|ps)?$/;

/** Families measured to be better off falling back to the document default than to a stand-in
 *  of their own class, because the stand-in's widths are further from theirs than the
 *  default's are.  Kept short and evidence-based, as docx4j's is. */
function leaveUnmapped(documentFontName: string | undefined): boolean {
  if (documentFontName === undefined) return false;
  const n = documentFontName.toLowerCase();
  // Lato is a narrow humanist sans; Arimo (Arial's widths) set a Lato document a page longer
  if (n.startsWith('lato')) return true;
  return POSTSCRIPT_NAME.test(n);
}

/** Whether this family is one docx4j deliberately leaves to the document default rather than
 *  standing in for it: a condensed face, or one of the measured exceptions. */
export function leftToTheDocumentDefault(documentFontName: string | undefined): boolean {
  return isCondensed(documentFontName) || leaveUnmapped(documentFontName);
}

/** The class `FontSubstitutions.xml` gives this font, or undefined where it does not say. */
function classFromTable(documentFontName: string | undefined): FontClass | undefined {
  const entry = vclEntryFor(documentFontName);
  if (!entry || entry.fontType.length === 0) return undefined;
  const type = entry.fontType.toLowerCase();
  // CJK, Symbol, Special, Decorative and Script faces have no stand-in of "their class"
  if (type.includes('cjk') || type.includes('symbol') || type.includes('special')
    || type.includes('decorative') || type.includes('script') || type.includes('ctl')) return 'UNKNOWN';
  if (type.includes('fixed') || type.includes('typewriter')) return 'MONO';
  if (type.includes('sansserif')) return 'SANS';
  if (type.includes('serif')) return 'SERIF';
  return undefined;
}

const MONO_WORDS = ['cousine', 'monospace', 'courier', 'consol', 'typewriter', 'terminal',
  'fixedsys', 'inconsolata', 'menlo'];
// NB no bare "gothic": in a Microsoft font name it is as often Japanese (MS Gothic, Yu Gothic)
// as it is a sans.  "urw gothic" is the whole family name of the Avant Garde clone.
const SANS_WORDS = ['arimo', 'carlito', 'cantarell', 'helvetica', 'arial', 'urw gothic',
  'selawik', 'verdana', 'tahoma', 'segoe', 'gadugi', 'trebuchet', 'calibri', 'grotesk',
  'grotesque', 'futura', 'frutiger', 'myriad', 'univers', 'avenir', 'lato', 'roboto',
  'franklin', 'open sans'];
// "p052" and "c059" are the URW Palatino and Century Schoolbook clones, "gelasio" Georgia's.
const SERIF_WORDS = ['tinos', 'caladea', 'charter', 'times', 'georgia', 'gelasio', 'garamond',
  'palatino', 'p052', 'c059', 'bookman', 'book antiqua', 'cambria', 'constantia', 'century',
  'baskerville', 'caslon', 'utopia', 'minion', 'sylfaen', 'roman'];

/**
 * The class a name says, sans before serif so that "HelveticaNeue LT 55 Roman" is a sans.
 *
 * @param generic whether a name which only ends in "Sans" or "Serif" counts.  It does when the
 *   question is what class an installed font is (Liberation Sans, DejaVu Serif), but not when
 *   it is whether to stand in for a document font: measured, that guess was worth less than the
 *   document default.
 */
export function classFromName(documentFontName: string | undefined, generic = true): FontClass {
  if (documentFontName === undefined) return 'UNKNOWN';
  const n = documentFontName.toLowerCase();
  for (const mono of MONO_WORDS) if (n.includes(mono)) return 'MONO';
  // "mono" only as a word: Monotype Corsiva is a script face, not a monospace one
  for (const word of n.split(/[^a-z0-9]+/)) if (word === 'mono') return 'MONO';
  for (const sans of SANS_WORDS) if (n.includes(sans)) return 'SANS';
  for (const serif of SERIF_WORDS) if (n.includes(serif)) return 'SERIF';
  if (generic) {
    if (n.includes('sans')) return 'SANS';
    if (n.includes('serif')) return 'SERIF';
  }
  return 'UNKNOWN';
}

/** The font's class, from `FontSubstitutions.xml` where it knows the font, else from its name. */
export function classOf(documentFontName: string | undefined): FontClass {
  return classFromTable(documentFontName) ?? classFromName(documentFontName);
}

/** The class to stand in for: the table's word, else what the name says for certain. */
export function substitutionClass(documentFontName: string | undefined): FontClass {
  return classFromTable(documentFontName) ?? classFromName(documentFontName, false);
}

const SANS_DEFAULTS = ['Arimo Regular', 'Arimo', 'Liberation Sans', 'Nimbus Sans', 'DejaVu Sans',
  'Noto Sans', 'FreeSans'];
const SERIF_DEFAULTS = ['Tinos Regular', 'Tinos', 'Liberation Serif', 'Nimbus Roman',
  'DejaVu Serif', 'Noto Serif', 'FreeSerif'];
const MONO_DEFAULTS = ['Cousine Regular', 'Cousine', 'Liberation Mono', 'Nimbus Mono PS',
  'DejaVu Sans Mono', 'Noto Mono', 'FreeMono'];

function defaultsFor(fontClass: FontClass): readonly string[] {
  switch (fontClass) {
    case 'SANS': return SANS_DEFAULTS;
    case 'SERIF': return SERIF_DEFAULTS;
    case 'MONO': return MONO_DEFAULTS;
    default: return [];
  }
}

/** The families `FontSubstitutions.xml` offers for this font, in its order, as short keys. */
function vclSubstituteKeys(documentFontName: string): readonly string[] {
  const entry = vclEntryFor(documentFontName);
  if (!entry || entry.substFonts.length === 0) return [];
  return entry.substFonts.split(';').map(shortKey).filter((k) => k.length > 0);
}

/**
 * A physical font of this document font's class, without regard to what it can render.  Used
 * where the characters are Latin (or Greek/Cyrillic), which any of these covers.
 *
 * `FontSubstitutions.xml`'s own list first - it is ordered by closeness - then the class
 * defaults.  Undefined where the font's class is unknown, or it is condensed or one of the
 * measured exceptions, or the registry has nothing of that class.
 */
export function selectByClass(registry: FontRegistry, documentFontName: string | undefined): PhysicalFont | undefined {
  if (documentFontName === undefined) return undefined;
  if (isCondensed(documentFontName) || leaveUnmapped(documentFontName)) return undefined;
  const fontClass = substitutionClass(documentFontName);
  if (fontClass === 'UNKNOWN') return undefined;

  const byKey = new Map<string, PhysicalFont>();
  for (const font of registry.all()) byKey.set(shortKey(font.name), font);
  for (const candidate of vclSubstituteKeys(documentFontName)) {
    const pf = byKey.get(candidate);
    if (pf && classOf(pf.name) === fontClass) return pf;
  }
  for (const candidate of defaultsFor(fontClass)) {
    const pf = registry.get(candidate);
    if (pf) return pf;
  }
  return undefined;
}

// ---------------------------------------------------------------------- coverage groups
//
// Used by `RunFontSelector.spanScript`, which cuts a span where the script changes.

/** U+2190-U+2BFF, the symbol blocks: Arrows, Mathematical Operators, Miscellaneous Technical,
 *  Box Drawing, Block Elements, Geometric Shapes, Miscellaneous Symbols, Dingbats, Braille and
 *  the supplemental arrow and maths blocks.  Every one of them is `COMMON`, so they are their
 *  own group rather than shared: the space and the digits beside an arrow are not dragged into
 *  the symbol font with it. */
export function isSymbol(cp: number): boolean {
  return cp >= 0x2190 && cp <= 0x2BFF;
}

/** U+1F000-U+1FAFF, the emoji blocks (`COMMON`, and like the symbols a group of their own). */
export function isEmoji(cp: number): boolean {
  return cp >= 0x1F000 && cp <= 0x1FAFF;
}

/** The East Asian punctuation and form blocks, which are `COMMON` - a fullwidth comma U+FF0C,
 *  an ideographic comma U+3001 - and which a Latin face does not have. */
export function isEastAsianForm(cp: number): boolean {
  return (cp >= 0x2E80 && cp <= 0x33FF) || (cp >= 0xFE10 && cp <= 0xFE4F)
    || (cp >= 0xFF00 && cp <= 0xFFEF);
}

export const SYMBOL_GROUP = 'SYMBOL';
export const EMOJI_GROUP = 'EMOJI';

/** The key a code point groups under: its Unicode script, except that the symbol and emoji
 *  blocks are their own groups. */
export function coverageGroupOf(cp: number): string {
  if (isSymbol(cp)) return SYMBOL_GROUP;
  if (isEmoji(cp)) return EMOJI_GROUP;
  return scriptOf(cp);
}
