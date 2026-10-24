// CR-008 phase 2: colours as Office draws them (docx4j `org.docx4j.model.color`, its CR-034): any colour reference in
// a docx, pptx or xlsx resolved to the RGB value (and alpha) Word, PowerPoint or Excel draws for it.
//
// A reference resolves against a {@link ThemeColors}: a theme's colour scheme and the colour map that leads to it,
// made for the reference's host by `await ThemeColors.of(pkg or part)` (a docx: the main document part's theme and
// the settings' `w:clrSchemeMapping`, which applies to WordprocessingML theme colours and DrawingML scheme colours
// alike, and is `PropertyResolver.getThemeColors()`; a slide, layout, master, notes slide, notes or handout master:
// its master's theme with the master's `p:clrMap`, then the layout's and the slide's `p:clrMapOvr`; an xlsx: the
// workbook's theme; a chart, `forChart` on its host's). A package with no theme part gets the theme Office supplies,
// the bundled one `pkg.fonts.defaultTheme` names (Office 2023's by default).
//
// - {@link WordColors}: `w:color`, `w:shd`, the borders and `w:u`. Word draws the theme reference and keeps the
//   value beside it as a cache, which can be stale; `w:themeTint` and `w:themeShade` change the HSL luminance, the
//   tint winning where both are given.
// - {@link DrawingMLColors}: the six DrawingML colour kinds and their transforms, applied in document order; a system
//   colour is the machine's, not its `lastClr`. A shape's style references (`a:fillRef`, `a:lnRef`, `a:fontRef`)
//   resolve through the theme's format scheme; an entry that is not a solid fill comes back as a {@link StyleFill} of
//   its kind, not guessed at.
// - {@link SpreadsheetColors}: `rgb`, `theme` (whose first four indexes Excel reads as lt1, dk1, lt2, dk2),
//   `indexed` (the legacy palette or the workbook's `x:indexedColors`) and `tint`, in Windows' integer HLS.
//
// Each returns a {@link ResolvedColor}, or undefined for "automatic", which is the caller's default. Every rule was
// read from Office's own renderings (docx4j CR-034 section 11); the probes, Office's readings and the corpus tint and
// shade triples are this package's vectors too (`test/theme-colors/`, CR-008 section 7).
export { ResolvedColor } from './ResolvedColor.mjs';
export { ThemeColors } from './ThemeColors.mjs';
export { WordColors } from './WordColors.mjs';
export { DrawingMLColors } from './DrawingMLColors.mjs';
export { SpreadsheetColors } from './SpreadsheetColors.mjs';
export { StyleFill, type StyleFillKind } from './StyleFill.mjs';
