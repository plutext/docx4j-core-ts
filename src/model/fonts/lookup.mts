// Reaching a package's `RunFontSelector` from the content API, structurally.
//
// `Font.name` resolves a theme reference, which needs the selector; the content API must not
// import `WordprocessingMLPackage` or `MainDocumentPart` (that would be a runtime cycle, since
// `parts/wml` builds the `Body` views), so the lookup is by shape, as `Body.propertyResolver`
// and the change tracker are.
import type { RunFontSelector } from './RunFontSelector.mjs';

interface PackageLike {
  getMainDocumentPart?: () => { runFontSelectorOrUndefined?: RunFontSelector };
}

/** The package's selector, or undefined where it has none yet (nothing has awaited
 *  `getRunFontSelector()`), where the package has no main document part, or where there is no
 *  package at all - a fragment built on its own. */
export function runFontSelectorOf(pkg: unknown): RunFontSelector | undefined {
  try {
    return (pkg as PackageLike | undefined)?.getMainDocumentPart?.().runFontSelectorOrUndefined;
  } catch {
    return undefined;
  }
}
