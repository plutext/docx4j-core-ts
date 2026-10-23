// Which link targets the Markdown export writes as a link (CR-012 section 21, the port of docx4j CR-056 phase 1:
// `org.docx4j.convert.out.common.writer.LinkTargets`, docx4j 84ee9741c). A document's `w:hyperlink` relationship
// target or HYPERLINK field target was written as the link's destination as it is, so a crafted
// `javascript:alert(1)` target ran script in a page that rendered the Markdown as HTML (CWE-79, the defect
// apache/poi#1455 fixed in POI). The rule is docx4j's allowlist of schemes; a refused target gives the link's text
// with no link.

/** docx4j's default schemes (POI's set). */
const DEFAULT_SCHEMES: readonly string[] = Object.freeze(['http', 'https', 'ftp', 'mailto', 'file', 'news', 'tel']);

/** Leading controls and spaces removed; tabs and line breaks removed anywhere. */
function normalise(target: string): string {
  let out = '';
  for (const c of target) {
    if (c === '\t' || c === '\n' || c === '\r') continue;
    if (c <= ' ' && out.length === 0) continue;
    out += c;
  }
  return out;
}

/**
 * docx4j `LinkTargets`. A target is written as a link when it has no scheme (a relative path, a fragment, a UNC
 * path), is a Windows drive path (`C:\...`, `C:/...`), or its scheme is `http`, `https`, `ftp`, `mailto`, `file`,
 * `news` or `tel`, or one the caller adds. docx4j adds them through its JVM-wide property
 * `docx4j.convert.out.linkSchemes`; this package has no properties file, so they are an argument here and
 * `MarkdownExportOptions.linkSchemes` on the export, for a deployment whose documents carry `onenote:`, `ms-word:` or
 * `notes:` links. Added, never removed: a default stays.
 *
 * The scheme is read as a browser reads one (WHATWG URL parsing): leading control characters and spaces are ignored,
 * and tabs and line breaks are dropped anywhere, so ` \u0001java\tscri\npt:` is `javascript:`; the comparison ignores
 * case. Anything up to the first `:` (where that comes before any `/`, `?` or `#`) is the scheme, so a target whose
 * scheme a browser would reject as malformed (`java script:`) is refused too, rather than second-guessed.
 */
export const LinkTargets = {
  /** The default schemes, frozen. */
  defaultSchemes(): readonly string[] {
    return DEFAULT_SCHEMES;
  },

  /**
   * The schemes a link may use: the defaults plus `extra`, each trimmed and lower-cased, a trailing `:` dropped
   * and an empty one ignored (docx4j reads its property's comma-separated list so).
   */
  schemes(extra: readonly string[] = []): readonly string[] {
    const set = new Set(DEFAULT_SCHEMES);
    for (const raw of extra) {
      let s = raw.trim().toLowerCase();
      if (s.endsWith(':')) s = s.slice(0, -1);
      if (s !== '') set.add(s);
    }
    return [...set];
  },

  /**
   * Whether an output may write this target as a link: it has no scheme, is a drive path, or its scheme (read as
   * a browser reads it) is a default one or in `extra`. An undefined or empty target is safe (there is nothing to
   * write).
   */
  isSafe(target: string | undefined, extra: readonly string[] = []): boolean {
    const scheme = LinkTargets.scheme(target);
    if (scheme === undefined) return true;
    if (scheme.length === 1 && /\p{L}/u.test(scheme)) {
      // a Windows drive path, C:\ or C:/
      const next = normalise(target!).charAt(2);
      return next === '\\' || next === '/';
    }
    return LinkTargets.schemes(extra).includes(scheme);
  },

  /**
   * The target's scheme, read as a browser reads it, lower-cased; undefined where it has none (a relative path, a
   * fragment, a UNC path, or nothing at all).
   */
  scheme(target: string | undefined): string | undefined {
    if (target === undefined) return undefined;
    const s = normalise(target);
    for (let i = 0; i < s.length; i++) {
      const c = s.charAt(i);
      if (c === ':') return i === 0 ? undefined : s.slice(0, i).toLowerCase();
      if (c === '/' || c === '?' || c === '#') return undefined;
    }
    return undefined;
  },

  /** A target for a log line: at most 80 characters (UTF-16 units, as docx4j's), control characters shown as spaces. */
  excerpt(target: string | undefined): string {
    if (target === undefined) return 'undefined';
    const s = target.length > 80 ? target.slice(0, 80) + '...' : target;
    return s.replace(/[\u0000-\u001f]/g, ' ');
  },

  /** The message the export logs, once per document, when it refuses a target. */
  refusedMessage(target: string | undefined): string {
    return 'Link target refused (scheme not in MarkdownExportOptions.linkSchemes): the text is written without a link: '
      + LinkTargets.excerpt(target);
  },
};
