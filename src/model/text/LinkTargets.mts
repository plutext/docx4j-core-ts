// Which link targets the Markdown export writes as a link (CR-012 section 21, the port of docx4j CR-056 phase 1:
// `org.docx4j.convert.out.common.writer.LinkTargets`, docx4j 84ee9741c). A document's `w:hyperlink` relationship
// target or HYPERLINK field target was written as the link's destination as it is, so a crafted
// `javascript:alert(1)` target ran script in a page that rendered the Markdown as HTML (CWE-79, the defect
// apache/poi#1455 fixed in POI). The rule is docx4j's allowlist of schemes, read as written and with its character
// references decoded (docx4j's CR-056 section 14, CR-012 section 21.2 here); a refused target gives the link's text with
// no link.

/** docx4j's default schemes (POI's set). */
const DEFAULT_SCHEMES: readonly string[] = Object.freeze(['http', 'https', 'ftp', 'mailto', 'file', 'news', 'tel']);

/**
 * The HTML5 named character references that decode to an ASCII character. A scheme is ASCII, so these are the ones that
 * can build one or end one (`&colon;`, `&sol;`, `&num;`, `&quest;`, `&Tab;`, `&NewLine;`); any other decodes to a
 * character no scheme holds, or stays as written, and either way leaves a scheme malformed, which is refused.
 */
const ASCII_REFERENCES: Readonly<Record<string, string>> = Object.freeze({
  Tab: '\t', NewLine: '\n', excl: '!', quot: '"', QUOT: '"', num: '#', dollar: '$', percnt: '%', amp: '&', AMP: '&',
  apos: "'", lpar: '(', rpar: ')', ast: '*', midast: '*', plus: '+', comma: ',', period: '.', sol: '/', colon: ':',
  semi: ';', lt: '<', LT: '<', equals: '=', gt: '>', GT: '>', quest: '?', commat: '@', lsqb: '[', lbrack: '[',
  bsol: '\\', rsqb: ']', rbrack: ']', Hat: '^', lowbar: '_', UnderBar: '_', grave: '`', DiacriticalGrave: '`',
  lcub: '{', lbrace: '{', verbar: '|', vert: '|', VerticalLine: '|', rcub: '}', rbrace: '}',
});

/** docx4j `isSafeAsWritten`, `isSafe` for one reading of the target: no scheme, a drive path, or an allowed scheme. */
function safeAsWritten(target: string, extra: readonly string[]): boolean {
  const scheme = LinkTargets.scheme(target);
  if (scheme === undefined) return true;
  if (scheme.length === 1 && /\p{L}/u.test(scheme)) {
    // a Windows drive path, C:\ or C:/
    const next = normalise(target!).charAt(2);
    return next === '\\' || next === '/';
  }
  return LinkTargets.schemes(extra).includes(scheme);
}

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
    if (target === undefined || target === '') return true;
    if (!safeAsWritten(target, extra)) return false;
    // as a CommonMark renderer reads a destination, its character references decoded (docx4j CR-056 section 14)
    const decoded = LinkTargets.decodeCharacterReferences(target)!;
    return decoded === target || safeAsWritten(decoded, extra);
  },

  /**
   * The target with its character references decoded as CommonMark decodes a link destination: a decimal (`&#58;`) or
   * hex (`&#x3A;`) reference to its character, U+FFFD for zero, out of range or a surrogate; a named reference that
   * stands for one ASCII character (`&colon;`, `&sol;`, `&Tab;`, ...) to that character. Other named references, and an
   * `&` that begins no reference, stay as written. Decoded once, as a renderer decodes once. docx4j
   * `decodeCharacterReferences`.
   */
  decodeCharacterReferences(target: string | undefined): string | undefined {
    if (target === undefined || !target.includes('&')) return target;
    let out = '';
    let i = 0;
    while (i < target.length) {
      const c = target.charAt(i);
      if (c !== '&') { out += c; i++; continue; }
      const semi = target.indexOf(';', i);
      if (semi < 0 || semi - i > 40) { out += c; i++; continue; }
      const ref = target.substring(i + 1, semi);
      let decoded: string | undefined;
      if (ref.length > 1 && ref.charAt(0) === '#') {
        const hex = ref.charAt(1) === 'x' || ref.charAt(1) === 'X';
        const digits = ref.substring(hex ? 2 : 1);
        if (digits !== '' && digits.length <= (hex ? 6 : 7) && (hex ? /^[0-9A-Fa-f]+$/ : /^[0-9]+$/).test(digits)) {
          const cp = parseInt(digits, hex ? 16 : 10);
          decoded = cp === 0 || cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff) ? '\ufffd' : String.fromCodePoint(cp);
        }
      } else {
        decoded = Object.prototype.hasOwnProperty.call(ASCII_REFERENCES, ref) ? ASCII_REFERENCES[ref] : undefined;
      }
      if (decoded === undefined) { out += c; i++; continue; }
      out += decoded;
      i = semi + 1;
    }
    return out;
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
