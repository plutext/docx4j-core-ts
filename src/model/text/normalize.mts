// CR-012 phase 2: WD REQ-1201's normalisation, docx4j's `Texts.normalize`.

/**
 * WD REQ-1201's normalisation, under which texts are compared (`mce.branches_differ`, REQ-307; the
 * conformance suite): CriticMarkup's delimiters, then white space (Java's `\s`, the six ASCII white-space
 * characters, and `\p{Z}`), `*`, `_`, `\`, `{`, `}` and `$` removed. No Unicode normalisation: zero-width
 * and tag characters, and the hyphens, are kept. JavaScript's `\s` would take U+FEFF too, so the
 * characters are named.
 */
export function normalize(value: string | undefined | null): string {
  if (value === undefined || value === null) return '';
  return value.replace(/\{\+\+|\+\+\}|\{--|--\}|\{>>|<<\}/g, '').replace(/[\t\n\x0B\f\r\p{Z}*_\\{}$]/gu, '');
}
