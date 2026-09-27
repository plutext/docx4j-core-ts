/** Office JS Word.SearchOptions, the subset honoured here, plus `matchRegExp` (CR-002 section 26). */
export interface SearchOptions {
  matchCase?: boolean;
  matchWholeWord?: boolean;
  /** Word wildcards: `?` one character, `*` any run, `[a-z]`, `[!x]`, `<` / `>` word start and end, `@` one or more of the previous. */
  matchWildcards?: boolean;
  /**
   * The search text is an ECMAScript regular expression, compiled with `u` (extension beyond
   * Office JS, as `replaceText` is; `Word.SearchOptions` has no such member, so it stays out of
   * `Word.supported`). Never `m` or `s`: a match stays inside one paragraph, as every search here
   * does (CR-002 sections 26 and 29). `matchWildcards` as well is an error - they are two
   * languages for the same argument.
   */
  matchRegExp?: boolean;
}

/** A regular expression for a search text under the options; global so it can iterate matches. */
export function searchPattern(text: string, options: SearchOptions = {}): RegExp {
  if (options.matchRegExp && options.matchWildcards) {
    throw new RangeError('matchRegExp and matchWildcards are two languages for the same argument; choose one');
  }
  // The `u` flag goes on this branch ONLY. Under it several escapes that are merely redundant become
  // syntax errors, and the other two branches write some: the wildcard `[...]` doubles backslashes
  // and `escapeRegExp` escapes `-`. So it must not be moved to the one `new RegExp` below.
  if (options.matchRegExp) {
    // A user expression may have a top-level alternation, so `matchWholeWord` has to wrap it in a
    // group: `(?<!\w)cat|dog(?!\w)` is two alternatives with one guard each and matches `dog` in
    // `dogma`. A non-capturing group keeps the group numbers `$1` refers to.
    const source = options.matchWholeWord ? `(?<![\\w])(?:${text})(?![\\w])` : text;
    return new RegExp(source, options.matchCase ? 'gu' : 'giu');
  }
  let source: string;
  if (options.matchWildcards) {
    source = '';
    for (let i = 0; i < text.length; i++) {
      const c = text[i]!;
      if (c === '?') source += '.';
      else if (c === '*') source += '.*?';
      else if (c === '<') source += '\\b(?=\\w)';
      else if (c === '>') source += '\\b(?<=\\w)';
      else if (c === '@') source += '+';
      else if (c === '[') {
        const close = text.indexOf(']', i);
        if (close < 0) { source += '\\['; continue; }
        let cls = text.substring(i + 1, close);
        if (cls.startsWith('!')) cls = '^' + cls.substring(1);
        source += `[${cls.replace(/\\/g, '\\\\')}]`;
        i = close;
      } else if (c === '{') {
        const close = text.indexOf('}', i);
        if (close < 0) { source += '\\{'; continue; }
        source += `{${text.substring(i + 1, close)}}`;
        i = close;
      } else if (c === '\\' && i + 1 < text.length) {
        source += '\\' + text[++i];
      } else if (c === '(' || c === ')') {
        // Word's wildcard grouping, which `\1` to `\9` in a replacement refer to (CR-002 section 26);
        // a literal parenthesis is `\(`, taken by the escape case above. Before section 26 both were
        // escaped to literals, so no expression that worked then changes meaning - one that used
        // them meant them literally and Word would not have matched it either.
        source += c;
      } else source += escapeRegExp(c);
    }
  } else {
    source = escapeRegExp(text);
  }
  if (options.matchWholeWord) source = `(?<![\\w])${source}(?![\\w])`;
  return new RegExp(source, options.matchCase ? 'g' : 'gi');
}

export function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * All [start, end) matches of a pattern in a text. Kept as it is - it is a public export, and every
 * caller that only wants spans would have to change - so the groups a replacement needs come from
 * `matchesOf` instead (CR-002 section 26's review note, trap 1).
 */
export function findAll(text: string, pattern: RegExp): [number, number][] {
  return matchesOf(text, pattern).map((m) => [m.index, m.index + m[0].length]);
}

/**
 * Every match of a pattern in a text, with its capture groups: what `replaceText` needs to expand
 * `$1` and Word's `\1`. An empty match advances by one rather than looping forever, as `findAll`
 * has always done.
 */
export function matchesOf(text: string, pattern: RegExp): RegExpExecArray[] {
  const out: RegExpExecArray[] = [];
  pattern.lastIndex = 0;
  for (let m = pattern.exec(text); m !== null; m = pattern.exec(text)) {
    if (m[0].length === 0) { pattern.lastIndex++; continue; }
    out.push(m);
  }
  return out;
}

/**
 * A replacement with its group references expanded against one match: `$1` to `$9`, `$&` and `$$`
 * for a regular expression (JavaScript's `String.prototype.replace` rules), and Word's `\1` to
 * `\9` for wildcards (CR-002 section 26). A reference to a group the pattern does not have expands
 * to nothing, as both languages do.
 */
export function expandReplacement(replacement: string, match: RegExpExecArray, options: SearchOptions = {}): string {
  if (options.matchRegExp) {
    return replacement.replace(/\$(\$|&|[1-9])/g, (_, token: string) =>
      token === '$' ? '$' : token === '&' ? match[0] : match[Number(token)] ?? '');
  }
  if (options.matchWildcards) {
    return replacement.replace(/\\([1-9])/g, (_, digit: string) => match[Number(digit)] ?? '');
  }
  return replacement;
}
