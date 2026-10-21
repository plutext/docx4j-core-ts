import { Docx4JException } from '../../opc/exceptions.mjs';

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
        // escaped to literals, so this changes what a pattern with them finds: `f(x)` matched the
        // text "f(x)" and now matches "fx", as Word's wildcards do. An unbalanced one is refused
        // below (CR-002 section 31).
        source += c;
      } else source += escapeRegExp(c);
    }
  } else {
    source = escapeRegExp(text);
  }
  if (options.matchWholeWord) source = `(?<![\\w])${source}(?![\\w])`;
  try {
    return new RegExp(source, options.matchCase ? 'g' : 'gi');
  } catch (cause) {
    // Only a wildcard pattern can fail here - a plain text is escaped whole - and the RegExp's own
    // message quotes the translation, a pattern the caller never wrote.
    throw new Docx4JException(`Not a valid wildcard pattern: ${text}`, { cause });
  }
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
 *
 * A match that starts or ends inside a grapheme cluster is dropped: half of an emoji's surrogate
 * pair (a wildcard `?` or class matches one UTF-16 unit), or an Arabic letter without the
 * combining mark after it (`\p{L}+` stops before the mark). Such a range is not text a person can
 * see, and replacing it splits a character - half a surrogate pair is saved as U+FFFD (CR-002
 * section 32 item 1, found by the editor's agreement suite). The clusters are `Intl.Segmenter`'s,
 * computed only for a text with a code unit at or above U+0300, below which every code unit is a
 * cluster of its own, so Latin text pays nothing; a runtime without `Intl.Segmenter` keeps every match.
 */
export function matchesOf(text: string, pattern: RegExp): RegExpExecArray[] {
  const out: RegExpExecArray[] = [];
  pattern.lastIndex = 0;
  for (let m = pattern.exec(text); m !== null; m = pattern.exec(text)) {
    if (m[0].length === 0) { pattern.lastIndex++; continue; }
    out.push(m);
  }
  const boundaries = out.length > 0 ? graphemeBoundaries(text) : undefined;
  return boundaries === undefined ? out : out.filter((m) => boundaries.has(m.index) && boundaries.has(m.index + m[0].length));
}

/** The code-unit offsets where a grapheme cluster starts, and the text's end; undefined where every unit is its own cluster. */
function graphemeBoundaries(text: string): Set<number> | undefined {
  if (!/[\u0300-\uffff]/.test(text)) return undefined;
  const Segmenter = (Intl as { Segmenter?: new (locale?: string, options?: { granularity: 'grapheme' }) => { segment(input: string): Iterable<{ index: number }> } }).Segmenter;
  if (Segmenter === undefined) return undefined;
  const boundaries = new Set<number>([text.length]);
  for (const segment of new Segmenter(undefined, { granularity: 'grapheme' }).segment(text)) boundaries.add(segment.index);
  return boundaries;
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

// --- Office JS's search over a paragraph (CR-012 phase 4, section 6.2 question 3; check 48) -------------

/** What `officeMatches` reads of a paragraph: its segments in Office JS's characters, the objects and hidden items listed. */
export interface SearchSegment {
  text: string;
  start: number;
  object?: boolean;
  hidden?: boolean;
  item: { name: { localPart: string } };
}

/** One match, as offsets in `Paragraph.text`, with the expression's match over the search's own text. */
export interface OfficeMatch { start: number; end: number; match: RegExpExecArray }

/** The object replacement character: a picture's place in the text `search` reads, which `^g` matches. */
const OBJECT = '￼';

/**
 * Word's Find codes a plain query may hold (check 48: each matched its character in Office JS's text):
 * `^f` a footnote reference, `^e` an endnote reference, `^l` a line break, `^m` a page break, `^t` a tab,
 * `^~` a non-breaking hyphen, `^-` an optional hyphen, `^g` a picture, `^^` a caret.
 */
const FIND_CODES: Record<string, string> = { f: '\u0002', e: '\u0002', l: '\u000B', m: '\u000C', t: '\t', '~': '\u001E', '-': '\u001F', g: OBJECT, '^': '^' };

/**
 * Every match of a query in a paragraph's text by Office JS's rules (check 48, Word's default view): the
 * text is `Paragraph.text` (hidden text left out, so a query matches across it and never in it), where a
 * picture is a place of no text that a plain query doesn't match across and `^g` matches, and an optional
 * hyphen is passed over unless the query names it (`^-`). A plain query takes Word's Find codes
 * (`FIND_CODES`), and its `-` matches a non-breaking hyphen too. Wildcards read the same text, the codes
 * not translated; a regular expression (`matchRegExp`, this package's extension) reads `Paragraph.text`
 * itself. Only text in [from, to) is read: a range's search.
 */
export function officeMatches(segments: readonly SearchSegment[], query: string, options: SearchOptions = {}, from = 0, to = Number.POSITIVE_INFINITY): OfficeMatch[] {
  const plain = !options.matchRegExp && !options.matchWildcards;
  const keepSoftHyphens = !plain || /\^-/.test(query.replace(/\^\^/g, ''));
  let pattern: RegExp;
  let tokens: (string | undefined)[] = [];
  if (plain) {
    const translated = plainQuery(query);
    tokens = translated.tokens;
    pattern = patternOf(translated.source, options);
  } else {
    pattern = searchPattern(query, options);
  }
  // the usual paragraph: nothing but text, no optional hyphen to pass over, no reference a code must tell
  // apart, so the text searched is Paragraph.text's own [from, to), each offset itself
  if (segments.every((seg) => !seg.object && !seg.hidden) && tokens.every((kind) => kind === undefined)) {
    const whole = segments.map((seg) => seg.text).join('');
    const start = Math.max(0, from);
    const span = whole.substring(start, Math.min(whole.length, to));
    if (keepSoftHyphens || !span.includes('\u001F')) {
      return matchesOf(span, pattern).map((match) => ({ start: start + match.index, end: start + match.index + match[0].length, match }));
    }
  }
  let text = '';
  const starts: number[] = [];
  const ends: number[] = [];
  const notes: (string | undefined)[] = [];
  for (const seg of segments) {
    if (seg.hidden) continue;
    if (seg.object) {
      if (!options.matchRegExp && from < seg.start && seg.start < to) { text += OBJECT; starts.push(seg.start); ends.push(seg.start); notes.push(undefined); }
      continue;
    }
    for (let i = 0; i < seg.text.length; i++) {
      const at = seg.start + i;
      if (at < from || at >= to) continue;
      const c = seg.text[i]!;
      if (c === '\u001F' && !keepSoftHyphens) continue;
      text += c;
      starts.push(at);
      ends.push(at + 1);
      notes.push(c === '\u0002' ? seg.item.name.localPart : undefined);
    }
  }
  const out: OfficeMatch[] = [];
  for (const match of matchesOf(text, pattern)) {
    const a = match.index;
    const b = a + match[0].length;
    // ^f and ^e both read U+0002: each matches only its own kind of reference
    if (tokens.some((kind, i) => kind !== undefined && !(notes[a + i] ?? '').startsWith(kind))) continue;
    out.push({ start: starts[a]!, end: ends[b - 1]!, match });
  }
  return out;
}

/** A plain query as an expression over `officeMatches`' text, and which of its characters are `^f` or `^e`. */
function plainQuery(query: string): { source: string; tokens: (string | undefined)[] } {
  let source = '';
  const tokens: (string | undefined)[] = [];
  for (let i = 0; i < query.length; i++) {
    const c = query[i]!;
    const code = c === '^' ? query[i + 1] : undefined;
    if (code !== undefined && FIND_CODES[code] !== undefined) {
      source += escapeRegExp(FIND_CODES[code]!);
      tokens.push(code === 'f' ? 'footnote' : code === 'e' ? 'endnote' : undefined);
      i++;
    } else if (c === '-') {
      source += '[-\\u001E]';                                         // a non-breaking hyphen is a hyphen to a plain query
      tokens.push(undefined);
    } else {
      source += escapeRegExp(c);
      tokens.push(undefined);
    }
  }
  return { source, tokens };
}

/** `searchPattern`'s flags and whole-word guard over a source already written as an expression. */
function patternOf(source: string, options: SearchOptions): RegExp {
  const guarded = options.matchWholeWord ? `(?<![\\w])${source}(?![\\w])` : source;
  return new RegExp(guarded, options.matchCase ? 'g' : 'gi');
}
