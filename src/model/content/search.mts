import { Docx4JException } from '../../opc/exceptions.mjs';
import { NotSupportedError } from '../../office-js/errors.mjs';

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
      // a code point, a surrogate pair whole, as Word's ? matches an emoji (check 58 case 13; CR-002 section 32, K4);
      // ? and * cross a paragraph's mark where a search reads one (`wildcardMatchesAcross`, check 59, K7)
      if (c === '?') source += '(?:[\\uD800-\\uDBFF][\\uDC00-\\uDFFF]|[\\s\\S])';
      else if (c === '*') source += '[\\s\\S]*?';
      // ^13, Word's wildcard code for a paragraph mark (check 59)
      else if (c === '^' && text.startsWith('13', i + 1)) { source += '\\r'; i += 2; }
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
 * Only the stretches between control characters that a match's ends fall in are segmented, each once
 * (`graphemeBoundaries`), so a search across a long body's paragraphs costs what each paragraph's does.
 */
export function matchesOf(text: string, pattern: RegExp): RegExpExecArray[] {
  const out: RegExpExecArray[] = [];
  pattern.lastIndex = 0;
  for (let m = pattern.exec(text); m !== null; m = pattern.exec(text)) {
    if (m[0].length === 0) { pattern.lastIndex++; continue; }
    out.push(m);
  }
  const isBoundary = out.length > 0 ? graphemeBoundaries(text) : undefined;
  return isBoundary === undefined ? out : out.filter((m) => isBoundary(m.index) && isBoundary(m.index + m[0].length));
}

/** A C0 or C1 control character: a grapheme cluster of its own, always a break before and after it but between CR and LF (UAX #29, GB3 to GB5). */
const isControl = (code: number): boolean => code < 0x20 || (code >= 0x7F && code < 0xA0);

/**
 * Whether a grapheme cluster starts at a code-unit offset (the text's end counts); undefined where every unit is its own
 * cluster. A control character always has a break before and after it, and no rule looks back past one, so each stretch
 * between controls (a paragraph's, in a search across marks) is segmented on its own, as it would be within the whole
 * text, and only the stretches asked about, each once. V8's segmenter is quadratic in its input's length: the whole of a
 * 510,000-character body didn't come back in 100 s (the editor's finding, 2026-10-26; CR-002 section 64).
 */
function graphemeBoundaries(text: string): ((at: number) => boolean) | undefined {
  if (!/[\u0300-\uffff]/.test(text)) return undefined;
  const Segmenter = (Intl as { Segmenter?: new (locale?: string, options?: { granularity: 'grapheme' }) => { segment(input: string): Iterable<{ index: number }> } }).Segmenter;
  if (Segmenter === undefined) return undefined;
  const segmenter = new Segmenter(undefined, { granularity: 'grapheme' });
  const controls: number[] = [];
  for (let i = 0; i < text.length; i++) if (isControl(text.charCodeAt(i))) controls.push(i);
  // each stretch's cluster starts, by the stretch's start; undefined where every unit of it is its own cluster
  const stretches = new Map<number, Set<number> | undefined>();
  return (at) => {
    if (at <= 0 || at >= text.length) return true;
    const before = text.charCodeAt(at - 1);
    const after = text.charCodeAt(at);
    if (isControl(before) || isControl(after)) return before !== 0x0D || after !== 0x0A;
    // the stretch holding `at`: from just after the last control before it to the first control after it
    let lo = 0;
    let hi = controls.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (controls[mid]! < at) lo = mid + 1; else hi = mid; }
    const start = lo > 0 ? controls[lo - 1]! + 1 : 0;
    if (!stretches.has(start)) {
      const stretch = text.slice(start, lo < controls.length ? controls[lo]! : text.length);
      let starts: Set<number> | undefined;
      if (/[\u0300-\uffff]/.test(stretch)) {
        starts = new Set<number>();
        for (const segment of segmenter.segment(stretch)) starts.add(start + segment.index);
      }
      stretches.set(start, starts);
    }
    const starts = stretches.get(start);
    return starts === undefined || starts.has(at);
  };
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

/** A paragraph's part in a wildcard search across its marks: its segments, the [from, to) of its text read, its text's length, whether its mark follows. */
export interface MarkedSegments { segments: readonly SearchSegment[]; from: number; to: number; length: number; mark: boolean }

/** A hit of a wildcard search across marks: from offset `start` in piece `first` to `end` in piece `last`, which holds its mark where `mark`. */
export interface MarkedMatch { first: number; start: number; last: number; end: number; mark: boolean; match: RegExpExecArray }

/**
 * A wildcard query's hits over consecutive paragraphs, each followed by its mark where `mark` says, as Word's wildcards
 * match a paragraph mark (check 59; CR-002 section 32, K7): `?`, `*` and a class cross one, `^13` is one. The text read
 * in each paragraph is `officeMatches`' for wildcards; the caller gives only paragraphs a hit may span (one cell's, or
 * none's) and only the marks it may hold (a paragraph's own, never a cell's or a row's end).
 */
export function wildcardMatchesAcross(pieces: readonly MarkedSegments[], query: string, options: SearchOptions): MarkedMatch[] {
  const pattern = searchPattern(query, { ...options, matchWildcards: true });
  let text = '';
  const starts: [number, number][] = [];
  const ends: [number, number][] = [];
  const marks: boolean[] = [];
  const put = (c: string, k: number, start: number, end: number, mark: boolean): void => { text += c; starts.push([k, start]); ends.push([k, end]); marks.push(mark); };
  pieces.forEach((piece, k) => {
    for (const seg of piece.segments) {
      if (seg.hidden) continue;
      if (seg.object) { if (piece.from < seg.start && seg.start < piece.to) put(OBJECT, k, seg.start, seg.start, false); continue; }
      for (let i = 0; i < seg.text.length; i++) {
        const at = seg.start + i;
        if (at >= piece.from && at < piece.to) put(seg.text[i]!, k, at, at + 1, false);
      }
    }
    if (piece.mark) put('\r', k, piece.length, piece.length, true);
  });
  return matchesOf(text, pattern).map((match) => {
    const a = match.index;
    const b = a + match[0].length;
    return { first: starts[a]![0], start: starts[a]![1], last: ends[b - 1]![0], end: ends[b - 1]![1], mark: marks[b - 1]!, match };
  });
}

/**
 * Refuses a wildcard replacement where a hit holds a paragraph mark (CR-002 section 32, K7): `search` finds it, as
 * Word's does (check 59), but what replacing a mark leaves isn't measured, as for a plain `^p` (`refuseMarkReplacement`).
 */
export function refuseWildcardMarkReplacement(className: string, hits: readonly { spansParagraphs: boolean; holdsEndMark: boolean }[]): void {
  if (!hits.some((hit) => hit.spansParagraphs || hit.holdsEndMark)) return;
  throw new NotSupportedError(className, 'replaceText', "a hit holds a paragraph mark, which Word's wildcards match (check 59), and what replacing "
    + 'one does isn\'t measured (CR-002 section 32); search for the hits and edit their ranges');
}

/**
 * A plain query split at Word's paragraph mark code, `^p`: its pieces, one more than the marks, each a plain
 * query of its own (`^^` stays an escaped caret); undefined where the query holds no `^p`. What a `Range`'s search
 * across paragraphs matches piece by piece (CR-014 phase 3).
 */
export function paragraphMarkPieces(query: string): string[] | undefined {
  const pieces: string[] = [];
  let piece = '';
  for (let i = 0; i < query.length; i++) {
    const c = query[i]!;
    if (c === '^' && query[i + 1] === 'p') { pieces.push(piece); piece = ''; i++; continue; }
    if (c === '^' && i + 1 < query.length) { piece += c + query[i + 1]!; i++; continue; }
    piece += c;
  }
  if (pieces.length === 0) return undefined;
  pieces.push(piece);
  return pieces;
}

/**
 * Refuses a replacement whose query holds `^p` (CR-014 phase 3): `search` finds such a hit, but what Office JS's
 * replacement of a paragraph mark leaves isn't measured, and `insertText('Replace')` keeps a last mark (check 52).
 * Until phase 3 such a query matched the two characters `^p`, where Word matches a mark.
 */
export function refuseMarkReplacement(className: string, find: string, options?: SearchOptions): void {
  if (options?.matchWildcards || options?.matchRegExp || !paragraphMarkPieces(find)) return;
  throw new NotSupportedError(className, 'replaceText', "the query holds ^p, Word's paragraph mark, and what replacing one does isn't "
    + "measured (CR-014 phase 3); search for the hits and edit their ranges");
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
