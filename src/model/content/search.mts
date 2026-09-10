/** Office JS Word.SearchOptions, the subset honoured here. */
export interface SearchOptions {
  matchCase?: boolean;
  matchWholeWord?: boolean;
  /** Word wildcards: `?` one character, `*` any run, `[a-z]`, `[!x]`, `<` / `>` word start and end, `@` one or more of the previous. */
  matchWildcards?: boolean;
}

/** A regular expression for a search text under the options; global so it can iterate matches. */
export function searchPattern(text: string, options: SearchOptions = {}): RegExp {
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

/** All [start, end) matches of a pattern in a text. */
export function findAll(text: string, pattern: RegExp): [number, number][] {
  const out: [number, number][] = [];
  pattern.lastIndex = 0;
  for (let m = pattern.exec(text); m !== null; m = pattern.exec(text)) {
    if (m[0].length === 0) { pattern.lastIndex++; continue; }
    out.push([m.index, m.index + m[0].length]);
  }
  return out;
}
