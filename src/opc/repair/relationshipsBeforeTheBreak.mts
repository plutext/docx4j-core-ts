import { Namespaces } from '../../parts/Namespaces.mjs';

/**
 * A relationships part that isn't well formed, read as Word's repair reads it (CR-032; WD 0.76's ruling on WD 0.73's item (i);
 * docx4j CR-075 D5 and D6, `RelationshipsPart.readRelationshipsBeforeTheBreak`): each whole `Relationship` before the break.
 *
 * docx4j reads the part with a streaming parser and keeps each `Relationship` whose end it met before the parser stopped. A DOM
 * parser has no such reading, so this finds the same thing another way: the part's text is cut after each child of the root,
 * found by a scan that knows quotes, comments, processing instructions and CDATA, and the longest cut that parses, closed with the
 * root's end tag, holds the relationships before the break. The parser decides, never the scan: a cut the scan got wrong
 * doesn't parse. The last cut is tried first, which for a part cut off is the one more parse it takes; then the last cut at the
 * parse error's position, where the parser gave one (xmldom's locator); then a binary search. Each failed parse is a line xmldom
 * prints, so the order keeps them few.
 *
 * As docx4j has it: a `Relationship` is taken by its local name, in any namespace, and kept with its `Id`, `Type`, `Target` and
 * `TargetMode`; one without an `Id`, a `Type` or a `Target` is passed over, as neither reader could use it; what the part held
 * besides is not kept. The kept relationships are returned as a part in the right namespace.
 */
export interface KeptRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

/** The relationships before the break, and the part they make, well formed and in the relationships namespace. */
export interface RelationshipsReadShort {
  kept: KeptRelationship[];
  xml: string;
}

export function relationshipsBeforeTheBreak(
  text: string,
  parse: (xml: string) => Document,
  errorAt?: { line?: number | undefined; column?: number | undefined },
): RelationshipsReadShort {
  const kept = keptRelationships(text, parse, errorAt);
  const attr = (name: string, value: string): string => ` ${name}="${escapeAttribute(value)}"`;
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + `<Relationships xmlns="${Namespaces.RELATIONSHIPS}">`
    + kept.map((r) => `<Relationship${attr('Id', r.id)}${attr('Type', r.type)}${attr('Target', r.target)}`
      + `${r.targetMode === undefined ? '' : attr('TargetMode', r.targetMode)}/>`).join('')
    + '</Relationships>';
  return { kept, xml };
}

function keptRelationships(text: string, parse: (xml: string) => Document, errorAt?: { line?: number | undefined; column?: number | undefined }): KeptRelationship[] {
  const scanned = scan(text);
  if (!scanned) return [];
  const { root, cuts } = scanned;
  const tried = new Map<number, Document | undefined>();
  const parsed = (k: number): Document | undefined => {
    if (!tried.has(k)) {
      let doc: Document | undefined;
      try {
        doc = parse(`${text.slice(0, cuts[k])}</${root}>`);
      } catch {
        doc = undefined;
      }
      tried.set(k, doc);
    }
    return tried.get(k);
  };
  // a part cut off: the last cut, after the last whole child, parses, and it is the one parse
  const last = cuts.length - 1;
  let best = parsed(last) ? last : -1;
  // the last cut at or before the error's position, where the parser gave one: xmldom puts it on the character it stopped at,
  // which can be a cut's last
  const at = offsetOf(text, errorAt);
  if (best < 0 && at !== undefined) {
    let k = last;
    while (k >= 0 && cuts[k]! > at + 1) k--;
    if (k >= 0 && parsed(k)) best = k;
  }
  if (best < 0) {
    // otherwise the longest cut that parses: a cut that parses has every shorter one parse, so a binary search finds it
    if (!parsed(0)) return [];
    let lo = 0;
    let hi = cuts.length - 1;
    while (lo < hi) {
      const mid = Math.ceil((lo + hi) / 2);
      if (parsed(mid)) lo = mid;
      else hi = mid - 1;
    }
    best = lo;
  }
  const doc = parsed(best)!;
  const out: KeptRelationship[] = [];
  for (let c = doc.documentElement?.firstChild ?? null; c; c = c.nextSibling) {
    if (c.nodeType !== 1 || (c as Element).localName !== 'Relationship') continue;
    const e = c as Element;
    if (!e.hasAttribute('Id') || !e.hasAttribute('Type') || !e.hasAttribute('Target')) continue;
    const r: KeptRelationship = { id: e.getAttribute('Id')!, type: e.getAttribute('Type')!, target: e.getAttribute('Target')! };
    if (e.hasAttribute('TargetMode')) r.targetMode = e.getAttribute('TargetMode')!;
    out.push(r);
  }
  return out;
}

/**
 * The root's qualified name and the cuts: the offset after the root's start tag, then after each child of the root. Undefined
 * where the text has no whole start tag for its root, or the root is empty (`<Relationships/>`), or a DTD comes first, which
 * docx4j's reader, with DTDs off, stops at.
 */
function scan(text: string): { root: string; cuts: number[] } | undefined {
  let i = 0;
  let root: string | undefined;
  let depth = 0;
  const cuts: number[] = [];
  while (i < text.length) {
    const lt = text.indexOf('<', i);
    if (lt < 0) break;
    if (text.startsWith('<!--', lt)) {
      const end = text.indexOf('-->', lt + 4);
      if (end < 0) break;
      i = end + 3;
    } else if (text.startsWith('<![CDATA[', lt)) {
      const end = text.indexOf(']]>', lt + 9);
      if (end < 0) break;
      i = end + 3;
    } else if (text.startsWith('<?', lt)) {
      const end = text.indexOf('?>', lt + 2);
      if (end < 0) break;
      i = end + 2;
    } else if (text.startsWith('<!', lt)) {
      return undefined; // a DTD
    } else if (text.startsWith('</', lt)) {
      const end = text.indexOf('>', lt + 2);
      if (end < 0) break;
      i = end + 1;
      depth--;
      if (depth <= 0) break; // the root's end
      if (depth === 1) cuts.push(i);
    } else {
      const end = tagEnd(text, lt + 1);
      if (end < 0) break;
      i = end + 1;
      const empty = text[end - 1] === '/';
      if (root === undefined) {
        if (empty) return undefined;
        root = /^[^\s/>]+/.exec(text.slice(lt + 1))?.[0];
        if (!root) return undefined;
        depth = 1;
        cuts.push(i);
      } else if (empty) {
        if (depth === 1) cuts.push(i);
      } else {
        depth++;
      }
    }
  }
  return root === undefined ? undefined : { root, cuts };
}

/** The offset of the `>` ending the tag that starts before `from`, quotes respected; -1 where the text ends first. */
function tagEnd(text: string, from: number): number {
  let quote: string | undefined;
  for (let i = from; i < text.length; i++) {
    const c = text[i];
    if (quote) {
      if (c === quote) quote = undefined;
    } else if (c === '"' || c === "'") {
      quote = c;
    } else if (c === '>') {
      return i;
    } else if (c === '<') {
      return -1; // a tag cut off, another begun
    }
  }
  return -1;
}

/** The offset of a 1-based line and column, where the parser gave both. */
function offsetOf(text: string, at?: { line?: number | undefined; column?: number | undefined }): number | undefined {
  if (at?.line === undefined || at.column === undefined) return undefined;
  let offset = 0;
  for (let line = 1; line < at.line; line++) {
    const nl = text.indexOf('\n', offset);
    if (nl < 0) return undefined;
    offset = nl + 1;
  }
  return offset + at.column - 1;
}

function escapeAttribute(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
}
