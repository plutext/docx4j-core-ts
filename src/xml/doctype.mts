// SPDX-License-Identifier: Apache-2.0
//
// A DOCTYPE read by this package, never by the platform's parser (CR-033; the portfolio's docs/dtd_handling_policy.md, sections
// 6.2 and 6.4, Jason's decisions of 2026-11-11). xmldom on Node expands no entity and fetches nothing, but a browser's DOMParser
// (libxml2, expat) may expand an internal entity, so what a part reads as would depend on the runtime, and a part built to
// expand without end would reach the platform's limits, not this package's. So the DOCTYPE is taken out of the text before any
// parser sees it, and what XML 1.0 (section 5.1) tells a non-validating processor to do with its internal subset is done here:
//
// - each internal general entity's replacement text is put in place of its references, in content and in attribute values,
//   within the limits (`DOCTYPE_LIMITS`): nesting 8 deep, 10,000 references and 1 MiB of replacement text a part. Past one, or
//   at an entity that names itself, the text is refused as XML not well formed is, naming the limit;
// - each attribute default the internal subset gives (a literal or `#FIXED` value) is applied, after the parse, to the elements
//   of that name that lack the attribute;
// - **nothing outside the text is ever read**: no external subset, no external entity, no parameter entity. An external
//   entity's references read as nothing; after a parameter entity's reference, which is not read, later declarations are not
//   processed, as section 5.1 has it. What wasn't read is listed in the summary.
//
// A reference to an entity no declaration names reads as its own text (`&name;`), as xmldom read it before.
import { Docx4JException } from '../opc/exceptions.mjs';

/** The bounds on expansion, a part at a time (the policy's D7). */
export const DOCTYPE_LIMITS = Object.freeze({ depth: 8, references: 10_000, replacement: 1024 * 1024 });

/** What a DOCTYPE holds: declarations of each kind, something outside the text named, or nothing (`bare`). */
export type DoctypeHolds = 'bare' | 'elements' | 'attributes' | 'entities' | 'external';

/** What reading a DOCTYPE found and did. */
export interface DoctypeSummary {
  /** The root element name the DOCTYPE gives. */
  name: string;
  /** What it holds, in this order; `['bare']` for none of them. */
  holds: DoctypeHolds[];
  /** The entity references expanded. */
  expanded: number;
  /** The attribute defaults applied, once the text is parsed. */
  defaulted: number;
  /** What wasn't read: the external subset (`SYSTEM "…"`), each external entity (`&x;`) and parameter entity (`%x;`). */
  notRead: string[];
}

/** An attribute default from an `<!ATTLIST>`, its value ready to set. */
export interface AttributeDefault {
  element: string;
  attribute: string;
  value: string;
}

/** The text with the DOCTYPE taken out and its entities expanded, what it held, and the defaults to apply after the parse. */
export interface DoctypeReading {
  text: string;
  summary: DoctypeSummary;
  defaults: AttributeDefault[];
}

const PREDEFINED = new Set(['lt', 'gt', 'amp', 'apos', 'quot']);
const NAME = /^[A-Za-z_:À-￿][\w.:\-·À-￿]*/;

function malformed(reason: string): Docx4JException {
  return new Docx4JException(`Malformed XML: ${reason} (CR-033)`);
}

/** Where the prolog's DOCTYPE starts, or -1: after a byte order mark, the XML declaration, whitespace, comments and PIs. */
function doctypeStart(text: string): number {
  let i = text.charCodeAt(0) === 0xfeff ? 1 : 0;
  for (;;) {
    while (i < text.length && (text[i] === ' ' || text[i] === '\t' || text[i] === '\n' || text[i] === '\r')) i++;
    if (text.startsWith('<?', i)) {
      const end = text.indexOf('?>', i + 2);
      if (end < 0) return -1;
      i = end + 2;
    } else if (text.startsWith('<!--', i)) {
      const end = text.indexOf('-->', i + 4);
      if (end < 0) return -1;
      i = end + 3;
    } else break;
  }
  return text.startsWith('<!DOCTYPE', i) ? i : -1;
}

/** Whether the text's prolog holds a DOCTYPE, whole or not: what a part that may hold none is refused for (Part 2 §6.2.5). */
export function hasDoctype(text: string): boolean {
  return doctypeStart(text) >= 0;
}

/** The offset after a quoted literal starting at `i`. */
function afterLiteral(text: string, i: number): number {
  const end = text.indexOf(text[i]!, i + 1);
  if (end < 0) throw malformed('a DOCTYPE literal is not closed');
  return end + 1;
}

/** The offset after a markup declaration starting at `i`, its quoted literals respected. */
function afterDeclaration(text: string, i: number): number {
  for (let k = i + 2; k < text.length; k++) {
    const c = text[k];
    if (c === '"' || c === "'") k = afterLiteral(text, k) - 1;
    else if (c === '>') return k + 1;
  }
  throw malformed('a DOCTYPE declaration is not closed');
}

/** The DOCTYPE's parts: where it ends, its name, its external identifier, its internal subset. */
function doctypeAt(text: string, start: number): { end: number; name: string; external: string | undefined; subset: string } {
  let k = start + '<!DOCTYPE'.length;
  const ws = (): void => { while (k < text.length && /\s/.test(text[k]!)) k++; };
  ws();
  const name = NAME.exec(text.slice(k))?.[0];
  if (!name) throw malformed('a DOCTYPE names no root element');
  k += name.length;
  ws();
  let external: string | undefined;
  for (const keyword of ['SYSTEM', 'PUBLIC']) {
    if (!text.startsWith(keyword, k)) continue;
    const from = k;
    k += keyword.length;
    for (let n = keyword === 'PUBLIC' ? 2 : 1; n > 0; n--) {
      ws();
      if (text[k] !== '"' && text[k] !== "'") throw malformed(`a DOCTYPE's ${keyword} identifier has no literal`);
      k = afterLiteral(text, k);
    }
    external = text.slice(from, k);
    ws();
  }
  let subset = '';
  if (text[k] === '[') {
    const from = k + 1;
    for (k = from; ; k++) {
      if (k >= text.length) throw malformed('a DOCTYPE\'s internal subset is not closed');
      if (text.startsWith('<!--', k)) {
        const end = text.indexOf('-->', k + 4);
        if (end < 0) throw malformed('a comment in a DOCTYPE is not closed');
        k = end + 2;
      } else if (text.startsWith('<?', k)) {
        const end = text.indexOf('?>', k + 2);
        if (end < 0) throw malformed('a processing instruction in a DOCTYPE is not closed');
        k = end + 1;
      } else if (text.startsWith('<!', k)) {
        k = afterDeclaration(text, k) - 1;
      } else if (text[k] === ']') break;
    }
    subset = text.slice(from, k);
    k++;
    ws();
  }
  if (text[k] !== '>') throw malformed('a DOCTYPE is not closed');
  return { end: k + 1, name, external, subset };
}

/** A literal's text, character references decoded (as an entity value's are at its declaration, XML 1.0 §4.5). */
function decodeCharacterReferences(value: string): string {
  return value.replace(/&#(x[0-9A-Fa-f]+|[0-9]+);/g, (_, code: string) => String.fromCodePoint(code[0] === 'x' ? parseInt(code.slice(1), 16) : parseInt(code, 10)));
}

/** The tokens of a declaration's body: quoted literals, parenthesised groups, `#KEYWORD`s, `%name;` and names. */
function tokensOf(body: string): string[] {
  const out: string[] = [];
  for (let k = 0; k < body.length;) {
    const c = body[k]!;
    if (/\s/.test(c)) { k++; continue; }
    if (c === '"' || c === "'") { const end = afterLiteral(body, k); out.push(body.slice(k, end)); k = end; continue; }
    if (c === '(') { const end = body.indexOf(')', k); if (end < 0) throw malformed('a DOCTYPE group is not closed'); out.push(body.slice(k, end + 1)); k = end + 1; continue; }
    const m = /^[^\s"'()]+/.exec(body.slice(k));
    out.push(m![0]);
    k += m![0].length;
  }
  return out;
}

interface Entity { value?: string; external?: string }

/** The internal subset's declarations, as XML 1.0 §5.1 has a non-validating processor take them. */
function declarationsOf(subset: string, summary: DoctypeSummary, holds: Set<DoctypeHolds>): { entities: Map<string, Entity>; defaults: { element: string; attribute: string; raw: string }[] } {
  const entities = new Map<string, Entity>();
  const defaults: { element: string; attribute: string; raw: string }[] = [];
  const declared = new Set<string>();
  // after a parameter entity's reference, which isn't read, later declarations aren't processed (§5.1)
  let stopped = false;
  for (let k = 0; k < subset.length;) {
    const c = subset[k]!;
    if (/\s/.test(c)) { k++; continue; }
    if (subset.startsWith('<!--', k)) { k = subset.indexOf('-->', k + 4) + 3; continue; }
    if (subset.startsWith('<?', k)) { k = subset.indexOf('?>', k + 2) + 2; continue; }
    if (c === '%') {
      const end = subset.indexOf(';', k);
      if (end < 0) throw malformed('a parameter entity reference is not closed');
      summary.notRead.push(subset.slice(k, end + 1));
      holds.add('external');
      stopped = true;
      k = end + 1;
      continue;
    }
    if (!subset.startsWith('<!', k)) throw malformed(`a DOCTYPE's internal subset holds ${JSON.stringify(subset.slice(k, k + 20))}`);
    const end = afterDeclaration(subset, k);
    const declaration = subset.slice(k, end - 1);
    k = end;
    const keyword = /^<!([A-Z]+)/.exec(declaration)?.[1];
    const tokens = tokensOf(declaration.slice(2 + (keyword?.length ?? 0)));
    if (keyword === 'ELEMENT') { holds.add('elements'); continue; }
    if (keyword === 'NOTATION') continue;
    if (keyword === 'ENTITY') {
      holds.add('entities');
      const parameter = tokens[0] === '%';
      const name = tokens[parameter ? 1 : 0];
      const rest = tokens.slice(parameter ? 2 : 1);
      if (!name) throw malformed('an entity declaration names no entity');
      const isExternal = rest[0] === 'SYSTEM' || rest[0] === 'PUBLIC';
      if (isExternal) {
        holds.add('external');
        summary.notRead.push(parameter ? `%${name};` : `&${name};`);
      }
      if (parameter || stopped || entities.has(name)) continue; // the first declaration binds (§4.2)
      if (isExternal) entities.set(name, { external: rest.slice(1).join(' ') });
      else if (rest[0] && (rest[0][0] === '"' || rest[0][0] === "'")) entities.set(name, { value: decodeCharacterReferences(rest[0].slice(1, -1)) });
      continue;
    }
    if (keyword === 'ATTLIST') {
      holds.add('attributes');
      const element = tokens[0];
      if (!element) throw malformed('an attribute-list declaration names no element');
      for (let t = 1; t < tokens.length;) {
        const attribute = tokens[t++];
        let type = tokens[t++];
        if (type === 'NOTATION') type = tokens[t++];
        const declaredDefault = tokens[t++];
        let raw: string | undefined;
        if (declaredDefault === '#FIXED') raw = tokens[t++];
        else if (declaredDefault && (declaredDefault[0] === '"' || declaredDefault[0] === "'")) raw = declaredDefault;
        if (!attribute || type === undefined) break;
        const key = `${element} ${attribute}`;
        if (raw === undefined || stopped || declared.has(key)) continue;
        declared.add(key);
        defaults.push({ element, attribute, raw: raw.slice(1, -1) });
      }
      continue;
    }
    throw malformed(`a DOCTYPE's internal subset holds an unknown declaration <!${keyword ?? ''}`);
  }
  return { entities, defaults };
}

/** The expansion of a text's entity references, within the limits; `attribute` for a quoted value's (the quote given). */
class Expander {
  references = 0;
  replacement = 0;

  constructor(private readonly entities: Map<string, Entity>) {}

  /** A reference's text: an internal entity's replacement text expanded, an external one's nothing, an unknown one's own. */
  private resolve(name: string, depth: number, stack: string[], attribute: string | undefined): string {
    if (PREDEFINED.has(name)) return `&${name};`;
    const entity = this.entities.get(name);
    if (!entity) return `&amp;${name};`;
    if (entity.value === undefined) return '';
    if (stack.includes(name)) throw malformed(`the entity ${name} refers to itself`);
    if (depth >= DOCTYPE_LIMITS.depth) throw malformed(`entities nested more than ${DOCTYPE_LIMITS.depth} deep`);
    if (++this.references > DOCTYPE_LIMITS.references) throw malformed(`more than ${DOCTYPE_LIMITS.references} entity references in one part`);
    const inner = [...stack, name];
    let text = attribute === undefined ? this.content(entity.value, depth + 1, inner) : this.value(entity.value, depth + 1, inner, attribute);
    if (attribute !== undefined) text = text.replace(/</g, '&lt;').replace(attribute === '"' ? /"/g : /'/g, attribute === '"' ? '&quot;' : '&apos;');
    this.replacement += text.length;
    if (this.replacement > DOCTYPE_LIMITS.replacement) throw malformed(`more than ${DOCTYPE_LIMITS.replacement} characters of entity replacement text in one part`);
    return text;
  }

  /** A reference starting at `i`: its name and where it ends, or undefined for a character reference or something else. */
  private referenceAt(text: string, i: number): { name: string; end: number } | undefined {
    const m = NAME.exec(text.slice(i + 1, i + 1 + 256));
    if (!m || text[i + 1 + m[0].length] !== ';') return undefined;
    return { name: m[0], end: i + 1 + m[0].length + 1 };
  }

  /** An attribute value's text, its references expanded. */
  value(text: string, depth: number, stack: string[], quote: string): string {
    let out = '';
    for (let i = 0; i < text.length; i++) {
      const ref = text[i] === '&' ? this.referenceAt(text, i) : undefined;
      if (!ref) { out += text[i]; continue; }
      out += this.resolve(ref.name, depth, stack, quote);
      i = ref.end - 1;
    }
    return out;
  }

  /** Content's text, its references expanded: outside comments, processing instructions and CDATA, in text and in attribute values. */
  content(text: string, depth: number, stack: string[]): string {
    let out = '';
    let i = 0;
    const copyTo = (marker: string, from: number): void => {
      const end = text.indexOf(marker, from);
      const stop = end < 0 ? text.length : end + marker.length;
      out += text.slice(i, stop);
      i = stop;
    };
    while (i < text.length) {
      const c = text[i]!;
      if (c === '<') {
        if (text.startsWith('<!--', i)) { copyTo('-->', i + 4); continue; }
        if (text.startsWith('<![CDATA[', i)) { copyTo(']]>', i + 9); continue; }
        if (text.startsWith('<?', i)) { copyTo('?>', i + 2); continue; }
        // a tag: its attribute values' references expanded, the rest copied
        let k = i + 1;
        out += '<';
        while (k < text.length && text[k] !== '>') {
          const q = text[k]!;
          if (q === '"' || q === "'") {
            const end = text.indexOf(q, k + 1);
            const stop = end < 0 ? text.length : end;
            out += q + this.value(text.slice(k + 1, stop), depth, stack, q) + (end < 0 ? '' : q);
            k = stop + 1;
          } else if (q === '<') break; // a tag cut off: the parser says so
          else { out += q; k++; }
        }
        if (text[k] === '>') { out += '>'; k++; }
        i = k;
        continue;
      }
      if (c === '&') {
        const ref = this.referenceAt(text, i);
        if (ref) { out += this.resolve(ref.name, depth, stack, undefined); i = ref.end; continue; }
      }
      out += c;
      i++;
    }
    return out;
  }
}

/** An attribute value's characters: what the parser makes of `&lt;`, `&#38;` and the rest, for `setAttribute`. */
function decodeValue(value: string): string {
  const named: Record<string, string> = { lt: '<', gt: '>', amp: '&', apos: "'", quot: '"' };
  return value.replace(/&(#x[0-9A-Fa-f]+|#[0-9]+|lt|gt|amp|apos|quot);/g, (_, ref: string) => {
    if (ref[0] !== '#') return named[ref]!;
    return String.fromCodePoint(ref[1] === 'x' ? parseInt(ref.slice(2), 16) : parseInt(ref.slice(1), 10));
  });
}

/**
 * The text with its prolog's DOCTYPE taken out (its lines kept, so a parse error's line is the source's) and its entities
 * expanded, what the DOCTYPE held, and the attribute defaults to apply after the parse; undefined for a text holding none.
 * Throws `Malformed XML` where the DOCTYPE isn't well formed or the expansion passes a limit.
 */
export function readDoctype(text: string): DoctypeReading | undefined {
  const start = doctypeStart(text);
  if (start < 0) return undefined;
  const { end, name, external, subset } = doctypeAt(text, start);
  const summary: DoctypeSummary = { name, holds: [], expanded: 0, defaulted: 0, notRead: [] };
  const holds = new Set<DoctypeHolds>();
  if (external !== undefined) { holds.add('external'); summary.notRead.push(external); }
  const { entities, defaults } = declarationsOf(subset, summary, holds);
  const expander = new Expander(entities);
  const lines = text.slice(start, end).replace(/[^\n]/g, '');
  const rest = expander.content(text.slice(end), 0, []);
  summary.expanded = expander.references;
  summary.holds = (['elements', 'attributes', 'entities', 'external'] as const).filter((h) => holds.has(h));
  if (summary.holds.length === 0) summary.holds = ['bare'];
  return {
    text: text.slice(0, start) + lines + rest,
    summary,
    defaults: defaults.map((d) => ({ element: d.element, attribute: d.attribute, value: decodeValue(expander.value(d.raw, 0, [], '"')) })),
  };
}

/** The defaults applied to the parsed document: each element of the name lacking the attribute given it. The count applied. */
export function applyAttributeDefaults(doc: Document, defaults: readonly AttributeDefault[]): number {
  let applied = 0;
  for (const { element, attribute, value } of defaults) {
    // a namespace declaration given by default would change what every name below means: not applied
    if (attribute === 'xmlns' || attribute.startsWith('xmlns:')) continue;
    const colon = attribute.indexOf(':');
    const elements = doc.getElementsByTagName(element);
    for (let i = 0; i < elements.length; i++) {
      const e = elements[i]!;
      if (e.hasAttribute(attribute)) continue;
      if (colon > 0) {
        const ns = e.lookupNamespaceURI(attribute.slice(0, colon));
        if (ns === null) continue;
        e.setAttributeNS(ns, attribute, value);
      } else e.setAttribute(attribute, value);
      applied++;
    }
  }
  return applied;
}
