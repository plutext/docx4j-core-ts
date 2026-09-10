import { InvalidFormatException } from './exceptions.mjs';

/**
 * An OPC part name (docx4j org.docx4j.openpackaging.parts.PartName): a leading '/', no trailing
 * '/', no empty segments, no segment ending in '.', percent-encoding allowed. Equality is
 * case-insensitive, as Word writes mixed case (`[Content_Types].xml`, `_rels/.rels`).
 */
export class PartName {
  /** The name as given, e.g. '/word/document.xml'. */
  readonly name: string;
  /** Lower-cased name, the key for maps. */
  readonly key: string;

  private constructor(name: string) {
    this.name = name;
    this.key = name.toLowerCase();
  }

  /** The package root, the source of the package relationships. Not a valid part name otherwise. */
  static readonly ROOT = new PartName('/');

  /** Validates and wraps a part name; a leading '/' is added when missing. */
  static of(name: string | PartName): PartName {
    if (name instanceof PartName) return name;
    if (typeof name !== 'string' || name.length === 0) throw new InvalidFormatException('A part name shall not be empty [M1.1]');
    const full = name.startsWith('/') ? name : '/' + name;
    if (full === '/') throw new InvalidFormatException('A part name shall not be empty [M1.1]');
    if (full.endsWith('/')) throw new InvalidFormatException(`A part name shall not have a forward slash as the last character [M1.5]: ${full}`);
    if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(name)) throw new InvalidFormatException(`Absolute URI forbidden: ${name}`);
    for (const seg of full.substring(1).split('/')) {
      if (seg.length === 0) throw new InvalidFormatException(`A part name shall not have empty segments [M1.3]: ${full}`);
      if (seg.endsWith('.')) throw new InvalidFormatException(`A segment shall not end with a dot ('.') character [M1.9]: ${full}`);
      if (/^\.+$/.test(seg)) throw new InvalidFormatException(`A segment shall include at least one non-dot character [M1.10]: ${full}`);
      if (/[\\#?\s]/.test(seg)) throw new InvalidFormatException(`A segment shall not hold any characters other than pchar characters [M1.6]: ${full}`);
      if (/%(?![0-9A-Fa-f]{2})/.test(seg)) throw new InvalidFormatException(`The segment ${seg} contains an invalid encoded character`);
      if (/%2[fF]|%5[cC]/.test(seg)) throw new InvalidFormatException(`A segment shall not contain a percent-encoded forward or backward slash [M1.7]: ${full}`);
    }
    return new PartName(full);
  }

  /** 'xml' for '/word/document.xml'; '' when there is no dot. */
  get extension(): string {
    const last = this.name.substring(this.name.lastIndexOf('/') + 1);
    const i = last.lastIndexOf('.');
    return i > -1 ? last.substring(i + 1) : '';
  }

  /** '/word' for '/word/document.xml'; '/' for '/[Content_Types].xml'. */
  get directory(): string {
    const i = this.name.lastIndexOf('/');
    return i === 0 ? '/' : this.name.substring(0, i);
  }

  /** 'document.xml' for '/word/document.xml'. */
  get fileName(): string {
    return this.name.substring(this.name.lastIndexOf('/') + 1);
  }

  /** The name as stored in a container: without the leading '/'. */
  get storeName(): string {
    return this.name.substring(1);
  }

  /** True for '/_rels/.rels' and '/word/_rels/document.xml.rels'. */
  get isRelationshipsPart(): boolean {
    return /\/_rels\/[^/]*\.rels$/.test(this.key);
  }

  equals(other: PartName | string | undefined): boolean {
    if (other === undefined) return false;
    return this.key === (typeof other === 'string' ? other.toLowerCase() : other.key);
  }

  toString(): string {
    return this.name;
  }

  /** The relationships part name for a part: '/word/_rels/document.xml.rels'; for the root, '/_rels/.rels'. */
  static relsFor(source: PartName | string): PartName {
    const s = typeof source === 'string' ? source : source.name;
    if (s === '/') return PACKAGE_RELS;
    const pos = s.lastIndexOf('/');
    return new PartName(s.substring(0, pos) + '/_rels' + s.substring(pos) + '.rels');
  }

  /** The inverse of relsFor: '/word/document.xml' for '/word/_rels/document.xml.rels'; ROOT for '/_rels/.rels'. */
  static sourceOfRels(rels: PartName | string): PartName {
    const s = typeof rels === 'string' ? rels : rels.name;
    const m = /^(.*)\/_rels\/([^/]*)\.rels$/.exec(s);
    if (!m) throw new InvalidFormatException(`Not a relationships part name: ${s}`);
    if (m[1] === '' && m[2] === '') return PartName.ROOT;
    return new PartName(`${m[1]}/${m[2]}`);
  }

  /**
   * Resolves a relationship target against its source part (docx4j URIHelper.resolvePartUri):
   * relative to the source's directory, '.' and '..' collapsed; a target with a leading '/' is
   * absolute. The source may be the package root.
   */
  static resolve(source: PartName | string, target: string): PartName {
    const s = typeof source === 'string' ? source : source.name;
    const base = s === '/' ? '/' : s.substring(0, s.lastIndexOf('/') + 1);
    const path = target.startsWith('/') ? target : base + target;
    const out: string[] = [];
    for (const seg of path.split('/')) {
      if (seg === '' || seg === '.') continue;
      if (seg === '..') out.pop();
      else out.push(seg);
    }
    return PartName.of('/' + out.join('/'));
  }

  /**
   * The relationship target for a part from a source part (docx4j URIHelper.relativizeURI): from
   * '/word/document.xml' to '/word/media/image1.png' is 'media/image1.png'; from the root it is
   * the name without the leading '/'.
   */
  static relativize(source: PartName | string, target: PartName | string): string {
    const s = typeof source === 'string' ? source : source.name;
    const t = typeof target === 'string' ? target : target.name;
    if (s === '/') return t.substring(1);
    const from = s.substring(1, s.lastIndexOf('/') + 1).split('/').filter((x) => x !== '');
    const to = t.substring(1).split('/');
    let i = 0;
    while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
    return '../'.repeat(from.length - i) + to.slice(i).join('/');
  }
}

const PACKAGE_RELS = PartName.of('/_rels/.rels');
/** The container name of the content types stream; not a part, never in `parts`. */
export const CONTENT_TYPES_NAME = '[Content_Types].xml';
