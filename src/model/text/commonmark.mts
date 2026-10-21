// CR-012 phase 3: the Markdown writer docx4j's MarkdownExporter serializes with, a port of commonmark-java
// 0.30.0's org.commonmark.renderer.markdown (MarkdownWriter, CoreMarkdownNodeRenderer) and of the Markdown
// renderers of its GFM tables, strikethrough and footnotes extensions, with docx4j-markdown's math nodes
// (org.docx4j.markdown.math.MathExtension's renderer). Only the nodes docx4j's WmlToMarkdown builds are here,
// and each renders as commonmark-java renders it, so that this package's Markdown is docx4j's byte for byte.
//
// commonmark-java's licence, which its ported code keeps:
//
//   Copyright (c) 2015, Atlassian Pty Ltd
//   All rights reserved.
//
//   Redistribution and use in source and binary forms, with or without
//   modification, are permitted provided that the following conditions are met:
//
//   * Redistributions of source code must retain the above copyright notice, this
//     list of conditions and the following disclaimer.
//
//   * Redistributions in binary form must reproduce the above copyright notice,
//     this list of conditions and the following disclaimer in the documentation
//     and/or other materials provided with the distribution.
//
//   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
//   AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
//   IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
//   DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
//   FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
//   DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
//   SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
//   CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
//   OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
//   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

// ---------------------------------------------------------------- nodes (org.commonmark.node and the extensions')

/** A node of the Markdown tree: its parent and children (commonmark-java `Node`). */
export abstract class MdNode {
  parent: MdNode | undefined;
  readonly children: MdNode[] = [];
  /** What the node was made for, where a caller wants to know where its Markdown begins (`MarkdownExporter.blocks`). */
  tag: unknown;

  get firstChild(): MdNode | undefined {
    return this.children[0];
  }

  get lastChild(): MdNode | undefined {
    return this.children[this.children.length - 1];
  }

  get previous(): MdNode | undefined {
    const siblings = this.parent?.children;
    return siblings === undefined ? undefined : siblings[siblings.indexOf(this) - 1];
  }

  get next(): MdNode | undefined {
    const siblings = this.parent?.children;
    if (siblings === undefined) return undefined;
    const i = siblings.indexOf(this);
    return i < 0 ? undefined : siblings[i + 1];
  }

  appendChild(child: MdNode): void {
    child.unlink();
    child.parent = this;
    this.children.push(child);
  }

  prependChild(child: MdNode): void {
    child.unlink();
    child.parent = this;
    this.children.unshift(child);
  }

  private unlink(): void {
    if (this.parent === undefined) return;
    const i = this.parent.children.indexOf(this);
    if (i >= 0) this.parent.children.splice(i, 1);
    this.parent = undefined;
  }
}

export class MdDocument extends MdNode {}
export class MdParagraph extends MdNode {}
export class MdBlockQuote extends MdNode {}
export class MdHardLineBreak extends MdNode {}
export class MdStrongEmphasis extends MdNode {}
export class MdTableBlock extends MdNode {}
export class MdTableHead extends MdNode {}
export class MdTableBody extends MdNode {}
export class MdTableRow extends MdNode {}

export class MdHeading extends MdNode {
  constructor(readonly level: number) {
    super();
  }
}

/** A bullet or ordered list: tight when its items are not separated by blank lines. */
export abstract class MdListBlock extends MdNode {
  tight = true;
}

export class MdBulletList extends MdListBlock {
  /** `-` where unset. */
  marker: string | undefined;
}

export class MdOrderedList extends MdListBlock {
  /** 1 where unset. */
  markerStartNumber: number | undefined;
  /** `.` where unset. */
  markerDelimiter: string | undefined;
}

export class MdListItem extends MdNode {}

export class MdFencedCodeBlock extends MdNode {
  constructor(readonly literal: string) {
    super();
  }
}

export class MdThematicBreak extends MdNode {
  constructor(readonly literal: string | undefined) {
    super();
  }
}

export class MdHtmlInline extends MdNode {
  constructor(readonly literal: string) {
    super();
  }
}

export class MdText extends MdNode {
  constructor(public literal: string) {
    super();
  }
}

export class MdCode extends MdNode {
  constructor(readonly literal: string) {
    super();
  }
}

export class MdEmphasis extends MdNode {}

export class MdLink extends MdNode {
  constructor(readonly destination: string, readonly title: string | undefined) {
    super();
  }
}

export class MdImage extends MdNode {
  constructor(public destination: string, readonly title: string | undefined) {
    super();
  }
}

export type MdAlignment = 'LEFT' | 'CENTER' | 'RIGHT';

export class MdTableCell extends MdNode {
  alignment: MdAlignment | undefined;

  constructor(readonly header: boolean) {
    super();
  }
}

export class MdStrikethrough extends MdNode {
  constructor(readonly delimiter: string) {
    super();
  }
}

export class MdFootnoteDefinition extends MdNode {
  constructor(readonly label: string) {
    super();
  }
}

export class MdFootnoteReference extends MdNode {
  constructor(readonly label: string) {
    super();
  }
}

/** Inline math, `$...$` (docx4j `InlineMath`). */
export class MdInlineMath extends MdNode {
  constructor(readonly literal: string, readonly displayHint: boolean) {
    super();
  }
}

/** Display math, a `$$` block (docx4j `DisplayMath`). */
export class MdDisplayMath extends MdNode {
  constructor(readonly literal: string) {
    super();
  }
}

// ---------------------------------------------------------------- the writer (MarkdownWriter)

/** A set of ASCII characters (commonmark-java `AsciiMatcher`). */
type CharMatcher = ReadonlySet<string>;

const chars = (s: string): Set<string> => new Set(s.split(''));

/** Where each tagged node's Markdown begins: the start of the line its first character is written on. */
export type TagStarts = Map<unknown, number>;

class MarkdownWriter {
  out = '';
  private blockSeparator = 0;
  lastChar = '\u0000';
  atLineStart = true;
  private readonly prefixes: string[] = [];
  private readonly tight: boolean[] = [];
  private readonly rawEscapes: CharMatcher[] = [];
  /** A tag waiting for its node's first character. */
  pendingTag: unknown;
  /** Where the line being written starts in `out`. */
  private lineStart = 0;

  constructor(readonly starts: TagStarts) {}

  raw(s: string): void {
    this.flushBlockSeparator();
    this.write(s, undefined, true);
  }

  text(s: string, escape: CharMatcher): void {
    if (s === '') return;
    this.flushBlockSeparator();
    this.write(s, escape, true);
  }

  line(): void {
    this.write('\n', undefined, false);
    this.writePrefixes();
    this.atLineStart = true;
  }

  block(): void {
    this.blockSeparator = this.isTight() ? 1 : 2;
    this.atLineStart = true;
  }

  pushPrefix(prefix: string): void {
    this.prefixes.push(prefix);
  }

  writePrefix(prefix: string): void {
    const tmp = this.atLineStart;
    this.raw(prefix);
    this.atLineStart = tmp;
  }

  popPrefix(): void {
    this.prefixes.pop();
  }

  pushTight(tight: boolean): void {
    this.tight.push(tight);
  }

  popTight(): void {
    this.tight.pop();
  }

  pushRawEscape(escape: CharMatcher): void {
    this.rawEscapes.push(escape);
  }

  popRawEscape(): void {
    this.rawEscapes.pop();
  }

  /**
   * Writes `s`, escaping what `escape` and the raw escapes match. Content (`raw`, `text`) starts a pending
   * tag at the line it is written on; line ends, prefixes and block separators never do.
   */
  private write(s: string, escape: CharMatcher | undefined, content: boolean): void {
    if (s === '') {
      this.atLineStart = false;
      return;
    }
    if (content && this.pendingTag !== undefined) {
      if (!this.starts.has(this.pendingTag)) this.starts.set(this.pendingTag, this.lineStart);
      this.pendingTag = undefined;
    }
    let chunk = s;
    if (this.rawEscapes.length > 0 || escape !== undefined) {
      chunk = '';
      for (let i = 0; i < s.length; i++) {
        const c = s.charAt(i);
        if ((escape !== undefined && escape.has(c)) || this.rawEscapes.some((e) => e.has(c))) {
          chunk += c === '\n' ? '&#10;' : `\\${c}`;                    // a line end can't be escaped with \
        } else {
          chunk += c;
        }
      }
    }
    const newline = chunk.lastIndexOf('\n');
    if (newline >= 0) this.lineStart = this.out.length + newline + 1;
    this.out += chunk;
    this.lastChar = s.charAt(s.length - 1);
    this.atLineStart = false;
  }

  private writePrefixes(): void {
    for (const prefix of this.prefixes) this.write(prefix, undefined, false);
  }

  private flushBlockSeparator(): void {
    if (this.blockSeparator !== 0) {
      this.write('\n', undefined, false);
      this.writePrefixes();
      if (this.blockSeparator > 1) {
        this.write('\n', undefined, false);
        this.writePrefixes();
      }
      this.blockSeparator = 0;
    }
  }

  private isTight(): boolean {
    return this.tight.length > 0 && this.tight[this.tight.length - 1]!;
  }
}

// ---------------------------------------------------------------- the renderer (CoreMarkdownNodeRenderer and the extensions')

/**
 * The characters escaped in text: commonmark-java's own, and the extensions' special characters as docx4j
 * registers them (`|` the tables', `~` strikethrough's, `$` docx4j's math); `#` too in a heading.
 */
const TEXT_ESCAPE = chars('[]<>`*_&\n\\|~$');
const TEXT_ESCAPE_IN_HEADING = chars('[]<>`*_&\n\\|~$#');
const LINK_DESTINATION_NEEDS_ANGLE_BRACKETS = chars(' ()<>\n\\');
const LINK_DESTINATION_ESCAPE_IN_ANGLE_BRACKETS = chars('<>\n\\');
const LINK_TITLE_ESCAPE_IN_QUOTES = chars('"\n\\');
const PIPE = chars('|');
const ORDERED_LIST_MARKER = /^([0-9]{1,9})([.)])/;

type ListHolder = { parent: ListHolder | undefined; marker: string; number: number; delimiter: string | undefined };

/**
 * Renders a Markdown tree as commonmark-java's `MarkdownRenderer` does, with the tables, strikethrough,
 * footnotes and docx4j math extensions. `starts` receives, for each node with a `tag`, the offset of the
 * line its first character is written on.
 */
export function renderMarkdown(document: MdDocument, starts: TagStarts = new Map()): string {
  const r = new Renderer(new MarkdownWriter(starts));
  r.render(document);
  return r.writer.out;
}

class Renderer {
  private listHolder: ListHolder | undefined;
  /** The current table's column alignments, from its header cells. */
  private readonly columns: (MdAlignment | undefined)[] = [];

  constructor(readonly writer: MarkdownWriter) {}

  render(node: MdNode): void {
    const w = this.writer;
    // a tag whose node wrote nothing gives way to the next; a tag starts once
    const tagged = node.tag !== undefined && !w.starts.has(node.tag);
    if (tagged) w.pendingTag = node.tag;
    this.renderNode(node);
    if (tagged && w.pendingTag === node.tag) w.pendingTag = undefined;
  }

  private renderNode(node: MdNode): void {
    const w = this.writer;
    if (node instanceof MdDocument) {
      this.children(node);
      w.line();
    } else if (node instanceof MdThematicBreak) {
      w.raw(node.literal ?? '___');                                    // ___ isn't ambiguous with a list marker
      w.block();
    } else if (node instanceof MdHeading) {
      this.heading(node);
    } else if (node instanceof MdFencedCodeBlock) {
      this.fencedCodeBlock(node);
    } else if (node instanceof MdParagraph) {
      this.children(node);
      w.block();
    } else if (node instanceof MdBlockQuote) {
      w.writePrefix('> ');
      w.pushPrefix('> ');
      this.children(node);
      w.popPrefix();
      w.block();
    } else if (node instanceof MdBulletList || node instanceof MdOrderedList) {
      w.pushTight(node.tight);
      this.listHolder = node instanceof MdBulletList
        ? { parent: this.listHolder, marker: node.marker ?? '-', number: 0, delimiter: undefined }
        : { parent: this.listHolder, marker: '', number: node.markerStartNumber ?? 1, delimiter: node.markerDelimiter ?? '.' };
      this.children(node);
      this.listHolder = this.listHolder.parent;
      w.popTight();
      w.block();
    } else if (node instanceof MdListItem) {
      this.listItem(node);
    } else if (node instanceof MdCode) {
      this.code(node.literal);
    } else if (node instanceof MdEmphasis) {
      // nested emphasis takes the other delimiter
      const delimiter = w.lastChar === '*' ? '_' : '*';
      w.raw(delimiter);
      this.children(node);
      w.raw(delimiter);
    } else if (node instanceof MdStrongEmphasis) {
      w.raw('**');
      this.children(node);
      w.raw('**');
    } else if (node instanceof MdLink) {
      this.linkLike(node.title, node.destination, node, '[');
    } else if (node instanceof MdImage) {
      this.linkLike(node.title, node.destination, node, '![');
    } else if (node instanceof MdHtmlInline) {
      w.raw(node.literal);
    } else if (node instanceof MdHardLineBreak) {
      w.raw('  ');
      w.line();
    } else if (node instanceof MdText) {
      this.text(node);
    } else if (node instanceof MdTableBlock) {
      this.columns.length = 0;
      w.pushTight(true);
      this.children(node);
      w.popTight();
      w.block();
    } else if (node instanceof MdTableHead) {
      this.children(node);
      for (const alignment of this.columns) {
        w.raw('|');
        w.raw(alignment === 'LEFT' ? ':---' : alignment === 'RIGHT' ? '---:' : alignment === 'CENTER' ? ':---:' : '---');
      }
      w.raw('|');
      w.block();
    } else if (node instanceof MdTableBody) {
      this.children(node);
    } else if (node instanceof MdTableRow) {
      this.children(node);
      w.raw('|');                                                      // the | at the end of the line
      w.block();
    } else if (node instanceof MdTableCell) {
      if (node.parent?.parent instanceof MdTableHead) this.columns.push(node.alignment);
      w.raw('|');
      w.pushRawEscape(PIPE);
      this.children(node);
      w.popRawEscape();
    } else if (node instanceof MdStrikethrough) {
      w.raw(node.delimiter);
      this.children(node);
      w.raw(node.delimiter);
    } else if (node instanceof MdFootnoteReference) {
      w.raw('[^');
      w.raw(node.label);                                                // labels are written as they are
      w.raw(']');
    } else if (node instanceof MdFootnoteDefinition) {
      w.raw('[^');
      w.raw(node.label);
      w.raw(']: ');
      w.pushPrefix('    ');
      this.children(node);
      w.popPrefix();
    } else if (node instanceof MdInlineMath) {
      const delimiter = node.displayHint ? '$$' : '$';
      w.raw(delimiter);
      w.raw(node.literal.replace(/\n/g, ' '));
      w.raw(delimiter);
    } else if (node instanceof MdDisplayMath) {
      w.raw('$$');
      w.line();
      for (const line of node.literal.split('\n')) {
        w.raw(line);
        w.line();
      }
      w.raw('$$');
      w.block();
    } else {
      throw new Error(`No Markdown renderer for ${node.constructor.name}`);
    }
  }

  private children(parent: MdNode): void {
    for (const child of [...parent.children]) this.render(child);
  }

  private heading(heading: MdHeading): void {
    const w = this.writer;
    if (heading.level <= 2 && hasLineBreak(heading)) {
      // a Setext heading can have several lines, but only levels 1 and 2
      this.children(heading);
      w.line();
      w.raw(heading.level === 1 ? '===' : '---');
      w.block();
      return;
    }
    w.raw('#'.repeat(heading.level));
    w.raw(' ');
    this.children(heading);
    w.block();
  }

  private fencedCodeBlock(code: MdFencedCodeBlock): void {
    const w = this.writer;
    const literal = code.literal;
    // a fence longer than any run of backticks in the code
    const fence = '`'.repeat(Math.max(maxRunLength('`', literal) + 1, 3));
    w.raw(fence);
    w.line();
    if (literal !== '') {
      for (const line of linesOf(literal)) {
        w.raw(line);
        w.line();
      }
    }
    w.raw(fence);
    w.block();
  }

  private listItem(item: MdListItem): void {
    const w = this.writer;
    const holder = this.listHolder;
    if (holder === undefined) throw new Error('A list item outside a list');
    let marker: string;
    if (holder.delimiter === undefined) {
      marker = holder.marker;
    } else {
      marker = `${holder.number}${holder.delimiter}`;
      holder.number++;
    }
    const spaces = ' ';
    w.writePrefix(marker);
    w.writePrefix(spaces);
    w.pushPrefix(' '.repeat(marker.length + spaces.length));
    if (item.firstChild === undefined) {
      w.block();                                                        // an empty item
    } else {
      this.children(item);
    }
    w.popPrefix();
  }

  private code(literal: string): void {
    const w = this.writer;
    // backticks in the code: one more around it
    const backticks = maxRunLength('`', literal);
    w.raw('`'.repeat(backticks + 1));
    // a space inside where the code starts or ends with a backtick, or starts and ends with a space
    const addSpace = literal.startsWith('`') || literal.endsWith('`')
      || (literal.startsWith(' ') && literal.endsWith(' ') && /[^ ]/.test(literal));
    if (addSpace) w.raw(' ');
    w.raw(literal);
    if (addSpace) w.raw(' ');
    w.raw('`'.repeat(backticks + 1));
  }

  private text(text: MdText): void {
    // When in doubt, escape: a character special only at a line's start is escaped there alone.
    const w = this.writer;
    let literal = text.literal;
    if (w.atLineStart && literal !== '') {
      const c = literal.charAt(0);
      if (c === '-') {
        w.raw('\\-');                                                   // a bullet list marker
        literal = literal.substring(1);
      } else if (c === '#') {
        w.raw('\\#');                                                   // an ATX heading
        literal = literal.substring(1);
      } else if (c === '=') {
        if (text.previous !== undefined) {                              // a Setext heading, unless the block's first line
          w.raw('\\=');
          literal = literal.substring(1);
        }
      } else if (c >= '0' && c <= '9') {
        const m = ORDERED_LIST_MARKER.exec(literal);
        if (m) {
          w.raw(m[1]!);
          w.raw(`\\${m[2]!}`);
          literal = literal.substring(m[0].length);
        }
      } else if (c === '\t') {
        w.raw('&#9;');
        literal = literal.substring(1);
      } else if (c === ' ') {
        w.raw('&#32;');
        literal = literal.substring(1);
      }
    }
    const escape = text.parent instanceof MdHeading ? TEXT_ESCAPE_IN_HEADING : TEXT_ESCAPE;
    if (literal.endsWith('!') && text.next instanceof MdLink) {
      // an unescaped ! would make the link an image
      w.text(literal.substring(0, literal.length - 1), escape);
      w.raw('\\!');
    } else {
      w.text(literal, escape);
    }
  }

  private linkLike(title: string | undefined, destination: string, node: MdNode, opener: string): void {
    const w = this.writer;
    w.raw(opener);
    this.children(node);
    w.raw(']');
    w.raw('(');
    if ([...destination].some((c) => LINK_DESTINATION_NEEDS_ANGLE_BRACKETS.has(c))) {
      w.raw('<');
      w.text(destination, LINK_DESTINATION_ESCAPE_IN_ANGLE_BRACKETS);
      w.raw('>');
    } else {
      w.raw(destination);
    }
    if (title !== undefined) {
      w.raw(' ');
      w.raw('"');
      w.text(title, LINK_TITLE_ESCAPE_IN_QUOTES);
      w.raw('"');
    }
    w.raw(')');
  }
}

function hasLineBreak(node: MdNode): boolean {
  return node.children.some((c) => c instanceof MdHardLineBreak || hasLineBreak(c));
}

function maxRunLength(needle: string, s: string): number {
  let max = 0;
  let pos = 0;
  while (pos < s.length) {
    pos = s.indexOf(needle, pos);
    if (pos === -1) return max;
    let run = 0;
    do {
      pos += needle.length;
      run++;
    } while (s.startsWith(needle, pos));
    max = Math.max(run, max);
  }
  return max;
}

/** A literal's lines, `\n` being a terminator: "abc\n" is one line. */
function linesOf(literal: string): string[] {
  const parts = literal.split('\n');
  return parts[parts.length - 1] === '' ? parts.slice(0, -1) : parts;
}
