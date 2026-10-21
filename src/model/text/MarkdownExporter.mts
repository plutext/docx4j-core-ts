// CR-012 phase 3: Markdown from the block model (WD section 7.2, REQ-703 and REQ-704), docx4j-markdown's
// org.docx4j.markdown.MarkdownExporter and MarkdownExportOptions (VERSION_17_3_4, unchanged since 198d015a3).
import type { Part } from '../../parts/Part.mjs';
import { BinaryPart } from '../../parts/BinaryPart.mjs';
import { DefaultXmlPart } from '../../parts/DefaultXmlPart.mjs';
import { TextOptions, type TextOptionsInit, type TextModelView } from './TextOptions.mjs';
import { TextResult, type TextBlock, type TextRun } from './TextResult.mjs';
import type { TextPackage } from './TextContext.mjs';
import { Texts } from './Texts.mjs';
import { WmlToMarkdown } from './WmlToMarkdown.mjs';
import { renderMarkdown, type TagStarts } from './commonmark.mjs';
import { base64Encode } from '../../xml/dom.mjs';

const SVG = 'image/svg+xml';

/** A picture's image, for a `MarkdownExportOptions.imageDestination`. */
export interface MarkdownImage {
  /** The picture's run in the block model. */
  readonly run: TextRun;
  /** The image part. */
  readonly part: Part;
  readonly bytes: Uint8Array;
  readonly contentType: string;
  /** The picture's number among the export's images with an image part, from 1. */
  readonly index: number;
  /** docx4j's file name for it, `image1.png`: `image`, the index, and the extension of its content type. */
  readonly fileName: string;
}

/** Where a picture's Markdown image points: a data URI by default, or what a host makes of the image. */
export type MarkdownImageDestination = (image: MarkdownImage) => string | Promise<string>;

/** `MarkdownExportOptions.of`'s argument. */
export interface MarkdownExportOptionsInit {
  /** The text model's options the export reads a package with; the default is the text model's. */
  textOptions?: TextOptions | TextOptionsInit;
  /** The view, which sets `textOptions.view`. */
  view?: TextModelView;
  /** Where each picture points; a data URI where unset. */
  imageDestination?: MarkdownImageDestination;
}

/**
 * Options for `MarkdownExporter`. What text is written is the text model's, under its `TextOptions`
 * (`textOptions`): the view, the stories, hidden text, field codes, metadata as YAML front matter, under the
 * names of the WD's section 10. The options here are how it is written. docx4j `MarkdownExportOptions`.
 *
 * docx4j's `imageDirPath` and `imageTargetUri`, which write each image to a directory, are `imageDestination`
 * here: nothing outside `./node` writes a file (CR-012 section 16.4). A host that wants the files writes
 * them in the function, which is given docx4j's file name, and returns the link.
 */
export class MarkdownExportOptions {
  /**
   * The text model's options a package is read with (`MarkdownExporter.export(pkg)`): the default is the
   * text model's, the `accepted` view, the main story with its text boxes and notes, hidden text left out.
   * A result is written under the options it was read with.
   */
  textOptions: TextOptions = new TextOptions();
  /** Where each picture points; a data URI (`data:image/png;base64,...`) where unset. */
  imageDestination: MarkdownImageDestination | undefined;

  static of(init?: MarkdownExportOptionsInit | MarkdownExportOptions): MarkdownExportOptions {
    if (init instanceof MarkdownExportOptions) return init;
    const o = new MarkdownExportOptions();
    if (init?.textOptions !== undefined) o.textOptions = TextOptions.of(init.textOptions);
    if (init?.view !== undefined) o.view = init.view;
    o.imageDestination = init?.imageDestination;
    return o;
  }

  /** The revision view: `accepted` (the default), `original` or `markup`. */
  get view(): TextModelView {
    return this.textOptions.view;
  }

  set view(view: TextModelView) {
    this.textOptions.view = view;
  }
}

/** A block's Markdown (`MarkdownExporter.blocks`). */
export interface MarkdownBlock {
  readonly block: TextBlock;
  /**
   * Its Markdown: the lines from the start of the line its first character is on to the next block's, the
   * blank line after it included. Empty for a block that writes nothing: an empty paragraph, or a code
   * paragraph merged into the fence before it.
   */
  readonly markdown: string;
}

/** A result's Markdown block by block (`MarkdownExporter.blocks`). */
export interface MarkdownBlocks {
  /** The YAML front matter, under `metadata: true` (REQ-704); empty otherwise. */
  readonly frontMatter: string;
  /** Each block of the result, in its order, with its Markdown. */
  readonly blocks: readonly MarkdownBlock[];
  /** The whole: the front matter and the blocks' Markdown joined, what `export` gives. */
  readonly markdown: string;
}

/**
 * Exports WordprocessingML to Markdown, CommonMark with GFM's tables and strikethrough, footnotes and math
 * (WD section 7.2). The text is the text model's (`Texts.of`, WD REQ-202): its block model is written as a
 * Markdown tree, which a port of commonmark-java's Markdown renderer serializes, so the escaping and the list
 * markers are commonmark-java's, and the Markdown docx4j's byte for byte.
 *
 * - a heading from its outline level, not its style's name; a list from its numbering, kept apart from an
 *   adjacent list and going on at Word's count after a paragraph that interrupts it (REQ-703);
 * - bold, italic, struck-through and code text where a run adds them to its paragraph's style; a code-styled
 *   paragraph, or one in a mono paragraph style, a fenced code block;
 * - links from a `w:hyperlink` or a HYPERLINK field; pictures with the describer's text, their image a data
 *   URI unless `imageDestination` says otherwise;
 * - notes `[^1]` and `[^e1]`, their definitions after the body; GFM tables, the header the `w:tblHeader` rows
 *   or the first row, a cell's paragraphs `<br>` apart; block quotes from the Quote styles;
 * - in the `markup` view, CriticMarkup: `{++inserted++}`, `{--deleted--}` and `{>>comment<<}`;
 * - under `metadata: true`, the document properties as YAML front matter (REQ-704).
 *
 * Async because a picture's bytes are read from its part. The package is not changed (REQ-204).
 * docx4j `MarkdownExporter`.
 *
 * ```ts
 * const md = await new MarkdownExporter().export(pkg);
 * const result = await Texts.of(pkg, { view: 'markup' });
 * PlainText.of(result);                     // the plain text and the Markdown of one reading
 * await MarkdownExporter.export(result, pkg);
 * ```
 */
export class MarkdownExporter {
  readonly options: MarkdownExportOptions;

  constructor(options?: MarkdownExportOptions | MarkdownExportOptionsInit) {
    this.options = MarkdownExportOptions.of(options);
  }

  /**
   * The Markdown of a package (the text model's reading of it under `options.textOptions`), or of a result
   * the text model has made, under the options it was read with: its view, and its metadata as front matter
   * where it was read with `metadata: true`. A result's block quotes are nested by their indent beyond the
   * quote style's own where the package it was read from is given, by their indent alone where it isn't.
   */
  async export(source: TextPackage | TextResult, pkg?: TextPackage): Promise<string> {
    return (await this.blocks(source, pkg)).markdown;
  }

  /**
   * The Markdown of each block (CR-012 section 4.4, the editor's ask), the front matter apart: a host that
   * writes something before each block, or stops after some, has the pieces `export` joins.
   */
  async blocks(source: TextPackage | TextResult, pkg?: TextPackage): Promise<MarkdownBlocks> {
    let result: TextResult;
    let styles: TextPackage | undefined;
    if (source instanceof TextResult) {
      result = source;
      styles = pkg;
    } else {
      result = await Texts.of(source, this.options.textOptions);
      styles = source;
    }
    const writer = new WmlToMarkdown(result, await resolverOf(styles));
    const document = writer.convert();
    let index = 0;
    for (const { node, run } of writer.images) {
      node.destination = await this.destination(run, () => ++index);
    }
    const starts: TagStarts = new Map();
    const body = renderMarkdown(document, starts);
    const frontMatter = result.options.metadata ? MarkdownExporter.frontMatter(result) : '';
    return { frontMatter, blocks: sliced(result.blocks, starts, body), markdown: frontMatter + body };
  }

  /** `new MarkdownExporter().export(...)`. */
  static async export(source: TextPackage | TextResult, pkg?: TextPackage): Promise<string> {
    return new MarkdownExporter().export(source, pkg);
  }

  /** `new MarkdownExporter().blocks(...)`. */
  static async blocks(source: TextPackage | TextResult, pkg?: TextPackage): Promise<MarkdownBlocks> {
    return new MarkdownExporter().blocks(source, pkg);
  }

  /**
   * The document properties as a YAML front-matter block (WD REQ-704, REQ-515): a mapping for each kind of
   * property the package has, `core`, `app` and `custom`, in that order, each property by its element's
   * local name (core, application) or its name (custom), its value a double-quoted string. Empty where the
   * result's metadata has none.
   */
  static frontMatter(result: TextResult): string {
    let sb = '';
    for (const kind of [TextResult.CORE, TextResult.APP, TextResult.CUSTOM]) {
      const properties = result.metadata[kind];
      if (properties === null || typeof properties !== 'object') continue;
      const entries = properties instanceof Map ? [...properties.entries()] : Object.entries(properties);
      if (entries.length === 0) continue;
      sb += `${kind}:\n`;
      for (const [k, v] of entries) sb += `  ${yamlKey(String(k))}: ${yamlString(String(v))}\n`;
    }
    return sb === '' ? '' : `---\n${sb}---\n\n`;
  }

  /**
   * A picture's destination: its image part's bytes (an SVG picture's SVG where it has no PNG copy, docx4j
   * #699) as a data URI or as `imageDestination` says; empty where its image can't be had, so that its alt
   * text is kept.
   */
  private async destination(run: TextRun, next: () => number): Promise<string> {
    const part = run.image?.part;
    let bytes: Uint8Array | undefined;
    try {
      if (part instanceof BinaryPart || (part instanceof DefaultXmlPart && part.contentType === SVG)) bytes = await part.getBytes();
    } catch {
      bytes = undefined;                                               // its alt text, with no destination
    }
    if (part === undefined || bytes === undefined) return '';
    const contentType = part.contentType;
    if (this.options.imageDestination === undefined) return `data:${contentType};base64,${base64Encode(bytes)}`;
    const index = next();
    return this.options.imageDestination({ run, part, bytes, contentType, index, fileName: `image${index}.${extensionFor(contentType, part)}` });
  }
}

/** The resolver a block quote's style indent is asked of: the package's, where it has a styles part. */
async function resolverOf(pkg: TextPackage | undefined) {
  if (pkg === undefined) return undefined;
  try {
    // asking a package with no styles part for a resolver would add one to it
    if (!pkg.getMainDocumentPart().styleDefinitionsPart) return undefined;
    return await pkg.getPropertyResolver();
  } catch {
    return undefined;
  }
}

/** Each block's Markdown: from its start to the next start, the last to the end; a start is one block's. */
function sliced(blocks: readonly TextBlock[], starts: TagStarts, body: string): MarkdownBlock[] {
  const ordered = [...new Set(starts.values())].sort((a, b) => a - b);
  const end = new Map<number, number>();
  ordered.forEach((s, i) => end.set(s, ordered[i + 1] ?? body.length));
  const taken = new Set<number>();
  return blocks.map((block) => {
    const start = starts.get(block);
    if (start === undefined || taken.has(start)) return { block, markdown: '' };
    taken.add(start);
    return { block, markdown: body.slice(start, end.get(start)) };
  });
}

function yamlKey(key: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.-]*$/.test(key) ? key : yamlString(key);
}

/** A YAML double-quoted scalar: `\`, `"` and the control characters escaped. */
function yamlString(value: string): string {
  let sb = '"';
  for (let i = 0; i < value.length; i++) {
    const c = value.charAt(i);
    const code = value.charCodeAt(i);
    switch (c) {
      case '"': sb += '\\"'; break;
      case '\\': sb += '\\\\'; break;
      case '\n': sb += '\\n'; break;
      case '\r': sb += '\\r'; break;
      case '\t': sb += '\\t'; break;
      default:
        if (code < 0x20 || code === 0x7f || code === 0x85 || code === 0x2028 || code === 0x2029) sb += `\\u${code.toString(16).padStart(4, '0')}`;
        else sb += c;
    }
  }
  return `${sb}"`;
}

function extensionFor(contentType: string, part: Part): string {
  switch (contentType) {
    case 'image/png': return 'png';
    case 'image/jpeg': return 'jpg';
    case 'image/gif': return 'gif';
    case 'image/bmp': return 'bmp';
    case 'image/tiff': return 'tif';
    case 'image/svg+xml': return 'svg';
    case 'image/x-emf': return 'emf';
    case 'image/x-wmf': return 'wmf';
    default: {
      const name = part.partName.name;
      const dot = name.lastIndexOf('.');
      return dot >= 0 ? name.substring(dot + 1) : 'bin';
    }
  }
}
