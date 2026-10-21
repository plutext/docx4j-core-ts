// CR-012 phase 2: the text model's facade, docx4j's org.docx4j.model.text.Texts.
import { Part } from '../../parts/Part.mjs';
import { Docx4JException } from '../../opc/exceptions.mjs';
import { TextOptions, TextOptionError, type TextOptionsInit } from './TextOptions.mjs';
import { TextResult } from './TextResult.mjs';
import { Extractor } from './Extractor.mjs';
import { PlainText } from './PlainText.mjs';
import { PackageTrees, TextContext, type TextPackage } from './TextContext.mjs';
import { normalize } from './normalize.mjs';

type Options = TextOptionsInit | TextOptions | undefined;

const WML_CONTENT_TYPES = {
  main: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', 'application/vnd.ms-word.document.macroEnabled.main+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml', 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml'],
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  endnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
  comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
};

function optionsOf(options: Options): TextOptions {
  const o = TextOptions.of(options);
  o.checkImplemented();
  return o;
}

/** A content object: a typed value (`TYPE_NAME`) or an element pair. */
function isContent(value: unknown): boolean {
  if (typeof value !== 'object' || value === null || value instanceof TextOptions) return false;
  if (typeof (value as { TYPE_NAME?: unknown }).TYPE_NAME === 'string') return true;
  return 'name' in value && 'value' in value && typeof (value as { value?: { TYPE_NAME?: unknown } }).value?.TYPE_NAME === 'string';
}

function isPackage(value: unknown): value is TextPackage {
  return typeof value === 'object' && value !== null && typeof (value as TextPackage).getMainDocumentPart === 'function'
    && typeof (value as TextPackage).getPropertyResolver === 'function';
}

/**
 * docx4j's text model (its CR-046; the portfolio's *WordprocessingML Text Extraction*, WD): a text is
 * `serialize(scope(view(resolve(package))))` (WD REQ-201). `Texts.of` makes the block model
 * (`TextResult`); `PlainText.of` writes plain text from it.
 *
 * ```ts
 * const result = await Texts.of(pkg, { view: 'markup', include: ['headers', 'footers'], hidden: true });
 * PlainText.of(result);
 * result.warnings;
 * ```
 *
 * - **resolve**: each `mc:AlternateContent` is read from the branch the content API reads
 *   (`selectedBranchOf`, CR-010), so that the text model and `Paragraph.text` take the same branch.
 * - **view**: `accepted` (the default), `original` or `markup`; the first two are the text of the document
 *   `acceptAllRevisions` or `rejectAllRevisions` leaves, worked out without resolving anything
 *   (`ResolvedView`), held equal to that copy route by the agreement test (REQ-406).
 * - **scope**: the main story, each text box once after its anchoring paragraph (after the table for one in
 *   a cell), then the footnotes and endnotes in reference order; headers, footers and comments where
 *   included, and the glossary, last, under `include: glossary`; hidden text excluded; placeholders
 *   included; a bound control's value from its custom XML part; images' descriptions by an
 *   `ImageDescriber`; the document properties in the metadata.
 *
 * **Not yet** (CR-012 section 5): `fields: update` and `embedded: true` are refused (REQ-1001); an equation
 * is its text inside `$...$`, not LaTeX (phase 5b); a chart or SmartArt contributes nothing (5c); an
 * `altChunk` is not expanded and raises `altchunk.not_expanded` (5e).
 *
 * **The package is never changed** (REQ-204): every part is read with `readContents()`, the live tree where
 * the part is unmarshalled and a private one where it is not, so a save afterwards is byte for byte where it
 * was before.
 *
 * **Cost.** Each call reads every story and, in the `accepted` and `original` views, their tracked changes
 * once, whether it reads the whole package or one paragraph of it. A `TextContext` (`contextOf`) holds what
 * was read, so calls under one context read it once.
 *
 * **Not Office JS's text.** `Paragraph.text` and the other Office JS readers keep Office JS's text (CR-012
 * D2, REQ-202's emulated-API exemption); each `TextRun` records where its text lies in its paragraph's
 * `Paragraph.text` (`officeOffset`), so that the two map onto each other.
 */
export class Texts {
  private constructor() {}

  /** The text of a package (REQ-501's scope), or of one story part, or of content in a package. */
  static async of(pkg: TextPackage, options?: Options, context?: TextContext): Promise<TextResult>;
  /**
   * The text of one story part, as its own story: the main document part (its story and text boxes; the
   * notes it references numbered in the context, not read), a header or a footer (addressed
   * `header:rId3/...`), the footnotes or endnotes part (the notes the context's stories referenced, in
   * reference order; under a context that has not read the main story, the main story is read first, for
   * the order alone), or the comments part (the comments the view includes, in `TextResult.comments`).
   * The part must belong to a package. Read them in REQ-501's order under one context and
   * `TextResult.concat` the results to have the package's text.
   */
  static async of(part: Part, options?: Options, context?: TextContext): Promise<TextResult>;
  /**
   * The text of content in a package (a paragraph, a table, a body, a header's content), as the main story:
   * the view is the package's, so a join or a move is resolved as in the whole document. Its addresses
   * start `content`. The content is the package's own tree, its part unmarshalled (`getContents()`): a part
   * that is not is read privately, afresh for each context, and the views' plan is made over that reading.
   */
  static async of(pkg: TextPackage, content: object, options?: Options, context?: TextContext): Promise<TextResult>;
  /**
   * The text of content with no package: in the `accepted` and `original` views only run-level revisions
   * are resolved (no paragraph is joined, a move is its two halves), and hidden text is only what the runs
   * themselves hide, since there are no styles.
   */
  static async of(content: object, options?: Options): Promise<TextResult>;
  static async of(first: unknown, second?: unknown, third?: unknown, fourth?: unknown): Promise<TextResult> {
    if (first instanceof Part) return Texts.ofPart(first, second as Options, third as TextContext | undefined);
    if (isPackage(first)) {
      if (isContent(second)) return Texts.ofContent(first, second as object, third as Options, fourth as TextContext | undefined);
      const o = optionsOf(second as Options);
      const context = await Texts.context(first, third as TextContext | undefined, o);
      return new Extractor(context.trees, o, 'view', context).extract();
    }
    if (isContent(first)) {
      const o = optionsOf(second as Options);
      return new Extractor(PackageTrees.none(), o, 'planless', new TextContext(PackageTrees.none())).extractContent(first, 'MAIN', 'content', undefined);
    }
    throw new Docx4JException('Texts.of takes a WordprocessingMLPackage, a story part of one, or content');
  }

  private static async ofContent(pkg: TextPackage, content: object, options: Options, given: TextContext | undefined): Promise<TextResult> {
    const o = optionsOf(options);
    const context = await Texts.context(pkg, given, o);
    const v = 'name' in content && 'value' in content ? (content as { value: unknown }).value : content;
    return new Extractor(context.trees, o, 'view', context).extractContent(content, 'MAIN', 'content', context.trees.partOf(v));
  }

  private static async ofPart(part: Part, options: Options, given: TextContext | undefined): Promise<TextResult> {
    const o = optionsOf(options);
    const pkg = part.package as unknown;
    if (!isPackage(pkg)) throw new Docx4JException(`${part.partName.name} belongs to no WordprocessingML package; read content with Texts.of(content)`);
    const context = await Texts.context(pkg, given, o);
    const trees = context.trees;
    const main = trees.main;
    const ct = part.contentType;
    if (main && part === (main.part as unknown)) {
      return new Extractor(trees, o, 'view', context).extractContent(main.document, 'MAIN', 'body', part);
    }
    if (ct === WML_CONTENT_TYPES.header || ct === WML_CONTENT_TYPES.footer) {
      const rId = main ? relationshipIdOf(main.part as unknown as Part, part) : undefined;
      return new Extractor(trees, o, 'view', context).headerOrFooterPart(part, ct === WML_CONTENT_TYPES.header ? 'HEADER' : 'FOOTER', rId ?? part.partName.name);
    }
    if (ct === WML_CONTENT_TYPES.footnotes || ct === WML_CONTENT_TYPES.endnotes) {
      if (!context.mainRead && main) {
        // the notes are numbered in their references' order: read the main story for it, its blocks unused
        const first = context.sibling();
        new Extractor(trees, o, 'view', first).extractContent(main.document, 'MAIN', 'body', main.part as unknown as Part);
        context.adoptNotes(first);
        context.mainRead = true;
      }
      return new Extractor(trees, o, 'view', context).extractNotes(ct === WML_CONTENT_TYPES.footnotes ? 'footnote' : 'endnote');
    }
    if (ct === WML_CONTENT_TYPES.comments) return new Extractor(trees, o, 'view', context).extractComments();
    throw new Docx4JException(`Not a story part: ${part.partName.name}; use Texts.of(package)`);
  }

  /**
   * The text of a package that `acceptAllRevisions` or `rejectAllRevisions` has already resolved, read with
   * no view logic: the copy route of WD REQ-406, which the views of `Texts.of` are held equal to. The
   * options' view says which resolution it was (`accepted` or `original`).
   */
  static async ofResolved(resolved: TextPackage, options?: Options): Promise<TextResult> {
    const o = optionsOf(options);
    if (o.view === 'markup') throw new TextOptionError('A resolved package has no markup view', 'view', 'REQ-406', 'invalid');
    const context = await Texts.context(resolved, undefined, o);
    return new Extractor(context.trees, o, 'resolved', context).extract();
  }

  /** Plain text of a package (WD section 7.1): `PlainText.of(await Texts.of(pkg, options))`. */
  static async plainText(pkg: TextPackage, options?: Options): Promise<string> {
    return PlainText.of(await Texts.of(pkg, options));
  }

  /**
   * A context for calls over one package (CR-012 section 4.4): it reads the package's stories once, makes
   * each view's plan once, and carries the note numbers, numbering states and bookmarks from one call to the
   * next. A snapshot: after the package changes, make another.
   */
  static async contextOf(pkg: TextPackage): Promise<TextContext> {
    return new TextContext(await PackageTrees.read(pkg));
  }

  /** WD REQ-1201's normalisation, under which texts are compared. */
  static normalize(value: string | undefined | null): string {
    return normalize(value);
  }

  private static async context(pkg: TextPackage, given: TextContext | undefined, o: TextOptions): Promise<TextContext> {
    return contextFor(pkg, given, o);
  }
}

/**
 * The context a call reads under: the one given, checked against the package and the view, or a new one
 * (internal: `Texts` and the outline reader share it).
 */
export async function contextFor(pkg: TextPackage, given: TextContext | undefined, o: TextOptions): Promise<TextContext> {
  if (given !== undefined && given.pkg !== pkg) throw new Docx4JException('This TextContext was made for another package');
  const context = given ?? await Texts.contextOf(pkg);
  if (context.view !== undefined && context.view !== o.view) {
    throw new TextOptionError(`This TextContext read in the ${context.view} view, and its note numbers and list counts are that view's: make another for ${o.view}`,
      'view', 'REQ-401', 'invalid');
  }
  context.view = o.view;
  if (o.includes.has('glossary')) await context.trees.readGlossary();
  return context;
}

/** Options checked as `Texts.of` checks them (internal). */
export { optionsOf, isPackage };

/** The id of the main document part's relationship to a part. */
function relationshipIdOf(main: Part, part: Part): string | undefined {
  const rels = main.getRelationshipsPart(false);
  for (const rel of rels?.relationships.relationship ?? []) {
    if (rels?.getPart(rel) === part) return rel.id;
  }
  return undefined;
}

