// CR-012 phase 2: what one reading of a package needs, read before the walk begins (`PackageTrees`), and
// what a story needs from the rest of the document (`TextContext`, CR-012 section 4.4): the note references'
// order, the numbering states per story, the bookmarks, and the views' plans.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as w15 from '@docx4j/generated-objects-ts/modules/org_docx4j_w15';
import type * as w16cid from '@docx4j/generated-objects-ts/modules/org_docx4j_w16cid';
import type * as w16cex from '@docx4j/generated-objects-ts/modules/org_docx4j_w16cex';
import type * as core from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_core';
import type * as extended from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_extended';
import type * as custom from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_custom';
import type { Part } from '../../parts/Part.mjs';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import * as f from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import { PropertyResolver } from '../properties/PropertyResolver.mjs';
import { NumberingDefinitions } from '../listnumbering/definitions.mjs';
import { Emulator, type NumberingPartLike } from '../listnumbering/Emulator.mjs';
import { NumberingStates } from '../listnumbering/state.mjs';
import { Body } from '../content/Body.mjs';
import { ResolvedView } from '../content/ResolvedView.mjs';
import type { XPathEngine } from '../customxml/xpath.mjs';
import { parsePrefixMappings } from '../customxml/xpath.mjs';
import { CORE_PROPERTIES_STOREITEMID, EXTENDED_PROPERTIES_STOREITEMID, StoreItemResolver, isPropertiesStoreItem, normaliseStoreItemId, type StoreItemCandidate } from '../customxml/storeItems.mjs';
import { parseXml, decodeXmlText } from '../../xml/dom.mjs';
import { walk } from '@docx4j/generated-objects-ts/builders/wml';
import type { NoteType } from './TextResult.mjs';

/** As much of a `WordprocessingMLPackage` as a reading uses, so that this module does not import the package. */
export interface TextPackage {
  getMainDocumentPart(): MainLike;
  getPropertyResolver(): Promise<PropertyResolver>;
  getNumberingEmulator(): Promise<unknown>;
  readonly propertyResolverOrUndefined?: PropertyResolver;
  readonly docPropsCorePart?: XmlPart<core.CoreProperties>;
  readonly docPropsExtendedPart?: XmlPart<extended.Properties>;
  readonly docPropsCustomPart?: XmlPart<custom.Properties>;
  readonly customXmlDataStorageParts: ReadonlyMap<string, { parseDocument(): Promise<Document>; readonly partName: { readonly name: string } }>;
  xpathEngine: XPathEngine;
  /** The names the container held twice, read from the first under `duplicateEntries: 'first'` (CR-024). */
  readonly duplicateEntries?: readonly { readonly entry: string; readonly count: number }[];
}

/** As much of a `MainDocumentPart` as a reading uses. */
export interface MainLike extends XmlPart<wml.Document> {
  readonly styleDefinitionsPart?: Part;
  readonly numberingDefinitionsPart?: XmlPart<wml.Numbering> & NumberingPartLike & { getDefinitions(): Promise<NumberingDefinitions> };
  readonly documentSettingsPart?: XmlPart<wml.CTSettings>;
  readonly footnotesPart?: XmlPart<wml.CTFootnotes>;
  readonly endnotesPart?: XmlPart<wml.CTEndnotes>;
  readonly commentsPart?: XmlPart<wml.Comments>;
  readonly commentsExtendedPart?: XmlPart<w15.CTCommentsEx>;
  readonly commentsIdsPart?: XmlPart<w16cid.CTCommentsIds>;
  readonly commentsExtensiblePart?: XmlPart<w16cex.CTCommentsExtensible>;
  readonly glossaryDocumentPart?: XmlPart<wml.GlossaryDocument>;
  readonly headerParts: XmlPart<wml.Hdr>[];
  readonly footerParts: XmlPart<wml.Ftr>[];
}

/**
 * A part a reading may need that it could not read (CR-012 section 28, WD REQ-801): its role, by what it supplies. Each is
 * raised as `part.not_read`, naming the part, where the reading needs it.
 */
export type UnreadPart = 'styles' | 'numbering' | 'settings' | 'core' | 'app' | 'custom';

/** Where a package keeps its numbering part: what `part.absent` names where it has none (WD 0.76). */
const NUMBERING_PART_NAME = '/word/numbering.xml';

/** `readContents()`, undefined for a part that cannot be read. */
async function readQuietly<T>(part: XmlPart<T> | undefined): Promise<T | undefined> {
  if (!part) return undefined;
  try {
    return await part.readContents();
  } catch {
    return undefined;
  }
}

/**
 * The trees one reading walks, each read **privately** where its part was not unmarshalled
 * (`XmlPart.readContents`), so that reading a document costs no part its byte-for-byte round trip
 * (REQ-204), and the live tree where it was. Every story is read, so that the views' plan covers them
 * all ("the view is the package's"); the glossary only when it is included.
 */
export class PackageTrees {
  main: { part: MainLike; document: wml.Document } | undefined;
  footnotes: { part: Part; tree: wml.CTFootnotes } | undefined;
  endnotes: { part: Part; tree: wml.CTEndnotes } | undefined;
  comments: { part: Part; tree: wml.Comments } | undefined;
  commentsExtended: w15.CTCommentsEx | undefined;
  /**
   * Each comment's `w16cex:dateUtc`, by its last paragraph's `w14:paraId` (upper case): the comments ids part gives the
   * paragraph's durable id, the comments extensible part the durable id's date (REQ-702's offset; docx4j's `commentDatesUtc`).
   */
  readonly commentDatesUtc = new Map<string, wml.XmlCalendar>();
  /** Each header and footer part related from the main document part, with its content. */
  readonly headersAndFooters = new Map<Part, wml.Hdr | wml.Ftr>();
  glossary: { part: Part; tree: wml.GlossaryDocument } | undefined;
  /** Each story part that could not be read, with the reason: a reading taking it raises `content.not_read` (REQ-801). */
  readonly unreadable = new Map<Part, unknown>();
  settings: wml.CTSettings | undefined;
  /** The document properties under `core`, `app` and `custom` (REQ-515). */
  properties: Record<string, unknown> | undefined;
  /** The resolver, where the package has a styles part that can be read (docx4j makes none without one). */
  resolver: PropertyResolver | undefined;
  /**
   * What the numbering resolves a style's `w:numPr` through: the package's resolver, else, where a styles or numbering part
   * cannot be read, the reading's own without it (CR-012 section 28); undefined where the package has neither part.
   */
  private numberingResolver: PropertyResolver | undefined;
  numbering: NumberingDefinitions | undefined;
  /**
   * The numbering this reading counts by: the numbering part's emulator, one of the reading's own where the styles cannot
   * be read, or, where the numbering part is absent or cannot be read, one over Word's default definition, made at the
   * first numbered paragraph (`numberingFor`).
   */
  emulator: Emulator | undefined;
  /** Where the numbering part is absent or cannot be read: the warning the first numbered paragraph raises, and the part named. */
  numberingMissing: { kind: 'part.absent' | 'part.not_read'; partName: string } | undefined;
  /** The parts this reading could not read that it may need, by role, each with its part's name (CR-012 section 28). */
  readonly notRead = new Map<UnreadPart, string>();
  /** The custom XML stores and properties parts a binding could not be read against, by normalised store item id: their part names. */
  private readonly unreadStores = new Map<string, string>();
  /** The core and application properties' DOMs, by their normalised store item ids. */
  private readonly stores = new Map<string, Document | undefined>();
  /** The custom XML parts' DOMs, in the package's order: what a binding resolves to (CR-020). */
  private readonly customStores: StoreItemCandidate<Document | undefined>[] = [];
  private readonly storeItems = new StoreItemResolver<Document | undefined>(() => this.customStores);
  private xpath: XPathEngine | undefined;
  private anchored: Set<string> | undefined;

  private constructor(readonly pkg: TextPackage | undefined) {}

  /** The trees of a package. */
  static async read(pkg: TextPackage): Promise<PackageTrees> {
    const trees = new PackageTrees(pkg);
    let main: MainLike | undefined;
    try {
      main = pkg.getMainDocumentPart();
    } catch {
      main = undefined;
    }
    if (main) {
      trees.main = { part: main, document: await main.readContents() };
      await trees.readStylesAndNumbering(pkg, main);
      trees.settings = await trees.readOrNote(main.documentSettingsPart, 'settings');
      const footnotes = await trees.readStory(main.footnotesPart);
      if (footnotes && main.footnotesPart) trees.footnotes = { part: main.footnotesPart, tree: footnotes };
      const endnotes = await trees.readStory(main.endnotesPart);
      if (endnotes && main.endnotesPart) trees.endnotes = { part: main.endnotesPart, tree: endnotes };
      const comments = await trees.readStory(main.commentsPart);
      if (comments && main.commentsPart) trees.comments = { part: main.commentsPart, tree: comments };
      trees.commentsExtended = await readQuietly(main.commentsExtendedPart);   // passed over silently: CR-012 section 28's D3 (a), not ruled
      await trees.readCommentDates(main);
      for (const part of [...main.headerParts, ...main.footerParts] as XmlPart<wml.Hdr | wml.Ftr>[]) {
        const content = await trees.readStory(part);
        if (content) trees.headersAndFooters.set(part, content);
      }
    }
    trees.properties = await properties(pkg, (part, role) => trees.readOrNote(part, role));
    await trees.readStores(pkg);
    return trees;
  }

  /**
   * The resolver and the numbering (CR-012 section 28, WD REQ-801). The package's own where they can be built; where a
   * styles or numbering part cannot be read, a resolver of the reading's own without it, as for a package that lacks the
   * part, the part kept in `notRead`. The package's resolver is never set from it, so the content API still refuses the
   * part (CR-001 section 26), and a part put right is read again by the next reading. Where the styles cannot be read the
   * numbering's definitions are read directly and counted by an emulator of the reading's own, so that nothing is kept on
   * the numbering part. A numbering part absent or unread is Word's default definition (`numberingFor`, WD 0.76).
   */
  private async readStylesAndNumbering(pkg: TextPackage, main: MainLike): Promise<void> {
    const stylesPart = main.styleDefinitionsPart;
    const numberingPart = main.numberingDefinitionsPart;
    if (stylesPart || numberingPart) {
      let packages = true;
      try {
        this.numberingResolver = await pkg.getPropertyResolver();
      } catch {
        packages = false;
        try {
          this.numberingResolver = await PropertyResolver.create(pkg as unknown as Parameters<typeof PropertyResolver.create>[0], {
            onUnreadable: (role) => {
              const part = role === 'styles' ? stylesPart : numberingPart;
              if (part) this.notRead.set(role, part.partName.name);
            },
          });
        } catch {
          this.numberingResolver = undefined;                          // hidden text and indents not resolved through styles
        }
      }
      if (stylesPart && !this.notRead.has('styles')) this.resolver = this.numberingResolver;
      if (numberingPart && !this.notRead.has('numbering')) {
        try {
          if (packages) {
            await pkg.getNumberingEmulator();
            this.numbering = await numberingPart.getDefinitions();
            this.emulator = numberingPart.getEmulator();
          } else {
            const resolver = this.numberingResolver;
            this.numbering = new NumberingDefinitions(await numberingPart.readContents(), (id) => resolver?.getStyle(id));
            this.emulator = new Emulator(this.numbering, resolver);
          }
        } catch {
          this.numbering = undefined;
          this.emulator = undefined;
          this.notRead.set('numbering', numberingPart.partName.name);
        }
      }
    }
    if (this.emulator === undefined) {
      this.numberingMissing = numberingPart
        ? { kind: 'part.not_read', partName: numberingPart.partName.name }
        : { kind: 'part.absent', partName: NUMBERING_PART_NAME };
    }
  }

  /** `commentDatesUtc`, from the comments ids and extensible parts, read privately; neither read, as D3 (a) leaves them, nothing. */
  private async readCommentDates(main: MainLike): Promise<void> {
    const ids = await readQuietly(main.commentsIdsPart);
    const extensible = await readQuietly(main.commentsExtensiblePart);
    if (!ids || !extensible) return;
    const byDurableId = new Map<string, wml.XmlCalendar>();
    for (const ce of extensible.commentExtensible ?? []) {
      if (ce.durableId && ce.dateUtc) byDurableId.set(ce.durableId.toUpperCase(), ce.dateUtc as wml.XmlCalendar);
    }
    for (const ci of ids.commentId ?? []) {
      if (!ci.paraId || !ci.durableId) continue;
      const utc = byDurableId.get(ci.durableId.toUpperCase());
      if (utc) this.commentDatesUtc.set(ci.paraId.toUpperCase(), utc);
    }
  }

  /** `readContents()`, undefined for a part that cannot be read, which is then kept in `notRead` under its role. */
  private async readOrNote<T>(part: XmlPart<T> | undefined, role: UnreadPart): Promise<T | undefined> {
    if (!part) return undefined;
    try {
      return await part.readContents();
    } catch {
      this.notRead.set(role, part.partName.name);
      return undefined;
    }
  }

  /**
   * The numbering a numbered paragraph is counted by. Where the numbering part is absent or cannot be read, Word's default
   * definition (WD 0.76, REQ-510), made here once, at the first numbered paragraph, and never written to the package: one
   * abstract definition of nine decimal levels, each `w:start` 1 and `w:lvlText` `%{level+1}.`, and one instance per
   * `w:numId` other than 0 the document names, plain where a paragraph style's `w:numPr` names it, else with a
   * `w:startOverride` of 1 at every level, so that the direct paragraphs restart at their first and then share the style's
   * count.
   */
  numberingFor(): Emulator | undefined {
    if (this.emulator !== undefined || this.numberingMissing === undefined) return this.emulator;
    const resolver = this.numberingResolver;
    const plain = new Set<number>();
    for (const style of resolver?.getStylesInOrder() ?? []) {
      const numId = Number(style.type === 'paragraph' ? style.pPr?.numPr?.numId?.val : undefined);
      if (Number.isInteger(numId) && numId !== 0) plain.add(numId);
    }
    const direct = new Set<number>();
    const roots: unknown[] = [this.main?.document, ...this.headersAndFooters.values(), this.footnotes?.tree, this.endnotes?.tree,
      this.comments?.tree, this.glossary?.tree];
    for (const root of roots) {
      if (root === undefined) continue;
      walk(root, (v) => {
        if ((v as { TYPE_NAME?: string }).TYPE_NAME !== 'org_docx4j_wml.PPrBase.NumPr') return;
        const numId = Number((v as wml.PPrBase.NumPr).numId?.val);
        if (Number.isInteger(numId) && numId !== 0 && !plain.has(numId)) direct.add(numId);
      });
    }
    const levels = [0, 1, 2, 3, 4, 5, 6, 7, 8];
    const num = (numId: number, restart: boolean): wml.Numbering.Num => f.createNumberingNum({
      numId,
      abstractNumId: f.createNumberingNumAbstractNumId({ val: 0 }),
      lvlOverride: restart
        ? levels.map((ilvl) => f.createNumberingNumLvlOverride({ ilvl, startOverride: f.createNumberingNumLvlOverrideStartOverride({ val: 1 }) }))
        : [],
    });
    const tree = f.createNumbering({
      abstractNum: [f.createNumberingAbstractNum({
        abstractNumId: 0,
        lvl: levels.map((ilvl) => f.createLvl({
          ilvl, start: f.createLvlStart({ val: 1 }), numFmt: f.createNumFmt({ val: 'decimal' }), lvlText: f.createLvlLvlText({ val: `%${ilvl + 1}.` }),
        })),
      })],
      num: [...[...plain].map((n) => num(n, false)), ...[...direct].map((n) => num(n, true))],
    });
    this.numbering = new NumberingDefinitions(tree, (id) => resolver?.getStyle(id));
    this.emulator = new Emulator(this.numbering, resolver);
    return this.emulator;
  }

  /** Trees with no package: content read as it stands (docx4j's planless `Texts.of(content)`). */
  static none(): PackageTrees {
    return new PackageTrees(undefined);
  }

  /** The glossary document, read when `include: glossary` first asks for it. */
  async readGlossary(): Promise<void> {
    const part = this.main?.part.glossaryDocumentPart;
    if (this.glossary || !part) return;
    const tree = await this.readStory(part);
    if (tree) this.glossary = { part, tree };
  }

  /** A story part's content; undefined where it has none, and where it could not be read, then kept in `unreadable`. */
  private async readStory<T>(part: XmlPart<T> | undefined): Promise<T | undefined> {
    if (!part) return undefined;
    try {
      return await part.readContents();
    } catch (e) {
      this.unreadable.set(part, e);
      return undefined;
    }
  }

  // --- bound values (REQ-509) ---------------------------------------------------------------------

  private async readStores(pkg: TextPackage): Promise<void> {
    let context: Node | undefined;
    for (const [id, part] of pkg.customXmlDataStorageParts) {
      let doc: Document;
      try {
        doc = await part.parseDocument();
        context ??= doc;
      } catch {
        // read as absent (CR-012 section 28, D3; WD 0.69): no candidate for a binding, which REQ-509's repair resolves to the
        // store its XPath fits, else its surface; its id kept, for the part.not_read a binding naming it raises
        this.unreadStores.set(normaliseStoreItemId(id), part.partName.name);
        continue;
      }
      this.customStores.push({ part: doc, id, partName: part.partName.name, root: doc.documentElement, selects: (xpath, mappings) => this.selectIn(doc, xpath, mappings).length > 0 });
    }
    for (const [id, part] of [[CORE_PROPERTIES_STOREITEMID, pkg.docPropsCorePart], [EXTENDED_PROPERTIES_STOREITEMID, pkg.docPropsExtendedPart]] as const) {
      if (!part) continue;
      try {
        // the part's XML as it would be saved, parsed apart from it: the part keeps its round trip
        const doc = parseXml(decodeXmlText(await part.getBytes()));
        this.stores.set(normaliseStoreItemId(id), doc);
        context ??= doc;
      } catch {
        this.stores.set(normaliseStoreItemId(id), undefined);
        this.unreadStores.set(normaliseStoreItemId(id), part.partName.name);
      }
    }
    if (context === undefined) return;
    try {
      await pkg.xpathEngine.ready(context);
      this.xpath = pkg.xpathEngine;
    } catch {
      this.xpath = undefined;                                         // no XPath here: every control's surface is read
    }
  }

  /**
   * A data binding's value, docx4j `BoundValues`: the binding's XPath, with its own prefix mappings,
   * against the custom XML part its `w:storeItemID` names (or the core or application properties), or,
   * where it names none, the part its XPath fits, as Word repairs it (CR-020; nothing is written back);
   * the string value of the first node selected. Undefined where the binding does not resolve.
   */
  boundValue(binding: wml.CTDataBinding): string | undefined {
    if (!binding.xpath || !binding.storeItemID || !this.xpath) return undefined;
    const doc = isPropertiesStoreItem(binding.storeItemID)
      ? this.stores.get(normaliseStoreItemId(binding.storeItemID))
      : this.storeItems.resolve(binding.storeItemID, binding.xpath, binding.prefixMappings)?.part;
    try {
      const node = this.selectIn(doc, binding.xpath, binding.prefixMappings)[0];
      if (!node) return undefined;
      return node.nodeType === 2 ? (node as Attr).value : (node.textContent ?? '');
    } catch {
      return undefined;
    }
  }

  /**
   * The name of the part a binding names that could not be read: a custom XML store, or the core or application
   * properties (CR-012 section 28, D3); undefined where it names a part that was read, or none.
   */
  unreadStoreOf(binding: wml.CTDataBinding): string | undefined {
    return binding.storeItemID ? this.unreadStores.get(normaliseStoreItemId(binding.storeItemID)) : undefined;
  }

  /** The nodes an XPath selects in a store, with a binding's prefix mappings (the default namespace's left out). */
  private selectIn(doc: Document | undefined, xpath: string, prefixMappings: string | undefined): Node[] {
    if (!doc || !this.xpath) return [];
    const mappings = parsePrefixMappings(prefixMappings);
    delete mappings[''];
    return this.xpath.select(xpath, doc, mappings);
  }

  // --- the comments anchored in the document (REQ-504) -----------------------------------------------

  /**
   * The ids of the comments anchored in the document: a `w:commentReference` or a `w:commentRangeStart`
   * in the main document, a header or footer, or a note, in any branch.
   */
  anchoredComments(): Set<string> {
    if (this.anchored) return this.anchored;
    const ids = new Set<string>();
    const roots: unknown[] = [this.main?.document, ...this.headersAndFooters.values(), this.footnotes?.tree, this.endnotes?.tree];
    for (const root of roots) {
      if (root === undefined) continue;
      walk(root, (v) => {
        const tn = (v as { TYPE_NAME?: string }).TYPE_NAME;
        if (tn === 'org_docx4j_wml.R.CommentReference' || tn === 'org_docx4j_wml.CommentRangeStart') {
          const id = (v as { id?: unknown }).id;
          if (id !== undefined && id !== null) ids.add(String(id));
        }
      });
    }
    return (this.anchored = ids);
  }

  /** Every story's body, over these trees, for the views' plan: the main document, each note and comment, the headers and footers. */
  stories(): Body[] {
    const pkg = this.pkg as unknown as ConstructorParameters<typeof Body>[3];
    const out: Body[] = [];
    if (this.main?.document.body) out.push(new Body(this.main.part, this.main.document.body as { content?: never[] }, 'body', pkg));
    for (const [holder, kind] of [[this.footnotes, 'footnote'], [this.endnotes, 'endnote']] as const) {
      if (!holder) continue;
      const notes = (kind === 'footnote' ? (holder.tree as wml.CTFootnotes).footnote : (holder.tree as wml.CTEndnotes).endnote) ?? [];
      for (const note of notes) out.push(new Body(holder.part as XmlPart<unknown>, note as { content?: never[] }, `${kind}:${note.id}`, pkg));
    }
    for (const comment of this.comments?.tree.comment ?? []) {
      out.push(new Body(this.comments!.part as XmlPart<unknown>, comment as { content?: never[] }, `comment:${comment.id}`, pkg));
    }
    for (const [part, content] of this.headersAndFooters) out.push(new Body(part as XmlPart<unknown>, content as { content?: never[] }, 'story', pkg));
    return out;
  }

  /** The part whose tree holds an object, by its `PARENT`s, or undefined. */
  partOf(object: unknown): Part | undefined {
    let root: unknown = object;
    for (let o = object as { PARENT?: unknown } | undefined; o; o = o.PARENT as { PARENT?: unknown } | undefined) root = o;
    if (this.main && (root === this.main.document || root === this.main.document.body)) return this.main.part;
    for (const [part, content] of this.headersAndFooters) if (root === content) return part;
    if (this.footnotes && root === this.footnotes.tree) return this.footnotes.part;
    if (this.endnotes && root === this.endnotes.tree) return this.endnotes.part;
    if (this.comments && root === this.comments.tree) return this.comments.part;
    if (this.glossary && root === this.glossary.tree) return this.glossary.part;
    return undefined;
  }
}

/**
 * What a story needs from the rest of the document (CR-012 section 4.4): the note references' order, for
 * `[n]`; the numbering states per story, for list labels and `value`; the bookmarks, for `fields: update`;
 * and, made once, each view's plan of what Accept All or Reject All would leave. A call given a context
 * uses it and adds to it, so that the editor's calls over one part at a time, in the order the
 * specification reads the stories (REQ-501), number the notes and count the lists as one reading of the
 * whole package does; `TextResult.concat` assembles their results.
 *
 * A context is a snapshot of the package as it was read (`Texts.contextOf`): after the package changes,
 * make another. It is for one reading, in one view: the calls under it take the same options, as the parts
 * of one `Texts.of(pkg)` would.
 */
export class TextContext {
  /** Notes referenced, in reference order: id to number, 0 for a note with a custom mark. */
  readonly footnotes = new Map<string, number>();
  readonly endnotes = new Map<string, number>();
  private footnoteCount = 0;
  private endnoteCount = 0;
  /** The numbering states, one per story (`Emulator`). */
  readonly numberingStates = new NumberingStates();
  /** The `w:lastRenderedPageBreak`s of the main story so far (REQ-708). */
  pageMarkers = 0;
  /** The bookmarks' text, by lower-cased name (REQ-506's REF, CR-012 phase 5a). */
  readonly bookmarks = new Map<string, string>();
  /** Whether the main story has been read under this context, which numbers the notes. */
  mainRead = false;
  /**
   * The view the context's first call read in: its note numbers and counts are that view's, so a call in
   * another view is refused (`Texts.of`).
   */
  view: 'accepted' | 'original' | 'markup' | undefined;
  private readonly plans = new Map<boolean, ResolvedView>();
  /** Per container, then per options, what the blocks begin inside (`Extractor.fieldsAtBlocks`), made once; shared with siblings. */
  private starts = new WeakMap<object, Map<string, unknown>>();

  constructor(readonly trees: PackageTrees) {}

  /** What `make` gives for a container's blocks under options keyed `key`, made once in this context (CR-012 section 17.13). */
  blockStarts<T>(container: object, key: string, make: () => T): T {
    let byKey = this.starts.get(container);
    if (!byKey) this.starts.set(container, byKey = new Map());
    if (!byKey.has(key)) byKey.set(key, make());
    return byKey.get(key) as T;
  }

  /** The package this context reads. */
  get pkg(): TextPackage | undefined {
    return this.trees.pkg;
  }

  /** A note's number, given at its first reference: in reference order, footnotes and endnotes apart, 0 for a custom mark (REQ-503). */
  noteNumber(type: NoteType, id: string, customMark: boolean): number {
    const numbers = type === 'footnote' ? this.footnotes : this.endnotes;
    let n = numbers.get(id);
    if (n === undefined) {
      n = customMark ? 0 : type === 'footnote' ? ++this.footnoteCount : ++this.endnoteCount;
      numbers.set(id, n);
    }
    return n;
  }

  /** The plan of Accept All (`accepted`) or Reject All, over every story, made once. */
  plan(accepted: boolean): ResolvedView {
    let plan = this.plans.get(accepted);
    if (!plan) {
      const stories = this.trees.stories();
      plan = accepted ? ResolvedView.accepted(stories) : ResolvedView.original(stories);
      this.plans.set(accepted, plan);
    }
    return plan;
  }

  /** A context over the same trees and plans, its own notes and states: for a reading whose blocks are not the result's. */
  sibling(): TextContext {
    const c = new TextContext(this.trees);
    for (const [k, v] of this.plans) c.plans.set(k, v);
    c.starts = this.starts;
    return c;
  }

  /** Adopts the note numbers another context gave (`sibling`'s, for a notes part read before its story). */
  adoptNotes(other: TextContext): void {
    for (const [id, n] of other.footnotes) if (!this.footnotes.has(id)) this.footnotes.set(id, n);
    for (const [id, n] of other.endnotes) if (!this.endnotes.has(id)) this.endnotes.set(id, n);
    this.footnoteCount = Math.max(this.footnoteCount, other.footnoteCount);
    this.endnoteCount = Math.max(this.endnoteCount, other.endnoteCount);
  }
}

// --- the document properties (REQ-515; docx4j `Metadata`) --------------------------------------------

/** How `properties` reads a part: undefined where it cannot be, the part noted under its role (CR-012 section 28.5). */
type PartReader = <T>(part: XmlPart<T> | undefined, role: UnreadPart) => Promise<T | undefined>;

/** The properties under `core`, `app` and `custom`, each where the package has that part and it can be read. */
async function properties(pkg: TextPackage, read: PartReader): Promise<Record<string, unknown>> {
  const out: Record<string, unknown> = {};
  const c = await read(pkg.docPropsCorePart, 'core');
  if (c) out.core = coreProperties(c);
  const a = await read(pkg.docPropsExtendedPart, 'app');
  if (a) out.app = appProperties(a);
  const x = await read(pkg.docPropsCustomPart, 'custom');
  if (x) out.custom = customProperties(x);
  return out;
}

function put(m: Record<string, unknown>, name: string, value: unknown): void {
  if (value === undefined || value === null || value === '') return;
  m[name] = value;
}

function literal(o: unknown): string | undefined {
  const v = typeof o === 'object' && o !== null && 'value' in o && 'name' in o ? (o as { value: unknown }).value : o;
  const content = (v as { content?: string[] } | undefined)?.content;
  return content === undefined || content.length === 0 ? undefined : content.join('');
}

/** The core properties in the schema's order. */
function coreProperties(c: core.CoreProperties): Record<string, unknown> {
  const m: Record<string, unknown> = {};
  put(m, 'title', literal(c.title));
  put(m, 'subject', literal(c.subject));
  put(m, 'creator', literal(c.creator));
  put(m, 'keywords', c.keywords);
  put(m, 'description', literal(c.description));
  put(m, 'category', c.category);
  put(m, 'contentStatus', c.contentStatus);
  put(m, 'language', literal(c.language));
  put(m, 'identifier', literal(c.identifier));
  put(m, 'version', c.version);
  put(m, 'revision', c.revision);
  put(m, 'lastModifiedBy', c.lastModifiedBy);
  put(m, 'created', literal(c.created));
  put(m, 'modified', literal(c.modified));
  put(m, 'lastPrinted', calendarText(c.lastPrinted));
  return m;
}

/** The application properties that describe the document; the statistics, which depend on layout, left out. */
function appProperties(p: extended.Properties): Record<string, unknown> {
  const m: Record<string, unknown> = {};
  put(m, 'Application', p.application);
  put(m, 'AppVersion', p.appVersion);
  put(m, 'Company', p.company);
  put(m, 'Manager', p.manager);
  put(m, 'Template', p.template);
  put(m, 'HyperlinkBase', p.hyperlinkBase);
  put(m, 'TotalTime', p.totalTime);
  put(m, 'DocSecurity', p.docSecurity);
  return m;
}

/** Each custom property's typed value, in the part's order; values of other types left out. */
function customProperties(p: custom.Properties): Record<string, unknown> {
  const m: Record<string, unknown> = {};
  for (const property of p.property ?? []) {
    if (property.name === undefined) continue;
    put(m, property.name, customValue(property));
  }
  return m;
}

function customValue(p: custom.Properties.Property): unknown {
  for (const key of ['lpwstr', 'lpstr', 'bstr', 'bool', 'i1', 'i2', 'i4', 'i8', '_int', 'ui1', 'ui2', 'ui4', 'ui8', 'uint', 'r4', 'r8', 'decimal'] as const) {
    const v = p[key];
    if (v !== undefined && v !== null) return v;
  }
  if (p.filetime !== undefined) return calendarText(p.filetime);
  if (p.date !== undefined) return calendarText(p.date);
  if (p.cy !== undefined) return p.cy;
  if (p.error !== undefined) return p.error;
  return undefined;
}

/** An XML calendar as its lexical form. */
export function calendarText(date: unknown): string | undefined {
  if (date === undefined || date === null) return undefined;
  if (typeof date === 'string') return date;
  const c = date as { year?: number; month?: number; day?: number; hour?: number; minute?: number; second?: number; fractionalSecond?: number; timezone?: number };
  if (typeof c.year !== 'number') return String(date);
  const p2 = (n: number | undefined): string => String(n ?? 0).padStart(2, '0');
  let out = `${String(c.year).padStart(4, '0')}-${p2(c.month)}-${p2(c.day)}`;
  if (c.hour !== undefined) {
    out += `T${p2(c.hour)}:${p2(c.minute)}:${p2(c.second)}`;
    if (c.fractionalSecond) out += String(c.fractionalSecond).replace(/^0/, '');
  }
  if (c.timezone !== undefined) {
    if (c.timezone === 0) out += 'Z';
    else {
      const sign = c.timezone < 0 ? '-' : '+';
      const minutes = Math.abs(c.timezone);
      out += `${sign}${p2(Math.floor(minutes / 60))}:${p2(minutes % 60)}`;
    }
  }
  return out;
}
