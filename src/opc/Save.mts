import type { PartSink } from './PartStore.mjs';
import { CONTENT_TYPES_NAME } from './PartName.mjs';
import { ContentTypes } from './ContentTypes.mjs';
import { Docx4JException, MisplacedContentException, SaveLossException, type ContentLoss } from './exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import { RelationshipsPart, isExternal } from '../parts/RelationshipsPart.mjs';
import type { Part } from '../parts/Part.mjs';
import { XmlPart, valueOutOfRange, misplacedContent } from '../parts/XmlPart.mjs';
import type { ValueOutOfRangeWarning, MisplacedContentWarning } from './mce/preprocessor.mjs';
import type { RelationshipsNotReadWarning } from './exceptions.mjs';
import type { ContentSkippedWarning } from './skipped.mjs';
import type { Normalisation } from '@docx4j/generated-objects-ts';
import type { OpcPackage } from '../packages/OpcPackage.mjs';
import { encodeText, decodeXmlText } from '../xml/dom.mjs';
import { hasDoctype } from '../xml/doctype.mjs';
import { DefaultXmlPart } from '../parts/DefaultXmlPart.mjs';

/**
 * A value a save wrote in its normal form, or left out (CR-015 D2 and D8; the objects facade's normaliser, its
 * CR-013 phase C, docx4j CR-061's rule): `path` from the part's root, `before` as read, `after` as written, null
 * where the attribute or the element was left out. The code is this package's (CR-015 D8); docx4j-python CR-018 reports
 * the same as `package.last_save.normalised`.
 */
export interface ValueNormalised {
  code: 'value.normalised';
  partName: string;
  path: string;
  before: string | undefined;
  after: string | null;
  message: string;
}

/**
 * What the last save of a package did beyond writing it: each value it normalised (CR-015 D8), each integer
 * attribute value it wrote that Word 365 reads differently, written as held (CR-017 P3, E3), with
 * `checkContentOnSave` off, each child it wrote where its container never admits one (CR-016 D2), each with its path,
 * each item the read dropped that a changed part was written without as Word's own save writes it (CR-025 section 10), and,
 * with `checkLossOnSave` off, each it was written without that Word's save keeps or nobody has measured (CR-025 D8).
 */
export interface SaveReport {
  normalised: ValueNormalised[];
  outOfRange: ValueOutOfRangeWarning[];
  misplaced: MisplacedContentWarning[];
  /** Items a changed part was written without, Word's own save dropping them too (`onSave: 'dropped'`, CR-025 D14): never refused. */
  dropped: ContentSkippedWarning[];
  /** Items a changed part was written without under `checkLossOnSave` off (`onSave: 'lost'`). */
  lost: ContentSkippedWarning[];
  /**
   * Each relationships part read short that the save wrote as read, the relationships before the break (CR-032): Word's repaired
   * form under the load's `repair: 'word'` or `'all'`, else with `checkLossOnSave` off.
   */
  relationshipsNotRead: RelationshipsNotReadWarning[];
  /**
   * Each part whose source held a DOCTYPE that this save wrote afresh without it (CR-033 D4): after an edit, every such part, so
   * that Word 365, which refuses a DOCTYPE in the main part, a header or a footer, opens the file. An untouched save writes none
   * afresh, passing each through as it came.
   */
  doctypes: string[];
}

function valueNormalised(partName: string, n: Normalisation): ValueNormalised {
  const after = n.after === null ? 'left out' : `written ${JSON.stringify(n.after)}`;
  return {
    code: 'value.normalised', partName, path: n.path, before: n.before, after: n.after,
    message: `${partName}: ${n.path} ${n.before === undefined ? '(no value)' : JSON.stringify(n.before)} ${after}, as Word reads it`,
  };
}

/**
 * `pkg.writeTimesOnSave` (CR-018; docx4j `DocPropsCorePart.writeTimes` under `docx4j.dcterms.write`): `dcterms:modified`
 * now, UTC to the second in the W3CDTF form Office writes, and `dcterms:created` too for a package created here (one with
 * no source container) that has none, its core properties part made for it. A loaded package without the part is left
 * without one, as docx4j leaves it. The part is unmarshalled for the write, so it is marshalled by this save.
 */
async function writeTimes(pkg: OpcPackage): Promise<void> {
  const isNew = pkg.sourcePartStore === undefined;
  let part = pkg.docPropsCorePart;
  if (!part) {
    if (!isNew) return;
    part = pkg.addDocPropsCorePart();
  }
  await part.getContents();
  const props = part.contents as { created?: unknown; modified?: unknown };
  const now = (): { TYPE_NAME: string; content: string[] } => ({
    TYPE_NAME: 'org_docx4j_docProps_core_dc_terms.W3CDTF', content: [new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')],
  });
  if (props.created === undefined && isNew) props.created = now();
  props.modified = now();
}

/**
 * Writes a package to a sink (docx4j io3.Save): `[Content_Types].xml` first, then the package
 * relationships, then every part reachable through relationships, each once. Relationship parts
 * and unmarshalled XML parts are marshalled; every other part's bytes come from its source
 * container unchanged.
 */
export async function savePackage<R>(pkg: OpcPackage, sink: PartSink<R>): Promise<R> {
  if (pkg.writeTimesOnSave) await writeTimes(pkg);
  const saver = new Saver(pkg, sink);
  await saver.dropDoctypes(); // first: a part it reads afresh is then one the losses are settled for
  await saver.settleLosses(); // before anything is written, whatever the sink
  sink.put(CONTENT_TYPES_NAME, encodeText(pkg.contentTypeManager.toXml()), { contentType: ContentTypes.CONTENT_TYPES_PART });
  await saver.saveRels(pkg.relationshipsPart);
  const result = await sink.finish();
  pkg.lastSave = {
    normalised: saver.normalised, outOfRange: saver.outOfRange, misplaced: saver.misplaced, dropped: saver.dropped, lost: saver.lost,
    relationshipsNotRead: saver.relationshipsNotRead, doctypes: saver.doctypes,
  };
  return result;
}

class Saver {
  private readonly handled = new Set<string>();
  readonly normalised: ValueNormalised[] = [];
  readonly outOfRange: ValueOutOfRangeWarning[] = [];
  readonly misplaced: MisplacedContentWarning[] = [];
  readonly dropped: ContentSkippedWarning[] = [];
  readonly lost: ContentSkippedWarning[] = [];
  readonly relationshipsNotRead: RelationshipsNotReadWarning[] = [];
  /** Parts holding skipped content, read and not changed: written from their source (CR-025 section 2.2). */
  private readonly copied = new Set<string>();
  readonly doctypes: string[] = [];
  /** Parts read afresh so that their DOCTYPE isn't written (CR-033 D4): never copied from their source. */
  private readonly rewritten = new Set<string>();

  constructor(private readonly pkg: OpcPackage, private readonly sink: PartSink<unknown>) {}

  /**
   * CR-025 sections 2.2, 2.3 and 10, over the parts this save reaches: one holding what its read dropped is copied if it hasn't
   * changed. If it has, what Word's own save drops too is written without and listed (`dropped`), and for the rest the save is
   * refused for all such parts at once, or, with `checkLossOnSave` off, says what went (`lost`).
   */
  async settleLosses(): Promise<void> {
    const losses: ContentLoss[] = [];
    const seen = new Set<string>();
    const visit = async (rp: RelationshipsPart): Promise<void> => {
      for (const r of rp.list) {
        if (r.type === Namespaces.HYPERLINK || isExternal(r)) continue;
        const part = this.pkg.getPart(rp.resolveTarget(r));
        if (!part || seen.has(part.partName.key)) continue;
        seen.add(part.partName.key);
        if (part instanceof XmlPart) {
          const skipped = part.skippedByRead;
          if (skipped.length > 0) {
            if (!this.rewritten.has(part.partName.key) && await part.isUnchangedSinceRead()) this.copied.add(part.partName.key);
            else {
              this.dropped.push(...skipped.filter((w) => w.onSave === 'dropped'));
              const lost = skipped.filter((w) => w.onSave === 'lost');
              if (lost.length > 0 && this.pkg.checkLossOnSave) losses.push({ partName: part.partName.name, kind: 'skipped', count: lost.length });
              else this.lost.push(...lost);
            }
          }
        }
        const rrp = part.relationshipsPart;
        if (rrp && rrp.size > 0) await visit(rrp);
      }
    };
    await visit(this.pkg.relationshipsPart);
    // a relationships part read short (CR-032, docx4j CR-075 D8): a save would drop what a lost relationship named. Under the load's
    // repair 'word' or 'all' it writes the repaired form, as Word's save of its repair does; otherwise it is refused, unless
    // checkLossOnSave is off. A part given contents since has no warning, which is how a caller who has put the relationships back says so
    const followsWord = (this.pkg.loadOptions.repair ?? 'refuse') !== 'refuse';
    for (const w of this.pkg.relationshipsNotRead) {
      if (!followsWord && this.pkg.checkLossOnSave) losses.push({ partName: w.partName, kind: 'relationships', count: w.count });
      else this.relationshipsNotRead.push(w);
    }
    if (losses.length > 0) throw new SaveLossException(losses);
  }

  /**
   * CR-033 D4 (the DTD handling policy's section 6.5): a save after an edit writes every part whose source holds a DOCTYPE afresh,
   * without it, read with its entities expanded and its defaults applied, so that the edited file opens in Word, which refuses one
   * in the main part, a header or a footer. An edit is a part written afresh: one unmarshalled or given content, or one the source
   * didn't hold. An untouched save passes every part through. A signature part is left as it came: its bytes are what it signs.
   */
  async dropDoctypes(): Promise<void> {
    const store = this.pkg.sourcePartStore;
    const parts = [...this.pkg.parts];
    const writtenAfresh = (p: Part): boolean => (p instanceof XmlPart && !p.isUntouched)
      || (p instanceof DefaultXmlPart && p.isLoaded) || (store !== undefined && !store.has(p.partName.storeName));
    if (store !== undefined && !parts.some(writtenAfresh)) return;
    for (const part of parts) {
      if (part.contentType === ContentTypes.DIGITAL_SIGNATURE_XML_SIGNATURE_PART) continue;
      if (!(part instanceof XmlPart) && !(part instanceof DefaultXmlPart)) continue;
      if (writtenAfresh(part)) {
        if ((part instanceof XmlPart && part.doctypeRead) || (part instanceof DefaultXmlPart && part.doctype)) this.doctypes.push(part.partName.name);
        continue;
      }
      if (!hasDoctype(decodeXmlText(await part.getBytes()))) continue;
      if (part instanceof XmlPart) await part.getContents();
      else { await part.parseDocument(); part.markModified(); }
      this.rewritten.add(part.partName.key);
      this.doctypes.push(part.partName.name);
    }
  }

  async saveRels(rp: RelationshipsPart): Promise<void> {
    this.sink.put(rp.partName.storeName, await rp.getBytes(), { contentType: ContentTypes.RELATIONSHIPS_PART });
    for (const r of rp.list) {
      if (r.type === Namespaces.HYPERLINK || isExternal(r)) continue;
      const part = this.pkg.getPart(rp.resolveTarget(r));
      if (!part) throw new Docx4JException(`Part ${rp.resolveTarget(r)} (target of ${r.id} in ${rp.partName}) not found`);
      await this.savePart(part);
    }
  }

  private async savePart(part: Part): Promise<void> {
    if (this.handled.has(part.partName.key)) return;
    this.handled.add(part.partName.key);
    if (this.copied.has(part.partName.key)) {
      this.sink.put(part.partName.storeName, (await (part as XmlPart<unknown>).bytesRead()) as Uint8Array, { compress: part.compress, contentType: part.contentType });
      const held = part.relationshipsPart;
      if (held && held.size > 0) await this.saveRels(held);
      return;
    }
    let bytes: Uint8Array;
    // CR-016 D2: a child its container never admits refuses the save, every one in the part listed, unless
    // pkg.checkContentOnSave is off, when each is written as held and reported; the runtime's own refusal (a string in a
    // run) names what the walk reported either way, where before it said nothing of the cause
    const misplaced: MisplacedContentWarning[] = [];
    try {
      bytes = await part.getBytes({
        onNormalised: (n) => this.normalised.push(valueNormalised(part.partName.name, n)),
        onOutOfRange: (r) => this.outOfRange.push(valueOutOfRange(r, part.partName.name)),
        onMisplaced: (r) => misplaced.push(misplacedContent(r, part.partName.name)),
      });
    } catch (e) {
      if (misplaced.length > 0 && this.pkg.checkContentOnSave) throw new MisplacedContentException(part.partName.name, misplaced, { cause: e });
      const because = misplaced.length > 0 ? `: ${misplaced.map((f) => f.message).join('; ')}` : '';
      throw new Docx4JException(`Problem saving part ${part.partName}${because}`, { cause: e });
    }
    if (misplaced.length > 0) {
      if (this.pkg.checkContentOnSave) throw new MisplacedContentException(part.partName.name, misplaced);
      this.misplaced.push(...misplaced);
    }
    this.sink.put(part.partName.storeName, bytes, { compress: part.compress, contentType: part.contentType });
    const rrp = part.relationshipsPart;
    if (rrp && rrp.size > 0) await this.saveRels(rrp);
  }
}
