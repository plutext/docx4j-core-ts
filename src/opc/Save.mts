import type { PartSink } from './PartStore.mjs';
import { CONTENT_TYPES_NAME } from './PartName.mjs';
import { ContentTypes } from './ContentTypes.mjs';
import { Docx4JException, MisplacedContentException, SaveLossException, type ContentLoss } from './exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import { RelationshipsPart, isExternal } from '../parts/RelationshipsPart.mjs';
import type { Part } from '../parts/Part.mjs';
import { XmlPart, valueOutOfRange, misplacedContent } from '../parts/XmlPart.mjs';
import type { ValueOutOfRangeWarning, MisplacedContentWarning } from './mce/preprocessor.mjs';
import type { ContentSkippedWarning } from './skipped.mjs';
import type { Normalisation } from '@docx4j/generated-objects-ts';
import type { OpcPackage } from '../packages/OpcPackage.mjs';
import { encodeText } from '../xml/dom.mjs';

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
  await saver.settleLosses(); // before anything is written, whatever the sink
  sink.put(CONTENT_TYPES_NAME, encodeText(pkg.contentTypeManager.toXml()), { contentType: ContentTypes.CONTENT_TYPES_PART });
  await saver.saveRels(pkg.relationshipsPart);
  const result = await sink.finish();
  pkg.lastSave = { normalised: saver.normalised, outOfRange: saver.outOfRange, misplaced: saver.misplaced, dropped: saver.dropped, lost: saver.lost };
  return result;
}

class Saver {
  private readonly handled = new Set<string>();
  readonly normalised: ValueNormalised[] = [];
  readonly outOfRange: ValueOutOfRangeWarning[] = [];
  readonly misplaced: MisplacedContentWarning[] = [];
  readonly dropped: ContentSkippedWarning[] = [];
  readonly lost: ContentSkippedWarning[] = [];
  /** Parts holding skipped content, read and not changed: written from their source (CR-025 section 2.2). */
  private readonly copied = new Set<string>();

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
            if (await part.isUnchangedSinceRead()) this.copied.add(part.partName.key);
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
    if (losses.length > 0) throw new SaveLossException(losses);
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
