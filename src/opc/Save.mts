import type { PartSink } from './PartStore.mjs';
import { CONTENT_TYPES_NAME } from './PartName.mjs';
import { ContentTypes } from './ContentTypes.mjs';
import { Docx4JException } from './exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import { RelationshipsPart, isExternal } from '../parts/RelationshipsPart.mjs';
import type { Part } from '../parts/Part.mjs';
import { valueOutOfRange } from '../parts/XmlPart.mjs';
import type { ValueOutOfRangeWarning } from './mce/preprocessor.mjs';
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
 * What the last save of a package did beyond writing it: each value it normalised (CR-015 D8), and each integer
 * attribute value it wrote that Word 365 reads differently, written as held (CR-017 P3, E3), each with its path.
 */
export interface SaveReport {
  normalised: ValueNormalised[];
  outOfRange: ValueOutOfRangeWarning[];
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
  sink.put(CONTENT_TYPES_NAME, encodeText(pkg.contentTypeManager.toXml()), { contentType: ContentTypes.CONTENT_TYPES_PART });
  await saver.saveRels(pkg.relationshipsPart);
  const result = await sink.finish();
  pkg.lastSave = { normalised: saver.normalised, outOfRange: saver.outOfRange };
  return result;
}

class Saver {
  private readonly handled = new Set<string>();
  readonly normalised: ValueNormalised[] = [];
  readonly outOfRange: ValueOutOfRangeWarning[] = [];

  constructor(private readonly pkg: OpcPackage, private readonly sink: PartSink<unknown>) {}

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
    let bytes: Uint8Array;
    try {
      bytes = await part.getBytes({
        onNormalised: (n) => this.normalised.push(valueNormalised(part.partName.name, n)),
        onOutOfRange: (r) => this.outOfRange.push(valueOutOfRange(r, part.partName.name)),
      });
    } catch (e) {
      throw new Docx4JException(`Problem saving part ${part.partName}`, { cause: e });
    }
    this.sink.put(part.partName.storeName, bytes, { compress: part.compress, contentType: part.contentType });
    const rrp = part.relationshipsPart;
    if (rrp && rrp.size > 0) await this.saveRels(rrp);
  }
}
