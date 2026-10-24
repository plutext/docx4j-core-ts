import type { PartSink } from './PartStore.mjs';
import { CONTENT_TYPES_NAME } from './PartName.mjs';
import { ContentTypes } from './ContentTypes.mjs';
import { Docx4JException } from './exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import { RelationshipsPart, isExternal } from '../parts/RelationshipsPart.mjs';
import type { Part } from '../parts/Part.mjs';
import type { Normalisation } from '@docx4j/generated-objects-ts';
import type { OpcPackage } from '../packages/OpcPackage.mjs';
import { encodeText } from '../xml/dom.mjs';

/**
 * A value a save wrote in its normal form, or left out (CR-015 D2 and D8; the objects facade's normaliser, its
 * CR-013 phase C, docx4j CR-061's rule): `path` from the part's root, `before` as read, `after` as written, null
 * where the attribute or the element was left out. The code is docx4j-python CR-018's.
 */
export interface ValueNormalised {
  code: 'value.normalised';
  partName: string;
  path: string;
  before: string | undefined;
  after: string | null;
  message: string;
}

/** What the last save of a package did beyond writing it: each value it normalised (CR-015 D8). */
export interface SaveReport {
  normalised: ValueNormalised[];
}

function valueNormalised(partName: string, n: Normalisation): ValueNormalised {
  const after = n.after === null ? 'left out' : `written ${JSON.stringify(n.after)}`;
  return {
    code: 'value.normalised', partName, path: n.path, before: n.before, after: n.after,
    message: `${partName}: ${n.path} ${n.before === undefined ? '(no value)' : JSON.stringify(n.before)} ${after}, as Word reads it`,
  };
}

/**
 * Writes a package to a sink (docx4j io3.Save): `[Content_Types].xml` first, then the package
 * relationships, then every part reachable through relationships, each once. Relationship parts
 * and unmarshalled XML parts are marshalled; every other part's bytes come from its source
 * container unchanged.
 */
export async function savePackage<R>(pkg: OpcPackage, sink: PartSink<R>): Promise<R> {
  const saver = new Saver(pkg, sink);
  sink.put(CONTENT_TYPES_NAME, encodeText(pkg.contentTypeManager.toXml()), { contentType: ContentTypes.CONTENT_TYPES_PART });
  await saver.saveRels(pkg.relationshipsPart);
  const result = await sink.finish();
  pkg.lastSave = { normalised: saver.normalised };
  return result;
}

class Saver {
  private readonly handled = new Set<string>();
  readonly normalised: ValueNormalised[] = [];

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
      bytes = await part.getBytes({ onNormalised: (n) => this.normalised.push(valueNormalised(part.partName.name, n)) });
    } catch (e) {
      throw new Docx4JException(`Problem saving part ${part.partName}`, { cause: e });
    }
    this.sink.put(part.partName.storeName, bytes, { compress: part.compress, contentType: part.contentType });
    const rrp = part.relationshipsPart;
    if (rrp && rrp.size > 0) await this.saveRels(rrp);
  }
}
