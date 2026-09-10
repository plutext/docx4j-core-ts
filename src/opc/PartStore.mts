/**
 * Where a package's bytes come from (docx4j io3.stores.PartStore, the load half). Names are as
 * stored, without the leading '/' (`word/document.xml`, `[Content_Types].xml`, `_rels/.rels`).
 */
export interface PartStore {
  /** Names in container order. */
  partNames(): Iterable<string>;
  has(partName: string): boolean;
  /** Raw bytes of a part. Implementations may inflate lazily on first access. */
  load(partName: string): Promise<Uint8Array>;
  /** Size if known without loading, else undefined. */
  size?(partName: string): number | undefined;
}

export interface PutOptions {
  /** Deflate in a zip (default true); ignored by containers that do not compress. */
  compress?: boolean;
  /** The part's content type, needed by containers that carry it per part (flat OPC). */
  contentType?: string;
}

/** Where a package's bytes go (docx4j PartStore, the save half). `R` is what `finish` produces. */
export interface PartSink<R = unknown> {
  put(partName: string, bytes: Uint8Array, options?: PutOptions): void;
  /** Completes the container. */
  finish(): Promise<R>;
}

/** An in-memory container: new packages, tests, and the result of `MemoryPartSink`. */
export class MemoryPartStore implements PartStore, PartSink<MemoryPartStore> {
  private readonly parts = new Map<string, Uint8Array>();
  /** Content types as put, for stores filled through the sink side. */
  readonly contentTypes = new Map<string, string>();

  constructor(entries?: Iterable<[string, Uint8Array]>) {
    if (entries) for (const [name, bytes] of entries) this.parts.set(normalize(name), bytes);
  }

  partNames(): Iterable<string> {
    return this.parts.keys();
  }

  has(partName: string): boolean {
    return this.parts.has(normalize(partName));
  }

  async load(partName: string): Promise<Uint8Array> {
    const bytes = this.parts.get(normalize(partName));
    if (bytes === undefined) throw new Error(`No part ${partName}`);
    return bytes;
  }

  size(partName: string): number | undefined {
    return this.parts.get(normalize(partName))?.length;
  }

  put(partName: string, bytes: Uint8Array, options?: PutOptions): void {
    const name = normalize(partName);
    this.parts.set(name, bytes);
    if (options?.contentType !== undefined) this.contentTypes.set(name, options.contentType);
  }

  delete(partName: string): boolean {
    return this.parts.delete(normalize(partName));
  }

  async finish(): Promise<MemoryPartStore> {
    return this;
  }
}

/** A sink that collects into a fresh `MemoryPartStore`. */
export class MemoryPartSink implements PartSink<MemoryPartStore> {
  readonly store = new MemoryPartStore();
  put(partName: string, bytes: Uint8Array, options?: PutOptions): void {
    this.store.put(partName, bytes, options);
  }
  async finish(): Promise<MemoryPartStore> {
    return this.store;
  }
}

function normalize(name: string): string {
  return name.startsWith('/') ? name.substring(1) : name;
}
