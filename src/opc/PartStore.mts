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
  /**
   * The names the container holds more than one entry of, OPC part names compared without case (WD REQ-206; CR-024): each
   * first entry's spelling, with the number of entries, in container order. The load refuses a package with any, unless
   * `LoadOptions.duplicateEntries` is `'first'`. Every store here has it; a store without it is not checked.
   */
  duplicates?(): ReadonlyMap<string, number>;
}

/** An entry's key: its name without a leading '/', lower-cased, as `PartName`'s equality compares (CR-024 section 2.1). */
export function entryKey(name: string): string {
  return (name.startsWith('/') ? name.substring(1) : name).toLowerCase();
}

/**
 * A container's entries by key (`entryKey`), the first entry of each key kept with its spelling and its value and the
 * rest counted: what each store reads its container into, so that `has`, `load` and `size` agree and look up without
 * case, and `duplicates()` reports a name held twice (CR-024, WD REQ-206).
 */
export class EntryIndex<T> {
  private readonly byKey = new Map<string, { name: string; value: T; count: number }>();

  /** Adds an entry; a second entry of a key is counted, its value not kept. */
  add(name: string, value: T): void {
    const key = entryKey(name);
    const entry = this.byKey.get(key);
    if (entry) entry.count++;
    else this.byKey.set(key, { name: name.startsWith('/') ? name.substring(1) : name, value, count: 1 });
  }

  /** Sets an entry's value, replacing the one its key holds (a sink's write), its first spelling kept. */
  set(name: string, value: T): void {
    const entry = this.byKey.get(entryKey(name));
    if (entry) entry.value = value;
    else this.add(name, value);
  }

  get(name: string): T | undefined {
    return this.byKey.get(entryKey(name))?.value;
  }

  has(name: string): boolean {
    return this.byKey.has(entryKey(name));
  }

  delete(name: string): boolean {
    return this.byKey.delete(entryKey(name));
  }

  /** Each key's first spelling, in the order the entries were added. */
  names(): string[] {
    return [...this.byKey.values()].map((e) => e.name);
  }

  /** Each key's value (its first entry's), in the same order. */
  values(): T[] {
    return [...this.byKey.values()].map((e) => e.value);
  }

  /** The keys added more than once: the first spelling and the count. */
  duplicates(): Map<string, number> {
    const out = new Map<string, number>();
    for (const e of this.byKey.values()) if (e.count > 1) out.set(e.name, e.count);
    return out;
  }
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

/**
 * An in-memory container: new packages, tests, and the result of `MemoryPartSink`. The constructor's entries are read as a
 * container's (a second of one name counted, `duplicates()`); `put`, a sink's write, replaces.
 */
export class MemoryPartStore implements PartStore, PartSink<MemoryPartStore> {
  private readonly parts = new EntryIndex<Uint8Array>();
  /** Content types as put, for stores filled through the sink side. */
  readonly contentTypes = new Map<string, string>();

  constructor(entries?: Iterable<[string, Uint8Array]>) {
    if (entries) for (const [name, bytes] of entries) this.parts.add(normalize(name), bytes);
  }

  partNames(): Iterable<string> {
    return this.parts.names();
  }

  has(partName: string): boolean {
    return this.parts.has(partName);
  }

  async load(partName: string): Promise<Uint8Array> {
    const bytes = this.parts.get(partName);
    if (bytes === undefined) throw new Error(`No part ${partName}`);
    return bytes;
  }

  size(partName: string): number | undefined {
    return this.parts.get(partName)?.length;
  }

  duplicates(): ReadonlyMap<string, number> {
    return this.parts.duplicates();
  }

  put(partName: string, bytes: Uint8Array, options?: PutOptions): void {
    const name = normalize(partName);
    this.parts.set(name, bytes);
    if (options?.contentType !== undefined) this.contentTypes.set(name, options.contentType);
  }

  delete(partName: string): boolean {
    return this.parts.delete(partName);
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
