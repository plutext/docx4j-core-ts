// An unzipped package on the file system (docx4j io3.stores.UnzippedPartStore). Node only: it is
// the one module here that imports `node:` builtins, and it is exported from the `./node` subpath
// only, never from `.` or `./opc`, so a browser or add-in bundle never sees them.
import { readFile, readdir, stat, mkdir, writeFile } from 'node:fs/promises';
import { join, dirname, sep } from 'node:path';
import type { PartStore, PartSink, PutOptions } from './PartStore.mjs';
import { CONTENT_TYPES_NAME } from './PartName.mjs';
import { Docx4JException } from './exceptions.mjs';

/**
 * A package spread over a directory: `word/document.xml` is the file `<dir>/word/document.xml`,
 * `[Content_Types].xml` is read from the directory like any other entry, and a part's size comes
 * from the file rather than from a central directory.
 *
 * ```ts
 * const pkg = await WordprocessingMLPackage.load(await DirectoryPartStore.open('unzipped'));
 * ```
 *
 * The directory is scanned once (`open`), because a `PartStore` lists its names synchronously.
 * A part that is never touched is copied from here to the sink byte for byte, as from a zip, so
 * zip to directory to zip reproduces every part.
 */
export class DirectoryPartStore implements PartStore {
  /** Store name (no leading '/') to [absolute file path, size]. */
  private readonly entries = new Map<string, [path: string, size: number]>();

  private constructor(readonly dir: string) {}

  /** Scans a directory and returns a store over it. */
  static async open(dir: string): Promise<DirectoryPartStore> {
    const store = new DirectoryPartStore(dir);
    await store.scan('');
    return store;
  }

  private async scan(relative: string): Promise<void> {
    const here = relative === '' ? this.dir : join(this.dir, relative);
    let listing;
    try {
      listing = await readdir(here, { withFileTypes: true });
    } catch (e) {
      throw new Docx4JException(`Cannot read package directory ${here}`, { cause: e });
    }
    // Names in a stable order, with [Content_Types].xml first as a zip written by Word has it.
    const names = listing.map((entry) => entry.name).sort(compareNames);
    for (const name of names) {
      const childRelative = relative === '' ? name : `${relative}/${name}`;
      const childPath = join(this.dir, childRelative);
      const info = await stat(childPath);
      if (info.isDirectory()) await this.scan(childRelative);
      else if (info.isFile()) this.entries.set(childRelative, [childPath, info.size]);
    }
  }

  partNames(): Iterable<string> {
    return this.entries.keys();
  }

  has(partName: string): boolean {
    return this.entries.has(normalize(partName));
  }

  async load(partName: string): Promise<Uint8Array> {
    const entry = this.entries.get(normalize(partName));
    if (!entry) throw new Docx4JException(`No part ${partName} in ${this.dir}`);
    return new Uint8Array(await readFile(entry[0]));
  }

  size(partName: string): number | undefined {
    return this.entries.get(normalize(partName))?.[1];
  }
}

/**
 * Writes a package to a directory (docx4j `UnzippedPartStore`'s save half): one file per part,
 * directories created as needed, `[Content_Types].xml` first. `finish()` resolves to the
 * directory path.
 *
 * ```ts
 * await pkg.saveTo(new DirectoryPartSink('unzipped'));
 * ```
 *
 * Nothing already in the directory is removed: a part the package no longer has stays on disk.
 */
export class DirectoryPartSink implements PartSink<string> {
  /** Insertion-ordered, so `[Content_Types].xml` (which `savePackage` puts first) is written first. */
  private readonly pending = new Map<string, Uint8Array>();

  constructor(readonly dir: string) {}

  put(partName: string, bytes: Uint8Array, _options?: PutOptions): void {
    this.pending.set(normalize(partName), bytes);
  }

  async finish(): Promise<string> {
    const first = this.pending.get(CONTENT_TYPES_NAME);
    if (first !== undefined) {
      // written first even if a caller put it later
      this.pending.delete(CONTENT_TYPES_NAME);
      await this.write(CONTENT_TYPES_NAME, first);
    }
    for (const [name, bytes] of this.pending) await this.write(name, bytes);
    this.pending.clear();
    return this.dir;
  }

  private async write(partName: string, bytes: Uint8Array): Promise<void> {
    const path = join(this.dir, partName);
    // A part name that climbs out of the directory (docx4j's "Zip Slip" check).
    if (path !== this.dir && !path.startsWith(this.dir.endsWith(sep) ? this.dir : this.dir + sep)) {
      throw new Docx4JException(`Part ${partName} is outside the target directory`);
    }
    await mkdir(dirname(path), { recursive: true });
    await writeFile(path, bytes);
  }
}

function normalize(name: string): string {
  return name.startsWith('/') ? name.substring(1) : name;
}

/** `[Content_Types].xml` first, then by name; the order the entries are listed in. */
function compareNames(a: string, b: string): number {
  if (a === CONTENT_TYPES_NAME) return -1;
  if (b === CONTENT_TYPES_NAME) return 1;
  return a < b ? -1 : a > b ? 1 : 0;
}
