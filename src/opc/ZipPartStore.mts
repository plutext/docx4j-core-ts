import { inflateSync, zipSync, type Zippable } from 'fflate';
import type { PartStore, PartSink, PutOptions } from './PartStore.mjs';
import { InvalidFormatException } from './exceptions.mjs';

interface Entry {
  name: string;
  method: number;
  compressedSize: number;
  size: number;
  localHeaderOffset: number;
}

/**
 * A zip container read from bytes (docx4j io3.stores.ZipPartStore). The central directory is
 * read once; an entry is inflated on first `load` and then cached. ZIP64 archives are rejected.
 */
export class ZipPartStore implements PartStore {
  private readonly data: Uint8Array;
  private readonly entries = new Map<string, Entry>();
  private readonly order: string[] = [];
  private readonly cache = new Map<string, Uint8Array>();

  constructor(source: Uint8Array | ArrayBuffer) {
    this.data = source instanceof Uint8Array ? source : new Uint8Array(source);
    this.readCentralDirectory();
  }

  partNames(): Iterable<string> {
    return this.order;
  }

  has(partName: string): boolean {
    return this.entries.has(partName);
  }

  size(partName: string): number | undefined {
    return this.entries.get(partName)?.size;
  }

  async load(partName: string): Promise<Uint8Array> {
    return this.loadSync(partName);
  }

  loadSync(partName: string): Uint8Array {
    const cached = this.cache.get(partName);
    if (cached) return cached;
    const e = this.entries.get(partName);
    if (!e) throw new Error(`No entry ${partName} in zip`);
    const d = this.data;
    const off = e.localHeaderOffset;
    if (readU32(d, off) !== 0x04034b50) throw new InvalidFormatException(`Bad local header for ${partName}`);
    const nameLen = readU16(d, off + 26);
    const extraLen = readU16(d, off + 28);
    const start = off + 30 + nameLen + extraLen;
    const raw = d.subarray(start, start + e.compressedSize);
    let bytes: Uint8Array;
    if (e.method === 0) bytes = raw.slice();
    else if (e.method === 8) bytes = inflateSync(raw);
    else throw new InvalidFormatException(`Unsupported compression method ${e.method} for ${partName}`);
    this.cache.set(partName, bytes);
    return bytes;
  }

  private readCentralDirectory(): void {
    const d = this.data;
    // End of central directory record: signature 0x06054b50, at most 64 KB of comment after it.
    let eocd = -1;
    for (let i = d.length - 22; i >= 0 && i >= d.length - 22 - 0xffff; i--) {
      if (readU32(d, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new InvalidFormatException('Not a zip file (no end of central directory record)');
    const count = readU16(d, eocd + 10);
    const cdOffset = readU32(d, eocd + 16);
    if (count === 0xffff || cdOffset === 0xffffffff) throw new InvalidFormatException('ZIP64 archives are not supported');
    let p = cdOffset;
    const decoder = new TextDecoder('utf-8');
    for (let i = 0; i < count; i++) {
      if (readU32(d, p) !== 0x02014b50) throw new InvalidFormatException('Bad central directory entry');
      const method = readU16(d, p + 10);
      const compressedSize = readU32(d, p + 20);
      const size = readU32(d, p + 24);
      const nameLen = readU16(d, p + 28);
      const extraLen = readU16(d, p + 30);
      const commentLen = readU16(d, p + 32);
      const localHeaderOffset = readU32(d, p + 42);
      if (compressedSize === 0xffffffff || size === 0xffffffff || localHeaderOffset === 0xffffffff) {
        throw new InvalidFormatException('ZIP64 archives are not supported');
      }
      const name = decoder.decode(d.subarray(p + 46, p + 46 + nameLen));
      p += 46 + nameLen + extraLen + commentLen;
      if (name.endsWith('/')) continue; // directory entry
      this.entries.set(name, { name, method, compressedSize, size, localHeaderOffset });
      this.order.push(name);
    }
  }
}

/**
 * Writes a zip (docx4j ZipPartStore, the save half). Entries are written in `put` order, so the
 * saver puts `[Content_Types].xml` first, as Word does; entries are deflated unless `compress`
 * is false (already-compressed media).
 */
export class ZipPartSink implements PartSink<Uint8Array> {
  private readonly files: Zippable = {};
  private readonly level: number;

  constructor(options?: { level?: number }) {
    this.level = options?.level ?? 6;
  }

  put(partName: string, bytes: Uint8Array, options?: PutOptions): void {
    const name = partName.startsWith('/') ? partName.substring(1) : partName;
    this.files[name] = [bytes, { level: options?.compress === false ? 0 : (this.level as 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9) }];
  }

  async finish(): Promise<Uint8Array> {
    return zipSync(this.files);
  }
}

function readU16(d: Uint8Array, p: number): number {
  return d[p]! | (d[p + 1]! << 8);
}

function readU32(d: Uint8Array, p: number): number {
  return (d[p]! | (d[p + 1]! << 8) | (d[p + 2]! << 16) | (d[p + 3]! << 24)) >>> 0;
}
