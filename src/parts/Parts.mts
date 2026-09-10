import type { Part } from './Part.mjs';
import { PartName } from '../opc/PartName.mjs';

/** The parts of a package by name (docx4j Parts), case-insensitive. */
export class Parts implements Iterable<Part> {
  private readonly map = new Map<string, Part>();

  put(part: Part): void {
    this.map.set(part.partName.key, part);
  }

  get(partName: PartName | string): Part | undefined {
    return this.map.get(PartName.of(partName).key);
  }

  has(partName: PartName | string): boolean {
    return this.map.has(PartName.of(partName).key);
  }

  remove(partName: PartName | string): boolean {
    return this.map.delete(PartName.of(partName).key);
  }

  get size(): number {
    return this.map.size;
  }

  values(): IterableIterator<Part> {
    return this.map.values();
  }

  [Symbol.iterator](): IterableIterator<Part> {
    return this.map.values();
  }
}
