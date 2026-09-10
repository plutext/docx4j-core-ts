import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

export const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), 'fixtures');

export async function fixture(name) {
  return new Uint8Array(await readFile(join(fixturesDir, name)));
}

export function bytesEqual(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** A structural copy without PARENT pointers and TYPE_NAME noise, for deep equality. */
export function plain(value) {
  return JSON.parse(JSON.stringify(value, (k, v) => (k === 'PARENT' ? undefined : v)));
}
