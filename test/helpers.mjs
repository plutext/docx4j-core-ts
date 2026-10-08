import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { deflateSync } from 'node:zlib';

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

/**
 * A solid-colour 8-bit RGB PNG at 96 dpi (pHYs 3780 px/m), deflated with node:zlib, as base64: the
 * picture binding tests and Word check 38's files need images of known sizes and colours.
 */
export function pngOf(width, height, [r, g, b]) {
  const table = new Int32Array(256).map((_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c; });
  const crc32 = (bytes) => { let c = -1; for (const byte of bytes) c = table[(c ^ byte) & 0xff] ^ (c >>> 8); return (c ^ -1) >>> 0; };
  const chunk = (type, data) => {
    const t = Buffer.from(type, 'latin1');
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(Buffer.concat([t, data])));
    return Buffer.concat([len, t, data, crc]);
  };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(width, 0); ihdr.writeUInt32BE(height, 4); ihdr[8] = 8; ihdr[9] = 2;
  const phys = Buffer.alloc(9); phys.writeUInt32BE(3780, 0); phys.writeUInt32BE(3780, 4); phys[8] = 1;
  const row = Buffer.alloc(1 + width * 3);
  for (let x = 0; x < width; x++) { row[1 + x * 3] = r; row[2 + x * 3] = g; row[3 + x * 3] = b; }
  const raw = Buffer.concat(Array.from({ length: height }, () => row));
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr), chunk('pHYs', phys), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0)),
  ]).toString('base64');
}
