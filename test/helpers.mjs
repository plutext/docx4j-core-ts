import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { deflateSync } from 'node:zlib';
import { parseXml } from '../dist/index.mjs';
import { strToU8 } from 'fflate';

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

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * docx4j BrOutsideARunTest's comparison over a story part's XML (CR-021): text, breaks with their type, field
 * characters and paragraph ends, in order, each text and break with its run's bold, italic and size; a run's text
 * joins the token before where the element and the properties are the same, so Word's run boundaries and ours may
 * differ. Property elements (`w:pPr`, `w:sectPr`, ...) and anything outside WordprocessingML are passed over.
 */
export function breakTokens(xml) {
  const doc = parseXml(xml);
  const out = [];
  const children = (n) => [...n.childNodes].filter((c) => c.nodeType === 1 && c.namespaceURI === W_NS);
  const props = (r) => {
    const rPr = children(r).find((c) => c.localName === 'rPr');
    if (!rPr) return '';
    const get = (name) => children(rPr).find((c) => c.localName === name);
    const on = (name) => get(name) !== undefined && !['0', 'false', 'off'].includes(get(name).getAttributeNS(W_NS, 'val'));
    const sz = get('sz');
    const shown = [on('b') ? 'b' : '', on('i') ? 'i' : '', sz ? `sz${sz.getAttributeNS(W_NS, 'val')}` : ''].filter(Boolean).join(' ');
    return shown === '' ? '' : `[${shown}]`;
  };
  const walk = (n) => {
    for (const c of children(n)) {
      if (c.localName === 'r') {
        const p = props(c);
        for (const item of children(c)) {
          if (['t', 'instrText', 'delText', 'delInstrText'].includes(item.localName)) {
            const head = `${item.localName}${p}:`;
            if (out.length > 0 && out[out.length - 1].startsWith(head)) out[out.length - 1] += item.textContent;
            else out.push(head + item.textContent);
          } else if (item.localName === 'br') {
            const type = item.getAttributeNS(W_NS, 'type');
            out.push(`br${type ? `(${type})` : ''}${p}`);
          } else if (item.localName === 'fldChar') out.push(`fld:${item.getAttributeNS(W_NS, 'fldCharType')}`);
        }
      } else if (!c.localName.endsWith('Pr') && c.localName !== 'tblGrid') {
        walk(c);
        if (c.localName === 'p') out.push('¶');
      }
    }
  };
  walk(doc.documentElement);
  return out;
}

/**
 * A stored zip of `[name, bytes]` entries in order, repeated names and case variants kept as given, each name in UTF-8 with
 * the language encoding flag set: what fflate's writers refuse to make (CR-024's duplicate entries, Word check 75).
 */
export function storedZip(entries) {
  const table = new Uint32Array(256).map((_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc32 = (b) => {
    let c = 0xffffffff;
    for (const x of b) c = table[(c ^ x) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const u16 = (v) => [v & 0xff, (v >>> 8) & 0xff];
  const u32 = (v) => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
  const chunks = [];
  const central = [];
  let offset = 0;
  for (const [name, data] of entries) {
    const n = strToU8(name);
    const crc = crc32(data);
    // version 2.0, flag 0x0800 (UTF-8 names), stored, no time
    const local = Uint8Array.from([...u32(0x04034b50), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21),
      ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(n.length), ...u16(0)]);
    chunks.push(local, n, data);
    central.push(Uint8Array.from([...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21),
      ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(n.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0),
      ...u32(offset)]), n);
    offset += local.length + n.length + data.length;
  }
  const size = central.reduce((s, c) => s + c.length, 0);
  const end = Uint8Array.from([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(entries.length), ...u16(entries.length), ...u32(size),
    ...u32(offset), ...u16(0)]);
  const all = [...chunks, ...central, end];
  const out = new Uint8Array(all.reduce((s, c) => s + c.length, 0));
  let p = 0;
  for (const c of all) { out.set(c, p); p += c.length; }
  return out;
}
