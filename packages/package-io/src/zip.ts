// DOCX package IO. Entry order is preserved end to end (Map insertion order ==
// central-directory order on read AND write) — JavaScript objects reorder
// integer-like keys ("2" before "10"), so both directions avoid object
// conversion: read order comes from a central-directory scan, and the writer
// assembles the archive itself. Content compression uses fflate's raw DEFLATE.
// Every timestamp field is fixed: the same entries always produce the same
// bytes.
import { deflateSync, unzipSync } from "fflate";

export type Entries = Map<string, Uint8Array>;

export function readPackage(bytes: Uint8Array): Entries {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const unzipped = unzipSync(data);
  const order = centralDirectoryOrder(data);
  const entries: Entries = new Map();
  if (order) {
    for (const name of order) {
      const content = unzipped[name];
      if (content !== undefined) entries.set(name, content);
    }
    // any entry the directory scan missed (unexpected zip layout) is appended
    for (const [name, content] of Object.entries(unzipped)) {
      if (!entries.has(name)) entries.set(name, content);
    }
  } else {
    for (const [name, content] of Object.entries(unzipped)) entries.set(name, content);
  }
  return entries;
}

export function writePackage(entries: Entries): Uint8Array {
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const [name, content] of entries) {
    const nameBytes = new TextEncoder().encode(name);
    const deflated = content.length > 0 ? deflateSync(content, { level: 6 }) : content;
    const method = content.length > 0 ? 8 : 0;
    const crc = crc32(content);
    const local = [...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(method), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(deflated.length), ...u32(content.length), ...u16(nameBytes.length), ...u16(0), ...nameBytes];
    chunks.push(new Uint8Array(local), deflated);
    const cd = [...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(method), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(deflated.length), ...u32(content.length), ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset), ...nameBytes];
    central.push(new Uint8Array(cd));
    offset += local.length + deflated.length;
  }
  const cdStart = offset;
  let cdSize = 0;
  for (const c of central) cdSize += c.length;
  const eocd = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(entries.size), ...u16(entries.size), ...u32(cdSize), ...u32(cdStart), ...u16(0)]);
  const total = new Uint8Array(offset + cdSize + eocd.length);
  let at = 0;
  for (const c of [...chunks, ...central, eocd]) {
    total.set(c, at);
    at += c.length;
  }
  return total;
}

const u16 = (v: number): Uint8Array => new Uint8Array([v & 0xff, (v >>> 8) & 0xff]);
const u32 = (v: number): Uint8Array => new Uint8Array([v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff]);

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(data: Uint8Array): number {
  let c = ~0;
  for (let i = 0; i < data.length; i++) c = CRC_TABLE[(c ^ data[i]!) & 0xff]! ^ (c >>> 8);
  return ~c >>> 0;
}

/**
 * Ordered part names from the archive's central directory, or null when the
 * layout is not a plain (non-zip64) archive we can scan. This exists because
 * Object.keys reorders integer-like names ("2" before "10"), which would
 * silently reorder parts on read and write.
 */
function centralDirectoryOrder(data: Uint8Array): string[] | null {
  // locate End of Central Directory (signature 0x06054b50) by scanning back;
  // the comment field may be up to 65535 bytes
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let eocd = -1;
  const min = Math.max(0, data.length - 22 - 65535);
  for (let i = data.length - 22; i >= min; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return null;
  const count = view.getUint16(eocd + 10, true);
  if (count === 0xffff) return null; // zip64: fall back to object order
  let ptr = view.getUint32(eocd + 16, true);
  const names: string[] = [];
  const decoder = new TextDecoder();
  for (let n = 0; n < count; n++) {
    if (ptr + 46 > data.length || view.getUint32(ptr, true) !== 0x02014b50) return null;
    const nameLen = view.getUint16(ptr + 28, true);
    const extraLen = view.getUint16(ptr + 30, true);
    const commentLen = view.getUint16(ptr + 32, true);
    names.push(decoder.decode(data.subarray(ptr + 46, ptr + 46 + nameLen)));
    ptr += 46 + nameLen + extraLen + commentLen;
  }
  return names;
}

export function decodePart(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes);
}

export function encodePart(s: string): Uint8Array {
  return new TextEncoder().encode(s);
}

