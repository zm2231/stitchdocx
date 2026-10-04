// DOCX package IO. Entry order is preserved end to end (Map insertion order ==
// central-directory order on read AND write) — JavaScript objects reorder
// integer-like keys ("2" before "10"), so both directions avoid object
// conversion: read order comes from a central-directory scan, and the writer
// assembles the archive itself. Content compression uses fflate's raw DEFLATE.
// Every timestamp field is fixed: the same entries always produce the same
// bytes.
//
// Archives are read central-directory first (ZIP64 included), with input
// limits enforced before anything inflates: total and per-entry uncompressed
// size, entry count, duplicate names and encrypted entries are refused with a
// typed error. Entries stay compressed (RawEntry, which the writer passes
// through with its original bytes) until a part is asked for, every inflate
// is capped at the entry's recorded size, and an XML part over
// maxXmlPartBytes is refused before it inflates.
import { deflateSync, Inflate, inflateSync } from "fflate";

export type Entries = Map<string, EntryData>;

export class PackageLimitError extends Error {
  constructor(
    readonly code:
      | "package_too_large"
      | "too_many_entries"
      | "duplicate_entry"
      | "encrypted_entry"
      | "unreadable_archive"
      | "document_too_large",
    msg: string,
  ) {
    super(msg);
    this.name = "PackageLimitError";
  }
}

export const PACKAGE_LIMITS = {
  /** total uncompressed bytes across every entry */
  maxTotalUncompressed: 256 * 1024 * 1024,
  /** uncompressed bytes of one entry */
  maxEntryUncompressed: 64 * 1024 * 1024,
  /** entry count */
  maxEntries: 4096,
  /** uncompressed bytes of one XML part a command reads (main part, styles,
   *  numbering, settings, rels, story parts) */
  maxXmlPartBytes: 8 * 1024 * 1024,
  /** nodes one parse may build (see NodeBudget): an XML tree costs far more
   *  per node than per byte, so bytes alone do not bound it */
  maxPartNodes: 60_000,
};

/** an entry kept in its original compressed form; entryBytes() inflates it */
export class RawEntry {
  constructor(
    readonly compressed: Uint8Array,
    readonly method: number,
    readonly crc: number,
    readonly size: number,
  ) {}
}

export type EntryData = Uint8Array | RawEntry;

/** inflated bytes of an entry, inflating on demand when it is a RawEntry;
 *  never more than the size the archive records */
export function entryBytes(e: EntryData): Uint8Array {
  if (!(e instanceof RawEntry)) return e;
  let out: Uint8Array;
  if (e.method === 0) out = e.compressed;
  else if (e.method === 8) {
    try {
      out = inflateSync(e.compressed, { out: new Uint8Array(e.size + 1) });
    } catch (err) {
      throw new PackageLimitError("unreadable_archive", `entry does not inflate: ${String(err)}`);
    }
  } else throw new PackageLimitError("unreadable_archive", `unsupported compression method ${e.method}`);
  if (out.length !== e.size) throw new PackageLimitError("unreadable_archive", `entry holds ${out.length === e.size + 1 ? "more" : out.length} bytes, the archive records ${e.size}`);
  return out;
}

/** uncompressed size of an entry */
export function entrySize(e: EntryData): number {
  return e instanceof RawEntry ? e.size : e.length;
}

/** decoded text of an XML part; refuses document_too_large over
 *  maxXmlPartBytes before inflating */
export function xmlText(name: string, e: EntryData): string {
  const size = entrySize(e);
  if (size > PACKAGE_LIMITS.maxXmlPartBytes) {
    throw new PackageLimitError("document_too_large", `${name} inflates to ${size} bytes, over the limit of ${PACKAGE_LIMITS.maxXmlPartBytes}`);
  }
  return decodePart(entryBytes(e));
}

/** inflated content of an entry in chunks; a deflated entry inflates a few KB
 *  of input at a time, so no chunk is large and nothing holds the whole entry */
function* entryChunks(e: EntryData): Generator<Uint8Array> {
  if (!(e instanceof RawEntry) || e.method === 0) {
    yield entryBytes(e);
    return;
  }
  if (e.method !== 8) throw new PackageLimitError("unreadable_archive", `unsupported compression method ${e.method}`);
  const pending: Uint8Array[] = [];
  const inflater = new Inflate((chunk) => {
    pending.push(chunk);
  });
  const STEP = 4096;
  let produced = 0;
  for (let at = 0; at < e.compressed.length; at += STEP) {
    try {
      inflater.push(e.compressed.subarray(at, at + STEP), at + STEP >= e.compressed.length);
    } catch (err) {
      throw new PackageLimitError("unreadable_archive", `entry does not inflate: ${String(err)}`);
    }
    for (const chunk of pending.splice(0)) {
      produced += chunk.length;
      if (produced > e.size) throw new PackageLimitError("unreadable_archive", `entry holds more bytes than the archive records (${e.size})`);
      yield chunk;
    }
  }
  if (produced !== e.size) throw new PackageLimitError("unreadable_archive", `entry holds ${produced} bytes, the archive records ${e.size}`);
}

/** whether two entries hold the same content, compared chunk by chunk */
export function entriesEqual(a: EntryData, b: EntryData): boolean {
  if (entrySize(a) !== entrySize(b)) return false;
  if (a instanceof RawEntry && b instanceof RawEntry && a.method === b.method && a.crc === b.crc && bytesEqual(a.compressed, b.compressed)) return true;
  const ia = entryChunks(a);
  const ib = entryChunks(b);
  let ca: Uint8Array = new Uint8Array(0);
  let cb: Uint8Array = ca;
  for (;;) {
    if (ca.length === 0) {
      const n = ia.next();
      if (n.done) break;
      ca = n.value;
      continue;
    }
    if (cb.length === 0) {
      const n = ib.next();
      if (n.done) return false;
      cb = n.value;
      continue;
    }
    const k = Math.min(ca.length, cb.length);
    for (let i = 0; i < k; i++) if (ca[i] !== cb[i]) return false;
    ca = ca.subarray(k);
    cb = cb.subarray(k);
  }
  while (cb.length === 0) {
    const n = ib.next();
    if (n.done) return true;
    cb = n.value;
  }
  return false;
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

interface CentralEntry {
  name: string;
  method: number;
  flags: number;
  crc: number;
  size: number;
  compressedSize: number;
  localOffset: number;
}

/**
 * Central-directory-first package reader. The directory is parsed and every
 * input limit checked in the constructor; part bytes inflate only when asked.
 */
export class PackageReader {
  readonly names: string[];
  private central: Map<string, CentralEntry>;

  constructor(private data: Uint8Array) {
    const central = readCentralDirectory(data);
    this.names = central.map((e) => e.name);
    this.central = new Map(central.map((e) => [e.name, e]));
    let total = 0;
    const seen = new Set<string>();
    for (const e of central) {
      if (e.flags & 1) throw new PackageLimitError("encrypted_entry", `entry ${e.name} is encrypted`);
      if (seen.has(e.name)) throw new PackageLimitError("duplicate_entry", `archive holds duplicate entry ${e.name}`);
      seen.add(e.name);
      total += e.size;
      if (e.size > PACKAGE_LIMITS.maxEntryUncompressed) {
        throw new PackageLimitError("package_too_large", `entry ${e.name} inflates to ${e.size} bytes, over the per-entry limit`);
      }
    }
    if (total > PACKAGE_LIMITS.maxTotalUncompressed) {
      throw new PackageLimitError("package_too_large", `archive inflates to ${total} bytes, over the total limit`);
    }
  }

  has(name: string): boolean {
    return this.central.has(name);
  }

  /** decoded text of one XML part (see xmlText) */
  xml(name: string): string | undefined {
    const raw = this.raw(name);
    return raw ? xmlText(name, raw) : undefined;
  }

  /** one entry as stored: its compressed bytes, method, CRC and size, with
   *  nothing inflated */
  raw(name: string): RawEntry | undefined {
    const e = this.central.get(name);
    if (!e) return undefined;
    const local = e.localOffset;
    const view = new DataView(this.data.buffer, this.data.byteOffset, this.data.byteLength);
    if (local + 30 > this.data.length || view.getUint32(local, true) !== 0x04034b50) {
      throw new PackageLimitError("unreadable_archive", `entry ${name} has an unreadable local header`);
    }
    const nameLen = view.getUint16(local + 26, true);
    const extraLen = view.getUint16(local + 28, true);
    const start = local + 30 + nameLen + extraLen;
    if (start + e.compressedSize > this.data.length) {
      throw new PackageLimitError("unreadable_archive", `entry ${name} runs past the end of the archive`);
    }
    return new RawEntry(this.data.subarray(start, start + e.compressedSize), e.method, e.crc, e.size);
  }

  /** raw() under the registry's Parts interface */
  get(name: string): EntryData | undefined {
    return this.raw(name);
  }

  /** every entry as stored, in central-directory order */
  readAll(): Entries {
    const entries: Entries = new Map();
    for (const name of this.names) entries.set(name, this.raw(name)!);
    return entries;
  }
}

export function openPackage(bytes: Uint8Array): PackageReader {
  return new PackageReader(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
}

export function readPackage(bytes: Uint8Array): Entries {
  return openPackage(bytes).readAll();
}

export function writePackage(entries: Entries): Uint8Array {
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const [name, entry] of entries) {
    const nameBytes = new TextEncoder().encode(name);
    if (entry instanceof RawEntry) {
      // pass the original compressed bytes through; the central directory
      // carries the real sizes, so the data-descriptor flag is dropped
      const flags = 0;
      const local = [...u32(0x04034b50), ...u16(20), ...u16(flags), ...u16(entry.method), ...u16(0), ...u16(0x21), ...u32(entry.crc), ...u32(entry.compressed.length), ...u32(entry.size), ...u16(nameBytes.length), ...u16(0), ...nameBytes];
      chunks.push(new Uint8Array(local), entry.compressed);
      const cd = [...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(flags), ...u16(entry.method), ...u16(0), ...u16(0x21), ...u32(entry.crc), ...u32(entry.compressed.length), ...u32(entry.size), ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset), ...nameBytes];
      central.push(new Uint8Array(cd));
      offset += local.length + entry.compressed.length;
      continue;
    }
    const content = entry;
    const deflated = content.length > 0 ? deflateSync(content, { level: 6 }) : content;
    const method = content.length > 0 ? 8 : 0;
    const crc = crc32(content);
    // general-purpose bit 11 marks UTF-8 names; without it readers decode
    // high bytes as cp437 and a retained part effectively gets renamed
    const flags = nameBytes.some((b) => b > 0x7f) ? 0x0800 : 0;
    const local = [...u32(0x04034b50), ...u16(20), ...u16(flags), ...u16(method), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(deflated.length), ...u32(content.length), ...u16(nameBytes.length), ...u16(0), ...nameBytes];
    chunks.push(new Uint8Array(local), deflated);
    const cd = [...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(flags), ...u16(method), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(deflated.length), ...u32(content.length), ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset), ...nameBytes];
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

/** central-directory entries; ZIP64 sizes, offsets and counts are read from
 *  their ZIP64 records */
function readCentralDirectory(data: Uint8Array): CentralEntry[] {
  const unreadable = (msg: string): never => {
    throw new PackageLimitError("unreadable_archive", msg);
  };
  // locate End of Central Directory (signature 0x06054b50) by scanning back;
  // the comment field may be up to 65535 bytes
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const u64 = (at: number): number => {
    const v = view.getBigUint64(at, true);
    if (v > BigInt(Number.MAX_SAFE_INTEGER)) unreadable("ZIP64 field out of range");
    return Number(v);
  };
  let eocd = -1;
  const min = Math.max(0, data.length - 22 - 65535);
  for (let i = data.length - 22; i >= min; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) unreadable("no end of central directory record");
  let count = view.getUint16(eocd + 10, true);
  let ptr = view.getUint32(eocd + 16, true);
  if (count === 0xffff || ptr === 0xffffffff) {
    // ZIP64: the locator sits just before the EOCD and names the ZIP64 EOCD
    const loc = eocd - 20;
    if (loc < 0 || view.getUint32(loc, true) !== 0x07064b50) unreadable("ZIP64 locator missing");
    const eocd64 = u64(loc + 8);
    if (eocd64 + 56 > data.length || view.getUint32(eocd64, true) !== 0x06064b50) unreadable("ZIP64 end of central directory unreadable");
    count = u64(eocd64 + 32);
    ptr = u64(eocd64 + 48);
  }
  if (count > PACKAGE_LIMITS.maxEntries) {
    throw new PackageLimitError("too_many_entries", `archive holds ${count} entries, over the limit of ${PACKAGE_LIMITS.maxEntries}`);
  }
  const entries: CentralEntry[] = [];
  const decoder = new TextDecoder();
  for (let n = 0; n < count; n++) {
    if (ptr + 46 > data.length || view.getUint32(ptr, true) !== 0x02014b50) unreadable(`central directory entry ${n} unreadable`);
    const flags = view.getUint16(ptr + 8, true);
    const method = view.getUint16(ptr + 10, true);
    const crc = view.getUint32(ptr + 16, true);
    let compressedSize = view.getUint32(ptr + 20, true);
    let size = view.getUint32(ptr + 24, true);
    const nameLen = view.getUint16(ptr + 28, true);
    const extraLen = view.getUint16(ptr + 30, true);
    const commentLen = view.getUint16(ptr + 32, true);
    let localOffset = view.getUint32(ptr + 42, true);
    const extraStart = ptr + 46 + nameLen;
    const next = extraStart + extraLen + commentLen;
    if (next > data.length) unreadable(`central directory entry ${n} runs past the end of the archive`);
    if (size === 0xffffffff || compressedSize === 0xffffffff || localOffset === 0xffffffff) {
      // the ZIP64 extended-information extra field (0x0001) holds, in order,
      // each of these that is saturated
      let at = extraStart;
      let found = false;
      while (at + 4 <= extraStart + extraLen) {
        const tag = view.getUint16(at, true);
        const len = view.getUint16(at + 2, true);
        if (tag === 0x0001) {
          let f = at + 4;
          const end = f + len;
          const take = (): number => {
            if (f + 8 > end) unreadable(`central directory entry ${n} has a short ZIP64 field`);
            const v = u64(f);
            f += 8;
            return v;
          };
          if (size === 0xffffffff) size = take();
          if (compressedSize === 0xffffffff) compressedSize = take();
          if (localOffset === 0xffffffff) localOffset = take();
          found = true;
          break;
        }
        at += 4 + len;
      }
      if (!found) unreadable(`central directory entry ${n} lacks its ZIP64 field`);
    }
    entries.push({
      name: decoder.decode(data.subarray(ptr + 46, ptr + 46 + nameLen)),
      method,
      flags,
      crc,
      size,
      compressedSize,
      localOffset,
    });
    ptr = next;
  }
  return entries;
}

export function decodePart(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes);
}

export function encodePart(s: string): Uint8Array {
  return new TextEncoder().encode(s);
}
