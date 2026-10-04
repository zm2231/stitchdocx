// DOCX package IO. Entry order is preserved end to end (Map insertion order ==
// central-directory order on read AND write) — JavaScript objects reorder
// integer-like keys ("2" before "10"), so both directions avoid object
// conversion: read order comes from a central-directory scan, and the writer
// assembles the archive itself. Content compression uses fflate's raw DEFLATE.
// Every timestamp field is fixed: the same entries always produce the same
// bytes.
//
// Archives are read central-directory first, with input limits enforced before
// anything inflates: total and per-entry uncompressed size, entry count,
// duplicate names and encrypted entries are refused with a typed error, and
// parts a command does not need stay uninflated (RawEntry — the writer passes
// them through with their original compressed bytes).
import { deflateSync, inflateSync, unzipSync } from "fflate";

export type Entries = Map<string, EntryData>;

export class PackageLimitError extends Error {
  constructor(
    readonly code:
      | "package_too_large"
      | "too_many_entries"
      | "duplicate_entry"
      | "encrypted_entry"
      | "unreadable_archive",
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

/** inflated bytes of an entry, inflating on demand when it is a RawEntry */
export function entryBytes(e: EntryData): Uint8Array {
  if (e instanceof RawEntry) {
    if (e.method === 0) return e.compressed;
    if (e.method === 8) return inflateSync(e.compressed);
    throw new PackageLimitError("unreadable_archive", `unsupported compression method ${e.method}`);
  }
  return e;
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
  private eager: Map<string, Uint8Array> | null;

  constructor(private data: Uint8Array) {
    const central = readCentralDirectory(data);
    if (central) {
      this.eager = null;
      this.names = central.map((e) => e.name);
      this.central = new Map(central.map((e) => [e.name, e]));
      if (central.length > PACKAGE_LIMITS.maxEntries) {
        throw new PackageLimitError("too_many_entries", `archive holds ${central.length} entries, over the limit of ${PACKAGE_LIMITS.maxEntries}`);
      }
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
    } else {
      // zip64 or unusual layout: fall back to a full inflate, then apply the
      // same limits to what actually came out
      const unzipped = new Map(Object.entries(unzipSync(data)));
      this.eager = unzipped;
      this.central = new Map();
      this.names = [...unzipped.keys()];
      if (this.names.length > PACKAGE_LIMITS.maxEntries) {
        throw new PackageLimitError("too_many_entries", `archive holds ${this.names.length} entries, over the limit of ${PACKAGE_LIMITS.maxEntries}`);
      }
      const seen = new Set<string>();
      let total = 0;
      for (const n of this.names) {
        if (seen.has(n)) throw new PackageLimitError("duplicate_entry", `archive holds duplicate entry ${n}`);
        seen.add(n);
        total += unzipped.get(n)!.length;
        if (unzipped.get(n)!.length > PACKAGE_LIMITS.maxEntryUncompressed) {
          throw new PackageLimitError("package_too_large", `entry ${n} inflates to ${unzipped.get(n)!.length} bytes, over the per-entry limit`);
        }
      }
      if (total > PACKAGE_LIMITS.maxTotalUncompressed) {
        throw new PackageLimitError("package_too_large", `archive inflates to ${total} bytes, over the total limit`);
      }
    }
  }

  has(name: string): boolean {
    return this.eager ? this.eager.has(name) : this.central.has(name);
  }

  /** inflated bytes of one part, inflating on demand */
  part(name: string): Uint8Array | undefined {
    if (this.eager) return this.eager.get(name);
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
    const compressed = this.data.subarray(start, start + e.compressedSize);
    return entryBytes(new RawEntry(compressed, e.method, e.crc, e.size));
  }

  /** part() under the registry's Parts interface */
  get(name: string): Uint8Array | undefined {
    return this.part(name);
  }

  /**
   * Every entry, in central-directory order. Parts selected by `only` come
   * back inflated; the rest stay RawEntry so the writer can pass their
   * original compressed bytes through untouched.
   */
  readAll(only?: (name: string) => boolean): Entries {
    const entries: Entries = new Map();
    for (const name of this.names) {
      if (!only || only(name)) entries.set(name, this.part(name)!);
      else if (this.eager) entries.set(name, this.eager.get(name)!);
      else {
        const e = this.central.get(name)!;
        const local = e.localOffset;
        const view = new DataView(this.data.buffer, this.data.byteOffset, this.data.byteLength);
        const nameLen = view.getUint16(local + 26, true);
        const extraLen = view.getUint16(local + 28, true);
        const start = local + 30 + nameLen + extraLen;
        entries.set(name, new RawEntry(this.data.subarray(start, start + e.compressedSize), e.method, e.crc, e.size));
      }
    }
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

/** central-directory entries, or null when the layout is not a plain
 *  (non-zip64) archive we can scan */
function readCentralDirectory(data: Uint8Array): CentralEntry[] | null {
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
  if (count === 0xffff) return null; // zip64: fall back to a full inflate
  let ptr = view.getUint32(eocd + 16, true);
  const entries: CentralEntry[] = [];
  const decoder = new TextDecoder();
  for (let n = 0; n < count; n++) {
    if (ptr + 46 > data.length || view.getUint32(ptr, true) !== 0x02014b50) return null;
    const flags = view.getUint16(ptr + 8, true);
    const method = view.getUint16(ptr + 10, true);
    const crc = view.getUint32(ptr + 16, true);
    const compressedSize = view.getUint32(ptr + 20, true);
    const size = view.getUint32(ptr + 24, true);
    const nameLen = view.getUint16(ptr + 28, true);
    const extraLen = view.getUint16(ptr + 30, true);
    const commentLen = view.getUint16(ptr + 32, true);
    const localOffset = view.getUint32(ptr + 42, true);
    entries.push({
      name: decoder.decode(data.subarray(ptr + 46, ptr + 46 + nameLen)),
      method,
      flags,
      crc,
      size,
      compressedSize,
      localOffset,
    });
    ptr += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

export function decodePart(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes);
}

export function encodePart(s: string): Uint8Array {
  return new TextEncoder().encode(s);
}
