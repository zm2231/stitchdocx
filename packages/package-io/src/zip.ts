// DOCX package IO. Entry order is preserved (Map insertion order == archive
// order); content is what matters for the contract — every entry except the
// one being edited must round-trip byte-identically. fflate is deterministic
// EXCEPT for the per-entry mtime it embeds by default (2-second DOS
// granularity), so every entry gets a fixed timestamp: the same entries always
// produce the same bytes.
import { unzipSync, zipSync } from "fflate";

export type Entries = Map<string, Uint8Array>;

const FIXED_MTIME = new Date(2000, 0, 1); // local-construct: fflate validates local date parts

export function readPackage(bytes: Uint8Array): Entries {
  const unzipped = unzipSync(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  return new Map(Object.entries(unzipped));
}

export function writePackage(entries: Entries): Uint8Array {
  const obj: Record<string, [Uint8Array, { mtime: Date }]> = {};
  for (const [k, v] of entries) obj[k] = [v, { mtime: FIXED_MTIME }];
  return zipSync(obj as never);
}

export function decodePart(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes);
}

export function encodePart(s: string): Uint8Array {
  return new TextEncoder().encode(s);
}
