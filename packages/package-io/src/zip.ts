// DOCX package IO. Entry order is preserved (Map insertion order == archive
// order); content is what matters for the contract — every entry except the
// one being edited must round-trip byte-identically, and zipSync is
// deterministic (fflate stores no timestamps), so the same entries always
// produce the same bytes.
import { unzipSync, zipSync } from "fflate";

export type Entries = Map<string, Uint8Array>;

export function readPackage(bytes: Uint8Array): Entries {
  const unzipped = unzipSync(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  return new Map(Object.entries(unzipped));
}

export function writePackage(entries: Entries): Uint8Array {
  const obj: Record<string, Uint8Array> = {};
  for (const [k, v] of entries) obj[k] = v;
  return zipSync(obj);
}

export function decodePart(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes);
}

export function encodePart(s: string): Uint8Array {
  return new TextEncoder().encode(s);
}
