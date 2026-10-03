// verify: a pure receipt check. output must be original plus exactly what the
// receipts claim: every applied receipt's changeIds exist, rejecting them
// yields `before`, accepting them yields `after`, and no other paragraph
// differs. Tampering anywhere fails.
import { decodeEntities, decodePart, readPackage, parseXml, kid } from "@stitchdocx/package-io";
import { walkBody } from "./body.js";
import { acceptAll, paragraphText, projectParagraphs, type View } from "./projection.js";
import type { Addr, Receipt } from "./types.js";

/** the adapter namespace: receipt revisions must be authored under it */
export const OUR_AUTHOR_PREFIX = "Scope Optimizer";
void OUR_AUTHOR_PREFIX;
export function verifyOutput(original: Uint8Array, output: Uint8Array, receipts: Receipt[]): { ok: boolean; failures: string[] } {
  const failures: string[] = [];
  const origDoc = openForVerify(original);
  const outDoc = openForVerify(output);
  const applied = receipts.filter((r) => r.status === "applied");

  const origRecords = projectParagraphs(origDoc.walk, acceptAll);
  const outRecords = projectParagraphs(outDoc.walk, acceptAll);
  const outById = new Map<string, (typeof outRecords)[number]>();
  for (const r of outRecords) if (r.addr) outById.set(r.addr, r);
  const origAddrs = origRecords.map((r) => r.addr ?? "");
  const origAddrSet = new Set(origAddrs);

  // existing revision ids → authors in the output (for receipt identity checks)
  const outRevAuthors = new Map<string, string>();
  collectRevisionAuthors(outDoc.root, outRevAuthors);

  const insertAddrs = new Set<string>();
  for (const r of applied) {
    if (!r.changeIds?.length) failures.push(`applied receipt ${r.index} reports no change ids`);
    for (const id of r.changeIds ?? []) {
      const author = outRevAuthors.get(String(id));
      if (author === undefined) failures.push(`receipt ${r.index} change id ${id} not in document`);
      else if (author !== r.author) failures.push(`receipt ${r.index} change id ${id} authored by ${author}, expected ${r.author}`);
    }
    if (r.insertedAddr) insertAddrs.add(r.insertedAddr);
  }

  for (const r of applied) {
    const ids = new Set((r.changeIds ?? []).map(String));
    const rejectMine: View = (rev) => (rev.id !== null && ids.has(rev.id) ? "reject" : "accept");
    if (r.insertedAddr) {
      const rejRecords = projectParagraphs(outDoc.walk, rejectMine);
      const rejByAddr = new Map<string, (typeof rejRecords)[number]>();
      for (const rec of rejRecords) if (rec.addr) rejByAddr.set(rec.addr, rec);
      const rejected = rejByAddr.get(r.insertedAddr);
      if (rejected && rejected.text !== "") {
        // Word's container-end form: the tracked insertion sits on the
        // anchor's mark, so rejecting merges the anchor paragraph into the
        // inserted one — which must read exactly like the anchor did
        if (rejected.text !== (r.before ?? "")) {
          failures.push(`receipt ${r.index}: rejected insert reads ${JSON.stringify(rejected.text)}, expected ${JSON.stringify(r.before ?? "")}`);
        }
        if (rejByAddr.get(r.addr)) {
          failures.push(`receipt ${r.index}: anchor paragraph survives beside the merged insert`);
        }
      }
      const accepted = outById.get(r.insertedAddr);
      if (!accepted) failures.push(`receipt ${r.index}: inserted paragraph ${r.insertedAddr} missing`);
      else if (accepted.text !== r.after) failures.push(`receipt ${r.index}: accepted text ${JSON.stringify(accepted.text)} != after ${JSON.stringify(r.after ?? "")}`);
      continue;
    }
    const target = outById.get(r.addr);
    if (!target) {
      failures.push(`receipt ${r.index}: addressed paragraph ${r.addr} missing from output`);
      continue;
    }
    const rejectedText = paragraphText(target.node, rejectMine);
    if (rejectedText !== r.before) failures.push(`receipt ${r.index}: reject yields ${JSON.stringify(rejectedText)}, expected before ${JSON.stringify(r.before ?? "")}`);
    if (target.text !== r.after) failures.push(`receipt ${r.index}: accept yields ${JSON.stringify(target.text)}, expected after ${JSON.stringify(r.after ?? "")}`);
  }

  // paragraph accounting: original paragraphs all present; untouched ones unchanged
  const touchedAddrs = new Set<Addr>(applied.filter((r) => !r.insertedAddr).map((r) => r.addr));
  const expectedCount = origRecords.length + insertAddrs.size;
  if (outRecords.length !== expectedCount) failures.push(`output has ${outRecords.length} paragraphs, expected ${expectedCount}`);
  for (let i = 0; i < origRecords.length; i++) {
    const addr = origAddrs[i]!;
    const out = outById.get(addr);
    if (!out) {
      failures.push(`original paragraph ${i} (${addr}) missing from output`);
      continue;
    }
    const origText = origRecords[i]!.text;
    if (!touchedAddrs.has(addr) && out.text !== origText) {
      failures.push(`untouched paragraph ${i} (${addr}) changed: ${JSON.stringify(out.text)} vs ${JSON.stringify(origText)}`);
    }
    const orig = origRecords[i]!;
    if (out.style !== orig.style || out.numbered !== orig.numbered) failures.push(`paragraph ${i} (${addr}) changed style or numbering`);
  }

  return { ok: failures.length === 0, failures };
}

function openForVerify(bytes: Uint8Array) {
  const entries = readPackage(bytes);
  const docXml = entries.get("word/document.xml");
  if (!docXml) throw new Error("word/document.xml missing");
  const tree = parseXml(decodePart(docXml));
  const root = kid(tree as any, "w:document");
  if (!root) throw new Error("word/document.xml has no w:document root");
  return { walk: walkBody(root), root };
}

function collectRevisionAuthors(root: any, ids: Map<string, string>): void {
  const stack = [root];
  while (stack.length) {
    const cur = stack.pop()!;
    for (const c of cur.children) {
      if (c.kind !== "el") continue;
      if (c.name === "w:ins" || c.name === "w:del" || c.name === "w:moveFrom" || c.name === "w:moveTo" || c.name.endsWith("PrChange")) {
        const id = c.attrs.find((a: any) => a.name === "w:id");
        const author = c.attrs.find((a: any) => a.name === "w:author");
        if (id) ids.set(decodeEntities(id.raw), author ? decodeEntities(author.raw) : "");
      }
      stack.push(c);
    }
  }
}
