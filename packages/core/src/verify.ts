// verify: a pure receipt check. output must be original plus exactly what the
// receipts claim: every applied receipt's changeIds exist, rejecting them
// yields `before`, accepting them yields `after`, no other paragraph differs,
// every other package entry is byte-identical and none is added or removed,
// and the main part outside paragraphs is canonically identical. Tampering
// anywhere fails.
import { cloneNode, decodePart, decodeEntities, kid, mainPartName, openPackage, parseXml, serializeXml, removeAttr, setAttr, textNode, type XElement } from "@stitchdocx/package-io";
import { walkBody } from "./body.js";
import { acceptAll, attrVal, markRevision, paragraphText, projectParagraphs, wordText, type View } from "./projection.js";
import { isRevisionEl } from "./revisions.js";
import type { Addr, Receipt } from "./types.js";

export function verifyOutput(original: Uint8Array, output: Uint8Array, receipts: Receipt[]): { ok: boolean; failures: string[] } {
  const failures: string[] = [];
  const applied = receipts.filter((r) => r.status === "applied");

  // — package accounting: same entries, same order, everything but the main
  // part byte-identical
  const origPkg = openPackage(original);
  const outPkg = openPackage(output);
  if (origPkg.names.join("\u0000") !== outPkg.names.join("\u0000")) {
    failures.push("package entries differ (added, removed or reordered)");
  } else {
    const mainPart = mainPartName(origPkg);
    for (const name of origPkg.names) {
      if (name === mainPart) continue;
      const a = origPkg.part(name)!;
      const b = outPkg.part(name)!;
      if (!bytesEqual(a, b)) failures.push(`package entry ${name} was modified`);
    }
  }

  const origDoc = openForVerify(original);
  const outDoc = openForVerify(output);

  // — main part outside paragraphs: canonically identical (paragraphs are
  // removed wholesale from both trees; everything that remains — body
  // properties, section breaks, tables' shells, text-box anchors — must match)
  if (stripParagraphs(origDoc.root) !== stripParagraphs(outDoc.root)) {
    failures.push("document content outside paragraphs was modified");
  }

  const origRecords = projectParagraphs(origDoc.walk, acceptAll);
  const outRecords = projectParagraphs(outDoc.walk, acceptAll);
  const outById = new Map<string, (typeof outRecords)[number]>();
  for (const r of outRecords) if (r.addr) outById.set(r.addr, r);
  const origByAddr = new Map<string, (typeof origRecords)[number]>();
  for (const r of origRecords) if (r.addr) origByAddr.set(r.addr, r);
  const origAddrs = origRecords.map((r) => r.addr ?? "");

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
      const anchor = outById.get(r.addr);
      if (!anchor) failures.push(`receipt ${r.index}: anchor paragraph ${r.addr} missing`);
      else if (anchor.text !== (r.before ?? "")) failures.push(`receipt ${r.index}: anchor reads ${JSON.stringify(anchor.text)}, expected before ${JSON.stringify(r.before ?? "")}`);
      continue;
    }
    const target = outById.get(r.addr);
    if (!target) {
      failures.push(`receipt ${r.index}: addressed paragraph ${r.addr} missing from output`);
      continue;
    }
    // receipts report in projection-record space: a removed mark may merge
    // other paragraphs' content into this record under either view
    const rejectedRecord = projectParagraphs(outDoc.walk, rejectMine).find((rec) => rec.addr === r.addr);
    const rejectedText = rejectedRecord ? rejectedRecord.text : paragraphText(target.node, rejectMine);
    if (rejectedText !== r.before) failures.push(`receipt ${r.index}: reject yields ${JSON.stringify(rejectedText)}, expected before ${JSON.stringify(r.before ?? "")}`);
    if (target.text !== r.after) failures.push(`receipt ${r.index}: accept yields ${JSON.stringify(target.text)}, expected after ${JSON.stringify(r.after ?? "")}`);
  }

  // paragraph accounting: original paragraphs all present; untouched ones
  // byte-identical; touched ones carry ONLY the receipt's revisions (the
  // paragraph with those revisions inverted must equal the original,
  // canonicalized). An insert's anchor counts as touched ONLY when the
  // receipt's revisions actually reach it — the container-end form tracks the
  // anchor's own mark; mid-container and table-scope inserts leave the anchor
  // byte-identical.
  const touchedAddrs = new Set<Addr>(applied.filter((r) => !r.insertedAddr).map((r) => r.addr));
  const changeIdsByAddr = new Map<Addr, Set<string>>();
  for (const r of applied) {
    if (r.insertedAddr) continue;
    const ids = changeIdsByAddr.get(r.addr) ?? new Set<string>();
    for (const id of r.changeIds ?? []) ids.add(String(id));
    changeIdsByAddr.set(r.addr, ids);
  }
  for (const r of applied) {
    if (!r.insertedAddr) continue;
    const anchor = outById.get(r.addr);
    const mark = anchor ? markRevision(anchor.node) : null;
    const markId = mark ? attrVal(mark, "w:id") : null;
    if (markId !== null && (r.changeIds ?? []).map(String).includes(markId)) {
      touchedAddrs.add(r.addr);
      const ids = changeIdsByAddr.get(r.addr) ?? new Set<string>();
      ids.add(markId);
      changeIdsByAddr.set(r.addr, ids);
    }
  }
  const expectedCount = origRecords.length + insertAddrs.size;
  if (outRecords.length !== expectedCount) failures.push(`output has ${outRecords.length} paragraphs, expected ${expectedCount}`);
  for (let i = 0; i < origRecords.length; i++) {
    const addr = origAddrs[i]!;
    const out = outById.get(addr);
    if (!out) {
      failures.push(`original paragraph ${i} (${addr}) missing from output`);
      continue;
    }
    const orig = origRecords[i]!;
    if (!touchedAddrs.has(addr)) {
      if (out.text !== orig.text) {
        failures.push(`untouched paragraph ${i} (${addr}) changed: ${JSON.stringify(out.text)} vs ${JSON.stringify(orig.text)}`);
      }
      if (serializeXml(orig.node) !== serializeXml(out.node)) {
        failures.push(`untouched paragraph ${i} (${addr}) was modified`);
      }
    } else {
      const ours = changeIdsByAddr.get(addr) ?? new Set<string>();
      const inverted = canonicalParagraph(out.node, ours);
      const source = canonicalParagraph(orig.node, new Set());
      if (inverted !== source) {
        failures.push(`touched paragraph ${i} (${addr}) carries changes beyond its receipt`);
      }
    }
    if (out.style !== orig.style || out.numbered !== orig.numbered) failures.push(`paragraph ${i} (${addr}) changed style or numbering`);
  }

  return { ok: failures.length === 0, failures };
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a === b) return true;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** serialize the document with every w:p element removed (a fresh clone; the
 *  caller's tree is untouched) */
function stripParagraphs(root: XElement): string {
  const clone = cloneNode(root) as XElement;
  const doomed: XElement[] = [];
  const walk = (node: XElement): void => {
    for (const c of node.children) {
      if (c.kind !== "el") continue;
      if (c.name === "w:p") doomed.push(c);
      else walk(c);
    }
  };
  walk(clone);
  for (const p of doomed) {
    const parent = p.parent;
    if (parent) {
      const i = parent.children.indexOf(p);
      if (i >= 0) parent.children.splice(i, 1);
    }
  }
  return serializeXml(clone);
}

/**
 * Canonical form of a touched paragraph: the receipt's revision elements are
 * inverted (insertions dropped, deletions restored to runs, property-change
 * markers removed), runs split by the edit are merged back together (their
 * duplicated rPr copies collapse into the first), adjacent w:t elements
 * coalesce, and w:t xml:space is normalized against the text it carries.
 * Equal to the canonicalized original when the edit touched exactly what the
 * receipt claims.
 */
function canonicalParagraph(node: XElement, ours: ReadonlySet<string>): string {
  const clone = cloneNode(node) as XElement;
  invert(clone, ours);
  dropEmptyMarkPr(clone);
  stripRunPrRevisionRecords(clone);
  // per-element first: each w:t becomes its Word-read characters under
  // preserve, so a split's read-form slices and an unpreserved original
  // (whose edge whitespace Word drops and whose newlines read as spaces)
  // land on the same canonical text before the halves coalesce
  for (const c of walkEls(clone)) if (c.name === "w:t") setCanonicalText(c);
  mergeSplitRuns(clone);
  for (const c of walkEls(clone)) coalesceText(c);
  // an emptied paragraph canonicalizes to the self-closing form Word writes
  clone.selfClosing = clone.children.length === 0;
  return serializeXml(clone);
}

/** rewrite a w:t to hold exactly the characters Word reads from it */
function setCanonicalText(t: XElement): void {
  const read = wordText(t);
  t.children = [textNode(read)];
  removeAttr(t, "xml:space");
  setAttr(t, "xml:space", "preserve");
}

/** the anchor-mark insert form creates w:pPr/w:rPr shells to hold the tracked
 *  mark; with the mark inverted they are empty and vanish */
function dropEmptyMarkPr(paragraph: XElement): void {
  const pPr = paragraph.children.find((c) => c.kind === "el" && c.name === "w:pPr");
  if (!pPr || pPr.kind !== "el") return;
  const rPr = pPr.children.find((c) => c.kind === "el" && c.name === "w:rPr");
  if (rPr && rPr.kind === "el" && rPr.children.length === 0) {
    pPr.children.splice(pPr.children.indexOf(rPr), 1);
  }
  if (pPr.children.length === 0) paragraph.children.splice(paragraph.children.indexOf(pPr), 1);
}

/** run-level formatting-change records are dropped from the canonical form on
 *  BOTH sides: a split run's copies carry re-allocated ids, so record
 *  identity cannot survive the merge — everything else in w:rPr compares */
function stripRunPrRevisionRecords(node: XElement): void {
  for (const c of walkEls(node)) {
    if (c.name !== "w:r") continue;
    const rPr = c.children[0];
    if (rPr && rPr.kind === "el" && rPr.name === "w:rPr") {
      rPr.children = rPr.children.filter((x) => !(x.kind === "el" && isRevisionEl(x)));
    }
  }
}

/** rewrite the paragraph clone in place: drop our insertions, unwrap our
 *  deletions (delText back to t), drop our property-change markers */
function invert(node: XElement, ours: ReadonlySet<string>): void {
  const kept: typeof node.children = [];
  for (const c of node.children) {
    if (c.kind !== "el") {
      kept.push(c);
      continue;
    }
    const id = attrVal(c, "w:id") ?? "";
    if (isRevisionEl(c) && ours.has(id)) {
      if (c.name === "w:del" || c.name === "w:moveFrom") {
        restoreDelText(c);
        for (const cc of c.children) {
          if (cc.kind === "el") {
            invert(cc, ours);
            cc.parent = node;
          }
          kept.push(cc);
        }
      }
      // w:ins / w:moveTo / *PrChange: our insertions and markers vanish
      continue;
    }
    invert(c, ours);
    kept.push(c);
  }
  node.children = kept;
}

function restoreDelText(root: XElement): void {
  for (const c of walkEls(root)) {
    if (c.name === "w:delText") renameTo(c, "w:t");
  }
}

function walkEls(root: XElement): XElement[] {
  const out: XElement[] = [];
  const visit = (n: XElement): void => {
    for (const c of n.children) {
      if (c.kind !== "el") continue;
      out.push(c);
      visit(c);
    }
  };
  visit(root);
  return out;
}

/** rename keeping the attribute source form (mirrors revisions.renameEl) */
function renameTo(e: XElement, newName: string): void {
  if (!e.attrsDirty) {
    const attrsPart = e.rawOpen.slice(e.name.length);
    e.rawOpen = newName + attrsPart;
  }
  e.name = newName;
  e.rawClose = newName;
}

/** merge adjacent runs with identical formatting (an edit splits runs; the
 *  original was one), dropping the right half's rPr copy, and coalesce the
 *  adjacent w:t elements a split produced */
function mergeSplitRuns(node: XElement): void {
  for (const c of node.children) if (c.kind === "el") mergeSplitRuns(c);
  const merged: typeof node.children = [];
  for (const c of node.children) {
    const prev = merged[merged.length - 1]!;
    if (c.kind === "el" && c.name === "w:r" && prev && prev.kind === "el" && prev.name === "w:r" && runKey(prev) === runKey(c)) {
      const head = c.children[0];
      const from = head && head.kind === "el" && head.name === "w:rPr" ? 1 : 0;
      prev.children.push(...c.children.slice(from));
      for (const cc of c.children.slice(from)) if (cc.kind === "el") cc.parent = prev;
      coalesceText(prev);
      continue;
    }
    merged.push(c);
  }
  node.children = merged;
}

/** join adjacent w:t children (their split halves read as one) */
function coalesceText(node: XElement): void {
  const out: typeof node.children = [];
  for (const c of node.children) {
    const prev = out[out.length - 1]!;
    if (c.kind === "el" && c.name === "w:t" && prev && prev.kind === "el" && prev.name === "w:t") {
      prev.children.push(...c.children);
      continue;
    }
    out.push(c);
  }
  node.children = out;
}

function runKey(run: XElement): string {
  const rPr = run.children[0];
  const has = rPr && rPr.kind === "el" && rPr.name === "w:rPr";
  return has ? serializeXml(rPr as XElement) : "";
}


function openForVerify(bytes: Uint8Array) {
  const reader = openPackage(bytes);
  const docXml = reader.part(mainPartName(reader));
  if (!docXml) throw new Error("main document part missing");
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
