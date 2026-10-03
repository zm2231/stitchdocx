// Regression tests for the round-1 review findings — one test per ledger
// Accept check. Builders construct minimal OOXML packages in-memory.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { homedir } from "node:os";
import { describe, it, expect } from "vitest";
import { encodePart, readPackage, writePackage } from "@stitchdocx/package-io";
import { normalizeDocx, openDocument, verifyOutput, type Op, type Receipt } from "./index.js";
import { buildSpanMap, checkSpan } from "./spanmap.js";
import { wordEdits } from "./worddiff.js";

const HANDOFF = process.env.STITCH_HANDOFF ?? join(homedir(), "Downloads", "scope-docx-engine-handoff");

const DATE = "2026-10-03T00:00:00Z";

function docx(bodyChildrenXml: string): Uint8Array {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>${bodyChildrenXml}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  const entries = new Map<string, Uint8Array>([
    ["[Content_Types].xml", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
    ["word/document.xml", encodePart(xml)],
  ]);
  return writePackage(entries);
}

const P = (paraId: string | null, inner: string): string =>
  `<w:p${paraId ? ` w14:paraId="${paraId}"` : ""}>${inner}</w:p>`;
const R = (text: string, rPr?: string): string => `<w:r>${rPr ?? ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;

function runOn(bodyXml: string, ops: Op[]): { receipts: Receipt[]; out: Uint8Array } {
  const { docx: norm } = normalizeDocx(docx(bodyXml));
  const doc = openDocument(norm);
  const receipts = doc.apply(ops, { date: DATE });
  return { receipts, out: doc.save() };
}

describe("review round 1 regression tests", () => {
  it("M1: two zero-length replacements at one offset become sibling ins revisions with independent receipts", () => {
    const body = P("10000001", R("Payments begin on the start date."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const op = (flag: string, text: string): Op => ({
      op: "replace", flagId: flag, author: `Scope Optimizer [${flag}]`, addr: "10000001",
      start: 16, end: 16, expectText: "", text,
    });
    const receipts = doc.apply([op("f1", "AAA "), op("f2", "BBB ")], { date: DATE });
    expect(receipts.every((r) => r.status === "applied")).toBe(true);
    const out1 = doc.save();
    // sibling: no w:ins directly inside another w:ins
    const xml = new TextDecoder().decode(readPackage(out1).get("word/document.xml")!);
    expect(xml).not.toMatch(/<w:ins [^>]*><w:ins /);
    // independent reject of f1 leaves f2's text
    const doc2 = openDocument(out1);
    doc2.decide([{ author: "Scope Optimizer [f1]", decision: "reject" }]);
    const text = doc2.projection().find((p) => p.addr === "10000001")!.text;
    expect(text).toContain("BBB ");
    expect(text).not.toContain("AAA ");
  });

  it("M2: span crossing from plain text into a hyperlink field result refuses span_crosses_structure", () => {
    // hyperlink containing a complex field whose result text is inside the link
    const body = P("10000001",
      R("See ") +
      `<w:hyperlink r:id="r1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      R("the portal") +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>` +
      R(" for details."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const para = doc.projectionRecords()[0]!;
    expect(para.text).toBe("See the portal for details.");
    const map = buildSpanMap(para.node, () => "accept");
    // span crosses from plain text into the hyperlink's field result
    const issue = checkSpan(map, 4, 16, []);
    expect(issue?.reason).toBe("span_crosses_structure");
  });

  it("M3: span across inline customXml and a sibling run refuses, no partial mutation", () => {
    const body = P("10000001",
      `<w:customXml w:uri="urn:x" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` + R("ab") + `</w:customXml>` + R("cd"));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const para = doc.projectionRecords()[0]!;
    expect(para.text).toBe("abcd");
    const map = buildSpanMap(para.node, () => "accept");
    const issue = checkSpan(map, 1, 3, []); // covers "b" (inside customXml) + "c" (outside)
    expect(issue?.reason).toBe("span_crosses_structure");
    const before = doc.save();
    const { receipts } = (() => {
      const d2 = openDocument(norm);
      const r = d2.apply([{ op: "delete", flagId: "f", author: "Scope Optimizer [f]", addr: para.addr!, start: 1, end: 3, expectText: "bc" }], { date: DATE });
      return { receipts: r };
    })();
    expect(receipts[0]!.status).toBe("refused");
    expect(receipts[0]!.reason).toBe("span_crosses_structure");
    expect(Buffer.from(doc.save()).equals(Buffer.from(before))).toBe(true);
  });

  it("RM1: CR/CRLF in text projects LF-normalized like the oracle", () => {
    const body = P("10000001", `<w:r><w:t>A\r\nB\rC</w:t></w:r>`);
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    expect(doc.projection()[0]!.text).toBe("A\nB\nC");
  });

  it("RM3: normalize reassigns duplicate paraIds deterministically and stays idempotent", () => {
    const body = P("10000001", R("one")) + P("10000001", R("two"));
    const first = normalizeDocx(docx(body));
    expect(first.assigned).toBe(1); // the duplicate reassigned
    const xml = new TextDecoder().decode(readPackage(first.docx).get("word/document.xml")!);
    const ids = [...xml.matchAll(/w14:paraId="([0-9A-F]+)"/g)].map((m) => m[1]!);
    expect(new Set(ids).size).toBe(ids.length);
    const second = normalizeDocx(first.docx);
    expect(Buffer.from(second.docx).equals(Buffer.from(first.docx))).toBe(true);
  });

  it("RM4: paragraphs never emits an empty addr, even pre-normalization (CLI path)", () => {
    const { docx: normalized } = normalizeDocx(docx(P(null, R("no id here"))));
    const doc = openDocument(normalized);
    for (const p of doc.projection()) expect(p.addr).toMatch(/^[0-9A-F]{8}$/);
  });

  it("V1: verify rejects author tampering on a receipt revision", () => {
    const body = P("10000001", R("Hold harmless the indemnified party under all claims."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([{ op: "replace", flagId: "f", author: "Scope Optimizer [f]", addr: "10000001", start: 0, end: 13, expectText: "Hold harmless", text: "Indemnify" }], { date: DATE });
    const out = doc.save();
    expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    // tamper: change the revision's author
    const entries = readPackage(out);
    let xml = new TextDecoder().decode(entries.get("word/document.xml")!);
    xml = xml.replace('w:author="Scope Optimizer [f]"', 'w:author="Other Author"');
    entries.set("word/document.xml", encodePart(xml));
    const tampered = writePackage(entries);
    expect(verifyOutput(norm, tampered, receipts).ok).toBe(false);
  });

  it("V2: acceptWith on an insertion-only edit inserts the member wording", () => {
    const body = P("10000001", R("Payments begin on the start date."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    doc.apply([{ op: "replace", flagId: "f", author: "Scope Optimizer [f]", addr: "10000001", start: 16, end: 16, expectText: "", text: "AAA " }], { date: DATE });
    doc.acceptWithWording({ flagId: "f", author: "Scope Optimizer [f]", decision: "accept", text: "BBB " });
    const text = doc.projection().find((p) => p.addr === "10000001")!.text;
    // offset 16 sits after the first "o" of "on": original text with BBB at 16
    expect(text).toBe("Payments begin oBBB n the start date.");
    const xml = new TextDecoder().decode(readPackage(doc.save()).get("word/document.xml")!);
    expect(xml).not.toContain("Scope Optimizer");
  });

  it("V3: finalize sweep removes move and property-change revisions under the prefix", () => {
    const body = P("10000001",
      `<w:moveFrom w:id="90" w:author="Scope Optimizer [legacy]" w:date="${DATE}">` + R("gone") + `</w:moveFrom>` +
      `<w:r><w:rPr><w:b/><w:rPrChange w:id="91" w:author="Scope Optimizer [legacy]" w:date="${DATE}"><w:rPr/></w:rPrChange></w:rPr>` + R("kept") + `</w:r>`);
    const doc = openDocument(docx(body));
    const swept = doc.decideByPrefix("Scope Optimizer", "reject");
    expect(swept).toBeGreaterThanOrEqual(2);
    const xml = new TextDecoder().decode(readPackage(doc.save()).get("word/document.xml")!);
    expect(xml).not.toContain("Scope Optimizer");
    expect(xml).toContain("gone"); // reject moveFrom restores the text
  });

  it("V4: decide accepts changeIds-based decisions and returns receipts", () => {
    const body = P("10000001", R("The fee is due."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([{ op: "replace", flagId: "f", author: "Scope Optimizer [f]", addr: "10000001", start: 4, end: 7, expectText: "fee", text: "charge" }], { date: DATE });
    const ids = receipts[0]!.changeIds;
    const doc2 = openDocument(doc.save());
    const decideReceipts = doc2.decide([{ changeIds: ids, decision: "reject" }]);
    expect(decideReceipts[0]!.changeIds.length).toBeGreaterThan(0);
    expect(doc2.projection()[0]!.text).toBe("The fee is due.");
  });

  it("P1: engine-authored text with XML-invalid characters is sanitized", () => {
    const body = P("10000001", R("Rate card:"));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([{ op: "replace", flagId: "f", author: "Scope Optimizer [f]", addr: "10000001", start: 0, end: 9, expectText: "Rate card", text: "Rates\u0001\r\n2026" }], { date: DATE });
    expect(receipts[0]!.status).toBe("applied");
    const xml = new TextDecoder().decode(readPackage(doc.save()).get("word/document.xml")!);
    expect(xml).not.toContain("\u0001");
    expect(xml).not.toMatch(/\r/);
  });

  it("S03 word-diff offsets: pure insertion keeps old offsets", () => {
    const edits = wordEdits("Sessions not used within the month are forfeited.", "Sessions not used within the month roll over once and are then forfeited.");
    expect(edits).toEqual([
      { insAt: 35, insText: "roll over once and " },
      { insAt: 39, insText: "then " },
    ]);
  });
});
