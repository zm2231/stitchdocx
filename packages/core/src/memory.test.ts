// The memory architecture: paragraphs stay source until an edit or a decision
// reaches them. Listing, verifying and seeding ids read paragraphs through
// transient views; only the paragraphs a command changes are parsed in place.
// A main part over the size limit is refused before it is inflated, and every
// parse is bounded in nodes, so a document whose tree would not fit is refused
// rather than built.
import { describe, expect, it } from "vitest";
import { encodePart, isPending, openPackage, PACKAGE_LIMITS, PackageLimitError, pendingElements, writePackage, type XElement } from "@stitchdocx/package-io";
import { DOCUMENT_LIMITS, finalizeDocx, normalizeDocx, openDocument, verifyOutput, type Op } from "./index.js";
import type { Document } from "./document.js";
import { acceptAll, projectParagraphs } from "./projection.js";
import type { BodyWalk } from "./body.js";

const DATE = "2026-10-03T00:00:00Z";
const AUTHOR = "Reviewer A";
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** a package with the given body and, optionally, a header and a styles part */
const docx = (body: string, extra: { header?: string; styles?: string } = {}): Uint8Array => {
  const parts = new Map([
    ["[Content_Types].xml", encodePart(`${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
    ["word/document.xml", encodePart(`${DECL}<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`)],
  ]);
  if (extra.header !== undefined) {
    parts.set("word/_rels/document.xml.rels", encodePart(`${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/header" Target="header1.xml"/></Relationships>`));
    parts.set("word/header1.xml", encodePart(`${DECL}<w:hdr ${W}>${extra.header}</w:hdr>`));
  }
  if (extra.styles !== undefined) parts.set("word/styles.xml", encodePart(`${DECL}<w:styles ${W}>${extra.styles}</w:styles>`));
  return writePackage(parts);
};

const refused = (fn: () => unknown): string => {
  try {
    fn();
  } catch (e) {
    return e instanceof PackageLimitError ? e.code : String(e);
  }
  return "accepted";
};

/** run `fn` with one limit lowered */
const withLimit = <L extends Record<K, number>, K extends string>(limits: L, key: K, value: number, fn: () => void): void => {
  const saved = limits[key];
  limits[key] = value as L[K];
  try {
    fn();
  } finally {
    limits[key] = saved;
  }
};
const withMainLimit = (value: number, fn: () => void) => withLimit(DOCUMENT_LIMITS, "maxMainPartNodes", value, fn);
const withPartLimit = (value: number, fn: () => void) => withLimit(PACKAGE_LIMITS, "maxPartNodes", value, fn);

const P = (id: string, inner: string): string => `<w:p w14:paraId="${id}">${inner}</w:p>`;
const R = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const fc = (t: string): string => `<w:r><w:fldChar w:fldCharType="${t}"/></w:r>`;

const body =
  P("10000001", R("Fees are $10,000.")) +
  `<w:tbl><w:tr><w:tc>${P("10000002", R("Cell text."))}</w:tc></w:tr></w:tbl>` +
  P("10000003", `${fc("begin")}<w:r><w:instrText> TOC </w:instrText></w:r>${fc("separate")}${R("1. Scope")}`) +
  P("10000004", `${R("2. Terms")}${fc("end")}`) +
  P("10000005", R("Payment within 30 days."));

/** the ids of body paragraphs parsed in place */
const parsed = (doc: Document): string[] => {
  const root = (doc as unknown as { root: XElement }).root;
  const pending = new Set([...pendingElements(root)]);
  const out: string[] = [];
  const visit = (n: XElement): void => {
    if (n.name === "w:p" && !pending.has(n)) out.push(n.attrs.find((a) => a.name === "w14:paraId")?.raw ?? "?");
    if (isPending(n)) return;
    for (const c of n.children) if (c.kind === "el") visit(c);
  };
  visit(root);
  return out;
};

describe("paragraphs stay source until a command changes them", () => {
  it("listing under any view parses no paragraph in place", () => {
    const doc = openDocument(normalizeDocx(docx(body)).docx);
    expect(doc.projection().map((p) => p.text)).toEqual(["Fees are $10,000.", "Cell text.", "1. Scope", "2. Terms", "Payment within 30 days."]);
    doc.projectionRecords(() => "reject");
    expect(parsed(doc)).toEqual([]);
  });

  it("apply parses only the paragraphs its ops change, and save copies the rest verbatim", () => {
    const doc = openDocument(normalizeDocx(docx(body)).docx);
    const ops: Op[] = [{ op: "replace", ref: "f", author: AUTHOR, addr: "10000005", start: 15, end: 17, expectText: "30", text: "45" }];
    const receipts = doc.apply(ops, { date: DATE });
    expect(receipts[0]!.status).toBe("applied");
    expect(parsed(doc)).toEqual(["10000005"]);
    const out = doc.save();
    expect(verifyOutput(normalizeDocx(docx(body)).docx, out, receipts).ok).toBe(true);
  });

  it("a decision parses only the paragraph holding its changes", () => {
    const doc = openDocument(normalizeDocx(docx(body)).docx);
    const receipts = doc.apply([{ op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 9, end: 16, expectText: "$10,000", text: "$12,000" }], { date: DATE });
    const redline = openDocument(doc.save());
    redline.decide([{ ref: "f", changeIds: receipts[0]!.changeIds, decision: "accept" }]);
    expect(parsed(redline)).toEqual(["10000001"]);
    expect(openDocument(redline.save()).projection()[0]!.text).toBe("Fees are $12,000.");
  });

  it("a paragraph that is not well-formed fails the first read that reaches it", () => {
    const doc = openDocument(docx(P("10000001", "<w:r><w:t>open</w:r>")));
    expect(() => doc.projection()).toThrow(/xml parse/);
  });
});

describe("document_too_large", () => {
  const commands = (bytes: Uint8Array): string[] => [
    refused(() => normalizeDocx(bytes)),
    refused(() => openDocument(bytes)),
    refused(() => finalizeDocx(bytes, [])),
    refused(() => verifyOutput(bytes, bytes, [])),
  ];

  it("every command refuses a main part over the byte limit before inflating it", () => {
    const bytes = docx(P("10000001", R("Fees are $10,000.")));
    withLimit(PACKAGE_LIMITS, "maxXmlPartBytes", 100, () => expect(commands(bytes)).toEqual(Array(4).fill("document_too_large")));
  });

  it("every command that reads styles refuses a styles part over the byte limit, the main part under it", () => {
    const styles = `<w:style w:styleId="Normal"/>`.repeat(200);
    const bytes = docx(P("10000001", R("Fees are $10,000.")), { styles });
    withLimit(PACKAGE_LIMITS, "maxXmlPartBytes", 2000, () => expect(commands(bytes)).toEqual(["accepted", ...Array(3).fill("document_too_large")]));
  });

  it("a paragraph nested past the depth limit is refused, not failed, once it is read", () => {
    const deep = "<w:sdt><w:sdtContent>".repeat(300) + R("x") + "</w:sdtContent></w:sdt>".repeat(300);
    const bytes = docx(P("10000001", deep));
    expect(refused(() => openDocument(bytes).projection())).toBe("document_too_large");
    expect(refused(() => verifyOutput(bytes, bytes, []))).toBe("document_too_large");
  });
});

describe("every parse is bounded in nodes", () => {
  const cells = (n: number): string => `<w:tbl><w:tr>${Array.from({ length: n }, () => "<w:tc><w:p/></w:tc>").join("")}</w:tr></w:tbl>`;

  it("a dense grid of tiny cells is refused for its node count, though its bytes are within the limit", () => {
    const bytes = docx(P("10000001", R("Fees apply.")) + cells(200));
    expect(refused(() => openDocument(normalizeDocx(bytes).docx))).toBe("accepted");
    withMainLimit(300, () => {
      expect(refused(() => normalizeDocx(bytes))).toBe("document_too_large");
      expect(refused(() => openDocument(bytes))).toBe("document_too_large");
      expect(refused(() => verifyOutput(bytes, bytes, []))).toBe("document_too_large");
    });
  });

  it("a transient read of one paragraph is bounded on its own", () => {
    const bytes = normalizeDocx(docx(P("10000001", R("a").repeat(40)) + P("10000002", R("Fees apply.")))).docx;
    withMainLimit(100, () => {
      const doc = openDocument(bytes);
      expect(refused(() => doc.projection())).toBe("document_too_large");
    });
  });

  it("paragraphs parsed in place draw on the main part's budget, and spending it mid-batch poisons the document", () => {
    const ids = Array.from({ length: 20 }, (_, i) => String(10000001 + i));
    const bytes = normalizeDocx(docx(ids.map((id) => P(id, R("Fee one. ").repeat(5))).join(""))).docx;
    const op = (addr: string): Op => ({ op: "replace", ref: addr, author: AUTHOR, addr, start: 0, end: 3, expectText: "Fee", text: "Cost" });
    withMainLimit(60, () => {
      const one = openDocument(bytes);
      expect(one.apply([op(ids[0]!)], { date: DATE })[0]!.status).toBe("applied");
      const all = openDocument(bytes);
      expect(refused(() => all.apply(ids.map(op), { date: DATE }))).toBe("document_too_large");
      expect(() => all.save()).toThrow(/unusable after a failed edit/);
    });
  });

  it("a story part is parsed only when its ids need repair, and then bounded", () => {
    const big = Array.from({ length: 50 }, (_, i) => P(String(20000001 + i), R("Header."))).join("");
    withPartLimit(100, () => {
      expect(refused(() => normalizeDocx(docx(P("10000001", R("Body.")), { header: big })))).toBe("accepted");
      const dup = P("10000001", R("Header.")) + big;
      expect(refused(() => normalizeDocx(docx(P("10000001", R("Body.")), { header: dup })))).toBe("document_too_large");
    });
  });

  it("styles are bounded", () => {
    const styles = Array.from({ length: 50 }, (_, i) => `<w:style w:type="paragraph" w:styleId="S${i}"><w:name w:val="S${i}"/></w:style>`).join("");
    const bytes = normalizeDocx(docx(P("10000001", R("Body.")), { styles })).docx;
    expect(refused(() => openDocument(bytes))).toBe("accepted");
    withPartLimit(100, () => expect(refused(() => openDocument(bytes))).toBe("document_too_large"));
  });
});

describe("normalize repairs story-part ids in document order", () => {
  it("the first occurrence in a part keeps its id, and the part keeps its declaration", () => {
    const header = P("30000001", R("First.")) + P("30000001", R("Second."));
    const out = normalizeDocx(docx(P("10000001", R("Body.")), { header }));
    expect(out.assigned).toBe(1);
    const xml = openPackage(out.docx).xml("word/header1.xml")!;
    expect(xml.startsWith(DECL)).toBe(true);
    const ids = [...xml.matchAll(/w14:paraId="([0-9A-F]+)"/g)].map((m) => m[1]);
    expect(ids[0]).toBe("30000001");
    expect(ids[1]).not.toBe("30000001");
  });
});

describe("a decision parses only paragraphs naming its ids as w:id", () => {
  it("an attribute that merely equals the id leaves a paragraph unparsed", () => {
    const doc = openDocument(normalizeDocx(docx(body)).docx);
    const receipts = doc.apply([{ op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 9, end: 16, expectText: "$10,000", text: "$12,000" }], { date: DATE });
    const ids = receipts[0]!.changeIds!;
    const decoy = P("10000009", `<w:pPr><w:spacing w:after="${ids[0]}"/></w:pPr>${R("Decoy.")}`);
    const redline = openDocument(normalizeDocx(docx(body + decoy)).docx);
    const r2 = redline.apply([{ op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 9, end: 16, expectText: "$10,000", text: "$12,000" }], { date: DATE });
    const reopened = openDocument(redline.save());
    reopened.decide([{ ref: "f", changeIds: r2[0]!.changeIds, decision: "accept" }]);
    expect(parsed(reopened)).toEqual(["10000001"]);
  });
});

describe("a projection restricted to some records", () => {
  it("still threads a field instruction from a paragraph it does not read for text", () => {
    const fieldBody =
      P("10000001", `${fc("begin")}<w:r><w:instrText> HYPERLINK </w:instrText></w:r>`) +
      P("10000002", `${R("Instruction.")}${fc("separate")}${R("Shown.")}${fc("end")}`) +
      P("10000003", R("Plain."));
    const doc = openDocument(normalizeDocx(docx(fieldBody)).docx);
    const walk = (doc as unknown as { walk: BodyWalk }).walk;
    expect(projectParagraphs(walk, acceptAll, doc.styles).map((r) => r.text)).toEqual(["", "Shown.", "Plain."]);
    expect(projectParagraphs(walk, acceptAll, doc.styles, new Set(["10000002"])).map((r) => r.text)).toEqual(["Shown."]);
  });

  it("passes over paragraphs that cannot reach the ones asked for, without reading them", () => {
    // the middle paragraph is not well-formed: any read of it throws
    const doc = openDocument(docx(P("10000001", R("Fees.")) + P("10000002", "<w:r><w:t>open</w:r>") + P("10000003", R("Terms."))));
    const walk = (doc as unknown as { walk: BodyWalk }).walk;
    expect(() => projectParagraphs(walk, acceptAll, doc.styles)).toThrow(/xml parse/);
    expect(projectParagraphs(walk, acceptAll, doc.styles, new Set(["10000003"])).map((r) => r.text)).toEqual(["Terms."]);
  });
});
