// Real-world markup that the first two rounds of hardening missed: fields
// that span paragraphs, hidden runs, symbols, tracked table rows, content
// controls around rows and cells, per-part revision and paragraph ids, whole
// span of an op under accept-with-wording, package-level verify, and input
// limits. Builders construct minimal OOXML packages in-memory.
import { describe, expect, it } from "vitest";
import { encodePart, entryBytes, PackageLimitError, readPackage, writePackage } from "@stitchdocx/package-io";
import { finalizeDocx, normalizeDocx, openDocument, verifyOutput, type Op } from "./index.js";
import { cloneRunPr } from "./revisions.js";
import { el, kid, textNode } from "@stitchdocx/package-io";

const DATE = "2026-10-03T00:00:00Z";
const AUTHOR = "Reviewer A";
const CLIENT = "Client Author";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15"';

function entriesOf(body: string, extra: Record<string, string> = {}): Map<string, Uint8Array> {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  const entries = new Map<string, Uint8Array>([
    ["[Content_Types].xml", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${Object.keys(extra).filter((k) => k.startsWith("word/") && k.endsWith(".xml")).map((k) => `<Override PartName="/${k}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${k.includes("numbering") ? "numbering" : k.includes("comments") ? "comments" : k.includes("styles") ? "styles" : "header"}+xml"/>`).join("")}</Types>`)],
    ["word/document.xml", encodePart(xml)],
    ["word/_rels/document.xml.rels", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${Object.keys(extra).filter((k) => k.startsWith("word/") && k.endsWith(".xml")).map((k, i) => `<Relationship Id="rX${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${k.includes("numbering") ? "numbering" : k.includes("comments") ? "comments" : k.includes("styles") ? "styles" : "header"}" Target="${k.slice(5)}"/>`).join("")}</Relationships>`)],
  ]);
  for (const [k, v] of Object.entries(extra)) entries.set(k, encodePart(v));
  return entries;
}

const docx = (body: string, extra?: Record<string, string>): Uint8Array => writePackage(entriesOf(body, extra));

const P = (id: string, inner: string, pPr = ""): string => `<w:p w14:paraId="${id}">${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${inner}</w:p>`;
const R = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const runApply = (bodyXml: string, ops: Op[], extra?: Record<string, string>) => {
  const { docx: norm } = normalizeDocx(docx(bodyXml, extra));
  const doc = openDocument(norm);
  const receipts = doc.apply(ops, { date: DATE });
  return { doc, receipts, out: doc.save(), norm };
};

const textsOf = (bytes: Uint8Array): string[] => openDocument(bytes).projection().map((p) => [p.text, p.editable] as const).map((t) => `${t[0]}|${t[1]}`);

describe("fields spanning paragraphs and simple fields are never editable (F3)", () => {
  const fc = (t: string): string => `<w:r><w:fldChar w:fldCharType="${t}"/></w:r>`;
  const body =
    P("10000001", `${R("Client: ")}<w:fldSimple w:instr=" DOCPROPERTY Company "><w:r><w:t>Acme Corp</w:t></w:r></w:fldSimple>`) +
    P("10000002", `${fc("begin")}<w:r><w:instrText xml:space="preserve"> TOC \\o </w:instrText></w:r>${fc("separate")}${R("1. Scope")}`) +
    P("10000003", R("2. Payment")) +
    P("10000004", `${R("3. Termination")}${fc("end")}`) +
    P("10000005", R("Body text."));

  it("field text is listed in reader order but every span inside a field refuses", () => {
    const doc = openDocument(docx(body));
    expect(doc.projection().map((p) => p.text)).toEqual(["Client: Acme Corp", "1. Scope", "2. Payment", "3. Termination", "Body text."]);
    const { receipts } = runApply(body, [
      { op: "replace", ref: "f1", author: AUTHOR, addr: "10000001", start: 8, end: 12, expectText: "Acme", text: "Beta" },
      { op: "replace", ref: "f2", author: AUTHOR, addr: "10000002", start: 0, end: 1, expectText: "1", text: "2" },
      { op: "replace", ref: "f3", author: AUTHOR, addr: "10000003", start: 0, end: 1, expectText: "2", text: "3" },
      { op: "replace", ref: "f4", author: AUTHOR, addr: "10000004", start: 0, end: 1, expectText: "3", text: "4" },
      { op: "replace", ref: "f5", author: AUTHOR, addr: "10000005", start: 0, end: 4, expectText: "Body", text: "Text" },
    ]);
    expect(receipts.map((r) => `${r.ref}:${r.status}${r.reason ? "/" + r.reason : ""}`)).toEqual([
      "f1:refused/span_in_field",
      "f2:refused/span_in_field",
      "f3:refused/span_in_field",
      "f4:refused/span_in_field",
      "f5:applied",
    ]);
  });

  it("field instructions across paragraphs stay hidden, results show", () => {
    const cross = P("10000001", fc("begin") + R("hidden instr") + fc("separate")) + P("10000002", R("shown result") + fc("end"));
    expect(openDocument(docx(cross)).projection().map((p) => p.text)).toEqual(["", "shown result"]);
    const instrOnly = P("10000001", fc("begin") + R("hidden instr")) + P("10000002", R("still instr") + fc("end"));
    expect(openDocument(docx(instrOnly)).projection().map((p) => p.text)).toEqual(["", ""]);
  });
});

describe("hidden runs are neither listed nor formatted onto insertions (F4)", () => {
  it("a hidden drafting note is invisible to the listing and to offsets", () => {
    const body = P("10000001", R("Payment due in 30 days.") + `<w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space="preserve"> [note: late fee]</w:t></w:r>`);
    const doc = openDocument(docx(body));
    expect(doc.projection().map((p) => p.text)).toEqual(["Payment due in 30 days."]);
    // the hidden note's offsets do not exist: an edit of the visible text
    // applies and leaves the hidden run byte-identical
    const { receipts, out, norm } = runApply(body, [
      { op: "replace", ref: "g", author: AUTHOR, addr: "10000001", start: 15, end: 17, expectText: "30", text: "45" },
    ]);
    expect(receipts[0]!.status).toBe("applied");
    const before = new TextDecoder().decode(entryBytes(readPackage(norm).get("word/document.xml")!)).match(/<w:r><w:rPr><w:vanish\/>[\s\S]*?<\/w:r>/)![0];
    const after = new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!)).match(/<w:r><w:rPr><w:vanish\/>[\s\S]*?<\/w:r>/)![0];
    expect(after).toBe(before);
    expect(verifyOutput(norm, out, receipts).ok).toBe(true);
  });

  it("cloneRunPr never carries hidden properties or revision records", () => {
    const run = el("w:r", {}, [
      (() => {
        const rPr = el("w:rPr", {}, [el("w:b"), el("w:vanish"), el("w:webHidden"), el("w:rPrChange", { "w:id": "7", "w:author": CLIENT, "w:date": DATE })]);
        return rPr;
      })(),
      (() => {
        const t = el("w:t", { "xml:space": "preserve" });
        t.children.push(textNode("x"));
        return t;
      })(),
    ]);
    const clone = cloneRunPr(run);
    expect(clone).not.toBeNull();
    const names = clone!.children.map((c) => (c.kind === "el" ? c.name : "?"));
    expect(names).toEqual(["w:b"]);
  });
});

describe("w:sym is one listed character and tracks like any run content (F11)", () => {
  const body = P("10000001", R("Box ") + `<w:r><w:sym w:font="Wingdings" w:char="F0FE"/></w:r>` + R(" Approved"));

  it("the symbol lists as its private-use code point", () => {
    expect(openDocument(docx(body)).projection().map((p) => p.text)).toEqual(["Box \uF0FE Approved"]);
  });

  it("a delete spanning the symbol tracks it", () => {
    const { receipts, out } = runApply(body, [
      { op: "delete", ref: "u", author: AUTHOR, addr: "10000001", start: 4, end: 5, expectText: "\uF0FE" },
    ]);
    expect(receipts[0]!.status).toBe("applied");
    const xml = new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!));
    expect(/<w:del[^>]*>(?:(?!<\/w:del>)[\s\S])*<w:sym/.test(xml)).toBe(true);
    expect(verifyOutput(normalizeDocx(docx(body)).docx, out, receipts).ok).toBe(true);
  });
});

describe("tracked rows follow Word's accept/reject view (F5)", () => {
  const d = `w:author="${CLIENT}" w:date="${DATE}"`;
  const body =
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>` +
    `<w:tr><w:tc>${P("10000001", R("Phase 1 $40,000"))}</w:tc></w:tr>` +
    `<w:tr><w:trPr><w:del w:id="1" ${d}/></w:trPr><w:tc>${P("10000002", R("Phase 2 $90,000"))}</w:tc></w:tr>` +
    `<w:tr><w:trPr><w:ins w:id="2" ${d}/></w:trPr><w:tc>${P("10000003", R("Phase 3 added"))}</w:tc></w:tr>` +
    `</w:tbl>` + P("10000004", R("End"));

  it("a tracked-deleted row drops from the accept view; a tracked-inserted row stays but is locked", () => {
    const doc = openDocument(docx(body));
    expect(doc.projection().map((p) => [p.text, p.editable])).toEqual([
      ["Phase 1 $40,000", true],
      ["Phase 3 added", true],
      ["End", true],
    ]);
  });

  it("spans inside a tracked row refuse; the untracked row edits", () => {
    const { receipts } = runApply(body, [
      { op: "replace", ref: "r1", author: AUTHOR, addr: "10000003", start: 8, end: 13, expectText: "added", text: "extra" },
      { op: "replace", ref: "r2", author: AUTHOR, addr: "10000001", start: 8, end: 15, expectText: "$40,000", text: "$50,000" },
    ]);
    expect(receipts.map((r) => `${r.ref}:${r.status}${r.reason ? "/" + r.reason : ""}`)).toEqual([
      "r1:refused/span_overlaps_revision",
      "r2:applied",
    ]);
  });

  it("the reject view lists the deleted row again", () => {
    const theirsReject = openDocument(docx(body)).projectionRecords((rev) => (rev.author === CLIENT ? "reject" : "accept"));
    expect(theirsReject.map((p) => p.text)).toEqual(["Phase 1 $40,000", "Phase 2 $90,000", "End"]);
  });
});

describe("content controls around rows and cells stay visible (F6)", () => {
  const tc = (t: string): string => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p>${R(t)}</w:p></w:tc>`;
  const body =
    P("10000001", R("Pricing")) +
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
    `<w:tr>${tc("Item")}${tc("Fee")}</w:tr>` +
    `<w:sdt><w:sdtPr><w15:repeatingSection/></w:sdtPr><w:sdtContent><w:tr>${tc("Discovery")}${tc("$40,000")}</w:tr></w:sdtContent></w:sdt>` +
    `<w:tr><w:sdt><w:sdtPr/><w:sdtContent>${tc("Build")}</w:sdtContent></w:sdt>${tc("$90,000")}</w:tr>` +
    `</w:tbl>` + P("10000002", R("End"));

  it("every cell lists with the right column, and normalize covers the wrapped paragraphs", () => {
    const { docx: norm, assigned } = normalizeDocx(docx(body));
    expect(assigned).toBe(6);
    const listed = openDocument(norm).projection();
    expect(listed.map((p) => p.text)).toEqual(["Pricing", "Item", "Fee", "Discovery", "$40,000", "Build", "$90,000", "End"]);
    const cols = listed.filter((p) => p.table).map((p) => p.table!.col);
    expect(cols).toEqual([0, 1, 0, 1, 0, 1]);
  });
});

describe("revision ids seed across every part and revision type (F9)", () => {
  const hdr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:ins w:id="41" w:author="${CLIENT}" w:date="${DATE}"><w:r><w:t>Draft</w:t></w:r></w:ins></w:p></w:hdr>`;
  const body =
    P("10000001", R("Fee is $10,000.")) +
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcPrChange w:id="17" w:author="${CLIENT}" w:date="${DATE}"><w:tcPr/></w:tcPrChange></w:tcPr>${P("10000002", R("Cell"))}</w:tc></w:tr></w:tbl>`;

  it("our ids land above the largest id anywhere, including headers and tcPrChange", () => {
    const { receipts, out } = runApply(body, [
      { op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 7, end: 14, expectText: "$10,000", text: "$12,000" },
    ], { "word/header1.xml": hdr });
    expect(receipts[0]!.changeIds.map(Number)).toEqual([42, 43]);
    const xml = new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!));
    expect(xml).toContain(`w:del w:id="42"`);
  });
});

describe("paragraph ids are unique across parts and always valid (F16)", () => {
  const comments = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="0" w:author="${CLIENT}"><w:p w14:paraId="10000003"><w:r><w:t>Please confirm</w:t></w:r></w:p></w:comment></w:comments>`;

  it("a body id colliding with a comment paragraph is reassigned; invalid ids are repaired", () => {
    const body = P("10000001", R("A")) + P("10000003", R("B")) + P("8000ABCD", R("C")) + P("xyz", R("D"));
    const { docx: norm, assigned } = normalizeDocx(docx(body, { "word/comments.xml": comments }));
    expect(assigned).toBe(3);
    const xml = new TextDecoder().decode(entryBytes(readPackage(norm).get("word/document.xml")!));
    const ids = [...xml.matchAll(/w14:paraId="([^"]+)"/g)].map((m) => m[1]!);
    expect(ids).toContain("10000001");
    // the body keeps its id (the main part plans first); the comment's
    // colliding paragraph is reassigned, and every id is valid 8-hex
    for (const id of ids) expect(id).toMatch(/^[0-9A-F]{8}$/);
    const commentXml = new TextDecoder().decode(entryBytes(readPackage(norm).get("word/comments.xml")!));
    const commentIds = [...commentXml.matchAll(/w14:paraId="([^"]+)"/g)].map((m) => m[1]!);
    expect(commentIds.filter((i) => i === "10000003")).toHaveLength(0);
    expect(new Set([...ids, ...commentIds]).size).toBe(ids.length + commentIds.length);
  });
});

describe("numbering labels match Word (F7)", () => {
  const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="Section %1"/></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride></w:num>
</w:numbering>`;
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="ListHeading"><w:basedOn w:val="Base"/></w:style><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="1"/></w:numPr></w:pPr></w:style></w:styles>`;
  const np = (l: number, n = 1): string => `<w:numPr><w:ilvl w:val="${l}"/><w:numId w:val="${n}"/></w:numPr>`;

  it("level values, %n indices, restarts, formats, overrides and style inheritance", () => {
    const body =
      P("10000001", R("Scope"), np(0)) +
      P("10000002", R("Fees"), np(0)) +
      P("10000003", R("sub one"), np(1)) +
      P("10000004", R("sub two"), np(1)) +
      P("10000005", R("Term"), np(0)) +
      P("10000006", R("Roman"), np(0, 2)) +
      P("10000007", R("Inherited"), `<w:pStyle w:val="ListHeading"/>`);
    const doc = openDocument(docx(body, { "word/numbering.xml": numbering, "word/styles.xml": styles }));
    expect(doc.projection().map((p) => p.numbering)).toEqual(["1.", "2.", "(a)", "(b)", "3.", "Section VII", "(a)"]);
    // a second projection recomputes from scratch
    expect(doc.projection().map((p) => p.numbering)).toEqual(["1.", "2.", "(a)", "(b)", "3.", "Section VII", "(a)"]);
  });
});

describe("accept-with-wording replaces the whole changed span (F1)", () => {
  const applyWording = (body: string, op: Op, text: string) => {
    const first = runApply(body, [op]);
    expect(first.receipts[0]!.status).toBe("applied");
    const expectText = op.op === "insertParagraphAfter" ? undefined : op.expectText;
    const { docx: out } = finalizeDocx(first.out, [{ ref: op.ref, changeIds: first.receipts[0]!.changeIds, decision: "accept", text, expectText }]);
    return { out };
  };

  it("a span-granularity replace reads as the original with the wording in place", () => {
    const body = P("10000001", R("Payment is due within thirty days."));
    const { out } = applyWording(body, { op: "replace", ref: "j", author: AUTHOR, addr: "10000001", start: 11, end: 28, expectText: "due within thirty", text: "payable on receipt" }, "payable on receipt");
    expect(openDocument(out).projection().map((p) => p.text)).toEqual(["Payment is payable on receipt days."]);
    const xml = new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!));
    expect(xml).not.toContain(`w:author="${AUTHOR}"`);
  });

  it("a word-granularity replace replaces every deleted segment and the words between, as one run", () => {
    const body = P("10000001", R("Payment is due within thirty days of invoice receipt."));
    const op: Op = { op: "replace", ref: "j", author: AUTHOR, addr: "10000001", start: 22, end: 28, expectText: "thirty", text: "fifteen", granularity: "word" };
    const { out } = applyWording(body, op, "forty");
    // the changed region ("thirty" -> "fifteen") reads as the wording; the
    // untouched words outside the op's span survive exactly
    expect(openDocument(out).projection().map((p) => p.text)).toEqual(["Payment is due within forty days of invoice receipt."]);
    const xml = new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!));
    expect(xml).not.toContain(`w:author="${AUTHOR}"`);
    expect(xml).not.toContain("w:del");
    expect(xml).not.toContain("w:ins");
  });

  it("the wording replaces the op's whole span, including words the word diff left unchanged", () => {
    const body = P("10000001", R("Payment is due within thirty days of invoice receipt."));
    const span = "due within thirty days of invoice";
    const op: Op = { op: "replace", ref: "j", author: AUTHOR, addr: "10000001", start: 11, end: 11 + span.length, expectText: span, text: "due within fifteen days of the invoice", granularity: "word" };
    const { out } = applyWording(body, op, "payable on receipt");
    expect(openDocument(out).projection().map((p) => p.text)).toEqual(["Payment is payable on receipt receipt."]);
  });

  it("a delete accepted with wording puts the wording where the span was", () => {
    const body = P("10000001", R("Fees are due on receipt of invoice."));
    const op: Op = { op: "delete", ref: "d", author: AUTHOR, addr: "10000001", start: 9, end: 23, expectText: "due on receipt" };
    const { out } = applyWording(body, op, "payable in 30 days");
    expect(openDocument(out).projection().map((p) => p.text)).toEqual(["Fees are payable in 30 days of invoice."]);
  });

  it("wording fails, naming the decision, when expectText is missing or does not cover the change", () => {
    const body = P("10000001", R("Fee A and Fee B apply."));
    const op: Op = { op: "replace", ref: "j", author: AUTHOR, addr: "10000001", start: 14, end: 15, expectText: "B", text: "C" };
    const first = runApply(body, [op]);
    const ids = first.receipts[0]!.changeIds;
    expect(() => finalizeDocx(first.out, [{ ref: "j", changeIds: ids, decision: "accept", text: "X" }])).toThrow(/j\) accepts with wording but gives no expectText/);
    expect(() => finalizeDocx(first.out, [{ ref: "j", changeIds: ids, decision: "accept", text: "X", expectText: "Fee A" }])).toThrow(/covers its changes nowhere/);
    expect(() => finalizeDocx(first.out, [{ ref: "j", changeIds: ["999"], decision: "accept" }])).toThrow(/change 999, which is not a revision/);
  });
});

describe("verify is a whole-package gate (F8)", () => {
  const hdr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>ACME Confidential</w:t></w:r></w:p></w:hdr>`;
  const body = P("10000001", R("Fee is $10,000 per month.")) + P("10000002", R("End"));
  const apply1 = () =>
    runApply(body, [{ op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 7, end: 14, expectText: "$10,000", text: "$12,000" }], { "word/header1.xml": hdr });

  it("an honest output verifies", () => {
    const { norm, out, receipts } = apply1();
    expect(verifyOutput(norm, out, receipts).ok).toBe(true);
  });

  it("a tampered header, a hidden run in the touched paragraph, or an extra zip entry each fail", () => {
    const { norm, out, receipts } = apply1();
    const decode = (bytes: Uint8Array): Map<string, Uint8Array> => new Map([...readPackage(bytes)].map(([n, b]) => [n, b] as [string, Uint8Array]));

    const hdrTampered = decode(out);
    hdrTampered.set("word/header1.xml", encodePart(hdr.replace("ACME Confidential", "ACME Public")));
    expect(verifyOutput(norm, writePackage(hdrTampered), receipts).ok).toBe(false);

    const touchedTampered = decode(out);
    const xml = new TextDecoder().decode(touchedTampered.get("word/document.xml")!);
    const injected = xml.replace("</w:p>", '<w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space="preserve"></w:t></w:r></w:p>');
    touchedTampered.set("word/document.xml", encodePart(injected));
    expect(verifyOutput(norm, writePackage(touchedTampered), receipts).ok).toBe(false);

    const extraEntry = decode(out);
    extraEntry.set("word/evil.bin", encodePart("payload"));
    expect(verifyOutput(norm, writePackage(extraEntry), receipts).ok).toBe(false);
  });
});

describe("verify checks inserted paragraphs by markup, not text", () => {
  const body = P("10000001", R("Scope.")) + P("10000002", R("Fees."));
  const insertOne = () => runApply(body, [{ op: "insertParagraphAfter", ref: "i", author: AUTHOR, addr: "10000001", text: "New clause." }]);
  const tamper = (out: Uint8Array, addr: string, extra: string): Uint8Array => {
    const entries = readPackage(out);
    const xml = new TextDecoder().decode(entryBytes(entries.get("word/document.xml")!));
    const at = xml.indexOf("</w:p>", xml.indexOf(`w14:paraId="${addr}"`));
    entries.set("word/document.xml", encodePart(xml.slice(0, at) + extra + xml.slice(at)));
    return writePackage(entries);
  };

  it("an honest insert verifies", () => {
    const { norm, out, receipts } = insertOne();
    expect(verifyOutput(norm, out, receipts).ok).toBe(true);
  });

  it("hidden text or an empty run added to the inserted paragraph fails", () => {
    const { norm, out, receipts } = insertOne();
    const addr = receipts[0]!.insertedAddr!;
    const hidden = verifyOutput(norm, tamper(out, addr, '<w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r>'), receipts);
    expect(hidden.ok).toBe(false);
    expect(hidden.failures.join(" ")).toMatch(/outside its revisions/);
    expect(verifyOutput(norm, tamper(out, addr, "<w:r><w:rPr><w:b/></w:rPr></w:r>"), receipts).ok).toBe(false);
  });

  it("hidden formatting inside our insertion fails", () => {
    const { norm, out, receipts } = insertOne();
    const entries = readPackage(out);
    const xml = new TextDecoder().decode(entryBytes(entries.get("word/document.xml")!));
    const hiddenIns = xml.replace(/(<w:ins [^>]*w:author="[^"]*"[^>]*><w:r>)(<w:rPr>)?/, (m, head, rPr) => `${head}<w:rPr><w:vanish/>${rPr ? "" : "</w:rPr>"}`);
    expect(hiddenIns).not.toBe(xml);
    entries.set("word/document.xml", encodePart(hiddenIns));
    const result = verifyOutput(norm, writePackage(entries), receipts);
    expect(result.ok).toBe(false);
  });
});

describe("input limits refuse hostile packages with typed errors (F13)", () => {
  const base = docx(P("10000001", R("Fee is $10,000.")));

  it("duplicate entry names are refused", () => {
    // a Map cannot hold a name twice: write a same-length stand-in name, then
    // rename it in the local header and the central directory
    const entries = readPackage(base);
    entries.set("word/documenX.xml", entries.get("word/document.xml")!);
    const bytes = writePackage(entries);
    const from = new TextEncoder().encode("word/documenX.xml");
    const to = new TextEncoder().encode("word/document.xml");
    let renamed = 0;
    for (let i = 0; i + from.length <= bytes.length; i++) {
      if (from.every((b, k) => bytes[i + k] === b)) {
        bytes.set(to, i);
        renamed++;
      }
    }
    expect(renamed).toBe(2);
    expect(() => normalizeDocx(bytes)).toThrow(PackageLimitError);
    expect(() => normalizeDocx(bytes)).toThrow(/duplicate/);
  });

  it("an entry over the per-entry limit is refused", () => {
    const entries = readPackage(base);
    entries.set("word/media/pad.bin", new Uint8Array(65 * 1024 * 1024));
    expect(() => normalizeDocx(writePackage(entries))).toThrow(PackageLimitError);
  });
});

describe("explicit paragraph scope and surrogate offsets (F2, F10)", () => {
  it('scope "paragraph" hits the same refusals as no scope', () => {
    const control = (p: string): string => `<w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent>${p}</w:sdtContent></w:sdt>`;
    const body = control(P("10000001", R("Client name"))) + P("10000002", R("After"));
    for (const scope of [undefined, "paragraph"] as const) {
      const { receipts } = runApply(body, [
        { op: "insertParagraphAfter", ref: "a", author: AUTHOR, addr: "10000001", text: "New para", ...(scope ? { scope } : {}) },
      ]);
      expect(receipts[0]!.status).toBe("refused");
      expect(receipts[0]!.reason).toBe("insert_in_plain_text_control");
    }
  });

  it("an offset inside a surrogate pair refuses as out of range", () => {
    const body = P("10000001", R("Rate \u{1F600} applies"));
    const listed = openDocument(docx(body)).projection()[0]!.text;
    expect(listed).toBe("Rate \u{1F600} applies");
    const { receipts } = runApply(body, [
      { op: "delete", ref: "y", author: AUTHOR, addr: "10000001", start: 6, end: 15, expectText: listed.slice(6, 15) },
    ]);
    expect(receipts[0]!.status).toBe("refused");
    expect(receipts[0]!.reason).toBe("span_out_of_range");
    // whole-character offsets still work
    const ok = runApply(body, [
      { op: "replace", ref: "z", author: AUTHOR, addr: "10000001", start: 5, end: 7, expectText: "\u{1F600}", text: "😎" },
    ]);
    expect(ok.receipts[0]!.status).toBe("applied");
    expect(verifyOutput(normalizeDocx(docx(body)).docx, ok.out, ok.receipts).ok).toBe(true);
  });
});

describe("AlternateContent branches and office math are listed but locked (F17, F18)", () => {
  it("a supported Choice is read once and its spans refuse", () => {
    const body =
      P("10000001", R("Before ") + `<mc:AlternateContent><mc:Choice Requires="w14"><w:r><w:t>chosen words</w:t></w:r></mc:Choice><mc:Choice Requires="wps"><w:r><w:t>shape fallback</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>fallback words</w:t></w:r></mc:Fallback></mc:AlternateContent>` + R(" after"));
    const doc = openDocument(docx(body));
    expect(doc.projection().map((p) => p.text)).toEqual(["Before chosen words after"]);
    const { receipts } = runApply(body, [
      { op: "replace", ref: "c1", author: AUTHOR, addr: "10000001", start: 7, end: 13, expectText: "chosen", text: "X" },
      { op: "replace", ref: "c2", author: AUTHOR, addr: "10000001", start: 0, end: 6, expectText: "Before", text: "Y" },
    ]);
    expect(receipts.map((r) => `${r.ref}:${r.status}${r.reason ? "/" + r.reason : ""}`)).toEqual([
      "c1:refused/span_crosses_structure",
      "c2:applied",
    ]);
  });

  it("a Choice requiring wps is the branch Word renders", () => {
    const body = P("10000001", R("Fee ") + `<mc:AlternateContent><mc:Choice Requires="wps"><w:r><w:t>shape</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>picture</w:t></w:r></mc:Fallback></mc:AlternateContent>`);
    expect(openDocument(docx(body)).projection().map((p) => p.text)).toEqual(["Fee shape"]);
  });

  it("a w16se:symEx reads as its character", () => {
    const body = P("10000001", R("Rate ") + `<w:r><mc:AlternateContent><mc:Choice Requires="w16se"><w16se:symEx xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" w16se:font="Segoe UI Emoji" w16se:char="1F600"/></mc:Choice><mc:Fallback><w:t>?</w:t></mc:Fallback></mc:AlternateContent></w:r>` + R(" applies."));
    expect(openDocument(docx(body)).projection().map((p) => p.text)).toEqual(["Rate \u{1F600} applies."]);
  });

  it("an unsupported Requires falls back, and math text lists locked", () => {
    const body =
      P("10000001", R("Fee ") + `<mc:AlternateContent><mc:Choice Requires="xyz"><w:r><w:t>shape</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>picture</w:t></w:r></mc:Fallback></mc:AlternateContent>` + R(" and ") + `<m:oMath><m:r><m:t>E=mc^2</m:t></m:r></m:oMath>`) + P("10000002", R("End"));
    const doc = openDocument(docx(body));
    expect(doc.projection().map((p) => p.text)).toEqual(["Fee picture and E=mc^2", "End"]);
    const { receipts } = runApply(body, [
      { op: "replace", ref: "m1", author: AUTHOR, addr: "10000001", start: 16, end: 19, expectText: "E=m", text: "X" },
      { op: "replace", ref: "m2", author: AUTHOR, addr: "10000002", start: 0, end: 3, expectText: "End", text: "Fin" },
    ]);
    expect(receipts.map((r) => `${r.ref}:${r.status}${r.reason ? "/" + r.reason : ""}`)).toEqual([
      "m1:refused/span_crosses_structure",
      "m2:applied",
    ]);
  });
});

describe("hidden text follows Word's style rules (R4)", () => {
  const cs = (id: string, based: string, rPr: string) => `<w:style w:type="character" w:styleId="${id}"><w:name w:val="${id}"/>${based ? `<w:basedOn w:val="${based}"/>` : ""}<w:rPr>${rPr}</w:rPr></w:style>`;
  const ps = (id: string, based: string, rPr: string, def = false) => `<w:style w:type="paragraph"${def ? ' w:default="1"' : ""} w:styleId="${id}"><w:name w:val="${id}"/>${based ? `<w:basedOn w:val="${based}"/>` : ""}<w:rPr>${rPr}</w:rPr></w:style>`;
  const V = "<w:vanish/>";
  const V0 = '<w:vanish w:val="0"/>';
  const styles = (inner: string, defaults = "") =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${defaults}${inner}</w:styles>`;
  const run = (text: string, rPr = "") => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const rs = (id: string) => `<w:rStyle w:val="${id}"/>`;
  const pst = (id: string) => `<w:pStyle w:val="${id}"/>`;

  // expected values are Word's (font.hidden and a PDF render of the same document)
  it("resolves direct, character style, paragraph style and docDefaults as Word does", () => {
    const sheet = styles(
      ps("Normal", "", "", true) + cs("CV", "", V) + cs("CV2", "CV", "") + cs("CV3", "CV", V) + cs("CV0", "CV", V0) +
        ps("PV", "", V) + ps("PV2", "PV", "") + ps("PV3", "PV", V) + ps("PV0", "PV", V0),
    );
    const cases: Array<[string, string, string, boolean]> = [
      ["", rs("CV"), "a01", true], ["", rs("CV2"), "a02", true], ["", rs("CV3"), "a03", true], ["", rs("CV0"), "a04", false],
      [pst("PV"), "", "a05", true], [pst("PV2"), "", "a06", true], [pst("PV3"), "", "a07", true], [pst("PV0"), "", "a08", false],
      [pst("PV"), rs("CV"), "a09", false], [pst("PV"), V0, "a10", false], [pst("PV"), rs("CV0"), "a11", true],
      ["", rs("CV") + V, "a12", true], ["", rs("CV") + V0, "a13", false], [pst("PV"), V, "a14", true], ["", "", "a15", false],
      [pst("PV3"), rs("CV"), "a16", false], ["", "<w:specVanish/>", "a17", false], ["", "<w:webHidden/>", "a18", false],
    ];
    const body = cases.map(([pPr, rPr, t], i) => P(String(10000001 + i), run(`${t} `, rPr) + run("x", V0), pPr)).join("");
    const listed = openDocument(docx(body, { "word/styles.xml": sheet })).projection().map((p) => p.text);
    expect(listed).toEqual(cases.map(([, , t, hidden]) => (hidden ? "x" : `${t} x`)));
  });

  it("uses the default paragraph style and docDefaults", () => {
    const b = styles(ps("Normal", "", V, true) + ps("Other", "", "") + ps("Based", "Normal", "") + cs("CV", "", V));
    const bBody = P("10000001", run("b01 ") + run("x")) + P("10000002", run("b02 ") + run("x"), pst("Other")) +
      P("10000003", run("b03 ") + run("x"), pst("Based")) + P("10000004", run("b04 ", rs("CV")) + run("x", V0)) +
      P("10000005", run("b05 ") + run("x", V0), pst("Missing"));
    expect(openDocument(docx(bBody, { "word/styles.xml": b })).projection().map((p) => p.text)).toEqual(["", "b02 x", "", "b04 x", "x"]);
    const c = styles(ps("Normal", "", "", true) + ps("PV", "", V) + cs("CV", "", V) + ps("P0", "", V0), `<w:docDefaults><w:rPrDefault><w:rPr>${V}</w:rPr></w:rPrDefault></w:docDefaults>`);
    const cBody = P("10000001", run("c01 ") + run("x", V0)) + P("10000002", run("c02 ") + run("x", V0), pst("PV")) +
      P("10000003", run("c03 ", rs("CV")) + run("x", V0)) + P("10000004", run("c04 ") + run("x"), pst("P0")) + P("10000005", run("c05 ", V0));
    expect(openDocument(docx(cBody, { "word/styles.xml": c })).projection().map((p) => p.text)).toEqual(["x", "x", "x", "c04 x", "c05 "]);
  });

  it("our insertions stay visible where the styles would hide them, and verify checks it", () => {
    const sheet = styles(ps("Normal", "", "", true) + cs("CV", "", V));
    const body = P("10000001", run("Fee is due.", rs("CV") + V0)) + P("10000002", run("End"));
    const { norm, out, receipts } = runApply(body, [
      { op: "replace", ref: "r", author: AUTHOR, addr: "10000001", start: 7, end: 10, expectText: "due", text: "owed" },
      { op: "insertParagraphAfter", ref: "i", author: AUTHOR, addr: "10000001", text: "New clause." },
    ], { "word/styles.xml": sheet });
    expect(receipts.map((r) => r.status)).toEqual(["applied", "applied"]);
    expect(openDocument(out).projection().map((p) => p.text)).toEqual(["Fee is owed.", "New clause.", "End"]);
    expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    const entries = readPackage(out);
    const xml = new TextDecoder().decode(entryBytes(entries.get("word/document.xml")!));
    const stripped = xml.replace(/(<w:ins [^>]*w:author="Reviewer A"[^>]*><w:r><w:rPr>(?:(?!<\/w:rPr>).)*?)<w:vanish w:val="0"(?:\/>|><\/w:vanish>)/g, "$1");
    expect(stripped).not.toBe(xml);
    entries.set("word/document.xml", encodePart(stripped));
    const tampered = verifyOutput(norm, writePackage(entries), receipts);
    expect(tampered.ok).toBe(false);
    expect(tampered.failures.join(" ")).toMatch(/inserts hidden text/);
  });
});

describe("a span ending where a formatted run begins splits nothing", () => {
  it("leaves no empty run behind", () => {
    const body = P("10000001", R("Fee is due ") + `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">now.</w:t></w:r>`);
    const { out, receipts } = runApply(body, [{ op: "replace", ref: "r", author: AUTHOR, addr: "10000001", start: 7, end: 11, expectText: "due ", text: "owed " }]);
    expect(receipts[0]!.status).toBe("applied");
    const xml = new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!));
    expect(xml.match(/<w:b\/>/g)).toHaveLength(1);
    expect(xml).not.toMatch(/<w:r><w:rPr><w:b\/><\/w:rPr><\/w:r>/);
  });
});

describe("text that reads empty occupies no offset (fuzz find)", () => {
  it("lone unpreserved spaces between runs neither list nor offset", () => {
    const body =
      P("10000001", `<w:r><w:t xml:space="preserve">Action/to</w:t></w:r><w:r><w:t> </w:t></w:r><w:r><w:t xml:space="preserve">do</w:t></w:r><w:r><w:t> </w:t></w:r><w:r><w:t xml:space="preserve">list</w:t></w:r>`);
    const doc = openDocument(docx(body));
    expect(doc.projection().map((p) => p.text)).toEqual(["Action/todolist"]);
    const { receipts, out, norm } = runApply(body, [
      { op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 0, end: 15, expectText: "Action/todolist", text: "net-30" },
    ]);
    expect(receipts[0]!.status).toBe("applied");
    expect(openDocument(out).projectionRecords((r) => (r.author === AUTHOR ? "accept" : "accept")).map((p) => p.text)).toEqual(["net-30"]);
    expect(openDocument(out).projectionRecords((r) => (r.author === AUTHOR ? "reject" : "accept")).map((p) => p.text)).toEqual(["Action/todolist"]);
    expect(verifyOutput(norm, out, receipts).ok).toBe(true);
  });
});
