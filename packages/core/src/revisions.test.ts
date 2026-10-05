// Listing revisions: every element decide() can name, with the author that wrote it.
import { describe, it, expect } from "vitest";
import { encodePart, writePackage } from "@stitchdocx/package-io";
import { normalizeDocx, openDocument } from "./index.js";

const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"`;
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const R = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const INS = (id: number, author: string, t: string) => `<w:ins w:id="${id}" w:author="${author}" w:date="2026-10-01T00:00:00Z">${R(t)}</w:ins>`;
const DEL = (id: number, author: string, t: string) => `<w:del w:id="${id}" w:author="${author}" w:date="2026-10-01T00:00:00Z"><w:r><w:delText xml:space="preserve">${t}</w:delText></w:r></w:del>`;
const P = (inner: string, mark = "") => `<w:p>${mark ? `<w:pPr><w:rPr>${mark}</w:rPr></w:pPr>` : ""}${inner}</w:p>`;

function docx(body: string): Uint8Array {
  return normalizeDocx(writePackage(new Map<string, Uint8Array>([
    ["[Content_Types].xml", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
    ["_rels/.rels", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ["word/document.xml", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`)],
  ]))).docx;
}

describe("Document.revisions", () => {
  const body =
    P(R("Fee ") + DEL(1, "Scope Optimizer [flag_a1]", "TBD") + INS(2, "Scope Optimizer [flag_a1]", "$5,000")) +
    P(R("Term ") + INS(3, "Client &amp; Counsel", "twelve months")) +
    P(R("New clause"), `<w:ins w:id="4" w:author="Scope Optimizer [flag_b2]" w:date="2026-10-01T00:00:00Z"/>`) +
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${P(INS(5, "Scope Optimizer [flag_b2]", "cell text"))}</w:tc></w:tr></w:tbl>`;

  it("lists each revision with its id, decoded author and element name", () => {
    const listed = openDocument(docx(body)).revisions().map((r) => `${r.id}|${r.author}|${r.name}`).sort();
    expect(listed).toEqual([
      "1|Scope Optimizer [flag_a1]|w:del",
      "2|Scope Optimizer [flag_a1]|w:ins",
      "3|Client & Counsel|w:ins",
      "4|Scope Optimizer [flag_b2]|w:ins",
      "5|Scope Optimizer [flag_b2]|w:ins",
    ]);
  });

  it("names ids that decide() resolves, leaving other authors' changes", () => {
    const doc = openDocument(docx(body));
    const ids = [...new Set(doc.revisions().filter((r) => r.author?.includes("[flag_a1]")).map((r) => r.id))];
    doc.decide([{ ref: "flag_a1", changeIds: ids, decision: "accept" }]);
    const after = openDocument(doc.save());
    expect(after.projection()[0]?.text).toBe("Fee $5,000");
    expect(after.revisions().map((r) => r.id).sort()).toEqual(["3", "4", "5"]);
  });

  it("is empty for a document with no tracked changes", () => {
    expect(openDocument(docx(P(R("Plain")))).revisions()).toEqual([]);
  });
});
