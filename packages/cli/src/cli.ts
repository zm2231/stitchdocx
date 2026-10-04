#!/usr/bin/env node
// Redline adapter CLI: five commands over a .docx file, printing one JSON
// object per command on stdout. This is the only place Node built-ins appear;
// the core stays runtime-neutral. A non-zero exit means the engine itself
// failed (refusals are not failures). The wire contract is documented in
// packages/cli/README.md.
//
// The wire format names each operation's correlation field `flagId`; the core
// calls it `ref` (a neutral caller id echoed on receipts). This adapter maps
// between the two and is the only place the wire name exists.
import { readFileSync, writeFileSync } from "node:fs";
import { finalizeDocx, isIsoDate, normalizeDocx, openDocument, verifyOutput, type Op, type Receipt, type Decision } from "@stitchdocx/core";

const [cmd, ...args] = process.argv.slice(2);

function fail(msg: string): never {
  console.error(`stitchdocx: ${msg}`);
  process.exit(1);
}

const out = (json: unknown): void => {
  process.stdout.write(JSON.stringify(json));
};

type WireOp = Omit<Op, "ref"> & { ref?: string; flagId?: string };
const toCoreOp = (o: WireOp, index: number): Op => {
  const { flagId, ...rest } = o;
  const ref = rest.ref ?? flagId;
  if (!ref) fail(`op ${index} is missing its flagId`);
  if (typeof rest.author !== "string" || rest.author.trim() === "") fail(`op ${index} (flagId ${ref}) needs a non-empty author`);
  return { ...rest, ref } as Op;
};

type WireReceipt = Omit<Receipt, "ref"> & { ref?: string; flagId?: string };
const toCoreReceipt = (r: WireReceipt, index: number): Receipt => {
  const { flagId, ...rest } = r;
  const ref = rest.ref ?? flagId;
  if (!ref) fail(`receipt ${index} is missing its flagId`);
  return { ...rest, ref } as Receipt;
};
const toWireReceipt = (r: Receipt): WireReceipt => {
  const { ref, ...rest } = r;
  return { ...rest, flagId: ref };
};

type WireDecision = Omit<Decision, "ref"> & { flagId?: string };
const toCoreDecision = (d: WireDecision, index: number): Decision => {
  const { flagId, ...rest } = d;
  if (!flagId) fail(`decision ${index} is missing its flagId`);
  if (!Array.isArray(rest.changeIds) || rest.changeIds.length === 0) fail(`decision ${index} (flagId ${flagId}) needs the receipt's changeIds`);
  if (rest.decision !== "accept" && rest.decision !== "reject") fail(`decision ${index} (flagId ${flagId}) must be accept or reject`);
  return { ...rest, changeIds: rest.changeIds.map(String), ref: flagId };
};

const requireDate = (value: unknown, what: string): string => {
  if (!isIsoDate(value)) fail(`${what} must be an ISO 8601 timestamp (e.g. 2026-01-01T00:00:00Z), got ${JSON.stringify(value)}`);
  return value;
};

try {
  switch (cmd) {
    case "normalize": {
      const [inp, outp] = args;
      if (!inp || !outp) fail("usage: normalize <in.docx> <out.docx>");
      const { docx, assigned, protection } = normalizeDocx(new Uint8Array(readFileSync(inp)));
      writeFileSync(outp, docx);
      out({ assigned, protection });
      break;
    }
    case "paragraphs": {
      const [inp] = args;
      if (!inp) fail("usage: paragraphs <in.docx>");
      // normalize in memory first: every projected paragraph carries a valid
      // address (the stored file is not rewritten — paragraphs is read-only)
      const { docx } = normalizeDocx(new Uint8Array(readFileSync(inp)));
      const doc = openDocument(docx);
      out({ paragraphs: doc.projection() });
      break;
    }
    case "apply": {
      const [inp, opsPath, outp] = args;
      if (!inp || !opsPath || !outp) fail("usage: apply <in.docx> <ops.json> <out.docx>");
      const spec = JSON.parse(readFileSync(opsPath, "utf8")) as { date: string; ops: WireOp[] };
      const date = requireDate(spec.date, "date");
      const doc = openDocument(new Uint8Array(readFileSync(inp)));
      const receipts: Receipt[] = doc.apply((spec.ops ?? []).map(toCoreOp), { date });
      writeFileSync(outp, doc.save());
      out({ receipts: receipts.map(toWireReceipt) });
      break;
    }
    case "finalize": {
      const [inp, decPath, outp] = args;
      if (!inp || !decPath || !outp) fail("usage: finalize <in.docx> <decisions.json> <out.docx>");
      const spec = JSON.parse(readFileSync(decPath, "utf8")) as { date?: string; decisions: WireDecision[] };
      if (spec.date !== undefined) requireDate(spec.date, "date");
      const { docx, results } = finalizeDocx(new Uint8Array(readFileSync(inp)), (spec.decisions ?? []).map(toCoreDecision));
      writeFileSync(outp, docx);
      out({ results: results.map(({ ref, ...rest }) => ({ ...rest, flagId: ref })) });
      break;
    }
    case "verify": {
      const [origPath, outPath, receiptsPath] = args;
      if (!origPath || !outPath || !receiptsPath) fail("usage: verify <original.docx> <output.docx> <receipts.json>");
      const spec = JSON.parse(readFileSync(receiptsPath, "utf8")) as { receipts: WireReceipt[] };
      const result = verifyOutput(new Uint8Array(readFileSync(origPath)), new Uint8Array(readFileSync(outPath)), (spec.receipts ?? []).map((r, i) => toCoreReceipt(r, i)));
      out(result);
      break;
    }
    default:
      fail(`unknown command ${String(cmd)}; expected normalize|paragraphs|apply|finalize|verify`);
  }
} catch (err) {
  fail(err instanceof Error ? err.message : String(err));
}
