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
import { normalizeDocx, openDocument, verifyOutput, type Op, type Receipt, type Decision, type RevisionDecision } from "@stitchdocx/core";

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
      const doc = openDocument(new Uint8Array(readFileSync(inp)));
      const receipts: Receipt[] = doc.apply((spec.ops ?? []).map(toCoreOp), { date: spec.date });
      writeFileSync(outp, doc.save());
      out({ receipts: receipts.map(toWireReceipt) });
      break;
    }
    case "finalize": {
      const [inp, decPath, outp] = args;
      if (!inp || !decPath || !outp) fail("usage: finalize <in.docx> <decisions.json> <out.docx>");
      const spec = JSON.parse(readFileSync(decPath, "utf8")) as { date?: string; decisions: Array<Omit<Decision, "ref"> & { ref?: string; flagId?: string }> };
      const doc = openDocument(new Uint8Array(readFileSync(inp)));
      const results: Array<{ flagId: string; author: string; decision: string; resolved: number }> = [];
      // resolve exactly the authors named in the decisions — revisions by
      // anyone else (including earlier runs' work in a re-uploaded redline)
      // must survive untouched
      for (const [index, d] of (spec.decisions ?? []).entries()) {
        const ref = d.ref ?? d.flagId;
        if (!ref) fail(`decision ${index} is missing its flagId`);
        const core: Decision = { ...d, ref };
        if (core.text !== undefined) {
          const resolved = doc.acceptWithText(core);
          results.push({ flagId: d.flagId ?? ref, author: d.author, decision: "accept", resolved });
        } else {
          const engineDecision: RevisionDecision = { author: d.author, decision: d.decision };
          const decided = doc.decide([engineDecision]);
          const resolved = decided.reduce((n, r) => n + r.changeIds.length, 0);
          results.push({ flagId: d.flagId ?? ref, author: d.author, decision: d.decision, resolved });
        }
      }
      writeFileSync(outp, doc.save());
      out({ results });
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
