#!/usr/bin/env node
// Redline adapter CLI — the five-op contract the handoff harness tests (see
// the handoff API.md). Thin over the stitchdocx core; the only place Node
// built-ins appear. Every command prints one JSON object and exits 0; a
// non-zero exit means the engine itself failed (refusals are not failures).
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

try {
  switch (cmd) {
    case "normalize": {
      const [inp, outp] = args;
      if (!inp || !outp) fail("usage: normalize <in.docx> <out.docx>");
      const { docx, assigned } = normalizeDocx(new Uint8Array(readFileSync(inp)));
      writeFileSync(outp, docx);
      out({ assigned });
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
      const spec = JSON.parse(readFileSync(opsPath, "utf8")) as { date: string; ops: Op[] };
      const doc = openDocument(new Uint8Array(readFileSync(inp)));
      const receipts: Receipt[] = doc.apply(spec.ops ?? [], { date: spec.date });
      writeFileSync(outp, doc.save());
      out({ receipts });
      break;
    }
    case "finalize": {
      const [inp, decPath, outp] = args;
      if (!inp || !decPath || !outp) fail("usage: finalize <in.docx> <decisions.json> <out.docx>");
      const spec = JSON.parse(readFileSync(decPath, "utf8")) as { date?: string; decisions: Decision[] };
      const doc = openDocument(new Uint8Array(readFileSync(inp)));
      const results: Array<{ flagId: string; author: string; decision: string; resolved: number }> = [];
      for (const d of spec.decisions ?? []) {
        if (d.text !== undefined) {
          const resolved = doc.acceptWithWording(d);
          results.push({ flagId: d.flagId, author: d.author, decision: "accept", resolved });
        } else {
          const engineDecision: RevisionDecision = { author: d.author, decision: d.decision };
          const decided = doc.decide([engineDecision]);
          const resolved = decided.reduce((n, r) => n + r.changeIds.length, 0);
          results.push({ flagId: d.flagId, author: d.author, decision: d.decision, resolved });
        }
      }
      // sweep: no revision by any of our authors may remain (undecided flags default to reject)
      const swept = doc.decideByPrefix("Scope Optimizer", "reject");
      writeFileSync(outp, doc.save());
      out({ results, swept });
      break;
    }
    case "verify": {
      const [origPath, outPath, receiptsPath] = args;
      if (!origPath || !outPath || !receiptsPath) fail("usage: verify <original.docx> <output.docx> <receipts.json>");
      const spec = JSON.parse(readFileSync(receiptsPath, "utf8")) as { receipts: Receipt[] };
      const result = verifyOutput(new Uint8Array(readFileSync(origPath)), new Uint8Array(readFileSync(outPath)), spec.receipts ?? []);
      out(result);
      break;
    }
    default:
      fail(`unknown command ${String(cmd)}; expected normalize|paragraphs|apply|finalize|verify`);
  }
} catch (err) {
  fail(err instanceof Error ? err.message : String(err));
}
