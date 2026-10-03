# stitchdocx

**Status: Phase 1 goal complete — the engine core + redline adapter pass the
independent handoff harness 40/40.**

The engine is a pure-TypeScript, Worker-compatible DOCX (OOXML) library:

- `packages/package-io` — byte-faithful XML tree (400/400 fixture parts
  round-trip byte-identically), zip package IO with deterministic output,
  light part registry.
- `packages/core` — paraId addressing, reader-exact projection, hardened span
  resolution, tracked-change ops (`replace` span/word, `delete`,
  `insertParagraphAfter`) with receipts and a closed refusal taxonomy, pure
  accept/reject resolution, receipt verification, deterministic save.
- `packages/cli` — the five-op redline adapter CLI (`normalize`, `paragraphs`,
  `apply`, `finalize`, `verify`) that the acceptance harness drives.

## Acceptance gate (continuously runnable)

```bash
cd $DOCX_ACCEPTANCE_DIR
npm run selftest                                        # independent reader tests: 6/6
npm run verify -- --engine "node $REPO/packages/cli/dist/cli.js"   # 40/40
```

Rules: zero edits to `cases/` or `harness/`. Additional
gates: `pnpm test` (fidelity/round-trip/parity-oracle suites),
`wrangler deploy --dry-run` (Worker bundle, `tools/worker-entry`), permissive-
only deps (fflate MIT, jsdiff BSD-3).
