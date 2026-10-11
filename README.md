# stitchdocx

A pure-TypeScript, Worker-compatible DOCX (OOXML) engine for reading,
tracked-change editing and verifying Word documents. The XML tree is the
document: untouched subtrees and package parts re-serialize byte-identically,
so an edit changes exactly what it claims to change.

- `packages/package-io` — byte-faithful XML tree (every fixture part round-
  trips byte-identically), zip package IO with deterministic output, light
  part registry.
- `packages/core` — paraId addressing, Word-faithful projection, hardened
  span resolution, tracked-change ops (`replace` span/word, `delete`,
  `insertParagraphAfter`) with receipts and a closed refusal taxonomy, pure
  accept/reject resolution, receipt verification, deterministic save.
- `packages/cli` — the five-op adapter CLI (`normalize`, `paragraphs`,
  `apply`, `finalize`, `verify`); see [packages/cli/README.md](packages/cli/README.md).

## Development

```bash
pnpm install
pnpm build   # bundles packages/cli to dist/cli.js
pnpm test    # unit + property suites
```

The acceptance-corpus suites (parser byte-fidelity, package round-trip over
the fixture corpus, projection parity against the independent verification
reader) run only when `DOCX_ACCEPTANCE_DIR` points at an acceptance folder;
without it they are skipped and the synthetic suites still cover the
invariants.

The acceptance gate itself:

```bash
DOCX_ACCEPTANCE_DIR=… pnpm verify:acceptance
```

## License

[PolyForm Noncommercial 1.0.0](LICENSE.md). Free for noncommercial use. For commercial use, contact the author.
