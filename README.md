# stitchdocx

stitchdocx puts tracked changes into someone else's Word file and hands it back otherwise exactly as it came in. Each edit lands as a Word revision under the author you name. Paragraphs you didn't touch keep their exact text, and every other part of the package is byte-for-byte the original. It's plain TypeScript with no native dependencies, so it runs in Node and inside a Cloudflare Worker.

It exists for one job: a program (usually a model) decides what a contract should say, and the client still has to open their own file in Word, see each suggestion as a redline, and accept or reject it. The hard part is making those edits without disturbing the numbering/styles/existing revisions that the client's file already carries. The engine either makes an edit exactly as asked or refuses it and tells you why.

## Try it in five minutes

```bash
pnpm install
pnpm build                     # builds packages/cli/dist/cli.js
alias stitchdocx="node $PWD/packages/cli/dist/cli.js"

stitchdocx normalize contract.docx work.docx     # give every paragraph an address
stitchdocx paragraphs work.docx                  # list addresses and text
```

`paragraphs` prints what Word shows, one entry per paragraph:

```json
{ "addr": "10000008", "kind": "body", "text": "The engagement is billed at $3,650 Fixed Fee, payable in advance.", "editable": true }
```

Write the edits against those addresses. Offsets count characters in that paragraph's text, and `expectText` must match the span exactly, so an edit aimed at the wrong place is refused instead of applied somewhere close.

```json
{
  "date": "2026-10-11T12:00:00Z",
  "ops": [
    { "op": "replace", "flagId": "fee", "author": "Reviewer", "addr": "10000008",
      "start": 28, "end": 34, "expectText": "$3,650", "text": "$4,200" },
    { "op": "insertParagraphAfter", "flagId": "late", "author": "Reviewer", "addr": "10000008",
      "text": "Invoices unpaid after 30 days accrue 1% interest per month." }
  ]
}
```

```bash
stitchdocx apply work.docx ops.json out.docx > receipts.json
stitchdocx verify work.docx out.docx receipts.json    # {"ok":true,"failures":[]}
```

Open `out.docx` in Word: `$3,650` is struck, `$4,200` is inserted, and the late-fee clause is a new tracked paragraph, all by Reviewer. Each receipt names the revision ids it wrote:

```json
{ "flagId": "fee", "status": "applied", "changeIds": ["1", "2"],
  "before": "The engagement is billed at $3,650 Fixed Fee, payable in advance.",
  "after":  "The engagement is billed at $4,200 Fixed Fee, payable in advance." }
```

Those ids are how you resolve the edit later. `finalize` accepts or rejects exactly the revisions you name and leaves every other one, including the client's own, where it was:

```json
{ "decisions": [
  { "flagId": "fee",  "changeIds": ["1", "2"], "decision": "accept" },
  { "flagId": "late", "changeIds": ["3", "4"], "decision": "reject" }
] }
```

```bash
stitchdocx finalize out.docx decisions.json final.docx
```

## The five commands

| Command | What it does |
|---|---|
| `normalize` | Gives every paragraph a unique `w14:paraId` address (adding the namespace declaration if it's missing) and reports any editing restriction set in the file. |
| `paragraphs` | Lists every paragraph Word shows, in reading order, with its address, text, kind, style, list label and table position. |
| `apply` | Writes `replace`, `delete` and `insertParagraphAfter` edits as tracked changes and returns one receipt per edit. |
| `finalize` | Accepts or rejects the revisions a receipt names. A decision can also accept with different wording. |
| `verify` | Fails unless the output is the original plus exactly what the receipts claim. |

Every command prints one JSON object and exits 0. A refused edit is a normal result with a reason, and a non-zero exit means the engine itself broke. The full field reference is in [packages/cli/README.md](packages/cli/README.md).

## What gets refused, and why

A refusal means the engine could not make that edit without changing something you didn't ask it to. The other edits in the batch still apply.

| Situation (the common ones) | Result |
|---|---|
| `expectText` doesn't match the span, or the address doesn't exist | `span_text_mismatch`, `addr_not_found` |
| The span overlaps a revision already in the file, or an earlier edit in the same batch | `span_overlaps_revision`, `span_overlaps_prior_op` |
| The span is a field result (a page number, a table of contents entry) | `span_in_field` |
| The span reaches into a content control, office math or other fixed structure | `span_in_content_control`, `span_crosses_structure` |
| The span crosses the edge of a hyperlink | Applied. Each side is edited in place and the link survives. |

Text Word hides never shows up in `paragraphs`, so it can't be targeted. Inserted words take their formatting from the words they replace, or from the text beside them, as typing in Word would.

## Limits

A document too big to hold safely is refused whole with `{ "refused": { "reason", "detail" } }`, so the caller can route it elsewhere. The main document part can be up to 8 MiB, and node budgets cap how much XML any one parse builds. The limits are sized for a 128 MB Worker: the heaviest document built to sit just inside them (a 7.9 MB main part with 19,000 table cells) runs normalize, apply and verify in one Worker request.

## Using it as a library

The CLI is a thin wrapper. `@stitchdocx/core` takes and returns bytes and never touches the filesystem, which is what lets it run in a Worker. The packages aren't on npm yet, so use them from this workspace or bundle `packages/core` into your project.

```ts
import { normalizeDocx, openDocument } from "@stitchdocx/core";

const doc = openDocument(normalizeDocx(bytes).docx);
const fees = doc.projection().find((p) => p.text.includes("$3,650"))!;
const start = fees.text.indexOf("$3,650");
const [receipt] = doc.apply(
  [{ op: "replace", ref: "fee", author: "Reviewer", addr: fees.addr, start, end: start + 6, expectText: "$3,650", text: "$4,200" }],
  { date: new Date().toISOString() },
);
const out = doc.save();
```

`ref` here is the same field the CLI calls `flagId`.

| Package | Contents |
|---|---|
| `packages/package-io` | Reads and writes the zip package and an XML tree that re-serializes byte-identically. |
| `packages/core` | Addressing, the paragraph view Word shows, edits with receipts, accept/reject, verification. |
| `packages/cli` | The five commands above. |

## Development

```bash
pnpm test
```

The acceptance suites (byte-fidelity over a fixture corpus, and parity with an independent reader) run when `DOCX_ACCEPTANCE_DIR` points at an acceptance folder and are skipped otherwise:

```bash
DOCX_ACCEPTANCE_DIR=/path/to/acceptance pnpm verify:acceptance
```

## License

[PolyForm Noncommercial 1.0.0](LICENSE.md). Free for noncommercial use. For commercial use, contact the author.
