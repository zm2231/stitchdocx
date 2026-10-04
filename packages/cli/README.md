# @stitchdocx/cli

Adapter CLI over the stitchdocx core. Five commands; each prints one JSON
object on stdout and exits 0. A non-zero exit means the engine itself failed —
per-op refusals are reported in the output, not as failures.

The core is a library: this CLI is the only place that touches the filesystem
and the only place the wire field `flagId` exists (the core's neutral name for
the same correlation field is `ref`; every op carries it and every receipt
echoes it).

## Commands

### `normalize <in.docx> <out.docx>`

Assigns a valid unique `w14:paraId` to every paragraph that lacks one (and
repairs duplicates), adds the `w14` namespace declaration when missing, and
writes the result. Untouched parts stay byte-identical. `protection` is the
enforced editing restriction from `word/settings.xml` (`readOnly`,
`comments`, `trackedChanges` or `forms`), or null — surfaced so the caller
can ask the document owner before processing; it is never changed.

```json
{ "assigned": 3, "protection": null }
```

### `paragraphs <in.docx>`

Read-only projection: every paragraph in reader order with its address,
text, kind, style, numbering label and table position. The input is
normalized in memory first, so every paragraph carries an address without
rewriting the stored file.

```json
{ "paragraphs": [ { "addr": "00000001", "text": "…", "kind": "body", "editable": true } ] }
```

`editable` is `false` (with `readOnlyReason: "merged_removed_mark"`) for
display-only records whose text was carried across a removed paragraph mark —
for example a deleted heading merged into the first cell of the table that
followed it. Their text is what Word shows on reject, but their offsets span
more than one physical paragraph, so skip them when deriving operations; the
resolver and `apply` refuse such spans with `paragraph_not_editable`. The
underlying paragraphs keep their own addresses and are editable in their own
coordinates.

### `apply <in.docx> <ops.json> <out.docx>`

`ops.json` is `{ "date": "<ISO timestamp stamped on every revision>", "ops": [...] }`.
Ops (all require `flagId`, `author`, `addr`):

- `{ "op": "replace", "addr", "start", "end", "expectText", "text", "granularity"? }`
  — track-deletes `[start,end)` and track-inserts `text` after it. With
  `"granularity": "word"` only the changed tokens are struck.
- `{ "op": "delete", "addr", "start", "end", "expectText" }`
- `{ "op": "insertParagraphAfter", "addr", "text", "scope"? }` — a new
  tracked-inserted paragraph after `addr` (or after the containing table with
  `"scope": "table"`).

Spans are in the paragraph's original text coordinates and must match
`expectText` exactly — address only records listed with `editable: true`
(see `paragraphs`). Output is `{ "receipts": [...] }`; each receipt reports
`index`, `flagId`, `author`, `status` (`applied`/`refused`), `addr`,
`changeIds` (the `w:id`s the op wrote), and for applied ops `before`/`after`
paragraph text and `insertedAddr`.

Refusals carry a machine-readable `reason`: `addr_not_found`,
`span_out_of_range`, `span_text_mismatch`, `empty_text`,
`span_overlaps_revision`, `span_overlaps_prior_op`, `span_crosses_structure`,
`span_in_field`, `span_in_content_control`, `paragraph_not_editable`,
`insert_in_plain_text_control` (a plain-text content control may hold only
one paragraph).

### `finalize <in.docx> <decisions.json> <out.docx>`

`decisions.json` is `{ "date"?, "decisions": [{ "flagId", "author",
"decision": "accept" | "reject", "text"? }] }`. Each decision resolves every
revision authored by `author` — exactly that author, nothing else; revisions
by anyone else survive untouched. `"text"` accepts-with-wording: the flagged
span reads as the original with the given text in place of our revision.
Output is `{ "results": [{ "flagId", "author", "decision", "resolved" }] }`.

### `verify <original.docx> <output.docx> <receipts.json>`

Pure receipt check: `receipts.json` is `{ "receipts": [...] }` as produced by
`apply`. Fails ( `{ "ok": false, "failures": [...] }` ) unless the output is
the original plus exactly what the receipts claim — every changeId present
and authored by the receipt's author, rejecting one receipt's revisions yields
its `before`, accepting them yields its `after`, untouched paragraphs byte-
identical in text, and only `word/document.xml` differs.
