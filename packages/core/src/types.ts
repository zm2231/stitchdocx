// Public contract types (see research/04-engine-api-adjudication.md).

export type Addr = string; // w14:paraId, 8 hex digits, < 0x80000000

export interface Paragraph {
  addr: Addr;
  text: string;
  kind: "body" | "heading" | "listItem" | "cell";
  table?: { index: number; row: number; col: number };
  /** rendered numbering label, e.g. "2." — never part of text */
  numbering?: string;
  style?: string;
  editable: boolean;
  readOnlyReason?: string;
}

/** finalize decision: the revisions it resolves are the receipt's changeIds */
export interface Decision {
  /** caller's id, echoed on the result */
  ref: string;
  changeIds: string[];
  decision: "accept" | "reject";
  /** accept with the caller's wording in place of the flagged span or inserted paragraph */
  text?: string;
  /** the op's expectText; with `text`, it locates the whole flagged span of a replace or delete */
  expectText?: string;
}

export type Op =
  | {
      op: "replace";
      ref: string;
      author: string;
      addr: Addr;
      start: number;
      end: number;
      expectText: string;
      text: string;
      granularity?: "span" | "word";
    }
  | {
      op: "delete";
      ref: string;
      author: string;
      addr: Addr;
      start: number;
      end: number;
      expectText: string;
    }
  | {
      op: "insertParagraphAfter";
      ref: string;
      author: string;
      addr: Addr;
      text: string;
      scope?: "paragraph" | "table";
    };

export type RefusalReason =
  | "addr_not_found"
  | "span_out_of_range"
  | "span_text_mismatch"
  | "empty_text"
  | "span_overlaps_revision"
  | "span_overlaps_prior_op"
  | "span_crosses_structure"
  | "span_in_field"
  | "span_in_content_control"
  | "paragraph_not_editable"
  | "insert_in_plain_text_control";

export interface Receipt {
  index: number;
  /** the op's ref, echoed */
  ref: string;
  /** the op's author — binds every changeId to the revision identity that wrote it */
  author: string;
  status: "applied" | "refused";
  reason?: RefusalReason;
  addr: Addr;
  /** w:id of every revision this op wrote; [] when refused */
  changeIds: string[];
  /** paragraph text with this receipt's revisions rejected and everything else accepted */
  before?: string;
  /** paragraph text with everything accepted */
  after?: string;
  /** insertParagraphAfter only */
  insertedAddr?: Addr;
}

export interface DecisionResult {
  index: number;
  ref: string;
  decision: "accept" | "reject";
  /** w:id of every revision this decision resolved */
  changeIds: string[];
}
