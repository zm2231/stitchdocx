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

/** engine-level decision: by changeIds or by author (research/04 §3) */
export type RevisionDecision =
  | { changeIds: string[]; decision: "accept" | "reject" }
  | { author: string; decision: "accept" | "reject" };

/** finalize decision for one author: plain accept/reject, or accept with the
 *  caller's replacement wording */
export interface Decision {
  /** caller's id, echoed on the finalize result */
  ref: string;
  author: string;
  decision: "accept" | "reject";
  /** accept-with-caller-wording (adapter-level finalize composition) */
  text?: string;
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
  | "paragraph_not_editable";

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

/** result of Document.decide — a decision is document-wide, so it has no op identity */
export interface DecisionResult {
  index: number;
  decision: "accept" | "reject";
  /** w:id of every revision this decision resolved */
  changeIds: string[];
}
