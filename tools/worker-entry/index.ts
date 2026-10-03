// C1 bundle check: the core imports cleanly into a Cloudflare Worker with no
// Node built-ins. Never deployed — `wrangler deploy --dry-run` only.
import { normalizeDocx, openDocument } from "@stitchdocx/core";

export default {
  async fetch(req: Request): Promise<Response> {
    const bytes = new Uint8Array(await req.arrayBuffer());
    const { docx, assigned } = normalizeDocx(bytes);
    const doc = openDocument(docx);
    return Response.json({ assigned, paragraphs: doc.projection().length });
  },
};
