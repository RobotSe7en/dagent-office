import type { WorkbookCommentMention } from "../contracts/workbook-comments.js";
import { isTextBoundary } from "../document/text.js";

/** Resolve preserved native offsets against the current exact UTF-16 body.
 * Invalid offsets remain provenance on needs_reanchor threads; never clamp
 * them or infer a person, notification target or permission from these IDs. */
export function resolveWorkbookCommentMention(text: string, mention: Pick<WorkbookCommentMention, "startIndex" | "length">): Readonly<{ from: number; to: number }> | undefined {
  if (typeof text !== "string" || !mention || ![mention.startIndex, mention.length].every(value => Number.isSafeInteger(value) && value >= 0 && value <= 4294967295)) return undefined;
  const from = mention.startIndex, to = from + mention.length;
  return to <= text.length && isTextBoundary(text, from) && isTextBoundary(text, to) ? Object.freeze({ from, to }) : undefined;
}
