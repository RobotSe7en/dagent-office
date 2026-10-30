import type { Annotation, TextRange } from "../contracts/annotations.js";
import { validTextRanges } from "../document/annotations.js";
import type { ParsedHtml } from "../document/html.js";
import { isTextBoundary } from "../document/text.js";
import type { DeepReadonly, TextOffsetMapping } from "../runtime.js";

export function annotationNodes(annotation: DeepReadonly<Annotation>): readonly string[] {
  const anchor = annotation.anchor;
  return anchor.kind === "text" || anchor.kind === "node" ? [anchor.nodeId]
    : anchor.kind === "text-set" ? anchor.ranges.map((range) => range.nodeId)
      : anchor.kind === "region" ? [anchor.surfaceId, ...anchor.nodeIds] : anchor.nodeIds;
}

function mapRange(range: TextRange, edits: readonly TextOffsetMapping[], after: ParsedHtml): TextRange | undefined {
  let { from, to } = range;
  for (const edit of edits) {
    if (edit.nodeId !== range.nodeId) continue;
    const delta = edit.insertedLength - (edit.to - edit.from);
    if (from === to) {
      // Point anchors attach to the following source; deleted interior points are uncertain.
      if (edit.from < from && from < edit.to) return undefined;
      if (edit.to <= from) from = to = from + delta;
      else if (edit.from === from) from = to = edit.from + edit.insertedLength;
    } else if (edit.to <= from) {
      from += delta;
      to += delta;
    } else if (edit.from >= to) {
      continue;
    } else if (edit.from <= from && edit.to >= to) {
      return undefined;
    } else {
      // Partial replacements remain selected; insertions at either edge stay outside.
      if (edit.from < from) from = edit.from;
      to = edit.to < to ? to + delta : edit.from + edit.insertedLength;
    }
  }
  const text = after.getNode(range.nodeId)?.text;
  if (text === undefined || from > to || to > text.length || !isTextBoundary(text, from) || !isTextBoundary(text, to)) return undefined;
  return { nodeId: range.nodeId, from, to, quote: text.slice(from, to) };
}

/** Map a complete annotation or retain its complete original anchor for manual recovery. */
export function mapAnnotations(
  annotations: Annotation[], before: ParsedHtml, after: ParsedHtml,
  edits: { readonly textEdits: readonly TextOffsetMapping[]; readonly replacedTextNodes: ReadonlySet<string>; readonly structuralTextNodes: ReadonlySet<string> },
  explicit: ReadonlySet<string>,
): void {
  const layoutChanged = before.html !== after.html;
  const byNode = new Map<string, TextOffsetMapping[]>();
  for (const edit of edits.textEdits) {
    const items = byNode.get(edit.nodeId) ?? [];
    items.push(edit);
    byNode.set(edit.nodeId, items);
  }
  for (const annotation of annotations) {
    if (annotation.status === "needs_reanchor" || explicit.has(annotation.annotationId)) continue;
    const anchor = annotation.anchor;
    let invalid = annotationNodes(annotation).some((id) => !after.hasNode(id)) || anchor.kind === "region" && layoutChanged;
    if (!invalid && (anchor.kind === "text" || anchor.kind === "text-set") && layoutChanged) {
      const ranges = anchor.kind === "text" ? [anchor] : anchor.ranges;
      // Activating recovery data without a new anchor must not manufacture a matching quote.
      if (!validTextRanges(ranges, before)) continue;
      const mapped = ranges.map((range) => edits.replacedTextNodes.has(range.nodeId) ? undefined
        : mapRange(range, !edits.structuralTextNodes.has(range.nodeId) && before.getNode(range.nodeId)?.text === after.getNode(range.nodeId)?.text ? [] : byNode.get(range.nodeId) ?? [], after));
      invalid = mapped.some((range) => range === undefined) || !validTextRanges(mapped as TextRange[], after);
      if (!invalid) annotation.anchor = anchor.kind === "text"
        ? { kind: "text", ...mapped[0]! } : { kind: "text-set", ranges: mapped as TextRange[] };
    }
    if (invalid) annotation.status = "needs_reanchor";
  }
}
