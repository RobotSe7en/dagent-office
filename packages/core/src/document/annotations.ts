import type { Annotation, TextRange } from "../contracts/annotations.js";
import type { DocumentManifest } from "../contracts/document.js";
import type { ParsedHtml } from "./html.js";
import { isTextBoundary } from "./text.js";
import { invalid } from "./validation.js";

/** A text-set denotes one ordered selection, including ranges on nested source nodes. */
export function validTextRanges(ranges: readonly TextRange[], html: ParsedHtml): boolean {
  let order: Map<string, number> | undefined;
  const ancestorOf = (ancestor: string, id: string): boolean => {
    for (let current: string | null = id; current !== null; current = html.peekNode(current)!.parentNodeId) {
      if (current === ancestor) return true;
    }
    return false;
  };
  let previous: { nodeId: string; from: number; to: number; story: string } | undefined;
  for (const range of ranges) {
    const text = html.getNode(range.nodeId)?.text;
    if (text === undefined || range.from > range.to || range.to > text.length || text.slice(range.from, range.to) !== range.quote
      || !isTextBoundary(text, range.from) || !isTextBoundary(text, range.to)) return false;
    const offset = html.textOffset(range.nodeId)!;
    const story = html.textScope(range.nodeId, range.from, range.to);
    if (story === undefined) return false;
    const current = { nodeId: range.nodeId, from: offset + range.from, to: offset + range.to, story };
    if (previous && (current.story !== previous.story || current.from < previous.to)) return false;
    if (previous && current.from === previous.from) {
      order ??= new Map(html.nodeIds().map((id, index) => [id, index]));
      if (order.get(current.nodeId)! <= order.get(previous.nodeId)! || ancestorOf(previous.nodeId, current.nodeId)) return false;
    }
    previous = current;
  }
  return true;
}

/** Validate source anchors, independent of native Word comment metadata. */
export function validateAnnotations(annotations: readonly Annotation[], manifest: DocumentManifest, html: ParsedHtml): void {
  const identities = new Set<string>();
  const inside = (id: string, surface: string): boolean => {
    for (let current: string | null = id; current !== null; current = html.peekNode(current)!.parentNodeId) {
      if (current === surface) return true;
    }
    return false;
  };
  for (const annotation of annotations) {
    if (identities.has(annotation.annotationId) || annotation.documentId !== manifest.documentId || annotation.baseRevision > manifest.revision) {
      invalid("Annotations require unique identities and must belong to this document and its history.");
    }
    identities.add(annotation.annotationId);
    // Recovery data intentionally refers to an older source, including deleted nodes.
    if (annotation.status === "needs_reanchor") continue;
    const anchor = annotation.anchor;
    const ids = anchor.kind === "text" || anchor.kind === "node" ? [anchor.nodeId]
      : anchor.kind === "text-set" ? anchor.ranges.map((range) => range.nodeId) : anchor.nodeIds;
    if (ids.some((id) => !html.hasNode(id))) invalid("An active or resolved annotation refers to a missing document node.");
    if ((anchor.kind === "node-set" || anchor.kind === "region") && new Set(ids).size !== ids.length) {
      invalid("Annotation node sets must not repeat identities.");
    }
    if (anchor.kind === "region") {
      if (!html.isContainer(anchor.surfaceId) || ids.some((id) => !inside(id, anchor.surfaceId))) {
        invalid("A region requires an existing container surface and hits inside that surface.");
      }
    }
    if (anchor.kind === "text" || anchor.kind === "text-set") {
      const ranges = anchor.kind === "text" ? [anchor] : anchor.ranges;
      if (!validTextRanges(ranges, html)) invalid("Annotation text ranges must match UTF-16 source boundaries, be ordered, non-overlapping and in one content story.");
    }
  }
}
