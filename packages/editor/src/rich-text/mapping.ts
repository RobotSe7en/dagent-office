import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import type { Transaction } from "@tiptap/pm/state";
import type { ReconciledTextEdit } from "@dagent-office/core";
import { attributes, canonicalOffset, canonicalText, sourceId } from "./source.js";

function index(document: ProseMirrorNode): Map<string, { node: ProseMirrorNode; position: number }> {
  const result = new Map<string, { node: ProseMirrorNode; position: number }>();
  document.descendants((node, position) => { const id = sourceId(node); if (id && !result.has(id)) result.set(id, { node, position }); });
  return result;
}
export function transactionProofs(transaction: Transaction): ReconciledTextEdit[] {
  const result: ReconciledTextEdit[] = [];
  for (const [stepIndex, step] of transaction.steps.entries()) {
    const before = transaction.docs[stepIndex]!, after = transaction.docs[stepIndex + 1] ?? transaction.doc;
    const previous = index(before), following = index(after);
    let from = Number.MAX_SAFE_INTEGER, to = 0;
    step.getMap().forEach((start, end) => { from = Math.min(from, start); to = Math.max(to, end); });
    for (const [nodeId, entry] of previous) {
      const next = following.get(nodeId); if (!next) continue;
      const oldText = canonicalText(entry.node), newText = canonicalText(next.node);
      if (oldText === newText) continue;
      let start = canonicalOffset(entry.node, Math.max(0, Math.min(entry.node.content.size, from - entry.position - 1)));
      let end = canonicalOffset(entry.node, Math.max(0, Math.min(entry.node.content.size, to - entry.position - 1)));
      let length = newText.length - oldText.length + end - start;
      if (length < 0) { end = Math.min(oldText.length, end - length); length = newText.length - oldText.length + end - start; }
      if (length < 0 || oldText.slice(0, start) !== newText.slice(0, start)
        || oldText.slice(end) !== newText.slice(start + length)) { start = 0; end = oldText.length; length = newText.length; }
      result.push({ nodeId, from: start, to: end, expectedText: oldText.slice(start, end), text: newText.slice(start, start + length) });
    }
  }
  return result;
}

/** Find the derived editor position corresponding to a canonical UTF-16 offset. */
function locate(node: ProseMirrorNode, position: number, requested: number): number {
    if (node.isText) return position + Math.max(0, Math.min(node.text!.length, requested));
    if (node.isLeaf) return position + (requested > 0 ? node.nodeSize : 0);
    let remaining = requested, found: number | undefined;
    node.forEach((child, childOffset) => {
      if (found !== undefined) return;
      const prefix = attributes(child).beforeText?.length ?? 0, length = canonicalText(child).length;
      remaining = Math.max(0, remaining - prefix);
      if (remaining <= length) found = locate(child, position + (node.type.name === "doc" ? 0 : 1) + childOffset, remaining);
      else remaining -= length;
    });
    return found ?? position + (node.type.name === "doc" ? node.content.size : node.nodeSize - 1);
}
export function positionAtOffset(document: ProseMirrorNode, offset: number): number {
  return Math.max(0, Math.min(document.content.size, locate(document, 0, offset)));
}

export interface TextBookmark {
  readonly nodeId?: string;
  readonly blockPath: readonly number[];
  readonly offset: number;
  readonly fallback: number;
}

/** Paragraph boundaries have the same plain-text offset; retain a structural location too. */
export function textBookmark(document: ProseMirrorNode, position: number): TextBookmark {
  const resolved = document.resolve(position), fallback = canonicalOffset(document, position);
  let depth = resolved.depth;
  while (depth > 0 && !resolved.node(depth).isTextblock) depth--;
  if (!depth) return { blockPath: [], offset: fallback, fallback };
  const node = resolved.node(depth), id = sourceId(node), path: number[] = [];
  let identities = 0;
  if (id) document.descendants((child) => { if (sourceId(child) === id) identities++; });
  for (let level = 0; level < depth; level++) {
    let ordinal = 0;
    resolved.node(level).forEach((child, _offset, index) => { if (index < resolved.index(level) && child.isBlock) ordinal++; });
    path.push(ordinal);
  }
  return { ...(id && identities === 1 ? { nodeId: id } : {}), blockPath: path,
    offset: canonicalOffset(node, position - resolved.start(depth)), fallback };
}

export function bookmarkPosition(document: ProseMirrorNode, bookmark: TextBookmark): number {
  let target: { node: ProseMirrorNode; position: number } | undefined;
  if (bookmark.nodeId) document.descendants((node, position) => {
    if (sourceId(node) === bookmark.nodeId && node.isTextblock) target = { node, position };
  });
  if (!target) {
    let current = document, position = -1, valid = true;
    for (const ordinal of bookmark.blockPath) {
      let chosen: { node: ProseMirrorNode; position: number } | undefined, index = 0;
      current.forEach((child, offset) => {
        if (child.isBlock && index++ === ordinal) chosen = { node: child, position: position + 1 + offset };
      });
      if (!chosen) { valid = false; break; }
      current = chosen.node; position = chosen.position;
    }
    if (valid && current.isTextblock) target = { node: current, position };
  }
  return target ? Math.max(0, Math.min(document.content.size, locate(target.node, target.position, bookmark.offset)))
    : positionAtOffset(document, bookmark.fallback);
}
