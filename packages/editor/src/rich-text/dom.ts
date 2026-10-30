import type { ReconciledTextEdit } from "@dagent-office/core";

export type IgnoreNode = (node: Node) => boolean;
export const blockSelector = "p,h1,h2,h3,h4,h5,h6,li,td,th";
export function sourceText(node: Node, ignored: IgnoreNode): string {
  if (ignored(node)) return "";
  if (node.nodeType === 3) return node.textContent ?? "";
  if (node.nodeType === 1) {
    const element = node as Element;
    if (element.hasAttribute("data-word-tab")) return "\t";
    if (element.localName === "br") return "\n";
    if (element.localName === "img") return "\ufffc";
  }
  return [...node.childNodes].map((child) => sourceText(child, ignored)).join("");
}

/** Canonical UTF-16 position, excluding presentation children and tab leaders. */
export function offsetAt(root: Node, target: Node, offset: number, ignored: IgnoreNode): number {
  let total = 0, found = false;
  const visit = (node: Node): void => {
    if (found || ignored(node)) return;
    if (node === target) {
      total += node.nodeType === 3 ? Math.min(offset, node.textContent!.length)
        : [...node.childNodes].slice(0, offset).reduce((sum, child) => sum + sourceText(child, ignored).length, 0);
      found = true; return;
    }
    if (node.nodeType === 3 || node.nodeType === 1 && (node as Element).matches("br,img,[data-word-tab]")) {
      total += sourceText(node, ignored).length; return;
    }
    for (const child of node.childNodes) visit(child);
  };
  visit(root); return total;
}

export interface InputStep {
  readonly from: number;
  readonly to: number;
  readonly nodes: ReadonlyMap<string, { readonly text: string; readonly start: number }>;
}
function sourceIndex(root: HTMLElement, ignored: IgnoreNode): Map<string, { text: string; start: number }> {
  const result = new Map<string, { text: string; start: number }>();
  let position = 0;
  const visit = (node: Node): string => {
    if (ignored(node)) return "";
    const id = node.nodeType === 1 ? (node as HTMLElement).dataset.nodeId : undefined;
    const entry = { text: "", start: position };
    if (id && !result.has(id)) result.set(id, entry);
    if (node.nodeType === 3 || node.nodeType === 1 && (node as Element).matches("br,img,[data-word-tab]")) {
      entry.text = sourceText(node, ignored); position += entry.text.length;
    } else entry.text = [...node.childNodes].map(visit).join("");
    return entry.text;
  };
  visit(root); return result;
}
export function inputStep(root: HTMLElement, range: Range, ignored: IgnoreNode): InputStep {
  return { nodes: sourceIndex(root, ignored), from: offsetAt(root, range.startContainer, range.startOffset, ignored), to: offsetAt(root, range.endContainer, range.endOffset, ignored) };
}

/** Evidence uses the actual target range. Ambiguous structural moves receive no guessed proof. */
export function inputProofs(root: HTMLElement, step: InputStep, ignored: IgnoreNode): ReconciledTextEdit[] {
  const result: ReconciledTextEdit[] = [];
  for (const [nodeId, next] of sourceIndex(root, ignored)) {
    const previous = step.nodes.get(nodeId); if (!previous) continue;
    const value = next.text, old = previous.text;
    if (old === value) continue;
    const from = Math.max(0, Math.min(old.length, step.from - previous.start));
    const to = Math.max(from, Math.min(old.length, step.to - previous.start));
    const length = value.length - old.length + to - from;
    if (length < 0 || value.slice(0, from) !== old.slice(0, from) || value.slice(from + length) !== old.slice(to)) continue;
    result.push({ nodeId, from, to, expectedText: old.slice(from, to), text: value.slice(from, from + length) });
  }
  return result;
}

export interface CaretBookmark { readonly nodeId?: string; readonly path: readonly number[]; readonly offset: number; readonly affinity?: "forward" }
export interface SelectionBookmark { readonly anchor: CaretBookmark; readonly focus: CaretBookmark }
export function bookmark(root: HTMLElement, ignored: IgnoreNode): SelectionBookmark | undefined {
  const selected = root.ownerDocument.getSelection();
  if (!selected?.anchorNode || !selected.focusNode || !root.contains(selected.anchorNode) || !root.contains(selected.focusNode)) return;
  const point = (node: Node, offset: number): CaretBookmark => {
    const parent = node.nodeType === 1 ? node as Element : node.parentElement;
    const block = parent?.closest<HTMLElement>(blockSelector), target = block && root.contains(block) ? block : root;
    const path: number[] = [];
    for (let child = target; child !== root && child.parentElement; child = child.parentElement)
      path.unshift([...child.parentElement.children].filter((item) => !ignored(item)).indexOf(child));
    return { ...(target.dataset.nodeId ? { nodeId: target.dataset.nodeId } : {}), path, offset: offsetAt(target, node, offset, ignored), ...(node.nodeType === 3 && offset === 0 ? { affinity: "forward" as const } : {}) };
  };
  return { anchor: point(selected.anchorNode, selected.anchorOffset), focus: point(selected.focusNode, selected.focusOffset) };
}
export function caretAt(root: Node, requested: number, ignored: IgnoreNode, affinity?: "forward"): [Node, number] {
  let remaining = requested, fallback: [Node, number] | undefined;
  const visit = (node: Node): [Node, number] | undefined => {
    if (ignored(node)) return;
    if (node.nodeType === 3) {
      if (remaining < node.textContent!.length || remaining === node.textContent!.length && affinity !== "forward") return [node, remaining];
      if (remaining === node.textContent!.length) fallback = [node, remaining];
      remaining -= node.textContent!.length; return;
    }
    if (node.nodeType === 1 && (node as Element).matches("br,img,[data-word-tab]")) {
      if (remaining === 0) return [node.parentNode!, [...node.parentNode!.childNodes].indexOf(node as ChildNode)];
      remaining--; return;
    }
    for (const child of node.childNodes) { const found = visit(child); if (found) return found; }
    return;
  };
  return visit(root) ?? fallback ?? [root, root.childNodes.length];
}
export function restoreBookmark(root: HTMLElement, value: SelectionBookmark | undefined, ignored: IgnoreNode): void {
  if (!value) return;
  const point = (entry: CaretBookmark): [Node, number] => {
    let target = entry.nodeId === root.dataset.nodeId ? root
      : entry.nodeId ? root.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(entry.nodeId) + '"]') : null;
    if (!target) {
      target = root;
      for (const ordinal of entry.path) target = [...target.children].filter((item) => !ignored(item))[ordinal] as HTMLElement ?? target;
    }
    return caretAt(target, entry.offset, ignored, entry.affinity);
  };
  const selected = root.ownerDocument.getSelection()!;
  selected.setBaseAndExtent(...point(value.anchor), ...point(value.focus));
}
