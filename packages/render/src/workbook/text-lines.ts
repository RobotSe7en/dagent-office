import { DocumentError } from "@dagent-office/core";
import type { WorkbookTextProjection, WorkbookTextView } from "./text-view.js";

/** Split already shaped lines without flattening rich spans or inventing source
 * offsets. Browser line boxes decide wrapping; this layer only positions the
 * resulting lines. All glyph probes and materialized ancestors are bounded. */
export function distributeWorkbookTextLines(text: WorkbookTextView, availableHeight: number, maxNodes: number): { readonly text: WorkbookTextView; readonly naturalHeight: number } {
  const root = text.element, document = root.ownerDocument, view = document.defaultView!, naturalHeight = root.getBoundingClientRect().height;
  if (!Number.isFinite(availableHeight) || availableHeight < 0 || !Number.isFinite(naturalHeight)) throw new DocumentError("InvalidOperation", "Vertical text distribution requires finite line geometry.");
  const walker = document.createTreeWalker(root, view.NodeFilter.SHOW_TEXT | view.NodeFilter.SHOW_ELEMENT), source: (Text | HTMLElement)[] = [], rectangles: { top: number; bottom: number }[] = [];
  const range = document.createRange(), mapped = new Map(text.projections.map(projection => [projection.node, projection]));
  let probes = 0;
  const probe = (): void => { if (++probes > Math.min(100_000, maxNodes * 8)) throw new DocumentError("ResourceLimitExceeded", "Worksheet line inspection exceeds its geometry budget."); };
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    probe();
    if (node.nodeType !== 3) { if ((node as HTMLElement).hasAttribute("data-office-number-space")) source.push(node as HTMLElement); continue; }
    source.push(node as Text); range.selectNodeContents(node);
    for (const rect of range.getClientRects()) if (rect.height > 0) { probe(); rectangles.push({ top: rect.top, bottom: rect.bottom }); }
  }
  rectangles.sort((a, b) => a.top - b.top || a.bottom - b.bottom);
  const lines: { top: number; bottom: number }[] = [];
  for (const rect of rectangles) {
    const last = lines.at(-1);
    if (last && rect.top < last.bottom - .1 && rect.bottom > last.top + .1) { last.top = Math.min(last.top, rect.top); last.bottom = Math.max(last.bottom, rect.bottom); }
    else lines.push({ ...rect });
  }
  if (lines.length <= 1 || naturalHeight >= availableHeight) return { text, naturalHeight };
  if (lines.length * 2 + 1 > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Distributed lines exceed their node budget.");
  let previousLine = 0;
  const lineAt = (node: Text, offset: number): number => {
    probe(); range.setStart(node, offset); range.setEnd(node, Math.min(node.length, offset + 1));
    let rect = [...range.getClientRects()].find(rect => rect.height > 0);
    if (!rect && offset > 0) { probe(); range.setStart(node, offset - 1); rect = [...range.getClientRects()].find(rect => rect.height > 0); }
    if (!rect) return previousLine;
    let low = 0, high = lines.length;
    while (low < high) { const mid = (low + high) >>> 1; if (lines[mid]!.bottom <= rect.top + .1) low = mid + 1; else high = mid; }
    return Math.min(low, lines.length - 1);
  };
  let nodes = 1;
  const account = (): void => { if (++nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Distributed rich text exceeds its node budget."); };
  const result = root.cloneNode(false) as HTMLElement, projections: WorkbookTextProjection[] = [];
  const rows = lines.map(line => {
    account(); const element = root.cloneNode(false) as HTMLElement;
    Object.assign(element.style, { position: "relative", display: "block", left: "auto", top: "auto", width: "100%", height: "auto", transform: "none", whiteSpace: "pre", flex: "none" });
    element.style.minHeight = Math.max(0, line.bottom - line.top) + "px";
    result.append(element); return { element, parents: new Map<HTMLElement, HTMLElement>([[root, element]]) };
  });
  function parentFor(original: HTMLElement, line: number): HTMLElement {
    const item = rows[line]!, known = item.parents.get(original); if (known) return known;
    if (!root.contains(original) || !original.parentElement) throw new DocumentError("InvalidOperation", "A distributed line escaped its source text.");
    const parent = parentFor(original.parentElement, line); account(); const clone = original.cloneNode(false) as HTMLElement;
    parent.append(clone); item.parents.set(original, clone); return clone;
  }
  for (const entry of source) {
    if (entry.nodeType !== 3) {
      const spacer = entry as HTMLElement; probe(); const rect = spacer.getBoundingClientRect(); let line = lines.findIndex(item => item.bottom > rect.top + .1); if (line < 0) line = lines.length - 1;
      account(); parentFor(spacer.parentElement!, line).append(spacer.cloneNode(false)); continue;
    }
    const node = entry as Text;
    let from = 0;
    while (from < node.length) {
      const line = lineAt(node, from); let low = from + 1, high = node.length;
      // Find the next line transition using UTF-16 range positions. A browser
      // cannot wrap within a grapheme; surrogate/variation-sequence halves have
      // the same line and therefore remain in the same final fragment.
      while (low < high) { const middle = (low + high) >>> 1; if (lineAt(node, middle) === line) low = middle + 1; else high = middle; }
      const to = low, content = node.data.slice(from, to), parent = parentFor(node.parentElement!, line);
      previousLine = line;
      let offset = from;
      for (const segment of content.match(/[^\r\n]+|[\r\n]+/g) ?? []) {
        account(); const clone = document.createTextNode(segment), mapping = mapped.get(node);
        if (/^[\r\n]+$/.test(segment)) {
          // Keep native newline units without generating a second line break.
          account(); const span = document.createElement("span"); span.style.whiteSpace = "normal"; span.append(clone); parent.append(span);
        } else parent.append(clone);
        if (mapping) projections.push(Object.freeze({ node: clone, from: mapping.from + offset, to: mapping.from + offset + segment.length }));
        offset += segment.length;
      }
      from = to;
    }
  }
  Object.assign(result.style, { display: "flex", flexDirection: "column", justifyContent: "space-between", height: availableHeight + "px" });
  root.replaceWith(result);
  return { text: Object.freeze({ element: result, nodes, projections: Object.freeze(projections) }), naturalHeight };
}
