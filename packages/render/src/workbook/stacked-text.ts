import { DocumentError } from "@dagent-office/core";
import type { WorkbookTextProjection, WorkbookTextView } from "./text-view.js";

/** Stack complete graphemes using horizontal font metrics. A writing-mode
 * switch would substitute vertical glyph advances and punctuation forms.
 * Derived breaks add no source units; original text and UTF-16 offsets remain. */
export function stackWorkbookText(text: WorkbookTextView, maxNodes: number): WorkbookTextView {
  if (typeof Intl.Segmenter !== "function") throw new DocumentError("UnsupportedFeature", "Stacked worksheet text requires Unicode grapheme segmentation.");
  let source = text.projections;
  const mapped = source.length > 0;
  if (!mapped) {
    const derived: WorkbookTextProjection[] = [], document = text.element.ownerDocument;
    const walker = document.createTreeWalker(text.element, document.defaultView!.NodeFilter.SHOW_TEXT); let offset = 0;
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (derived.length >= maxNodes) throw new DocumentError("ResourceLimitExceeded", "Stacked numeric text exceeds its node budget.");
      const next = offset + (node as Text).length; derived.push({ node: node as Text, from: offset, to: next }); offset = next;
    }
    source = derived;
  }
  const content = source.map(projection => projection.node.data).join(""), boundaries: number[] = [];
  let nodes = text.nodes;
  for (const item of new Intl.Segmenter("und", { granularity: "grapheme" }).segment(content)) {
    const end = item.index + item.segment.length;
    if (end < content.length) boundaries.push(end);
    // Every break needs at least one node. Stop segmentation before a hostile
    // long cell can allocate an unbounded boundary array.
    if (nodes + boundaries.length > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Stacked worksheet text exceeds its node budget.");
  }
  const projections: WorkbookTextProjection[] = [], document = text.element.ownerDocument;
  const account = (): void => { if (++nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Stacked worksheet text exceeds its node budget."); };
  let index = 0;
  for (const projection of source) {
    let node = projection.node, from = projection.from;
    while (from < projection.to) {
      const boundary = boundaries[index], end = Math.min(boundary ?? projection.to, projection.to);
      let next: Text | undefined;
      if (end < projection.to) { account(); next = node.splitText(end - from); }
      let anchor: Node = node;
      if (/^[\r\n]+$/.test(node.data)) {
        // A native stacked line break occupies one empty character position.
        // Keep its actual text for projection while preventing a second break.
        account(); const span = document.createElement("span"); span.style.whiteSpace = "normal";
        node.parentNode!.insertBefore(span, node); span.append(node); anchor = span;
      }
      projections.push(Object.freeze({ node, from, to: end }));
      if (boundary === end) { account(); anchor.parentNode!.insertBefore(document.createElement("br"), anchor.nextSibling); index++; }
      from = end; if (next) node = next;
    }
  }
  return Object.freeze({ element: text.element, projections: Object.freeze(mapped ? projections : []), nodes });
}
