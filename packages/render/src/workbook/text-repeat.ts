import { DocumentError } from "@dagent-office/core";
import type { WorkbookTextProjection, WorkbookTextView } from "./text-view.js";

/** Repeat complete displayed values. Clones retain local rich styles and every
 * original @/text offset; generated repetitions never become persisted text. */
export function repeatWorkbookText(text: WorkbookTextView, availableWidth: number, maxNodes: number): WorkbookTextView {
  const width = text.element.getBoundingClientRect().width, units = text.element.textContent?.length ?? 0;
  if (!Number.isFinite(width) || width < 0 || !Number.isFinite(availableWidth) || availableWidth < 0) throw new DocumentError("InvalidOperation", "Fill alignment requires finite text geometry.");
  if (!width || !units) return text;
  const count = Math.max(1, Math.floor((availableWidth + 1e-7) / width));
  if (count === 1) return text;
  const nodes = 1 + count * text.nodes;
  if (!Number.isSafeInteger(nodes) || nodes > maxNodes || count * units > 32767) throw new DocumentError("ResourceLimitExceeded", "Filled cell text exceeds its display budget.");
  const document = text.element.ownerDocument, root = text.element.cloneNode(false) as HTMLElement, projections: WorkbookTextProjection[] = [];
  const mappings = new Map(text.projections.map(projection => [projection.node, projection]));
  const original = document.createTreeWalker(text.element, document.defaultView!.NodeFilter.SHOW_TEXT), source: Text[] = [];
  for (let node = original.nextNode(); node; node = original.nextNode()) source.push(node as Text);
  for (let index = 0; index < count; index++) {
    const clone = text.element.cloneNode(true) as HTMLElement;
    Object.assign(clone.style, { display: "inline", position: "static", width: "auto", height: "auto", transform: "none", whiteSpace: "pre" });
    const walker = document.createTreeWalker(clone, document.defaultView!.NodeFilter.SHOW_TEXT); let part = 0;
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const mapping = mappings.get(source[part++]!); if (mapping) projections.push(Object.freeze({ node: node as Text, from: mapping.from, to: mapping.to }));
    }
    root.append(clone);
  }
  text.element.replaceWith(root);
  return Object.freeze({ element: root, projections: Object.freeze(projections), nodes });
}
