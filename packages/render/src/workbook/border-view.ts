import { DocumentError } from "@dagent-office/core";
import { svgElement } from "../presentation/svg.js";
import { markViewNode } from "../view-nodes.js";
import type { WorkbookBorderKind } from "./border-priority.js";
import type { WorkbookRect } from "./projection.js";

export interface WorkbookBorderEdge {
  readonly kind: WorkbookBorderKind; readonly color: string;
  readonly from: { readonly x: number; readonly y: number }; readonly to: { readonly x: number; readonly y: number };
}
interface Stroke { readonly width: number; readonly offset: number; readonly dash?: readonly number[] }
const profiles: Readonly<Partial<Record<WorkbookBorderKind, readonly Stroke[]>>> = {
  none: [], hair: [{ width: .1, offset: 0 }], thin: [{ width: .75, offset: 0 }], medium: [{ width: 1.75, offset: 0 }], thick: [{ width: 2.5, offset: 0 }],
  double: [{ width: .5, offset: -.625 }, { width: .5, offset: .625 }],
  dashed: [{ width: .75, offset: 0, dash: [3, 1] }], dotted: [{ width: .75, offset: 0, dash: [.5, 1] }],
  mediumDashed: [{ width: 1.75, offset: 0, dash: [8, 2.5] }],
  dashDot: [{ width: .75, offset: 0, dash: [8, 2.5, 2.5, 2.5] }], mediumDashDot: [{ width: 1.75, offset: 0, dash: [8, 2.5, 2.5, 2.5] }],
  dashDotDot: [{ width: .75, offset: 0, dash: [8, 2.5, 2.5, 2.5, 2.5, 2.5] }], mediumDashDotDot: [{ width: 1.75, offset: 0, dash: [8, 2.5, 2.5, 2.5, 2.5, 2.5] }],
};

/** Draw already-selected, physical edges in one clipped pane. Keep full edge
 * endpoints so scrolling a clipped dashed edge does not restart its pattern.
 * Point widths/patterns are calibrated against the pinned native PDF profile;
 * CSS collapsed borders cannot express these independent double-line widths.
 * Merge/CF/neighbor conflict composition belongs before this drawing boundary. */
export function createWorkbookBorderView(document: Document, bounds: WorkbookRect, edges: readonly WorkbookBorderEdge[], maxNodes = 40_001): SVGSVGElement {
  if (![bounds.x, bounds.y, bounds.width, bounds.height].every(Number.isFinite) || bounds.width <= 0 || bounds.height <= 0 || !Number.isSafeInteger(maxNodes) || maxNodes < 1)
    throw new DocumentError("InvalidOperation", "Worksheet borders require finite pane geometry and a positive node budget.");
  if (edges.length >= maxNodes) throw new DocumentError("ResourceLimitExceeded", "Worksheet borders exceed their edge budget.");
  const root = markViewNode(svgElement(document, "svg", { viewBox: `${bounds.x} ${bounds.y} ${bounds.width} ${bounds.height}`, "aria-hidden": "true" }));
  root.style.cssText = `position:absolute;pointer-events:none;overflow:hidden;left:${bounds.x}pt;top:${bounds.y}pt;width:${bounds.width}pt;height:${bounds.height}pt`;
  let nodes = 1;
  for (const edge of edges) {
    if (![edge.from.x, edge.from.y, edge.to.x, edge.to.y].every(Number.isFinite) || edge.from.x !== edge.to.x && edge.from.y !== edge.to.y)
      throw new DocumentError("InvalidDocument", "A worksheet border requires an axis-aligned finite edge.");
    if (!Object.hasOwn(profiles, edge.kind)) throw new DocumentError("UnsupportedFeature", "This worksheet border has no verified stroke geometry.");
    // Only a resolved sRGB color may reach SVG stroke. In particular, neither
    // a URL paint server nor an unresolved CSS variable can load a resource.
    const color = edge.color.length <= 80 ? /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*(\d*\.?\d+))?\s*\)$/.exec(edge.color) : null;
    if (!color || color.slice(1, 4).some(channel => Number(channel) > 255) || Number(color[4] ?? 1) > 1)
      throw new DocumentError("InvalidDocument", "A worksheet border requires a resolved sRGB color.");
    if (edge.from.x === edge.to.x && edge.from.y === edge.to.y) continue;
    for (const stroke of profiles[edge.kind]!) {
      if (++nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Worksheet border strokes exceed their node budget.");
      const x = edge.from.x === edge.to.x ? stroke.offset : 0, y = edge.from.y === edge.to.y ? stroke.offset : 0;
      root.append(svgElement(document, "line", { x1: edge.from.x + x, y1: edge.from.y + y, x2: edge.to.x + x, y2: edge.to.y + y,
        fill: "none", stroke: edge.color, "stroke-width": stroke.width, "stroke-linecap": "butt", ...(stroke.dash ? { "stroke-dasharray": stroke.dash.join(" ") } : {}) }));
    }
  }
  return root;
}
