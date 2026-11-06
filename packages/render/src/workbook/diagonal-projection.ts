import { DocumentError } from "@doffi/core";
import type { WorkbookPaneBorders } from "./border-projection.js";
import type { WorkbookCellDiagonals } from "./cell-borders.js";
import type { WorkbookGridCell } from "./projection.js";
import type { WorkbookViewportLayout } from "./viewport.js";
import type { WorkbookBorderEdge } from "./border-view.js";
import { workbookBorderStrokes } from "./border-style.js";
import type { WorkbookRect } from "./projection.js";

function clipEdge(edge: WorkbookBorderEdge, bounds: WorkbookRect): WorkbookBorderEdge | undefined {
  const strokes = workbookBorderStrokes(edge), padding = Math.max(0, ...strokes.map(stroke => Math.abs(stroke.offset) + stroke.width / 2));
  const dx = edge.to.x - edge.from.x, dy = edge.to.y - edge.from.y, length = Math.hypot(dx, dy);
  if (!Number.isFinite(length)) throw new DocumentError("ResourceLimitExceeded", "Worksheet diagonal exceeds finite geometry.");
  if (!length) return undefined;
  let start = 0, end = 1;
  for (const [origin, delta, minimum, maximum] of [[edge.from.x, dx, bounds.x - padding, bounds.x + bounds.width + padding], [edge.from.y, dy, bounds.y - padding, bounds.y + bounds.height + padding]]) {
    if (!delta) { if (origin! < minimum! || origin! > maximum!) return undefined; continue; }
    const a = (minimum! - origin!) / delta, b = (maximum! - origin!) / delta;
    start = Math.max(start, Math.min(a, b)); end = Math.min(end, Math.max(a, b));
  }
  if (end <= start) return undefined;
  // Ordinary cells keep identical endpoints in every pane, avoiding renderer
  // rounding changes at newly introduced diagonal endpoints. Only large
  // logical cells need their line geometry reduced before entering SVG.
  if (Math.max(Math.abs(edge.from.x - bounds.x), Math.abs(edge.to.x - bounds.x), Math.abs(edge.from.y - bounds.y), Math.abs(edge.to.y - bounds.y)) <= 1_000_000) return Object.freeze({ ...edge, from: Object.freeze(edge.from), to: Object.freeze(edge.to) });
  const period = strokes[0]?.dash?.reduce((sum, part) => sum + part, 0), phase = period ? length * start % period : 0;
  return Object.freeze({ ...edge, from: Object.freeze({ x: edge.from.x + dx * start, y: edge.from.y + dy * start }),
    to: Object.freeze({ x: edge.from.x + dx * end, y: edge.from.y + dy * end }), ...(period ? { dashOffsetPt: phase } : {}) });
}

/** Diagonals belong to the full cell/merge anchor, not its shared perimeter.
 * Native up/down describe physical directions even when columns run RTL. */
export function projectWorkbookDiagonals(cells: readonly WorkbookGridCell[], layout: WorkbookViewportLayout,
  read: (row: number, column: number) => WorkbookCellDiagonals, rightToLeft: boolean, maxEdges = 10_000): readonly WorkbookPaneBorders[] {
  if (!Number.isSafeInteger(maxEdges) || maxEdges < 1) throw new DocumentError("InvalidOperation", "Worksheet diagonal projection requires a positive edge budget.");
  let count = 0;
  const cache = new Map<number, WorkbookCellDiagonals>(), panes = new Map(layout.panes.map(pane => [pane.kind, { pane: pane.kind,
    bounds: Object.freeze({ x: rightToLeft ? layout.widthPt - pane.inlineStartPt - pane.widthPt : pane.inlineStartPt, y: pane.blockStartPt, width: pane.widthPt, height: pane.heightPt }), edges: [] as WorkbookBorderEdge[] }]));
  for (const cell of cells) {
    const key = cell.row * 16_384 + cell.column;
    let paint = cache.get(key); if (!paint) { paint = read(cell.row, cell.column); cache.set(key, paint); }
    const { x, y, width, height } = cell.rect, left = rightToLeft ? layout.widthPt - x - width : x, right = left + width;
    for (const direction of ["up", "down"] as const) {
      const border = paint[direction]; if (!border || border.kind === "none") continue;
      const pane = panes.get(cell.pane); if (!pane) throw new DocumentError("InvalidOperation", "A worksheet diagonal has no projected pane.");
      const edge = clipEdge({ ...border, from: { x: left, y: direction === "up" ? y + height : y }, to: { x: right, y: direction === "up" ? y : y + height } }, pane.bounds);
      if (!edge) continue;
      if (++count > maxEdges) throw new DocumentError("ResourceLimitExceeded", "Worksheet diagonals exceed their edge budget.");
      pane.edges.push(edge);
    }
  }
  return Object.freeze([...panes.values()].map(pane => Object.freeze({ ...pane, edges: Object.freeze(pane.edges) })));
}
