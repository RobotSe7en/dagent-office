import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookDrawingAnchor, WorkbookDrawingMarker } from "@doffi/core";
import type { AxisDimensions } from "./dimensions.js";
import { intersectRect, paneRect } from "./projection.js";
import type { WorkbookRect } from "./projection.js";
import type { WorkbookPane, WorkbookViewportLayout } from "./viewport.js";

/** Resolve native cell anchors against the same measured dimensions as the
 * grid. Move/size response flags govern later structural edits, not this read. */
export function workbookDrawingRect(anchor: DeepReadonly<WorkbookDrawingAnchor>, dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }): WorkbookRect {
  const marker = (value: DeepReadonly<WorkbookDrawingMarker>) => ({ x: dimensions.columns.position(value.column) + value.offsetXPt, y: dimensions.rows.position(value.row) + value.offsetYPt });
  if (anchor.kind === "absolute") return { x: anchor.xPt, y: anchor.yPt, width: anchor.widthPt, height: anchor.heightPt };
  const from = marker(anchor.from);
  if (anchor.kind === "one-cell") return { ...from, width: anchor.widthPt, height: anchor.heightPt };
  const to = marker(anchor.to), width = to.x - from.x, height = to.y - from.y;
  if (width < 0 || height < 0) throw new DocumentError("InvalidDocument", "Drawing anchor endpoints resolve to a negative extent with the current worksheet dimensions.");
  return { ...from, width, height };
}
export interface WorkbookDrawingProjection { readonly pane: WorkbookPane["kind"]; readonly rect: WorkbookRect; readonly clip: WorkbookRect }

/** Visibility uses the rotated bounding rectangle, while the object itself
 * keeps the original source geometry and transform center in every pane. */
export function projectWorkbookDrawing(rect: WorkbookRect, rotationDegrees: number, layout: WorkbookViewportLayout, outsetPt = 0): readonly WorkbookDrawingProjection[] {
  if (![rect.x, rect.y, rect.width, rect.height, rotationDegrees, outsetPt].every(Number.isFinite) || rect.width < 0 || rect.height < 0 || outsetPt < 0)
    throw new DocumentError("InvalidOperation", "Workbook drawing geometry must be finite with nonnegative extents.");
  if ((!rect.width || !rect.height) && !outsetPt) return [];
  const radians = (rotationDegrees % 360) * Math.PI / 180, cos = Math.abs(Math.cos(radians)), sin = Math.abs(Math.sin(radians));
  // Outlines and effects can extend beyond the anchor. Keep a conservative
  // local outset so a visible stroke/shadow is neither culled nor clipped away.
  const width = (rect.width + 2 * outsetPt) * cos + (rect.height + 2 * outsetPt) * sin, height = (rect.width + 2 * outsetPt) * sin + (rect.height + 2 * outsetPt) * cos;
  if (![width, height].every(Number.isFinite)) throw new DocumentError("ResourceLimitExceeded", "Workbook drawing bounds exceed finite geometry.");
  const result: WorkbookDrawingProjection[] = [];
  for (const pane of layout.panes) {
    const placed = { ...rect, x: rect.x - pane.sourceLeftPt + pane.inlineStartPt, y: rect.y - pane.sourceTopPt + pane.blockStartPt };
    const bounds = { x: placed.x + (rect.width - width) / 2, y: placed.y + (rect.height - height) / 2, width, height };
    const clip = intersectRect(bounds, paneRect(pane));
    if (clip) result.push({ pane: pane.kind, rect: placed, clip });
  }
  return result;
}
