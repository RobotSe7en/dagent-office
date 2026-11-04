import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WorkbookDrawingAnchor, WorkbookDrawingMarker } from "@dagent-office/core";
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
export function projectWorkbookDrawing(rect: WorkbookRect, rotationDegrees: number, layout: WorkbookViewportLayout): readonly WorkbookDrawingProjection[] {
  if (![rect.x, rect.y, rect.width, rect.height, rotationDegrees].every(Number.isFinite) || rect.width < 0 || rect.height < 0)
    throw new DocumentError("InvalidOperation", "Workbook drawing geometry must be finite with nonnegative extents.");
  if (!rect.width || !rect.height) return [];
  const radians = (rotationDegrees % 360) * Math.PI / 180, cos = Math.abs(Math.cos(radians)), sin = Math.abs(Math.sin(radians));
  const width = rect.width * cos + rect.height * sin, height = rect.width * sin + rect.height * cos;
  const result: WorkbookDrawingProjection[] = [];
  for (const pane of layout.panes) {
    const placed = { ...rect, x: rect.x - pane.sourceLeftPt + pane.inlineStartPt, y: rect.y - pane.sourceTopPt + pane.blockStartPt };
    const bounds = { x: placed.x + (rect.width - width) / 2, y: placed.y + (rect.height - height) / 2, width, height };
    const clip = intersectRect(bounds, paneRect(pane));
    if (clip) result.push({ pane: pane.kind, rect: placed, clip });
  }
  return result;
}
