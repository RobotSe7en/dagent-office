import { DocumentError } from "@dagent-office/core";
import type { WorkbookWindowFrame } from "./window-loader.js";
import type { WorkbookGridCell } from "./projection.js";
import { intersectRect, paneRect } from "./projection.js";
import type { WorkbookCellFill } from "./cell-fill.js";
import type { WorkbookCellShape } from "./cell-shape.js";
import type { WorkbookTextLayout } from "./text-layout.js";
import { createWorkbookShapedFillView } from "./shaped-cell-fill.js";
import { markViewNode } from "../view-nodes.js";

export interface WorkbookPaintedCell {
  readonly projection: WorkbookGridCell;
  readonly shape: WorkbookCellShape;
  readonly fill: WorkbookCellFill;
  readonly text?: WorkbookTextLayout;
}

/** Compose detached layers, retaining full-cell text geometry but allocating
 * only visible fill fragments. A rotated cell may spill sideways within its
 * pane; the row band, frozen boundaries and viewport still clip its display.
 * The caller publishes this whole result only after all other layers succeed. */
export function paintWorkbookWindow(document: Document, frame: WorkbookWindowFrame, cells: readonly WorkbookPaintedCell[], rightToLeft: boolean,
  borders: readonly SVGSVGElement[], maxNodes = 50_000, gridLines: readonly SVGSVGElement[] = []): HTMLElement {
  if (!Number.isSafeInteger(maxNodes) || maxNodes < 1 || cells.length !== frame.projections.length
    || cells.some((cell, index) => cell.projection !== frame.projections[index])) throw new DocumentError("InvalidOperation", "Workbook paint requires one prepared cell per projection and a positive node budget.");
  const root = markViewNode(document.createElement("div")), fills = document.createElement("div"), rows = new Map<number, HTMLElement>(), accessible = new Set<number>();
  root.dataset.officeView = "workbook-window"; root.dataset.sheetId = frame.sheetId;
  root.setAttribute("role", "grid"); root.setAttribute("aria-readonly", "true"); root.setAttribute("aria-rowcount", "1048576"); root.setAttribute("aria-colcount", "16384");
  root.dir = rightToLeft ? "rtl" : "ltr";
  root.style.cssText = `position:relative;overflow:hidden;isolation:isolate;background:white;width:${frame.layout.widthPt}pt;height:${frame.layout.heightPt}pt`;
  fills.dataset.officeView = "workbook-fills"; fills.setAttribute("aria-hidden", "true"); fills.style.display = "contents";
  let nodes = 2, id = 0;
  const account = (amount: number): void => { nodes += amount; if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook window display exceeds its cumulative node budget."); };
  const panes = new Map(frame.layout.panes.map(pane => { const rect = paneRect(pane); return [pane.kind, { ...rect, x: rightToLeft ? frame.layout.widthPt - rect.x - rect.width : rect.x }] as const; }));
  try {
    for (const grid of gridLines) { account(1 + grid.querySelectorAll("*").length); root.append(grid); }
    root.append(fills);
    for (const cell of cells) {
      const { projection, shape, text } = cell, { rect } = projection;
      const pane = panes.get(projection.pane)!;
      const x = rightToLeft ? frame.layout.widthPt - rect.x - rect.width : rect.x;
      const bounds = intersectRect({ x: x + Math.min(0, shape.topOffsetPt), y: rect.y, width: rect.width + Math.abs(shape.topOffsetPt), height: rect.height }, pane)!;
      // The projection guarantees a visible rectangular cell; the full slant
      // envelope contains that rectangle and thus always intersects its pane.
      const fill = createWorkbookShapedFillView(document, cell.fill, shape, `workbook-fill-${++id}`, {
        xPt: bounds.x - x, yPt: bounds.y - rect.y, widthPt: bounds.width, heightPt: bounds.height,
      }, Math.max(1, maxNodes - nodes));
      account(fill.nodes); fill.element.style.left = `${bounds.x}pt`; fill.element.style.top = `${bounds.y}pt`; fills.append(fill.element);
      let row = rows.get(projection.row);
      if (!row) {
        account(1); row = document.createElement("div"); row.setAttribute("role", "row"); row.setAttribute("aria-rowindex", String(projection.row + 1)); row.style.display = "contents"; rows.set(projection.row, row);
      }
      account(2 + (text ? 1 + text.text.nodes : 0));
      const clip = document.createElement("div"), body = document.createElement("div"), key = projection.row * 16_384 + projection.column;
      clip.dataset.officeView = "workbook-cell"; clip.dataset.row = String(projection.row); clip.dataset.column = String(projection.column); clip.dataset.pane = projection.pane;
      if (accessible.has(key)) clip.setAttribute("aria-hidden", "true");
      else {
        accessible.add(key); clip.setAttribute("role", "gridcell"); clip.setAttribute("aria-colindex", String(projection.column + 1));
        if (projection.range.to.row !== projection.row) clip.setAttribute("aria-rowspan", String(projection.range.to.row - projection.row + 1));
        if (projection.range.to.column !== projection.column) clip.setAttribute("aria-colspan", String(projection.range.to.column - projection.column + 1));
      }
      // The text's own clip enforces its row band. Sideways text may extend
      // beyond the parallelogram because glyphs can exceed the cell width.
      const textBounds = shape.topOffsetPt ? { x: pane.x, y: bounds.y, width: pane.width, height: bounds.height }
        : { ...projection.clip, x: rightToLeft ? frame.layout.widthPt - projection.clip.x - projection.clip.width : projection.clip.x };
      clip.style.cssText = `position:absolute;overflow:hidden;left:${textBounds.x}pt;top:${textBounds.y}pt;width:${textBounds.width}pt;height:${textBounds.height}pt`;
      body.style.cssText = `position:absolute;left:${x - textBounds.x}pt;top:${rect.y - textBounds.y}pt;width:${rect.width}pt;height:${rect.height}pt`;
      if (text) body.append(text.element);
      clip.append(body); row.append(clip);
    }
    for (const [, row] of [...rows].sort((a, b) => a[0] - b[0])) root.append(row);
    for (const border of borders) { account(1 + border.querySelectorAll("*").length); root.append(border); }
    return root;
  } catch (error) { root.replaceChildren(); throw error; }
}
