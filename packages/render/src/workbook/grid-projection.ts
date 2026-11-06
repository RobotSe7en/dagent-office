import { DocumentError } from "@dagent-office/core";
import type { WorkbookPaneBorders } from "./border-projection.js";
import type { WorkbookBorderEdge } from "./border-view.js";
import type { WorkbookViewportLayout } from "./viewport.js";
export const WORKBOOK_GRID_PAINT: Readonly<{ kind: "css"; style: "solid"; widthPt: number; color: string }> = Object.freeze({ kind: "css", style: "solid", widthPt: .75, color: "rgb(220, 224, 230)" });

/** Uniform grid lines on a sheet without merges. Visible axes already omit
 * hidden dimensions. Each boundary spans its pane's complete visible axes;
 * clipping still happens in the border view, including partial first cells.
 * Merged interiors require the ordinary shared-edge projection instead. */
export function projectWorkbookGridLines(layout: WorkbookViewportLayout, rightToLeft: boolean, maxEdges = 20_000): readonly WorkbookPaneBorders[] {
  if (!Number.isSafeInteger(maxEdges) || maxEdges < 1) throw new DocumentError("InvalidOperation", "Worksheet grid projection requires a positive edge budget.");
  let count = 0;
  return Object.freeze(layout.panes.map(pane => {
    const bounds = Object.freeze({ x: rightToLeft ? layout.widthPt - pane.inlineStartPt - pane.widthPt : pane.inlineStartPt,
      y: pane.blockStartPt, width: pane.widthPt, height: pane.heightPt });
    const edges: WorkbookBorderEdge[] = [];
    if (pane.rows.length && pane.columns.length) {
      count += pane.rows.length + pane.columns.length + 2;
      if (count > maxEdges) throw new DocumentError("ResourceLimitExceeded", "Worksheet grid projection exceeds its edge budget.");
      const x = (position: number): number => { const logical = position - pane.sourceLeftPt + pane.inlineStartPt; return rightToLeft ? layout.widthPt - logical : logical; };
      const y = (position: number): number => position - pane.sourceTopPt + pane.blockStartPt;
      const firstRow = pane.rows[0]!, lastRow = pane.rows.at(-1)!, firstColumn = pane.columns[0]!, lastColumn = pane.columns.at(-1)!;
      const a = x(firstColumn.startPt), b = x(lastColumn.startPt + lastColumn.sizePt), top = y(firstRow.startPt), bottom = y(lastRow.startPt + lastRow.sizePt);
      const append = (x1: number, y1: number, x2: number, y2: number): void => {
        edges.push(Object.freeze({ ...WORKBOOK_GRID_PAINT, from: Object.freeze({ x: x1, y: y1 }), to: Object.freeze({ x: x2, y: y2 }) }));
      };
      append(Math.min(a, b), top, Math.max(a, b), top);
      for (const row of pane.rows) { const position = y(row.startPt + row.sizePt); append(Math.min(a, b), position, Math.max(a, b), position); }
      append(a, top, a, bottom);
      for (const column of pane.columns) { const position = x(column.startPt + column.sizePt); append(position, top, position, bottom); }
    }
    return Object.freeze({ pane: pane.kind, bounds, edges: Object.freeze(edges) });
  }));
}
