import { DocumentError } from "@dagent-office/core";
import type { WorkbookRange } from "@dagent-office/core";
import type { AxisDimensions } from "./dimensions.js";
import type { WorkbookCellBorders } from "./cell-borders.js";
import type { WorkbookBorderPaint } from "./border-style.js";
import type { WorkbookBorderEdge } from "./border-view.js";
import type { WorkbookMergeIndex, WorkbookRect } from "./projection.js";
import type { WorkbookPane, WorkbookViewportLayout } from "./viewport.js";
import { selectWorkbookBorder } from "./border-priority.js";
import { workbookCellShape, workbookCellShapeOffset } from "./cell-shape.js";
import type { WorkbookCellShape } from "./cell-shape.js";

export interface WorkbookPaneBorders { readonly pane: WorkbookPane["kind"]; readonly bounds: WorkbookRect; readonly edges: readonly WorkbookBorderEdge[] }
type Owner = WorkbookBorderPaint & { readonly range: Readonly<WorkbookRange> };

/** Compose shared physical edges from bounded visible axes and source corners.
 * A merge uses its upper-leading corner for top/before, upper-trailing corner
 * for after and lower-leading corner for bottom, as in the pinned native probe.
 * Outside neighbors can replace individual segments. Never expand hidden runs
 * or a merged interior; use the appearance query's auxiliary source records. */
export function projectWorkbookBorders(dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, layout: WorkbookViewportLayout, merges: WorkbookMergeIndex,
  read: (row: number, column: number) => WorkbookCellBorders | undefined, rightToLeft: boolean, maxEdges = 20_000,
  readRotation?: (row: number, column: number) => number): readonly WorkbookPaneBorders[] {
  if (!Number.isSafeInteger(maxEdges) || maxEdges < 1) throw new DocumentError("InvalidOperation", "Worksheet border projection requires a positive edge budget.");
  let count = 0;
  const cache = new Map<number, WorkbookCellBorders>();
  const shapes = new Map<number, WorkbookCellShape>();
  const shapeFor = (range: Readonly<WorkbookRange>): WorkbookCellShape => {
    const key = range.from.row * 16_384 + range.from.column, existing = shapes.get(key); if (existing) return existing;
    const height = dimensions.rows.position(range.to.row + 1) - dimensions.rows.position(range.from.row);
    const width = dimensions.columns.position(range.to.column + 1) - dimensions.columns.position(range.from.column);
    const shape = workbookCellShape(width, height, readRotation!(range.from.row, range.from.column), true);
    shapes.set(key, shape); return shape;
  };
  const source = (row: number, column: number): WorkbookCellBorders => {
    const key = row * 16_384 + column, existing = cache.get(key); if (existing) return existing;
    const value = read(row, column); if (!value) throw new DocumentError("InvalidOperation", "A worksheet border source is missing from its appearance window.");
    cache.set(key, value); return value;
  };
  const owner = (row: number | undefined, column: number | undefined, side: keyof WorkbookCellBorders): Owner | undefined => {
    if (row === undefined || column === undefined) return undefined;
    const merge = merges.at(row, column), range = merge ?? { from: { row, column }, to: { row, column } };
    const cellRow = side === "bottom" ? range.to.row : range.from.row, cellColumn = side === "after" ? range.to.column : range.from.column;
    const border = source(cellRow, cellColumn)[side];
    return border ? { ...border, range } : undefined;
  };
  const panes = layout.panes.map(pane => {
    const edges: WorkbookBorderEdge[] = [], seenRows = new Set<number>(), seenColumns = new Set<number>();
    const x = (position: number): number => {
      const logical = position - pane.sourceLeftPt + pane.inlineStartPt; return rightToLeft ? layout.widthPt - logical : logical;
    };
    const y = (position: number): number => position - pane.sourceTopPt + pane.blockStartPt;
    const append = (border: Owner | undefined, from: { x: number; y: number }, to: { x: number; y: number }, axis: "row" | "column"): void => {
      if (!border || border.kind === "none") return;
      if (++count > maxEdges) throw new DocumentError("ResourceLimitExceeded", "Worksheet border projection exceeds its edge budget.");
      const { range, ...paint } = border;
      const origin = axis === "row" ? x(dimensions.columns.position(rightToLeft ? range.to.column + 1 : range.from.column)) : y(dimensions.rows.position(range.from.row));
      let dashOffsetPt = (axis === "row" ? from.x : from.y) - origin;
      const shape = readRotation ? shapeFor(range) : undefined;
      if (shape?.topOffsetPt) {
        const top = y(dimensions.rows.position(range.from.row));
        from = { x: from.x + workbookCellShapeOffset(shape, from.y - top), y: from.y };
        to = { x: to.x + workbookCellShapeOffset(shape, to.y - top), y: to.y };
        if (axis === "column") dashOffsetPt *= Math.hypot(1, shape.topOffsetPt / shape.heightPt);
      }
      edges.push(Object.freeze({ ...paint, from: Object.freeze(from), to: Object.freeze(to), dashOffsetPt }));
    };
    for (const row of pane.rows) {
      for (const [first, second, boundary] of [
        [dimensions.rows.neighbor(row.index, -1), row.index, row.startPt],
        [row.index, dimensions.rows.neighbor(row.index, 1), row.startPt + row.sizePt],
      ] as const) {
        const key = second ?? dimensions.rows.count;
        if (seenRows.has(key)) continue; seenRows.add(key);
        for (const column of pane.columns) {
          const merge = first === undefined ? undefined : merges.at(first, column.index);
          if (merge && second !== undefined && merge === merges.at(second, column.index)) continue;
          const selected = selectWorkbookBorder(owner(first, column.index, "bottom"), owner(second, column.index, "top"), "row", rightToLeft);
          const a = x(column.startPt), b = x(column.startPt + column.sizePt), top = y(boundary);
          append(selected, { x: Math.min(a, b), y: top }, { x: Math.max(a, b), y: top }, "row");
        }
      }
    }
    for (const column of pane.columns) {
      for (const [first, second, boundary] of [
        [dimensions.columns.neighbor(column.index, -1), column.index, column.startPt],
        [column.index, dimensions.columns.neighbor(column.index, 1), column.startPt + column.sizePt],
      ] as const) {
        const key = second ?? dimensions.columns.count;
        if (seenColumns.has(key)) continue; seenColumns.add(key);
        for (const row of pane.rows) {
          const merge = first === undefined ? undefined : merges.at(row.index, first);
          if (merge && second !== undefined && merge === merges.at(row.index, second)) continue;
          const selected = selectWorkbookBorder(owner(row.index, first, "after"), owner(row.index, second, "before"), "column", rightToLeft), left = x(boundary);
          append(selected, { x: left, y: y(row.startPt) }, { x: left, y: y(row.startPt + row.sizePt) }, "column");
        }
      }
    }
    const bounds = Object.freeze({ x: rightToLeft ? layout.widthPt - pane.inlineStartPt - pane.widthPt : pane.inlineStartPt, y: pane.blockStartPt, width: pane.widthPt, height: pane.heightPt });
    return Object.freeze({ pane: pane.kind, bounds, edges: Object.freeze(edges) });
  });
  return Object.freeze(panes);
}
