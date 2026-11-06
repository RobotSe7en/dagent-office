import { DocumentError } from "@doffi/core";
import type { WorkbookCoordinate } from "@doffi/core";
import type { AxisCell, AxisDimensions } from "./dimensions.js";
import type { WorkbookGridCell, WorkbookMergeIndex } from "./projection.js";
import type { WorkbookViewportLayout } from "./viewport.js";

/** Content anchors are insufficient for cell borders: covered perimeter cells
 * and adjacent cells can own a visible edge. Read those sources separately from
 * the DOM projection budget. Even a sheet-sized merge visits only visible axis
 * entries and its corners; hidden runs and the merge interior are not expanded.
 * This supplies source facts without deciding merge/CF border precedence. */
export function workbookAppearanceCoordinates(dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, layout: WorkbookViewportLayout,
  projections: readonly WorkbookGridCell[], merges: WorkbookMergeIndex, maxCells: number,
  extraColumns?: ReadonlyMap<WorkbookGridCell["pane"], readonly AxisCell[]>): readonly Readonly<WorkbookCoordinate>[] {
  if (!Number.isSafeInteger(maxCells) || maxCells <= 0) throw new DocumentError("InvalidOperation", "Worksheet appearance requires a positive source-cell budget.");
  const coordinates = new Map<number, Readonly<WorkbookCoordinate>>(), panes = new Map(layout.panes.map(pane => [pane.kind, pane]));
  const add = (row: number, column: number): void => {
    const key = row * 16_384 + column; if (coordinates.has(key)) return;
    if (coordinates.size >= maxCells) throw new DocumentError("ResourceLimitExceeded", "Worksheet appearance sources exceed their cell budget.");
    coordinates.set(key, Object.freeze({ row, column }));
  };
  const corners = (range: WorkbookGridCell["range"]): void => {
    add(range.from.row, range.from.column); add(range.from.row, range.to.column);
    add(range.to.row, range.from.column); add(range.to.row, range.to.column);
  };
  const neighbor = (row: number | undefined, column: number | undefined): void => {
    if (row === undefined || column === undefined) return;
    add(row, column); const merge = merges.at(row, column); if (merge) corners(merge);
  };
  const collect = (range: WorkbookGridCell["range"], pane: WorkbookViewportLayout["panes"][number]): void => {
    const beforeRow = dimensions.rows.neighbor(range.from.row, -1), afterRow = dimensions.rows.neighbor(range.to.row, 1);
    const beforeColumn = dimensions.columns.neighbor(range.from.column, -1), afterColumn = dimensions.columns.neighbor(range.to.column, 1);
    if (range.from.row === range.to.row && range.from.column === range.to.column) {
      add(range.from.row, range.from.column);
      neighbor(beforeRow, range.from.column); neighbor(afterRow, range.from.column); neighbor(range.from.row, beforeColumn); neighbor(range.from.row, afterColumn);
      return;
    }
    corners(range);
    for (const column of within(extraColumns?.get(pane.kind) ?? pane.columns, range.from.column, range.to.column)) {
      add(range.from.row, column); add(range.to.row, column); neighbor(beforeRow, column); neighbor(afterRow, column);
    }
    for (const row of within(pane.rows, range.from.row, range.to.row)) {
      add(row, range.from.column); add(row, range.to.column); neighbor(row, beforeColumn); neighbor(row, afterColumn);
    }
  };
  for (const projection of projections) {
    const pane = panes.get(projection.pane);
    if (!pane) throw new DocumentError("InvalidOperation", "A worksheet appearance projection has no matching pane.");
    collect(projection.range, pane);
  }
  if (extraColumns) for (const pane of layout.panes) {
    const seen = new Set<WorkbookGridCell["range"]>();
    for (const row of pane.rows) for (const column of extraColumns.get(pane.kind) ?? []) {
      const merge = merges.at(row.index, column.index); if (merge && seen.has(merge)) continue;
      if (merge) seen.add(merge);
      collect(merge ?? { from: { row: row.index, column: column.index }, to: { row: row.index, column: column.index } }, pane);
    }
  }
  return Object.freeze([...coordinates.values()].sort((a, b) => a.row - b.row || a.column - b.column));
}

function* within(cells: readonly AxisCell[], from: number, to: number): Generator<number> {
  let low = 0, high = cells.length;
  while (low < high) { const middle = (low + high) >>> 1; if (cells[middle]!.index < from) low = middle + 1; else high = middle; }
  for (let index = low; index < cells.length && cells[index]!.index <= to; index++) yield cells[index]!.index;
}
