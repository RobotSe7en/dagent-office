import { DocumentError } from "@dagent-office/core";
import type { WorkbookWindow } from "@dagent-office/core";
import type { AxisCell, AxisDimensions } from "./dimensions.js";

export interface WorkbookViewport {
  readonly widthPt: number; readonly heightPt: number;
  readonly scrollLeftPt: number; readonly scrollTopPt: number;
}
export interface WorkbookPane {
  readonly kind: "corner" | "frozenRows" | "frozenColumns" | "body";
  /** Logical inline coordinates work for both left-to-right and RTL sheets. */
  readonly inlineStartPt: number; readonly blockStartPt: number;
  readonly widthPt: number; readonly heightPt: number;
  readonly sourceLeftPt: number; readonly sourceTopPt: number;
  readonly rows: readonly AxisCell[]; readonly columns: readonly AxisCell[];
}
export interface WorkbookViewportLayout {
  readonly panes: readonly WorkbookPane[];
  readonly scrollLeftPt: number; readonly scrollTopPt: number;
  readonly maxScrollLeftPt: number; readonly maxScrollTopPt: number;
  readonly cellCount: number;
}

/** Partition frozen and scrolling regions in source coordinates. The same
 * aggregate budget covers all four panes, including their duplicated borders. */
export function workbookViewport(dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, frozen: { readonly rows: number; readonly columns: number }, viewport: WorkbookViewport, maxCells: number): WorkbookViewportLayout {
  if (![viewport.widthPt, viewport.heightPt, viewport.scrollLeftPt, viewport.scrollTopPt].every(Number.isFinite) || viewport.widthPt < 0 || viewport.heightPt < 0
    || !Number.isSafeInteger(maxCells) || maxCells <= 0 || !Number.isSafeInteger(frozen.rows) || !Number.isSafeInteger(frozen.columns)
    || frozen.rows < 0 || frozen.columns < 0 || frozen.rows >= dimensions.rows.count || frozen.columns >= dimensions.columns.count) throw new DocumentError("InvalidOperation", "Invalid workbook viewport or frozen boundaries.");
  const frozenWidth = dimensions.columns.position(frozen.columns), frozenHeight = dimensions.rows.position(frozen.rows);
  const fixedWidth = Math.min(viewport.widthPt, frozenWidth), fixedHeight = Math.min(viewport.heightPt, frozenHeight);
  const bodyWidth = viewport.widthPt - fixedWidth, bodyHeight = viewport.heightPt - fixedHeight;
  const maxScrollLeftPt = bodyWidth ? Math.max(0, dimensions.columns.extentPt - frozenWidth - bodyWidth) : 0;
  const maxScrollTopPt = bodyHeight ? Math.max(0, dimensions.rows.extentPt - frozenHeight - bodyHeight) : 0;
  const scrollLeftPt = Math.max(0, Math.min(maxScrollLeftPt, viewport.scrollLeftPt)), scrollTopPt = Math.max(0, Math.min(maxScrollTopPt, viewport.scrollTopPt));
  const panes: WorkbookPane[] = []; let cellCount = 0;
  const pane = (kind: WorkbookPane["kind"], left: number, top: number, width: number, height: number, sourceLeft: number, sourceTop: number): void => {
    if (!width || !height || sourceTop >= dimensions.rows.extentPt || sourceLeft >= dimensions.columns.extentPt) return;
    const rows = dimensions.rows.visible(sourceTop, sourceTop + height, maxCells), columns = dimensions.columns.visible(sourceLeft, sourceLeft + width, maxCells);
    if (!rows.length || !columns.length) return;
    cellCount += rows.length * columns.length;
    if (cellCount > maxCells) throw new DocumentError("ResourceLimitExceeded", "All workbook panes together exceed their visible-cell budget.");
    panes.push({ kind, inlineStartPt: left, blockStartPt: top, widthPt: width, heightPt: height, sourceLeftPt: sourceLeft, sourceTopPt: sourceTop, rows, columns });
  };
  pane("corner", 0, 0, fixedWidth, fixedHeight, 0, 0);
  pane("frozenRows", fixedWidth, 0, bodyWidth, fixedHeight, frozenWidth + scrollLeftPt, 0);
  pane("frozenColumns", 0, fixedHeight, fixedWidth, bodyHeight, 0, frozenHeight + scrollTopPt);
  pane("body", fixedWidth, fixedHeight, bodyWidth, bodyHeight, frozenWidth + scrollLeftPt, frozenHeight + scrollTopPt);
  return { panes, scrollLeftPt, scrollTopPt, maxScrollLeftPt, maxScrollTopPt, cellCount };
}

function runs(cells: readonly AxisCell[]): { start: number; count: number }[] {
  const result: { start: number; count: number }[] = [];
  for (const cell of cells) {
    const previous = result.at(-1);
    if (previous && previous.start + previous.count === cell.index) previous.count++;
    else result.push({ start: cell.index, count: 1 });
  }
  return result;
}
/** Hidden coordinate gaps become separate small queries, never a bounding
 * rectangle spanning a million rows. A Worker can transport these as one batch. */
export function viewportWindows(sheetId: string, layout: WorkbookViewportLayout): readonly WorkbookWindow[] {
  const windows: WorkbookWindow[] = [];
  for (const pane of layout.panes) {
    const columns = runs(pane.columns);
    for (const rows of runs(pane.rows)) for (const column of columns) windows.push({ sheetId, row: rows.start, column: column.start, rows: rows.count, columns: column.count });
  }
  return windows;
}
