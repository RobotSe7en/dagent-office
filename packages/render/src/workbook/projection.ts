import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WorkbookCoordinate, WorkbookRange, WorkbookWindow } from "@dagent-office/core";
import type { AxisDimensions } from "./dimensions.js";
import type { WorkbookPane, WorkbookViewportLayout } from "./viewport.js";

type Range = DeepReadonly<WorkbookRange>;
export interface WorkbookRect { readonly x: number; readonly y: number; readonly width: number; readonly height: number }
export interface WorkbookGridCell {
  readonly pane: WorkbookPane["kind"];
  /** Source anchor, including a merged cell whose anchor is outside the viewport. */
  readonly row: number; readonly column: number;
  readonly range: Range;
  /** Logical inline/block coordinates relative to the entire viewport. */
  readonly rect: WorkbookRect; readonly clip: WorkbookRect;
}
interface Branch { readonly range: Range; readonly start: number; readonly end: number; readonly fromColumn: number; readonly toColumn: number; readonly left?: Branch; readonly right?: Branch }

/** The core already validates nonoverlapping merges. This derived interval tree
 * indexes their row spans without expanding any covered source cells. */
export class WorkbookMergeIndex {
  private readonly root: Branch | undefined;
  constructor(ranges: readonly Range[]) {
    const sorted = [...ranges].sort((a, b) => a.from.row - b.from.row || a.from.column - b.from.column);
    function build(from: number, to: number): Branch | undefined {
      if (from === to) return undefined;
      const middle = (from + to) >>> 1, range = sorted[middle]!, left = build(from, middle), right = build(middle + 1, to);
      return { range, start: Math.min(range.from.row, left?.start ?? Infinity), end: Math.max(range.to.row, left?.end ?? -1, right?.end ?? -1),
        fromColumn: Math.min(range.from.column, left?.fromColumn ?? Infinity, right?.fromColumn ?? Infinity),
        toColumn: Math.max(range.to.column, left?.toColumn ?? -1, right?.toColumn ?? -1), ...(left ? { left } : {}), ...(right ? { right } : {}) };
    }
    this.root = build(0, sorted.length);
  }
  at(row: number, column: number): Range | undefined {
    function find(branch: Branch | undefined): Range | undefined {
      if (!branch || branch.start > row || branch.end < row || branch.fromColumn > column || branch.toColumn < column) return undefined;
      const left = find(branch.left); if (left) return left;
      const range = branch.range;
      if (range.from.row > row) return undefined;
      if (range.to.row >= row && range.from.column <= column && column <= range.to.column) return range;
      return find(branch.right);
    }
    return find(this.root);
  }
}

export function intersectRect(a: WorkbookRect, b: WorkbookRect): WorkbookRect | undefined {
  const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y), right = Math.min(a.x + a.width, b.x + b.width), bottom = Math.min(a.y + a.height, b.y + b.height);
  return right > x && bottom > y ? { x, y, width: right - x, height: bottom - y } : undefined;
}
export function paneRect(pane: WorkbookPane): WorkbookRect { return { x: pane.inlineStartPt, y: pane.blockStartPt, width: pane.widthPt, height: pane.heightPt }; }

/** Each merged source cell may have several clipped pane projections, but it
 * always keeps the same anchor and full geometry. Text is never re-centered in
 * the clipped fragment. Invisible source cells are never materialized. */
export function projectWorkbookCells(dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, layout: WorkbookViewportLayout, merges: WorkbookMergeIndex, maxCells: number): readonly WorkbookGridCell[] {
  if (!Number.isSafeInteger(maxCells) || maxCells <= 0) throw new DocumentError("InvalidOperation", "A workbook projection requires a positive cell budget.");
  const result: WorkbookGridCell[] = [];
  for (const pane of layout.panes) {
    const projected = new Set<Range>(), bounds = paneRect(pane);
    for (const row of pane.rows) for (const column of pane.columns) {
      const merge = merges.at(row.index, column.index);
      if (merge && projected.has(merge)) continue;
      if (merge) projected.add(merge);
      const range = merge ?? { from: { row: row.index, column: column.index }, to: { row: row.index, column: column.index } };
      const x = dimensions.columns.position(range.from.column), y = dimensions.rows.position(range.from.row);
      const rect = { x: x - pane.sourceLeftPt + pane.inlineStartPt, y: y - pane.sourceTopPt + pane.blockStartPt,
        width: dimensions.columns.position(range.to.column + 1) - x, height: dimensions.rows.position(range.to.row + 1) - y };
      const clip = intersectRect(rect, bounds); if (!clip) continue;
      if (result.length >= maxCells) throw new DocumentError("ResourceLimitExceeded", "The workbook source projection exceeds its cell budget.");
      result.push({ pane: pane.kind, row: range.from.row, column: range.from.column, range, rect, clip });
    }
  }
  return result;
}

/** Query each displayed source anchor once, including hidden/offscreen merge
 * origins. Adjacent anchors share windows; covered cells are not read. */
export function projectionWindows(sheetId: string, cells: readonly WorkbookGridCell[]): readonly WorkbookWindow[] {
  const byRow = new Map<number, Set<number>>();
  for (const cell of cells) {
    let columns = byRow.get(cell.row); if (!columns) { columns = new Set(); byRow.set(cell.row, columns); }
    columns.add(cell.column);
  }
  const result: { sheetId: string; row: number; column: number; rows: number; columns: number }[] = [];
  let previous = new Map<string, number>();
  for (const [row, values] of [...byRow].sort((a, b) => a[0] - b[0])) {
    const columns = [...values].sort((a, b) => a - b), next = new Map<string, number>();
    for (let start = 0; start < columns.length;) {
      let end = start + 1; while (end < columns.length && columns[end] === columns[end - 1]! + 1) end++;
      const column = columns[start]!, count = end - start, key = `${column}:${count}`, existing = previous.get(key), window = existing === undefined ? undefined : result[existing];
      if (window && window.row + window.rows === row) { window.rows++; next.set(key, existing!); }
      else { next.set(key, result.length); result.push({ sheetId, row, column, rows: 1, columns: count }); }
      start = end;
    }
    previous = next;
  }
  return result;
}

/** Input is logical inline/block distance, independent of DOM RTL scrollLeft.
 * Half-open pane bounds make a frozen edge belong to exactly one pane. */
export function hitWorkbookCell(dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, layout: WorkbookViewportLayout, merges: WorkbookMergeIndex, x: number, y: number): Readonly<WorkbookCoordinate> | undefined {
  if (!Number.isFinite(x) || !Number.isFinite(y)) throw new DocumentError("InvalidOperation", "Workbook hit coordinates must be finite.");
  const pane = layout.panes.find(pane => x >= pane.inlineStartPt && x < pane.inlineStartPt + pane.widthPt && y >= pane.blockStartPt && y < pane.blockStartPt + pane.heightPt);
  if (!pane) return undefined;
  const row = dimensions.rows.hit(y - pane.blockStartPt + pane.sourceTopPt), column = dimensions.columns.hit(x - pane.inlineStartPt + pane.sourceLeftPt);
  if (row === undefined || column === undefined) return undefined;
  return merges.at(row, column)?.from ?? { row, column };
}
