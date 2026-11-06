import type { WorkbookRange } from "../contracts/workbook.js";
import { invalid } from "../document/validation.js";
import { MAX_SHEET_COLUMNS } from "./chunk-index.js";

export function checkRange(range: WorkbookRange): void {
  if (range.from.row > range.to.row || range.from.column > range.to.column) invalid("Workbook ranges must be ordered.");
}
/** Sweep rows with a column segment tree; no quadratic pairwise merge scan. */
export function checkNonOverlappingRanges(ranges: readonly WorkbookRange[]): void {
  const events = ranges.flatMap(range => { checkRange(range); return [{ row: range.from.row, delta: 1, range }, { row: range.to.row + 1, delta: -1, range }]; }).sort((a, b) => a.row - b.row || a.delta - b.delta);
  const maximum = new Int32Array(MAX_SHEET_COLUMNS * 4), lazy = new Int32Array(MAX_SHEET_COLUMNS * 4);
  function update(node: number, left: number, right: number, from: number, to: number, delta: number): void {
    if (from <= left && right <= to) { maximum[node]! += delta; lazy[node]! += delta; return; }
    const middle = (left + right) >>> 1;
    if (from <= middle) update(node * 2, left, middle, from, to, delta);
    if (to > middle) update(node * 2 + 1, middle + 1, right, from, to, delta);
    maximum[node] = lazy[node]! + Math.max(maximum[node * 2]!, maximum[node * 2 + 1]!);
  }
  for (const event of events) {
    update(1, 0, MAX_SHEET_COLUMNS - 1, event.range.from.column, event.range.to.column, event.delta);
    if (maximum[1]! > 1) invalid("Workbook regions must not overlap.");
  }
}
