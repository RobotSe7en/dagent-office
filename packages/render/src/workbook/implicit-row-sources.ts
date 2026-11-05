import type { DeepReadonly, WorkbookCoordinate, WorkbookSheet } from "@dagent-office/core";
import type { AxisDimensions } from "./dimensions.js";
import { checkWindowCancellation } from "./window-source.js";

/** Column formatting contributes only inside the used row interval. Explicit
 * row-format intervals also count when they contain no stored cell. A simple
 * native class has one typography per region; custom selectors need every
 * implicit coordinate and remain subject to the enclosing cumulative budget. */
export async function* implicitWorkbookRowSources(sheet: DeepReadonly<WorkbookSheet>, columns: AxisDimensions, occupied: ReadonlyMap<number, ReadonlySet<number>>,
  lastUsedRow: number, simpleClasses: boolean, automatic: (row: number) => boolean, mergeEnd: (row: number, column: number) => number | undefined,
  admit: () => void, signal: AbortSignal): AsyncGenerator<Readonly<WorkbookCoordinate>> {
  checkWindowCancellation(signal);
  const styledColumns = sheet.columnDimensions.filter(run => run.cssClass && !run.hidden);
  const rows = sheet.rowDimensions.filter(run => run.cssClass).map(run => ({ from: run.from, to: run.to }));
  if (styledColumns.length && lastUsedRow >= 0) rows.push({ from: 0, to: lastUsedRow });
  rows.sort((a, b) => a.from - b.from);
  const spans: { from: number; to: number }[] = [];
  for (const span of rows) {
    const previous = spans.at(-1);
    if (previous && span.from <= previous.to + 1) previous.to = Math.max(previous.to, span.to); else spans.push({ ...span });
  }
  let checkpoint = performance.now(), visitedRows = 0;
  for (const span of spans) for (let row = span.from; row <= span.to; row++) {
    if (++visitedRows % 256 === 0) {
      checkWindowCancellation(signal);
      if (performance.now() - checkpoint >= 8) { await new Promise<void>(resolve => setTimeout(resolve, 0)); checkpoint = performance.now(); }
    }
    if (!automatic(row)) continue;
    let low = 0, high = sheet.rowDimensions.length;
    while (low < high) { const mid = (low + high) >>> 1; if (sheet.rowDimensions[mid]!.to < row) low = mid + 1; else high = mid; }
    const run = sheet.rowDimensions[low], rowClass = run && run.from <= row ? run.cssClass : undefined;
    const regions = rowClass ? [{ from: 0, to: 16_383 }] : styledColumns;
    for (const region of regions) for (let column = region.from; column <= region.to;) {
      if (!columns.size(column)) { const next = columns.neighbor(column, 1); if (next === undefined) break; column = next; continue; }
      const end = mergeEnd(row, column); if (end !== undefined) { column = end + 1; continue; }
      if (occupied.get(row)?.has(column)) { column++; continue; }
      admit(); yield Object.freeze({ row, column });
      if (simpleClasses) break;
      column++;
    }
  }
  checkWindowCancellation(signal);
}
