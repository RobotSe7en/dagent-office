import type { DeepReadonly, WorkbookCoordinate, WorkbookRange, WorkbookSheet, WorkbookSnapshot } from "@dagent-office/core";
import type { AxisDimensions } from "./dimensions.js";
import { checkWindowCancellation } from "./window-source.js";

/** A color-only native rule cannot change a blank cell's font leading. For
 * general source selectors keep the whole declared region: selector ancestry
 * can change which font declaration wins. Tables also have regional/named
 * style layers, so their declared rectangles are resolved by the cell reader. */
export function implicitWorkbookFontRegions(snapshot: WorkbookSnapshot, sheet: DeepReadonly<WorkbookSheet>, simpleClasses: boolean): readonly DeepReadonly<WorkbookRange>[] {
  const fontClasses = new Set<string>();
  if (simpleClasses) for (const style of snapshot.styles) for (const match of style.css.matchAll(/\.([a-zA-Z_][-\w]*)\s*\{([^{}]*)\}/g)) {
    if (/(?:^|;)\s*(?:font(?:-[\w-]+)?|line-height|letter-spacing|word-spacing|--office-(?:font|minor|major)-[\w-]+)\s*:/i.test(match[2]!)) fontClasses.add(match[1]!);
  }
  return Object.freeze([...(sheet.tables ?? []).map(table => table.range), ...(sheet.conditionalRules ?? []).flatMap(rule =>
    rule.kind !== "colorScale" && rule.cssClass && (!simpleClasses || fontClasses.has(rule.cssClass)) ? rule.ranges : [])]);
}

/** Column formatting contributes only inside the used row interval. Explicit
 * row-format intervals also count when they contain no stored cell. A simple
 * native class has one typography per region; custom selectors and regional
 * table/conditional appearance need every implicit coordinate and remain
 * subject to the enclosing cumulative discovery budget. */
export async function* implicitWorkbookRowSources(sheet: DeepReadonly<WorkbookSheet>, columns: AxisDimensions, occupied: ReadonlyMap<number, ReadonlySet<number>>,
  lastUsedRow: number, simpleClasses: boolean, automatic: (row: number) => boolean, mergeEnd: (row: number, column: number) => number | undefined,
  admit: () => void, signal: AbortSignal, extraRegions: readonly DeepReadonly<WorkbookRange>[] = []): AsyncGenerator<Readonly<WorkbookCoordinate>> {
  checkWindowCancellation(signal);
  const styledColumns = sheet.columnDimensions.filter(run => run.cssClass);
  const rows = sheet.rowDimensions.filter(run => run.cssClass).map(run => ({ from: run.from, to: run.to }));
  rows.push(...extraRegions.map(region => ({ from: region.from.row, to: region.to.row })));
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
    const regions = (rowClass ? [{ from: 0, to: 16_383 }] : row <= lastUsedRow ? styledColumns : []).map(region => ({ from: region.from, to: region.to }));
    for (const region of extraRegions) {
      // Bound region inspection as well as emitted coordinates. Many sparse,
      // overlapping ranges must not multiply unaccounted work for every row.
      admit();
      if (row >= region.from.row && row <= region.to.row) regions.push({ from: region.from.column, to: region.to.column });
    }
    if (extraRegions.length) {
      regions.sort((a, b) => a.from - b.from); let count = 0;
      for (const region of regions) {
        const previous = regions[count - 1];
        if (previous && region.from <= previous.to + 1) previous.to = Math.max(previous.to, region.to);
        else regions[count++] = { from: region.from, to: region.to };
      }
      regions.length = count;
    }
    for (const region of regions) for (let column = region.from; column <= region.to;) {
      if (!columns.size(column)) { const next = columns.neighbor(column, 1); if (next === undefined) break; column = next; continue; }
      const end = mergeEnd(row, column); if (end !== undefined) { column = end + 1; continue; }
      if (occupied.get(row)?.has(column)) { column++; continue; }
      admit(); yield Object.freeze({ row, column });
      if (simpleClasses && !extraRegions.length) break;
      column++;
    }
  }
  checkWindowCancellation(signal);
}
