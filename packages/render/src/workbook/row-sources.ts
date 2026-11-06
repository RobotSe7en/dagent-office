import { DocumentError } from "@dagent-office/core";
import type { WorkbookCoordinate, WorkbookSnapshot, WorkbookWindow } from "@dagent-office/core";
import type { WorkbookViewportLayout } from "./viewport.js";

export interface WorkbookRowSources {
  readonly coordinates: readonly Readonly<WorkbookCoordinate>[];
  readonly required: readonly Readonly<WorkbookCoordinate>[];
}

/** All actual chunk columns on displayed rows are needed to discover text
 * whose anchor is outside a viewport and to measure automatic row heights.
 * Empty worksheet space has no chunks and is never expanded. Query slots and
 * retained actual cells have independent caps. */
export function workbookRowSourceCoordinates(snapshot: WorkbookSnapshot, sheetId: string, layout: WorkbookViewportLayout,
  appearance: readonly Readonly<WorkbookCoordinate>[], maxReadCells: number, maxRequiredCells: number): WorkbookRowSources {
  if (![maxReadCells, maxRequiredCells].every(value => Number.isSafeInteger(value) && value > 0)) throw new DocumentError("InvalidOperation", "Row source discovery needs positive source budgets.");
  const rows = [...new Set(layout.panes.flatMap(pane => pane.rows.map(row => row.index)))].sort((a, b) => a - b);
  const required = new Map(appearance.map(point => [point.row * 16_384 + point.column, point]));
  for (const merge of snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)?.merges ?? []) {
    let low = 0, high = rows.length;
    while (low < high) { const middle = (low + high) >>> 1; if (rows[middle]! < merge.from.row) low = middle + 1; else high = middle; }
    if (low === rows.length || rows[low]! > merge.to.row) continue;
    for (const row of [merge.from.row, merge.to.row]) for (const column of [merge.from.column, merge.to.column]) {
      const key = row * 16_384 + column; if (required.has(key)) continue;
      if (required.size >= maxRequiredCells) throw new DocumentError("ResourceLimitExceeded", "Offscreen merge contexts exceed their retained source budget.");
      required.set(key, Object.freeze({ row, column }));
    }
  }
  const points = new Map(required);
  if (points.size > maxReadCells) throw new DocumentError("ResourceLimitExceeded", "Worksheet row discovery exceeds its query-slot budget.");
  for (const chunk of snapshot.manifest.chunks) {
    if (chunk.sheetId !== sheetId) continue;
    let low = 0, high = rows.length;
    while (low < high) { const middle = (low + high) >>> 1; if (rows[middle]! < chunk.row) low = middle + 1; else high = middle; }
    for (let index = low; index < rows.length && rows[index]! < chunk.row + 256; index++) for (let column = chunk.column; column < chunk.column + 32; column++) {
      const row = rows[index]!, key = row * 16_384 + column; if (points.has(key)) continue;
      if (points.size >= maxReadCells) throw new DocumentError("ResourceLimitExceeded", "Worksheet row discovery exceeds its query-slot budget.");
      points.set(key, Object.freeze({ row, column }));
    }
  }
  return Object.freeze({ coordinates: Object.freeze([...points.values()].sort((a, b) => a.row - b.row || a.column - b.column)), required: Object.freeze([...required.values()]) });
}

/** Preserve endpoint caps even when a sparse row query spans more coordinates
 * than one authorized transport request may admit. Input windows are disjoint;
 * each resulting slot is queried exactly once. */
export function workbookWindowBatches(windows: readonly WorkbookWindow[], maxCells: number): readonly (readonly WorkbookWindow[])[] {
  if (!Number.isSafeInteger(maxCells) || maxCells < 1) throw new DocumentError("InvalidOperation", "Window batching needs a positive coordinate budget.");
  const result: WorkbookWindow[][] = []; let current: WorkbookWindow[] = [], count = 0;
  for (const window of windows) {
    const columns = Math.min(window.columns, maxCells), rows = Math.max(1, Math.floor(maxCells / columns));
    for (let row = window.row; row < window.row + window.rows; row += rows) for (let column = window.column; column < window.column + window.columns; column += columns) {
      const part = Object.freeze({ sheetId: window.sheetId, row, column, rows: Math.min(rows, window.row + window.rows - row), columns: Math.min(columns, window.column + window.columns - column) });
      const size = part.rows * part.columns;
      if (count + size > maxCells) { result.push(current); current = []; count = 0; }
      current.push(part); count += size;
    }
  }
  if (current.length) result.push(current);
  return Object.freeze(result.map(batch => Object.freeze(batch)));
}
