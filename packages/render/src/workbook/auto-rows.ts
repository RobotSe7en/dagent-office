import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WorkbookCell, WorkbookCoordinate, WorkbookNumberFormatter, WorkbookRuleReadLimits, WorkbookSheet, WorkbookWindow } from "@dagent-office/core";
import type { FontResolution } from "../contracts.js";
import { WorkbookCellSources } from "./cell-source.js";
import { layoutWorkbookRowTexts } from "./row-text-layout.js";
import type { WorkbookColumnFont } from "./column-font.js";
import { worksheetDimensions } from "./dimensions.js";
import type { WorkbookSurface } from "./surface.js";
import { WorkbookMergeIndex } from "./projection.js";
import { checkedWorkbookRules } from "./rule-response.js";
import type { WorkbookRuleSource } from "./rule-source.js";
import { cellSourceUnits, checkWindowCancellation, checkWindowIdentity } from "./window-source.js";
import type { WorkbookWindowSource } from "./window-source.js";
import { implicitWorkbookRowSources } from "./implicit-row-sources.js";
import { WorkbookRowSourcePrototypes } from "./row-source-prototypes.js";
import type { WorkbookSourceCell } from "./cell-source.js";
import { workbookRowFontKey, workbookRowMetricKey } from "./row-metric-cache.js";
import type { WorkbookRowTypography } from "./row-metric-cache.js";

export interface WorkbookAutoRowLimits {
  readonly maxCells: number;
  readonly maxReadCells: number;
  readonly maxCodeUnits: number;
}
export const DEFAULT_AUTO_ROW_LIMITS: WorkbookAutoRowLimits = Object.freeze({ maxCells: 2_500_000, maxReadCells: 8_000_000, maxCodeUnits: 256 * 1024 * 1024 });
export interface WorkbookAutoRows {
  readonly dimensions: ReturnType<typeof worksheetDimensions>;
  readonly fonts: readonly FontResolution[];
  readonly usedFallbackLineMetrics: boolean;
  readonly measuredCells: number;
}

/** Source height overrides remain sparse, and a hidden row is never expanded. */
function automatic(sheet: DeepReadonly<WorkbookSheet>, row: number): boolean {
  let low = 0, high = sheet.rowDimensions.length;
  while (low < high) { const mid = (low + high) >>> 1; if (sheet.rowDimensions[mid]!.to < row) low = mid + 1; else high = mid; }
  const run = sheet.rowDimensions[low], dimension = run && run.from <= row ? run : undefined;
  return !(dimension?.hidden ?? sheet.defaultRowsHidden) && dimension?.heightPt === undefined
    && (dimension?.autoHeight ?? !sheet.defaultRowHeightCustom);
}

/** Build one stable index before publishing any viewport for a worksheet. Read
 * only declared chunks, retain one bounded batch, and release every temporary
 * source/display node. Horizontal scrolling cannot change row prefix sums.
 * Imported formula caches remain imported; this pass does not mark them fresh. */
export async function measureWorkbookAutoRows(surface: Pick<WorkbookSurface, "source" | "view">, source: WorkbookWindowSource,
  sheetId: string, normal: WorkbookColumnFont, formatter: WorkbookNumberFormatter, signal: AbortSignal,
  options: Partial<WorkbookAutoRowLimits> = {}, rules?: { readonly source: WorkbookRuleSource; readonly limits: WorkbookRuleReadLimits }): Promise<WorkbookAutoRows> {
  const limits = { ...DEFAULT_AUTO_ROW_LIMITS, ...options }, snapshot = source.snapshot;
  if (Object.keys(limits).some(key => !Object.hasOwn(DEFAULT_AUTO_ROW_LIMITS, key)) || !Object.values(limits).every(value => Number.isSafeInteger(value) && value > 0))
    throw new DocumentError("InvalidOperation", "Automatic worksheet rows require supported positive budgets.");
  const selectedSheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId);
  if (!selectedSheet) throw new DocumentError("InvalidOperation", "Automatic rows require a source worksheet.");
  const sheet = selectedSheet;
  const dimensions = worksheetDimensions(sheet, normal.maximumDigitWidthPx), merges = new WorkbookMergeIndex(sheet.merges);
  const heights = new Map<number, number>(), fonts = new Map<string, FontResolution>();
  const formats = new Map(snapshot.metadata.workbook.numberFormats.map(format => [format.formatId, format]));
  // Native class-only styles have no coordinate/ancestor selectors. A custom
  // package with any other selector still resolves every cell independently.
  const simpleClasses = snapshot.styles.every(style => /^(?:\s*\.[a-zA-Z_][-\w]*\s*\{[^{}]*\}\s*)*$/.test(style.css));
  const typography = simpleClasses ? new Map<string, WorkbookRowTypography>() : undefined;
  const metricCache = new Map<string, number>(); let metricCacheUnits = 0;
  const fontCache = new Map<string, number>();
  const columnFonts = sheet.columnDimensions.some(run => run.cssClass && !run.hidden);
  const implicitFonts = columnFonts || sheet.rowDimensions.some(run => run.cssClass);
  const occupied = new Map<number, Set<number>>(); let lastUsedRow = -1;
  let visited = 0, units = 0, measuredCells = 0, usedFallbackLineMetrics = false, checkpoint = performance.now();
  const chunks = snapshot.manifest.chunks.filter(chunk => chunk.sheetId === sheetId);
  const identity = { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision };
  // Preflight sparse query slots before the first provider request. Fixed or
  // hidden rows need no measurement; column fonts still need their used bounds.
  const windows: WorkbookWindow[] = [];
  for (const [index, chunk] of chunks.entries()) {
    checkWindowCancellation(signal);
    if (index && index % 32 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); checkWindowCancellation(signal); }
    let start: number | undefined;
    for (let row = chunk.row; row <= chunk.row + 256; row++) {
      const include = row < chunk.row + 256 && (columnFonts || automatic(sheet, row));
      if (include && start === undefined) start = row;
      if (!include && start !== undefined) {
        visited += (row - start) * 32;
        if (visited > limits.maxReadCells) throw new DocumentError("ResourceLimitExceeded", "Automatic row discovery exceeds its cumulative query budget.");
        for (let from = start; from < row; from += 128) windows.push({ sheetId, row: from, column: chunk.column, rows: Math.min(128, row - from), columns: 32 });
        start = undefined;
      }
    }
  }
  const prototypes = simpleClasses && !sheet.conditionalRules?.length && !sheet.tables?.length ? new WorkbookRowSourcePrototypes(surface.source, snapshot, sheetId) : undefined;
  async function measure(coordinates: readonly Readonly<WorkbookCoordinate>[], byCoordinate: ReadonlyMap<number, WorkbookCell>): Promise<void> {
    checkWindowCancellation(signal);
    let current;
    if (sheet.conditionalRules?.length) {
      if (!rules) throw new DocumentError("InvalidOperation", "Automatic rule appearance requires an authorized current-value provider.");
      const request = Object.freeze({ ...identity, sheetId, ...rules.limits, coordinates: Object.freeze(coordinates) });
      current = checkedWorkbookRules(await rules.source.read(request, signal), request, snapshot); checkWindowCancellation(signal);
    }
    const shared = new Map<number, WorkbookSourceCell>(), ordinary: Readonly<WorkbookCoordinate>[] = [];
    for (const point of coordinates) {
      const key = point.row * 16_384 + point.column, cell = byCoordinate.get(key), source = cell && prototypes?.get(cell);
      if (source) shared.set(key, source); else ordinary.push(point);
    }
    const sources = ordinary.length ? new WorkbookCellSources(surface.source, snapshot, { ...identity, sheetId, sourceCoordinates: ordinary,
      getCell: (row, column) => byCoordinate.get(row * 16_384 + column) }, 50_000 - (prototypes?.nodes ?? 0), current) : undefined;
    try {
      const values = new Map(current?.cells.map(cell => [cell.row * 16_384 + cell.column, cell.current]));
      const inputs = coordinates.map(point => {
        const cell = byCoordinate.get(point.row * 16_384 + point.column), result = values.get(point.row * 16_384 + point.column);
        const key = point.row * 16_384 + point.column;
        return { source: shared.get(key) ?? sources!.get(point.row, point.column)!, value: result?.kind === "value" ? result.value : cell?.value, formula: Boolean(cell?.formula) || shared.has(key),
          widthPt: dimensions.columns.size(point.column), heightPt: sheet.defaultRowHeightPt, bordered: false };
      });
      const selected: typeof inputs = [];
      const keys = inputs.map(input => workbookRowMetricKey(input, formatter, formats, typography)), pending = new Map<string, number>();
      const indices = inputs.map((input, index) => {
        const key = keys[index]; if (key !== undefined && metricCache.has(key)) return -1;
        if (key !== undefined && pending.has(key)) return pending.get(key)!;
        const at = selected.length; selected.push(input); if (key !== undefined) pending.set(key, at); return at;
      });
      const measured = layoutWorkbookRowTexts(surface.view, snapshot, selected, formatter, normal.spaceWidthPx, fontCache, input => workbookRowFontKey(input, typography));
      usedFallbackLineMetrics ||= measured.usedFallbackLineMetrics;
      for (const font of measured.fonts) {
        const key = JSON.stringify(font);
        if (fonts.size >= 50_000 && !fonts.has(key)) throw new DocumentError("ResourceLimitExceeded", "Automatic row font diagnostics exceed their budget.");
        fonts.set(key, font);
      }
      for (const [index, cell] of coordinates.entries()) {
        const key = keys[index], required = indices[index] === -1 ? metricCache.get(key!) : measured.heights[indices[index]!];
        if (required === undefined) continue;
        if (key !== undefined && !metricCache.has(key) && metricCache.size < 4096 && metricCacheUnits + key.length <= 4 * 1024 * 1024) {
          metricCache.set(key, required); metricCacheUnits += key.length;
        }
        // Excel's native maximum is 409pt. Quantize upward to a CSS pixel so
        // the last line cannot be lost to fractional prefix-sum rounding.
        const height = Math.min(409, Math.max(sheet.defaultRowHeightPt, Math.ceil(required / .75) * .75));
        if (height > (heights.get(cell.row) ?? sheet.defaultRowHeightPt)) heights.set(cell.row, height);
      }
    } finally { sources?.dispose(); }
    if (performance.now() - checkpoint >= 8) { await new Promise<void>(resolve => setTimeout(resolve, 0)); checkpoint = performance.now(); }
  }
  try {
    for (const window of windows) {
      checkWindowCancellation(signal);
      const response = await source.read({ ...identity, windows: [window], maxCells: 4096, maxCodeUnits: Math.min(limits.maxCodeUnits, 8 * 1024 * 1024) }, signal);
      checkWindowCancellation(signal); checkWindowIdentity(response, snapshot);
      if (response.cells.length > window.rows * window.columns) throw new DocumentError("InvalidDocument", "Automatic row data exceeds its requested coordinates.");
      const seen = new Set<number>();
      const cells: WorkbookCell[] = [];
      for (const cell of response.cells) {
        if (cell.sheetId !== sheetId || !Number.isSafeInteger(cell.row) || !Number.isSafeInteger(cell.column) || cell.row < window.row || cell.row >= window.row + window.rows
          || cell.column < window.column || cell.column >= window.column + window.columns || seen.has(cell.row * 16_384 + cell.column))
          throw new DocumentError("InvalidDocument", "Automatic row data contains duplicate or unrequested coordinates.");
        seen.add(cell.row * 16_384 + cell.column);
        lastUsedRow = Math.max(lastUsedRow, cell.row);
        if (implicitFonts) { let columns = occupied.get(cell.row); if (!columns) { columns = new Set(); occupied.set(cell.row, columns); } columns.add(cell.column); }
        if (++measuredCells > limits.maxCells) throw new DocumentError("ResourceLimitExceeded", "Automatic rows exceed their cumulative source-cell budget.");
        units += cellSourceUnits(cell);
        if (units > limits.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Automatic rows exceed their cumulative source-text budget.");
        // Merged cells do not drive native AutoFit; their complete source extent
        // is retained by ordinary merge layout rather than distributed over rows.
        if (automatic(sheet, cell.row) && dimensions.columns.size(cell.column) > 0 && !merges.at(cell.row, cell.column)) cells.push(cell);
      }
      for (let offset = 0; offset < cells.length; offset += 128) {
        const batch = cells.slice(offset, offset + 128);
        await measure(batch.map(cell => Object.freeze({ row: cell.row, column: cell.column })), new Map(batch.map(cell => [cell.row * 16_384 + cell.column, cell])));
      }
    }
    let implicit: Readonly<WorkbookCoordinate>[] = [];
    const admit = (): void => {
      if (++visited > limits.maxReadCells) throw new DocumentError("ResourceLimitExceeded", "Automatic row discovery exceeds its cumulative query and implicit-coordinate budget.");
    };
    for await (const point of implicitWorkbookRowSources(sheet, dimensions.columns, occupied, lastUsedRow,
      simpleClasses && !sheet.conditionalRules?.length && !sheet.tables?.length, row => automatic(sheet, row), (row, column) => merges.at(row, column)?.to.column, admit, signal)) {
      implicit.push(point);
      if (implicit.length === 128) { await measure(implicit, new Map()); implicit = []; }
    }
    if (implicit.length) await measure(implicit, new Map());

    checkWindowCancellation(signal);
    return Object.freeze({ dimensions: worksheetDimensions(sheet, normal.maximumDigitWidthPx, heights), fonts: Object.freeze([...fonts.values()]), usedFallbackLineMetrics, measuredCells });
  } finally { prototypes?.dispose(); }
}
