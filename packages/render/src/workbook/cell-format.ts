import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WorkbookCell, WorkbookCoordinate, WorkbookSheet } from "@dagent-office/core";

export interface WorkbookCellFormat {
  readonly cssClass: string;
  readonly numberFormatId?: string;
}

type Dimension = Pick<DeepReadonly<WorkbookSheet>["rowDimensions"][number], "from" | "to" | "cssClass" | "numberFormatId">;
function dimensionAt(dimensions: readonly Dimension[], coordinate: number): Dimension | undefined {
  let low = 0, high = dimensions.length;
  while (low < high) { const mid = (low + high) >>> 1; if (dimensions[mid]!.to < coordinate) low = mid + 1; else high = mid; }
  const item = dimensions[low]; return item && item.from <= coordinate ? item : undefined;
}

/** Choose source format references, not CSS declarations. A cell's selected
 * class list still uses the source stylesheet's ordinary CSS cascade; class
 * order must never be mistaken for spreadsheet format precedence. Dimensions
 * come from a validated immutable snapshot and are never expanded per cell. */
export function workbookCellFormat(sheet: DeepReadonly<WorkbookSheet>, coordinate: WorkbookCoordinate, cell?: WorkbookCell): WorkbookCellFormat {
  const { row, column } = coordinate;
  if (!Number.isSafeInteger(row) || row < 0 || row >= 1_048_576 || !Number.isSafeInteger(column) || column < 0 || column >= 16_384
    || cell && (cell.sheetId !== sheet.sheetId || cell.row !== row || cell.column !== column)) {
    throw new DocumentError("InvalidOperation", "Cell format resolution requires a matching worksheet coordinate.");
  }
  const rowStyle = dimensionAt(sheet.rowDimensions, row), columnStyle = dimensionAt(sheet.columnDimensions, column);
  const cssClass = cell?.cssClass ?? rowStyle?.cssClass ?? columnStyle?.cssClass ?? sheet.defaultCssClass;
  const numberFormatId = cell?.formatId ?? rowStyle?.numberFormatId ?? columnStyle?.numberFormatId ?? sheet.defaultNumberFormatId;
  return Object.freeze({ cssClass, ...(numberFormatId === undefined ? {} : { numberFormatId }) });
}
