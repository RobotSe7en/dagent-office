import type { WorkbookAutoFilter, WorkbookFilterDate, WorkbookMetadata, WorkbookRange, WorkbookSort, WorkbookTable, WorkbookTableStyleReference } from "../contracts/workbook.js";
import { checkLimit, invalid } from "../document/validation.js";
import { parseFormula } from "./formula.js";
import { checkRange, checkNonOverlappingRanges } from "./ranges.js";

function contains(outer: WorkbookRange, inner: WorkbookRange): boolean {
  return inner.from.row >= outer.from.row && inner.to.row <= outer.to.row && inner.from.column >= outer.from.column && inner.to.column <= outer.to.column;
}
function sortDefinition(sort: WorkbookSort, outer: WorkbookRange): number {
  checkRange(sort.range); if (!contains(outer, sort.range)) invalid("A sort definition exceeds its owning range.");
  const columns = new Set<number>();
  for (const key of sort.keys) {
    if (key.column < sort.range.from.column || key.column > sort.range.to.column || columns.has(key.column)) invalid("Sort keys must have unique columns within their range.");
    columns.add(key.column);
  }
  return columns.size + 1;
}
function dateFilter(date: WorkbookFilterDate): void {
  const fields = ["year", "month", "day", "hour", "minute", "second"] as const, end = fields.indexOf(date.grouping);
  for (const [index, key] of fields.entries()) if ((date[key] !== undefined) !== (index <= end)) invalid("A date filter must specify exactly the components of its grouping.");
  if (date.day !== undefined) {
    const leap = date.year % 4 === 0 && (date.year % 100 !== 0 || date.year % 400 === 0), days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    // Excel's historical 1900-02-29 is a supported serial-calendar value.
    if (date.day > days[date.month! - 1]! && !(date.year === 1900 && date.month === 2 && date.day === 29)) invalid("A date filter has an impossible calendar day.");
  }
}
function filterDefinition(filter: WorkbookAutoFilter, maxItems: number): number {
  checkRange(filter.range); let count = 1; const columns = new Set<number>();
  for (const column of filter.columns) {
    if (column.column < filter.range.from.column || column.column > filter.range.to.column || columns.has(column.column)) invalid("Filter columns must be unique and within their range.");
    columns.add(column.column); count++;
    const criterion = column.criterion;
    if (criterion?.kind === "values") { count += criterion.values.length + criterion.dates.length; checkLimit(count, maxItems); for (const date of criterion.dates) dateFilter(date); }
    if (criterion?.kind === "custom") count += criterion.tests.length;
    if (criterion?.kind === "top" && criterion.percent && criterion.count > 100) invalid("A percentage filter exceeds one hundred percent.");
    checkLimit(count, maxItems);
  }
  if (filter.sort) count += sortDefinition(filter.sort, filter.range);
  return count;
}
export function validateWorkbookTables(workbook: WorkbookMetadata, classExists: (name: string) => void, formatExists: (name: string | undefined) => void, maxItems: number): number {
  let count = 0;
  const cellStyles = new Set(workbook.cellStyles?.definitions.map(style => style.styleId));
  const regionCellStyles = (region: Pick<WorkbookTable, "headerCellStyleId" | "dataCellStyleId" | "totalCellStyleId">): void => {
    for (const key of ["headerCellStyleId", "dataCellStyleId", "totalCellStyleId"] as const) if (region[key] !== undefined) {
      count++; if (!cellStyles.has(region[key])) invalid("A table region refers to an undefined named cell style.");
    }
  };
  const styles = new Set<string>(), styleNames = new Set<string>(), tableIds = new Set<string>(), names = new Set((workbook.names ?? []).map(name => name.name.toUpperCase())), objectNames = new Set<string>();
  const differential = (style: { cssClass: string; numberFormatId?: string | undefined } | undefined): void => {
    if (style) { classExists(style.cssClass); formatExists(style.numberFormatId); count++; }
  };
  for (const style of workbook.tableStyles ?? []) {
    count += 1 + style.elements.length; checkLimit(count, maxItems);
    if (styles.has(style.styleId) || styleNames.has(style.name.toUpperCase())) invalid("Custom table style identities and names must be unique.");
    styles.add(style.styleId); styleNames.add(style.name.toUpperCase()); const regions = new Set<string>();
    for (const element of style.elements) {
      if (regions.has(element.region)) invalid("Custom table style regions must be unique."); regions.add(element.region); differential(element.style);
      if (!element.region.endsWith("Stripe") && element.size !== 1) invalid("Only table stripes may have a size greater than one.");
    }
  }
  const styleReference = (style: WorkbookTableStyleReference | undefined): void => {
    if (style?.kind === "custom" && !styles.has(style.styleId)) invalid("A custom table style reference is undefined.");
  };
  styleReference(workbook.defaultTableStyle);
  for (const sheet of workbook.sheets) {
    const tables = sheet.tables ?? [];
    checkNonOverlappingRanges([...tables.map(table => table.range), ...sheet.merges, ...(sheet.autoFilter ? [sheet.autoFilter.range] : [])]);
    if (sheet.autoFilter) count += filterDefinition(sheet.autoFilter, maxItems);
    for (const table of tables) {
      count += 1 + table.columns.length; checkLimit(count, maxItems); checkRange(table.range);
      const objectName = (table.objectName ?? table.name).toUpperCase(), token = parseFormula("=" + table.name, { row: 0, column: 0 });
      if (tableIds.has(table.tableId) || names.has(table.name.toUpperCase()) || objectNames.has(objectName) || token.kind !== "name" || token.sheet !== undefined || token.name !== table.name || /^[RC]$/i.test(table.name) || /^R\d+C\d+$/i.test(table.name) || table.name.toLowerCase().startsWith("_xlnm.")) invalid("Table identities, formula names and object names must be valid and unique.");
      tableIds.add(table.tableId); names.add(table.name.toUpperCase()); objectNames.add(objectName);
      const width = table.range.to.column - table.range.from.column + 1, height = table.range.to.row - table.range.from.row + 1;
      if (width !== table.columns.length || height < Number(table.header) + Number(table.totals) || table.totals && !table.totalsPreviouslyShown) invalid("Table dimensions, columns or total-row state are inconsistent.");
      styleReference(table.style);
      regionCellStyles(table);
      for (const key of ["headerStyle", "dataStyle", "totalStyle", "borderStyle", "headerBorderStyle", "totalBorderStyle"] as const) differential(table[key]);
      if (table.autoFilter) {
        const range = table.autoFilter.range, first = table.range.from, last = table.range.to;
        if (range.from.row !== first.row || range.from.column !== first.column || range.to.column !== last.column || range.to.row < last.row - Number(table.totals) || range.to.row > last.row) invalid("A table auto-filter must cover its data columns and body.");
        count += filterDefinition(table.autoFilter, maxItems);
      }
      if (table.sort) count += sortDefinition(table.sort, table.range);
      const columnIds = new Set<string>(), columnNames = new Set<string>();
      for (const [index, column] of table.columns.entries()) {
        if (columnIds.has(column.columnId) || columnNames.has(column.name.toUpperCase())) invalid("Table column identities and names must be unique within their table.");
        columnIds.add(column.columnId); columnNames.add(column.name.toUpperCase());
        if (column.totalFunction === "custom" && !column.totalFormula || column.totalFormula && column.totalFunction !== "custom") invalid("A custom table total requires exactly one custom formula source.");
        if (column.totalFormula && column.totalLabel !== undefined) invalid("A table total cell cannot contain both a formula and a label.");
        if (column.calculatedFormula) parseFormula(column.calculatedFormula, { row: table.range.from.row + Number(table.header), column: table.range.from.column + index });
        if (column.totalFormula) parseFormula(column.totalFormula, { row: table.range.to.row, column: table.range.from.column + index });
        for (const key of ["headerStyle", "dataStyle", "totalStyle"] as const) differential(column[key]);
        regionCellStyles(column);
      }
      checkLimit(count, maxItems);
    }
  }
  checkLimit(count, maxItems); return count;
}
