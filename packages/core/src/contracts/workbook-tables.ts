import { z } from "zod";
import { IdentifierSchema } from "./primitives.js";
import { WorkbookColumnSchema, WorkbookRangeSchema } from "./workbook-coordinates.js";
import { WorkbookRuleFormulaSchema } from "./workbook-rules.js";

export const WorkbookFilterComparisonSchema = z.enum(["equal", "notEqual", "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual"]);
export const WorkbookFilterDateSchema = z.strictObject({
  grouping: z.enum(["year", "month", "day", "hour", "minute", "second"]), year: z.number().int().min(1).max(9999),
  month: z.number().int().min(1).max(12).optional(), day: z.number().int().min(1).max(31).optional(),
  hour: z.number().int().min(0).max(23).optional(), minute: z.number().int().min(0).max(59).optional(), second: z.number().int().min(0).max(59).optional(),
});
export const WorkbookDynamicFilterSchema = z.enum(["aboveAverage", "belowAverage", "tomorrow", "today", "yesterday", "nextWeek", "thisWeek", "lastWeek", "nextMonth", "thisMonth", "lastMonth", "nextQuarter", "thisQuarter", "lastQuarter", "nextYear", "thisYear", "lastYear", "yearToDate", "Q1", "Q2", "Q3", "Q4", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12"]);
export const WorkbookFilterCriterionSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("values"), blank: z.boolean(), values: z.array(z.string().max(32767)), dates: z.array(WorkbookFilterDateSchema) }),
  z.strictObject({ kind: z.literal("custom"), and: z.boolean(), tests: z.array(z.strictObject({ operator: WorkbookFilterComparisonSchema, value: z.string().max(32767) })).min(1).max(2) }),
  z.strictObject({ kind: z.literal("top"), top: z.boolean(), percent: z.boolean(), count: z.number().min(1).max(500) }),
  z.strictObject({ kind: z.literal("dynamic"), type: WorkbookDynamicFilterSchema }),
]);
export const WorkbookSortSchema = z.strictObject({
  range: WorkbookRangeSchema, caseSensitive: z.boolean(), method: z.enum(["none", "pinYin", "stroke"]),
  keys: z.array(z.strictObject({ column: WorkbookColumnSchema, descending: z.boolean() })).min(1).max(64),
});
export const WorkbookAutoFilterSchema = z.strictObject({
  range: WorkbookRangeSchema,
  columns: z.array(z.strictObject({ column: WorkbookColumnSchema, hiddenButton: z.boolean(), showButton: z.boolean(), criterion: WorkbookFilterCriterionSchema.optional() })),
  sort: WorkbookSortSchema.optional(),
});
export type WorkbookAutoFilter = z.infer<typeof WorkbookAutoFilterSchema>;
export type WorkbookFilterCriterion = z.infer<typeof WorkbookFilterCriterionSchema>;
export type WorkbookFilterDate = z.infer<typeof WorkbookFilterDateSchema>;
export type WorkbookSort = z.infer<typeof WorkbookSortSchema>;

export const WorkbookDifferentialSchema = z.strictObject({ cssClass: IdentifierSchema, numberFormatId: IdentifierSchema.optional() });
export const WorkbookTableStyleRegionSchema = z.enum(["wholeTable", "headerRow", "totalRow", "firstColumn", "lastColumn", "firstRowStripe", "secondRowStripe", "firstColumnStripe", "secondColumnStripe", "firstHeaderCell", "lastHeaderCell", "firstTotalCell", "lastTotalCell"]);
export const WorkbookTableStyleSchema = z.strictObject({
  styleId: IdentifierSchema, name: z.string().min(1).max(255), pivotCompatible: z.boolean(),
  elements: z.array(z.strictObject({ region: WorkbookTableStyleRegionSchema, size: z.number().int().min(1).max(9), style: WorkbookDifferentialSchema.optional() })),
});
export type WorkbookTableStyle = z.infer<typeof WorkbookTableStyleSchema>;
export const WorkbookTableStyleReferenceSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("builtin"), name: z.string().regex(/^TableStyle(?:Light(?:[1-9]|1\d|2[01])|Medium(?:[1-9]|1\d|2[0-8])|Dark(?:[1-9]|1[01]))$/) }),
  z.strictObject({ kind: z.literal("custom"), styleId: IdentifierSchema }),
]);
export type WorkbookTableStyleReference = z.infer<typeof WorkbookTableStyleReferenceSchema>;
const regionStyles = {
  headerStyle: WorkbookDifferentialSchema.optional(), dataStyle: WorkbookDifferentialSchema.optional(), totalStyle: WorkbookDifferentialSchema.optional(),
  /** Native name tokens, including unresolved aliases. Missing definitions use
   * table-style fallback (ECMA-376 §18.5.1.2/3); appearance stays in shared CSS. */
  headerCellStyleName: z.string().max(255).optional(), dataCellStyleName: z.string().max(255).optional(), totalCellStyleName: z.string().max(255).optional(),
};
export const WorkbookTableColumnSchema = z.strictObject({
  columnId: IdentifierSchema, name: z.string().min(1).max(255),
  calculatedFormula: WorkbookRuleFormulaSchema.optional(), totalFormula: WorkbookRuleFormulaSchema.optional(),
  totalFunction: z.enum(["none", "sum", "min", "max", "average", "count", "countNums", "stdDev", "var", "custom"]),
  totalLabel: z.string().max(32767).optional(), ...regionStyles,
});
export const WorkbookTableSchema = z.strictObject({
  tableId: IdentifierSchema, name: z.string().min(1).max(255), objectName: z.string().min(1).max(255).optional(),
  description: z.string().max(255).optional(),
  range: WorkbookRangeSchema, header: z.boolean(), totals: z.boolean(), totalsPreviouslyShown: z.boolean(),
  columns: z.array(WorkbookTableColumnSchema).min(1), autoFilter: WorkbookAutoFilterSchema.optional(), sort: WorkbookSortSchema.optional(),
  style: WorkbookTableStyleReferenceSchema.optional(), showFirstColumn: z.boolean(), showLastColumn: z.boolean(), rowStripes: z.boolean(), columnStripes: z.boolean(),
  ...regionStyles, borderStyle: WorkbookDifferentialSchema.optional(), headerBorderStyle: WorkbookDifferentialSchema.optional(), totalBorderStyle: WorkbookDifferentialSchema.optional(),
});
export type WorkbookTable = z.infer<typeof WorkbookTableSchema>;
