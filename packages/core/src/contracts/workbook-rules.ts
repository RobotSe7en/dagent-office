import { z } from "zod";
import { IdentifierSchema } from "./primitives.js";
import { WorkbookCoordinateSchema, WorkbookRangeSchema } from "./workbook-coordinates.js";

export const WorkbookRuleFormulaSchema = z.string().min(2).max(8192).startsWith("=");
export const WorkbookComparisonSchema = z.enum(["between", "notBetween", "equal", "notEqual", "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual"]);
const region = { ruleId: IdentifierSchema, ranges: z.array(WorkbookRangeSchema).min(1), origin: WorkbookCoordinateSchema };
const differential = { cssClass: IdentifierSchema.optional(), numberFormatId: IdentifierSchema.optional() };
const conditional = { ...region, priority: z.number().int().min(1).max(2147483647), stopIfTrue: z.boolean() };
export const WorkbookColorThresholdSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.enum(["min", "max"]), cssClass: IdentifierSchema }),
  z.strictObject({ kind: z.literal("number"), value: z.number(), cssClass: IdentifierSchema }),
  z.strictObject({ kind: z.enum(["percent", "percentile"]), value: z.number().min(0).max(100), cssClass: IdentifierSchema }),
  z.strictObject({ kind: z.literal("formula"), expression: WorkbookRuleFormulaSchema, cssClass: IdentifierSchema }),
]);
export const WorkbookConditionalRuleSchema = z.discriminatedUnion("kind", [
  z.strictObject({ ...conditional, ...differential, kind: z.literal("cellIs"), operator: WorkbookComparisonSchema, formulas: z.array(WorkbookRuleFormulaSchema).min(1).max(2) }),
  z.strictObject({ ...conditional, ...differential, kind: z.literal("expression"), expression: WorkbookRuleFormulaSchema }),
  z.strictObject({ ...conditional, kind: z.literal("colorScale"), stops: z.array(WorkbookColorThresholdSchema).min(2).max(3) }),
]);
export type WorkbookConditionalRule = z.infer<typeof WorkbookConditionalRuleSchema>;
export type WorkbookColorThreshold = z.infer<typeof WorkbookColorThresholdSchema>;

const validation = { ...region, allowBlank: z.boolean(), showErrorMessage: z.boolean(), showInputMessage: z.boolean(),
  errorStyle: z.enum(["stop", "warning", "information"]), errorTitle: z.string().max(32), error: z.string().max(225), promptTitle: z.string().max(32), prompt: z.string().max(255) };
export const WorkbookDataValidationSchema = z.discriminatedUnion("kind", [
  z.strictObject({ ...validation, kind: z.literal("none") }),
  z.strictObject({ ...validation, kind: z.literal("list"), expression: WorkbookRuleFormulaSchema, dropdown: z.boolean() }),
  z.strictObject({ ...validation, kind: z.literal("custom"), expression: WorkbookRuleFormulaSchema }),
  z.strictObject({ ...validation, kind: z.enum(["whole", "decimal", "date", "time", "textLength"]), operator: WorkbookComparisonSchema, formulas: z.array(WorkbookRuleFormulaSchema).min(1).max(2) }),
]);
export type WorkbookDataValidation = z.infer<typeof WorkbookDataValidationSchema>;
