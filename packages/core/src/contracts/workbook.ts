import { z } from "zod";
import { IdentifierSchema, PackagePathSchema } from "./primitives.js";

export const WorkbookRowSchema = z.number().int().min(0).max(1_048_575);
export const WorkbookColumnSchema = z.number().int().min(0).max(16_383);
/** Coordinates are zero based; inclusive range ordering is a semantic check. */
export const WorkbookCoordinateSchema = z.strictObject({ row: WorkbookRowSchema, column: WorkbookColumnSchema });
export const WorkbookRangeSchema = z.strictObject({ from: WorkbookCoordinateSchema, to: WorkbookCoordinateSchema });
export type WorkbookCoordinate = z.infer<typeof WorkbookCoordinateSchema>;
export type WorkbookRange = z.infer<typeof WorkbookRangeSchema>;

export const WorkbookChunkDescriptorSchema = z.strictObject({
  sheetId: IdentifierSchema,
  row: WorkbookRowSchema.multipleOf(256),
  column: WorkbookColumnSchema.multipleOf(32),
  path: PackagePathSchema,
  sha256: z.string().length(64).regex(/^[a-f0-9]{64}$/),
});
export type WorkbookChunkDescriptor = z.infer<typeof WorkbookChunkDescriptorSchema>;

export const WorkbookNumberFormatSchema = z.strictObject({ formatId: IdentifierSchema, code: z.string().min(1).max(2048) });
export const WorkbookRowDimensionSchema = z.strictObject({
  from: WorkbookRowSchema, to: WorkbookRowSchema,
  heightPt: z.number().min(0).max(409).optional(), hidden: z.boolean().optional(), cssClass: IdentifierSchema.optional(),
});
export const WorkbookColumnDimensionSchema = z.strictObject({
  from: WorkbookColumnSchema, to: WorkbookColumnSchema,
  width: z.number().min(0).max(255).optional(), unit: z.literal("characters"), hidden: z.boolean().optional(), cssClass: IdentifierSchema.optional(),
});
export const WorkbookSheetSchema = z.strictObject({
  sheetId: IdentifierSchema,
  name: z.string().min(1).max(31).regex(/^[^\\/\[\]:?*\u0000-\u001f]+(?![\s\S])/),
  visibility: z.enum(["visible", "hidden", "veryHidden"]),
  defaultRowHeightPt: z.number().positive().max(409),
  defaultColumnWidth: z.number().positive().max(255),
  defaultCssClass: IdentifierSchema,
  rowDimensions: z.array(WorkbookRowDimensionSchema),
  columnDimensions: z.array(WorkbookColumnDimensionSchema),
  merges: z.array(WorkbookRangeSchema),
  frozen: z.strictObject({ rows: z.number().int().min(0).max(1_048_575), columns: z.number().int().min(0).max(16_383) }),
  showGridLines: z.boolean(),
  rightToLeft: z.boolean(),
  tabColor: z.string().length(7).regex(/^#[a-fA-F0-9]{6}$/).optional(),
});
export type WorkbookSheet = z.infer<typeof WorkbookSheetSchema>;
export const WorkbookMetadataSchema = z.strictObject({
  dateSystem: z.enum(["1900", "1904"]),
  calculationMode: z.enum(["auto", "manual"]),
  activeSheetId: IdentifierSchema,
  sheets: z.array(WorkbookSheetSchema).min(1),
  numberFormats: z.array(WorkbookNumberFormatSchema),
});
export type WorkbookMetadata = z.infer<typeof WorkbookMetadataSchema>;
export const WorkbookDocumentMetadataSchema = z.strictObject({
  title: z.string().optional(), locale: IdentifierSchema.optional(), workbook: WorkbookMetadataSchema,
});
export type WorkbookDocumentMetadata = z.infer<typeof WorkbookDocumentMetadataSchema> & { readonly word?: never; readonly presentation?: never };

export const WorkbookErrorValueSchema = z.enum(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"]);
export const WorkbookValueSchema = z.discriminatedUnion("type", [
  z.strictObject({ type: z.literal("blank") }),
  z.strictObject({ type: z.literal("text"), value: z.string().max(32_767) }),
  z.strictObject({ type: z.literal("number"), value: z.number() }),
  z.strictObject({ type: z.literal("boolean"), value: z.boolean() }),
  z.strictObject({ type: z.literal("error"), value: WorkbookErrorValueSchema }),
]);
export type WorkbookValue = z.infer<typeof WorkbookValueSchema>;
export const WorkbookFormulaSourceSchema = z.strictObject({
  expression: z.string().min(2).max(8192).startsWith("="),
  cacheState: z.enum(["imported", "valid", "stale", "unavailable"]),
});
export type WorkbookFormulaSource = z.infer<typeof WorkbookFormulaSourceSchema>;
