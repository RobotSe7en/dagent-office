import { z } from "zod";
export const WorkbookRowSchema = z.number().int().min(0).max(1_048_575);
export const WorkbookColumnSchema = z.number().int().min(0).max(16_383);
/** Coordinates are zero based; inclusive range ordering is a semantic check. */
export const WorkbookCoordinateSchema = z.strictObject({ row: WorkbookRowSchema, column: WorkbookColumnSchema });
export const WorkbookRangeSchema = z.strictObject({ from: WorkbookCoordinateSchema, to: WorkbookCoordinateSchema });
export type WorkbookCoordinate = z.infer<typeof WorkbookCoordinateSchema>;
export type WorkbookRange = z.infer<typeof WorkbookRangeSchema>;
