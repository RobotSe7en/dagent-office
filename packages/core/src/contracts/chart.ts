import { z } from "zod";
import { IdentifierSchema } from "./primitives.js";

/** Shared basic chart appearance, independent of the authoritative data source. */
export const ChartAppearanceSchema = z.strictObject({
  kind: z.enum(["column", "line", "pie"]),
  titleNodeId: IdentifierSchema.optional(), titleVisible: z.boolean().optional(), legend: z.enum(["none", "left", "right", "top", "bottom"]),
  showLabels: z.boolean(), numberFormat: z.string().optional(),
  minimum: z.number().optional(), maximum: z.number().optional(), majorUnit: z.number().positive().optional(),
  plotCssClass: IdentifierSchema.optional(), legendCssClass: IdentifierSchema.optional(), textCssClass: IdentifierSchema.optional(), labelCssClass: IdentifierSchema.optional(),
  categoryAxisCssClass: IdentifierSchema.optional(), valueAxisCssClass: IdentifierSchema.optional(), gridCssClass: IdentifierSchema.optional(),
  showCategoryAxis: z.boolean().optional(), showValueAxis: z.boolean().optional(),
  blankAs: z.enum(["gap", "zero", "span"]).optional(),
  reverseCategories: z.boolean().optional(), reverseValues: z.boolean().optional(),
  gapWidth: z.number().min(0).max(500).optional(), overlap: z.number().min(-100).max(100).optional(), firstSliceAngle: z.number().min(0).max(360).optional(),
  labelContents: z.array(z.enum(["value", "category", "series", "percent"])).optional(),
  seriesStyles: z.array(z.strictObject({ seriesId: IdentifierSchema, cssClass: IdentifierSchema,
    labelCssClass: IdentifierSchema.optional(),
    marker: z.enum(["none", "circle", "square", "diamond", "triangle", "x", "star", "plus", "dot", "dash"]).optional(), markerSize: z.number().positive().optional() })).optional(),
  pointStyles: z.array(z.strictObject({ seriesId: IdentifierSchema, index: z.number().int().min(0), cssClass: IdentifierSchema, labelCssClass: IdentifierSchema.optional() })).optional(),
});
export type ChartAppearance = z.infer<typeof ChartAppearanceSchema>;
