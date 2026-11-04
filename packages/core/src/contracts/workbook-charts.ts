import { z } from "zod";
import { IdentifierSchema } from "./primitives.js";
import { ChartAppearanceSchema } from "./chart.js";

const index = z.number().int().min(0).max(0xffff_ffff);
export const WorkbookChartCacheSchema = z.strictObject({
  state: z.enum(["imported", "valid", "stale"]), count: index,
  points: z.array(z.strictObject({ index, value: z.union([z.number(), z.string().max(32767), z.null()]), numberFormat: z.string().max(2048).optional() })),
  numberFormat: z.string().max(2048).optional(),
});
export type WorkbookChartCache = z.infer<typeof WorkbookChartCacheSchema>;
/** A reference is the source. Native caches are explicitly derived values and
 * never become an independently writable chart data table. Origin is A1 on the
 * owning worksheet, including for local defined-name lookup. */
export const WorkbookChartReferenceSchema = z.strictObject({
  expression: z.string().min(2).max(8192).startsWith("="), valueType: z.enum(["text", "number"]), cache: WorkbookChartCacheSchema.optional(),
});
export type WorkbookChartReference = z.infer<typeof WorkbookChartReferenceSchema>;
export const WorkbookChartLabelSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("text"), text: z.string().max(32767) }),
  z.strictObject({ kind: z.literal("reference"), reference: WorkbookChartReferenceSchema }),
]);
export type WorkbookChartLabel = z.infer<typeof WorkbookChartLabelSchema>;
export const WorkbookChartSchema = ChartAppearanceSchema.extend({
  chartId: IdentifierSchema, nodeId: IdentifierSchema,
  categories: WorkbookChartReferenceSchema.optional(),
  series: z.array(z.strictObject({ seriesId: IdentifierSchema, nativeIndex: index.optional(), name: WorkbookChartLabelSchema.optional(), values: WorkbookChartReferenceSchema })).min(1),
  titleReference: WorkbookChartReferenceSchema.optional(),
  visibleCellsOnly: z.boolean(), varyColors: z.boolean(), showLabelsOverMaximum: z.boolean(),
  language: z.string().max(255).optional(), dateSystem: z.enum(["1900", "1904"]).optional(),
});
export type WorkbookChart = z.infer<typeof WorkbookChartSchema>;
