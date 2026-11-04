import { z } from "zod";
import { WorkbookColumnSchema, WorkbookRowSchema } from "./workbook-coordinates.js";
import { IdentifierSchema } from "./primitives.js";
import { DrawingPresetSchema, DrawingTextPropertiesSchema } from "./drawing.js";

// DrawingML's ST_Coordinate bound, converted from EMU to document points.
const point = z.number().min(-27273042329600 / 12700).max(27273042316900 / 12700);
const extent = z.number().nonnegative().max(27273042316900 / 12700);
export const WorkbookDrawingMarkerSchema = z.strictObject({
  row: WorkbookRowSchema, column: WorkbookColumnSchema, offsetXPt: point, offsetYPt: point,
});
export type WorkbookDrawingMarker = z.infer<typeof WorkbookDrawingMarkerSchema>;
export const WorkbookDrawingAnchorSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("absolute"), xPt: point, yPt: point, widthPt: extent, heightPt: extent }),
  z.strictObject({ kind: z.literal("one-cell"), from: WorkbookDrawingMarkerSchema, widthPt: extent, heightPt: extent }),
  z.strictObject({ kind: z.literal("two-cell"), from: WorkbookDrawingMarkerSchema, to: WorkbookDrawingMarkerSchema,
    moveWithCells: z.boolean(), sizeWithCells: z.boolean(),
  }),
]);
export type WorkbookDrawingAnchor = z.infer<typeof WorkbookDrawingAnchorSchema>;

export const WorkbookPresetSchema = DrawingPresetSchema;
export const WorkbookDrawingLocksSchema = z.strictObject({
  noGrp: z.boolean().optional(), noSelect: z.boolean().optional(), noRot: z.boolean().optional(), noChangeAspect: z.boolean().optional(),
  noMove: z.boolean().optional(), noResize: z.boolean().optional(), noEditPoints: z.boolean().optional(), noAdjustHandles: z.boolean().optional(),
  noChangeArrowheads: z.boolean().optional(), noChangeShapeType: z.boolean().optional(), noTextEdit: z.boolean().optional(), noCrop: z.boolean().optional(),
});
const object = {
  nodeId: IdentifierSchema, sheetId: IdentifierSchema, anchor: WorkbookDrawingAnchorSchema,
  name: z.string().max(32767).optional(), nativeId: IdentifierSchema.optional(),
  locked: z.boolean(), printable: z.boolean(),
  locks: WorkbookDrawingLocksSchema.optional(),
};
/** Anchor behavior lives in metadata; appearance, rotation/flips, crop and rich
 * text remain in the referenced HTML/CSS source, as for presentation objects. */
export const WorkbookObjectSchema = z.discriminatedUnion("kind", [
  z.strictObject({ ...object, kind: z.literal("text"), text: DrawingTextPropertiesSchema.optional() }),
  z.strictObject({ ...object, kind: z.literal("geometry"), preset: DrawingPresetSchema, text: DrawingTextPropertiesSchema.optional() }),
  z.strictObject({ ...object, kind: z.literal("image"), preferRelativeResize: z.boolean().optional() }),
  z.strictObject({ ...object, kind: z.literal("chart"), chartId: IdentifierSchema }),
]);
export type WorkbookObject = z.infer<typeof WorkbookObjectSchema>;
