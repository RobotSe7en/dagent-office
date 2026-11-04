import { z } from "zod";
import { WorkbookColumnSchema, WorkbookRowSchema } from "./workbook-coordinates.js";

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
