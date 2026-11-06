import { z } from "zod";
import { IdentifierSchema } from "./primitives.js";

/** These native update flags govern propagation of a later named-style edit.
 * Current appearance remains the cell/row/column CSS and number-format source. */
export const WorkbookStyleApplicationSchema = z.strictObject({
  font: z.boolean(), fill: z.boolean(), border: z.boolean(), alignment: z.boolean(), numberFormat: z.boolean(), protection: z.boolean(),
});
export type WorkbookStyleApplication = z.infer<typeof WorkbookStyleApplicationSchema>;
export const WorkbookCellStylesSchema = z.strictObject({
  definitions: z.array(z.strictObject({
    styleId: IdentifierSchema, cssClass: IdentifierSchema, numberFormatId: IdentifierSchema, apply: WorkbookStyleApplicationSchema,
  })),
  names: z.array(z.strictObject({
    name: z.string().min(1).max(255), styleId: IdentifierSchema, hidden: z.boolean(), customBuiltin: z.boolean(),
    builtinId: z.number().int().min(0).max(0xffff_ffff).optional(), level: z.number().int().min(0).max(0xffff_ffff).optional(),
  })),
  /** A selected cell, row or column class binds to one style definition.
   * This records identity/update behavior, never a second effective appearance. */
  bindings: z.array(z.strictObject({ cssClass: IdentifierSchema, styleId: IdentifierSchema, apply: WorkbookStyleApplicationSchema })),
});
export type WorkbookCellStyles = z.infer<typeof WorkbookCellStylesSchema>;
