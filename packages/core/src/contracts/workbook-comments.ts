import { z } from "zod";
import { WorkbookDrawingAnchorSchema } from "./workbook-drawing.js";
import { IdentifierSchema } from "./primitives.js";

/** Native identities are opaque provenance, never authorization or credentials. */
export const WorkbookCommentAuthorSchema = z.strictObject({
  authorId: IdentifierSchema, format: z.enum(["classic", "modern"]), name: z.string().max(32767),
  nativeId: IdentifierSchema.optional(), userId: z.string().max(32767).optional(), providerId: z.string().max(32767).optional(),
});
export type WorkbookCommentAuthor = z.infer<typeof WorkbookCommentAuthorSchema>;
export const WorkbookNoteBoxSchema = z.strictObject({
  anchor: WorkbookDrawingAnchorSchema, cssClass: IdentifierSchema.optional(), visible: z.boolean(), printable: z.boolean(),
  locked: z.boolean(), lockText: z.boolean(), defaultSize: z.boolean(), disabled: z.boolean(), autoFill: z.boolean(), autoLine: z.boolean(),
  autoScale: z.boolean(), justifyLastLine: z.boolean(), altText: z.string().max(32767).optional(),
});
export type WorkbookNoteBox = z.infer<typeof WorkbookNoteBoxSchema>;
const identity = {
  annotationId: IdentifierSchema, authorId: IdentifierSchema.optional(),
  recoveryStatus: z.enum(["open", "resolved", "closed"]).optional(),
  recoveryReasons: z.array(z.enum(["missing-location", "missing-placeholder", "mismatched-placeholder", "duplicate-placeholder", "orphan-placeholder"])).min(1).max(5).optional(),
};
export const WorkbookCommentSchema = z.discriminatedUnion("format", [
  z.strictObject({ ...identity, format: z.literal("classic"), box: WorkbookNoteBoxSchema.optional(), nativeShapeId: z.number().int().min(0).max(4294967295).optional(),
    compatibilityThreadId: IdentifierSchema.optional() }),
  z.strictObject({ ...identity, format: z.literal("modern"),
    replies: z.array(z.strictObject({ replyId: IdentifierSchema, authorId: IdentifierSchema.optional(), parentReplyId: IdentifierSchema.optional() })),
  }),
]);
export type WorkbookComment = z.infer<typeof WorkbookCommentSchema>;
