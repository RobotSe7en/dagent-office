import { z } from "zod";
import { IdentifierSchema } from "./primitives.js";

/** Native identities are opaque provenance, never authorization or credentials. */
export const WorkbookCommentAuthorSchema = z.strictObject({
  authorId: IdentifierSchema, format: z.enum(["classic", "modern"]), name: z.string().max(32767),
  nativeId: IdentifierSchema.optional(), userId: z.string().max(32767).optional(), providerId: z.string().max(32767).optional(),
});
export type WorkbookCommentAuthor = z.infer<typeof WorkbookCommentAuthorSchema>;
const identity = {
  annotationId: IdentifierSchema, authorId: IdentifierSchema.optional(),
  recoveryStatus: z.enum(["open", "resolved", "closed"]).optional(),
};
export const WorkbookCommentSchema = z.discriminatedUnion("format", [
  z.strictObject({ ...identity, format: z.literal("classic") }),
  z.strictObject({ ...identity, format: z.literal("modern"),
    replies: z.array(z.strictObject({ replyId: IdentifierSchema, authorId: IdentifierSchema.optional(), parentReplyId: IdentifierSchema.optional() })),
  }),
]);
export type WorkbookComment = z.infer<typeof WorkbookCommentSchema>;
