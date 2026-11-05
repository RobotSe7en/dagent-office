import { z } from "zod";
import { IdentifierSchema, RectangleSchema, RevisionSchema } from "./primitives.js";
import { WorkbookRangeSchema } from "./workbook-coordinates.js";

export const TextRangeSchema = z.strictObject({
  nodeId: IdentifierSchema,
  from: RevisionSchema,
  to: RevisionSchema,
  quote: z.string(),
});
export type TextRange = z.infer<typeof TextRangeSchema>;

export const AnchorSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("cells"), sheetId: IdentifierSchema, ranges: z.array(WorkbookRangeSchema).min(1) }),
  TextRangeSchema.extend({ kind: z.literal("text") }),
  z.strictObject({
    kind: z.literal("text-set"),
    ranges: z.array(TextRangeSchema).min(1),
  }),
  z.strictObject({ kind: z.literal("node"), nodeId: IdentifierSchema }),
  z.strictObject({
    kind: z.literal("node-set"),
    nodeIds: z.array(IdentifierSchema).min(1),
  }),
  z.strictObject({
    kind: z.literal("region"),
    surfaceId: IdentifierSchema,
    rect: RectangleSchema,
    nodeIds: z.array(IdentifierSchema),
    layoutKey: IdentifierSchema,
    pageId: IdentifierSchema.optional(),
  }),
]);
export type Anchor = z.infer<typeof AnchorSchema>;

export const AnnotationStatusSchema = z.enum(["open", "resolved", "closed", "needs_reanchor"]);
export type AnnotationStatus = z.infer<typeof AnnotationStatusSchema>;

const author = z.strictObject({ name: z.string(), initials: z.string().optional() });
const replyIdentity = {
  replyId: IdentifierSchema, author: author.optional(), date: z.string().optional(), sourceId: IdentifierSchema.optional(),
  status: z.enum(["open", "resolved", "closed"]).optional(),
};
export const AnnotationReplySchema = z.union([
  z.strictObject({ ...replyIdentity, instruction: z.string().min(1), bodyNodeId: z.never().optional() }),
  z.strictObject({ ...replyIdentity, bodyNodeId: IdentifierSchema, instruction: z.never().optional() }),
]);
export type AnnotationReply = z.infer<typeof AnnotationReplySchema>;
const annotationIdentity = {
  annotationId: IdentifierSchema,
  documentId: IdentifierSchema,
  baseRevision: RevisionSchema,
  status: AnnotationStatusSchema,
  anchor: AnchorSchema,
  author: author.optional(),
  date: z.string().optional(),
  sourceId: IdentifierSchema.optional(),
  replies: z.array(AnnotationReplySchema).optional(),
};
const instructionBody = { instruction: z.string().min(1), bodyNodeId: z.never().optional() };
const storedBody = { bodyNodeId: IdentifierSchema, instruction: z.never().optional() };
// A native thread can have no recorded cell. Represent that absence only in
// an explicitly invalidated annotation; normal anchors/requests still require
// at least one range. Never fabricate an A1 location for an unlocated thread.
const unlocatedCells = { ...annotationIdentity, status: z.literal("needs_reanchor"),
  // Keep the kind dispatch even for this one permitted variant. Bounded array
  // preflight must not interpret a Word/PPT text-set's ranges as coordinates
  // while examining this alternative of the enclosing annotation union.
  anchor: z.discriminatedUnion("kind", [z.strictObject({ kind: z.literal("cells"), sheetId: IdentifierSchema, ranges: z.array(WorkbookRangeSchema).length(0) })]) };
export const AnnotationSchema = z.union([
  z.strictObject({ ...annotationIdentity, ...instructionBody }), z.strictObject({ ...annotationIdentity, ...storedBody }),
  z.strictObject({ ...unlocatedCells, ...instructionBody }), z.strictObject({ ...unlocatedCells, ...storedBody }),
]);
export type Annotation = z.infer<typeof AnnotationSchema>;
export const AnnotationsSchema = z.array(AnnotationSchema);
