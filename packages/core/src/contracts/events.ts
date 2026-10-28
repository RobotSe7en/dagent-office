import { z } from "zod";
import { IdentifierSchema, PROTOCOL_VERSION, RevisionSchema } from "./primitives.js";

export const DocumentErrorCodeSchema = z.enum([
  "UnsupportedFormat", "UnsupportedVersion", "UnsupportedFeature", "InvalidDocument",
  "InvalidOperation", "InvalidProposal", "AnchorNotFound", "RevisionConflict",
  "AssetUnavailable", "ResourceLimitExceeded", "HostSubmissionFailed", "ExportFailed", "HostCommitFailed",
]);
export type DocumentErrorCode = z.infer<typeof DocumentErrorCodeSchema>;

export const ValidationIssueSchema = z.strictObject({
  code: z.string().min(1).max(64),
  path: z.array(z.union([z.string().max(128), RevisionSchema])).max(16),
  message: z.string().min(1).max(256),
});
export type ValidationIssue = z.infer<typeof ValidationIssueSchema>;

export const DocumentErrorDataSchema = z.strictObject({
  code: DocumentErrorCodeSchema,
  message: z.string().min(1).max(256),
  details: z.array(ValidationIssueSchema).max(16),
});
export type DocumentErrorData = z.infer<typeof DocumentErrorDataSchema>;

const eventIdentity = {
  protocolVersion: z.literal(PROTOCOL_VERSION),
  documentId: IdentifierSchema,
  revision: RevisionSchema,
};
export const DocumentEventSchema = z.discriminatedUnion("type", [
  z.strictObject({
    ...eventIdentity,
    type: z.literal("documentChanged"),
    affectedNodeIds: z.array(IdentifierSchema),
  }),
  z.strictObject({
    ...eventIdentity,
    type: z.literal("annotationChanged"),
    annotationIds: z.array(IdentifierSchema),
  }),
  z.strictObject({
    ...eventIdentity,
    type: z.literal("error"),
    error: DocumentErrorDataSchema,
  }),
]);
export type DocumentEvent = z.infer<typeof DocumentEventSchema>;
