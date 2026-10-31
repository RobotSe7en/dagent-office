import { z } from "zod";
import { AnnotationSchema } from "./annotations.js";
import { AssetDescriptorSchema, WordDocumentMetadataSchema } from "./document.js";
import { ContentOperationNameSchema, DocumentContentOperationSchema } from "./operations.js";
import { FORMAT_VERSION, IdentifierSchema, PackagePathSchema, PROTOCOL_VERSION, RevisionSchema } from "./primitives.js";

export const ContextBlockSchema = z.strictObject({
  nodeId: IdentifierSchema,
  html: z.string().min(1),
});
export type ContextBlock = z.infer<typeof ContextBlockSchema>;

export const ContextStyleSchema = z.strictObject({
  path: PackagePathSchema,
  css: z.string(),
});
export type ContextStyle = z.infer<typeof ContextStyleSchema>;

export const InteractionContextSchema = z.strictObject({
  formatVersion: z.literal(FORMAT_VERSION),
  kind: z.literal("document"),
  blocks: z.array(ContextBlockSchema).min(1),
  styles: z.array(ContextStyleSchema),
  metadata: WordDocumentMetadataSchema,
  assets: z.array(AssetDescriptorSchema),
});
export type InteractionContext = z.infer<typeof InteractionContextSchema>;

export const DocumentInteractionRequestSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  requestId: IdentifierSchema,
  documentId: IdentifierSchema,
  baseRevision: RevisionSchema,
  annotations: z.array(z.union([
    AnnotationSchema.options[0].extend({ status: z.literal("open") }),
    AnnotationSchema.options[1].extend({ status: z.literal("open") }),
  ])),
  context: InteractionContextSchema,
  supportedOperations: z.array(ContentOperationNameSchema),
});
export type DocumentInteractionRequest = z.infer<typeof DocumentInteractionRequestSchema>;

export const EditProposalSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  requestId: IdentifierSchema,
  proposalId: IdentifierSchema,
  documentId: IdentifierSchema,
  baseRevision: RevisionSchema,
  annotationIds: z.array(IdentifierSchema),
  summary: z.string(),
  operations: z.array(DocumentContentOperationSchema),
});
export type EditProposal = z.infer<typeof EditProposalSchema>;
