import { z } from "zod";
import { AnnotationSchema } from "./annotations.js";
import { AssetDescriptorSchema } from "./document.js";
import { PresentationDocumentMetadataSchema } from "./presentation.js";
import { DocumentContentOperationSchema } from "./operations.js";
import { PresentationOperationSchema } from "./presentation-operations.js";
import { ContextStyleSchema } from "./interaction.js";
import { FORMAT_VERSION, IdentifierSchema, PROTOCOL_VERSION, RevisionSchema } from "./primitives.js";
export const PresentationProposalOperationSchema = z.discriminatedUnion("op", [...DocumentContentOperationSchema.options, ...PresentationOperationSchema.options]);
export const PresentationOperationNameSchema = z.enum(PresentationProposalOperationSchema.options.map(schema => schema.shape.op.value));
export type PresentationOperationName = z.infer<typeof PresentationOperationNameSchema>;
export const PresentationInteractionContextSchema = z.strictObject({
    formatVersion: z.literal(FORMAT_VERSION), kind: z.literal("presentation"),
    scope: z.strictObject({ kind: z.enum(["slides", "global"]), slideIds: z.array(IdentifierSchema).min(1) }),
    blocks: z.array(z.strictObject({ nodeId: IdentifierSchema, slideId: IdentifierSchema, html: z.string().min(1) })).min(1),
    dependencies: z.array(z.union([
        z.strictObject({ kind: z.literal("theme"), definitionId: IdentifierSchema, access: z.enum(["read-only", "editable"]) }),
        z.strictObject({ kind: z.enum(["master", "layout", "notes-master", "notes", "comment"]), definitionId: IdentifierSchema, nodeId: IdentifierSchema, html: z.string().min(1), access: z.enum(["read-only", "editable"]) }),
    ])),
    styles: z.array(ContextStyleSchema), metadata: PresentationDocumentMetadataSchema, assets: z.array(AssetDescriptorSchema), providedAssets: z.array(AssetDescriptorSchema),
});
export type PresentationInteractionContext = z.infer<typeof PresentationInteractionContextSchema>;
export const PresentationInteractionRequestSchema = z.strictObject({
    protocolVersion: z.literal(PROTOCOL_VERSION), requestId: IdentifierSchema, documentId: IdentifierSchema, baseRevision: RevisionSchema,
    annotations: z.array(z.union([AnnotationSchema.options[0].extend({ status: z.literal("open") }), AnnotationSchema.options[1].extend({ status: z.literal("open") })])).min(1),
    context: PresentationInteractionContextSchema, supportedOperations: z.array(PresentationOperationNameSchema),
});
export type PresentationInteractionRequest = z.infer<typeof PresentationInteractionRequestSchema>;
export const PresentationEditProposalSchema = z.strictObject({
    protocolVersion: z.literal(PROTOCOL_VERSION), requestId: IdentifierSchema, proposalId: IdentifierSchema, documentId: IdentifierSchema, baseRevision: RevisionSchema,
    annotationIds: z.array(IdentifierSchema), summary: z.string(), operations: z.array(PresentationProposalOperationSchema),
});
export type PresentationEditProposal = z.infer<typeof PresentationEditProposalSchema>;
