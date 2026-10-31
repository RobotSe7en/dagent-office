import { z } from "zod";
import { AnchorSchema, AnnotationSchema, AnnotationStatusSchema } from "./annotations.js";
import { IdentifierSchema, PackagePathSchema, PROTOCOL_VERSION, RevisionSchema } from "./primitives.js";
import { PresentationGeometrySchema, PresentationSizeSchema } from "./presentation.js";

export const ReconciledTextEditSchema = z.strictObject({
  nodeId: IdentifierSchema, from: RevisionSchema, to: RevisionSchema,
  expectedText: z.string(), text: z.string(),
});
export type ReconciledTextEdit = z.infer<typeof ReconciledTextEditSchema>;

export const DocumentContentOperationSchema = z.discriminatedUnion("op", [
  z.strictObject({
    op: z.literal("replaceText"),
    nodeId: IdentifierSchema,
    from: RevisionSchema,
    to: RevisionSchema,
    expectedText: z.string(),
    text: z.string(),
  }),
  z.strictObject({
    op: z.literal("replaceFragment"),
    nodeId: IdentifierSchema,
    html: z.string(),
  }),
  z.strictObject({
    op: z.literal("reconcileFragment"),
    nodeId: IdentifierSchema,
    html: z.string(),
    textEdits: z.array(ReconciledTextEditSchema),
  }),
  z.strictObject({
    op: z.literal("insertFragment"),
    nodeId: IdentifierSchema,
    position: z.enum(["before", "after", "start", "end"]),
    html: z.string().min(1),
  }),
  z.strictObject({ op: z.literal("removeNode"), nodeId: IdentifierSchema }),
  z.strictObject({
    op: z.literal("setStyle"),
    nodeId: IdentifierSchema,
    property: z.string().min(1),
    value: z.string().min(1),
  }),
  z.strictObject({
    op: z.literal("replaceAsset"),
    nodeId: IdentifierSchema,
    assetPath: PackagePathSchema,
  }),
]);
export type DocumentContentOperation = z.infer<typeof DocumentContentOperationSchema>;

export const ContentOperationNameSchema = z.enum(
  DocumentContentOperationSchema.options.map((schema) => schema.shape.op.value),
);
export type ContentOperationName = z.infer<typeof ContentOperationNameSchema>;

export const AnnotationOperationSchema = z.discriminatedUnion("op", [
  z.strictObject({ op: z.literal("addAnnotation"), annotation: AnnotationSchema }),
  z.strictObject({
    op: z.literal("updateAnnotation"),
    annotationId: IdentifierSchema,
    changes: z.strictObject({
      anchor: AnchorSchema.optional(),
      instruction: z.string().min(1).optional(),
      status: AnnotationStatusSchema.optional(),
    }),
  }),
  z.strictObject({ op: z.literal("removeAnnotation"), annotationId: IdentifierSchema }),
]);
export type AnnotationOperation = z.infer<typeof AnnotationOperationSchema>;

/** Format operations remain separate from the V1 Word proposal operation set. */
export const PresentationOperationSchema = z.discriminatedUnion("op", [
  z.strictObject({ op: z.literal("setPresentationGeometry"), nodeId: IdentifierSchema, geometry: PresentationGeometrySchema,
    childCoordinates: z.strictObject({ x: z.number(), y: z.number(), width: z.number().positive(), height: z.number().positive() }).optional() }),
  z.strictObject({ op: z.literal("setPresentationSize"), size: PresentationSizeSchema }),
]);
export type PresentationOperation = z.infer<typeof PresentationOperationSchema>;

export const DocumentOperationSchema = z.discriminatedUnion("op", [
  ...DocumentContentOperationSchema.options,
  ...AnnotationOperationSchema.options,
  ...PresentationOperationSchema.options,
]);
export type DocumentOperation = z.infer<typeof DocumentOperationSchema>;

export const DocumentTransactionSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION),
  documentId: IdentifierSchema,
  baseRevision: RevisionSchema,
  operations: z.array(DocumentOperationSchema),
});
export type DocumentTransaction = z.infer<typeof DocumentTransactionSchema>;
