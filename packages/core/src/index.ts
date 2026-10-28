export { FORMAT_VERSION, PROTOCOL_VERSION, DocumentKindSchema, PackagePathSchema, RectangleSchema } from "./contracts/primitives.js";
export type { DocumentKind, Rectangle } from "./contracts/primitives.js";
export * from "./contracts/document.js";
export * from "./contracts/annotations.js";
export * from "./contracts/operations.js";
export * from "./contracts/interaction.js";
export * from "./contracts/events.js";
export { DocumentError } from "./errors.js";
export * from "./schema-registry.js";
export { DEFAULT_DOCUMENT_READ_LIMITS, openDocument } from "./document-model.js";
export { collectDocumentAssets, createAssetDescriptor } from "./assets.js";
export type {
  AnnotationSubmissionHandler, AssetProvider, CandidateChange, ChangeMapping, DeepReadonly,
  DocumentEventListener, DocumentPackage, DocumentSelection, DocumentSession,
  DocumentModel, DocumentNode, DocumentReadLimits, DocumentSnapshot, InteractionRequestInput,
  OpenDocumentOptions, TextOffsetMapping, TransactionReceipt,
} from "./runtime.js";
