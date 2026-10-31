export { FORMAT_VERSION, PROTOCOL_VERSION, DocumentKindSchema, PackagePathSchema, RectangleSchema } from "./contracts/primitives.js";
export type { DocumentKind, Rectangle } from "./contracts/primitives.js";
export * from "./contracts/document.js";
export * from "./contracts/word.js";
export * from "./contracts/presentation.js";
export * from "./contracts/annotations.js";
export * from "./contracts/operations.js";
export * from "./contracts/interaction.js";
export * from "./contracts/events.js";
export { DocumentError } from "./errors.js";
export { createOpaqueId } from "./identifiers.js";
export * from "./schema-registry.js";
export { DEFAULT_DOCUMENT_READ_LIMITS, openDocument } from "./document-model.js";
export { createPresentation } from "./presentation.js";
export type { CreatePresentationOptions } from "./presentation.js";
export { readPresentationGeometry } from "./document/presentation.js";
export { createDocumentSession, DEFAULT_DOCUMENT_SESSION_LIMITS, DEFAULT_INTERACTION_LIMITS } from "./document-session.js";
export { collectDocumentAssets, createAssetDescriptor } from "./assets.js";
export type {
  AnnotationSubmissionHandler, AssetProvider, CandidateChange, ChangeMapping, DeepReadonly,
  DocumentEventListener, DocumentPackage, DocumentSelection, DocumentSession,
  DocumentModel, DocumentNode, DocumentReadLimits, DocumentSnapshot, InteractionRequestInput,
  OpenDocumentOptions, TextOffsetMapping, TransactionReceipt, DocumentEditingSession,
  DocumentNodeChange, DocumentSessionLimits, DocumentSessionOptions, DocumentHistoryState,
} from "./runtime.js";
