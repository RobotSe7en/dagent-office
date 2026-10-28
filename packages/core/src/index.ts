export { FORMAT_VERSION, PROTOCOL_VERSION, DocumentKindSchema, RectangleSchema } from "./contracts/primitives.js";
export type { DocumentKind, Rectangle } from "./contracts/primitives.js";
export * from "./contracts/document.js";
export * from "./contracts/annotations.js";
export * from "./contracts/operations.js";
export * from "./contracts/interaction.js";
export * from "./contracts/events.js";
export { DocumentError } from "./errors.js";
export * from "./schema-registry.js";
export type {
  AnnotationSubmissionHandler, CandidateChange, ChangeMapping, DeepReadonly,
  DocumentEventListener, DocumentPackage, DocumentSelection, DocumentSession,
  DocumentSnapshot, InteractionRequestInput, TextOffsetMapping, TransactionReceipt,
} from "./runtime.js";
