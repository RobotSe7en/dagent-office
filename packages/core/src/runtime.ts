import type { Annotation, Anchor } from "./contracts/annotations.js";
import type { DocumentManifest, DocumentMetadata } from "./contracts/document.js";
import type { DocumentEvent } from "./contracts/events.js";
import type { ContextStyle, DocumentInteractionRequest, EditProposal } from "./contracts/interaction.js";
import type { ContentOperationName, DocumentTransaction } from "./contracts/operations.js";

export type DeepReadonly<Value> = Value extends readonly (infer Item)[]
  ? readonly DeepReadonly<Item>[]
  : Value extends object ? { readonly [Key in keyof Value]: DeepReadonly<Value[Key]> } : Value;

/** Virtual package files; manifest.json is contained here, never duplicated as authority. */
export type DocumentPackage = ReadonlyMap<string, Uint8Array>;

/** A derived, read-only view, not a second persistence format. */
export interface DocumentSnapshot {
  readonly manifest: DeepReadonly<DocumentManifest>;
  readonly html: string;
  readonly styles: readonly DeepReadonly<ContextStyle>[];
  readonly metadata: DeepReadonly<DocumentMetadata>;
  readonly annotations: readonly DeepReadonly<Annotation>[];
}

export interface DocumentReadLimits {
  readonly maxFiles: number;
  readonly maxPackageBytes: number;
  readonly maxFileBytes: number;
  readonly maxNodes: number;
  readonly maxDepth: number;
  readonly maxCssNodes: number;
}

export interface OpenDocumentOptions {
  readonly limits?: Partial<DocumentReadLimits>;
}

/** A derived node view. Text offsets use UTF-16; br is \n and img is U+FFFC. */
export interface DocumentNode {
  readonly nodeId: string;
  readonly tagName: string;
  readonly attributes: Readonly<Record<string, string>>;
  readonly parentNodeId: string | null;
  readonly childNodeIds: readonly string[];
  readonly text: string;
  readonly html: string;
}

/** Read-only document loading and serialization; editing belongs to DocumentSession. */
export interface DocumentModel {
  snapshot(): DocumentSnapshot;
  getNode(nodeId: string): DocumentNode | undefined;
  serialize(): DocumentPackage;
}

export interface DocumentSelection {
  readonly documentId: string;
  readonly baseRevision: number;
  readonly anchor: DeepReadonly<Anchor>;
}

export interface InteractionRequestInput {
  readonly annotationIds: readonly string[];
  readonly supportedOperations?: readonly ContentOperationName[];
  readonly limits?: {
    readonly maxBytes: number;
    readonly maxNodes: number;
    readonly maxAssets: number;
  };
}

// Type-only opacity; the future session must also bind candidates to private runtime state.
declare const candidateBrand: unique symbol;
export interface CandidateChange {
  readonly [candidateBrand]: true;
  readonly documentId: string;
  readonly baseRevision: number;
  readonly preview: DocumentSnapshot;
  readonly affectedNodeIds: readonly string[];
}

export interface TextOffsetMapping {
  readonly nodeId: string;
  readonly from: number;
  readonly to: number;
  readonly insertedLength: number;
}

export interface ChangeMapping {
  readonly nodes: ReadonlyMap<string, string | null>;
  readonly textEdits: readonly TextOffsetMapping[];
}

export interface TransactionReceipt {
  readonly documentId: string;
  readonly revision: number;
  readonly changed: boolean;
  readonly affectedNodeIds: readonly string[];
  readonly mapping: ChangeMapping;
}

export type DocumentEventListener = (event: DeepReadonly<DocumentEvent>) => void;
export type AnnotationSubmissionHandler = (request: DeepReadonly<DocumentInteractionRequest>) => void | Promise<void>;

/** Contract only. No session factory, editing, or I/O implementation is supplied in R1-03. */
export interface DocumentSession {
  snapshot(): DocumentSnapshot;
  getAnnotations(): readonly DeepReadonly<Annotation>[];
  buildInteractionRequest(input: InteractionRequestInput): DocumentInteractionRequest;
  prepare(transaction: DocumentTransaction): CandidateChange;
  prepareProposal(request: DocumentInteractionRequest, proposal: EditProposal): CandidateChange;
  commit(candidate: CandidateChange): TransactionReceipt;
  undo(): TransactionReceipt;
  redo(): TransactionReceipt;
  serialize(): DocumentPackage;
  subscribe(listener: DocumentEventListener): () => void;
  dispose(): void;
}
