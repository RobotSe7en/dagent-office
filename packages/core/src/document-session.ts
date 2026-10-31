import type { Annotation } from "./contracts/annotations.js";
import type { DocumentMetadata } from "./contracts/document.js";
import type { DocumentEvent } from "./contracts/events.js";
import type { DocumentOperation, DocumentTransaction } from "./contracts/operations.js";
import type { CandidateChange, DocumentSession, DocumentEventListener, DocumentModel, DocumentSessionOptions, TextOffsetMapping, TransactionReceipt } from "./runtime.js";
import { DocumentError } from "./errors.js";
import { parseContract } from "./schema-registry.js";
import { DocumentEdit } from "./document/edit.js";
import { parseHtml } from "./document/html.js";
import type { ParsedHtml } from "./document/html.js";
import { copyPackage, encode } from "./document/package.js";
import { buildDocumentState } from "./document/state.js";
import type { DocumentState } from "./document/state.js";
import { checkLimit, freeze, readLimits } from "./document/validation.js";
import { compareStates, retainedBytes } from "./session/changes.js";
import { checkTransactionSize, sessionLimits } from "./session/limits.js";
import { annotationNodes, mapAnnotations } from "./session/anchors.js";
import { buildInteractionRequest, validateProposal } from "./session/interaction.js";
import type { DocumentInteractionRequest, EditProposal } from "./contracts/interaction.js";
import type { InteractionRequestInput } from "./runtime.js";
export { DEFAULT_INTERACTION_LIMITS } from "./session/interaction.js";

export { DEFAULT_DOCUMENT_SESSION_LIMITS } from "./session/limits.js";

function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }

/** Remove definitions only when their complete source object was deleted. Partial damage fails validation. */
function updateWordMetadata(metadata: DocumentMetadata, before: DocumentState, after: ParsedHtml): void {
  const word = metadata.word;
  if (!word) return;
  word.hyperlinks = word.hyperlinks.filter((item) => after.hasNode(item.nodeId));
  word.revisions = word.revisions.filter((item) => after.hasNode(item.nodeId));
  word.tables = word.tables.filter((item) => after.hasNode(item.nodeId));
  word.tabs = word.tabs.filter((item) => after.hasNode(item.paragraphNodeId));
  word.bookmarks = word.bookmarks.filter((item) => after.hasNode(item.startNodeId) || after.hasNode(item.endNodeId));
  word.fields = word.fields.filter((item) => [item.nodeId, ...(item.endNodeId ? [item.endNodeId] : []), ...(item.resultNodeIds ?? [])]
    .some((id) => after.hasNode(id)));
  for (const field of word.fields) {
    const ids = [field.nodeId, ...(field.resultNodeIds ?? [])];
    if (ids.some((id) => before.html.getNode(id)?.text !== after.getNode(id)?.text)) field.dirty = true;
  }
}

function rebase(state: DocumentState, revision: number, limits: ReturnType<typeof readLimits>): DocumentState {
  const files = new Map(state.files);
  const manifest = { ...state.snapshot.manifest, revision };
  const annotations = state.snapshot.annotations.map((annotation) => annotation.status === "needs_reanchor"
    ? annotation : { ...annotation, baseRevision: revision });
  files.set("manifest.json", encode(JSON.stringify(manifest)));
  files.set(manifest.annotations, encode(JSON.stringify(annotations)));
  return buildDocumentState(files, limits);
}

interface Prepared {
  readonly baseRevision: number;
  consumed: boolean;
}
interface HistoryEntry {
  readonly before: DocumentState;
  readonly after: DocumentState;
  readonly textEdits: readonly TextOffsetMapping[];
}

/** Create an isolated transaction session from an authorized, validated document source. */
export function createDocumentSession(document: DocumentModel, options: DocumentSessionOptions = {}): DocumentSession {
  const limits = readLimits(options.limits);
  const budgets = sessionLimits(options.sessionLimits);
  let current: DocumentState | undefined = buildDocumentState(copyPackage(document.serialize(), limits), limits);
  if (current.snapshot.metadata.presentation !== undefined) throw new DocumentError("UnsupportedFeature", "Presentation editing is not implemented yet.");
  let undo: HistoryEntry[] = [];
  let redo: HistoryEntry[] = [];
  const activeCandidates = new Map<CandidateChange, DocumentState>();
  let issued = new WeakMap<CandidateChange, Prepared>();
  const listeners = new Set<DocumentEventListener>();
  const reservedIds = new Set(current.html.nodeIds());
  let sequence = 0;
  let publishing = false;

  function state(): DocumentState {
    if (!current) invalid("This document session has been disposed.");
    return current;
  }
  function editable(): DocumentState {
    const result = state();
    if (publishing) invalid("A session cannot be changed while its events are being delivered.");
    return result;
  }
  function emit(event: DocumentEvent): void {
    publishing = true;
    try {
      const frozen = freeze(event);
      for (const listener of [...listeners]) {
        // A consumer callback cannot turn an already committed edit into an apparent failure.
        try { listener(frozen); } catch { /* Consumer owns its callback errors. */ }
      }
    } finally { publishing = false; }
  }
  function execute<Value>(action: () => Value): Value {
    editable();
    try { return action(); }
    catch (error) {
      const bounded = error instanceof DocumentError
        ? error.code === "InvalidDocument" ? new DocumentError("InvalidOperation", error.message) : error
        : new DocumentError("InvalidOperation", "The document operation could not be prepared or committed.");
      const snapshot = state().snapshot;
      emit({ protocolVersion: 1, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, type: "error", error: bounded.toJSON() });
      throw bounded;
    }
  }
  function allocate(): string {
    let id: string;
    do {
      if (sequence >= Number.MAX_SAFE_INTEGER) throw new DocumentError("ResourceLimitExceeded", "Node identity allocation is exhausted.");
      id = "edit-" + ++sequence;
    } while (reservedIds.has(id));
    return id;
  }
  function historyStates(entries: readonly HistoryEntry[]): DocumentState[] {
    return entries.flatMap((entry) => [entry.before, entry.after]);
  }
  function emptyReceipt(): TransactionReceipt {
    const base = state();
    return freeze({ documentId: base.snapshot.manifest.documentId, revision: base.snapshot.manifest.revision, changed: false,
      affectedNodeIds: [], mapping: compareStates(base, base).mapping });
  }
  function publish(next: DocumentState, textEdits: readonly TextOffsetMapping[]): TransactionReceipt {
    const previous = state();
    const diff = compareStates(previous, next, textEdits);
    current = next;
    activeCandidates.clear();
    const identity = { protocolVersion: 1 as const, documentId: next.snapshot.manifest.documentId, revision: next.snapshot.manifest.revision };
    const receipt: TransactionReceipt = freeze({ documentId: identity.documentId, revision: identity.revision, changed: true,
      affectedNodeIds: diff.affectedNodeIds, mapping: diff.mapping });
    emit({ ...identity, type: "documentChanged", affectedNodeIds: [...diff.affectedNodeIds] });
    if (diff.changedAnnotationIds.length) emit({ ...identity, type: "annotationChanged", annotationIds: [...diff.changedAnnotationIds] });
    return receipt;
  }
  function nextRevision(): number {
    const revision = state().snapshot.manifest.revision;
    if (revision >= Number.MAX_SAFE_INTEGER) throw new DocumentError("ResourceLimitExceeded", "The document revision is exhausted.");
    return revision + 1;
  }
  function applyAnnotation(operation: Exclude<DocumentOperation, { nodeId: string }>, annotations: Annotation[], explicit: Set<string>): void {
    const base = state().snapshot.manifest;
    if (operation.op === "addAnnotation") {
      const annotation = operation.annotation;
      if (annotations.some((item) => item.annotationId === annotation.annotationId)
        || annotation.documentId !== base.documentId || annotation.baseRevision !== base.revision) invalid("The new annotation identity or revision is invalid.");
      if (annotation.status !== "needs_reanchor" && annotationNodes(annotation).some((id) => !state().html.hasNode(id))) invalid("Annotation targets must belong to the base snapshot.");
      annotations.push(annotation);
      explicit.add(annotation.annotationId);
    } else {
      const index = annotations.findIndex((item) => item.annotationId === operation.annotationId);
      if (index < 0) invalid("The annotation operation targets a missing annotation.");
      if (operation.op === "removeAnnotation") annotations.splice(index, 1);
      else {
        const annotation = parseContract("annotation", { ...annotations[index], ...operation.changes });
        if (operation.changes.anchor) {
          if (annotation.status !== "needs_reanchor" && annotationNodes(annotation).some((id) => !state().html.hasNode(id))) invalid("Annotation targets must belong to the base snapshot.");
          explicit.add(operation.annotationId);
        }
        annotations[index] = annotation;
      }
    }
  }

  function prepareTransaction(input: DocumentTransaction, resolvedIds?: readonly string[]): CandidateChange {
    if (input && Array.isArray(input.operations)) {
      let count = input.operations.length;
      for (const operation of input.operations) {
        if (operation?.op === "reconcileFragment" && Array.isArray(operation.textEdits)) count += operation.textEdits.length;
        checkLimit(count, budgets.maxOperations);
      }
      checkLimit(count, budgets.maxOperations);
    }
    checkTransactionSize(input, budgets.maxTransactionBytes, limits.maxDepth);
    const transaction = parseContract("document-transaction", input);
    const base = state();
    if (transaction.documentId !== base.snapshot.manifest.documentId) invalid("The transaction belongs to another document.");
    if (transaction.baseRevision !== base.snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "The transaction is based on a different document revision.");
    checkLimit(activeCandidates.size + 1, budgets.maxCandidates);
    // Validate every target against the base before applying any operation.
    for (const operation of transaction.operations) {
      if ("nodeId" in operation && !base.html.hasNode(operation.nodeId)) invalid("An operation targets a node outside the base snapshot.");
    }
    const edit = new DocumentEdit(base, limits, allocate);
    let annotations = parseContract("annotations", base.snapshot.annotations);
    const metadata = parseContract("document-metadata", base.snapshot.metadata);
    const explicit = new Set<string>();
    const unresolved = new Set(base.snapshot.annotations.filter((item) => item.status === "needs_reanchor").map((item) => item.annotationId));
    for (const operation of transaction.operations) {
      if ("nodeId" in operation) edit.apply(operation);
      else applyAnnotation(operation, annotations, explicit);
    }
    for (const annotation of annotations) {
      if (annotation.status !== "needs_reanchor" && unresolved.has(annotation.annotationId) && !explicit.has(annotation.annotationId)) {
        invalid("Reactivating an unresolved annotation requires an explicit anchor.");
      }
      if (explicit.has(annotation.annotationId) && annotation.status !== "needs_reanchor"
        && annotationNodes(annotation).some((id) => !base.html.hasNode(id))) invalid("Annotation targets must belong to the base snapshot.");
    }
    const files = new Map(base.files);
    const source = edit.html();
    files.set(base.snapshot.manifest.entrypoint, encode(source));
    const parsed = parseHtml(source, parseContract("document-manifest", base.snapshot.manifest), files, limits, { nodes: 0 });
    mapAnnotations(annotations, base.html, parsed, edit, explicit);
    updateWordMetadata(metadata, base, parsed);
    if (resolvedIds) {
      const resolved = new Set(resolvedIds);
      const contentChanged = parsed.html !== base.snapshot.html || JSON.stringify(metadata) !== JSON.stringify(base.snapshot.metadata);
      if (!contentChanged) annotations = parseContract("annotations", base.snapshot.annotations);
      else for (const annotation of annotations) {
        if (resolved.has(annotation.annotationId) && annotation.status === "open") annotation.status = "resolved";
      }
    }
    files.set(base.snapshot.manifest.metadata, encode(JSON.stringify(metadata)));
    files.set(base.snapshot.manifest.annotations, encode(JSON.stringify(annotations)));
    const next = buildDocumentState(files, limits);
    checkLimit(retainedBytes([...activeCandidates.values(), next]), budgets.maxCandidateBytes);
    const diff = compareStates(base, next, edit.textEdits);
    const candidate = freeze({ documentId: transaction.documentId, baseRevision: transaction.baseRevision,
      preview: next.snapshot, ...diff }) as unknown as CandidateChange;
    issued.set(candidate, { baseRevision: transaction.baseRevision, consumed: false });
    activeCandidates.set(candidate, next);
    return candidate;
  }

  return Object.freeze({
    snapshot: () => state().snapshot,
    getNode: (id: string) => state().html.getNode(id),
    getAnnotations: () => state().snapshot.annotations,
    serialize: () => new Map([...state().files].map(([path, bytes]) => [path, new Uint8Array(bytes)])),
    prepare: (input: DocumentTransaction): CandidateChange => execute(() => prepareTransaction(input)),
    buildInteractionRequest: (input: InteractionRequestInput): DocumentInteractionRequest => execute(() => {
      checkTransactionSize(input, budgets.maxTransactionBytes, limits.maxDepth);
      const request = buildInteractionRequest(state(), input);
      checkLimit(encode(JSON.stringify(request)).byteLength, limits.maxPackageBytes);
      return request;
    }),
    prepareProposal: (requestInput: DocumentInteractionRequest, proposalInput: EditProposal): CandidateChange => execute(() => {
      checkTransactionSize(requestInput, limits.maxPackageBytes * 8, limits.maxDepth + 16);
      checkLimit(encode(JSON.stringify(requestInput)).byteLength, limits.maxPackageBytes);
      checkTransactionSize(proposalInput, budgets.maxTransactionBytes, limits.maxDepth);
      const request = parseContract("interaction-request", requestInput);
      const proposal = parseContract("edit-proposal", proposalInput);
      validateProposal(state(), request, proposal);
      return prepareTransaction({ protocolVersion: 1, documentId: proposal.documentId,
        baseRevision: proposal.baseRevision, operations: proposal.operations }, proposal.annotationIds);
    }),
    commit: (candidate: CandidateChange): TransactionReceipt => execute(() => {
      const record = issued.get(candidate);
      if (!record || record.consumed) invalid("The candidate is foreign, discarded or already consumed.");
      const base = state();
      if (record.baseRevision !== base.snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "The candidate is based on an older document revision.");
      const prepared = activeCandidates.get(candidate);
      if (!prepared) invalid("The candidate is no longer available.");
      if (JSON.stringify(base.snapshot) === JSON.stringify(prepared.snapshot)) {
        record.consumed = true;
        activeCandidates.delete(candidate);
        return emptyReceipt();
      }
      const next = rebase(prepared, nextRevision(), limits);
      const entries = [...undo, { before: base, after: next, textEdits: candidate.mapping.textEdits }];
      checkLimit(entries.length, budgets.maxHistorySteps);
      checkLimit(retainedBytes(historyStates(entries)), budgets.maxHistoryBytes);
      record.consumed = true;
      undo = entries;
      redo = [];
      return publish(next, candidate.mapping.textEdits);
    }),
    discard: (candidate: CandidateChange): void => execute(() => {
      const record = issued.get(candidate);
      if (!record || record.consumed) invalid("The candidate is foreign, discarded or already consumed.");
      record.consumed = true;
      activeCandidates.delete(candidate);
    }),
    undo: (): TransactionReceipt => execute(() => {
      const entry = undo.at(-1);
      if (!entry) return emptyReceipt();
      const next = rebase(entry.before, nextRevision(), limits);
      const edits = [...entry.textEdits].reverse().map((item) => ({ nodeId: item.nodeId, from: item.from, to: item.from + item.insertedLength, insertedLength: item.to - item.from }));
      undo = undo.slice(0, -1);
      redo = [...redo, entry];
      return publish(next, edits);
    }),
    redo: (): TransactionReceipt => execute(() => {
      const entry = redo.at(-1);
      if (!entry) return emptyReceipt();
      const next = rebase(entry.after, nextRevision(), limits);
      redo = redo.slice(0, -1);
      undo = [...undo, entry];
      return publish(next, entry.textEdits);
    }),
    history: () => {
      state();
      return freeze({ undoSteps: undo.length, redoSteps: redo.length, retainedBytes: retainedBytes(historyStates([...undo, ...redo])) });
    },
    clearHistory: (): void => execute(() => { undo = []; redo = []; }),
    subscribe: (listener: DocumentEventListener): (() => void) => {
      state();
      listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    dispose: (): void => {
      if (!current) return;
      editable();
      current = undefined;
      undo = [];
      redo = [];
      activeCandidates.clear();
      issued = new WeakMap();
      reservedIds.clear();
      listeners.clear();
    },
  });
}
