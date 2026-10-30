import type { DocumentModel, DocumentPackage, OpenDocumentOptions } from "./runtime.js";
import { copyPackage } from "./document/package.js";
import { buildDocumentState } from "./document/state.js";
import { readLimits } from "./document/validation.js";

export { DEFAULT_DOCUMENT_READ_LIMITS } from "./document/validation.js";

/** Open authorized in-memory files. Performs no filesystem, network or model I/O. */
export function openDocument(input: DocumentPackage, options: OpenDocumentOptions = {}): DocumentModel {
  const limits = readLimits(options.limits);
  const state = buildDocumentState(copyPackage(input, limits), limits);
  return Object.freeze({
    snapshot: () => state.snapshot,
    getNode: (nodeId: string) => state.html.getNode(nodeId),
    serialize: () => new Map([...state.files].map(([path, bytes]) => [path, new Uint8Array(bytes)])),
  });
}
