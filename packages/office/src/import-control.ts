import { DocumentError } from "@dagent-office/core";

export interface OfficeImportProgress {
  readonly phase: "zip" | "metadata" | "worksheet" | "validate";
  readonly completed: number;
  readonly total: number;
  readonly sheetId?: string;
}
export interface OfficeImportControl {
  /** Structurally accepts AbortSignal without requiring DOM types in the SDK. */
  readonly signal?: { readonly aborted: boolean };
  readonly onProgress?: (progress: OfficeImportProgress) => void;
}

/** Cooperative import work; an unfinished package is never exposed to the host. */
export class ImportControl {
  constructor(readonly options: OfficeImportControl = {}) {
    if (options.signal !== undefined && (options.signal === null || typeof options.signal !== "object" || typeof options.signal.aborted !== "boolean")
      || options.onProgress !== undefined && typeof options.onProgress !== "function") throw new DocumentError("InvalidOperation", "Invalid import control options.");
  }
  check(): void {
    if (this.options.signal?.aborted) throw new DocumentError("OperationCancelled", "Document import was cancelled.");
  }
  async checkpoint(progress: OfficeImportProgress): Promise<void> {
    this.check();
    try { this.options.onProgress?.(Object.freeze({ ...progress })); }
    catch { throw new DocumentError("HostSubmissionFailed", "The host import progress callback failed."); }
    // A task boundary allows Worker cancellation messages and browser input to
    // run. Promise-only yielding would keep both waiting in the microtask queue.
    const scheduler = globalThis as unknown as { setTimeout(callback: () => void, delay: number): unknown };
    await new Promise<void>(resolve => scheduler.setTimeout(resolve, 0));
    this.check();
  }
}
