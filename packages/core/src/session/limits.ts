import type { DocumentSessionLimits } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { checkLimit } from "../document/validation.js";

export const DEFAULT_DOCUMENT_SESSION_LIMITS: DocumentSessionLimits = Object.freeze({
  maxOperations: 1024,
  maxTransactionBytes: 8 * 1024 * 1024,
  maxCandidates: 8,
  maxCandidateBytes: 64 * 1024 * 1024,
  maxHistorySteps: 100,
  maxHistoryBytes: 64 * 1024 * 1024,
});

export function sessionLimits(input: Partial<DocumentSessionLimits> = {}): DocumentSessionLimits {
  for (const [key, value] of Object.entries(input)) {
    if (!Object.hasOwn(DEFAULT_DOCUMENT_SESSION_LIMITS, key) || !Number.isSafeInteger(value) || value <= 0) {
      throw new DocumentError("InvalidOperation", "Session limits must be positive safe integers with supported names.");
    }
  }
  return Object.freeze({ ...DEFAULT_DOCUMENT_SESSION_LIMITS, ...input });
}

/** Bound allocation before structural parsing or JSON.stringify; strings use an escape upper bound. */
export function checkTransactionSize(value: unknown, maximum: number, maxDepth: number): void {
  const active = new Set<object>();
  let total = 0;
  const visit = (item: unknown, depth: number): void => {
    checkLimit(depth, maxDepth);
    total += typeof item === "string" ? item.length * 6 + 2 : 32;
    checkLimit(total, maximum);
    if (item === null || typeof item !== "object") return;
    if (active.has(item)) throw new DocumentError("InvalidOperation", "Transaction data must be acyclic.");
    active.add(item);
    if (Array.isArray(item)) {
      checkLimit(item.length, maximum);
      for (const child of item) visit(child, depth + 1);
    } else for (const key of Object.keys(item)) {
      total += key.length * 6 + 3;
      checkLimit(total, maximum);
      visit((item as Record<string, unknown>)[key], depth + 1);
    }
    active.delete(item);
  };
  visit(value, 0);
}
