import type { ZodIssue } from "zod";
import type { DocumentErrorCode, DocumentErrorData, ValidationIssue } from "./contracts/events.js";

const issueMessages: Readonly<Record<string, string>> = {
  unrecognized_keys: "Unknown fields are not allowed.",
  invalid_type: "Expected the declared value type.",
  invalid_value: "Value is outside the declared choices.",
  invalid_format: "Value does not match the required format.",
  too_small: "Value is below the permitted bound.",
  too_big: "Value exceeds the permitted bound.",
  invalid_union: "Value does not match a supported contract variant.",
};

export function summarizeIssues(issues: readonly ZodIssue[]): ValidationIssue[] {
  return issues.slice(0, 16).map((issue) => ({
    code: issue.code.slice(0, 64),
    path: issue.path.slice(0, 16).map((part) =>
      typeof part === "number" ? part : String(part).slice(0, 128)),
    message: issueMessages[issue.code] ?? "Value violates the structural contract.",
  }));
}

/** A bounded SDK error; neither input values nor a raw validator error are retained. */
export class DocumentError extends Error {
  readonly code: DocumentErrorCode;
  readonly details: readonly ValidationIssue[];

  constructor(code: DocumentErrorCode, message: string, details: readonly ValidationIssue[] = []) {
    super((message || code).slice(0, 256));
    this.name = "DocumentError";
    this.code = code;
    this.details = details.slice(0, 16).map((issue) => ({
      code: issue.code.slice(0, 64),
      path: issue.path.slice(0, 16).map((part) =>
        typeof part === "number" ? part : part.slice(0, 128)),
      message: (issue.message || "Invalid value.").slice(0, 256),
    }));
  }

  toJSON(): DocumentErrorData {
    return {
      code: this.code,
      message: this.message,
      details: this.details.map((issue) => ({ ...issue, path: [...issue.path] })),
    };
  }
}
