import { z } from "zod";
import { AnchorSchema, AnnotationSchema, AnnotationsSchema } from "./contracts/annotations.js";
import { AssetDescriptorSchema, DocumentManifestSchema, DocumentMetadataSchema } from "./contracts/document.js";
import { DocumentErrorDataSchema, DocumentEventSchema } from "./contracts/events.js";
import type { DocumentErrorCode } from "./contracts/events.js";
import { DocumentInteractionRequestSchema, EditProposalSchema } from "./contracts/interaction.js";
import { DocumentContentOperationSchema, DocumentOperationSchema, DocumentTransactionSchema } from "./contracts/operations.js";
import { FORMAT_VERSION, PROTOCOL_VERSION } from "./contracts/primitives.js";
import { DocumentError, summarizeIssues } from "./errors.js";

export const contractSchemas = Object.freeze({
  "asset-descriptor": AssetDescriptorSchema,
  "document-manifest": DocumentManifestSchema,
  "document-metadata": DocumentMetadataSchema,
  anchor: AnchorSchema,
  annotation: AnnotationSchema,
  annotations: AnnotationsSchema,
  "content-operation": DocumentContentOperationSchema,
  "document-operation": DocumentOperationSchema,
  "document-transaction": DocumentTransactionSchema,
  "interaction-request": DocumentInteractionRequestSchema,
  "edit-proposal": EditProposalSchema,
  "document-error": DocumentErrorDataSchema,
  "document-event": DocumentEventSchema,
});
export type ContractName = keyof typeof contractSchemas;
export type ContractValue<Name extends ContractName> = z.infer<(typeof contractSchemas)[Name]>;
export const CONTRACT_NAMES: readonly ContractName[] = Object.freeze(
  Object.keys(contractSchemas) as ContractName[],
);
const formatContracts: ReadonlySet<ContractName> = new Set([
  "asset-descriptor", "document-manifest", "document-metadata", "anchor", "annotation", "annotations",
]);

export interface ContractJsonSchema {
  readonly $schema: string;
  readonly $id: string;
  readonly [key: string]: unknown;
}

function schemaFor<Name extends ContractName>(name: Name): (typeof contractSchemas)[Name] {
  if (!Object.hasOwn(contractSchemas, name)) {
    throw new DocumentError("UnsupportedFeature", "Unknown structural contract.");
  }
  return contractSchemas[name];
}

function object(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown> : undefined;
}

function checkVersion(value: unknown, supported: number): void {
  if (typeof value === "number" && Number.isSafeInteger(value) && value !== supported) {
    throw new DocumentError("UnsupportedVersion", "The declared version is not supported.");
  }
}

function checkEnvelope(name: ContractName, input: unknown): void {
  const data = object(input);
  if (!data) return;
  if (name === "document-manifest") {
    if (typeof data.format === "string" && data.format !== "dagent-office") {
      throw new DocumentError("UnsupportedFormat", "The document format is not supported.");
    }
    checkVersion(data.formatVersion, FORMAT_VERSION);
  }
  if (["document-transaction", "interaction-request", "edit-proposal", "document-event"].includes(name)) {
    checkVersion(data.protocolVersion, PROTOCOL_VERSION);
  }
  if (name === "interaction-request") {
    const context = object(data.context);
    if (context) {
      checkVersion(context.formatVersion, FORMAT_VERSION);
      if (context.kind === "presentation" || context.kind === "workbook") {
        throw new DocumentError("UnsupportedFeature", "This interaction context is not implemented yet.");
      }
    }
  }
}

/** Structural validation only; it does not open a package or execute an operation. */
export function parseContract<Name extends ContractName>(name: Name, input: unknown): ContractValue<Name> {
  const schema = schemaFor(name);
  checkEnvelope(name, input);
  const result = schema.safeParse(input);
  if (result.success) {
    // The selected registry entry determines the corresponding generic output.
    return result.data as ContractValue<Name>;
  }
  const code: DocumentErrorCode = name === "edit-proposal" ? "InvalidProposal"
    : ["content-operation", "document-operation", "document-transaction"].includes(name)
      ? "InvalidOperation" : "InvalidDocument";
  throw new DocumentError(code, "Invalid " + name + " data.", summarizeIssues(result.error.issues));
}

/** Each call produces a self-contained schema with no network references or relaxed checks. */
export function getContractJsonSchema(name: ContractName): ContractJsonSchema {
  const schema = z.toJSONSchema(schemaFor(name), {
    target: "draft-2020-12",
    io: "input",
    unrepresentable: "throw",
    cycles: "throw",
    reused: "inline",
  });
  return {
    ...schema,
    $schema: "https://json-schema.org/draft/2020-12/schema",
    $id: "urn:dagent-office:contract:" + name + ":"
      + (formatContracts.has(name) ? FORMAT_VERSION : PROTOCOL_VERSION),
  };
}
