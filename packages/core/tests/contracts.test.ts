import { describe, expect, expectTypeOf, test } from "vitest";
import {
  CONTRACT_NAMES, DocumentError, getContractJsonSchema, parseContract,
} from "@dagent-office/core";
import type {
  ContractName, DocumentContentOperation, DocumentErrorCode, DocumentManifest, EditProposal,
} from "@dagent-office/core";
import { annotations, invalidCases, manifest, metadata, proposal, request, validCases } from "./fixtures.js";

function failure(name: ContractName, input: unknown, code?: DocumentErrorCode): DocumentError {
  try {
    parseContract(name, input);
  } catch (error) {
    if (!(error instanceof DocumentError)) throw error;
    if (code) expect(error.code).toBe(code);
    return error;
  }
  throw new Error("Expected a contract error.");
}

function freeze<Value>(value: Value): Value {
  if (value && typeof value === "object") {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
}

describe("strict structural contracts", () => {
  test.each(validCases)("accepts the %s sample without normalization", (name, input) => {
    const original = JSON.stringify(input);
    expect(parseContract(name, freeze(structuredClone(input)))).toEqual(input);
    expect(JSON.stringify(input)).toBe(original);
  });

  test.each(invalidCases.map(([name, input], index) => [index, name, input] as const))(
    "rejects invalid case %i (%s)", (_index, name, input) => {
      expect(() => parseContract(name, input)).toThrow(DocumentError);
    },
  );

  test("has independent valid examples covering every exported contract", () => {
    expect([...new Set(validCases.map(([name]) => name))].sort()).toEqual([...CONTRACT_NAMES].sort());
  });

  test("returns a separate nested result without sharing caller-owned document content", () => {
    const input = freeze(structuredClone(request));
    const result = parseContract("interaction-request", input);
    expect(result).not.toBe(input);
    expect(result.context).not.toBe(input.context);
    result.annotations[0]!.instruction = "Modified copy";
    expect(input.annotations[0]!.instruction).toBe(request.annotations[0]!.instruction);
  });

  test.each([
    { kind: "text", nodeId: "n", from: 0, to: 0, quote: "" },
    { kind: "text-set", ranges: [{ nodeId: "n", from: 0, to: 2, quote: "😀" }] },
    { kind: "node", nodeId: "n" },
    { kind: "node-set", nodeIds: ["n1", "n2"] },
    { kind: "region", surfaceId: "root", rect: { x: 0, y: 0, width: 10, height: 20 }, nodeIds: [], layoutKey: "fixed", pageId: "page-1" },
  ])("supports the defined $kind anchor", (input) => {
    expect(parseContract("anchor", input)).toEqual(input);
  });

  test.each([
    { op: "replaceText", nodeId: "n", from: 0, to: 0, expectedText: "", text: "" },
    { op: "replaceFragment", nodeId: "n", html: "" },
    { op: "insertFragment", nodeId: "n", position: "after", html: "<p>x</p>" },
    { op: "removeNode", nodeId: "n" },
    { op: "setStyle", nodeId: "n", property: "color", value: "#1f4b99" },
    { op: "replaceAsset", nodeId: "n", assetPath: "assets/new.png" },
  ])("supports the $op content contract", (input) => {
    expect(parseContract("content-operation", input)).toEqual(input);
  });

  test("keeps annotation operations in transactions and out of proposals", () => {
    const operations = [
      { op: "addAnnotation", annotation: annotations[0] },
      { op: "updateAnnotation", annotationId: "a-summary", changes: { status: "resolved" } },
      { op: "removeAnnotation", annotationId: "a-summary" },
    ];
    expect(parseContract("document-transaction", {
      protocolVersion: 1, documentId: "d", baseRevision: 0, operations,
    }).operations).toEqual(operations);
    failure("edit-proposal", { ...proposal, operations }, "InvalidProposal");
    expectTypeOf<Extract<DocumentContentOperation, { op: "addAnnotation" }>>().toEqualTypeOf<never>();
  });

  test("infers the selected output type and preserves optional metadata without defaults", () => {
    expectTypeOf(parseContract("document-manifest", manifest)).toEqualTypeOf<DocumentManifest>();
    expectTypeOf(parseContract("edit-proposal", proposal)).toEqualTypeOf<EditProposal>();
    const { title: _title, locale: _locale, ...withoutOptional } = metadata;
    expect(parseContract("document-metadata", withoutOptional)).toEqual(withoutOptional);
  });

  test.each(["/x", "../x", "a/../../x", "./x", "a/./x", "a//x", "a/", "C:\\x", "a\\x", "https://x", "x?query", "x#fragment", "x\u0000y", "x\n", "x\r\n", "x\t"])(
    "rejects noncanonical package path %j", (entrypoint) => {
      failure("document-manifest", { ...manifest, entrypoint }, "InvalidDocument");
    },
  );

  test("preserves spaces and Unicode in canonical package paths", () => {
    const input = { ...manifest, entrypoint: "正文/项目 报告.html" };
    expect(parseContract("document-manifest", input)).toEqual(input);
  });

  test("accepts structurally valid empty transaction batches unchanged", () => {
    const input = { protocolVersion: 1, documentId: "d", baseRevision: 3, operations: [] };
    expect(parseContract("document-transaction", input)).toEqual(input);
  });

  test("does not confuse structural acceptance with HTML, reference or transaction validation", () => {
    const input = { op: "replaceFragment", nodeId: "missing", html: "<script>untrustedData()</script>" };
    expect(parseContract("content-operation", input)).toEqual(input);
    // Opening/sanitizing HTML and executing operations belong to R1-04/R1-06.
  });
});

describe("version and error boundary", () => {
  test.each([0, 2, 999])("rejects unsupported format and protocol version %i", (version) => {
    failure("document-manifest", { ...manifest, formatVersion: version }, "UnsupportedVersion");
    failure("interaction-request", { ...request, protocolVersion: version }, "UnsupportedVersion");
    failure("edit-proposal", { ...proposal, protocolVersion: version }, "UnsupportedVersion");
    failure("interaction-request", {
      ...request, context: { ...request.context, formatVersion: version },
    }, "UnsupportedVersion");
  });

  test("does not guess or coerce an absent or string version", () => {
    failure("document-manifest", { ...manifest, formatVersion: undefined }, "InvalidDocument");
    failure("document-manifest", { ...manifest, formatVersion: "1" }, "InvalidDocument");
  });

  test("classifies unsupported formats and not-yet-supported interaction kinds", () => {
    failure("document-manifest", { ...manifest, format: "other" }, "UnsupportedFormat");
    failure("interaction-request", { ...request, context: { ...request.context, kind: "presentation" } }, "UnsupportedFeature");
  });

  test("rejects unknown fields without changing caller data or leaking it into errors", () => {
    const input = freeze({ ...manifest, titleTypo: "private text ".repeat(1000) });
    const before = JSON.stringify(input);
    const error = failure("document-manifest", input, "InvalidDocument");
    expect(JSON.stringify(input)).toBe(before);
    expect(JSON.stringify(error)).not.toContain("private text");
    expect(error).not.toHaveProperty("cause");
    expect(parseContract("document-error", error.toJSON())).toEqual(error.toJSON());
  });

  test("bounds error details even when many fields are invalid", () => {
    const error = failure("annotations", Array.from({ length: 100 }, () => ({})), "InvalidDocument");
    expect(error.details.length).toBeGreaterThan(0);
    expect(error.details.length).toBeLessThanOrEqual(16);
    expect(parseContract("document-error", error.toJSON())).toEqual(error.toJSON());
  });

  test.each([
    ["annotations", (items: unknown[]) => items, "InvalidDocument", [0]],
    ["document-metadata", (items: unknown[]) => ({ ...metadata, sections: items }), "InvalidDocument", ["sections", 0]],
    ["anchor", (items: unknown[]) => ({ kind: "text-set", ranges: items }), "InvalidDocument", ["ranges", 0]],
    ["interaction-request", (items: unknown[]) => ({ ...request, context: { ...request.context, blocks: items } }),
      "InvalidDocument", ["context", "blocks", 0]],
    ["edit-proposal", (items: unknown[]) => ({ ...proposal, operations: items }), "InvalidProposal", ["operations", 0]],
    ["document-transaction", (items: unknown[]) => ({ protocolVersion: 1, documentId: "d", baseRevision: 0,
      operations: [{ op: "addAnnotation", annotation: { ...annotations[0], anchor: { kind: "text-set", ranges: items } } }] }),
      "InvalidOperation", ["operations", 0, "annotation", "anchor", "ranges", 0]],
  ] as const)("stops at an invalid array item in %s before reading its tail", (name, input, code, prefix) => {
    const items: unknown[] = [{}];
    Object.defineProperty(items, 1, { get() { throw new Error("The array tail must not be read."); } });
    const error = failure(name, input(items), code);
    expect(error.details.length).toBeGreaterThan(0);
    expect(error.details.every((issue) => prefix.every((part, index) => issue.path[index] === part))).toBe(true);
  });

  test("unknown contract names cannot resolve inherited registry properties", () => {
    const name = "constructor" as ContractName;
    failure(name, {}, "UnsupportedFeature");
    expect(() => getContractJsonSchema(name)).toThrow(DocumentError);
  });
});
