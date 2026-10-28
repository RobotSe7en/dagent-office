import { readFileSync } from "node:fs";
import { join } from "node:path";
import { Ajv2020 } from "ajv/dist/2020.js";
import { expect, test } from "vitest";
import { CONTRACT_NAMES, DocumentError, getContractJsonSchema, parseContract } from "@dagent-office/core";
import { invalidCases, manifest, repository, validCases } from "./fixtures.js";

const ajv = new Ajv2020({ strict: true, allErrors: true });
const validators = new Map(CONTRACT_NAMES.map((name) => [name, ajv.compile(getContractJsonSchema(name))]));

test.each(validCases)("independently validates %s with JSON Schema", (name, input) => {
  const validate = validators.get(name)!;
  expect(validate(input), JSON.stringify(validate.errors)).toBe(true);
});

test.each(invalidCases.map(([name, input], index) => [index, name, input] as const))(
  "JSON Schema rejects invalid case %i (%s)", (_index, name, input) => {
    expect(validators.get(name)!(input)).toBe(false);
  },
);

test.each(["\u2028", "\u2029"])("both validators reject dot segments after %j", (separator) => {
  const validate = validators.get("document-manifest")!;
  for (const suffix of ["../../outside.html", "./outside.html", ".", ".."]) {
    const input = { ...manifest, entrypoint: "a" + separator + "/" + suffix };
    expect(validate(input)).toBe(false);
    expect(() => parseContract("document-manifest", input)).toThrow(DocumentError);
  }
  const canonical = { ...manifest, entrypoint: "a" + separator + "/inside.html" };
  expect(validate(canonical)).toBe(true);
  expect(parseContract("document-manifest", canonical)).toEqual(canonical);
});

test.each([
  { field: "message", minimum: 1, maximum: 256 },
  { field: "issue.code", minimum: 1, maximum: 64 },
  { field: "issue.path", minimum: 0, maximum: 128 },
  { field: "issue.message", minimum: 1, maximum: 256 },
])("both validators count Unicode code points in $field", ({ field, minimum, maximum }) => {
  for (const character of ["a", "中", "😀", "\n", "\u2028", "\u2029"]) {
    for (const length of [0, Math.floor(maximum / 2) + 1, maximum - 1, maximum, maximum + 1]) {
      const value = character.repeat(length);
      const error = {
        code: "InvalidDocument",
        message: field === "message" ? value : "Invalid document.",
        details: [{
          code: field === "issue.code" ? value : "invalid_type",
          path: [field === "issue.path" ? value : "message"],
          message: field === "issue.message" ? value : "Invalid value.",
        }],
      };
      const event = { protocolVersion: 1, documentId: "d", revision: 0, type: "error", error };
      const valid = length >= minimum && length <= maximum;
      for (const [name, input] of [["document-error", error], ["document-event", event]] as const) {
        expect(validators.get(name)!(input), name + ": " + field + ", " + JSON.stringify(value)).toBe(valid);
        if (valid) expect(parseContract(name, input)).toEqual(input);
        else expect(() => parseContract(name, input)).toThrow(DocumentError);
      }
    }
  }
});

test.each(CONTRACT_NAMES)("build artifact %s matches the current source contract", (name) => {
  const path = join(repository, "packages/core/dist/schemas", name + ".json");
  expect(JSON.parse(readFileSync(path, "utf8"))).toEqual(getContractJsonSchema(name));
});

test("schemas have unique versioned identifiers and generation does not share mutable output", () => {
  const ids = CONTRACT_NAMES.map((name) => getContractJsonSchema(name).$id);
  expect(new Set(ids).size).toBe(CONTRACT_NAMES.length);
  const schema = getContractJsonSchema("document-manifest");
  const properties = schema.properties as Record<string, Record<string, unknown>>;
  properties.format!.const = "changed";
  delete (schema as Record<string, unknown>).properties;
  expect(getContractJsonSchema("document-manifest")).toHaveProperty("properties.format.const", "dagent-office");
});
