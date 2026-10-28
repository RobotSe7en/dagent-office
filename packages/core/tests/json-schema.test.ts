import { readFileSync } from "node:fs";
import { join } from "node:path";
import { Ajv2020 } from "ajv/dist/2020.js";
import { expect, test } from "vitest";
import { CONTRACT_NAMES, getContractJsonSchema } from "@dagent-office/core";
import { invalidCases, repository, validCases } from "./fixtures.js";

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
