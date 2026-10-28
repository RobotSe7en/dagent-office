import { mkdirSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { CONTRACT_NAMES, getContractJsonSchema } from "../packages/core/dist/index.js";

const directory = new URL("../packages/core/dist/schemas/", import.meta.url);
const schemas = CONTRACT_NAMES.map((name) => [name, getContractJsonSchema(name)]);
mkdirSync(directory, { recursive: true });
for (const [name, schema] of schemas) {
  writeFileSync(new URL(name + ".json", directory), JSON.stringify(schema, null, 2) + "\n");
}
console.log("Generated " + schemas.length + " structural schemas in " + fileURLToPath(directory));
