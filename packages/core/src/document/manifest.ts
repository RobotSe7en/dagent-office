import type { DocumentManifest } from "../contracts/document.js";
import type { DocumentPackage } from "../runtime.js";
import { parseContract } from "../schema-registry.js";
import { readJson } from "./package.js";
import { invalid } from "./validation.js";

export function readManifest(files: DocumentPackage): DocumentManifest {
  const manifest = parseContract("document-manifest", readJson(files, "manifest.json"));
  const roles = ["manifest.json", manifest.entrypoint, manifest.metadata, manifest.annotations,
    ...manifest.styles, ...manifest.assets.map((asset) => asset.path)];
  if (new Set(roles).size !== roles.length) invalid("Manifest file roles and resource paths must not overlap or repeat.");
  return manifest;
}
