import type { AssetDescriptor } from "./contracts/document.js";
import { DocumentError } from "./errors.js";
import type { AssetProvider, DocumentPackage, OpenDocumentOptions } from "./runtime.js";
import { parseContract } from "./schema-registry.js";
import { assetDigest, validateAsset, validateAssets } from "./document/assets.js";
import { readManifest } from "./document/manifest.js";
import { copyPackage } from "./document/package.js";
import { checkLimit, freeze, invalid, readLimits } from "./document/validation.js";

/** Describe authorized bytes. A digest records integrity; it does not grant access. */
export function createAssetDescriptor(path: string, mimeType: string, bytes: Uint8Array): AssetDescriptor {
  if (!(bytes instanceof Uint8Array)) invalid("Resource contents must be bytes.");
  const asset = parseContract("asset-descriptor", { path, mimeType, sha256: assetDigest(bytes) });
  validateAsset(asset, bytes);
  return freeze(asset);
}

/** Collect missing declared resources through this host's authorized provider. Never fetch URLs. */
export async function collectDocumentAssets(input: DocumentPackage, provider: AssetProvider, options: OpenDocumentOptions = {}): Promise<DocumentPackage> {
  const limits = readLimits(options.limits);
  const files = copyPackage(input, limits);
  const manifest = readManifest(files);
  const missing = manifest.assets.filter((asset) => !files.has(asset.path));
  checkLimit(files.size + missing.length, limits.maxFiles);
  // Validate supplied resources before contacting the host for missing ones.
  validateAssets(manifest.assets.filter((asset) => files.has(asset.path)), files);
  let total = [...files.values()].reduce((sum, bytes) => sum + bytes.byteLength, 0);
  for (const asset of missing) {
    let supplied: Uint8Array | undefined;
    try {
      supplied = await provider.readAsset(freeze({ ...asset }));
    } catch {
      throw new DocumentError("AssetUnavailable", "The host could not supply a declared resource.");
    }
    if (!(supplied instanceof Uint8Array)) throw new DocumentError("AssetUnavailable", "A declared resource is unavailable from the host.");
    checkLimit(supplied.byteLength, limits.maxFileBytes);
    total += supplied.byteLength;
    checkLimit(total, limits.maxPackageBytes);
    const bytes = new Uint8Array(supplied);
    validateAsset(asset, bytes);
    files.set(asset.path, bytes);
  }
  return files;
}
