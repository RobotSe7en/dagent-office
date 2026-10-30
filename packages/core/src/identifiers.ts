import { DocumentError } from "./errors.js";

/** Opaque correlation IDs never grant access. getRandomValues remains available
 * in browsers that omit randomUUID on an HTTP development origin. */
export function createOpaqueId(): string {
  const crypto = (globalThis as unknown as { crypto?: { randomUUID?(): string; getRandomValues?(target: Uint8Array): Uint8Array } }).crypto;
  if (typeof crypto?.randomUUID === "function") return crypto.randomUUID();
  if (typeof crypto?.getRandomValues !== "function") throw new DocumentError("UnsupportedFeature", "Identifier generation requires the host runtime's Web Crypto random values.");
  const bytes = crypto.getRandomValues(new Uint8Array(16)); bytes[6] = (bytes[6]! & 15) | 64; bytes[8] = (bytes[8]! & 63) | 128;
  const hex = [...bytes].map((value) => value.toString(16).padStart(2, "0"));
  return [hex.slice(0, 4), hex.slice(4, 6), hex.slice(6, 8), hex.slice(8, 10), hex.slice(10)].map((part) => part.join("")).join("-");
}
