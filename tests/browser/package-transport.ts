/** Test protocol representation; public SDK inputs remain authorized bytes. */
export function encodeTestPackage(files: ReadonlyMap<string, Uint8Array>): [string, string][] {
  return Array.from(files, ([path, bytes]) => [path, Buffer.from(bytes).toString('base64')]);
}
