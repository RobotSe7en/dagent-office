/** Avoid one Playwright protocol object for every byte. Legacy small inputs
 * remain accepted while all SDK calls still receive independent byte arrays. */
export type TestPackage = readonly (readonly [string, string | readonly number[]])[];

export function decodeTestBytes(value: string | readonly number[]): Uint8Array {
  if (typeof value !== 'string') return Uint8Array.from(value);
  const text = atob(value), bytes = new Uint8Array(text.length);
  for (let index = 0; index < text.length; index++) bytes[index] = text.charCodeAt(index);
  return bytes;
}

export function decodeTestPackage(entries: TestPackage): Map<string, Uint8Array> {
  return new Map(entries.map(([path, value]) => [path, decodeTestBytes(value)]));
}
