import { readFileSync } from "node:fs";
import { zipSync } from "fflate";
import { expect, it } from "vitest";
import { readDocumentZip, readDocumentZipAsync } from "../src/document-zip.js";
import { ImportControl } from "../src/import-control.js";

const source = (): Uint8Array => readFileSync(new URL("../../../examples/xlsx/15-ten-thousand-rows.xlsx", import.meta.url));
const options = { limits: { maxFileBytes: 32 * 1024 * 1024, maxPackageBytes: 64 * 1024 * 1024 } };
it("extracts the same authorized bytes with asynchronous progress and responsive task turns", async () => {
  const bytes = source(), expected = readDocumentZip(bytes, options), progress: number[] = [];
  let ticks = 0; const timer = setInterval(() => ticks++, 0);
  try {
    const actual = await readDocumentZipAsync(bytes, options, new ImportControl({ onProgress(event) {
      expect(event.phase).toBe("zip"); expect(event.completed).toBeLessThanOrEqual(event.total); expect(Object.isFrozen(event)).toBe(true); progress.push(event.completed);
    } }));
    expect(actual.size).toBe(expected.size);
    for (const [path, content] of expected) expect(Buffer.compare(actual.get(path)!, content)).toBe(0);
    expect(progress).toEqual([...progress].sort((a, b) => a - b)); expect(ticks).toBeGreaterThan(2);
  } finally { clearInterval(timer); }
});

it("cancels inside a large entry, returns no partial package, and can open another document", async () => {
  const bytes = source(), controller = new AbortController(); let inside = false;
  await expect(readDocumentZipAsync(bytes, options, new ImportControl({ signal: controller.signal, onProgress(event) {
    if (event.completed > 1024 * 1024 && event.completed < event.total) { inside = true; controller.abort(); }
  } }))).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(inside).toBe(true);
  expect((await readDocumentZipAsync(zipSync({ "small.xml": new TextEncoder().encode("<ok/>") }), {}, new ImportControl())).size).toBe(1);
});

it("owns the compressed input before yielding and bounds declared and actual expansion", async () => {
  const bytes = zipSync({ "text.xml": new TextEncoder().encode("<value>private-source</value>") });
  const actual = await readDocumentZipAsync(bytes, {}, new ImportControl({ onProgress() { bytes.fill(0); } }));
  expect(new TextDecoder().decode(actual.get("text.xml"))).toBe("<value>private-source</value>");
  await expect(readDocumentZipAsync(source(), { ...options, maxZipBytes: 5 }, new ImportControl())).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  await expect(readDocumentZipAsync(source(), { limits: { maxFileBytes: 5 } }, new ImportControl())).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
});

it("preserves native ZIP error classification and isolates callback failures", async () => {
  const malformed = source().slice(0, -22);
  await expect(readDocumentZipAsync(malformed, options, new ImportControl())).rejects.toMatchObject({ code: "InvalidDocument" });
  const bytes = zipSync({ "x": new Uint8Array([1, 2, 3]) }, { level: 0 });
  bytes[31] = 99;
  await expect(readDocumentZipAsync(bytes, {}, new ImportControl())).rejects.toMatchObject({ code: "InvalidDocument" });
  await expect(readDocumentZipAsync(source(), options, new ImportControl({ onProgress() { throw new Error("host private data"); } }))).rejects.toMatchObject({ code: "HostSubmissionFailed", message: "The host import progress callback failed." });
});
