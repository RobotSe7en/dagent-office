import { openSdkHost } from './sdk-host.js';
import { expect, test } from "@playwright/test";
import { build } from "vite";
import { fileURLToPath } from "node:url";
import { writeFile } from "node:fs/promises";

let workerCode = "";
test.beforeAll(async () => {
  const bundle = await build({ configFile: false, build: { write: false, minify: false,
    lib: { entry: fileURLToPath(new URL("./fixtures/workbook-storage-worker.ts", import.meta.url)), formats: ["iife"], name: "StorageProbe" } } });
  if (!Array.isArray(bundle) && !("output" in bundle)) throw new Error("Unexpected storage probe bundle");
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output;
  workerCode = output.filter(part => part.type === "chunk").map(part => part.code).join("\n");
});

test("10,000 actual rows stay in a Worker; windows and cancellation have bounded replies", { tag: "@scale" }, async ({ page }, info) => {
  test.setTimeout(90_000);
  await openSdkHost(page);
  const report = await page.evaluate(async code => {
    const url = URL.createObjectURL(new Blob([code], { type: "text/javascript" })), worker = new Worker(url);
    const messages: Record<string, unknown>[] = [], delays: number[] = [];
    let previous = performance.now(), nextId = 0, ticks = 0, cancellationId: number | undefined;
    const timer = setInterval(() => { const now = performance.now(); delays.push(Math.max(0, now - previous - 10)); previous = now; ticks++; }, 10);
    worker.addEventListener("message", event => {
      messages.push(event.data as Record<string, unknown>);
      if (event.data.kind === "progress" && event.data.id === cancellationId) { cancellationId = undefined; worker.postMessage({ kind: "cancel", id: 99 }); }
    });
    function request(kind: "open" | "window", row?: number): Promise<Record<string, unknown>> {
      const id = ++nextId, expected = kind === "open" ? "opened" : "window";
      return new Promise((resolve, reject) => {
        const timeout = setTimeout(() => { cleanup(); reject(new Error("Storage Worker timed out")); }, 60_000);
        const error = (event: ErrorEvent) => { cleanup(); reject(new Error(event.message)); };
        const message = (event: MessageEvent<Record<string, unknown>>) => {
          if (event.data.id !== id) return;
          if (event.data.kind === expected) { cleanup(); resolve(event.data); }
          else if (event.data.kind === "error") { cleanup(); reject(new Error(String(event.data.message))); }
        };
        function cleanup() { clearTimeout(timeout); worker.removeEventListener("message", message); worker.removeEventListener("error", error); }
        worker.addEventListener("message", message); worker.addEventListener("error", error); worker.postMessage({ kind, id, row });
      });
    }
    try {
      const opened = await request("open"), last = await request("window", 9999);
      cancellationId = ++nextId; const cancelledOpenId = cancellationId;
      const cancelled = new Promise<Record<string, unknown>>((resolve, reject) => {
        const timeout = setTimeout(() => reject(new Error("Cancellation timed out")), 10_000);
        const listener = (event: MessageEvent<Record<string, unknown>>) => { if (event.data.kind === "cancelled") { clearTimeout(timeout); worker.removeEventListener("message", listener); resolve(event.data); } };
        worker.addEventListener("message", listener);
      });
      worker.postMessage({ kind: "open", id: cancelledOpenId });
      const cancellation = await cancelled, retained = await request("window", 9999);
      delays.sort((a, b) => a - b);
      return { opened, last, retained, cancellation, cancelledPublished: messages.some(m => m.kind === "opened" && m.id === cancelledOpenId), ticks,
        eventLoopDelayP95Ms: delays[Math.floor(delays.length * .95)] ?? 0,
        largestReplyBytes: Math.max(...messages.map(message => new TextEncoder().encode(JSON.stringify(message)).byteLength)),
        browser: navigator.userAgent };
    } finally { clearInterval(timer); worker.terminate(); URL.revokeObjectURL(url); }
  }, workerCode);
  expect(report.opened.cellCount).toBe(200_000); expect(report.opened.formulaCount).toBe(20_000);
  expect(report.opened.indexBytes).toBe(2_120_000); expect(report.opened.formulaTemplates).toBe(2); expect(report.last.cells).toHaveLength(20);
  expect(JSON.stringify(report.last.cells)).toContain("客户 9999 😀禰󠄀");
  expect(report.retained.cells).toEqual(report.last.cells); expect(report.cancellation.retainedChunks).toBe(40);
  expect(report.cancelledPublished).toBe(false); expect(report.ticks).toBeGreaterThan(10); expect(report.largestReplyBytes).toBeLessThan(10_000);
  await info.attach("workbook-storage-probe.json", { body: JSON.stringify(report, null, 2), contentType: "application/json" });
  await writeFile(info.outputPath("workbook-storage-probe.json"), JSON.stringify(report, null, 2) + "\n");
});
