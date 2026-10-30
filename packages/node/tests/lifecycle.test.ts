import { readFileSync } from "node:fs";
import { chromium } from "playwright-core";
import type { Browser, BrowserContext } from "playwright-core";
import { expect, test, vi } from "vitest";
import { openDocument } from "@dagent-office/core";
import { createDocumentConverter } from "@dagent-office/node";
function deferred<T>(): { promise: Promise<T>; resolve(value: T): void } {
  let resolve: (value: T) => void = () => {};
  const promise = new Promise<T>((done) => { resolve = done; }); return { promise, resolve };
}

test("a timed-out pending context keeps its quota slot until backend work and cleanup finish", async () => {
  const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((path) => [path,
    new Uint8Array(readFileSync(new URL("../../../examples/report/" + path, import.meta.url))),
  ]));
  const context = deferred<BrowserContext>(), cleanup = deferred<void>();
  const contextClose = vi.fn(async () => {}), browserClose = vi.fn(() => cleanup.promise);
  const browser = { newContext: vi.fn(() => context.promise), close: browserClose } as unknown as Browser;
  const launch = vi.spyOn(chromium, "launch").mockResolvedValue(browser);
  const converter = createDocumentConverter({ maxConcurrency: 1, timeoutMs: 1000 }), model = openDocument(files);
  try {
    const pending = converter.renderPdf(model, { timeoutMs: 30 });
    await expect(pending).rejects.toMatchObject({ code: "Timeout" });
    expect(browserClose).toHaveBeenCalledTimes(1);
    await expect(converter.renderPdf(model)).rejects.toMatchObject({ code: "Busy" });
    expect(launch).toHaveBeenCalledTimes(1);
    let closed = false; const closing = converter.close().then(() => { closed = true; });
    await Promise.resolve(); expect(closed).toBe(false);
    context.resolve({ close: contextClose } as unknown as BrowserContext);
    await Promise.resolve(); expect(closed).toBe(false);
    cleanup.resolve(); await closing;
    expect(contextClose).toHaveBeenCalledTimes(1); expect(browserClose).toHaveBeenCalledTimes(1);
  } finally {
    context.resolve({ close: contextClose } as unknown as BrowserContext); cleanup.resolve(); await converter.close(); launch.mockRestore();
  }
});
