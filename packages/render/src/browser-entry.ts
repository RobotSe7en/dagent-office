import { DocumentError, openDocument } from "@dagent-office/core";
import type { BrowserRenderBridge, BrowserRenderResult } from "./contracts.js";
import { mountPreview } from "./preview.js";

declare global { interface Window { __dagentOfficeRender: BrowserRenderBridge } }
let renderer: ReturnType<typeof mountPreview> | undefined;
function failure(reason: unknown): BrowserRenderResult {
  const error = reason instanceof DocumentError ? reason : new DocumentError("ExportFailed", "Browser rendering failed.");
  return { ok: false, error: error.toJSON() };
}
window.__dagentOfficeRender = {
  async mount(files, options) {
    try {
      renderer?.dispose();
      const input = new Map(files.map(([path, base64]) => [path, Uint8Array.from(atob(base64), (value) => value.charCodeAt(0))]));
      document.body.style.margin = "0";
      const host = document.createElement("div"); document.body.replaceChildren(host);
      renderer = mountPreview(host, openDocument(input), options);
      const report = await renderer.settled(); return { ok: true, report };
    } catch (reason) { return failure(reason); }
  },
  async preparePrint() {
    try {
      if (!renderer) throw new DocumentError("InvalidOperation", "A document must be rendered before printing.");
      const report = await renderer.settled(), print = renderer.printDocument();
      document.head.replaceWith(document.importNode(print.head, true)); document.body.replaceWith(document.importNode(print.body, true));
      await document.fonts.ready;
      await Promise.all([...document.images].map(async (image) => { if (!image.complete || !image.naturalWidth) await image.decode(); }));
      return { ok: true, report };
    } catch (reason) { return failure(reason); }
  },
  pageClip(index) {
    const page = renderer?.layout().pages[index], source = renderer?.source();
    if (!page || !source || !renderer) throw new DocumentError("InvalidOperation", "The requested page does not exist.");
    const outer = renderer.element.getBoundingClientRect(), inner = source.root.getBoundingClientRect();
    return { x: outer.left + inner.left + page.rect.x * 96 / 72, y: outer.top + inner.top + page.rect.y * 96 / 72,
      width: page.rect.width * 96 / 72, height: page.rect.height * 96 / 72 };
  },
  dispose() { renderer?.dispose(); renderer = undefined; },
};
