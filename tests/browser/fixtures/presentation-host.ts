import { openDocument } from "@dagent-office/core";
import type { DocumentModel } from "@dagent-office/core";
import { mountPreview, SourceProjection } from "@dagent-office/render";
import type { DocumentPreviewOptions } from "@dagent-office/render";

export interface PresentationHost {
  model: DocumentModel | undefined;
  view: ReturnType<typeof mountPreview> | undefined;
  mount(files: [string, number[]][], options?: DocumentPreviewOptions): Promise<void>;
  update(files: [string, number[]][]): void;
  projected(): string;
}
declare global { interface Window { presentationHost: PresentationHost } }
const container = document.createElement("div"); container.style.cssText = "width:1328px;margin:0"; document.body.replaceChildren(container); document.body.style.margin = "0";
window.presentationHost = {
  model: undefined, view: undefined,
  async mount(files, options = {}) {
    this.view?.dispose(); this.model = openDocument(new Map(files.map(([path, data]) => [path, new Uint8Array(data)])));
    this.view = mountPreview(container, this.model, options); await this.view.ready; await this.view.settled();
  },
  update(files) { const model = openDocument(new Map(files.map(([path, data]) => [path, new Uint8Array(data)]))); this.view!.update(model); this.model = model; },
  projected() {
    const source = this.view!.source()!, raw = new DOMParser().parseFromString(this.model!.snapshot().html, "text/html").querySelector<HTMLElement>('[data-office-kind="presentation"]')!;
    const projection = new SourceProjection(); projection.capture(source.root, raw); return projection.serialize(source.root).outerHTML;
  },
};
