import { createDocumentSession, openDocument } from "@dagent-office/core";
import type { DocumentModel, DocumentOperation, DocumentSession } from "@dagent-office/core";
import { mountPreview, SourceProjection } from "@dagent-office/render";
import type { DocumentPreviewOptions } from "@dagent-office/render";
import { mountDocumentInteraction, mountDocumentTextEditor, mountDocumentWorkbench } from "@dagent-office/editor";

export interface PresentationHost {
  model: DocumentModel | undefined;
  view: ReturnType<typeof mountPreview> | undefined;
  mount(files: [string, number[]][], options?: DocumentPreviewOptions): Promise<void>;
  update(files: [string, number[]][]): void;
  projected(): string;
  session?: DocumentSession;
  edit(operations: DocumentOperation[]): Promise<void>;
  history(action: "undo" | "redo"): Promise<void>;
  wordAdapterGates(): { code: string; children: number }[];
}
declare global { interface Window { presentationHost: PresentationHost } }
const container = document.createElement("div"); container.style.cssText = "width:1328px;margin:0"; document.body.replaceChildren(container); document.body.style.margin = "0";
window.presentationHost = {
  model: undefined, view: undefined,
  async mount(files, options = {}) {
    this.session?.dispose(); delete this.session;
    this.view?.dispose(); this.model = openDocument(new Map(files.map(([path, data]) => [path, new Uint8Array(data)])));
    this.view = mountPreview(container, this.model, options); await this.view.ready; await this.view.settled();
  },
  update(files) { const model = openDocument(new Map(files.map(([path, data]) => [path, new Uint8Array(data)]))); this.view!.update(model); this.model = model; },
  async edit(operations) {
    this.session ??= createDocumentSession(this.model!); const { documentId, revision } = this.session.snapshot().manifest;
    this.session.commit(this.session.prepare({ protocolVersion: 1, documentId, baseRevision: revision, operations }));
    this.model = this.session; this.view!.update(this.model); await this.view!.ready; await this.view!.settled();
  },
  async history(action) {
    this.session![action](); this.view!.update(this.session!); await this.view!.ready; await this.view!.settled();
  },
  wordAdapterGates() {
    const session = createDocumentSession(this.model!);
    try {
      return [mountDocumentInteraction, mountDocumentTextEditor, mountDocumentWorkbench].map(mount => {
        const target = document.createElement("div");
        try { mount(target, session).dispose(); return { code: "unexpected-success", children: target.childElementCount }; }
        catch (error) { return { code: String((error as { code?: string }).code), children: target.childElementCount }; }
      });
    } finally { session.dispose(); }
  },
  projected() {
    const source = this.view!.source()!, raw = new DOMParser().parseFromString(this.model!.snapshot().html, "text/html").querySelector<HTMLElement>('[data-office-kind="presentation"]')!;
    const projection = new SourceProjection(); projection.capture(source.root, raw); return projection.serialize(source.root).outerHTML;
  },
};
