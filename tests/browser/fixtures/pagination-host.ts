import { createDocumentSession, openDocument } from "@dagent-office/core";
import type { DocumentModel, DocumentSession } from "@dagent-office/core";
import { createLayoutFieldTransaction, mountDocumentInteraction, mountDocumentPreview, mountDocumentTextEditor } from "@dagent-office/editor";
import type { DocumentInteractionView, DocumentPreview, DocumentPreviewOptions, DocumentRenderReport, DocumentTextEditor } from "@dagent-office/editor";

export interface PaginationHost {
  model: DocumentModel | undefined;
  session: DocumentSession | undefined;
  view: DocumentPreview | DocumentInteractionView | DocumentTextEditor | undefined;
  interaction: DocumentInteractionView | undefined;
  errors: string[];
  mount(files: [string, number[]][], mode?: "preview" | "edit" | "interaction", options?: DocumentPreviewOptions): Promise<void>;
  layout(): DocumentRenderReport;
  updateFields(): Promise<void>;
}
declare global { interface Window { paginationHost: PaginationHost } }
const container = document.createElement("div"); container.style.cssText = "width:1050px;margin:20px"; document.body.replaceChildren(container);
window.paginationHost = {
  model: undefined, session: undefined, view: undefined, interaction: undefined, errors: [],
  async mount(files, mode = "preview", options = {}) {
    this.view?.dispose(); this.session?.dispose(); this.session = undefined; this.interaction = undefined; this.errors = [];
    this.model = openDocument(new Map(files.map(([path, bytes]) => [path, new Uint8Array(bytes)])));
    if (mode === "preview") this.view = mountDocumentPreview(container, this.model, options);
    else {
      this.session = createDocumentSession(this.model);
      const input = { ...options, onError: (error: Error) => this.errors.push(error.message) };
      this.view = mode === "edit" ? mountDocumentTextEditor(container, this.session, input) : this.interaction = mountDocumentInteraction(container, this.session, input);
    }
    await this.view.ready;
  },
  layout() { const result = this.view?.layout(); if (!result) throw new Error("Layout is not ready."); return result; },
  async updateFields() {
    const session = this.session!;
    session.commit(session.prepare(createLayoutFieldTransaction(session, this.layout())));
    await Promise.resolve(); await this.view!.ready;
  },
};
