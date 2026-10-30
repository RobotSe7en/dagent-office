import { createDocumentSession, openDocument } from "@dagent-office/core";
import type { DocumentSession } from "@dagent-office/core";
import { mountDocumentTextEditor } from "@dagent-office/editor";
import type { DocumentTextEditor } from "@dagent-office/editor";
export interface TextEditorHost {
  session: DocumentSession | undefined;
  view: DocumentTextEditor | undefined;
  errors: string[];
  mount(files: [string, number[]][]): Promise<void>;
}
declare global { interface Window { textEditorHost: TextEditorHost } }
const container = document.createElement("div"); container.style.cssText = "width:1050px;margin:20px"; document.body.replaceChildren(container);
window.textEditorHost = {
  session: undefined, view: undefined, errors: [],
  async mount(files) {
    this.view?.dispose(); this.session?.dispose();
    this.session = createDocumentSession(openDocument(new Map(files.map(([path, bytes]) => [path, new Uint8Array(bytes)]))));
    this.view = mountDocumentTextEditor(container, this.session, { onError: (error) => this.errors.push(error.message) });
    await this.view.ready;
  },
};
