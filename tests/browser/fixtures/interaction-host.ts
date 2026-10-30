import { createDocumentSession, openDocument } from "@dagent-office/core";
import type { DocumentEditingSession, DocumentSelection } from "@dagent-office/core";
import { mountDocumentInteraction } from "@dagent-office/editor";
import type { DocumentInteractionView } from "@dagent-office/editor";

/** A separate plain JavaScript host for public API/lifecycle browser acceptance. */
export interface InteractionHost {
  session: DocumentEditingSession | undefined;
  view: DocumentInteractionView | undefined;
  selections: (DocumentSelection | undefined)[];
  errors: string[];
  mount(files: [string, number[]][]): Promise<void>;
  settle(): Promise<void>;
}
declare global { interface Window { interactionHost: InteractionHost } }
const container = document.createElement("div");
container.style.cssText = "width:1000px;margin:20px";
document.body.replaceChildren(container);
window.interactionHost = {
  session: undefined, view: undefined, selections: [], errors: [],
  async mount(files) {
    this.view?.dispose();
    this.session?.dispose();
    this.session = createDocumentSession(openDocument(new Map(files.map(([path, bytes]) => [path, new Uint8Array(bytes)]))));
    this.view = mountDocumentInteraction(container, this.session, {
      onSelectionChange: (selection) => this.selections.push(selection), onError: (error) => this.errors.push(error.message),
    });
    await this.settle();
  },
  async settle() {
    await Promise.resolve();
    await this.view?.ready;
    await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
  },
};
