import { createDocumentSession, openDocument } from "@dagent-office/core";
import type { DocumentSession } from "@dagent-office/core";
import { mountPresentationEditor } from "@dagent-office/editor";
import type { PresentationEditorView, PresentationEditorOptions } from "@dagent-office/editor";
export interface PresentationEditorHost {
  session:DocumentSession|undefined;
  view:PresentationEditorView|undefined;
  errors:string[];
  mount(files:[string,number[]][],options?:PresentationEditorOptions):Promise<void>;
}
declare global {interface Window{presentationEditorHost:PresentationEditorHost}}
const container=document.createElement("div");container.style.width="1500px";document.body.replaceChildren(container);
window.presentationEditorHost={
  session:undefined,view:undefined,errors:[],
  async mount(files,options={}) {
    this.view?.dispose();this.session?.dispose();this.errors=[];
    this.session=createDocumentSession(openDocument(new Map(files.map(([path,bytes])=>[path,new Uint8Array(bytes)]))));
    this.view=mountPresentationEditor(container,this.session,{...options,onError:error=>{this.errors.push(error.message);}});await this.view.ready;
  },
};
