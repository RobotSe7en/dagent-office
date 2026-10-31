import { openDocument } from "@dagent-office/core";
import type { DocumentModel } from "@dagent-office/core";

export interface PreparedPreview { readonly model: DocumentModel; readonly document: Document; readonly urls: string[] }
export function packagePath(reference: string, entrypoint: string): string {
  const parts = entrypoint.split("/").slice(0, -1);
  for (const token of reference.split("/")) { const value = decodeURIComponent(token); if (value === "..") parts.pop(); else if (value !== ".") parts.push(value); }
  return parts.join("/");
}
export function safeStyleText(source: string): string { return source.replace(/<\/style/gi, value => "<\\" + value.slice(1)); }
/** Shared scriptless package preparation. Object URLs are scoped to this view,
 * including repeated references to one asset, and are revoked by its owner. */
export function preparePreviewPackage(input: DocumentModel): PreparedPreview {
  const model = openDocument(input.serialize()), snapshot = model.snapshot(), files = model.serialize();
  const document = new DOMParser().parseFromString(snapshot.html, "text/html");
  for (const link of document.querySelectorAll("link")) link.remove();
  const policy = document.createElement("meta"); policy.httpEquiv = "Content-Security-Policy";
  policy.content = "default-src 'none'; img-src blob:; style-src 'unsafe-inline'; font-src 'none'; base-uri 'none'; form-action 'none'"; document.head.prepend(policy);
  const style = document.createElement("style"); style.dataset.sdkSourceStyles = "";
  style.textContent = safeStyleText(snapshot.styles.map(s => s.css).join("\n")); document.head.append(style);
  const urls: string[] = [], assets = new Map(snapshot.manifest.assets.map(asset => [asset.path, asset])), byPath = new Map<string, string>();
  for (const image of document.querySelectorAll("img")) {
    const path = packagePath(image.getAttribute("src")!, snapshot.manifest.entrypoint);
    let url = byPath.get(path);
    if (!url) { const asset = assets.get(path)!, bytes = files.get(path)!; const buffer = new ArrayBuffer(bytes.length); new Uint8Array(buffer).set(bytes); url = URL.createObjectURL(new Blob([buffer], { type: asset.mimeType })); byPath.set(path, url); urls.push(url); }
    image.src = url;
  }
  return { model, document, urls };
}
