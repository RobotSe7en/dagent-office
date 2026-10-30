import { DocumentError } from "@dagent-office/core";
import type { RenderedPage, RenderLimits } from "./contracts.js";

/** Clip the measured canvas per page. Chromium retains searchable text and culls
 * content outside each page; the editable source and its flow remain unchanged. */
export function createPrintDocument(document: Document, root: HTMLElement, pages: readonly RenderedPage[], limits: Readonly<RenderLimits>): Document {
  const article = root.closest<HTMLElement>('[data-office-kind="document"]')!;
  const nodes = article.querySelectorAll("*").length + 1;
  if (pages.length * nodes > limits.maxPrintNodes) throw new DocumentError("ResourceLimitExceeded", "Print projection exceeds its node budget.");
  const output = document.implementation.createHTMLDocument("");
  output.head.replaceWith(output.importNode(document.head, true)); output.body.replaceChildren();
  const style = output.createElement("style");
  style.textContent = "html,body{margin:0!important;padding:0!important;background:white!important}.sdk-print-page{position:relative;overflow:hidden;break-after:page;box-sizing:border-box}.sdk-print-page:last-child{break-after:auto}.sdk-page-number{display:none!important}";
  const surface = root.getBoundingClientRect(), origin = article.getBoundingClientRect();
  for (const page of pages) {
    style.textContent += `@page sdk-page-${page.index}{size:${page.setup.width}pt ${page.setup.height}pt;margin:0}`;
    const paper = output.createElement("div"); paper.className = "sdk-print-page";
    paper.style.cssText = `width:${page.setup.width}pt;height:${page.setup.height}pt;page:sdk-page-${page.index}`;
    const background = document.querySelector<HTMLElement>('.sdk-paper[data-page-id="' + CSS.escape(page.pageId) + '"]');
    if (background) {
      const clone = output.importNode(background, true); clone.removeAttribute("aria-hidden");
      clone.style.cssText = `position:absolute;left:0;top:0;width:${page.setup.width}pt;height:${page.setup.height}pt;box-shadow:none;background:white`;
      paper.append(clone);
    }
    const content = output.importNode(article, true);
    content.style.cssText += `;position:absolute;left:${-(page.rect.x + (surface.left - origin.left) * 72 / 96)}pt;top:${-(page.rect.y + (surface.top - origin.top) * 72 / 96)}pt;margin:0;`;
    paper.append(content); output.body.append(paper);
  }
  output.head.append(style); return output;
}
