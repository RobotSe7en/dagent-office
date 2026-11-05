import { DocumentError } from "@dagent-office/core";
import { resolvedWorkbookColor } from "./color.js";

export function copyWorkbookTextStyle(original: HTMLElement, target: HTMLElement, background = false): CSSStyleDeclaration {
  target.dataset.officeTextStyle = "true";
  const document = original.ownerDocument, computed = document.defaultView!.getComputedStyle(original);
  for (const property of ["font-family", "font-size", "font-weight", "font-style", "font-variant", "line-height", "letter-spacing", "word-spacing", "text-decoration-line", "text-decoration-style", "text-decoration-color", "text-transform", "vertical-align"])
    target.style.setProperty(property, computed.getPropertyValue(property));
  target.style.color = resolvedWorkbookColor(document, computed, "--office-cell-foreground") ?? computed.color;
  if (background) target.style.backgroundColor = computed.backgroundColor;
  return computed;
}

export function decorateWorkbookText(computed: CSSStyleDeclaration, target: HTMLElement, account: (nodes: number) => void): HTMLElement {
  const underline = computed.getPropertyValue("--office-font-underline").trim() || "none", strike = computed.getPropertyValue("--office-font-strike").trim() || "false";
  if (!["none", "single", "double", "singleAccounting", "doubleAccounting"].includes(underline) || !["true", "false"].includes(strike)) throw new DocumentError("InvalidDocument", "A worksheet text decoration is invalid.");
  if (underline.endsWith("Accounting")) throw new DocumentError("UnsupportedFeature", "Accounting underlines require worksheet line geometry.");
  target.style.textDecorationLine = underline !== "none" ? "underline" : "none";
  target.style.textDecorationStyle = underline === "double" ? "double" : "solid";
  target.style.textDecorationColor = "currentColor";
  if (strike === "true") {
    account(1); const inner = target.ownerDocument.createElement("span"); inner.style.textDecoration = "line-through solid currentColor"; target.append(inner); return inner;
  }
  return target;
}
