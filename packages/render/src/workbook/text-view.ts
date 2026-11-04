import { DocumentError } from "@dagent-office/core";
import { markViewNode } from "../view-nodes.js";
import type { WorkbookSourceCell } from "./cell-source.js";
import { resolvedWorkbookColor } from "./color.js";
import { workbookTextFonts } from "./text-fonts.js";

export interface WorkbookTextProjection { readonly node: Text; readonly from: number; readonly to: number }
export interface WorkbookTextView { readonly element: HTMLElement; readonly projections: readonly WorkbookTextProjection[]; readonly nodes: number }

/** Build display-only rich text from computed source styles, before any virtual
 * grid ancestry can change their meaning. Keep source span nesting for CSS
 * decoration propagation; native font decorations belong to individual runs so
 * a rich-text run can explicitly turn off its cell's underline/strike. */
export function createWorkbookTextView(source: WorkbookSourceCell, maxNodes = 50_000): WorkbookTextView {
  const { element } = source, document = element.ownerDocument, view = document.defaultView;
  if (!view || !element.isConnected || !Number.isSafeInteger(maxNodes) || maxNodes < 1) throw new DocumentError("InvalidOperation", "Workbook text display requires a connected source and a positive node budget.");
  const runs = workbookTextFonts(source, maxNodes), root = markViewNode(document.createElement("span")), projections: WorkbookTextProjection[] = [];
  let nodes = 1, offset = 0, runIndex = 0;
  const account = (amount: number): void => { nodes += amount; if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook text display exceeds its node budget."); };
  const copyStyle = (original: HTMLElement, target: HTMLElement): CSSStyleDeclaration => {
    target.dataset.officeTextStyle = "true";
    const computed = view.getComputedStyle(original);
    for (const property of ["font-family", "font-size", "font-weight", "font-style", "font-variant", "line-height", "letter-spacing", "word-spacing", "text-decoration-line", "text-decoration-style", "text-decoration-color", "text-transform", "vertical-align"])
      target.style.setProperty(property, computed.getPropertyValue(property));
    target.style.color = resolvedWorkbookColor(document, computed, "--office-cell-foreground") ?? computed.color;
    if (original !== element) target.style.backgroundColor = computed.backgroundColor;
    return computed;
  };
  const decorate = (computed: CSSStyleDeclaration, target: HTMLElement): HTMLElement => {
    const underline = computed.getPropertyValue("--office-font-underline").trim() || "none", strike = computed.getPropertyValue("--office-font-strike").trim() || "false";
    if (!["none", "single", "double", "singleAccounting", "doubleAccounting"].includes(underline) || !["true", "false"].includes(strike)) throw new DocumentError("InvalidDocument", "A worksheet text decoration is invalid.");
    if (underline.endsWith("Accounting")) throw new DocumentError("UnsupportedFeature", "Accounting underlines require worksheet line geometry.");
    target.style.textDecorationLine = underline !== "none" ? "underline" : "none";
    target.style.textDecorationStyle = underline === "double" ? "double" : "solid";
    target.style.textDecorationColor = "currentColor";
    if (strike === "true") {
      // A double underline must not turn the independent strike into two lines.
      account(1); const inner = document.createElement("span"); inner.style.textDecoration = "line-through solid currentColor"; target.append(inner); return inner;
    }
    return target;
  };
  const append = (original: HTMLElement, target: HTMLElement, depth: number): void => {
    if (depth > 16) throw new DocumentError("ResourceLimitExceeded", "Workbook text display exceeds its nesting budget.");
    const computed = copyStyle(original, target);
    for (const child of original.childNodes) {
      if (child.nodeType === 3) {
        const text = child.textContent ?? "", end = offset + text.length;
        while (offset < end) {
          const run = runs[runIndex++];
          if (!run || run.from !== offset || run.to > end) throw new DocumentError("InvalidDocument", "Worksheet font runs do not match their source text.");
          account(2); const span = document.createElement("span"), node = document.createTextNode(run.text);
          span.dataset.officeScript = run.role; span.style.fontFamily = run.font.family; decorate(computed, span).append(node); target.append(span);
          projections.push(Object.freeze({ node, from: run.from, to: run.to })); offset = run.to;
        }
      } else if (child.nodeType === 1 && (child as Element).localName === "br") {
        const run = runs[runIndex++];
        if (!run || run.from !== offset || run.text !== "\n") throw new DocumentError("InvalidDocument", "Worksheet line breaks do not match their source text.");
        // A literal newline keeps one UTF-16 source unit and can participate in
        // DOM ranges. White-space handling belongs to the cell layout owner.
        account(1); const node = document.createTextNode("\n"); target.append(node);
        projections.push(Object.freeze({ node, from: offset, to: ++offset }));
      } else if (child.nodeType === 1 && (child as Element).localName === "span") {
        account(1); const span = document.createElement("span"); append(child as HTMLElement, span, depth + 1); target.append(span);
      } else throw new DocumentError("InvalidDocument", "Worksheet display text contains unsupported source content.");
    }
  };
  append(element, root, 0); root.style.whiteSpace = "pre";
  if (runIndex !== runs.length) throw new DocumentError("InvalidDocument", "Worksheet display did not consume all font runs.");
  return Object.freeze({ element: root, projections: Object.freeze(projections), nodes });
}
