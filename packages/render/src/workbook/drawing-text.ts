import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WorkbookObject } from "@dagent-office/core";
import { markViewNode } from "../view-nodes.js";
import { workbookRichTextFonts } from "./text-fonts.js";

export interface WorkbookDrawingTextProjection {
  readonly node: Text;
  readonly sourceNodeId: string;
  /** UTF-16 offsets in the whole object; paragraph boundaries add no units. */
  readonly from: number;
  readonly to: number;
}
export interface WorkbookDrawingTextView {
  readonly element: HTMLElement;
  readonly projections: readonly WorkbookDrawingTextProjection[];
  readonly nodes: number;
}

const inline = new Set(["span", "strong", "em", "u", "s", "sub", "sup", "br"]);
const typography = ["font-family", "font-size", "font-weight", "font-style", "font-variant", "line-height", "letter-spacing", "word-spacing", "color", "text-decoration-line", "text-decoration-style", "text-decoration-color", "text-transform", "vertical-align", "direction", "unicode-bidi", "tab-size"];
const paragraphStyle = ["text-align", "text-indent", "margin-left", "margin-right", "margin-top", "margin-bottom"];
const listAttributes = ["data-office-bullet", "data-office-numbering", "data-office-numbering-start", "data-office-level", "data-office-list-none"];

/** Materialize typography while the source still has its validated ancestry.
 * The detached result owns no source IDs/classes and carries no executable or
 * resource attributes. Layout, lists, local-font diagnostics and pane clipping
 * belong to its compositor; no source node is moved or changed here. */
export function createWorkbookDrawingText(source: HTMLElement, object: DeepReadonly<WorkbookObject>, maxNodes = 50_000): WorkbookDrawingTextView {
  const document = source.ownerDocument, view = document.defaultView;
  if (!source.isConnected || !view || !Number.isSafeInteger(maxNodes) || maxNodes < 1) throw new DocumentError("InvalidOperation", "Drawing text requires a connected source and a positive node budget.");
  if (object.kind !== "text" && object.kind !== "geometry" || source.dataset.nodeId !== object.nodeId || source.dataset.officeKind !== object.kind) throw new DocumentError("InvalidDocument", "Drawing text requires its own shape or text-box source.");
  const root = markViewNode(document.createElement("div")), projections: WorkbookDrawingTextProjection[] = [];
  root.dataset.officeView = "workbook-drawing-text"; root.dataset.sourceNodeId = object.nodeId;
  let nodes = 1, offset = 0;
  const account = (amount: number): void => { nodes += amount; if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Drawing text exceeds its node budget."); };
  const style = (original: HTMLElement, target: HTMLElement): CSSStyleDeclaration => {
    const computed = view.getComputedStyle(original); target.dataset.officeTextStyle = "true";
    for (const name of typography) target.style.setProperty(name, computed.getPropertyValue(name));
    if (original !== source) target.style.backgroundColor = computed.backgroundColor;
    if (original.lang) target.lang = original.lang;
    if (original.dataset.nodeId) target.dataset.sourceNodeId = original.dataset.nodeId;
    return computed;
  };
  style(source, root); root.style.whiteSpace = "pre-wrap";
  for (const child of source.childNodes) {
    if (child.nodeType === 3 && !child.textContent?.trim()) { offset += child.textContent?.length ?? 0; continue; }
    if (child.nodeType !== 1 || (child as Element).localName !== "p") throw new DocumentError("InvalidDocument", "Drawing text must consist of source paragraphs.");
    account(1);
    const paragraph = child as HTMLElement, runs = workbookRichTextFonts(paragraph, Math.max(1, maxNodes - nodes)), baseOffset = offset;
    let runIndex = 0, localOffset = 0;
    const target = document.createElement("p"), computed = style(paragraph, target);
    target.style.margin = "0"; target.style.minHeight = "1em"; target.style.flex = "none";
    for (const name of paragraphStyle) target.style.setProperty(name, computed.getPropertyValue(name));
    for (const name of listAttributes) { const value = paragraph.getAttribute(name); if (value !== null) target.setAttribute(name, value); }
    for (const name of computed) if (name.startsWith("--office-list-") || name.startsWith("--office-bullet-")) target.style.setProperty(name, computed.getPropertyValue(name));
    const appendText = (text: string, original: HTMLElement, parent: HTMLElement): void => {
      const end = localOffset + text.length;
      while (localOffset < end) {
        const run = runs[runIndex++];
        if (!run || run.from !== localOffset || run.to > end) throw new DocumentError("InvalidDocument", "Drawing font runs do not match the source text.");
        account(2); const span = document.createElement("span"), node = document.createTextNode(run.text);
        span.dataset.officeScript = run.role; span.style.fontFamily = run.font.family; span.append(node); parent.append(span);
        projections.push(Object.freeze({ node, sourceNodeId: original.dataset.nodeId!, from: baseOffset + run.from, to: baseOffset + run.to })); localOffset = run.to;
      }
    };
    const append = (original: HTMLElement, parent: HTMLElement, depth: number): void => {
      if (depth > 64) throw new DocumentError("ResourceLimitExceeded", "Drawing text exceeds its nesting budget.");
      for (const child of original.childNodes) {
        if (child.nodeType === 3) appendText(child.textContent ?? "", original, parent);
        else if (child.nodeType === 1 && inline.has((child as Element).localName)) {
          const element = child as HTMLElement; account(1); const span = document.createElement("span"); style(element, span); parent.append(span);
          if (element.localName === "br") appendText("\n", element, span);
          else append(element, span, depth + 1);
        } else throw new DocumentError("InvalidDocument", "Drawing text contains unsupported inline content.");
      }
    };
    append(paragraph, target, 0);
    if (runIndex !== runs.length) throw new DocumentError("InvalidDocument", "Drawing text did not consume all source font runs.");
    offset += localOffset; root.append(target);
  }
  return Object.freeze({ element: root, projections: Object.freeze(projections), nodes });
}
