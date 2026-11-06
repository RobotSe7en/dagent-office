import { DocumentError } from "@doffi/core";
import { markViewNode } from "../view-nodes.js";

export type WorkbookReadingOrder = "context" | "ltr" | "rtl";

/** SpreadsheetML context direction inspects the first non-whitespace
 * character, not the first strong character anywhere in the string. Use the
 * browser's Unicode bidi table on that one scalar only. This avoids both a
 * partial script-name approximation and host-locale/document-direction state.
 * The small cache belongs to one display scope and is cleared on disposal. */
export class WorkbookTextDirection {
  private readonly root: HTMLElement;
  private readonly probe: HTMLElement;
  private readonly cache = new Map<string, "ltr" | "rtl">();
  private disposed = false;
  constructor(scope: HTMLElement) {
    if (!scope.isConnected || !scope.ownerDocument.defaultView) throw new DocumentError("InvalidOperation", "Worksheet reading order requires a connected measurement scope.");
    this.root = markViewNode(scope.ownerDocument.createElement("span")); this.root.dir = "ltr";
    this.root.setAttribute("aria-hidden", "true"); this.root.dataset.officeView = "workbook-direction-measurement";
    this.root.style.cssText = "position:absolute;visibility:hidden;pointer-events:none;overflow:hidden;width:0;height:0";
    this.probe = scope.ownerDocument.createElement("span"); this.probe.dir = "auto"; this.root.append(this.probe); scope.append(this.root);
  }
  read(text: string, order: WorkbookReadingOrder): "ltr" | "rtl" {
    if (this.disposed || !this.root.isConnected) throw new DocumentError("InvalidOperation", "Worksheet reading-order measurement is no longer connected.");
    if (order === "ltr" || order === "rtl") return order;
    if (order !== "context") throw new DocumentError("InvalidDocument", "A worksheet reading order is invalid.");
    let first = "";
    for (const character of text) if (!/^\p{White_Space}$/u.test(character)) { first = character; break; }
    if (!first) return "ltr";
    const previous = this.cache.get(first); if (previous) return previous;
    this.probe.textContent = first;
    const direction = this.probe.ownerDocument.defaultView!.getComputedStyle(this.probe).direction;
    this.probe.replaceChildren();
    if (direction !== "ltr" && direction !== "rtl") throw new DocumentError("UnsupportedFeature", "This browser cannot resolve worksheet reading order.");
    if (this.cache.size >= 256) this.cache.clear(); this.cache.set(first, direction);
    return direction;
  }
  dispose(): void { this.root.remove(); this.root.replaceChildren(); this.cache.clear(); this.disposed = true; }
}
