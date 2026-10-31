import { DocumentError } from "@dagent-office/core";
import { markViewNode, ownViewStyle } from "../view-nodes.js";
import { property, resolvedColor } from "./appearance.js";

/** Native group scaling changes object rectangles; it does not mirror glyphs
 * or rewrite their point sizes. Resolve the cumulative linear view transform. */
export function textTransform(node: HTMLElement, rect: { x: number; y: number; width: number; height: number }): string {
  const view = node.ownerDocument.defaultView!; let matrix = new view.DOMMatrix(), flipH = false;
  for (let current: HTMLElement | null = node; current && current.dataset.officeKind !== "slide"; current = current.parentElement) {
    const transform = view.getComputedStyle(current).transform;
    if (transform !== "none") matrix = new view.DOMMatrix(transform).multiply(matrix);
    if (current.dataset.officeFlipH === "true") flipH = !flipH;
  }
  const scaleX = Math.hypot(matrix.a, matrix.b), scaleY = Math.hypot(matrix.c, matrix.d);
  if (!scaleX || !scaleY) throw new DocumentError("InvalidDocument", "A text instance has a singular group transform.");
  const angle = Math.atan2(matrix.b * (flipH ? -1 : 1), matrix.a * (flipH ? -1 : 1));
  const rotation = new view.DOMMatrix([Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0]);
  matrix.e = 0; matrix.f = 0;
  const local = matrix.inverse().multiply(rotation), width = rect.width * scaleX, height = rect.height * scaleY;
  return `left:${rect.x + (rect.width - width) / 2}pt;top:${rect.y + (rect.height - height) / 2}pt;width:${width}pt;height:${height}pt;transform-origin:50% 50%;transform:matrix(${local.a},${local.b},${local.c},${local.d},0,0)`;
}
function numberLabel(value: number, format: string): string {
  let label: string;
  if (format.startsWith("alpha")) {
    label = ""; for (let n = value; n > 0; n = Math.floor((n - 1) / 26)) label = String.fromCharCode(65 + (n - 1) % 26) + label;
    if (format.startsWith("alphaLc")) label = label.toLowerCase();
  } else if (format.startsWith("roman")) {
    if (value > 1_000_000) throw new DocumentError("ResourceLimitExceeded", "A presentation list label exceeds its budget.");
    label = ""; let remaining = value;
    for (const [amount, token] of [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]] as const) while (remaining >= amount) { label += token; remaining -= amount; }
    if (format.startsWith("romanLc")) label = label.toLowerCase();
  } else label = String(value);
  return format.endsWith("ParenBoth") ? "(" + label + ")" : format.endsWith("ParenR") ? label + ")" : format.endsWith("Plain") ? label : label + ".";
}
export function decorateLists(root: HTMLElement, account: (n: number) => void): void {
  const document = root.ownerDocument;
  for (const box of root.querySelectorAll<HTMLElement>(".sdk-ppt-text,.sdk-ppt-rich-story")) {
    const counters = new Map<string, number>();
    for (const paragraph of box.querySelectorAll<HTMLElement>("p")) {
      const s = document.defaultView!.getComputedStyle(paragraph), kind = paragraph.hasAttribute("data-office-list-none") ? "none" : paragraph.dataset.officeBullet ? "bullet" : paragraph.dataset.officeNumbering ? "number" : property(s, "--office-list-kind");
      if (kind !== "bullet" && kind !== "number") { counters.clear(); continue; }
      const marker = markViewNode(document.createElement("span")); marker.className = "sdk-ppt-marker";
      const character = paragraph.dataset.officeBullet ?? property(s, "--office-list-character", "•").replace(/^"|"$/g, "");
      const format = paragraph.dataset.officeNumbering ?? property(s, "--office-list-numbering", "arabicPeriod"), key = (paragraph.dataset.officeLevel ?? "0") + ":" + format;
      const value = counters.get(key) ?? Number(paragraph.dataset.officeNumberingStart ?? property(s, "--office-list-start", "1")); counters.set(key, value + 1);
      marker.textContent = kind === "bullet" ? character : numberLabel(value, format);
      const indent = parseFloat(s.textIndent) || -24, margin = parseFloat(s.marginLeft) || 24;
      marker.style.cssText = `left:${indent}px;width:${Math.abs(indent)}px;font-family:${property(s, "--office-bullet-font", s.fontFamily)};font-size:${property(s, "--office-bullet-size", "1em")};color:${resolvedColor(document, s, "--office-bullet", s.color)}`;
      ownViewStyle(paragraph, "position", "text-indent", "margin-left"); paragraph.style.position = "relative"; paragraph.style.textIndent = "0"; paragraph.style.marginLeft = margin + "px"; paragraph.prepend(marker); account(1);
    }
  }
}
