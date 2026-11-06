import { DocumentError } from "@doffi/core";
import { markViewNode } from "../view-nodes.js";
import { svgElement } from "../presentation/svg.js";

/** Keep underline ink inside bottom-aligned cells. The reserved band rotates
 * and shrinks with the complete text box, including automatic row measurement. */
export function workbookUnderlineGutter(text: HTMLElement): number {
  if (!text.dataset.officeUnderlinedText) return 0;
  const view = text.ownerDocument.defaultView!; let gutter = 0;
  for (const source of text.querySelectorAll<HTMLElement>("[data-office-native-underline]")) {
    const size = parseFloat(view.getComputedStyle(source).fontSize);
    if (!Number.isFinite(size) || size < 0) throw new DocumentError("InvalidDocument", "Worksheet underlines have an invalid font size.");
    gutter = Math.max(gutter, size * (["double", "doubleAccounting"].includes(source.dataset.officeNativeUnderline!) ? .35 : .125));
  }
  return gutter;
}

/** Accounting strokes span the complete cell, below the measured font boxes.
 * Merge script/font fragments on the same line before painting. Native rich
 * styles remain in the text root; the clipped graphic adds no text or offsets. */
export function appendWorkbookAccountingUnderlines(element: HTMLElement, text: HTMLElement,
  width: number, height: number, x: number, y: number, angle: number, scale: number,
  account: (nodes: number) => void, inspect: (keyUnits: number) => void): { readonly nodes: number; readonly codeUnits: number } {
  if (!text.dataset.officeAccountingText || scale === 0) return { nodes: 0, codeUnits: 0 };
  const document = text.ownerDocument, view = document.defaultView!, origin = text.getBoundingClientRect();
  const rectangles: { top: number; bottom: number; size: number; color: string; kind?: string }[] = [];
  for (const source of text.querySelectorAll<HTMLElement>("[data-office-script]")) {
    inspect(0);
    const style = view.getComputedStyle(source), size = parseFloat(style.fontSize), kind = source.dataset.officeAccountingUnderline;
    for (const rect of source.getClientRects()) {
      inspect(0);
      if (![rect.top, rect.bottom, size].every(Number.isFinite) || size < 0) throw new DocumentError("ResourceLimitExceeded", "Accounting underline font boxes exceed finite geometry.");
      if (rect.height > 0) rectangles.push({ top: rect.top - origin.top, bottom: rect.bottom - origin.top, size, color: style.color, ...(kind ? { kind } : {}) });
    }
  }
  if (!rectangles.some(rect => rect.kind)) return { nodes: 0, codeUnits: 0 };
  rectangles.sort((a, b) => a.top - b.top || a.bottom - b.bottom);
  const lines: { top: number; bottom: number }[] = [];
  for (const rect of rectangles) {
    const previous = lines.at(-1);
    if (previous && rect.top < previous.bottom - Math.min(.1, (rect.bottom - rect.top) / 4, (previous.bottom - previous.top) / 4)) previous.bottom = Math.max(previous.bottom, rect.bottom);
    else lines.push({ top: rect.top, bottom: rect.bottom });
  }
  const strokes = new Map<string, { bottom: number; size: number; color: string; kind: string }>();
  for (const rect of rectangles) if (rect.kind) {
    let low = 0, high = lines.length;
    while (low < high) { const mid = (low + high) >>> 1; if (lines[mid]!.bottom <= rect.top) low = mid + 1; else high = mid; }
    const line = lines[low]; if (!line) throw new DocumentError("InvalidOperation", "An accounting underline has no measured text line.");
    const key = JSON.stringify([low, rect.color, rect.kind]); inspect(key.length);
    const previous = strokes.get(key);
    if (!previous || rect.size > previous.size) strokes.set(key, { bottom: line.bottom, size: rect.size, color: rect.color, kind: rect.kind });
  }
  const cosine = Math.cos(angle), sine = Math.sin(angle);
  const corners = [[0, 0], [width, 0], [0, height], [width, height]].map(([px, py]) => (px! - x) * cosine + (py! - y) * sine);
  const from = Math.min(...corners), to = Math.max(...corners);
  if (![from, to].every(Number.isFinite) || to <= from) throw new DocumentError("ResourceLimitExceeded", "Accounting underlines exceed their finite cell geometry.");
  let nodes = 0;
  const create: typeof svgElement = (owner, tag, attributes) => { account(1); nodes++; return svgElement(owner, tag, attributes); };
  const layer = markViewNode(create(document, "svg", { width: "100%", height: "100%", viewBox: `0 0 ${width} ${height}`, "aria-hidden": "true" }));
  layer.dataset.officeView = "workbook-accounting-underlines"; layer.style.cssText = "position:absolute;inset:0;overflow:hidden;pointer-events:none;user-select:none"; element.append(layer);
  for (const stroke of strokes.values()) {
    const double = stroke.kind === "doubleAccounting", thickness = stroke.size * (double ? .035 : .055);
    if (thickness === 0) continue;
    const group = create(document, "g", {}); group.dataset.officeView = "workbook-accounting-line"; group.dataset.officeAccountingStyle = stroke.kind; layer.append(group);
    for (let index = 0; index < (double ? 2 : 1); index++) {
      const localY = stroke.bottom + stroke.size * (.05 + index * .2) + thickness / 2;
      const startX = x + cosine * from - sine * localY * scale, startY = y + sine * from + cosine * localY * scale;
      const endX = x + cosine * to - sine * localY * scale, endY = y + sine * to + cosine * localY * scale;
      if (![startX, startY, endX, endY, thickness * scale].every(Number.isFinite)) throw new DocumentError("ResourceLimitExceeded", "Accounting strokes exceed their coordinate scale.");
      group.append(create(document, "line", { x1: startX, y1: startY, x2: endX, y2: endY, stroke: stroke.color, "stroke-width": thickness * scale, "stroke-linecap": "butt" }));
    }
  }
  return { nodes, codeUnits: 0 };
}
