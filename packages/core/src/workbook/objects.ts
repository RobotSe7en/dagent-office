import type { WorkbookMetadata, WorkbookObject } from "../contracts/workbook.js";
import type { ParsedHtml } from "../document/html.js";
import { invalid } from "../document/validation.js";
import { evaluatePresentationPreset } from "../presentation-geometry.js";

const numeric = /^-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i;
function number(value: string): number {
  if (!numeric.test(value) || !Number.isFinite(Number(value))) invalid("A workbook drawing attribute must contain a finite number.");
  return Number(value);
}
/** Validate ownership and the source attributes that the object adapter uses.
 * Worksheet anchor coordinates are never duplicated as HTML x/y dimensions. */
export function workbookObjects(workbook: WorkbookMetadata, html: ParsedHtml | undefined): ReadonlyMap<string, WorkbookObject> {
  const sheets = new Set(workbook.sheets.map(sheet => sheet.sheetId)), objects = new Map<string, WorkbookObject>(), native = new Map<string, Set<string>>();
  for (const object of workbook.objects ?? []) {
    const node = html?.peekNode(object.nodeId), parent = node?.parentNodeId ? html?.peekNode(node.parentNodeId) : undefined;
    if (!sheets.has(object.sheetId) || objects.has(object.nodeId) || !node || node.tagName !== "div" || node.attributes["data-office-kind"] !== object.kind || parent?.attributes["data-office-kind"] !== "workbook-content" || parent.parentNodeId !== null) invalid("Workbook objects require unique source containers on an existing worksheet.");
    if (node.attributes["data-office-story"] !== undefined || html!.directText(object.nodeId)?.trim()) invalid("Workbook object text must belong to its rich-text paragraphs.");
    if (object.nativeId !== undefined) {
      const ids = native.get(object.sheetId) ?? new Set();
      if (ids.has(object.nativeId)) invalid("Native drawing identities must be unique within a worksheet.");
      ids.add(object.nativeId); native.set(object.sheetId, ids);
    }
    const adjustments: Record<string, number> = {};
    for (const [key, value] of Object.entries(node.attributes)) {
      if (!key.startsWith("data-office-")) continue;
      if (["data-office-kind", "data-office-alt"].includes(key)) continue;
      if (["data-office-flip-h", "data-office-flip-v", "data-office-hidden"].includes(key)) { if (!["true", "false"].includes(value)) invalid("Drawing boolean attributes must be true or false."); }
      else if (key === "data-office-rotation") number(value);
      else if (key.startsWith("data-office-adjust-") && ["text", "geometry"].includes(object.kind) && /^data-office-adjust-[a-z][a-z0-9]*$/.test(key)) {
        if (!/^-?\d+$/.test(value) || !Number.isSafeInteger(Number(value))) invalid("Shape adjustments require safe integer constants.");
        adjustments[key.slice("data-office-adjust-".length)] = Number(value);
      } else if (/^data-office-crop-(left|right|top|bottom)$/.test(key) && object.kind === "image") { if (number(value) < 0 || number(value) >= 1) invalid("Image crop fractions must lie in [0, 1)."); }
      else invalid("This workbook object attribute has no source representation.");
    }
    if (object.kind === "image") {
      if (node.childNodeIds.length !== 1 || html!.peekNode(node.childNodeIds[0]!)?.tagName !== "img") invalid("An image object requires exactly one image resource.");
      const crop = (side: string): number => Number(node.attributes[`data-office-crop-${side}`] ?? 0);
      if (crop("left") + crop("right") >= 1 || crop("top") + crop("bottom") >= 1) invalid("An image crop must retain a visible area.");
    } else {
      if (node.childNodeIds.some(id => html!.peekNode(id)?.tagName !== "p") || (object.kind === "text" || object.text !== undefined) && !node.childNodeIds.length) invalid("A drawing text body requires source paragraphs.");
      evaluatePresentationPreset(object.kind === "geometry" ? object.preset : "rect", 100, 100, adjustments);
    }
    objects.set(object.nodeId, object);
  }
  return objects;
}
