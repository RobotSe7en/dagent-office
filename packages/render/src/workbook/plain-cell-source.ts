import { DocumentError } from "@doffi/core";
import type { WorkbookCell } from "@doffi/core";

export const workbookCellSourceAttributes: ReadonlySet<string> = new Set(["data-column", "data-value-type", "data-value", "data-formula", "data-cache-state", "data-format-id", "class"]);
export function checkWorkbookCellSource(source: WorkbookCell, attribute: (name: string) => string | null, text: string): void {
  const value = attribute("data-value"), equalValue = source.value.type === "number" ? value !== null && value.trim() !== "" && Number(value) === source.value.value
    : source.value.type === "boolean" ? value === String(source.value.value) : source.value.type === "error" ? value === source.value.value : value === null;
  if (attribute("data-column") !== String(source.column) || attribute("data-value-type") !== source.value.type
    || (attribute("class") ?? undefined) !== source.cssClass || (attribute("data-format-id") ?? undefined) !== source.formatId
    || !equalValue || text !== source.text || source.value.type === "text" && source.value.value !== text
    || (attribute("data-formula") ?? undefined) !== source.formula?.expression || (attribute("data-cache-state") ?? undefined) !== source.formula?.cacheState)
    throw new DocumentError("InvalidDocument", "Workbook cell HTML does not match its queried source.");
}

/** Exact small fast path for canonical, entity-free plain cells. Anything
 * requiring HTML decoding or rich structure uses the existing inert DOM parser.
 * No input markup is inserted or silently stripped by this classifier. */
export function checkedPlainWorkbookAttributes(source: WorkbookCell): ReadonlyMap<string, string> | undefined {
  const html = source.html;
  if (typeof html !== "string" || html.length > 8 * 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "Workbook cell HTML exceeds its source budget.");
  if (!html.startsWith("<td") || !/[\t\n\f >]/.test(html[3] ?? "") || !html.endsWith("</td>") || /[&\r\u0000]/.test(html)) return undefined;
  const end = html.indexOf(">"), opening = html.slice(3, end), body = html.slice(end + 1, -5);
  if (end < 0 || opening.includes("<") || body.includes("<")) return undefined;
  const attributes = new Map<string, string>(), pattern = /[\t\n\f ]+([a-z][a-z0-9-]*)="([^"<>]*)"/gy; let cursor = 0;
  while (cursor < opening.length) {
    pattern.lastIndex = cursor; const match = pattern.exec(opening);
    if (!match) { if (!/^[\t\n\f ]*$/.test(opening.slice(cursor))) return undefined; break; }
    const name = match[1]!;
    if (!workbookCellSourceAttributes.has(name) || attributes.has(name)) throw new DocumentError("InvalidDocument", "A workbook plain cell has duplicate or unsupported attributes.");
    attributes.set(name, match[2]!); cursor = pattern.lastIndex;
  }
  checkWorkbookCellSource(source, name => attributes.get(name) ?? null, body);
  return attributes;
}
export function checkedPlainWorkbookCell(source: WorkbookCell): boolean { return checkedPlainWorkbookAttributes(source) !== undefined; }
