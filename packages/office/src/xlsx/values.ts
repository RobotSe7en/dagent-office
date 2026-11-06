import { DocumentError } from "@doffi/core";
import type { WorkbookCoordinate, WorkbookRange } from "@doffi/core";
import { attr, elements, fail, S, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { escapeHtml } from "../ooxml/styles.js";

export const value = (node: XmlNode | undefined, name: string): string | undefined => attr(node, name, "");
export function boolean(node: XmlNode | undefined, name: string, fallback = false): boolean {
  const raw = value(node, name); if (raw === undefined) return fallback;
  if (raw === "1" || raw === "true") return true;
  if (raw === "0" || raw === "false") return false;
  return fail(node!, "Invalid SpreadsheetML boolean.");
}
export function finite(raw: string, node: XmlNode): number {
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(raw.trim()) || !Number.isFinite(Number(raw))) fail(node, "Invalid SpreadsheetML numeric value.");
  return Number(raw);
}
export function number(node: XmlNode, name: string, fallback?: number): number {
  const raw = value(node, name); if (raw === undefined && fallback !== undefined) return fallback;
  if (raw === undefined) fail(node, "A SpreadsheetML numeric attribute is missing.");
  return finite(raw, node);
}
export function integer(node: XmlNode, name: string, minimum = 0, maximum = 0xffff_ffff, fallback?: number): number {
  const raw = value(node, name);
  if (raw !== undefined && !/^[+-]?\d+$/.test(raw.trim())) fail(node, "A SpreadsheetML integer has invalid lexical form.");
  const result = number(node, name, fallback);
  if (!Number.isSafeInteger(result) || result < minimum || result > maximum) fail(node, "A SpreadsheetML integer is outside its range.");
  return result;
}
export function one(node: XmlNode | undefined, name: string, uri: string = S): XmlNode | undefined {
  const found = elements(node, name, uri); if (found.length > 1) fail(node!, "A singleton SpreadsheetML element is duplicated.");
  return found[0];
}
export function attributes(node: XmlNode, names: readonly string[], namespace = ""): void {
  for (const [key] of node.attrs) {
    if (key.startsWith("http://www.w3.org/2000/xmlns/ ")) continue;
    if (!names.includes(key.slice(namespace.length + 1)) || !key.startsWith(namespace + " ")) unsupported(node, "This SpreadsheetML attribute is not supported.");
  }
}
export function address(source: string, node: Pick<XmlNode, "part" | "name">): WorkbookCoordinate {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(source);
  if (!match) fail(node, "Invalid worksheet cell address.");
  let column = 0;
  for (const character of match[1]!.toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64;
  const row = Number(match[2]);
  if (row > 1_048_576 || column > 16_384) fail(node, "Worksheet address exceeds native bounds.");
  return { row: row - 1, column: column - 1 };
}
export function range(source: string, node: Pick<XmlNode, "part" | "name">): WorkbookRange {
  const parts = source.split(":"); if (parts.length > 2) fail(node, "Invalid worksheet range.");
  const from = address(parts[0]!, node), to = address(parts[1] ?? parts[0]!, node);
  if (from.row > to.row || from.column > to.column) fail(node, "Worksheet range endpoints must be ordered.");
  return { from, to };
}
/** Decode native UTF-16 escapes exactly once, including literal _xNNNN_ text. */
export function spreadsheetText(source: string, node: XmlNode, maximum = 32767): string {
  if (source.length > maximum * 7) throw new DocumentError("ResourceLimitExceeded", "Spreadsheet text exceeds its encoded budget.");
  const decoded = source.replace(/_x([\da-fA-F]{4})_/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)));
  if (decoded.length > maximum) throw new DocumentError("ResourceLimitExceeded", "Spreadsheet text exceeds its UTF-16 budget.");
  for (let index = 0; index < decoded.length; index++) {
    const code = decoded.charCodeAt(index);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = decoded.charCodeAt(++index); if (!(next >= 0xdc00 && next <= 0xdfff)) fail(node, "Spreadsheet text contains an unpaired UTF-16 surrogate.");
    } else if (code >= 0xdc00 && code <= 0xdfff) fail(node, "Spreadsheet text contains an unpaired UTF-16 surrogate.");
    else if (code < 32 && ![9, 10, 13].includes(code) || code === 0xfffe || code === 0xffff) unsupported(node, "Spreadsheet text contains a character not representable in the document source.");
  }
  return decoded;
}
/** Preserve carriage returns through HTML parsing rather than normalizing them. */
export const textHtml = (source: string): string => escapeHtml(source).replaceAll("\r", "&#13;");
export const nativeText = (node: XmlNode, maximum = 32767): string => spreadsheetText(text(node), node, maximum);
export const encode = (source: string): Uint8Array => new (globalThis as unknown as { TextEncoder: new () => { encode(value: string): Uint8Array } }).TextEncoder().encode(source);
