import { attr, fail, number } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";

export const value = (node: XmlNode | undefined, name = "val"): string | undefined => attr(node, name, "");
export const integer = (node: XmlNode, name: string, fallback?: number): number => number(value(node, name), node, fallback);
export function unsigned(node: XmlNode, name: string, minimum = 0, maximum = 4294967295): number {
  const n = integer(node, name);
  if (n < minimum || n > maximum) fail(node, "A native unsigned identity or index is outside its range.");
  return n;
}
export const points = (node: XmlNode, name: string, fallback?: number): number => integer(node, name, fallback) / 12700;
export function bool(node: XmlNode, name: string, fallback = false): boolean {
  const v = value(node, name);
  if (v === undefined) return fallback;
  if (["true", "1"].includes(v)) return true;
  if (["false", "0"].includes(v)) return false;
  return fail(node, "A DrawingML boolean is invalid.");
}
export function rootName(node: XmlNode, name: string, uri: string): void {
  if (node.uri !== uri || node.name !== name) fail(node, "An OOXML part has the wrong root.");
}
export { cssFontName as fontName } from "../ooxml/css-values.js";
