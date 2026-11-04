import { escapeHtml } from "../ooxml/styles.js";

export interface Html { tag: string; attrs: Record<string, string>; children: (Html | string)[] }
export function serialize(node: Html): string {
  const attrs = Object.entries(node.attrs).map(([name, value]) => " " + name + '="' + escapeHtml(value) + '"').join("");
  return "<" + node.tag + attrs + ">" + (["br", "img"].includes(node.tag) ? "" : node.children.map(item => typeof item === "string" ? escapeHtml(item) : serialize(item)).join("") + "</" + node.tag + ">");
}
export const encode = (value: string): Uint8Array => new (globalThis as unknown as { TextEncoder: new () => { encode(value: string): Uint8Array } }).TextEncoder().encode(value);
